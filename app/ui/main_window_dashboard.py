"""
FALCON2 - ダッシュボード計算・HTML生成 Mixin
MainWindow から分離した Mixin クラス。
MainWindow のみが継承することを前提とし、self.* 属性は MainWindow.__init__ で初期化される。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import TYPE_CHECKING, Optional, Literal, List, Dict, Any, Tuple, Union
from pathlib import Path
from datetime import datetime, timedelta, date
import json
import logging
import re
import io
import threading
import webbrowser
import tempfile
import html

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False

if TYPE_CHECKING:
    from db.db_handler import DBHandler
    from modules.formula_engine import FormulaEngine
    from modules.rule_engine import RuleEngine

from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager

logger = logging.getLogger(__name__)

# 散布図の Plotly 表示用（乳検・ゲノム共通）
try:
    from modules.genome_report_html import _get_plotly_scatter_script, _build_plotly_scatter_divs_only
except ImportError:
    _get_plotly_scatter_script = None
    _build_plotly_scatter_divs_only = None


class DashboardMixin:
    """Mixin: FALCON2 - ダッシュボード計算・HTML生成 Mixin"""
    def _is_cow_disposed_for_dashboard(self, cow_auto_id: int) -> bool:
        """
        個体が売却または死亡廃用されているかをチェック（ダッシュボード用）
        
        Args:
            cow_auto_id: 牛の auto_id
            
        Returns:
            売却または死亡廃用されている場合True
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for event in events:
            event_number = event.get('event_number')
            if event_number in [RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD]:
                return True
        return False

    def _on_dashboard(self):
        """ダッシュボードメニューをクリック（HTML方式）"""
        try:
            # 農場名を取得
            settings_manager = SettingsManager(self.farm_path)
            farm_name = settings_manager.get("farm_name", self.farm_path.name)
            
            # 全個体を取得
            all_cows = self.db.get_all_cows()
            
            # 現存牛のみを取得（売却・死亡廃用を除く）
            existing_cows = [
                cow for cow in all_cows
                if not self._is_cow_disposed_for_dashboard(cow['auto_id'])
            ]
            
            # 経産牛（LACT>=1）を取得（現存牛のみ）
            lactated_cows = [cow for cow in existing_cows if cow.get('lact') is not None and cow.get('lact', 0) >= 1]
            
            # 産次ごとの頭数を集計
            parity_counts = {}
            for cow in lactated_cows:
                lact = cow.get('lact', 0)
                # LACTを整数に変換して確認
                try:
                    lact_int = int(lact) if lact is not None else 0
                except (ValueError, TypeError):
                    lact_int = 0
                
                if lact_int == 1:
                    parity = "1産"
                elif lact_int == 2:
                    parity = "2産"
                elif lact_int >= 3:
                    parity = "3産以上"
                else:
                    # LACTが1未満の場合はスキップ（本来はフィルタリングされているはず）
                    continue
                parity_counts[parity] = parity_counts.get(parity, 0) + 1
            
            total = sum(parity_counts.values())
            total_lact = sum(cow.get('lact', 0) for cow in lactated_cows)
            avg_parity = total_lact / total if total > 0 else 0
            
            # 産次ごとのドーナツチャート用データを準備
            parity_segments = []
            # モダンなカラーパレット（洗練された色）
            colors = ['#FF6B9D', '#4ECDC4', '#95E1D3']
            radius = 50
            circumference = 2 * 3.141592653589793 * radius
            offset = 0
            
            # 固定順序でソート（1産、2産、3産以上）
            parity_order = ["1産", "2産", "3産以上"]
            sorted_parities = []
            for parity in parity_order:
                if parity in parity_counts and parity_counts[parity] > 0:
                    sorted_parities.append((parity, parity_counts[parity]))
            
            for idx, (parity, count) in enumerate(sorted_parities):
                if count <= 0:
                    continue
                percent = (count / total * 100) if total > 0 else 0
                length = (percent / 100) * circumference
                parity_segments.append({
                    "label": parity,
                    "count": count,
                    "percent": round(percent, 1),
                    "color": colors[idx % len(colors)],
                    "length": length,
                    "offset": -offset  # 負の値として保存（stroke-dashoffsetで使用）
                })
                offset += length
            
            # 平均乳量を計算
            milk_stats = self._calculate_dashboard_milk_stats(lactated_cows)
            
            # 平均分娩後日数と妊娠牛の割合を計算
            herd_summary = self._calculate_herd_summary(lactated_cows)
            
            # 体細胞要約を計算
            scc_summary = self._calculate_dashboard_scc_stats(lactated_cows)
            
            # 散布図データを計算（乳検レポートと同じ：直近乳検日の DIM vs 乳量/リニアスコア）
            scatter_data = self._calculate_dashboard_scatter_data(lactated_cows)
            
            # 受胎率を計算（産次別）
            fertility_stats = self._calculate_dashboard_fertility_stats()
            
            # 月ごとの受胎率を計算
            monthly_fertility_stats = self._calculate_dashboard_monthly_fertility_stats()
            
            # 授精回数ごとの受胎率を計算
            insemination_count_fertility_stats = self._calculate_dashboard_insemination_count_fertility_stats()
            
            # 牛群動態（分娩予定月×産子種類）を計算
            herd_dynamics_data = None
            if self.farm_path:
                try:
                    from modules.herd_dynamics_report import build_herd_dynamics_data
                    herd_dynamics_data = build_herd_dynamics_data(self.db, self.formula_engine, self.farm_path)
                except Exception as e:
                    logging.debug(f"ダッシュボード: 牛群動態データ取得スキップ: {e}")
            
            # HTML生成
            html_content = self._build_dashboard_html(
                farm_name,
                parity_segments,
                avg_parity,
                total,
                milk_stats,
                fertility_stats,
                monthly_fertility_stats,
                herd_summary,
                scc_summary,
                scatter_data,
                insemination_count_fertility_stats,
                herd_dynamics_data,
            )
            
            # 一時HTMLファイルを作成
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.html', delete=False) as f:
                f.write(html_content)
                temp_file_path = f.name
            
            # ブラウザで開く
            webbrowser.open(f'file://{temp_file_path}')
        except Exception as e:
            logging.error(f"ダッシュボード表示エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ダッシュボードの準備中にエラーが発生しました: {e}")
    
    def _calculate_dashboard_milk_stats(self, lactated_cows):
        """平均乳量統計を計算（群全体の直近の乳検日の平均）"""
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "avg_milk": None,
                "avg_first_parity": None,
                "avg_multiparous": None,
                "latest_date": None,
                "milk_bins": {},
                "milk_yields": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "avg_milk": None,
                "avg_first_parity": None,
                "avg_multiparous": None,
                "latest_date": None,
                "milk_bins": {},
                "milk_yields": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        # 乳量データを取得
        milk_yields = []
        first_parity_milk_yields = []
        multiparous_milk_yields = []
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得（産次を確認するため）
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 経産牛（LACT>=1）のみを対象
            lact = cow.get('lact', 0)
            if lact < 1:
                continue
            
            # json_dataから乳量を取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                milk_yield = json_data.get('milk_yield') or json_data.get('milk_kg')
                if milk_yield is not None:
                    try:
                        milk_value = float(milk_yield)
                        if milk_value > 0:
                            milk_yields.append(milk_value)
                            
                            # 産次で分類
                            if lact == 1:
                                first_parity_milk_yields.append(milk_value)
                            elif lact >= 2:
                                multiparous_milk_yields.append(milk_value)
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        # 平均乳量を計算
        avg_milk = sum(milk_yields) / len(milk_yields) if milk_yields else None
        avg_first_parity = sum(first_parity_milk_yields) / len(first_parity_milk_yields) if first_parity_milk_yields else None
        avg_multiparous = sum(multiparous_milk_yields) / len(multiparous_milk_yields) if multiparous_milk_yields else None
        
        # 乳量階層を計算（~20kg, 20kg台, 30kg台, 40kg台, ~50kg）
        milk_bins = {
            "~20kg": 0,
            "20kg台": 0,
            "30kg台": 0,
            "40kg台": 0,
            "~50kg": 0
        }
        
        for milk_value in milk_yields:
            if milk_value < 20:
                milk_bins["~20kg"] += 1
            elif 20 <= milk_value < 30:
                milk_bins["20kg台"] += 1
            elif 30 <= milk_value < 40:
                milk_bins["30kg台"] += 1
            elif 40 <= milk_value < 50:
                milk_bins["40kg台"] += 1
            else:  # >= 50
                milk_bins["~50kg"] += 1
        
        return {
            "avg_milk": avg_milk,
            "avg_first_parity": avg_first_parity,
            "avg_multiparous": avg_multiparous,
            "latest_date": latest_milk_test_date,
            "milk_bins": milk_bins,
            "milk_yields": milk_yields
        }
    
    def _filter_existing_cows(self, cow_rows: List) -> List:
        """
        現存牛のみをフィルタリング（集計・リスト・グラフコマンド用）
        
        Args:
            cow_rows: 牛のデータ行のリスト（auto_idを含む辞書またはタプル）
            
        Returns:
            現存牛のみのリスト
        """
        existing_cows = []
        for cow_row in cow_rows:
            # cow_rowが辞書の場合
            if isinstance(cow_row, dict):
                cow_auto_id = cow_row.get('auto_id')
            # cow_rowがタプルやRowオブジェクトの場合
            else:
                # タプルの場合、最初の要素がauto_idと仮定
                try:
                    cow_auto_id = cow_row[0] if hasattr(cow_row, '__getitem__') else None
                except (IndexError, TypeError):
                    continue
            
            if cow_auto_id and not self._is_cow_disposed_for_dashboard(cow_auto_id):
                existing_cows.append(cow_row)
        
        return existing_cows
    
    def _calculate_dashboard_fertility_stats(self):
        """産次ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（産次別）
            result = self._calculate_conception_rate(
                self.db,
                "産次",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_dashboard_monthly_fertility_stats(self):
        """月ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（月別）
            result = self._calculate_conception_rate(
                self.db,
                "月",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"月ごと受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_dashboard_insemination_count_fertility_stats(self):
        """授精回数ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（授精回数別）
            result = self._calculate_conception_rate(
                self.db,
                "授精回数",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"授精回数ごと受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_herd_summary(self, lactated_cows):
        """牛群要約を計算（平均分娩後日数、妊娠牛の割合、分娩間隔、予定分娩間隔）"""
        from datetime import datetime, timedelta
        
        try:
            dim_values = []
            pregnant_count = 0
            total_count = len(lactated_cows)
            cci_values = []  # 分娩間隔
            pcci_values = []  # 予定分娩間隔
            
            today = datetime.now()
            
            for cow in lactated_cows:
                cow_auto_id = cow.get('auto_id')
                if not cow_auto_id:
                    continue
                
                # 繁殖コードを取得
                rc = cow.get('rc')
                if rc in [5, 6]:  # 5: Pregnant, 6: Dry
                    pregnant_count += 1
                
                # 分娩日を取得してDIMを計算
                clvd = cow.get('clvd')
                if clvd:
                    try:
                        clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
                        dim = (today - clvd_dt).days
                        if dim >= 0:  # 分娩日が今日以前の場合のみ
                            dim_values.append(dim)
                    except (ValueError, TypeError):
                        pass
                
                # 分娩間隔（CCI）を計算
                # 分娩イベントが2回以上ある場合のみ計算
                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                calving_events = [e for e in events if e.get('event_number') == self.rule_engine.EVENT_CALV]
                calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
                
                if len(calving_events) >= 2:
                    try:
                        prev_calv_date = calving_events[-2].get('event_date')
                        last_calv_date = calving_events[-1].get('event_date')
                        if prev_calv_date and last_calv_date:
                            prev_dt = datetime.strptime(prev_calv_date, '%Y-%m-%d')
                            last_dt = datetime.strptime(last_calv_date, '%Y-%m-%d')
                            cci = (last_dt - prev_dt).days
                            if cci > 0:
                                cci_values.append(cci)
                    except (ValueError, TypeError):
                        pass
                
                # 予定分娩間隔（PCCI）を計算
                # FormulaEngineのcalculateメソッドを使用（集計コマンドと同じロジック）
                if self.formula_engine:
                    try:
                        calculated = self.formula_engine.calculate(cow_auto_id, 'PCI')
                        if calculated and 'PCI' in calculated:
                            pcci = calculated['PCI']
                            if pcci is not None and isinstance(pcci, (int, float)) and pcci > 0:
                                pcci_values.append(int(pcci))
                    except Exception as e:
                        logging.debug(f"予定分娩間隔計算エラー: cow_auto_id={cow_auto_id}, error={e}")
                        pass
            
            # 平均分娩後日数を計算
            avg_dim = sum(dim_values) / len(dim_values) if dim_values else None
            
            # 妊娠牛の割合を計算
            pregnancy_rate = (pregnant_count / total_count * 100) if total_count > 0 else 0
            
            # 平均分娩間隔を計算
            avg_cci = sum(cci_values) / len(cci_values) if cci_values else None
            
            # 平均予定分娩間隔を計算
            avg_pcci = sum(pcci_values) / len(pcci_values) if pcci_values else None
            
            return {
                "avg_dim": avg_dim,
                "pregnancy_rate": pregnancy_rate,
                "pregnant_count": pregnant_count,
                "total_count": total_count,
                "avg_cci": avg_cci,
                "avg_pcci": avg_pcci
            }
        except Exception as e:
            logging.error(f"牛群要約計算エラー: {e}", exc_info=True)
            return {
                "avg_dim": None,
                "pregnancy_rate": 0,
                "pregnant_count": 0,
                "total_count": 0,
                "avg_cci": None,
                "avg_pcci": None
            }
    
    def _calculate_dashboard_scc_stats(self, lactated_cows):
        """体細胞要約を計算（リニアスコア階層、平均体細胞、平均リニアスコア）"""
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "avg_scc": None,
                "avg_ls": None,
                "avg_first_parity_ls": None,
                "avg_multiparous_ls": None,
                "latest_date": None,
                "ls_bins": {},
                "ls_values": [],
                "scc_values": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "avg_scc": None,
                "avg_ls": None,
                "avg_first_parity_ls": None,
                "avg_multiparous_ls": None,
                "latest_date": None,
                "ls_bins": {},
                "ls_values": [],
                "scc_values": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        # 体細胞とリニアスコアデータを取得
        scc_values = []
        ls_values = []
        first_parity_ls_values = []
        multiparous_ls_values = []
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得（産次を確認するため）
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 経産牛（LACT>=1）のみを対象
            lact = cow.get('lact', 0)
            if lact < 1:
                continue
            
            # json_dataから体細胞とリニアスコアを取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                scc_value = json_data.get('scc')
                ls_value = json_data.get('ls')
                
                if scc_value is not None:
                    try:
                        scc_num = float(scc_value)
                        if scc_num > 0:
                            scc_values.append(scc_num)
                    except (ValueError, TypeError):
                        pass
                
                if ls_value is not None:
                    try:
                        ls_num = float(ls_value)
                        if ls_num >= 0:
                            ls_values.append(ls_num)
                            
                            # 産次で分類
                            if lact == 1:
                                first_parity_ls_values.append(ls_num)
                            elif lact >= 2:
                                multiparous_ls_values.append(ls_num)
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        # 平均を計算
        avg_scc = sum(scc_values) / len(scc_values) if scc_values else None
        avg_ls = sum(ls_values) / len(ls_values) if ls_values else None
        avg_first_parity_ls = sum(first_parity_ls_values) / len(first_parity_ls_values) if first_parity_ls_values else None
        avg_multiparous_ls = sum(multiparous_ls_values) / len(multiparous_ls_values) if multiparous_ls_values else None
        
        # リニアスコア階層を計算（LS2以下、LS3-4、LS5以上）
        ls_bins = {
            "LS2以下": 0,
            "LS3-4": 0,
            "LS5以上": 0
        }
        
        for ls_value in ls_values:
            if ls_value <= 2:
                ls_bins["LS2以下"] += 1
            elif 3 <= ls_value <= 4:
                ls_bins["LS3-4"] += 1
            else:  # >= 5
                ls_bins["LS5以上"] += 1
        
        return {
            "avg_scc": avg_scc,
            "avg_ls": avg_ls,
            "avg_first_parity_ls": avg_first_parity_ls,
            "avg_multiparous_ls": avg_multiparous_ls,
            "latest_date": latest_milk_test_date,
            "ls_bins": ls_bins,
            "ls_values": ls_values,
            "scc_values": scc_values
        }
    
    def _calculate_dashboard_scatter_data(self, lactated_cows):
        """散布図データを計算（検定時のDIM、産次を含む）"""
        from datetime import datetime
        
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "milk_points": [],
                "ls_points": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "milk_points": [],
                "ls_points": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        milk_points = []  # (dim, milk_yield, parity)
        ls_points = []    # (dim, ls, parity)
        
        def calc_dim_on_date(clvd_date: str, event_date: str) -> int:
            """検定時のDIMを計算"""
            if not clvd_date:
                return None
            try:
                clvd_dt = datetime.strptime(clvd_date, "%Y-%m-%d")
                event_dt = datetime.strptime(event_date, "%Y-%m-%d")
            except (ValueError, TypeError):
                return None
            if event_dt < clvd_dt:
                return None
            return (event_dt - clvd_dt).days
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 経産牛（LACT>=1）のみを対象
            lact = cow.get('lact', 0)
            if lact < 1:
                continue
            
            # 産次を分類（1産、2産、3産以上）
            if lact == 1:
                parity = 1
            elif lact == 2:
                parity = 2
            else:  # lact >= 3
                parity = 3
            
            # 検定時のDIMを計算
            clvd = cow.get('clvd')
            dim = calc_dim_on_date(clvd, latest_milk_test_date)
            if dim is None:
                continue
            
            # json_dataから乳量とリニアスコアを取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                
                milk_yield = json_data.get('milk_yield')
                if milk_yield is not None:
                    try:
                        milk_value = float(milk_yield)
                        if milk_value > 0:
                            cow_id = str(cow.get("cow_id") or "")
                            jpn10 = str(cow.get("jpn10") or "")
                            milk_points.append((dim, milk_value, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass
                
                ls_value = json_data.get('ls')
                if ls_value is not None:
                    try:
                        ls_num = float(ls_value)
                        if ls_num >= 0:
                            cow_id = str(cow.get("cow_id") or "")
                            jpn10 = str(cow.get("jpn10") or "")
                            ls_points.append((dim, ls_num, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        return {
            "milk_points": milk_points,
            "ls_points": ls_points
        }
    
    def _build_dashboard_html(self, farm_name, parity_segments, avg_parity, total, milk_stats, fertility_stats, monthly_fertility_stats, herd_summary, scc_summary, scatter_data, insemination_count_fertility_stats, herd_dynamics_data=None):
        """ダッシュボードのHTMLを生成"""
        # 乳量・リニアスコア散布図（乳検レポートと同じ仕様：Plotly、ホバーでID表示・2グラフ連動、x_min/x_max で DIM 軸を固定）
        scatter_plotly_html = ""
        if _get_plotly_scatter_script and (scatter_data.get("milk_points") or scatter_data.get("ls_points")):
            def _unpack_pt(p):
                if len(p) >= 5:
                    return (p[0], p[1], p[2], str(p[3]), str(p[4]))
                return (p[0], p[1], p[2], "", "")
            configs = []
            milk_points = scatter_data.get("milk_points", [])
            ls_points = scatter_data.get("ls_points", [])
            if milk_points:
                m_pts = [_unpack_pt(p) for p in milk_points]
                configs.append({
                    "id": "scatter-dash-milk",
                    "data": {
                        "x": [int(p[0]) for p in m_pts],
                        "y": [float(p[1]) for p in m_pts],
                        "parity": [p[2] for p in m_pts],
                        "cow_ids": [p[3] for p in m_pts],
                        "jpn10s": [p[4] for p in m_pts],
                        "title": "乳量",
                        "xlabel": "DIM",
                        "ylabel": "乳量",
                        "x_min": 0,
                        "x_max": 400,
                        "y_start_zero": True,
                    },
                })
            if ls_points:
                l_pts = [_unpack_pt(p) for p in ls_points]
                configs.append({
                    "id": "scatter-dash-ls",
                    "data": {
                        "x": [int(p[0]) for p in l_pts],
                        "y": [float(p[1]) for p in l_pts],
                        "parity": [p[2] for p in l_pts],
                        "cow_ids": [p[3] for p in l_pts],
                        "jpn10s": [p[4] for p in l_pts],
                        "title": "リニアスコア",
                        "xlabel": "DIM",
                        "ylabel": "リニアスコア",
                        "x_min": 0,
                        "x_max": 400,
                        "y_start_zero": True,
                    },
                })
            if configs:
                config_json = json.dumps(configs, ensure_ascii=False).replace("</", "<\\/")
                script = _get_plotly_scatter_script(config_json)
                parts = []
                for c in configs:
                    tid = c["id"]
                    title = "乳量" if "milk" in tid else "リニアスコア"
                    parts.append(
                        f'<div class="scatter-card">'
                        f'<div class="chart-title">{html.escape(title)}</div>'
                        f'<div class="scatter-chart"><div id="{html.escape(tid)}" class="plotly-scatter-div" style="min-height:260px;"></div></div>'
                        f'<div class="scatter-legend">'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                        f'</div></div>'
                    )
                scatter_plotly_html = '<div class="scatter-stack">' + "".join(parts) + '</div>\n' + script
        if not scatter_plotly_html and (scatter_data.get("milk_points") or scatter_data.get("ls_points")):
            milk_pts = scatter_data.get("milk_points", [])
            ls_pts = scatter_data.get("ls_points", [])
            scatter_plotly_html = (
                '<div class="scatter-stack">'
                '<div class="scatter-card"><div class="chart-title">乳量</div><div class="scatter-chart">'
                + self._build_milk_scatter_svg([(p[0], p[1], p[2]) for p in milk_pts])
                + '</div><div class="scatter-legend">'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                '</div></div><div class="scatter-card"><div class="chart-title">リニアスコア</div><div class="scatter-chart">'
                + self._build_ls_scatter_svg([(p[0], p[1], p[2]) for p in ls_pts])
                + '</div><div class="scatter-legend">'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                '</div></div></div>'
            )

        # 産次ドーナツチャート
        parity_donut_svg = self._build_parity_donut_svg(parity_segments, avg_parity, total)
        parity_legend = self._build_parity_legend(parity_segments, total)
        
        # 平均乳量統計と乳量階層ドーナツチャート
        milk_stats_html = ""
        if milk_stats["latest_date"]:
            # 乳量階層のドーナツチャート用データを準備
            milk_segments = []
            milk_bins = milk_stats.get("milk_bins", {})
            milk_yields = milk_stats.get("milk_yields", [])
            total_milk_count = len(milk_yields)
            
            if total_milk_count > 0:
                # 乳量階層の色定義（モダンなカラーパレット）
                milk_colors = {
                    "~20kg": "#FFA07A",      # ライトサルモン
                    "20kg台": "#98D8C8",     # ミントグリーン
                    "30kg台": "#6C5CE7",     # パープル
                    "40kg台": "#A29BFE",     # ライトパープル
                    "~50kg": "#FD79A8"       # ピンク
                }
                
                # 階層の順序
                milk_order = ["~20kg", "20kg台", "30kg台", "40kg台", "~50kg"]
                radius = 50
                circumference = 2 * 3.141592653589793 * radius
                offset = 0
                
                for label in milk_order:
                    count = milk_bins.get(label, 0)
                    if count > 0:
                        percent = (count / total_milk_count * 100) if total_milk_count > 0 else 0
                        length = (percent / 100) * circumference
                        milk_segments.append({
                            "label": label,
                            "count": count,
                            "percent": round(percent, 1),
                            "color": milk_colors.get(label, "#999999"),
                            "length": length,
                            "offset": -offset
                        })
                        offset += length
                
                # 乳量階層ドーナツチャートと凡例を生成
                milk_donut_svg = self._build_milk_donut_svg(milk_segments, milk_stats["avg_milk"], total_milk_count)
                milk_legend = self._build_milk_legend(milk_segments, total_milk_count)
            else:
                milk_donut_svg = '<div class="subnote">データなし</div>'
                milk_legend = ""
            
            milk_stats_html = f"""
            <div class="milk-section">
                <div class="section-title">乳量要約（{milk_stats["latest_date"]}）</div>
                <div class="milk-chart-container">
                    <div class="milk-donut">
                        {milk_donut_svg}
                    </div>
                    <div class="legend">
                        {milk_legend}
                    </div>
                </div>
            <div class="milk-stats">
                <div class="stats-label"><span class="stats-label-text">平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_milk"]:.1f}kg</span></div>
                    {f'<div class="stats-label"><span class="stats-label-text">初産平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_first_parity"]:.1f}kg</span></div>' if milk_stats.get("avg_first_parity") else ''}
                    {f'<div class="stats-label"><span class="stats-label-text">2産以上平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_multiparous"]:.1f}kg</span></div>' if milk_stats.get("avg_multiparous") else ''}
                </div>
            </div>
            """
        else:
            milk_stats_html = '<div class="milk-section"><div class="section-title">乳量要約</div><div class="milk-stats"><div class="stats-label">乳検データがありません。</div></div></div>'
        
        # 受胎率表
        fertility_table_html = ""
        if fertility_stats and fertility_stats.get("stats"):
            stats = fertility_stats["stats"]
            start_date = fertility_stats["start_date"]
            end_date = fertility_stats["end_date"]
            
            # 分類をソート（産次は数値順）
            def sort_key(classification):
                if classification == "合計":
                    return (2, 0)
                try:
                    return (0, int(classification))
                except (ValueError, TypeError):
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            table_rows = []
            for classification in sorted_classifications:
                stat = stats[classification]
                total_count = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(str(classification))}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">産次ごとの受胎率（経産）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>産次</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            fertility_table_html = '<div class="fertility-section"><div class="subnote">受胎率データがありません。</div></div>'
        
        # 月ごとの受胎率表
        monthly_fertility_table_html = ""
        if monthly_fertility_stats and monthly_fertility_stats.get("stats"):
            stats = monthly_fertility_stats["stats"]
            start_date = monthly_fertility_stats["start_date"]
            end_date = monthly_fertility_stats["end_date"]
            
            # 期間内のすべての月を生成
            from datetime import datetime, timedelta
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            
            # 開始月から終了月までのすべての月を生成
            all_months = []
            current_dt = datetime(start_dt.year, start_dt.month, 1)
            end_month_dt = datetime(end_dt.year, end_dt.month, 1)
            
            while current_dt <= end_month_dt:
                month_str = current_dt.strftime('%Y-%m')
                all_months.append(month_str)
                # 次の月へ
                if current_dt.month == 12:
                    current_dt = datetime(current_dt.year + 1, 1, 1)
                else:
                    current_dt = datetime(current_dt.year, current_dt.month + 1, 1)
            
            table_rows = []
            for month_str in all_months:
                # データがある場合はそのデータを使用、ない場合は0で初期化
                if month_str in stats:
                    stat = stats[month_str]
                    total_count = stat['total']
                    conceived = stat['conceived']
                    not_conceived = stat['not_conceived']
                    other = stat['other']
                else:
                    total_count = 0
                    conceived = 0
                    not_conceived = 0
                    other = 0
                
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(month_str)}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            monthly_fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">月ごと受胎率（経産）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>授精月</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            monthly_fertility_table_html = '<div class="fertility-section"><div class="subnote">月ごと受胎率データがありません。</div></div>'
        
        # 授精回数ごとの受胎率表
        insemination_count_fertility_table_html = ""
        if insemination_count_fertility_stats and insemination_count_fertility_stats.get("stats"):
            stats = insemination_count_fertility_stats["stats"]
            start_date = insemination_count_fertility_stats["start_date"]
            end_date = insemination_count_fertility_stats["end_date"]
            
            # 分類をソート（授精回数は数値順）
            def sort_key(classification):
                if classification == "合計":
                    return (2, 0)
                try:
                    return (0, int(classification))
                except (ValueError, TypeError):
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            table_rows = []
            for classification in sorted_classifications:
                stat = stats[classification]
                total_count = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(str(classification))}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            insemination_count_fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">授精回数ごとの受胎率（経産）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>授精回数</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            insemination_count_fertility_table_html = '<div class="fertility-section"><div class="subnote">授精回数ごと受胎率データがありません。</div></div>'
        
        # 体細胞要約
        scc_summary_html = ""
        if scc_summary and scc_summary.get("latest_date"):
            # リニアスコア階層のドーナツチャート用データを準備
            ls_segments = []
            ls_bins = scc_summary.get("ls_bins", {})
            ls_values = scc_summary.get("ls_values", [])
            total_ls_count = len(ls_values)
            
            if total_ls_count > 0:
                # リニアスコア階層の色定義（モダンなカラーパレット）
                ls_colors = {
                    "LS2以下": "#4169E1",      # ブルー
                    "LS3-4": "#FFD700",        # イエロー
                    "LS5以上": "#FF6B6B"        # レッド
                }
                
                # 階層の順序
                ls_order = ["LS2以下", "LS3-4", "LS5以上"]
                radius = 50
                circumference = 2 * 3.141592653589793 * radius
                offset = 0
                
                for label in ls_order:
                    count = ls_bins.get(label, 0)
                    if count > 0:
                        percent = (count / total_ls_count * 100) if total_ls_count > 0 else 0
                        length = (percent / 100) * circumference
                        ls_segments.append({
                            "label": label,
                            "count": count,
                            "percent": round(percent, 1),
                            "color": ls_colors.get(label, "#999999"),
                            "length": length,
                            "offset": -offset
                        })
                        offset += length
                
                # リニアスコア階層ドーナツチャートと凡例を生成
                ls_donut_svg = self._build_ls_donut_svg(ls_segments, scc_summary["avg_ls"], total_ls_count)
                ls_legend = self._build_ls_legend(ls_segments, total_ls_count)
            else:
                ls_donut_svg = '<div class="subnote">データなし</div>'
                ls_legend = ""
            
            # 平均体細胞を計算（千単位で表示）
            import math
            avg_scc_display = None
            if scc_summary.get("avg_scc") is not None:
                avg_scc_display = int(math.floor(scc_summary["avg_scc"] + 0.5))
            
            scc_summary_html = f"""
            <div class="scc-section">
                <div class="section-title">体細胞要約（{scc_summary["latest_date"]}）</div>
                <div class="milk-chart-container">
                    <div class="milk-donut">
                        {ls_donut_svg}
                    </div>
                    <div class="legend">
                        {ls_legend}
                    </div>
                </div>
                <div class="milk-stats">
                    <div class="stats-label"><span class="stats-label-text">平均体細胞:</span> <span class="stats-label-value">{f"{avg_scc_display}千" if avg_scc_display is not None else "-"}</span></div>
                    <div class="stats-label"><span class="stats-label-text">平均リニアスコア:</span> <span class="stats-label-value">{f"{scc_summary['avg_ls']:.1f}" if scc_summary.get("avg_ls") is not None else "-"}</span></div>
                    {f'<div class="stats-label"><span class="stats-label-text">初産平均リニアスコア:</span> <span class="stats-label-value">{scc_summary["avg_first_parity_ls"]:.1f}</span></div>' if scc_summary.get("avg_first_parity_ls") is not None else ''}
                    {f'<div class="stats-label"><span class="stats-label-text">2産以上平均リニアスコア:</span> <span class="stats-label-value">{scc_summary["avg_multiparous_ls"]:.1f}</span></div>' if scc_summary.get("avg_multiparous_ls") is not None else ''}
                </div>
            </div>
            """
        else:
            scc_summary_html = '<div class="scc-section"><div class="section-title">体細胞要約</div><div class="milk-stats"><div class="stats-label">乳検データがありません。</div></div></div>'
        
        # 分娩予定月別・産子種類内訳グラフ（牛群動態）
        herd_dynamics_chart_html = self._build_herd_dynamics_chart_section(herd_dynamics_data)
        
        try:
            from modules.report_cow_bridge import DEFAULT_PORT as _report_open_cow_port
        except ImportError:
            _report_open_cow_port = 51985
        html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ダッシュボード</title>
    <style>
        /* ゲノム・乳検レポートと統一（フォント・配色・セクション見出し）。ダッシュボード用に可読性は維持 */
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Meiryo', 'Yu Gothic', sans-serif;
            font-size: 14px;
            margin: 20px;
            padding: 24px;
            background: #f8f9fa;
            color: #212529;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .dashboard-container {{
            max-width: 1400px;
            margin: 0 auto;
            background: #fff;
            padding: 32px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .header {{
            font-size: 1.4rem;
            font-weight: 600;
            margin-bottom: 28px;
            padding-bottom: 12px;
            border-bottom: 2px solid #0d6efd;
            color: #0d6efd;
        }}
        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
            margin-top: 20px;
        }}
        .dashboard-item {{
            display: flex;
            flex-direction: column;
            background: #fff;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
            transition: box-shadow 0.2s ease;
        }}
        .scatter-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 24px;
            margin-top: 0;
        }}
        .scatter-stack {{
            display: flex;
            flex-direction: row;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 16px;
            grid-column: 1 / -1;
            max-width: 100%;
            box-sizing: border-box;
        }}
        .scatter-stack .scatter-card {{
            flex: 1;
            min-width: 380px;
        }}
        .scatter-card {{
            border: 1px solid #dee2e6;
            border-radius: 6px;
            padding: 16px;
            background: #fff;
            overflow: hidden;
            box-sizing: border-box;
        }}
        .scatter-card .chart-title {{
            margin: 0 0 8px;
            font-size: 0.95rem;
            font-weight: 600;
            color: #212529;
        }}
        .scatter-card .scatter-chart {{
            width: 100%;
            max-width: 100%;
            overflow: hidden;
            box-sizing: border-box;
        }}
        .scatter-card .plotly-scatter-div {{
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box;
        }}
        .scatter-legend {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px 20px;
            margin-top: 10px;
            padding-bottom: 4px;
            font-size: 12px;
            color: #495057;
        }}
        .scatter-legend .legend-item {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}
        .scatter-legend .legend-dot {{
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            flex-shrink: 0;
        }}
        .fertility-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
            margin-top: 0;
        }}
        .herd-dynamics-row {{
            grid-column: 1 / -1;
            margin-top: 0;
        }}
        .dashboard-item:hover {{
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }}
        .parity-section {{
            margin-bottom: 0;
        }}
        .section-title {{
            font-size: 1.1rem;
            font-weight: 600;
            padding-bottom: 4px;
            border-bottom: 2px solid #0d6efd;
            margin-bottom: 20px;
            color: #0d6efd;
        }}
        .parity-chart-container {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
        }}
        .parity-donut {{
            flex-shrink: 0;
        }}
        .legend {{
            width: 100%;
        }}
        .legend-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 1px solid #e9ecef;
        }}
        .legend-row:last-child {{
            border-bottom: none;
        }}
        .legend-dot {{
            display: inline-block;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            margin-right: 10px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }}
        .legend-value {{
            display: flex;
            gap: 12px;
        }}
        .legend-count {{
            font-weight: bold;
        }}
        .legend-percent {{
            color: #666;
        }}
        .donut-center {{
            font-size: 12px;
            font-weight: bold;
            text-anchor: middle;
            fill: #333;
        }}
        .milk-section {{
            margin-bottom: 0;
        }}
        .milk-chart-container {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
        }}
        .milk-donut {{
            flex-shrink: 0;
        }}
        .milk-stats {{
            margin: 15px 0 0 0;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
        }}
        .herd-summary-stats {{
            margin-top: 20px;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
        }}
        .summary-stat-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            font-size: 14px;
        }}
        .summary-stat-item:not(:last-child) {{
            border-bottom: 1px solid #e9ecef;
        }}
        .summary-stat-label {{
            color: #495057;
            font-weight: 500;
            font-size: 14px;
        }}
        .summary-stat-value {{
            color: #212529;
            font-weight: 600;
            font-size: 14px;
        }}
        .stats-label {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            font-size: 14px;
            margin-bottom: 0;
            font-weight: normal;
        }}
        .stats-label:not(:last-child) {{
            border-bottom: 1px solid #e9ecef;
        }}
        .stats-label-text {{
            color: #495057;
            font-weight: 500;
            font-size: 14px;
        }}
        .stats-label-value {{
            color: #212529;
            font-weight: 600;
            font-size: 14px;
        }}
        .stats-note {{
            font-size: 11px;
            color: #666;
            margin-top: 6px;
        }}
        .fertility-section {{
            margin-top: 0;
        }}
        .scatter-section {{
            margin-top: 0;
        }}
        .scatter-chart {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin-top: 10px;
        }}
        .scatter-legend {{
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #e9ecef;
        }}
        .scatter-legend .legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            color: #495057;
        }}
        .scatter-legend .legend-dot {{
            display: inline-block;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            border: 1px solid rgba(0, 0, 0, 0.1);
        }}
        .subheader {{
            font-size: 12px;
            margin-bottom: 12px;
            color: #495057;
        }}
        .summary-table {{
            border-collapse: collapse;
            width: 100%;
            margin-top: 10px;
            font-size: 13px;
            border-radius: 8px;
            overflow: hidden;
        }}
        .summary-table th, .summary-table td {{
            border: 1px solid #dee2e6;
            padding: 10px 12px;
            text-align: left;
        }}
        .summary-table th {{
            background: #e9ecef;
            font-weight: 600;
            text-align: center;
            font-size: 12px;
            color: #495057;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .summary-table .num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .summary-table tbody tr:hover {{
            background: #f8f9fa;
        }}
        .total-row {{
            background: #f8f9fa;
            font-weight: 600;
        }}
        .subnote {{
            color: #6c757d;
            font-size: 11px;
            font-style: italic;
        }}
        @media print {{
            body {{
                margin: 0;
                padding: 5px;
                background: #fff;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .dashboard-container {{
                max-width: 100%;
                padding: 5px;
                box-shadow: none;
                transform: scale(0.85);
                transform-origin: top left;
                width: 117.65%;  /* 1 / 0.85 = 1.1765 */
            }}
            .header {{
                font-size: 18px;
                margin-bottom: 15px;
            }}
            .dashboard-grid {{
                grid-template-columns: repeat(3, 1fr);
                gap: 12px;
                page-break-inside: avoid;
            }}
            .dashboard-item {{
                page-break-inside: avoid;
                padding: 16px;
                box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
            }}
            .section-title {{
                font-size: 15px;
                margin-bottom: 12px;
            }}
            .parity-chart-container {{
                gap: 10px;
                margin-bottom: 10px;
            }}
            .legend-row {{
                padding: 5px 0;
                font-size: 12px;
            }}
            .milk-stats, .herd-summary-stats {{
                padding: 12px;
                margin-top: 10px;
            }}
            .stats-label, .summary-stat-item {{
                font-size: 12px;
            }}
            .summary-stat-value {{
                font-size: 13px;
            }}
            .summary-table {{
                font-size: 11px;
            }}
            .summary-table th, .summary-table td {{
                padding: 6px 8px;
            }}
            .summary-table th {{
                font-size: 10px;
            }}
            .subheader {{
                font-size: 11px;
                margin-bottom: 8px;
            }}
            .herd-dynamics-chart {{
                margin: 12px 0;
            }}
            .herd-dynamics-chart svg {{
                font-family: 'Meiryo', 'Yu Gothic', sans-serif;
            }}
            .herd-dynamics-legend {{
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 4px 16px;
                font-size: 12px;
                color: #546e7a;
                margin-top: 12px;
            }}
            .herd-dynamics-legend .legend-swatch {{
                display: inline-block;
                width: 14px;
                height: 14px;
                border: 1px solid #666;
                vertical-align: middle;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .herd-dynamics-legend .legend-label {{
                vertical-align: middle;
            }}
            @page {{
                size: portrait;
                margin: 0.5cm;
            }}
        }}
    </style>
    <script>var FALCON_OPEN_COW_PORT = {_report_open_cow_port};</script>
</head>
<body>
    <div class="dashboard-container">
        <div class="header">{html.escape(str(farm_name))}　ダッシュボード</div>
        
        <div class="dashboard-grid">
            <div class="dashboard-item">
        <div class="parity-section">
                    <div class="section-title">牛群要約</div>
            <div class="parity-chart-container">
                <div class="parity-donut">
                    {parity_donut_svg}
                </div>
                <div class="legend">
                    {parity_legend}
                </div>
            </div>
                    <div class="herd-summary-stats">
                        {f'<div class="summary-stat-item"><span class="summary-stat-label">経産牛頭数:</span> <span class="summary-stat-value">{total}頭</span></div>' if total > 0 else '<div class="summary-stat-item"><span class="summary-stat-label">経産牛頭数:</span> <span class="summary-stat-value">-</span></div>'}
                        {f'<div class="summary-stat-item"><span class="summary-stat-label">平均分娩後日数:</span> <span class="summary-stat-value">{herd_summary.get("avg_dim", 0):.1f}日</span></div>' if herd_summary.get("avg_dim") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">平均分娩後日数:</span> <span class="summary-stat-value">-</span></div>'}
                        {f'<div class="summary-stat-item"><span class="summary-stat-label">妊娠牛の割合:</span> <span class="summary-stat-value">{herd_summary.get("pregnancy_rate", 0):.1f}%</span></div>' if herd_summary.get("pregnancy_rate") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">妊娠牛の割合:</span> <span class="summary-stat-value">-</span></div>'}
                        {f'<div class="summary-stat-item"><span class="summary-stat-label">分娩間隔:</span> <span class="summary-stat-value">{herd_summary.get("avg_cci", 0):.1f}日</span></div>' if herd_summary.get("avg_cci") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">分娩間隔:</span> <span class="summary-stat-value">-</span></div>'}
                        {f'<div class="summary-stat-item"><span class="summary-stat-label">予定分娩間隔:</span> <span class="summary-stat-value">{herd_summary.get("avg_pcci", 0):.1f}日</span></div>' if herd_summary.get("avg_pcci") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">予定分娩間隔:</span> <span class="summary-stat-value">-</span></div>'}
                    </div>
                </div>
            </div>
            
            <div class="dashboard-item">
            {milk_stats_html}
        </div>
        
            <div class="dashboard-item">
                {scc_summary_html}
            </div>
            
            {scatter_plotly_html}
            
            <div class="herd-dynamics-row">
                <div class="dashboard-item">
                    <div class="herd-dynamics-section">
                        {herd_dynamics_chart_html}
                    </div>
                </div>
            </div>
            
            <div class="fertility-row">
                <div class="dashboard-item">
                    {monthly_fertility_table_html}
                </div>
                
                <div class="dashboard-item">
        {fertility_table_html}
                </div>
                
                <div class="dashboard-item">
                    {insemination_count_fertility_table_html}
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""
        return html_content
    
