"""
FALCON2 - 受胎率計算・表示 Mixin
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

from settings_manager import SettingsManager

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

logger = logging.getLogger(__name__)

class ConceptionRateMixin:
    """Mixin: FALCON2 - 受胎率計算・表示 Mixin"""
    def _execute_conception_rate_command(self):
        """受胎率コマンドを実行"""
        try:
            # 受胎率の種類を取得
            if not hasattr(self, 'conception_rate_type_entry') or not self.conception_rate_type_entry.winfo_exists():
                self.add_message(role="system", text="受胎率の種類が選択されていません")
                return
            
            rate_type = self.conception_rate_type_var.get().strip()
            if not rate_type:
                self.add_message(role="system", text="受胎率の種類を選択してください")
                return
            
            # 期間を取得
            start_date = None
            end_date = None
            if hasattr(self, 'start_date_entry') and hasattr(self, 'end_date_entry'):
                start_date_str = self.start_date_entry.get().strip()
                end_date_str = self.end_date_entry.get().strip()
                if start_date_str and start_date_str != "YYYY-MM-DD":
                    start_date = start_date_str
                if end_date_str and end_date_str != "YYYY-MM-DD":
                    end_date = end_date_str
            
            # 条件を取得
            condition_text = ""
            if hasattr(self, 'conception_rate_condition_entry') and self.conception_rate_condition_entry.winfo_exists():
                condition_text = self.conception_rate_condition_entry.get().strip()
                if condition_text == self.conception_rate_condition_placeholder:
                    condition_text = ""
            
            # ワーカースレッドで計算を実行
            self.result_notebook.select(0)  # 表タブを選択
            
            def calculate_in_thread():
                try:
                    # self.dbを直接使用（既に初期化済み）
                    result = self._calculate_conception_rate(
                        self.db, rate_type, start_date, end_date, condition_text
                    )
                    if result:
                        # 期間情報を文字列化
                        period_text = f"{result.get('start_date', '')} ～ {result.get('end_date', '')}"
                        self.root.after(0, lambda: self._display_conception_rate_result(result, rate_type, period_text))
                    else:
                        self.root.after(0, lambda: self.add_message(role="system", text="受胎率の計算に失敗しました"))
                except Exception as e:
                    error_msg = str(e)  # 例外メッセージを変数に保存
                    logging.error(f"受胎率計算エラー: {error_msg}", exc_info=True)
                    self.root.after(0, lambda msg=error_msg: self.add_message(role="system", text=f"受胎率の計算中にエラーが発生しました: {msg}"))
            
            import threading
            thread = threading.Thread(target=calculate_in_thread, daemon=True)
            thread.start()
            
        except Exception as e:
            logging.error(f"受胎率コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"受胎率コマンドの実行中にエラーが発生しました: {e}")
    
    def _evaluate_cow_condition_for_conception_rate(self, cow: Dict[str, Any], condition_text: str, db) -> bool:
        """
        受胎率計算用の条件評価
        
        Args:
            cow: 牛データ
            condition_text: 条件文字列（例：「DIM<150」「産次：初産」「フレッシュチェックNOTE>=MET」）
            db: DBHandlerインスタンス
        
        Returns:
            条件を満たす場合はTrue、満たさない場合はFalse
        """
        if not condition_text:
            return True
        
        try:
            import re
            import json
            
            # 表示名・項目キー解決用マッピングを構築（item_dictionary から）
            display_name_to_item_key = {}
            item_key_lower_to_item_key = {}
            item_dict = getattr(self, 'item_dictionary', None) or {}
            if not item_dict and getattr(self, 'item_dict_path', None):
                path = self.item_dict_path
                if path and getattr(path, 'exists', lambda: False) and path.exists():
                    try:
                        with open(path, 'r', encoding='utf-8') as f:
                            item_dict = json.load(f)
                    except Exception:
                        item_dict = {}
            for item_key, item_def in (item_dict or {}).items():
                display_name = (item_def or {}).get("display_name", "")
                if display_name:
                    display_name_to_item_key[display_name] = item_key
                    item_key_lower_to_item_key[item_key.lower()] = item_key
                alias = (item_def or {}).get("alias", "")
                if alias:
                    display_name_to_item_key[alias] = item_key
                    item_key_lower_to_item_key[alias.lower()] = item_key
            
            # 条件テキストをスペースで分割（複数条件に対応）
            condition_parts = condition_text.replace("　", " ").split()
            
            # すべての条件を満たす必要がある
            for part in condition_parts:
                part = part.strip()
                if not part:
                    continue
                
                # キーワード「受胎」「空胎」
                if part == "受胎":
                    rc = cow.get('rc')
                    if rc is None and cow.get('auto_id'):
                        try:
                            state = self.rule_engine.apply_events(cow['auto_id'])
                            rc = state.get('rc') if state else None
                        except Exception:
                            pass
                    try:
                        rc_int = int(rc) if rc is not None else None
                    except (ValueError, TypeError):
                        rc_int = None
                    if rc_int not in (5, 6):
                        return False
                    continue
                if part == "空胎":
                    rc = cow.get('rc')
                    if rc is None and cow.get('auto_id'):
                        try:
                            state = self.rule_engine.apply_events(cow['auto_id'])
                            rc = state.get('rc') if state else None
                        except Exception:
                            pass
                    try:
                        rc_int = int(rc) if rc is not None else None
                    except (ValueError, TypeError):
                        rc_int = None
                    if rc_int not in (1, 2, 3, 4):
                        return False
                    continue
                
                # 「項目名：値」の形式（例：産次：初産）
                if "：" in part or ":" in part:
                    separator = "：" if "：" in part else ":"
                    cond_parts = part.split(separator, 1)
                    if len(cond_parts) == 2:
                        item_name = cond_parts[0].strip()
                        value_str = cond_parts[1].strip()
                        
                        # 産次の特別処理
                        if item_name.lower() in ["lact", "産次"]:
                            cow_lact = cow.get('lact')
                            if cow_lact is None:
                                return False
                            
                            # 値が「初産」などの文字列の場合
                            if value_str in ["初産", "1", "1産"]:
                                if cow_lact != 1:
                                    return False
                            elif value_str in ["経産", "2産以上"]:
                                if cow_lact < 2:
                                    return False
                            else:
                                try:
                                    lact_value = int(value_str)
                                    if cow_lact != lact_value:
                                        return False
                                except ValueError:
                                    # 数値変換に失敗した場合は文字列比較
                                    if str(cow_lact) != value_str:
                                        return False
                            continue
                
                # 比較演算子を含む形式（項目名は日本語可。例：DIM>150, フレッシュチェックNOTE>=MET）
                match = re.match(r'^(.+?)\s*(>=|<=|!=|==|=|<|>)\s*(.*)$', part)
                if match:
                    item_name = match.group(1).strip()
                    operator = match.group(2)
                    value_str = match.group(3).strip()
                    
                    # 項目名を item_key に解決（表示名・別名・項目キーいずれでも可）
                    item_key = None
                    if item_name in item_dict:
                        item_key = item_name
                    elif item_name in display_name_to_item_key:
                        item_key = display_name_to_item_key[item_name]
                    elif item_name.lower() in item_key_lower_to_item_key:
                        item_key = item_key_lower_to_item_key[item_name.lower()]
                    else:
                        item_key = item_name
                    
                    # 項目の値を取得
                    row_value = None
                    if item_key.upper() == "LACT":
                        row_value = cow.get('lact')
                    else:
                        cow_auto_id = cow.get('auto_id')
                        if cow_auto_id and getattr(self, 'formula_engine', None):
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            if calculated:
                                row_value = calculated.get(item_key) or calculated.get(item_key.upper())
                    
                    if row_value is None:
                        return False
                    
                    # 条件を評価（数値なら比較演算子、文字列なら >= は部分一致）
                    try:
                        row_value_num = float(row_value)
                        condition_value_num = float(value_str)
                        
                        condition_match = False
                        if operator == "<":
                            condition_match = row_value_num < condition_value_num
                        elif operator == ">":
                            condition_match = row_value_num > condition_value_num
                        elif operator == "<=":
                            condition_match = row_value_num <= condition_value_num
                        elif operator == ">=":
                            condition_match = row_value_num >= condition_value_num
                        elif operator == "=" or operator == "==":
                            condition_match = row_value_num == condition_value_num
                        elif operator == "!=" or operator == "<>":
                            condition_match = row_value_num != condition_value_num
                        
                        if not condition_match:
                            return False
                    except (ValueError, TypeError):
                        # 数値でない場合は文字列比較。>= は部分一致（例：フレッシュチェックNOTE>=MET）
                        condition_match = False
                        if operator == "=" or operator == "==":
                            condition_match = (str(row_value).strip() == str(value_str).strip())
                        elif operator == "!=" or operator == "<>":
                            condition_match = (str(row_value).strip() != str(value_str).strip())
                        elif operator == ">=":
                            condition_match = str(value_str).strip() in str(row_value)
                        elif operator == "<=":
                            condition_match = str(value_str).strip() in str(row_value)
                        
                        if not condition_match:
                            return False
                    continue
            
            return True
            
        except Exception as e:
            logging.error(f"条件評価エラー: {e}", exc_info=True)
            # エラーが発生した場合は条件を満たさないと判定
            return False
    
    def _get_pregnancy_check_days(self) -> int:
        """農場設定から妊娠鑑定日数（授精後何日で鑑定するか）を取得。デフォルト35日。"""
        default = 35
        if not getattr(self, 'farm_path', None) or not self.farm_path:
            return default
        try:
            sm = SettingsManager(self.farm_path)
            for t in sm.get_repro_templates():
                for c in t.get('conditions') or []:
                    if c.get('type') == 'preg_check_days' and c.get('enabled', True):
                        v = c.get('value')
                        if v is not None:
                            return int(v)
        except Exception:
            pass
        for c in SettingsManager.get_default_template_conditions():
            if c.get('type') == 'preg_check_days':
                return int(c.get('value', default))
        return default

    def _is_insemination_type_r(self, ai_type_code: str) -> bool:
        """授精種類がR（連続授精）に該当するか。コードが'R'または農場設定の表示名がR/連続授精ならTrue。＊対象外の①。"""
        if not ai_type_code:
            return False
        raw = str(ai_type_code).strip()
        # 「コード：名称」形式の場合はコード部と名称部に分割
        code_part, name_part = raw, ""
        for sep in ("：", ":"):
            if sep in raw:
                parts = raw.split(sep, 1)
                if len(parts) >= 2:
                    code_part = (parts[0] or "").strip()
                    name_part = (parts[1] or "").strip()
                break
        s = code_part.upper()
        # コードがR、または名称がR/連続授精なら①
        if s == 'R':
            return True
        if name_part:
            np = name_part.upper()
            if np == 'R' or '連続授精' in name_part or '連続' in name_part:
                return True
        if not getattr(self, 'farm_path', None) or not self.farm_path:
            return False
        try:
            settings_file = self.farm_path / "insemination_settings.json"
            if not settings_file.exists():
                return False
            with open(settings_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            types_map = data.get('insemination_types', {}) or {}
            for code_key, name_value in types_map.items():
                if str(code_key).strip().upper() != s:
                    continue
                name_str = (name_value or '').strip()
                if name_str.upper() == 'R' or '連続授精' in name_str or '連続' in name_str:
                    return True
                return False
        except Exception:
            pass
        return False

    def _calculate_conception_rate(self, db, rate_type: str, start_date: Optional[str], 
                                   end_date: Optional[str], condition_text: str) -> Optional[Dict[str, Any]]:
        """
        受胎率を計算。
        その他 = ①授精種類コードR（連続授精）および7日以内間隔のAI/ET + ②妊娠鑑定日数経過にもかかわらずO/P/Rいずれも未記録（未鑑定）。
        再授精の扱い: 8日以上あいた後のAI/ETがあれば、その前の授精は空胎確定（O）とする。Rは授精間隔7日以内の場合のみ（1つの授精として扱う）。
        ＊は②が存在する場合のみ表示。①は＊対象外。
        受胎率 = 受胎数÷(受胎数＋不受胎数)。その他は母数に入れない。
        """
        try:
            from datetime import datetime, timedelta
            import json
            
            # デフォルト期間を設定（30日前からさかのぼって1年間）
            if not end_date:
                today = datetime.now()
                end_date_dt = today - timedelta(days=30)
                end_date = end_date_dt.strftime('%Y-%m-%d')
            if not start_date:
                end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')
                start_date_dt = end_date_dt - timedelta(days=365)
                start_date = start_date_dt.strftime('%Y-%m-%d')
            
            preg_check_days = self._get_pregnancy_check_days()
            end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')
            
            # 期間内の全AI/ETイベントを取得
            all_events = db.get_events_by_period(start_date, end_date, include_deleted=False)
            ai_et_events = [
                e for e in all_events 
                if e.get('event_number') in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET]
            ]
            
            # 日付順にソート
            ai_et_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
            
            # 各牛ごとにイベントをグループ化
            cow_events = {}  # {cow_auto_id: [events]}
            for event in ai_et_events:
                cow_auto_id = event.get('cow_auto_id')
                if cow_auto_id not in cow_events:
                    cow_events[cow_auto_id] = []
                cow_events[cow_auto_id].append(event)
            
            # 各AI/ETイベントについて受胎判定と分類値を計算
            insemination_data = []  # 各授精のデータ。other_reason = None | 'R' | 'undetermined_overdue'
            
            for cow_auto_id, events in cow_events.items():
                cow = db.get_cow_by_auto_id(cow_auto_id)
                if not cow:
                    continue
                
                # 経産牛のみ（LACT>=1）を対象とする
                lact = cow.get('lact')
                if lact is None or lact < 1:
                    continue
                
                # 条件チェック
                if condition_text:
                    if not self._evaluate_cow_condition_for_conception_rate(cow, condition_text, db):
                        continue
                
                # その牛の全イベントを取得（受胎判定のため）
                all_cow_events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
                all_cow_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
                
                # 授精回数は個体カードと同一定義（8日以内は同じ回数）で事前計算
                precomputed_insemination_counts = self.rule_engine.calculate_insemination_counts(cow_auto_id)
                
                # 各AI/ETイベントを処理
                for ai_event in events:
                    event_date = ai_event.get('event_date')
                    event_number = ai_event.get('event_number')
                    json_data_str = ai_event.get('json_data') or '{}'
                    try:
                        json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                    except Exception:
                        json_data = {}
                    
                    # 授精種類コード（R=連続授精）。農場設定のコード/表示名でR判定
                    ai_type_code = json_data.get('insemination_type_code') or json_data.get('type') or ''
                    ai_type_code = str(ai_type_code).strip() if ai_type_code is not None else ''
                    
                    # その他①: 連続授精（コードR）→ その他に計上（＊はつけない）
                    if self._is_insemination_type_r(ai_type_code):
                        classification_value = self._get_conception_rate_classification(
                            ai_event, cow, rate_type, all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        )
                        if classification_value is None:
                            continue
                        insemination_count_value = self._get_conception_rate_classification(
                            ai_event, cow, "授精回数", all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        ) or ""
                        insemination_data.append({
                            'classification': classification_value,
                            'conceived': False,
                            'other_reason': 'R',
                            'ai_event': ai_event,
                            'cow': cow,
                            'insemination_count': insemination_count_value,
                        })
                        continue
                    
                    # 受胎判定：このAI/ETイベント以降にPDP/PDP2/PAGPがあるか、または8日以上後の再授精でO確定
                    ai_event_date_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    conceived = False
                    undetermined = False
                    is_continuous_insemination = False  # 7日以内の連続授精（R扱い）
                    
                    for event in all_cow_events:
                        event_date_other = event.get('event_date')
                        if event_date_other <= event_date:
                            continue
                        event_number_other = event.get('event_number')
                        if event_number_other in [self.rule_engine.EVENT_PDP, self.rule_engine.EVENT_PDP2, self.rule_engine.EVENT_PAGP]:
                            conceived = True
                            break
                        elif event_number_other in [self.rule_engine.EVENT_PDN, self.rule_engine.EVENT_PAGN]:
                            break
                        elif event_number_other == self.rule_engine.EVENT_ABRT:
                            break
                        elif event_number_other in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET]:
                            try:
                                other_dt = datetime.strptime(event_date_other, '%Y-%m-%d')
                                interval_days = (other_dt - ai_event_date_dt).days
                            except (ValueError, TypeError):
                                interval_days = 0
                            # 8日以上後のAI/ET＝再授精→前の授精はO確定（undeterminedにしない）
                            if interval_days >= 8:
                                break
                            # 7日以内＝連続授精（R）。1つの授精として扱う
                            is_continuous_insemination = True
                            break
                        elif event_number_other == self.rule_engine.EVENT_CALV:
                            break
                    
                    # 7日以内の連続授精（R）→ その他に計上
                    if is_continuous_insemination:
                        classification_value = self._get_conception_rate_classification(
                            ai_event, cow, rate_type, all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        )
                        if classification_value is None:
                            continue
                        insemination_count_value = self._get_conception_rate_classification(
                            ai_event, cow, "授精回数", all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        ) or ""
                        insemination_data.append({
                            'classification': classification_value,
                            'conceived': False,
                            'other_reason': 'R',
                            'ai_event': ai_event,
                            'cow': cow,
                            'insemination_count': insemination_count_value,
                        })
                        continue
                    
                    # その他(1): 未確定のうち、妊娠鑑定日数が過ぎているのに妊娠鑑定未実施
                    if undetermined and not conceived:
                        days_since_ai = (end_date_dt - ai_event_date_dt).days
                        if days_since_ai < preg_check_days:
                            continue  # 鑑定日数未経過の未確定は表に含めない
                        classification_value = self._get_conception_rate_classification(
                            ai_event, cow, rate_type, all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        )
                        if classification_value is None:
                            continue
                        insemination_count_value = self._get_conception_rate_classification(
                            ai_event, cow, "授精回数", all_cow_events,
                            precomputed_insemination_counts=precomputed_insemination_counts
                        ) or ""
                        insemination_data.append({
                            'classification': classification_value,
                            'conceived': False,
                            'other_reason': 'undetermined_overdue',
                            'ai_event': ai_event,
                            'cow': cow,
                            'insemination_count': insemination_count_value,
                        })
                        continue
                    
                    # 分類値を計算（受胎/不受胎）
                    classification_value = self._get_conception_rate_classification(
                        ai_event, cow, rate_type, all_cow_events,
                        precomputed_insemination_counts=precomputed_insemination_counts
                    )
                    if classification_value is None:
                        continue
                    insemination_count_value = self._get_conception_rate_classification(
                        ai_event, cow, "授精回数", all_cow_events,
                        precomputed_insemination_counts=precomputed_insemination_counts
                    ) or ""
                    
                    # 受胎・不受胎がまだ確定していない場合（outcome無し／未鑑定）はその他に振り分け、母数に入れない
                    other_reason_val = None
                    if not conceived:
                        has_outcome = bool(json_data.get('outcome'))
                        dc305_no_result = bool(json_data.get('_dc305_no_result'))
                        if not has_outcome or dc305_no_result:
                            other_reason_val = 'undetermined_no_result'
                    
                    insemination_data.append({
                        'classification': classification_value,
                        'conceived': conceived,
                        'other_reason': other_reason_val,
                        'ai_event': ai_event,
                        'cow': cow,
                        'insemination_count': insemination_count_value,
                    })
            
            # 分類別に集計（受胎率＝受胎/(受胎＋不受胎)、その他は母数に入れない）
            # ＊は②の件数が1以上の場合のみ。①(R)のみの行には＊をつけない
            classification_stats = {}
            for data in insemination_data:
                classification = data['classification']
                if classification not in classification_stats:
                    classification_stats[classification] = {
                        'total': 0,
                        'conceived': 0,
                        'not_conceived': 0,
                        'other': 0,
                        'other_undetermined_overdue': 0,  # ②の件数のみ（＊の根拠）
                        'has_undetermined_overdue': False,
                        'events': [],
                    }
                
                classification_stats[classification]['total'] += 1
                other_reason = data.get('other_reason')  # None | 'R'(①) | 'undetermined_overdue'(②) | 'undetermined_no_result'(未鑑定)
                if other_reason:
                    classification_stats[classification]['other'] += 1
                    if other_reason == 'undetermined_overdue':
                        classification_stats[classification]['other_undetermined_overdue'] += 1
                        classification_stats[classification]['has_undetermined_overdue'] = True  # ②が1件でもあれば＊
                elif data['conceived']:
                    classification_stats[classification]['conceived'] += 1
                else:
                    classification_stats[classification]['not_conceived'] += 1
                
                classification_stats[classification]['events'].append({
                    'ai_event': data['ai_event'],
                    'cow': data['cow'],
                    'conceived': data['conceived'],
                    'undetermined': other_reason == 'undetermined_overdue',
                    'other_reason': other_reason,
                    'insemination_count': data.get('insemination_count', ''),
                })
            
            return {
                'rate_type': rate_type,
                'start_date': start_date,
                'end_date': end_date,
                'condition_text': condition_text,
                'stats': classification_stats
            }
            
        except Exception as e:
            logging.error(f"受胎率計算エラー: {e}", exc_info=True)
            return None
    
    def _get_conception_rate_classification(self, ai_event: Dict[str, Any], cow: Dict[str, Any], 
                                           rate_type: str, all_events: List[Dict[str, Any]],
                                           precomputed_insemination_counts: Optional[Dict[int, int]] = None) -> Optional[str]:
        """
        受胎率の分類値を取得
        
        Args:
            ai_event: AI/ETイベント
            cow: 牛データ
            rate_type: 受胎率の種類
            all_events: その牛の全イベント
            precomputed_insemination_counts: 授精回数別のとき RuleEngine 算出の {event_id: 回数}（8日ルール適用）
        
        Returns:
            分類値（文字列）、取得できない場合はNone
        """
        try:
            from datetime import datetime
            import json
            
            event_date = ai_event.get('event_date')
            json_data_str = ai_event.get('json_data') or '{}'
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
            except:
                json_data = {}
            
            if rate_type == "月":
                # 月：YYYY-MM形式
                event_date_dt = datetime.strptime(event_date, '%Y-%m-%d')
                return event_date_dt.strftime('%Y-%m')
            
            elif rate_type == "産次":
                # 産次：授精日時点での産次を計算
                event_date = ai_event.get('event_date')
                cow_auto_id = ai_event.get('cow_auto_id')
                if not event_date or not cow_auto_id:
                    return None
                
                try:
                    # 授精日時点での状態を計算
                    state = self.rule_engine.apply_events_until_date(cow_auto_id, event_date)
                    lact = state.get('lact')
                    if lact is not None and lact >= 1:
                        return str(lact)
                    return None
                except Exception as e:
                    logging.debug(f"産次計算エラー (cow_auto_id={cow_auto_id}, event_date={event_date}): {e}")
                    return None
            
            elif rate_type == "授精回数":
                # 授精回数：個体カードと同じ定義（RuleEngine: 8日以内のAI/ETは同じ回数）
                if precomputed_insemination_counts is not None:
                    event_id = ai_event.get('id')
                    count = precomputed_insemination_counts.get(event_id) if event_id is not None else None
                    return str(count) if count else None
                # 未渡しの場合は RuleEngine をその場で呼ぶ（後方互換）
                cow_auto_id = ai_event.get('cow_auto_id')
                if not cow_auto_id:
                    return None
                try:
                    counts = self.rule_engine.calculate_insemination_counts(cow_auto_id)
                    event_id = ai_event.get('id')
                    count = counts.get(event_id) if event_id is not None else None
                    return str(count) if count else None
                except Exception as e:
                    logging.debug(f"授精回数計算エラー (cow_auto_id={cow_auto_id}): {e}")
                    return None
            
            elif rate_type == "DIMサイクル":
                # DIMサイクル：授精日時点でのDIMを計算、50日スタートで20日サイクル
                event_date = ai_event.get('event_date')
                cow_auto_id = ai_event.get('cow_auto_id')
                
                if not event_date or not cow_auto_id:
                    return None
                
                # 授精日時点でのDIMを計算
                event_dim = None
                try:
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    
                    # 授精日以前の分娩イベント（event_number=202）を探す
                    calv_events_before = []
                    for ev in all_events:
                        ev_date_str = ev.get('event_date', '')
                        ev_number = ev.get('event_number')
                        if ev_number == 202 and ev_date_str:  # CALVイベント
                            try:
                                ev_dt = datetime.strptime(ev_date_str, '%Y-%m-%d')
                                if ev_dt <= event_dt:  # 授精日以前の分娩のみ
                                    calv_events_before.append((ev_dt, ev_date_str))
                            except ValueError:
                                continue
                    
                    # 授精日以前の分娩イベントを日付の降順でソート（最新が最初）
                    if calv_events_before:
                        calv_events_before.sort(reverse=True)  # 日付の降順
                        # 最新の分娩日を取得（授精日時点での「現在の産次」の分娩日）
                        calv_date = calv_events_before[0][1]
                        
                        # DIMを計算
                        calv_dt = datetime.strptime(calv_date, '%Y-%m-%d')
                        dim = (event_dt - calv_dt).days
                        if dim >= 0:
                            event_dim = dim
                except Exception as e:
                    logging.debug(f"DIMサイクル計算エラー (cow_auto_id={cow_auto_id}, event_date={event_date}): {e}")
                    return None
                
                if event_dim is None:
                    return None
                
                # 50日未満は除外
                if event_dim < 50:
                    return None
                
                # 50-69, 70-89, 90-109, ... というように20日サイクル
                cycle_start = ((event_dim - 50) // 20) * 20 + 50
                cycle_end = cycle_start + 19
                return f"{cycle_start}-{cycle_end}"
            
            elif rate_type == "授精間隔":
                # 授精間隔：前回の授精からの日数を範囲に分類
                event_date_dt = datetime.strptime(event_date, '%Y-%m-%d')
                prev_ai_date = None
                
                # 前回のAI/ETイベントを探す
                for event in all_events:
                    if event.get('id') == ai_event.get('id'):
                        continue
                    if event.get('event_number') in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET]:
                        prev_date = event.get('event_date')
                        if prev_date and prev_date < event_date:
                            if prev_ai_date is None or prev_date > prev_ai_date:
                                prev_ai_date = prev_date
                
                if prev_ai_date:
                    prev_date_dt = datetime.strptime(prev_ai_date, '%Y-%m-%d')
                    interval = (event_date_dt - prev_date_dt).days
                    
                    # 範囲に分類
                    if interval <= 3:
                        return "1-3"
                    elif interval <= 17:
                        return "4-17"
                    elif interval <= 24:
                        return "18-24"
                    elif interval <= 35:
                        return "25-35"
                    elif interval <= 48:
                        return "36-48"
                    else:
                        return "49以上"
                else:
                    # 前回の授精がない場合（初回授精）は計算対象外
                    return None
            
            elif rate_type == "授精種類":
                # 授精種類：json_dataからinsemination_type_codeを取得して農場設定の表示名に変換
                if not isinstance(json_data, dict):
                    logging.warning(f"授精種類分類取得: json_dataが辞書ではありません: {json_data}")
                    return None
                
                # insemination_type_codeを取得（様々なキー名を試す）
                ai_type_code = (json_data.get('insemination_type_code') or 
                               json_data.get('inseminationTypeCode') or
                               json_data.get('type') or 
                               json_data.get('ai_type') or 
                               json_data.get('insemination_type') or 
                               json_data.get('inseminationType', ''))
                
                if not ai_type_code or ai_type_code == 'R':
                    # 連続授精（R）の場合は計算対象外
                    return None
                
                # insemination_settings.jsonから授精種類コード辞書を取得
                insemination_settings_file = self.farm_path / "insemination_settings.json"
                insemination_type_codes = {}
                
                if insemination_settings_file.exists():
                    try:
                        with open(insemination_settings_file, 'r', encoding='utf-8') as f:
                            insemination_settings = json.load(f)
                            insemination_type_codes = insemination_settings.get('insemination_types', {})
                    except Exception as e:
                        logging.warning(f"insemination_settings.json読み込みエラー: {e}")
                
                # もしinsemination_settings.jsonにデータがない場合は、farm_settings.jsonから取得（後方互換性のため）
                if not insemination_type_codes:
                    settings = SettingsManager(self.farm_path)
                    insemination_type_codes = settings.get('insemination_type_codes', {})
                
                logging.info(f"授精種類分類取得: ai_type_code={ai_type_code}, insemination_type_codes={insemination_type_codes}")
                
                # コードから表示名を取得
                ai_type_code_str = str(ai_type_code).strip().upper()  # 大文字に統一
                
                logging.info(f"授精種類分類取得: ai_type_code_str={ai_type_code_str}, insemination_type_codes={insemination_type_codes}, type={type(insemination_type_codes)}")
                
                # 文字列キーで検索（大文字小文字を区別しない）
                ai_type_name = None
                for code_key, name_value in insemination_type_codes.items():
                    code_key_str = str(code_key).strip().upper()
                    if code_key_str == ai_type_code_str:
                        ai_type_name = name_value
                        logging.info(f"授精種類名を取得: code={ai_type_code_str} -> name={ai_type_name}")
                        break
                
                if ai_type_name:
                    return ai_type_name
                else:
                    # コードが見つからない場合は、コードそのものを返す（デバッグ用）
                    logging.warning(f"授精種類コード '{ai_type_code_str}' が農場設定に見つかりません。insemination_type_codes={insemination_type_codes}, keys={list(insemination_type_codes.keys()) if isinstance(insemination_type_codes, dict) else 'not dict'}")
                    return ai_type_code_str
            
            elif rate_type == "授精師":
                # 授精師：json_dataからtechnician_codeを取得して農場設定の表示名に変換
                if not isinstance(json_data, dict):
                    logging.warning(f"授精師分類取得: json_dataが辞書ではありません: {json_data}")
                    return None
                
                # technician_codeを取得（様々なキー名を試す）
                technician_code = (json_data.get('technician_code') or 
                                  json_data.get('technicianCode') or
                                  json_data.get('technician') or 
                                  json_data.get('tech_code') or 
                                  json_data.get('techCode', ''))
                
                if not technician_code:
                    return None
                
                # insemination_settings.jsonから授精師コード辞書を取得
                insemination_settings_file = self.farm_path / "insemination_settings.json"
                inseminator_codes = {}
                
                if insemination_settings_file.exists():
                    try:
                        with open(insemination_settings_file, 'r', encoding='utf-8') as f:
                            insemination_settings = json.load(f)
                            inseminator_codes = insemination_settings.get('technicians', {})
                    except Exception as e:
                        logging.warning(f"insemination_settings.json読み込みエラー: {e}")
                
                # もしinsemination_settings.jsonにデータがない場合は、farm_settings.jsonから取得（後方互換性のため）
                if not inseminator_codes:
                    settings = SettingsManager(self.farm_path)
                    inseminator_codes = settings.get('inseminator_codes', {})
                
                logging.info(f"授精師分類取得: technician_code={technician_code}, inseminator_codes={inseminator_codes}")
                
                # コードから表示名を取得
                technician_code_str = str(technician_code).strip()
                
                logging.info(f"授精師分類取得: technician_code_str={technician_code_str}, inseminator_codes={inseminator_codes}, type={type(inseminator_codes)}")
                
                # 文字列キーで検索（大文字小文字を区別しない）
                technician_name = None
                for code_key, name_value in inseminator_codes.items():
                    code_key_str = str(code_key).strip()
                    if code_key_str == technician_code_str:
                        technician_name = name_value
                        logging.info(f"授精師名を取得: code={technician_code_str} -> name={technician_name}")
                        break
                
                # 見つからない場合は整数キーも試す
                if technician_name is None and technician_code_str.isdigit():
                    technician_name = inseminator_codes.get(int(technician_code_str), None)
                    if technician_name:
                        logging.info(f"授精師名を取得（整数キー）: code={technician_code_str} -> name={technician_name}")
                
                if technician_name:
                    return technician_name
                else:
                    # コードが見つからない場合は、コードそのものを返す（デバッグ用）
                    logging.warning(f"授精師コード '{technician_code_str}' が農場設定に見つかりません。inseminator_codes={inseminator_codes}, keys={list(inseminator_codes.keys()) if isinstance(inseminator_codes, dict) else 'not dict'}")
                    return technician_code_str
            
            elif rate_type == "SIRE":
                # SIRE：json_dataからsireを取得してそのまま返す
                if not isinstance(json_data, dict):
                    logging.warning(f"SIRE分類取得: json_dataが辞書ではありません: {json_data}")
                    return None
                
                sire = json_data.get('sire') or json_data.get('SIRE') or json_data.get('sire_code') or ''
                sire_str = str(sire).strip() if sire is not None else ''
                return sire_str if sire_str else None
            
            return None
            
        except Exception as e:
            logging.error(f"分類値取得エラー: {e}", exc_info=True)
            return None
    
    def _display_conception_rate_result(self, result: Dict[str, Any], rate_type: str, period: Optional[str] = None):
        """受胎率結果を表示"""
        try:
            stats = result.get('stats', {})
            if not stats:
                self.add_message(role="system", text="該当するデータがありません")
                return
            
            # 全授精の総数を先に計算（％全授精の計算に使用）
            total_all = sum(s['total'] for s in stats.values())
            
            # 表のヘッダーを設定（分類タイプに応じて動的に変更）
            classification_header = rate_type  # 分類タイプをそのまま使用
            headers = [classification_header, "受胎率", "受胎", "不受胎", "その他", "総数", "％全授精"]
            
            # データ行を作成
            rows = []
            
            # 分類をソート（DIMサイクルと授精間隔の場合は数値順にソート）
            def sort_key(classification):
                # DIMサイクルの場合（例："50-69"）は、開始値を数値として抽出してソート
                if rate_type == "DIMサイクル" and '-' in classification:
                    try:
                        # "50-69" から "50" を抽出
                        start_value = int(classification.split('-')[0])
                        return (0, start_value)  # 0は通常の分類を示す
                    except (ValueError, IndexError):
                        return (1, classification)  # 数値変換に失敗した場合は文字列としてソート
                # 授精間隔の場合（例："1-3", "4-17", "18-24", "25-35", "36-48", "49以上"）
                elif rate_type == "授精間隔":
                    if classification == "49以上":
                        return (0, 49)  # 49以上は最後に来る
                    elif '-' in classification:
                        try:
                            # "1-3" から "1" を抽出
                            start_value = int(classification.split('-')[0])
                            return (0, start_value)
                        except (ValueError, IndexError):
                            return (1, classification)
                    else:
                        return (1, classification)
                # 産次・産次別：数値順（1, 2, 3, … 9, 10）、合計は最後
                elif rate_type in ("産次", "産次別"):
                    if classification == "合計":
                        return (2, -1)  # 合計は最後
                    try:
                        return (0, int(classification))
                    except (ValueError, TypeError):
                        return (1, classification)
                else:
                    # その他の分類（月など）は文字列としてソート
                    # ただし「合計」は最後に来るようにする
                    if classification == "合計":
                        return (2, "")  # 2は合計を示す
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            for classification in sorted_classifications:
                stat = stats[classification]
                total = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)。その他は母数に入れない。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                
                # ％全授精を計算
                if total_all > 0:
                    percent_of_all = (total / total_all) * 100
                    percent_of_all_str = f"{percent_of_all:.1f}%"
                else:
                    percent_of_all_str = "0.0%"
                
                other_str = str(other)
                
                rows.append([
                    classification,
                    rate_str,
                    str(conceived),
                    str(not_conceived),
                    other_str,
                    str(total),
                    percent_of_all_str
                ])
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            
            total_other_str = str(total_other)
            
            rows.append([
                "合計",
                total_rate_str,
                str(total_conceived),
                str(total_not_conceived),
                total_other_str,
                str(total_all),
                "100.0%"  # 合計行の％全授精は100.0%
            ])
            
            # 分類ごとのイベントデータを保持（ダブルクリック時に使用）
            self.conception_rate_events_by_classification = {}
            for classification in stats.keys():
                if 'events' in stats[classification]:
                    self.conception_rate_events_by_classification[classification] = stats[classification]['events']
            
            # 受胎率結果のrate_typeを保持（ダブルクリック時に使用）
            self.current_conception_rate_type = rate_type
            
            # 受胎率結果データを保存（グラフ表示用）
            self.current_conception_rate_result = result
            self.current_conception_rate_period = period
            
            # 期間情報を取得（引数で渡されていない場合はresultから取得）
            if not period:
                start_date = result.get('start_date', '')
                end_date = result.get('end_date', '')
                if start_date and end_date:
                    period = f"{start_date} ～ {end_date}"
            
            # 表を表示（条件をタイトルに追加）
            condition_text = result.get('condition_text', '')
            title = f"受胎率（経産）（{rate_type}別）"
            if condition_text:
                title = f"{title}：{condition_text}"
            self._display_list_result_in_table(headers, rows, title, period=period)
            
            # 受胎率結果テーブルにダブルクリックイベントをバインド
            self.result_treeview.unbind("<Double-Button-1>")
            self.result_treeview.bind("<Double-Button-1>", self._on_conception_rate_cell_double_click)
            
            # グラフタブでも受胎率の結果を表示できるようにする
            self._display_conception_rate_graph(result, rate_type, period)
            
        except Exception as e:
            logging.error(f"受胎率結果表示エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"結果の表示中にエラーが発生しました: {e}")
    
    def _display_conception_rate_graph(self, result: Dict[str, Any], rate_type: str, period: Optional[str] = None):
        """
        受胎率の結果をグラフ表示
        
        Args:
            result: 受胎率計算結果
            rate_type: 受胎率の種類
            period: 対象期間
        """
        try:
            stats = result.get('stats', {})
            if not stats:
                return
            
            # グラフ表示用のデータを準備
            classifications = []
            rates = []
            conceived_counts = []
            not_conceived_counts = []
            other_counts = []
            total_counts = []
            
            # 分類をソート（DIMサイクル・授精間隔・産次は数値順）
            def sort_key(classification):
                if rate_type == "DIMサイクル" and '-' in classification:
                    try:
                        start_value = int(classification.split('-')[0])
                        return (0, start_value)
                    except (ValueError, IndexError):
                        return (1, classification)
                elif rate_type == "授精間隔":
                    if classification == "49以上":
                        return (0, 49)
                    elif '-' in classification:
                        try:
                            start_value = int(classification.split('-')[0])
                            return (0, start_value)
                        except (ValueError, IndexError):
                            return (1, classification)
                    else:
                        return (1, classification)
                elif rate_type in ("産次", "産次別"):
                    if classification == "合計":
                        return (2, -1)
                    try:
                        return (0, int(classification))
                    except (ValueError, TypeError):
                        return (1, classification)
                else:
                    if classification == "合計":
                        return (2, "")
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            for classification in sorted_classifications:
                if classification == "合計":
                    continue  # 合計はグラフに含めない
                
                stat = stats[classification]
                total = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                else:
                    rate = 0.0
                
                classifications.append(classification)
                rates.append(rate)
                conceived_counts.append(conceived)
                not_conceived_counts.append(not_conceived)
                other_counts.append(other)
                total_counts.append(total)
            
            if not classifications:
                return
            
            # グラフ表示用のデータ形式を作成（集計と同じ形式）
            graph_data = {
                'type': 'conception_rate',
                'x_axis': classifications,
                'y_axis': rates,
                'title': f"受胎率（経産）（{rate_type}別）",
                'x_label': rate_type,
                'y_label': '受胎率（%）',
                'additional_data': {
                    'conceived': conceived_counts,
                    'not_conceived': not_conceived_counts,
                    'other': other_counts,
                    'total': total_counts
                },
                'period': period
            }
            
            # グラフ表示用のデータを保存
            self.current_graph_data = graph_data
            self.current_graph_command = f"受胎率（経産）（{rate_type}別）"
            
            # グラフタブが存在する場合はグラフを表示
            if hasattr(self, 'graph_canvas') and self.graph_canvas and MATPLOTLIB_AVAILABLE:
                self._draw_conception_rate_graph(graph_data)
                
        except Exception as e:
            logging.error(f"受胎率グラフ表示エラー: {e}", exc_info=True)
    
    def _draw_conception_rate_graph(self, graph_data: Dict[str, Any]):
        """
        受胎率の結果をグラフ表示（新規ウィンドウ）
        
        Args:
            graph_data: グラフ表示用のデータ
        """
        try:
            if not MATPLOTLIB_AVAILABLE:
                logging.warning("[グラフ] matplotlibが利用できません")
                return
            
            x_labels = graph_data.get('x_axis', [])
            y_values = graph_data.get('y_axis', [])
            title = graph_data.get('title', '受胎率')
            x_label = graph_data.get('x_label', '分類')
            y_label = graph_data.get('y_label', '受胎率（%）')
            period = graph_data.get('period', '')
            command = self.current_graph_command or title
            
            if not x_labels or not y_values:
                logging.warning("[グラフ] グラフ表示用のデータがありません")
                return
            
            # 新しいウィンドウを作成
            graph_window = tk.Toplevel(self.root)
            graph_window.title(f"グラフ: {title}")
            graph_window.geometry("800x600")
            
            # コマンド表示
            command_text = f"コマンド: {command}"
            if period:
                command_text += f" | 対象期間：{period}"
            
            command_label = tk.Label(
                graph_window,
                text=command_text,
                font=(self._main_font_family, 9),
                anchor=tk.W
            )
            command_label.pack(fill=tk.X, padx=10, pady=5)
            
            # matplotlib FigureとCanvasを作成
            figure = Figure(figsize=(10, 6), dpi=100)
            canvas = FigureCanvasTkAgg(figure, graph_window)
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 日本語フォントの設定
            try:
                import matplotlib.pyplot as plt
                import matplotlib.font_manager as fm
                
                # Windowsで利用可能な日本語フォントを探す
                japanese_fonts = ['MS Gothic', 'MS PGothic', 'Yu Gothic', 'Meiryo', 'Takao']
                font_found = None
                for font_name in japanese_fonts:
                    try:
                        font_path = fm.findfont(fm.FontProperties(family=font_name))
                        if font_path:
                            font_found = font_name
                            break
                    except:
                        continue
                
                if font_found:
                    plt.rcParams['font.family'] = font_found
                    logging.debug(f"[グラフ] 日本語フォントを設定: {font_found}")
                else:
                    logging.warning("[グラフ] 日本語フォントが見つかりません")
            except Exception as e:
                logging.warning(f"[グラフ] フォント設定エラー: {e}")
            
            # グラフを作成
            ax = figure.add_subplot(111)
            
            # 棒グラフを作成
            bars = ax.bar(range(len(x_labels)), y_values, color='steelblue', alpha=0.7, edgecolor='black', linewidth=0.5)
            
            # X軸のティック位置とラベルを設定
            ax.set_xticks(range(len(x_labels)))
            ax.set_xticklabels(x_labels, rotation=45, ha='right')
            
            # 各棒の上に値を表示
            for i, (bar, value) in enumerate(zip(bars, y_values)):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2., height,
                       f'{value:.1f}%',
                       ha='center', va='bottom', fontsize=9)
            
            # タイトルとラベルを設定
            full_title = title
            if period:
                full_title = f"{title} | 対象期間：{period}"
            ax.set_title(full_title, fontsize=12, fontweight='bold', pad=15)
            ax.set_xlabel(x_label, fontsize=11, fontweight='bold')
            ax.set_ylabel(y_label, fontsize=11, fontweight='bold')
            
            # グリッドを追加
            ax.grid(True, linestyle='--', alpha=0.3, axis='y')
            ax.set_axisbelow(True)
            
            # Y軸の範囲を設定（0-100%）
            y_max = max(y_values) if y_values else 100
            ax.set_ylim(0, max(100, y_max * 1.2))
            
            # レイアウトを調整
            figure.tight_layout()
            canvas.draw()
            
            logging.info(f"[グラフ] 受胎率グラフ描画完了: {len(x_labels)}件のデータを表示")
            
        except Exception as e:
            logging.error(f"[グラフ] 受胎率グラフ表示エラー: {e}", exc_info=True)
            import traceback
            traceback.print_exc()
            # エラー時も何か表示する
            try:
                if 'graph_window' in locals():
                    error_label = tk.Label(
                        graph_window,
                        text=f"グラフ表示エラー: {str(e)[:100]}",
                        font=(self._main_font_family, 10),
                        fg='red',
                        wraplength=700
                    )
                    error_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            except Exception as e2:
                logging.error(f"[グラフ] エラー表示も失敗: {e2}")
    
    def _on_conception_rate_cell_double_click(self, event):
        """受胎率結果テーブルのセルをダブルクリックした時の処理"""
        try:
            # クリックされたセルを取得
            item = self.result_treeview.selection()[0] if self.result_treeview.selection() else None
            if not item:
                return
            
            # クリックされた列を取得（#1, #2, ... 形式。空や不正な場合はスキップ）
            column = self.result_treeview.identify_column(event.x) or ""
            col_str = (column.replace("#", "") or "").strip()
            try:
                column_index = int(col_str) - 1 if col_str else -1
            except (ValueError, TypeError):
                column_index = -1
            if column_index < 0:
                return

            # 行の値を取得
            values = self.result_treeview.item(item, 'values')
            if not values:
                return
            
            # 分類値を取得（最初の列）
            classification = values[0] if len(values) > 0 else None
            if not classification or classification == "合計":
                return
            
            # 該当する分類のイベントリストを取得
            if not hasattr(self, 'conception_rate_events_by_classification'):
                return
            
            events_data = self.conception_rate_events_by_classification.get(classification, [])
            if not events_data:
                return
            
            # イベントリストを表示するウィンドウを開く
            self._show_conception_rate_events_window(classification, events_data)
            
        except Exception as e:
            logging.error(f"受胎率セルダブルクリックエラー: {e}", exc_info=True)
    
    def _show_conception_rate_events_window(self, classification: str, events_data: List[Dict[str, Any]]):
        """
        受胎率の分類に該当するイベントリストを表示するウィンドウを開く
        
        Args:
            classification: 分類値
            events_data: イベントデータのリスト
        """
        try:
            import tkinter as tk
            from tkinter import ttk
            from datetime import datetime
            import json
            from settings_manager import SettingsManager
            
            # 新しいウィンドウを作成
            window = tk.Toplevel(self.root)
            window.title(f"受胎率イベントリスト - {classification}")
            window.geometry("1000x600")
            
            # フレームを作成
            main_frame = ttk.Frame(window, padding="10")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # タイトル
            title_label = ttk.Label(main_frame, text=f"分類: {classification}", font=(self._main_font_family, self.default_font_size + 2, "bold"))
            title_label.pack(pady=(0, 10))
            
            # テーブルフレーム
            table_frame = ttk.Frame(main_frame)
            table_frame.pack(fill=tk.BOTH, expand=True)
            
            # スクロールバー
            scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            
            scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            
            # テーブル
            columns = ['ID', '日付', 'DIM', 'SIRE', '授精師', '授精回数', '授精種類', '結果']
            treeview = ttk.Treeview(table_frame, columns=columns, show='headings', 
                                   yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
            treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar_y.config(command=treeview.yview)
            scrollbar_x.config(command=treeview.xview)
            
            # 列ヘッダーを設定
            for col in columns:
                treeview.heading(col, text=col)
                treeview.column(col, width=100, anchor=tk.W)
            
            # 農場設定を読み込む
            settings = SettingsManager(self.farm_path)
            inseminator_codes = settings.get('inseminator_codes', {})
            
            # insemination_settings.jsonから授精種類コード辞書を取得
            insemination_settings_file = self.farm_path / "insemination_settings.json"
            insemination_type_codes = {}
            
            if insemination_settings_file.exists():
                try:
                    with open(insemination_settings_file, 'r', encoding='utf-8') as f:
                        insemination_settings = json.load(f)
                        insemination_type_codes = insemination_settings.get('insemination_types', {})
                except Exception as e:
                    logging.warning(f"insemination_settings.json読み込みエラー: {e}")
            
            # もしinsemination_settings.jsonにデータがない場合は、farm_settings.jsonから取得（後方互換性のため）
            if not insemination_type_codes:
                insemination_type_codes = settings.get('insemination_type_codes', {})
            
            # イベントデータを表示
            for event_data in events_data:
                ai_event = event_data['ai_event']
                cow = event_data['cow']
                
                # ID
                cow_id = cow.get('cow_id', '')
                
                # 日付
                event_date = ai_event.get('event_date', '')
                
                # DIMを計算
                dim_display = ""
                if event_date:
                    try:
                        event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        cow_auto_id = cow.get('auto_id')
                        if cow_auto_id:
                            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                            calv_events_before = []
                            for ev in all_events:
                                ev_date_str = ev.get('event_date', '')
                                ev_number = ev.get('event_number')
                                if ev_number == 202 and ev_date_str:  # CALVイベント
                                    try:
                                        ev_dt = datetime.strptime(ev_date_str, '%Y-%m-%d')
                                        if ev_dt <= event_dt:
                                            calv_events_before.append((ev_dt, ev_date_str))
                                    except ValueError:
                                        continue
                            
                            if calv_events_before:
                                calv_events_before.sort(reverse=True)
                                calv_date = calv_events_before[0][1]
                                calv_dt = datetime.strptime(calv_date, '%Y-%m-%d')
                                dim = (event_dt - calv_dt).days
                                if dim >= 0:
                                    dim_display = str(dim)
                    except Exception as e:
                        logging.debug(f"DIM計算エラー: {e}")
                
                # SIRE
                json_data_str = ai_event.get('json_data') or '{}'
                try:
                    json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                except:
                    json_data = {}
                sire = json_data.get('sire', '') or ''
                
                # 授精師
                technician_code = json_data.get('technician') or json_data.get('technician_code', '')
                technician_name = "不明"
                if technician_code:
                    technician_code_str = str(technician_code).strip()
                    technician_name = inseminator_codes.get(technician_code_str, technician_code_str)
                    if technician_name == technician_code_str and technician_code_str.isdigit():
                        technician_name = inseminator_codes.get(int(technician_code_str), technician_code_str)
                
                # 授精回数（保持されている値を使用）
                insemination_count = event_data.get('insemination_count', '') or ''
                
                # 授精種類
                # 様々なキー名を試す（insemination_type_codeも含む）
                ai_type_code = (json_data.get('type') or 
                               json_data.get('ai_type') or 
                               json_data.get('insemination_type') or 
                               json_data.get('inseminationType') or
                               json_data.get('insemination_type_code') or
                               json_data.get('inseminationTypeCode', ''))
                
                logging.info(f"授精種類取得: json_data.keys()={list(json_data.keys()) if isinstance(json_data, dict) else 'not dict'}, ai_type_code={ai_type_code}, insemination_type_codes={insemination_type_codes}")
                
                ai_type_name = "不明"
                if ai_type_code and ai_type_code != 'R':
                    ai_type_code_str = str(ai_type_code).strip()
                    # 文字列キーで検索
                    ai_type_name = insemination_type_codes.get(ai_type_code_str, None)
                    
                    # 見つからない場合は、大文字小文字を区別せずに検索
                    if ai_type_name is None:
                        for code_key, name_value in insemination_type_codes.items():
                            if str(code_key).upper() == ai_type_code_str.upper():
                                ai_type_name = name_value
                                logging.info(f"授精種類名を取得（大文字小文字無視）: code={ai_type_code_str} -> name={ai_type_name}")
                                break
                    
                    # それでも見つからない場合は、コードそのものを表示
                    if ai_type_name is None or ai_type_name == ai_type_code_str:
                        logging.warning(f"授精種類コード '{ai_type_code_str}' が農場設定に見つかりません。insemination_type_codes={insemination_type_codes}")
                        if ai_type_name is None:
                            ai_type_name = ai_type_code_str
                else:
                    if not ai_type_code:
                        logging.warning(f"授精種類コードが取得できません: json_data={json_data}")
                    elif ai_type_code == 'R':
                        logging.debug(f"授精種類が'R'（連続授精）のためスキップ")
                
                # 結果（P, O, R, 未鑑定）
                result = ""
                if event_data['conceived']:
                    result = "P"
                elif event_data.get('other_reason') == 'R' or event_data.get('undetermined'):
                    # 7日以内の連続授精（other_reason='R'）および旧ロジック互換の未確定フラグはR表示
                    result = "R"
                elif event_data.get('other_reason') == 'undetermined_no_result':
                    result = "未鑑定"
                else:
                    # DBのoutcomeを参照し、無い場合は未鑑定表示（既存データの整合性のため）
                    ai_json = event_data.get('ai_event', {}).get('json_data') or {}
                    if isinstance(ai_json, str):
                        try:
                            ai_json = json.loads(ai_json)
                        except Exception:
                            ai_json = {}
                    if not ai_json.get('outcome') or ai_json.get('_dc305_no_result'):
                        result = "未鑑定"
                    else:
                        result = "O"
                
                # 行を追加
                item_id = treeview.insert('', tk.END, values=[
                    cow_id,
                    event_date,
                    dim_display,
                    sire,
                    technician_name,
                    insemination_count,
                    ai_type_name,
                    result
                ], tags=(cow_id,))  # タグにcow_idを保存
            
            # ダブルクリックイベントをバインド（個体カードを開く）
            def on_row_double_click(event):
                selection = treeview.selection()
                if selection:
                    item = selection[0]
                    tags = treeview.item(item, 'tags')
                    if tags:
                        cow_id = tags[0]
                        self._jump_to_cow_card(cow_id)
            
            treeview.bind("<Double-Button-1>", on_row_double_click)
            
        except Exception as e:
            logging.error(f"受胎率イベントリスト表示エラー: {e}", exc_info=True)
            from tkinter import messagebox
            messagebox.showerror("エラー", f"イベントリストの表示中にエラーが発生しました: {e}")
    
