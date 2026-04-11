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

# 集計グラフ（MainWindow._display_aggregate_graph）と受胎率棒グラフで統一する Matplotlib スタイル
FALCON_BAR_CHART_KW = dict(color="steelblue", alpha=0.7, edgecolor="black", linewidth=0.5)
FALCON_BAR_CHART_GRID_KW = dict(linestyle="--", alpha=0.3, axis="y")
FALCON_BAR_CHART_TITLE_KW = dict(fontsize=12, fontweight="bold", pad=15)
FALCON_BAR_CHART_AXIS_LABEL_FS = 11
FALCON_BAR_CHART_VALUE_FS = 9


class ConceptionRateMixin:
    """Mixin: FALCON2 - 受胎率計算・表示 Mixin"""

    def _execute_pregnancy_rate_command(self):
        """妊娠率コマンドを実行（21日妊娠率: DairyComp305互換）"""
        try:
            # VWP取得
            vwp_str = self.pregnancy_rate_vwp_var.get().strip() if hasattr(self, 'pregnancy_rate_vwp_var') else "50"
            try:
                vwp = int(vwp_str)
                if vwp < 1:
                    vwp = 50
            except ValueError:
                vwp = 50

            # 期間取得
            start_date = None
            end_date = None
            if hasattr(self, 'start_date_entry') and self.start_date_entry.winfo_exists():
                s = self.start_date_entry.get().strip()
                if s and not self.start_date_placeholder_shown:
                    start_date = s
            if hasattr(self, 'end_date_entry') and self.end_date_entry.winfo_exists():
                e = self.end_date_entry.get().strip()
                if e and not self.end_date_placeholder_shown:
                    end_date = e

            # 条件取得（受胎率と同じ評価で経産牛リストを絞り込み）
            condition_text = ""
            if hasattr(self, 'pregnancy_rate_condition_entry') and self.pregnancy_rate_condition_entry.winfo_exists():
                if not self.pregnancy_rate_condition_placeholder_shown:
                    condition_text = self.pregnancy_rate_condition_entry.get().strip()

            period_text = f"{start_date} ～ {end_date}" if start_date and end_date else "全期間"
            self.add_message(role="system", text=f"妊娠率を計算しています… VWP={vwp}日 / 期間: {period_text}")
            self.result_notebook.select(0)  # 表タブを選択

            def calculate_in_thread():
                try:
                    from modules.reproduction_analysis import ReproductionAnalysis
                    analyzer = ReproductionAnalysis(
                        db=self.db,
                        rule_engine=self.rule_engine,
                        formula_engine=self.formula_engine,
                        vwp=vwp
                    )
                    cow_filter = None
                    if condition_text:
                        def _cf(cow):
                            return self._evaluate_cow_condition_for_conception_rate(
                                cow, condition_text, self.db)
                        cow_filter = _cf
                    results = analyzer.analyze(start_date, end_date, cow_filter=cow_filter)
                    self.root.after(0, lambda res=results, cf=cow_filter: self._display_pregnancy_rate_result(
                        res, vwp, period_text, condition_text, cf
                    ))
                except Exception as ex:
                    err = str(ex)
                    logging.error(f"妊娠率計算エラー: {err}", exc_info=True)
                    self.root.after(0, lambda msg=err: self.add_message(
                        role="system", text=f"妊娠率の計算中にエラーが発生しました: {msg}"
                    ))

            thread = threading.Thread(target=calculate_in_thread, daemon=True)
            thread.start()

        except Exception as e:
            logging.error(f"妊娠率コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"妊娠率の計算中にエラーが発生しました: {e}")

    def _display_pregnancy_rate_result(self, results, vwp: int, period_text: str, condition_text: str,
                                       cow_filter=None):
        """妊娠率結果を表とグラフで表示"""
        try:
            self._pr_cow_filter = cow_filter
            if not results:
                self.add_message(role="system", text="該当するデータがありません")
                return

            headers = ["開始日", "終了日", "繁殖対象", "授精", "授精率%", "妊娠対象", "妊娠", "受胎率%", "妊娠率%", "損耗"]
            rows = []

            total_br_el = 0
            total_bred = 0
            total_preg_eligible = 0
            total_preg = 0
            total_loss = 0

            def fmt_date(d: str) -> str:
                try:
                    from datetime import datetime as _dt
                    dt = _dt.strptime(d, '%Y-%m-%d')
                    return f"{dt.month}/{dt.day}"
                except Exception:
                    return d

            for r in results:
                hdr_str = f"{int(r.hdr)}%" if r.br_el > 0 else "-"
                pr_str = f"{int(r.pr)}%" if r.preg_eligible > 0 else "-"
                cr_str = f"{round(r.preg / r.bred * 100)}%" if r.bred > 0 and r.preg_eligible > 0 else "-"
                rows.append([
                    fmt_date(r.start_date),
                    fmt_date(r.end_date),
                    str(r.br_el),
                    str(r.bred),
                    hdr_str,
                    str(r.preg_eligible) if r.preg_eligible > 0 else "-",
                    str(r.preg) if r.preg_eligible > 0 else "-",
                    cr_str,
                    pr_str,
                    str(r.loss) if r.preg_eligible > 0 else "-",
                ])
                total_br_el += r.br_el
                total_bred += r.bred
                total_preg_eligible += r.preg_eligible
                total_preg += r.preg
                total_loss += r.loss

            # 合計行
            total_hdr_str = f"{round(total_bred / total_br_el * 100)}%" if total_br_el > 0 else "-"
            total_pr_str = f"{round(total_preg / total_preg_eligible * 100)}%" if total_preg_eligible > 0 else "-"
            total_cr_str = f"{round(total_preg / total_bred * 100)}%" if total_bred > 0 and total_preg_eligible > 0 else "-"
            rows.append([
                "【合計】",
                "",
                str(total_br_el),
                str(total_bred),
                total_hdr_str,
                str(total_preg_eligible) if total_preg_eligible > 0 else "-",
                str(total_preg) if total_preg_eligible > 0 else "-",
                total_cr_str,
                total_pr_str,
                str(total_loss) if total_preg_eligible > 0 else "-",
            ])

            title = f"妊娠率（21日サイクル）VWP={vwp}日"
            if condition_text:
                title = f"{title}：{condition_text}"
            self._display_list_result_in_table(headers, rows, title, period=period_text)

            # ヘッダー着色・合計行太字
            self._apply_pr_table_styling(headers)

            # ヒートマップ着色（妊娠率%列: ≥20%→薄緑、未満→無色）
            self._apply_pregnancy_rate_heatmap()

            # ドリルダウン用: 元データとVWPを保存し、シングルクリックをバインド
            self._pr_cycle_results = list(results)
            self._pr_vwp = vwp
            if hasattr(self, 'result_treeview') and self.result_treeview.winfo_exists():
                self.result_treeview.unbind("<Double-Button-1>")
                self.result_treeview.unbind("<Button-1>")
                self.result_treeview.unbind("<ButtonRelease-1>")
                self.result_treeview.bind("<ButtonRelease-1>", self._on_pregnancy_rate_row_click)

            # グラフ表示
            if MATPLOTLIB_AVAILABLE:
                self._display_pregnancy_rate_graph(results, title=title, period_text=period_text)

        except Exception as e:
            logging.error(f"妊娠率結果表示エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"結果の表示中にエラーが発生しました: {e}")

    def _display_pregnancy_rate_graph(self, results, title: str, period_text: str):
        """
        妊娠率・授精率の積み上げ棒グラフを別ウィンドウで表示。
        ウィンドウ内に目標ライン入力欄を設置し、入力後「補助線を更新」で再描画できる。
        """
        try:
            if not MATPLOTLIB_AVAILABLE or not results:
                return

            # 日本語フォント設定
            try:
                import matplotlib.font_manager as fm
                for font_name in ['MS Gothic', 'MS PGothic', 'Yu Gothic', 'Meiryo']:
                    try:
                        if fm.findfont(fm.FontProperties(family=font_name)):
                            plt.rcParams['font.family'] = font_name
                            break
                    except Exception:
                        continue
            except Exception:
                pass

            def fmt_date(d: str) -> str:
                try:
                    from datetime import datetime as _dt
                    dt = _dt.strptime(d, '%Y-%m-%d')
                    return f"{dt.month}/{dt.day}"
                except Exception:
                    return d

            x_labels = [fmt_date(r.start_date) for r in results]
            pr_vals = [r.pr for r in results]
            hdr_vals = [r.hdr for r in results]
            x = list(range(len(results)))

            # 妊娠率 95%信頼区間（Wilson スコア法）
            from math import sqrt as _sqrt
            def _wilson_ci(preg, n):
                if n == 0:
                    return None, None
                p = preg / n
                z = 1.96
                denom = 1 + z * z / n
                center = (p + z * z / (2 * n)) / denom
                margin = z * _sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / denom
                return max(0.0, (center - margin) * 100), min(100.0, (center + margin) * 100)

            _ci_lo, _ci_hi = [], []
            for r in results:
                lo, hi = _wilson_ci(r.preg, r.preg_eligible)
                _ci_lo.append(lo if lo is not None else r.pr)
                _ci_hi.append(hi if hi is not None else r.pr)
            _yerr_lo = [max(0.0, p - lo) for p, lo in zip(pr_vals, _ci_lo)]
            _yerr_hi = [max(0.0, hi - p) for p, hi in zip(pr_vals, _ci_hi)]

            # ---- ウィンドウ構築 ----
            graph_win = tk.Toplevel(self.root)
            full_title = f"{title} | {period_text}" if period_text else title
            graph_win.title(full_title)
            graph_win.transient(self.root)

            self.root.update_idletasks()
            _win_w, _win_h = 920, 620
            _screen_w = self.root.winfo_screenwidth()
            _screen_h = self.root.winfo_screenheight()
            _gx = self.root.winfo_x() + int(self.root.winfo_width() * 0.45)
            _gy = self.root.winfo_y() + 180
            if _gx + _win_w > _screen_w:
                _gx = max(0, _screen_w - _win_w - 10)
            if _gy + _win_h > _screen_h:
                _gy = max(0, _screen_h - _win_h - 40)
            graph_win.geometry(f"{_win_w}x{_win_h}+{_gx}+{_gy}")

            # タイトルラベル
            tk.Label(
                graph_win,
                text=full_title,
                font=(getattr(self, '_main_font_family', 'Meiryo UI'), 9),
                anchor=tk.W
            ).pack(fill=tk.X, padx=10, pady=(6, 0))

            # matplotlib キャンバス
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            figure = Figure(figsize=(11, 4.8), dpi=100)
            canvas = FigureCanvasTkAgg(figure, graph_win)
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=(6, 2))

            # ---- 補助線コントロールパネル ----
            ctrl_frame = ttk.LabelFrame(graph_win, text="目標補助線（0 = 非表示）")
            ctrl_frame.pack(fill=tk.X, padx=10, pady=(2, 8))

            _font = (getattr(self, '_main_font_family', 'Meiryo UI'), 9)

            ttk.Label(ctrl_frame, text="妊娠率 目標%:", font=_font).pack(side=tk.LEFT, padx=(12, 4))
            pr_target_var = tk.StringVar(value="0")
            pr_spin = ttk.Spinbox(ctrl_frame, from_=0, to=100, textvariable=pr_target_var,
                                  width=5, font=_font)
            pr_spin.pack(side=tk.LEFT, padx=(0, 16))

            ttk.Label(ctrl_frame, text="授精率 目標%:", font=_font).pack(side=tk.LEFT, padx=(0, 4))
            hdr_target_var = tk.StringVar(value="0")
            hdr_spin = ttk.Spinbox(ctrl_frame, from_=0, to=100, textvariable=hdr_target_var,
                                   width=5, font=_font)
            hdr_spin.pack(side=tk.LEFT, padx=(0, 16))

            ttk.Label(ctrl_frame, text="← 値を入力して「補助線を更新」を押してください",
                      font=(_font[0], _font[1] - 1), foreground='gray').pack(side=tk.LEFT, padx=(0, 12))

            update_btn = ttk.Button(ctrl_frame, text="補助線を更新")
            update_btn.pack(side=tk.LEFT, padx=(0, 12))

            ttk.Separator(ctrl_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=4)
            ttk.Button(ctrl_frame, text="📋 クリップボードにコピー",
                       command=lambda: self._copy_figure_to_clipboard(figure)).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(ctrl_frame, text="💾 画像を保存",
                       command=lambda: self._save_figure_as_file(figure)).pack(side=tk.LEFT, padx=(0, 6))

            # ---- 描画関数 ----
            def draw(pr_target: float = 0, hdr_target: float = 0):
                figure.clear()
                figure.patch.set_facecolor('white')
                ax = figure.add_subplot(111)
                ax.set_facecolor('white')
                hdr_above = [max(0.0, h - p) for h, p in zip(hdr_vals, pr_vals)]

                bar_w = 0.6
                ax.bar(x, pr_vals, width=bar_w, color='#1565c0', label='妊娠率')
                ax.bar(x, hdr_above, width=bar_w, bottom=pr_vals,
                       color='#90caf9', label='授精率（妊娠率との差）')

                # 妊娠率 95%信頼区間 エラーバー
                ax.errorbar(x, pr_vals,
                            yerr=[_yerr_lo, _yerr_hi],
                            fmt='none', color='#0d47a1', linewidth=1.2,
                            capsize=0, label='妊娠率 95%信頼区間')

                for i, (p, h) in enumerate(zip(pr_vals, hdr_vals)):
                    if h > 0:
                        ax.text(i, h + 0.5, f"{int(h)}%",
                                ha='center', va='bottom', fontsize=7, color='#455a64')
                    if p > 0:
                        ax.text(i, p / 2, f"{int(p)}%",
                                ha='center', va='center', fontsize=7,
                                color='white', fontweight='bold')

                if pr_target > 0:
                    ax.axhline(pr_target, color='#1565c0', linestyle='--', linewidth=1.5,
                               label=f'妊娠率目標 {int(pr_target)}%')
                    ax.text(len(x) - 0.5, pr_target + 0.5,
                            f"目標 {int(pr_target)}%", color='#1565c0',
                            fontsize=8, va='bottom', ha='right')
                if hdr_target > 0:
                    ax.axhline(hdr_target, color='#e65100', linestyle='--', linewidth=1.5,
                               label=f'授精率目標 {int(hdr_target)}%')
                    ax.text(len(x) - 0.5, hdr_target + 0.5,
                            f"目標 {int(hdr_target)}%", color='#e65100',
                            fontsize=8, va='bottom', ha='right')

                ax.set_xticks(x)
                ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=8)
                ax.set_xlabel("サイクル開始日", fontsize=9)
                ax.set_ylabel("%", fontsize=9)
                ax.set_title(full_title, fontsize=10, fontweight='bold', pad=10)
                y_max_data = max((max(hdr_vals) if hdr_vals else 0), pr_target, hdr_target)
                ax.set_ylim(0, max(100, y_max_data * 1.15))
                ax.grid(True, color='#e0e0e0', linewidth=0.8, axis='y')
                ax.set_axisbelow(True)
                ax.spines['top'].set_color('#cccccc')
                ax.spines['right'].set_color('#cccccc')
                ax.spines['left'].set_color('#cccccc')
                ax.spines['bottom'].set_color('#cccccc')
                ax.tick_params(colors='#555555')
                ax.legend(loc='upper left', fontsize=8, framealpha=0.7, edgecolor='none')

                try:
                    figure.tight_layout()
                except Exception:
                    pass
                canvas.draw()

            def on_update(_event=None):
                try:
                    pr_t = float(pr_target_var.get() or 0)
                except ValueError:
                    pr_t = 0
                try:
                    hdr_t = float(hdr_target_var.get() or 0)
                except ValueError:
                    hdr_t = 0
                draw(pr_t, hdr_t)

            update_btn.config(command=on_update)
            pr_spin.bind('<Return>', on_update)
            hdr_spin.bind('<Return>', on_update)

            # 初期描画（補助線なし）
            draw()

        except Exception as e:
            logging.error(f"妊娠率グラフ表示エラー: {e}", exc_info=True)

    def _execute_conception_rate_command(self, cow_type: str = "経産"):
        """受胎率コマンドを実行"""
        try:
            # 受胎率の種類を取得
            if not hasattr(self, 'conception_rate_type_entry') or not self.conception_rate_type_entry.winfo_exists():
                self.add_message(role="system", text="受胎率の種類が選択されていません")
                return

            rate_type = self.conception_rate_type_var.get().strip()
            # 未選択は「全体」（期間内合計）として実行（タイプ切替直後もそのまま実行可）
            if not rate_type:
                rate_type = "全体"

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
                        self.db, rate_type, start_date, end_date, condition_text, cow_type=cow_type
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
                                   end_date: Optional[str], condition_text: str,
                                   cow_type: str = "経産") -> Optional[Dict[str, Any]]:
        """
        受胎率を計算。
        cow_type: '経産'（lact>=1）または '未経産'（lact==0）
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
                
                # 経産/未経産フィルタ
                lact = cow.get('lact')
                if cow_type == "未経産":
                    if lact is None or lact != 0:
                        continue
                else:
                    # 経産牛のみ（LACT>=1）を対象とする
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
                entr_cutoff = (cow.get('entr') or '').strip()[:10]
                for ai_event in events:
                    event_date = ai_event.get('event_date')
                    event_number = ai_event.get('event_number')
                    # 登録日（entr）より前の授精は受胎率の母数に含めない
                    if entr_cutoff and event_date:
                        ed = (str(event_date)[:10])
                        if ed < entr_cutoff:
                            continue
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
                'cow_type': cow_type,
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
            
            if rate_type == "全体":
                return "全体"
            
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
            cow_type_label = result.get('cow_type', '経産')
            if rate_type == "全体":
                title = f"受胎率（{cow_type_label}）（全体）"
            else:
                title = f"受胎率（{cow_type_label}）（{rate_type}別）"
            if condition_text:
                title = f"{title}：{condition_text}"
            self._display_list_result_in_table(headers, rows, title, period=period)

            # ヒートマップ着色を適用
            self._apply_conception_rate_heatmap()

            # 受胎率結果テーブルにシングルクリックドリルダウンをバインド
            self.result_treeview.unbind("<Double-Button-1>")
            self.result_treeview.unbind("<ButtonRelease-1>")
            self.result_treeview.bind("<ButtonRelease-1>", self._on_conception_rate_row_click)

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

            cow_type_label = result.get("cow_type", "経産")
            if rate_type == "全体":
                graph_title = f"受胎率（{cow_type_label}）（全体）"
            else:
                graph_title = f"受胎率（{cow_type_label}）（{rate_type}別）"

            # グラフ表示用のデータ形式を作成（集計グラフと同じ見た目用に total で N= を付与）
            graph_data = {
                "type": "conception_rate",
                "x_axis": classifications,
                "y_axis": rates,
                "title": graph_title,
                "x_label": rate_type,
                "y_label": "受胎率（%）",
                "additional_data": {
                    "conceived": conceived_counts,
                    "not_conceived": not_conceived_counts,
                    "other": other_counts,
                    "total": total_counts,
                },
                "period": period,
            }

            self.current_graph_data = graph_data
            self.current_graph_command = graph_title

            # グラフをウィンドウで表示
            if MATPLOTLIB_AVAILABLE:
                self._draw_conception_rate_graph(graph_data)
                
        except Exception as e:
            logging.error(f"受胎率グラフ表示エラー: {e}", exc_info=True)
    
    def _draw_conception_rate_graph(self, graph_data: Dict[str, Any]):
        """
        受胎率の結果を別ウィンドウ（メイン画面の右側）に表示

        Args:
            graph_data: グラフ表示用のデータ
        """
        try:
            if not MATPLOTLIB_AVAILABLE:
                return

            x_labels = graph_data.get('x_axis', [])
            y_values = graph_data.get('y_axis', [])
            title = graph_data.get('title', '受胎率')
            x_label = graph_data.get('x_label', '分類')
            y_label = graph_data.get('y_label', '受胎率（%）')
            period = graph_data.get('period', '')
            command = self.current_graph_command or title

            if not x_labels or not y_values:
                return

            # 日本語フォントの設定
            try:
                import matplotlib.font_manager as fm
                japanese_fonts = ['MS Gothic', 'MS PGothic', 'Yu Gothic', 'Meiryo', 'Takao']
                for font_name in japanese_fonts:
                    try:
                        if fm.findfont(fm.FontProperties(family=font_name)):
                            plt.rcParams['font.family'] = font_name
                            break
                    except Exception:
                        continue
            except Exception:
                pass

            # グラフウィンドウをメイン画面の右側に配置
            graph_win = tk.Toplevel(self.root)
            full_title = f"{title} | 対象期間：{period}" if period else title
            graph_win.title(full_title)
            graph_win.transient(self.root)

            self.root.update_idletasks()
            _win_w, _win_h = 800, 620
            _screen_w = self.root.winfo_screenwidth()
            _screen_h = self.root.winfo_screenheight()
            # メイン画面の右半分に重ねて配置（表が左側に見えるように）
            _gx = self.root.winfo_x() + int(self.root.winfo_width() * 0.45)
            _gy = self.root.winfo_y() + 200
            # 画面外にはみ出す場合は補正
            if _gx + _win_w > _screen_w:
                _gx = max(0, _screen_w - _win_w - 10)
            if _gy + _win_h > _screen_h:
                _gy = max(0, _screen_h - _win_h - 40)
            graph_win.geometry(f"{_win_w}x{_win_h}+{_gx}+{_gy}")

            # コマンド表示
            cmd_text = f"コマンド: {command}"
            if period:
                cmd_text += f" | 対象期間：{period}"
            tk.Label(
                graph_win,
                text=cmd_text,
                font=(getattr(self, "_main_font_family", "Meiryo UI"), 9),
                anchor=tk.W,
            ).pack(fill=tk.X, padx=10, pady=5)

            # matplotlib Figure（集計グラフ _display_aggregate_graph と同一スタイル）
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            figure = Figure(figsize=(10, 6), dpi=100)
            canvas = FigureCanvasTkAgg(figure, graph_win)
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            ax = figure.add_subplot(111)
            figure.patch.set_facecolor("white")
            ax.set_facecolor("white")

            bars = ax.bar(range(len(x_labels)), y_values, **FALCON_BAR_CHART_KW)
            ax.set_xticks(range(len(x_labels)))

            add_data = graph_data.get("additional_data") or {}
            n_list = add_data.get("total") or []
            if len(n_list) == len(x_labels) and n_list:
                x_labels_display = [f"{lab}\n(N={n})" for lab, n in zip(x_labels, n_list)]
            else:
                x_labels_display = list(x_labels)

            ax.set_xticklabels(x_labels_display, rotation=0, ha="center")

            for bar, value in zip(bars, y_values):
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2.0,
                    height + 0.5,
                    f"{value:.1f}%",
                    ha="center",
                    va="bottom",
                    fontsize=FALCON_BAR_CHART_VALUE_FS,
                )

            ax.set_title(full_title, **FALCON_BAR_CHART_TITLE_KW)
            ax.set_xlabel(
                x_label,
                fontsize=FALCON_BAR_CHART_AXIS_LABEL_FS,
                fontweight="bold",
            )
            ax.set_ylabel(
                y_label,
                fontsize=FALCON_BAR_CHART_AXIS_LABEL_FS,
                fontweight="bold",
            )
            ax.grid(True, **FALCON_BAR_CHART_GRID_KW)
            ax.set_axisbelow(True)
            y_max = max(y_values) if y_values else 100
            ax.set_ylim(0, max(100, y_max * 1.2))

            try:
                figure.tight_layout()
            except Exception:
                pass

            canvas.draw()
            logging.info(f"[グラフ] 受胎率グラフ描画完了: {len(x_labels)}件のデータを表示")

            _btn_frame = ttk.Frame(graph_win)
            _btn_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
            _bf = (getattr(self, '_main_font_family', 'Meiryo UI'), 9)
            ttk.Button(_btn_frame, text="📋 クリップボードにコピー",
                       command=lambda: self._copy_figure_to_clipboard(figure)).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(_btn_frame, text="💾 画像を保存",
                       command=lambda: self._save_figure_as_file(figure)).pack(side=tk.LEFT, padx=(0, 6))

        except Exception as e:
            logging.error(f"[グラフ] 受胎率グラフ表示エラー: {e}", exc_info=True)
    
    def _apply_conception_rate_heatmap(self):
        """受胎率テーブルの行をパーセンタイルでヒートマップ着色する
        上位25%: 薄緑 / 下位25%: 薄赤 / 中間50%: 無色
        """
        try:
            if not hasattr(self, 'result_treeview') or not self.result_treeview.winfo_exists():
                return
            tree = self.result_treeview

            tree.tag_configure('cr_high',  background='#e8f5e9')  # 上位25%: 薄緑
            tree.tag_configure('cr_low',   background='#fce4ec')  # 下位25%: 薄赤
            tree.tag_configure('cr_total', background='#f0f0f0',
                               font=(self._main_font_family, 9, 'bold'))  # 合計行

            # 合計行を除いた受胎率の値を収集
            children = tree.get_children()
            rates: list[tuple] = []  # (item, rate)
            for item in children:
                values = tree.item(item, 'values')
                if not values:
                    continue
                if str(values[0]) == '合計':
                    tree.item(item, tags=('cr_total',))
                    continue
                rate_str = str(values[1]).replace('%', '').strip()
                try:
                    rates.append((item, float(rate_str)))
                except (ValueError, TypeError):
                    continue

            if not rates:
                return

            # 25・75パーセンタイルを計算
            sorted_rates = sorted(r for _, r in rates)
            n = len(sorted_rates)
            p25 = sorted_rates[max(0, int(n * 0.25) - 1)]
            p75 = sorted_rates[min(n - 1, int(n * 0.75))]

            for item, rate in rates:
                if rate >= p75:
                    tree.item(item, tags=('cr_high',))
                elif rate <= p25:
                    tree.item(item, tags=('cr_low',))
                # 中間: タグなし（無色）

        except Exception as e:
            logging.error(f"受胎率ヒートマップ適用エラー: {e}", exc_info=True)

    def _on_conception_rate_row_click(self, event):
        """受胎率結果テーブルの行をシングルクリックした時のドリルダウン処理"""
        try:
            tree = self.result_treeview
            # クリック位置の行を特定
            item = tree.identify_row(event.y)
            if not item:
                return

            values = tree.item(item, 'values')
            if not values:
                return

            classification = values[0] if len(values) > 0 else None
            if not classification or classification == '合計':
                return

            if not hasattr(self, 'conception_rate_events_by_classification'):
                return

            events_data = self.conception_rate_events_by_classification.get(classification, [])
            if not events_data:
                return

            self._show_conception_rate_events_window(classification, events_data)

        except Exception as e:
            logging.error(f"受胎率行クリックエラー: {e}", exc_info=True)

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

            # 結果による色分け
            treeview.tag_configure('result_P', background='#e8f5e9', foreground='#1b5e20')   # 妊娠：緑
            treeview.tag_configure('result_O', background='#fce4ec', foreground='#880e4f')   # 空胎：赤
            treeview.tag_configure('result_A', background='#fff3e0', foreground='#e65100')   # 流産：橙
            treeview.tag_configure('result_pending', background='#e3f2fd', foreground='#0d47a1')  # 連注/未鑑定：青
            
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
            
            # item_id → auto_id マップ（ダブルクリック時に使用。ソートのたびに再構築）
            auto_id_map: Dict[str, Any] = {}

            # 行レコード（ソート用）を構築
            row_records: List[Dict[str, Any]] = []
            _result_order_map = {'妊娠': 0, '空胎': 1, '流産': 2, '連注': 3, '未鑑定': 4}

            # イベントデータを走査して行レコードを作成
            for event_data in events_data:
                ai_event = event_data['ai_event']
                cow = event_data['cow']
                
                # ID
                cow_id = cow.get('cow_id', '')
                
                # 日付
                event_date = ai_event.get('event_date', '')
                event_dt_sort = None
                if event_date:
                    try:
                        event_dt_sort = datetime.strptime(event_date[:10], '%Y-%m-%d')
                    except (ValueError, TypeError):
                        pass
                
                # DIMを計算
                dim_display = ""
                dim_sort = -1
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
                                    dim_sort = dim
                    except Exception as e:
                        logging.debug(f"DIM計算エラー: {e}")
                
                # SIRE
                json_data_str = ai_event.get('json_data') or '{}'
                try:
                    json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                except Exception:
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
                insem_count_sort = -1
                try:
                    if str(insemination_count).strip() != '':
                        insem_count_sort = int(str(insemination_count).strip())
                except (ValueError, TypeError):
                    pass
                
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
                
                # 結果（妊娠/空胎/連注/流産/未鑑定）
                result = ""
                if event_data['conceived']:
                    result = "妊娠"
                elif event_data.get('other_reason') == 'R' or event_data.get('undetermined'):
                    # 7日以内の連続授精（other_reason='R'）および旧ロジック互換の未確定フラグは連注表示
                    result = "連注"
                elif event_data.get('other_reason') == 'undetermined_no_result':
                    result = "未鑑定"
                else:
                    # DBのoutcomeを参照し、S値に応じて日本語表示
                    ai_json = event_data.get('ai_event', {}).get('json_data') or {}
                    if isinstance(ai_json, str):
                        try:
                            ai_json = json.loads(ai_json)
                        except Exception:
                            ai_json = {}
                    outcome = (ai_json.get('outcome') or '').upper()
                    if not outcome or ai_json.get('_dc305_no_result'):
                        result = "未鑑定"
                    elif outcome == 'A':
                        result = "流産"
                    elif outcome == 'R':
                        result = "連注"
                    else:
                        result = "空胎"

                # 結果による色タグを決定
                if result == '妊娠':
                    result_tag = 'result_P'
                elif result == '空胎':
                    result_tag = 'result_O'
                elif result == '流産':
                    result_tag = 'result_A'
                else:  # 連注 / 未鑑定
                    result_tag = 'result_pending'

                row_records.append({
                    'values': (
                        cow_id,
                        event_date,
                        dim_display,
                        sire,
                        technician_name,
                        insemination_count,
                        ai_type_name,
                        result,
                    ),
                    'tag': result_tag,
                    'auto_id': cow.get('auto_id'),
                    'sort_date': event_dt_sort,
                    'sort_dim': dim_sort,
                    'sort_insem': insem_count_sort,
                    'sort_result': _result_order_map.get(result, 9),
                    'cow_id': cow_id,
                })

            # ソート状態: 列名 -> 0=元順 / 1=昇順 / 2=降順
            sort_state = {col: 0 for col in columns}
            _indexed = list(enumerate(row_records))

            def _sort_key(col_name: str, rec: Dict[str, Any]):
                if col_name == 'ID':
                    v = rec.get('cow_id') or ''
                    try:
                        return (0, int(v))
                    except (ValueError, TypeError):
                        return (1, str(v))
                if col_name == '日付':
                    dt = rec.get('sort_date')
                    return dt if dt is not None else datetime.min
                if col_name == 'DIM':
                    return rec.get('sort_dim', -1)
                if col_name == 'SIRE':
                    return rec['values'][3] or ''
                if col_name == '授精師':
                    return rec['values'][4] or ''
                if col_name == '授精回数':
                    return rec.get('sort_insem', -1)
                if col_name == '授精種類':
                    return rec['values'][6] or ''
                if col_name == '結果':
                    return rec.get('sort_result', 9)
                return ''

            def _populate(indexed_list):
                for iid in list(treeview.get_children()):
                    treeview.delete(iid)
                auto_id_map.clear()
                for _, rec in indexed_list:
                    item_id = treeview.insert('', tk.END, values=rec['values'], tags=(rec['tag'],))
                    auto_id_map[item_id] = rec['auto_id']

            def _on_header_click(col_name: str):
                prev = sort_state[col_name]
                for c in columns:
                    sort_state[c] = 0
                next_s = (prev + 1) % 3
                sort_state[col_name] = next_s
                for c in columns:
                    s = sort_state[c]
                    marker = ' ▲' if s == 1 else (' ▼' if s == 2 else '')
                    treeview.heading(c, text=f"{c}{marker}",
                                     command=lambda _c=c: _on_header_click(_c))
                if next_s == 0:
                    _populate(_indexed)
                else:
                    rev = (next_s == 2)
                    _populate(sorted(_indexed, key=lambda x: _sort_key(col_name, x[1]), reverse=rev))

            for col in columns:
                treeview.heading(col, text=col, command=lambda c=col: _on_header_click(c))

            _populate(_indexed)

            # 重複防止: auto_id → CowCardWindow
            _open_card_windows: Dict[int, Any] = {}

            # ダブルクリックで個体カードを独立ウィンドウで開く（妊娠率詳細と同方式）
            def on_row_double_click(event):
                item = treeview.identify_row(event.y)
                if not item:
                    return
                auto_id = auto_id_map.get(item)
                if not auto_id:
                    return
                # 既に開いていれば前面に
                existing = _open_card_windows.get(auto_id)
                if existing is not None:
                    try:
                        if existing.window.winfo_exists():
                            existing.window.lift()
                            existing.window.focus_set()
                            return
                    except Exception:
                        pass
                try:
                    from ui.cow_card_window import CowCardWindow
                    from constants import CONFIG_DEFAULT_DIR
                    event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
                    item_dict_path  = CONFIG_DEFAULT_DIR / "item_dictionary.json"
                    ccw = CowCardWindow(
                        parent=self.root,
                        db_handler=self.db,
                        formula_engine=self.formula_engine,
                        rule_engine=self.rule_engine,
                        event_dictionary_path=event_dict_path if event_dict_path.exists() else None,
                        item_dictionary_path=item_dict_path if item_dict_path.exists() else None,
                        cow_auto_id=auto_id,
                    )
                    _open_card_windows[auto_id] = ccw

                    def _on_card_close(_aid=auto_id):
                        _open_card_windows.pop(_aid, None)
                        ccw.window.destroy()

                    ccw.window.protocol("WM_DELETE_WINDOW", _on_card_close)
                    ccw.show()
                except Exception as ex:
                    logging.error(f"受胎率イベントリスト 個体カード起動エラー: {ex}", exc_info=True)
                    messagebox.showerror("エラー", str(ex), parent=window)

            treeview.bind("<Double-Button-1>", on_row_double_click)
            
        except Exception as e:
            logging.error(f"受胎率イベントリスト表示エラー: {e}", exc_info=True)
            from tkinter import messagebox
            messagebox.showerror("エラー", f"イベントリストの表示中にエラーが発生しました: {e}")

    # ------------------------------------------------------------------
    # 妊娠率（DIM別）
    # ------------------------------------------------------------------

    def _execute_pregnancy_rate_dim_command(self):
        """DIMレンジ別妊娠率コマンドを実行（DC305 Bredsumer互換）"""
        try:
            # VWP取得
            vwp_str = self.pregnancy_rate_vwp_var.get().strip() if hasattr(self, 'pregnancy_rate_vwp_var') else "50"
            try:
                vwp = int(vwp_str)
                if vwp < 1:
                    vwp = 50
            except ValueError:
                vwp = 50

            # 期間取得
            start_date = None
            end_date = None
            if hasattr(self, 'start_date_entry') and self.start_date_entry.winfo_exists():
                s = self.start_date_entry.get().strip()
                if s and not self.start_date_placeholder_shown:
                    start_date = s
            if hasattr(self, 'end_date_entry') and self.end_date_entry.winfo_exists():
                e = self.end_date_entry.get().strip()
                if e and not self.end_date_placeholder_shown:
                    end_date = e

            condition_text = ""
            if hasattr(self, 'pregnancy_rate_condition_entry') and self.pregnancy_rate_condition_entry.winfo_exists():
                if not self.pregnancy_rate_condition_placeholder_shown:
                    condition_text = self.pregnancy_rate_condition_entry.get().strip()

            period_text = f"{start_date} ～ {end_date}" if start_date and end_date else "全期間"
            self.add_message(role="system", text=f"妊娠率（DIM別）を計算しています… VWP={vwp}日 / 期間: {period_text}")
            self.result_notebook.select(0)

            def calculate_in_thread():
                try:
                    from modules.reproduction_analysis import ReproductionAnalysis
                    analyzer = ReproductionAnalysis(
                        db=self.db,
                        rule_engine=self.rule_engine,
                        formula_engine=self.formula_engine,
                        vwp=vwp
                    )
                    cow_filter = None
                    if condition_text:
                        def _cf(cow):
                            return self._evaluate_cow_condition_for_conception_rate(
                                cow, condition_text, self.db)
                        cow_filter = _cf
                    results = analyzer.analyze_by_dim(start_date, end_date, cow_filter=cow_filter)
                    self.root.after(0, lambda res=results, ct=condition_text: self._display_pregnancy_rate_dim_result(
                        res, vwp, period_text, ct
                    ))
                except Exception as ex:
                    err = str(ex)
                    logging.error(f"妊娠率（DIM別）計算エラー: {err}", exc_info=True)
                    self.root.after(0, lambda msg=err: self.add_message(
                        role="system", text=f"妊娠率（DIM別）の計算中にエラーが発生しました: {msg}"
                    ))

            thread = threading.Thread(target=calculate_in_thread, daemon=True)
            thread.start()

        except Exception as e:
            logging.error(f"妊娠率（DIM別）コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"妊娠率（DIM別）の計算中にエラーが発生しました: {e}")

    def _display_pregnancy_rate_dim_result(self, results, vwp: int, period_text: str,
                                           condition_text: str = ""):
        """DIMレンジ別妊娠率結果を表とグラフで表示"""
        try:
            if not results:
                self.add_message(role="system", text="該当するデータがありません")
                return

            def fmt_dim_range(dim_start: int, dim_end: int) -> str:
                if dim_end == 9999:
                    return f"DIM{dim_start}+"
                return f"DIM{dim_start}-{dim_end}"

            headers = ["DIMレンジ", "繁殖対象", "授精", "授精率%", "妊娠対象", "妊娠", "妊娠率%", "損耗"]
            rows = []

            total_br_el = 0
            total_bred = 0
            total_preg_eligible = 0
            total_preg = 0
            total_loss = 0

            for r in results:
                hdr_str = f"{int(r.hdr)}%" if r.br_el > 0 else "-"
                pr_str = f"{int(r.pr)}%" if r.preg_eligible > 0 else "-"
                rows.append([
                    fmt_dim_range(r.dim_start, r.dim_end),
                    str(r.br_el),
                    str(r.bred),
                    hdr_str,
                    str(r.preg_eligible) if r.preg_eligible > 0 else "-",
                    str(r.preg) if r.preg_eligible > 0 else "-",
                    pr_str,
                    str(r.loss) if r.preg_eligible > 0 else "-",
                ])
                total_br_el += r.br_el
                total_bred += r.bred
                total_preg_eligible += r.preg_eligible
                total_preg += r.preg
                total_loss += r.loss

            # 合計行
            total_hdr_str = f"{round(total_bred / total_br_el * 100)}%" if total_br_el > 0 else "-"
            total_pr_str = f"{round(total_preg / total_preg_eligible * 100)}%" if total_preg_eligible > 0 else "-"
            rows.append([
                "【合計】",
                str(total_br_el),
                str(total_bred),
                total_hdr_str,
                str(total_preg_eligible) if total_preg_eligible > 0 else "-",
                str(total_preg) if total_preg_eligible > 0 else "-",
                total_pr_str,
                str(total_loss) if total_preg_eligible > 0 else "-",
            ])

            title = f"妊娠率（DIM別）VWP={vwp}日"
            if condition_text:
                title = f"{title}：{condition_text}"
            self._display_list_result_in_table(headers, rows, title, period=period_text)

            # ヒートマップ着色（妊娠率%列: ≥20%→薄緑、未満→無色）
            self._apply_pregnancy_rate_heatmap()

            if MATPLOTLIB_AVAILABLE:
                self._display_pregnancy_rate_dim_graph(results, title=title, period_text=period_text)

        except Exception as e:
            logging.error(f"妊娠率（DIM別）結果表示エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"結果の表示中にエラーが発生しました: {e}")

    def _display_pregnancy_rate_dim_graph(self, results, title: str, period_text: str):
        """DIMレンジ別妊娠率グラフ（DC305 Bredsumer スタイル）

        - 積み上げ棒グラフ: 下=緑（Pregnancy Risk）、上=赤（Insemination Risk - PR）
        - 黒エラーバー: PR 95%信頼区間（Wilson 法）
        - 右Y軸・黒実線: Percent of Herd Still Open（累積空胎率）
        """
        try:
            if not MATPLOTLIB_AVAILABLE or not results:
                return

            from math import sqrt
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            import matplotlib.patches as mpatches
            import matplotlib.lines as mlines

            # 日本語フォント設定
            try:
                import matplotlib.font_manager as fm
                for _fn in ['Meiryo', 'MS Gothic', 'MS PGothic', 'Yu Gothic']:
                    try:
                        if fm.findfont(fm.FontProperties(family=_fn)):
                            plt.rcParams['font.family'] = _fn
                            break
                    except Exception:
                        continue
            except Exception:
                pass

            # ── データ準備 ──────────────────────────────────────────────
            # データがある（BR_EL > 0）レンジのみ対象
            active = [r for r in results if r.br_el > 0]
            if not active:
                self.add_message(role="system", text="表示できるデータがありません（全レンジ BR_EL=0）")
                return

            # X位置: DIMレンジの中点（実DIM値でプロット）
            x_mid = []
            for r in active:
                if r.dim_end == 9999:
                    x_mid.append(float(r.dim_start + 10))
                else:
                    x_mid.append((r.dim_start + r.dim_end) / 2.0)

            # 棒幅: DIMレンジ幅の80%
            _range_w = 21  # デフォルト21日レンジ
            if len(active) >= 2 and active[0].dim_end != 9999:
                _range_w = active[0].dim_end - active[0].dim_start + 1
            bar_w = _range_w * 0.80

            pr_vals  = [r.pr  for r in active]
            hdr_vals = [r.hdr for r in active]
            # 赤部分 = HDR - PR（授精したが非妊娠）
            red_vals = [max(0.0, h - p) for h, p in zip(hdr_vals, pr_vals)]

            # PR 95%信頼区間（Wilson スコア法）
            def _wilson_ci(preg: int, n: int):
                if n == 0:
                    return None, None
                p = preg / n
                z = 1.96
                denom = 1 + z * z / n
                center = (p + z * z / (2 * n)) / denom
                margin = z * sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / denom
                return max(0.0, (center - margin) * 100), min(100.0, (center + margin) * 100)

            ci_lo, ci_hi = [], []
            for r in active:
                lo, hi = _wilson_ci(r.preg, r.preg_eligible)
                ci_lo.append(lo if lo is not None else r.pr)
                ci_hi.append(hi if hi is not None else r.pr)

            yerr_lo = [max(0.0, p - lo) for p, lo in zip(pr_vals, ci_lo)]
            yerr_hi = [max(0.0, hi - p) for p, hi in zip(pr_vals, ci_hi)]

            # Percent of Herd Still Open（累積空胎率）
            # 各レンジで妊娠した割合を乗算して残存率を計算
            still_open = [100.0]
            for r in active:
                if r.preg_eligible > 0:
                    p_preg = r.preg / r.preg_eligible
                else:
                    p_preg = 0.0
                still_open.append(still_open[-1] * (1.0 - p_preg))

            # Still Open の X 座標: 最初のレンジ開始 → 各レンジ終端
            x_open = [float(active[0].dim_start)]
            for r in active:
                x_open.append(float(r.dim_end if r.dim_end != 9999 else r.dim_start + _range_w))

            # ── ウィンドウ構築 ───────────────────────────────────────────
            graph_win = tk.Toplevel(self.root)
            full_title = f"{title} | {period_text}" if period_text else title
            graph_win.title(full_title)
            graph_win.transient(self.root)

            self.root.update_idletasks()
            _win_w, _win_h = 980, 600
            _sw = self.root.winfo_screenwidth()
            _sh = self.root.winfo_screenheight()
            _gx = min(self.root.winfo_x() + int(self.root.winfo_width() * 0.45), _sw - _win_w - 10)
            _gy = min(self.root.winfo_y() + 160, _sh - _win_h - 40)
            graph_win.geometry(f"{_win_w}x{_win_h}+{max(0,_gx)}+{max(0,_gy)}")

            # matplotlib キャンバス
            figure = Figure(figsize=(11.5, 5.2), dpi=100)
            mpl_canvas = FigureCanvasTkAgg(figure, graph_win)
            mpl_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=(6, 2))

            # 目標補助線フレーム
            ctrl_frame = ttk.Frame(graph_win)
            ctrl_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
            _font = (getattr(self, '_main_font_family', 'Meiryo UI'), 9)
            ttk.Label(ctrl_frame, text="妊娠率 目標%:", font=_font).pack(side=tk.LEFT, padx=(12, 4))
            pr_target_var = tk.StringVar(value="0")
            ttk.Spinbox(ctrl_frame, from_=0, to=100, textvariable=pr_target_var,
                        width=5, font=_font).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Label(ctrl_frame, text="授精率 目標%:", font=_font).pack(side=tk.LEFT, padx=(0, 4))
            hdr_target_var = tk.StringVar(value="0")
            ttk.Spinbox(ctrl_frame, from_=0, to=100, textvariable=hdr_target_var,
                        width=5, font=_font).pack(side=tk.LEFT, padx=(0, 10))

            # ── 描画関数 ─────────────────────────────────────────────────
            def draw_graph():
                figure.clear()
                figure.patch.set_facecolor('white')
                ax1 = figure.add_subplot(111)
                ax1.set_facecolor('white')
                ax2 = ax1.twinx()

                # 積み上げ棒グラフ
                # 下段: 深ブルー（妊娠率）
                ax1.bar(x_mid, pr_vals, width=bar_w,
                        color='#1565c0', label='妊娠率', zorder=3)
                # 上段: 淡スカイブルー（授精率との差分）
                ax1.bar(x_mid, red_vals, width=bar_w, bottom=pr_vals,
                        color='#90caf9', label='授精率（差分）', zorder=3)

                # 妊娠率 95%信頼区間 エラーバー（ネイビー）
                ax1.errorbar(
                    x_mid, pr_vals,
                    yerr=[yerr_lo, yerr_hi],
                    fmt='none', color='#0d47a1', linewidth=1.4,
                    capsize=0, zorder=5, label='妊娠率 95%信頼区間'
                )

                # 空胎残存率（右Y軸・オレンジ実線）
                ax2.plot(x_open, still_open,
                         color='#e65100', linewidth=2.2,
                         label='空胎残存率', zorder=6)

                # 目標補助線
                try:
                    pt = int(pr_target_var.get())
                    if pt > 0:
                        ax1.axhline(pt, color='#0d47a1', linestyle='--',
                                    linewidth=1.2, label=f'妊娠率目標 {pt}%')
                except (ValueError, TypeError):
                    pass
                try:
                    ht = int(hdr_target_var.get())
                    if ht > 0:
                        ax1.axhline(ht, color='#e65100', linestyle='--',
                                    linewidth=1.2, label=f'授精率目標 {ht}%')
                except (ValueError, TypeError):
                    pass

                # ── 軸設定 ───────────────────────────────────────────────
                ax1.set_xlabel('搾乳日数（DIM）', fontsize=10)
                ax1.set_ylabel('割合（%）', fontsize=10)
                ax1.set_ylim(0, 105)
                ax1.yaxis.set_major_locator(plt.MultipleLocator(10))
                ax1.grid(axis='y', color='#e0e0e0', linewidth=0.8, zorder=0)
                ax1.spines['top'].set_color('#cccccc')
                ax1.spines['left'].set_color('#cccccc')
                ax1.spines['bottom'].set_color('#cccccc')
                ax1.tick_params(colors='#555555')

                # 右Y軸（オレンジラベル）
                _open_max = max(still_open) if still_open else 100
                ax2.set_ylim(0, _open_max * 1.06)
                ax2.set_ylabel('空胎残存率（%）',
                               fontsize=9, color='#e65100', rotation=270, labelpad=14)
                ax2.tick_params(axis='y', colors='#e65100', labelsize=8)
                ax2.spines['right'].set_color('#e65100')
                ax2.spines['right'].set_linewidth(1.2)

                # X軸: 全体範囲を少し広げて見やすく
                _x_min = max(0, (active[0].dim_start - _range_w * 1.5))
                _last  = active[-1]
                _x_max = (_last.dim_end if _last.dim_end != 9999 else _last.dim_start + _range_w * 2) + _range_w * 1.5
                ax1.set_xlim(_x_min, _x_max)
                ax2.set_xlim(_x_min, _x_max)

                # X軸メモリ: 40日ごと
                import numpy as _np
                _xticks = _np.arange(
                    int(_x_min // 40 + 1) * 40,
                    int(_x_max // 40 + 1) * 40,
                    40
                )
                ax1.set_xticks(_xticks)
                ax1.tick_params(axis='x', labelsize=9)

                # タイトル
                ax1.set_title(full_title, fontsize=10, pad=8)

                # 凡例（上部に1行で）
                _leg_handles = [
                    mpatches.Patch(color='#90caf9', label='授精率（差分）'),
                    mlines.Line2D([], [], color='#0d47a1', linewidth=1.4, label='妊娠率 95%信頼区間'),
                    mpatches.Patch(color='#1565c0', label='妊娠率'),
                    mlines.Line2D([], [], color='#e65100', linewidth=2.2, label='空胎残存率'),
                ]
                ax1.legend(handles=_leg_handles, loc='upper center',
                           bbox_to_anchor=(0.5, 1.13), ncol=4,
                           fontsize=8.5, framealpha=0.9,
                           handlelength=1.4, handleheight=0.9)

                figure.tight_layout(rect=[0, 0, 1, 0.95])
                mpl_canvas.draw()

            draw_graph()

            ttk.Button(ctrl_frame, text="補助線を更新",
                       command=draw_graph).pack(side=tk.LEFT, padx=(10, 0))

            ttk.Separator(ctrl_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=4)
            ttk.Button(ctrl_frame, text="📋 クリップボードにコピー",
                       command=lambda: self._copy_figure_to_clipboard(figure)).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(ctrl_frame, text="💾 画像を保存",
                       command=lambda: self._save_figure_as_file(figure)).pack(side=tk.LEFT, padx=(0, 6))

        except Exception as e:
            logging.error(f"妊娠率（DIM別）グラフ表示エラー: {e}", exc_info=True)

    # ─────────────────────────────────────────────────────────────────
    #  グラフ ユーティリティ
    # ─────────────────────────────────────────────────────────────────

    def _copy_figure_to_clipboard(self, figure):
        """matplotlib Figure を PNG 形式で Windows クリップボードにコピーする。
        PowerShell / System.Windows.Forms 経由でコピーし、バックグラウンドスレッドで実行。
        """
        import os
        import subprocess

        def _do_copy():
            tmp_path = None
            try:
                import tempfile
                buf = io.BytesIO()
                figure.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
                buf.seek(0)
                fd, tmp_path = tempfile.mkstemp(suffix='.png')
                with os.fdopen(fd, 'wb') as f:
                    f.write(buf.read())
                # PowerShell で .NET クリップボードにコピー
                ps_cmd = (
                    "Add-Type -AssemblyName System.Windows.Forms; "
                    "Add-Type -AssemblyName System.Drawing; "
                    f"$img = [System.Drawing.Image]::FromFile('{tmp_path}'); "
                    "[System.Windows.Forms.Clipboard]::SetImage($img); "
                    "$img.Dispose();"
                )
                result = subprocess.run(
                    ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_cmd],
                    timeout=20, capture_output=True
                )
                if result.returncode == 0:
                    self.root.after(0, lambda: self.add_message(
                        role="system", text="グラフをクリップボードにコピーしました。PowerPoint 等に貼り付けできます。"))
                else:
                    err = result.stderr.decode('utf-8', errors='replace').strip()
                    logging.error(f"クリップボードコピーエラー: {err}")
                    self.root.after(0, lambda: self.add_message(
                        role="system", text=f"クリップボードへのコピーに失敗しました: {err[:120]}"))
            except Exception as ex:
                logging.error(f"_copy_figure_to_clipboard エラー: {ex}", exc_info=True)
                self.root.after(0, lambda: self.add_message(
                    role="system", text=f"クリップボードへのコピー中にエラーが発生しました: {ex}"))
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass

        threading.Thread(target=_do_copy, daemon=True).start()

    def _save_figure_as_file(self, figure):
        """matplotlib Figure をファイルダイアログ経由で PNG / SVG として保存する。"""
        file_path = filedialog.asksaveasfilename(
            title="グラフを保存",
            defaultextension=".png",
            filetypes=[("PNG 画像", "*.png"), ("SVG ベクター", "*.svg"), ("全てのファイル", "*.*")],
        )
        if not file_path:
            return
        try:
            fmt = 'svg' if file_path.lower().endswith('.svg') else 'png'
            figure.savefig(file_path, format=fmt, dpi=150, bbox_inches='tight', facecolor='white')
            self.add_message(role="system", text=f"グラフを保存しました: {file_path}")
        except Exception as ex:
            logging.error(f"_save_figure_as_file エラー: {ex}", exc_info=True)
            self.add_message(role="system", text=f"グラフの保存中にエラーが発生しました: {ex}")

    # ------------------------------------------------------------------
    # 妊娠率表: ダブルクリック → サイクル個体詳細ウィンドウ
    # ------------------------------------------------------------------

    def _on_pregnancy_rate_row_double_click(self, event):
        """妊娠率表の行をダブルクリック → そのサイクルの個体一覧を表示"""
        if not hasattr(self, '_pr_cycle_results') or not self._pr_cycle_results:
            return
        tree = self.result_treeview
        item = tree.identify_row(event.y)
        if not item:
            return
        tree.selection_set(item)

        children = list(tree.get_children())
        try:
            row_index = children.index(item)
        except ValueError:
            return

        # 合計行（最後の行）は対象外
        if row_index >= len(self._pr_cycle_results):
            return

        r = self._pr_cycle_results[row_index]

        def fetch_in_thread():
            try:
                from modules.reproduction_analysis import ReproductionAnalysis, Cycle
                analyzer = ReproductionAnalysis(
                    db=self.db,
                    rule_engine=self.rule_engine,
                    formula_engine=self.formula_engine,
                    vwp=getattr(self, '_pr_vwp', 50)
                )
                cycle = Cycle(
                    cycle_number=r.cycle_number,
                    start_date=r.start_date,
                    end_date=r.end_date
                )
                # 経産牛のみ対象（未経産牛 lact=0 or None は除外）
                cows = [c for c in self.db.get_all_cows() if (c.get('lact') or 0) >= 1]
                cf = getattr(self, '_pr_cow_filter', None)
                if cf:
                    cows = [c for c in cows if cf(c)]
                today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                details = analyzer.get_cycle_cow_details(cycle, cows, today)
                self.root.after(0, lambda: self._show_pregnancy_rate_cycle_detail(r, details))
            except Exception as ex:
                logging.error(f"サイクル詳細取得エラー: {ex}", exc_info=True)
                self.root.after(0, lambda msg=str(ex): self.add_message(
                    role="system", text=f"サイクル詳細の取得中にエラーが発生しました: {msg}"
                ))

        threading.Thread(target=fetch_in_thread, daemon=True).start()

    def _show_pregnancy_rate_cycle_detail(self, r, details):
        """サイクル別個体詳細ウィンドウを表示"""
        try:
            def fmt_date(d):
                try:
                    dt = datetime.strptime(d, '%Y-%m-%d')
                    return f"{dt.month}/{dt.day}"
                except Exception:
                    return d

            win_title = f"サイクル詳細  {fmt_date(r.start_date)} ～ {fmt_date(r.end_date)}"
            win = tk.Toplevel(self.root)
            win.title(win_title)
            win.transient(self.root)

            self.root.update_idletasks()
            wx = self.root.winfo_x() + 80
            wy = self.root.winfo_y() + 80
            win.geometry(f"820x500+{wx}+{wy}")

            _font = (getattr(self, '_main_font_family', 'Meiryo UI'), 9)

            # 農場設定から授精師・授精種類コードを読み込む
            inseminator_codes: dict = {}
            insemination_type_codes: dict = {}
            try:
                from settings_manager import SettingsManager
                _sm = SettingsManager(self.farm_path)
                inseminator_codes = _sm.get('inseminator_codes', {}) or {}
                _ins_file = self.farm_path / "insemination_settings.json"
                if _ins_file.exists():
                    with open(_ins_file, 'r', encoding='utf-8') as _f:
                        insemination_type_codes = json.load(_f).get('insemination_types', {}) or {}
                if not insemination_type_codes:
                    insemination_type_codes = _sm.get('insemination_type_codes', {}) or {}
            except Exception:
                pass

            def _resolve_technician(code: str) -> str:
                if not code:
                    return ''
                name = inseminator_codes.get(code)
                if name is None and code.isdigit():
                    name = inseminator_codes.get(int(code))
                return name if name else code

            def _resolve_ai_type(code: str) -> str:
                if not code:
                    return ''
                name = insemination_type_codes.get(code)
                if name is None:
                    for k, v in insemination_type_codes.items():
                        if str(k).upper() == code.upper():
                            name = v
                            break
                return name if name else code

            # サマリー行
            parts = [f"繁殖対象 {r.br_el}頭", f"授精 {r.bred}頭"]
            if r.preg_eligible > 0:
                parts.append(f"妊娠 {r.preg}頭")
                parts.append(f"妊娠率 {int(r.pr)}%")
            tk.Label(win, text="  ".join(parts), font=_font, anchor=tk.W).pack(
                fill=tk.X, padx=10, pady=(8, 4)
            )

            # Treeview
            columns = ("ID", "産次", "DIM", "SIRE", "授精師", "授精種類", "授精回数", "結果")
            frame = ttk.Frame(win)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 2))

            vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            tree = ttk.Treeview(frame, columns=columns, show="headings",
                                yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            vsb.config(command=tree.yview)
            hsb.config(command=tree.xview)
            tree.pack(fill=tk.BOTH, expand=True)

            col_conf = {"ID": (65, tk.CENTER), "産次": (50, tk.CENTER),
                        "DIM": (50, tk.CENTER), "SIRE": (100, tk.W),
                        "授精師": (80, tk.W), "授精種類": (90, tk.W),
                        "授精回数": (65, tk.CENTER), "結果": (130, tk.W)}
            for col in columns:
                w, anc = col_conf[col]
                tree.column(col, width=w, anchor=anc, minwidth=w)

            tree.tag_configure('妊娠',         background='#e8f5e9', foreground='#1b5e20')
            tree.tag_configure('流産',         background='#fff3e0', foreground='#e65100')
            tree.tag_configure('授精済(未確定)', background='#e3f2fd', foreground='#0d47a1')
            tree.tag_configure('空胎',         background='#fce4ec', foreground='#880e4f')

            # ソート状態: col -> 0=元順 / 1=昇順 / 2=降順
            sort_state = {col: 0 for col in columns}
            _indexed = list(enumerate(details))  # 元インデックス付きリスト
            auto_id_map = {}

            _cat_order = {'妊娠': 0, '流産': 1, '授精済(未確定)': 2, '空胎': 3, '未授精': 4}

            def _sort_key(col, d):
                if col == 'ID':
                    v = d.get('cow_id') or ''
                    try:
                        return (0, int(v))
                    except (ValueError, TypeError):
                        return (1, v)
                elif col == '産次':
                    v = d.get('lact', '')
                    try:
                        return int(v) if v != '' else -1
                    except (ValueError, TypeError):
                        return -1
                elif col == 'DIM':
                    v = d.get('dim', '')
                    try:
                        return int(v) if v != '' else -1
                    except (ValueError, TypeError):
                        return -1
                elif col == '授精回数':
                    v = d.get('insemination_count', '')
                    try:
                        return int(v) if v != '' else -1
                    except (ValueError, TypeError):
                        return -1
                elif col == 'SIRE':
                    return d.get('sire', '') or ''
                elif col == '授精師':
                    return _resolve_technician(d.get('technician_code', ''))
                elif col == '授精種類':
                    return _resolve_ai_type(d.get('ai_type_code', ''))
                else:  # 結果
                    return _cat_order.get(d.get('category', ''), 9)

            def _populate(indexed_list):
                for iid in list(tree.get_children()):
                    tree.delete(iid)
                auto_id_map.clear()
                for _, d in indexed_list:
                    tag = d['category']
                    iid = tree.insert('', tk.END, values=(
                        d.get('cow_id', ''),
                        d.get('lact', ''),
                        d.get('dim', ''),
                        d.get('sire', ''),
                        _resolve_technician(d.get('technician_code', '')),
                        _resolve_ai_type(d.get('ai_type_code', '')),
                        d.get('insemination_count', ''),
                        d['category'],
                    ), tags=(tag,))
                    auto_id_map[iid] = d.get('auto_id')

            def _on_header_click(col):
                prev = sort_state[col]
                for c in columns:
                    sort_state[c] = 0
                next_s = (prev + 1) % 3
                sort_state[col] = next_s
                # ヘッダーラベル更新
                for c in columns:
                    s = sort_state[c]
                    marker = ' ▲' if s == 1 else (' ▼' if s == 2 else '')
                    tree.heading(c, text=f"{c}{marker}",
                                 command=lambda _c=c: _on_header_click(_c))
                if next_s == 0:
                    _populate(_indexed)
                else:
                    rev = (next_s == 2)
                    _populate(sorted(_indexed, key=lambda x: _sort_key(col, x[1]), reverse=rev))

            for col in columns:
                tree.heading(col, text=col, command=lambda c=col: _on_header_click(c))

            _populate(_indexed)

            tk.Label(win, text=f"合計 {len(details)} 頭（ダブルクリックで個体カード）",
                     font=(_font[0], _font[1] - 1), foreground='gray', anchor=tk.E).pack(
                fill=tk.X, padx=12, pady=(2, 6)
            )

            # 重複防止用: auto_id → CowCardWindow のマップ
            _open_card_windows: Dict[int, Any] = {}

            def on_row_double_click(event):
                sel = tree.selection()
                if not sel:
                    return
                auto_id = auto_id_map.get(sel[0])
                if not auto_id:
                    return
                # 既に開いていれば前面に出す
                existing = _open_card_windows.get(auto_id)
                if existing is not None:
                    try:
                        if existing.window.winfo_exists():
                            existing.window.lift()
                            existing.window.focus_set()
                            return
                    except Exception:
                        pass
                # 新規ウィンドウを開く
                try:
                    from ui.cow_card_window import CowCardWindow
                    from constants import CONFIG_DEFAULT_DIR
                    event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
                    item_dict_path  = CONFIG_DEFAULT_DIR / "item_dictionary.json"
                    ccw = CowCardWindow(
                        parent=self.root,
                        db_handler=self.db,
                        formula_engine=self.formula_engine,
                        rule_engine=self.rule_engine,
                        event_dictionary_path=event_dict_path if event_dict_path.exists() else None,
                        item_dictionary_path=item_dict_path if item_dict_path.exists() else None,
                        cow_auto_id=auto_id,
                    )
                    _open_card_windows[auto_id] = ccw

                    def _on_card_close(_aid=auto_id):
                        _open_card_windows.pop(_aid, None)
                        ccw.window.destroy()

                    ccw.window.protocol("WM_DELETE_WINDOW", _on_card_close)
                    ccw.show()
                except Exception as ex:
                    logging.error(f"個体カードウィンドウ起動エラー: {ex}", exc_info=True)

            tree.bind('<Double-Button-1>', on_row_double_click)

        except Exception as e:
            logging.error(f"サイクル詳細ウィンドウ表示エラー: {e}", exc_info=True)

    def _apply_pregnancy_rate_heatmap(self):
        """妊娠率テーブルの行をパーセンタイルでヒートマップ着色する
        上位25%: 薄緑 / 下位25%: 薄赤 / 中間50%: 無色
        """
        try:
            if not hasattr(self, 'result_treeview') or not self.result_treeview.winfo_exists():
                return
            tree = self.result_treeview

            cols = list(tree['columns'])
            try:
                pr_idx = cols.index('妊娠率%')
            except ValueError:
                return

            tree.tag_configure('pr_high',  background='#e8f5e9')  # 上位25%: 薄緑
            tree.tag_configure('pr_low',   background='#fce4ec')  # 下位25%: 薄赤
            tree.tag_configure('pr_total', background='#f0f0f0',
                               font=(self._main_font_family, 9, 'bold'))  # 合計行

            # 合計行を除いた妊娠率の値を収集（"-" は除外）
            rates: list[tuple] = []
            for item in tree.get_children():
                values = tree.item(item, 'values')
                if not values or len(values) <= pr_idx:
                    continue
                if str(values[0]) == '【合計】':
                    tree.item(item, tags=('pr_total',))
                    continue
                rate_str = str(values[pr_idx]).replace('%', '').strip()
                if rate_str == '-':
                    continue
                try:
                    rates.append((item, float(rate_str)))
                except (ValueError, TypeError):
                    continue

            if not rates:
                return

            # 25・75パーセンタイルを計算
            sorted_rates = sorted(r for _, r in rates)
            n = len(sorted_rates)
            p25 = sorted_rates[max(0, int(n * 0.25) - 1)]
            p75 = sorted_rates[min(n - 1, int(n * 0.75))]

            for item, rate in rates:
                if rate >= p75:
                    tree.item(item, tags=('pr_high',))
                elif rate <= p25:
                    tree.item(item, tags=('pr_low',))
                # 中間: タグなし（無色）

        except Exception as e:
            logging.error(f"妊娠率ヒートマップ適用エラー: {e}", exc_info=True)

    def _on_pregnancy_rate_row_click(self, event):
        """妊娠率表の行をシングルクリック → そのサイクルの個体一覧を表示"""
        if not hasattr(self, '_pr_cycle_results') or not self._pr_cycle_results:
            return
        tree = self.result_treeview
        item = tree.identify_row(event.y)
        if not item:
            return

        children = list(tree.get_children())
        try:
            row_index = children.index(item)
        except ValueError:
            return

        # 合計行（データ行数を超えるインデックス）は対象外
        if row_index >= len(self._pr_cycle_results):
            return

        r = self._pr_cycle_results[row_index]

        def fetch_in_thread():
            try:
                from modules.reproduction_analysis import ReproductionAnalysis, Cycle
                analyzer = ReproductionAnalysis(
                    db=self.db,
                    rule_engine=self.rule_engine,
                    formula_engine=self.formula_engine,
                    vwp=getattr(self, '_pr_vwp', 50)
                )
                cycle = Cycle(
                    cycle_number=r.cycle_number,
                    start_date=r.start_date,
                    end_date=r.end_date
                )
                cows = [c for c in self.db.get_all_cows() if (c.get('lact') or 0) >= 1]
                cf = getattr(self, '_pr_cow_filter', None)
                if cf:
                    cows = [c for c in cows if cf(c)]
                today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                details = analyzer.get_cycle_cow_details(cycle, cows, today)
                self.root.after(0, lambda: self._show_pregnancy_rate_cycle_detail(r, details))
            except Exception as ex:
                logging.error(f"サイクル詳細取得エラー: {ex}", exc_info=True)
                self.root.after(0, lambda msg=str(ex): self.add_message(
                    role="system", text=f"サイクル詳細の取得中にエラーが発生しました: {msg}"
                ))

        threading.Thread(target=fetch_in_thread, daemon=True).start()

    def _apply_pr_table_styling(self, headers: list):
        """妊娠率表の後処理: キー列ヘッダー着色 + 合計行（最終行）太字"""
        try:
            if not hasattr(self, 'result_treeview') or not self.result_treeview.winfo_exists():
                return
            tree = self.result_treeview

            # ── 合計行（最終行）を太字に ──
            children = tree.get_children()
            if children:
                font_bold = (self._main_font_family, 9, 'bold')
                tree.tag_configure('pr_total', font=font_bold, background='#f0f0f0')
                tree.item(children[-1], tags=('pr_total',))

            # ── キー列ヘッダーを PhotoImage で着色 ──
            HEADER_COLORS = {
                "授精率%": "#c8e6c9",   # 薄い緑
                "受胎率%": "#e1bee7",   # 薄い紫
                "妊娠率%": "#bbdefb",   # 薄い青
            }
            # 画像は GC されないよう self に保持
            if not hasattr(self, '_pr_header_images'):
                self._pr_header_images = {}
            else:
                self._pr_header_images.clear()

            # レイアウト確定後に幅を取得するため update_idletasks を挟む
            tree.update_idletasks()

            for col in headers:
                color = HEADER_COLORS.get(col)
                if not color:
                    continue
                try:
                    col_width = tree.column(col, 'width')
                    # PhotoImage で単色画像を作成（列幅 × ヘッダー高さ）
                    img = tk.PhotoImage(width=col_width, height=20)
                    row_str = '{' + ' '.join([color] * col_width) + '}'
                    img.put(' '.join([row_str] * 20))
                    self._pr_header_images[col] = img
                    tree.heading(col, image=img, compound='center')
                except Exception as e:
                    logging.debug(f"ヘッダー着色スキップ ({col}): {e}")

        except Exception as e:
            logging.error(f"妊娠率表スタイル適用エラー: {e}", exc_info=True)

