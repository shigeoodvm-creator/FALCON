"""
FALCON2 - 散布図描画 Mixin
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
import sys

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

class ScatterPlotMixin:
    """Mixin: FALCON2 - 散布図描画 Mixin"""

    def _copy_figure_image_to_clipboard(self, figure: Figure, dpi: int = 220):
        """matplotlib Figure を画像としてクリップボードにコピー（PowerPoint貼り付け向け）。"""
        try:
            buf = io.BytesIO()
            figure.savefig(
                buf,
                format="png",
                dpi=dpi,
                bbox_inches="tight",
                facecolor=figure.get_facecolor(),
                edgecolor="none",
            )
            buf.seek(0)
            data_png = buf.getvalue()
        except Exception as e:
            logging.error(f"散布図画像の生成に失敗: {e}", exc_info=True)
            messagebox.showerror("コピー失敗", f"散布図画像の生成に失敗しました。\n{e}")
            return

        if sys.platform != "win32":
            messagebox.showinfo(
                "画像コピー",
                "この環境では画像のクリップボードコピーに未対応です。\n"
                "「画像として保存」をご利用ください。"
            )
            return

        try:
            import win32clipboard  # type: ignore[reportMissingModuleSource]
            from PIL import Image
        except ImportError:
            messagebox.showinfo(
                "画像コピー",
                "画像コピーには追加パッケージが必要です。\n\n"
                "pip install pywin32 Pillow\n\n"
                "インストール後、アプリを再起動して再度お試しください。"
            )
            return

        try:
            img = Image.open(io.BytesIO(data_png))
            output = io.BytesIO()
            img.convert("RGB").save(output, "BMP")
            dib_data = output.getvalue()[14:]  # BMPヘッダー(14byte)を除去
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib_data)
            win32clipboard.CloseClipboard()
            messagebox.showinfo(
                "コピー完了",
                "散布図をクリップボードにコピーしました。\nPowerPointにそのまま貼り付けできます。"
            )
        except Exception as e:
            logging.error(f"散布図のクリップボードコピーに失敗: {e}", exc_info=True)
            messagebox.showerror("コピー失敗", f"クリップボードへのコピーに失敗しました。\n{e}")
    def _parse_scatter_plot_axes(self, query: str) -> Tuple[Optional[str], Optional[str]]:
        """
        ユーザー入力から散布図の縦軸・横軸をパース
        
        Args:
            query: ユーザー入力
        
        Returns:
            (横軸名, 縦軸名) のタプル。見つからない場合は (None, None)
        """
        # 「縦軸」「横軸」という語が含まれる場合
        if "縦軸" in query and "横軸" in query:
            # 「縦軸 X、横軸 Y」または「横軸 X、縦軸 Y」の形式を想定
            # 縦軸を抽出
            y_match = re.search(r'縦軸\s*([^、，,横軸]+)', query)
            y_axis = y_match.group(1).strip() if y_match else None
            
            # 横軸を抽出
            x_match = re.search(r'横軸\s*([^、，,縦軸散布図の]+)', query)
            x_axis = x_match.group(1).strip() if x_match else None
            
            if x_axis and y_axis:
                return (x_axis, y_axis)
        
        return (None, None)
    
    def _calculate_first_ai_days(self, cow_auto_id: int, db: DBHandler) -> Optional[int]:
        """
        初回授精日数を計算（授精回数1回目のAI/ETのみ対象）。
        分娩日から、授精回数1回目として記録された AI/ET イベント日までの日数。
        授精回数1回目のイベントがない個体（途中導入など）はNoneを返し集計から除外する。
        
        Args:
            cow_auto_id: 牛の auto_id
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            初回授精日数（日）。計算できない場合はNone
        """
        try:
            cow = db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                return None
            
            clvd = cow.get('clvd')
            if not clvd:
                return None
            
            try:
                clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
            except Exception:
                return None
            
            events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
            ai_et_events = []
            for event in events:
                event_number = event.get('event_number')
                event_date_str = event.get('event_date')
                if event_number not in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET] or not event_date_str:
                    continue
                try:
                    event_date = datetime.strptime(event_date_str, '%Y-%m-%d')
                    if event_date < clvd_date:
                        continue
                except Exception:
                    continue
                # 授精回数1回目のイベントのみ対象
                jd = event.get('json_data') or {}
                if isinstance(jd, str):
                    try:
                        jd = json.loads(jd)
                    except Exception:
                        jd = {}
                c = jd.get('insemination_count')
                if c is None:
                    continue
                try:
                    if int(c) != 1:
                        continue
                except (ValueError, TypeError):
                    continue
                ai_et_events.append((event_date, event_date_str))
            
            if not ai_et_events:
                return None
            
            ai_et_events.sort(key=lambda x: x[0])
            first_ai_date = ai_et_events[0][0]
            days = (first_ai_date - clvd_date).days
            if days < 0:
                return None
            return days
            
        except Exception as e:
            logging.error(f"初回授精日数計算エラー (cow_auto_id={cow_auto_id}): {e}")
            return None
    
    def _process_scatter_plot_query(self, query: str, analysis_context: Optional[Dict[str, Any]] = None, db: DBHandler = None) -> Optional[Dict[str, Any]]:
        """
        散布図クエリを処理
        
        Args:
            query: ユーザー入力
            analysis_context: 分析コンテキスト（matched_itemsを含む）
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            散布図データの辞書、処理できない場合はNone
        """
        try:
            # analysis_contextから項目を取得
            if analysis_context and analysis_context.get('matched_items'):
                matched_items = analysis_context['matched_items']
            else:
                matched_items = self._match_items_from_query(query)
            
            # 項目が2つ以上必要
            if len(matched_items) < 2:
                logging.warning(f"散布図クエリで項目が2つ以上必要: {query}")
                return None
            
            # 最初の2つの項目を使用
            item1 = matched_items[0]
            item2 = matched_items[1]
            
            # DIM × 初回授精日数の場合
            if (item1['item_key'] == 'DIM' and item2['item_key'] == 'DIMFAI') or \
               (item2['item_key'] == 'DIM' and item1['item_key'] == 'DIMFAI'):
                # DIMを横軸、初回授精日数を縦軸とする
                if item1['item_key'] == 'DIM':
                    return self._get_scatter_data_dim_vs_first_ai_days(
                        x_axis_name=item1['display_name'],
                        y_axis_name=item2['display_name'],
                        db=db
                    )
                else:
                    return self._get_scatter_data_dim_vs_first_ai_days(
                        x_axis_name=item2['display_name'],
                        y_axis_name=item1['display_name'],
                        db=db
                    )
            
            # その他の組み合わせは未対応（将来的に拡張可能）
            logging.warning(f"未対応の散布図軸組み合わせ: {item1['item_key']} × {item2['item_key']}")
            return None
            
        except Exception as e:
            logging.error(f"散布図クエリ処理エラー: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _get_scatter_data_dim_vs_first_ai_days(self, x_axis_name: str, y_axis_name: str, db: DBHandler) -> Optional[Dict[str, Any]]:
        """
        DIM × 初回授精日数の散布図用データを取得
        
        Args:
            x_axis_name: 横軸名
            y_axis_name: 縦軸名
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            散布図データの辞書（'x_data', 'y_data', 'cow_ids', 'x_label', 'y_label', 'title', 'csv_data'を含む）
            データが存在しない場合はNone
        """
        try:
            cows = db.get_all_cows()
            today = datetime.now()
            
            # 集計対象の牛を取得
            target_cows = []
            for cow in cows:
                # 搾乳中（RC=2: Fresh）の牛のみ対象
                if cow.get('rc') != 2:  # RC_FRESH
                    continue
                
                clvd = cow.get('clvd')
                cow_id = cow.get('cow_id', '')
                cow_auto_id = cow.get('auto_id')
                if not clvd or not cow_id or not cow_auto_id:
                    continue
                
                try:
                    clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
                    dim = (today - clvd_date).days
                    if dim < 0:
                        continue
                    
                    # 初回授精日数を計算
                    first_ai_days = self._calculate_first_ai_days(cow_auto_id, db)
                    if first_ai_days is None:
                        # 初回授精日数が計算できない個体は除外
                        continue
                    
                    target_cows.append({
                        'cow_auto_id': cow_auto_id,
                        'cow_id': cow_id,
                        'dim': dim,
                        'first_ai_days': first_ai_days,
                        'lact': cow.get('lact', 0)
                    })
                except:
                    continue
            
            if not target_cows:
                return None
            
            # 経産牛数をカウント
            parous_count = sum(1 for c in target_cows if c.get('lact', 0) > 0)
            
            # 基準日（集計日時）を取得
            base_date_str = today.strftime('%Y年%m月%d日')
            
            # データをリストに変換
            x_data = [c['dim'] for c in target_cows]
            y_data = [c['first_ai_days'] for c in target_cows]
            cow_ids = [c['cow_auto_id'] for c in target_cows]
            
            # CSVデータを生成
            csv_lines = [f"{x_axis_name},{y_axis_name},個体ID"]
            for c in target_cows:
                csv_lines.append(f"{c['dim']},{c['first_ai_days']},{c['cow_auto_id']}")
            csv_data = "\n".join(csv_lines)
            
            # タイトルを生成
            title = f"初回授精日数 × DIM 散布図\n対象牛：{len(target_cows)} 頭（経産牛 {parous_count} 頭）\n基準日：{base_date_str}"
            
            return {
                'x_data': x_data,
                'y_data': y_data,
                'cow_ids': cow_ids,
                'x_label': f"{x_axis_name}（日）",
                'y_label': f"{y_axis_name}（日）",
                'title': title,
                'csv_data': csv_data
            }
            
        except Exception as e:
            logging.error(f"散布図データ取得エラー: {e}")
            return None
    
    def _display_scatter_plot_from_result(self, result: Dict[str, Any], command: Optional[str] = None):
        """
        新しいQueryRouterシステムの結果から散布図を表示（クリック可能）
        
        Args:
            result: QueryExecutorV2の結果辞書（kind="scatter", rows, meta）
            command: 実行したコマンド（オプション）
        """
        # コマンド情報を保存
        if command:
            self.current_graph_command = command
            # コマンド表示テキストを更新
            if hasattr(self, 'graph_command_text') and self.graph_command_text.winfo_exists():
                self._update_command_text(self.graph_command_text, f"コマンド: {command}")
        if not MATPLOTLIB_AVAILABLE:
            self.add_message(role="system", text="matplotlib がインストールされていません。散布図機能は使用できません。")
            return
        
        try:
            rows = result.get("rows", [])
            meta = result.get("meta", {})
            
            if not rows:
                self.add_message(role="system", text="散布図データがありません。")
                return
            
            # 散布図表示用のFrameを取得（なければ作成）
            if not hasattr(self, 'scatter_plot_frame') or self.scatter_plot_frame is None:
                # chat_frameを取得
                if 'chat' not in self.views:
                    self._show_chat_view()
                chat_frame = self.views['chat']
                self.scatter_plot_frame = ttk.Frame(chat_frame)
            
            # 既存の散布図をクリア
            for widget in self.scatter_plot_frame.winfo_children():
                widget.destroy()
            
            # 既にpackされている場合は一旦pack_forget
            try:
                self.scatter_plot_frame.pack_forget()
            except:
                pass
            
            # cow_auto_idからcow_idを取得するためのマッピングを作成
            cow_id_map = {}  # cow_auto_id -> cow_id
            x_values = []
            y_values = []
            cow_ids = []  # インデックスに対応するcow_idのリスト
            
            for row in rows:
                cow_auto_id = row.get("cow_auto_id")
                if cow_auto_id is None:
                    continue
                
                # cow_idを取得（キャッシュがあれば使用）
                if cow_auto_id not in cow_id_map:
                    cow = self.db.get_cow_by_auto_id(cow_auto_id)
                    if cow:
                        cow_id_map[cow_auto_id] = cow.get("cow_id", "")
                    else:
                        cow_id_map[cow_auto_id] = None
                
                cow_id = cow_id_map[cow_auto_id]
                if cow_id:
                    x_values.append(row.get("x"))
                    y_values.append(row.get("y"))
                    cow_ids.append(cow_id)
            
            if not x_values:
                self.add_message(role="system", text="有効な散布図データがありません。")
                return
            
            # 散布図を描画
            fig = Figure(figsize=(8, 6))
            ax = fig.add_subplot(111)
            
            # pickerを有効化したscatterを作成
            scatter = ax.scatter(
                x_values,
                y_values,
                s=20,
                alpha=0.7,
                picker=True,
                pickradius=5
            )
            
            # ラベルとタイトルを設定
            x_item_key = meta.get("x_item_key", "")
            y_item_key = meta.get("y_item_key", "")
            if self.query_renderer:
                x_label = self.query_renderer._get_item_display_name(x_item_key)
                y_label = self.query_renderer._get_item_display_name(y_item_key)
            else:
                x_label = x_item_key
                y_label = y_item_key
            
            ax.set_xlabel(x_label)
            ax.set_ylabel(y_label)
            ax.set_title(f"{x_label} × {y_label}")
            ax.grid(True)
            
            # Canvasを作成してFrameに配置
            canvas = FigureCanvasTkAgg(fig, master=self.scatter_plot_frame)
            
            # pick_eventハンドラを設定
            def on_pick(event):
                """散布図の点がクリックされた時の処理"""
                try:
                    if event.ind is None or len(event.ind) == 0:
                        return
                    
                    # クリックされた点のインデックス（最初の1点のみ）
                    index = event.ind[0]
                    
                    if index >= len(cow_ids):
                        return
                    
                    cow_id = cow_ids[index]
                    if not cow_id:
                        return
                    
                    # 個体カードを開く
                    self._jump_to_cow_card(cow_id)
                    
                    # 視覚的フィードバック（点を一時的に強調）
                    try:
                        # クリックされた点のサイズを一時的に大きくする
                        sizes = scatter.get_sizes()
                        if isinstance(sizes, (list, tuple)) and len(sizes) > index:
                            original_size = sizes[index] if index < len(sizes) else 20
                            new_sizes = list(sizes)
                            new_sizes[index] = original_size * 2
                            scatter.set_sizes(new_sizes)
                            canvas.draw()
                            
                            # 0.3秒後に元に戻す
                            self.root.after(300, lambda: _reset_point_size(index, original_size))
                    except Exception as e:
                        logging.debug(f"視覚的フィードバックエラー（無視）: {e}")
                
                except Exception as e:
                    logging.error(f"散布図クリック処理エラー: {e}")
            
            def _reset_point_size(index: int, original_size: float):
                """点のサイズを元に戻す"""
                try:
                    sizes = scatter.get_sizes()
                    if isinstance(sizes, (list, tuple)) and len(sizes) > index:
                        new_sizes = list(sizes)
                        new_sizes[index] = original_size
                        scatter.set_sizes(new_sizes)
                        canvas.draw()
                except:
                    pass
            
            # pick_eventをバインド
            canvas.mpl_connect("pick_event", on_pick)
            
            # rowsデータを保持（後でCSVコピー用）
            self._current_scatter_rows = rows
            self._current_scatter_meta = meta
            
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # CSVコピーと画像保存のボタンを追加
            button_frame = ttk.Frame(self.scatter_plot_frame)
            button_frame.pack(fill=tk.X, padx=5, pady=5)
            
            def copy_csv_to_clipboard():
                """CSVデータをクリップボードにコピー"""
                try:
                    if self.query_renderer:
                        # QueryRendererでCSV形式に変換
                        csv_text = self.query_renderer.render(result, query_type="scatter")
                        self.root.clipboard_clear()
                        self.root.clipboard_append(csv_text)
                        self.root.update()
                        messagebox.showinfo("情報", "CSVデータをクリップボードにコピーしました")
                    else:
                        messagebox.showerror("エラー", "QueryRendererが利用できません。")
                except Exception as e:
                    logging.error(f"CSVコピーエラー: {e}")
                    messagebox.showerror("エラー", f"CSVデータのコピーに失敗しました: {e}")
            
            def save_plot_image():
                """散布図を画像ファイルとして保存"""
                try:
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".png",
                        filetypes=[("PNGファイル", "*.png"), ("すべてのファイル", "*.*")]
                    )
                    if filename:
                        fig.savefig(filename, dpi=150, bbox_inches='tight')
                        messagebox.showinfo("情報", f"画像を保存しました: {filename}")
                except Exception as e:
                    logging.error(f"画像保存エラー: {e}")
                    messagebox.showerror("エラー", f"画像の保存に失敗しました: {e}")
            
            csv_button = ttk.Button(button_frame, text="CSVコピー", command=copy_csv_to_clipboard)
            csv_button.pack(side=tk.LEFT, padx=5)

            copy_img_button = ttk.Button(
                button_frame,
                text="画像コピー",
                command=lambda: self._copy_figure_image_to_clipboard(fig)
            )
            copy_img_button.pack(side=tk.LEFT, padx=5)
            
            save_button = ttk.Button(button_frame, text="画像として保存", command=save_plot_image)
            save_button.pack(side=tk.LEFT, padx=5)

            # 十字基準線（任意表示）
            ref_frame = ttk.Frame(self.scatter_plot_frame)
            ref_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

            show_ref_var = tk.BooleanVar(value=False)
            x_ref_var = tk.StringVar()
            y_ref_var = tk.StringVar()
            ref_lines = {"v": None, "h": None}

            def _clear_ref_lines():
                for key in ("v", "h"):
                    ln = ref_lines.get(key)
                    if ln is not None:
                        try:
                            ln.remove()
                        except Exception:
                            pass
                        ref_lines[key] = None

            def _apply_ref_lines():
                _clear_ref_lines()
                if not show_ref_var.get():
                    canvas.draw_idle()
                    return

                x_text = x_ref_var.get().strip()
                y_text = y_ref_var.get().strip()
                try:
                    if x_text:
                        x_val = float(x_text)
                        ref_lines["v"] = ax.axvline(
                            x=x_val, color="#c62828", linestyle="--", linewidth=1.2, alpha=0.95
                        )
                    if y_text:
                        y_val = float(y_text)
                        ref_lines["h"] = ax.axhline(
                            y=y_val, color="#1565c0", linestyle="--", linewidth=1.2, alpha=0.95
                        )
                except ValueError:
                    messagebox.showwarning("入力エラー", "基準線の値は数値で入力してください。")
                    return

                canvas.draw_idle()

            ttk.Checkbutton(
                ref_frame,
                text="十字基準線を表示",
                variable=show_ref_var,
                command=_apply_ref_lines
            ).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Label(ref_frame, text="X:").pack(side=tk.LEFT)
            ttk.Entry(ref_frame, textvariable=x_ref_var, width=8).pack(side=tk.LEFT, padx=(2, 8))
            ttk.Label(ref_frame, text="Y:").pack(side=tk.LEFT)
            ttk.Entry(ref_frame, textvariable=y_ref_var, width=8).pack(side=tk.LEFT, padx=(2, 8))
            ttk.Button(ref_frame, text="適用", command=_apply_ref_lines).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(
                ref_frame,
                text="解除",
                command=lambda: (
                    show_ref_var.set(False),
                    x_ref_var.set(""),
                    y_ref_var.set(""),
                    _apply_ref_lines()
                )
            ).pack(side=tk.LEFT)
            
            # scatter_plot_frameをpack（チャット表示エリアの下に表示）
            self.scatter_plot_frame.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
            
        except Exception as e:
            logging.error(f"散布図表示エラー: {e}")
            import traceback
            traceback.print_exc()
            self.add_message(role="system", text=f"散布図の表示に失敗しました: {e}")
    
    def _display_scatter_plot(self, scatter_data: Dict[str, Any]):
        """
        散布図を表示する（後方互換性のため残す）
        
        Args:
            scatter_data: 散布図データの辞書（旧形式）
        """
        if not MATPLOTLIB_AVAILABLE:
            self.add_message(role="system", text="matplotlib がインストールされていません。散布図機能は使用できません。")
            return
        
        try:
            # 散布図表示用のFrameを取得（なければ作成）
            if not hasattr(self, 'scatter_plot_frame') or self.scatter_plot_frame is None:
                # chat_frameを取得
                if 'chat' not in self.views:
                    self._show_chat_view()
                chat_frame = self.views['chat']
                self.scatter_plot_frame = ttk.Frame(chat_frame)
            
            # 既存の散布図をクリア
            for widget in self.scatter_plot_frame.winfo_children():
                widget.destroy()
            
            # 既にpackされている場合は一旦pack_forget
            try:
                self.scatter_plot_frame.pack_forget()
            except:
                pass
            
            # 散布図を描画
            fig = Figure(figsize=(8, 6))
            ax = fig.add_subplot(111)
            
            ax.scatter(scatter_data['x_data'], scatter_data['y_data'], alpha=0.6)
            ax.set_xlabel(scatter_data['x_label'])
            ax.set_ylabel(scatter_data['y_label'])
            ax.set_title(scatter_data['title'])
            ax.grid(True)
            
            # Canvasを作成してFrameに配置
            canvas = FigureCanvasTkAgg(fig, master=self.scatter_plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # CSVコピーと画像保存のボタンを追加
            button_frame = ttk.Frame(self.scatter_plot_frame)
            button_frame.pack(fill=tk.X, padx=5, pady=5)
            
            def copy_csv_to_clipboard():
                """CSVデータをクリップボードにコピー"""
                try:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(scatter_data['csv_data'])
                    self.root.update()  # クリップボードを更新
                    messagebox.showinfo("情報", "CSVデータをクリップボードにコピーしました")
                except Exception as e:
                    logging.error(f"CSVコピーエラー: {e}")
                    messagebox.showerror("エラー", f"CSVデータのコピーに失敗しました: {e}")
            
            def save_plot_image():
                """散布図を画像ファイルとして保存"""
                try:
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".png",
                        filetypes=[("PNGファイル", "*.png"), ("すべてのファイル", "*.*")]
                    )
                    if filename:
                        fig.savefig(filename, dpi=150, bbox_inches='tight')
                        messagebox.showinfo("情報", f"画像を保存しました: {filename}")
                except Exception as e:
                    logging.error(f"画像保存エラー: {e}")
                    messagebox.showerror("エラー", f"画像の保存に失敗しました: {e}")
            
            csv_button = ttk.Button(button_frame, text="CSVコピー", command=copy_csv_to_clipboard)
            csv_button.pack(side=tk.LEFT, padx=5)

            copy_img_button = ttk.Button(
                button_frame,
                text="画像コピー",
                command=lambda: self._copy_figure_image_to_clipboard(fig)
            )
            copy_img_button.pack(side=tk.LEFT, padx=5)
            
            save_button = ttk.Button(button_frame, text="画像として保存", command=save_plot_image)
            save_button.pack(side=tk.LEFT, padx=5)

            # 十字基準線（任意表示）
            ref_frame = ttk.Frame(self.scatter_plot_frame)
            ref_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

            show_ref_var = tk.BooleanVar(value=False)
            x_ref_var = tk.StringVar()
            y_ref_var = tk.StringVar()
            ref_lines = {"v": None, "h": None}

            def _clear_ref_lines():
                for key in ("v", "h"):
                    ln = ref_lines.get(key)
                    if ln is not None:
                        try:
                            ln.remove()
                        except Exception:
                            pass
                        ref_lines[key] = None

            def _apply_ref_lines():
                _clear_ref_lines()
                if not show_ref_var.get():
                    canvas.draw_idle()
                    return

                x_text = x_ref_var.get().strip()
                y_text = y_ref_var.get().strip()
                try:
                    if x_text:
                        x_val = float(x_text)
                        ref_lines["v"] = ax.axvline(
                            x=x_val, color="#c62828", linestyle="--", linewidth=1.2, alpha=0.95
                        )
                    if y_text:
                        y_val = float(y_text)
                        ref_lines["h"] = ax.axhline(
                            y=y_val, color="#1565c0", linestyle="--", linewidth=1.2, alpha=0.95
                        )
                except ValueError:
                    messagebox.showwarning("入力エラー", "基準線の値は数値で入力してください。")
                    return

                canvas.draw_idle()

            ttk.Checkbutton(
                ref_frame,
                text="十字基準線を表示",
                variable=show_ref_var,
                command=_apply_ref_lines
            ).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Label(ref_frame, text="X:").pack(side=tk.LEFT)
            ttk.Entry(ref_frame, textvariable=x_ref_var, width=8).pack(side=tk.LEFT, padx=(2, 8))
            ttk.Label(ref_frame, text="Y:").pack(side=tk.LEFT)
            ttk.Entry(ref_frame, textvariable=y_ref_var, width=8).pack(side=tk.LEFT, padx=(2, 8))
            ttk.Button(ref_frame, text="適用", command=_apply_ref_lines).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(
                ref_frame,
                text="解除",
                command=lambda: (
                    show_ref_var.set(False),
                    x_ref_var.set(""),
                    y_ref_var.set(""),
                    _apply_ref_lines()
                )
            ).pack(side=tk.LEFT)
            
            # scatter_plot_frameをpack（チャット表示エリアの下に表示）
            self.scatter_plot_frame.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
            
        except Exception as e:
            logging.error(f"散布図表示エラー: {e}")
            import traceback
            traceback.print_exc()
            self.add_message(role="system", text=f"散布図の表示に失敗しました: {e}")
    
    def _display_table_report(self, table_data: Dict[str, Any]):
        """
        表レポートを表示する
        
        Args:
            table_data: 表レポートデータの辞書
        """
        try:
            from ui.report_table_window import ReportTableWindow
            
            ReportTableWindow(
                parent=self.root,
                report_title=table_data.get('title', 'レポート'),
                columns=table_data.get('columns', []),
                rows=table_data.get('rows', []),
                conditions=table_data.get('conditions'),
                period=table_data.get('period')
            ).show()
            
        except Exception as e:
            logging.error(f"表レポート表示エラー: {e}")
            import traceback
            traceback.print_exc()
            self.add_message(role="system", text=f"表レポートの表示に失敗しました: {e}")
    
    def _get_dim_scatter_data(self) -> str:
        """DIM × 個体ID の散布図用データを取得（後方互換性のため残す）"""
        try:
            cows = self.db.get_all_cows()
            
            data_lines = ["DIM,個体ID"]
            
            for cow in cows:
                # 搾乳中（RC=2: Fresh）の牛のみ対象
                if cow.get('rc') != 2:  # RC_FRESH
                    continue
                
                clvd = cow.get('clvd')
                cow_id = cow.get('cow_id', '')
                if not clvd or not cow_id:
                    continue
                
                try:
                    clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
                    today = datetime.now()
                    dim = (today - clvd_date).days
                    if dim >= 0:
                        data_lines.append(f"{dim},{cow_id}")
                except:
                    continue
            
            if len(data_lines) <= 1:
                return None
            
            # CSV形式のみを返す（説明文なし）
            return chr(10).join(data_lines)
            
        except Exception as e:
            logging.error(f"散布図データ取得エラー: {e}")
            return None
    
