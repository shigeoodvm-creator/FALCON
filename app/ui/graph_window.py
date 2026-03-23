"""
FALCON2 - グラフ表示ウィンドウ
グラフを表示する専用ウインドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime
import io
import logging
import sys
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import font_manager
import matplotlib.dates as mdates
import numpy as np
import warnings

# Log-rank検定用（オプション：lifelinesが無い場合はP値非表示）
try:
    from lifelines.statistics import multivariate_logrank_test
    _HAS_LIFELINES = True
except ImportError:
    _HAS_LIFELINES = False

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine

# 日本語フォントの設定（利用可能なフォントを動的に検出）
def _setup_japanese_font():
    """利用可能な日本語フォントを検出して設定"""
    # 優先順位の高いフォントリスト
    preferred_fonts = ['MS Gothic', 'MS PGothic', 'Meiryo', 'Yu Gothic', 'MS UI Gothic', 'Noto Sans CJK JP']
    
    # 利用可能なフォントを取得
    available_fonts = [f.name for f in font_manager.fontManager.ttflist]
    
    # 優先順位に従って利用可能なフォントを検索
    selected_font = None
    for font_name in preferred_fonts:
        if font_name in available_fonts:
            selected_font = font_name
            break
    
    # 利用可能なフォントが見つからない場合、sans-serifを使用
    if selected_font is None:
        selected_font = 'sans-serif'
        logging.warning("日本語フォントが見つかりません。sans-serifを使用します。")
    else:
        logging.info(f"日本語フォントとして '{selected_font}' を使用します。")
    
    # フォント警告を抑制（matplotlibのフォント警告を無視）
    warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib.font_manager')
    warnings.filterwarnings('ignore', message='.*findfont.*', category=UserWarning)
    
    # フォントを設定
    plt.rcParams['font.family'] = selected_font
    plt.rcParams['axes.unicode_minus'] = False  # マイナス記号の文字化けを防ぐ
    
    # matplotlibのログレベルを調整（フォント警告を抑制）
    matplotlib_logger = logging.getLogger('matplotlib.font_manager')
    matplotlib_logger.setLevel(logging.ERROR)

# フォント設定を実行
_setup_japanese_font()


class GraphWindow:
    """グラフ表示ウィンドウ（Toplevel）"""
    
    def __init__(self, parent: tk.Tk, graph_type: str,
                 x_item: str, y_item: str = "",
                 classification: str = "",
                 condition: str = "",
                 start_date: str = "",
                 end_date: str = "",
                 db=None,
                 formula_engine=None,
                 rule_engine=None,
                 farm_path=None,
                 event_dictionary_path=None,
                 on_cow_card_requested=None,
                 classification_event_number: Optional[int] = None,
                 classification_item_key: Optional[str] = None,
                 classification_bin_days: Optional[int] = None,
                 classification_threshold: Optional[float] = None):
        """
        初期化
        
        Args:
            classification_threshold: 生存曲線で項目を閾値で二分する値（例: 0.1 → 「0.1未満」「0.1以上」）
        """
        self.parent = parent
        self.graph_type = graph_type
        self.x_item = x_item
        self.y_item = y_item
        self.classification = classification
        self.condition = condition
        self.classification_event_number = classification_event_number
        self.classification_item_key = classification_item_key or None
        self.classification_bin_days = classification_bin_days
        self.classification_threshold = classification_threshold  # 閾値（未満/以上）
        self.start_date = start_date
        self.end_date = end_date
        self.db = db
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.farm_path = farm_path
        self.event_dictionary_path = event_dictionary_path
        self.on_cow_card_requested = on_cow_card_requested  # 個体カードを開くコールバック関数
        
        # ダブルクリック検出用の変数
        self.last_click_time = 0
        self.last_click_x = None
        self.last_click_y = None
        self.double_click_threshold = 0.3  # ダブルクリックとみなす時間間隔（秒）
        self.double_click_position_threshold = 5  # ダブルクリックとみなす位置の差（ピクセル）
        
        # ウィンドウ作成
        try:
            self.window = tk.Toplevel(parent)
            self.window.title(f"グラフ - {graph_type}")
            self.window.geometry("1000x700")
            
            self._create_widgets()
            self._load_data()
        except Exception as e:
            logging.error(f"グラフウィンドウ作成エラー: {e}", exc_info=True)
            import traceback
            error_detail = traceback.format_exc()
            logging.error(f"グラフウィンドウ作成エラー詳細: {error_detail}")
            from tkinter import messagebox
            messagebox.showerror("エラー", f"グラフウィンドウの作成に失敗しました: {e}\n\n詳細はログを確認してください。")
    
    def __del__(self):
        """GraphWindowインスタンスの削除"""
        pass
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 分類項目チェックボックスフレーム（散布図・生存曲線の場合のみ表示）
        self.classification_frame = None
        self.classification_vars: Dict[str, tk.BooleanVar] = {}
        if self.graph_type == "散布図" and self.classification:
            self.classification_frame = ttk.LabelFrame(main_frame, text="分類項目", padding=5)
            self.classification_frame.pack(fill=tk.X, padx=5, pady=5)
        if self.graph_type == "空胎日数生存曲線" and self.classification:
            if self.classification_frame is None:
                self.classification_frame = ttk.LabelFrame(main_frame, text="分類項目", padding=5)
            self.classification_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 近似曲線チェックボックス（散布図の場合のみ表示）
        self.trend_line_frame = None
        self.show_trend_line_var: Optional[tk.BooleanVar] = None
        if self.graph_type == "散布図":
            self.trend_line_frame = ttk.Frame(main_frame)
            self.trend_line_frame.pack(fill=tk.X, padx=5, pady=2)
            self.show_trend_line_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(
                self.trend_line_frame,
                text="近似曲線を表示（式・R²も表示）",
                variable=self.show_trend_line_var,
                command=self._on_trend_line_checkbox_changed
            ).pack(side=tk.LEFT, padx=5)
        
        # グラフ表示フレーム
        graph_frame = ttk.Frame(main_frame)
        graph_frame.pack(fill=tk.BOTH, expand=True)
        
        # matplotlibのFigureを作成
        self.fig = Figure(figsize=(10, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        
        # Canvasを作成
        self.canvas = FigureCanvasTkAgg(self.fig, graph_frame)
        self._init_canvas_id = id(self.canvas)  # 初期canvas IDを保存（差し替え検出用）
        self.canvas.draw()
        canvas_widget = self.canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)
        # 右クリックでコンテキストメニュー（グラフをコピー）
        canvas_widget.bind("<Button-3>", self._on_graph_right_click)
        
        # 散布図の場合のみ、クリックイベントとマウス移動イベントを接続（重複接続を防ぐ）
        
        if self.graph_type == "散布図":
            # 既存のイベント接続を切断（重複接続を防ぐため）
            if hasattr(self.canvas, "_falcon_scatter_cid_click") and self.canvas._falcon_scatter_cid_click is not None:
                try:
                    self.canvas.mpl_disconnect(self.canvas._falcon_scatter_cid_click)
                    logging.info(f"既存のクリックイベント接続を切断: cid={self.canvas._falcon_scatter_cid_click}")
                except Exception as e:
                    logging.debug(f"既存のクリックイベント切断エラー（無視）: {e}")
            
            if hasattr(self.canvas, "_falcon_scatter_cid_motion") and self.canvas._falcon_scatter_cid_motion is not None:
                try:
                    self.canvas.mpl_disconnect(self.canvas._falcon_scatter_cid_motion)
                    logging.info(f"既存のマウス移動イベント接続を切断: cid={self.canvas._falcon_scatter_cid_motion}")
                except Exception as e:
                    logging.debug(f"既存のマウス移動イベント切断エラー（無視）: {e}")
            
            # イベントハンドラーを接続
            cid_click = self.canvas.mpl_connect('button_press_event', self._on_scatter_click)
            cid_motion = self.canvas.mpl_connect('motion_notify_event', self._on_scatter_hover)
            self.canvas._falcon_scatter_cid_click = cid_click
            self.canvas._falcon_scatter_cid_motion = cid_motion
            self.canvas._falcon_scatter_connected = True
        
        # ツールチップ用の変数を初期化（散布図の場合のみ）
        if self.graph_type == "散布図":
            self.tooltip = None
            self.tooltip_annotation = None
    
    def _on_graph_right_click(self, event):
        """グラフ上で右クリックしたときにコンテキストメニューを表示"""
        menu = tk.Menu(self.window, tearoff=0)
        menu.add_command(label="グラフをコピー", command=self._copy_graph_to_clipboard)
        menu.add_command(label="画像として保存...", command=self._save_graph_as_image)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def _save_graph_as_image(self):
        """グラフをPNG画像ファイルとして保存"""
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG画像", "*.png"), ("すべてのファイル", "*.*")],
            title="グラフを画像として保存"
        )
        if not path:
            return
        try:
            self.fig.savefig(path, dpi=150, bbox_inches="tight",
                             facecolor=self.fig.get_facecolor(), edgecolor="none")
            messagebox.showinfo("保存完了", f"グラフを保存しました。\n{path}")
        except Exception as e:
            logging.error(f"グラフの保存に失敗: {e}", exc_info=True)
            messagebox.showerror("保存失敗", f"グラフの保存に失敗しました。\n{e}")
    
    def _copy_graph_to_clipboard(self):
        """現在のグラフを画像としてクリップボードにコピー（レポート・プレゼン用）"""
        try:
            buf = io.BytesIO()
            self.fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                             facecolor=self.fig.get_facecolor(), edgecolor="none")
            buf.seek(0)
            data_png = buf.getvalue()
        except Exception as e:
            logging.error(f"グラフ画像の生成に失敗: {e}", exc_info=True)
            messagebox.showerror("コピー失敗", f"グラフ画像の生成に失敗しました。\n{e}")
            return
        
        if sys.platform == "win32":
            try:
                import win32clipboard  # type: ignore[reportMissingModuleSource]
                from PIL import Image
            except ImportError as e:
                logging.warning(f"クリップボード用モジュールがありません: {e}")
                messagebox.showinfo(
                    "グラフをコピー",
                    "画像をクリップボードにコピーするには、次のパッケージが必要です。\n\n"
                    "pip install pywin32 Pillow\n\n"
                    "インストール後、アプリを再起動してから再度お試しください。"
                )
                return
            try:
                img = Image.open(io.BytesIO(data_png))
                output = io.BytesIO()
                img.convert("RGB").save(output, "BMP")
                bmp_data = output.getvalue()
                dib_data = bmp_data[14:]
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib_data)
                win32clipboard.CloseClipboard()
                messagebox.showinfo("コピー完了", "グラフをクリップボードにコピーしました。\nWord・Excel・PowerPointなどにそのまま貼り付けできます。")
            except Exception as e:
                logging.error(f"クリップボードへのコピーに失敗: {e}", exc_info=True)
                messagebox.showerror("コピー失敗", f"クリップボードへのコピーに失敗しました。\n{e}")
        else:
            self.window.clipboard_clear()
            self.window.clipboard_append("")
            messagebox.showinfo(
                "グラフをコピー",
                "この環境では画像のクリップボードコピーに未対応です。\n"
                "右クリックメニューの「画像として保存...」でPNGを保存してご利用ください。"
            )
    
    def _is_cow_disposed(self, cow_auto_id: int) -> bool:
        """
        個体が売却または死亡廃用されているかをチェック
        
        Args:
            cow_auto_id: 牛の auto_id
            
        Returns:
            売却または死亡廃用されている場合True
        """
        if not self.rule_engine:
            return False
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for event in events:
            event_number = event.get('event_number')
            if event_number in [RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD]:
                return True
        return False
    
    def _filter_existing_cows(self, cows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        現存牛のみをフィルタリング（除籍牛を除外）
        
        Args:
            cows: 牛のデータリスト
            
        Returns:
            現存牛のみのリスト
        """
        existing_cows = []
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if cow_auto_id and not self._is_cow_disposed(cow_auto_id):
                existing_cows.append(cow)
        return existing_cows
    
    def _load_data(self):
        """データを読み込んでグラフを描画"""
        try:
            # 全牛を取得
            if not self.db:
                logging.error("DBHandlerが設定されていません")
                messagebox.showerror("エラー", "データベース接続が設定されていません。")
                return
            all_cows = self.db.get_all_cows()
            logging.info(f"全牛取得: {len(all_cows)}頭")
            
            # 現存牛のみをフィルタリング（除籍牛を除外）- 生存曲線の場合は除籍牛も含め打ち切りに使う
            if self.graph_type != "空胎日数生存曲線":
                all_cows = self._filter_existing_cows(all_cows)
            logging.info(f"フィルタリング後: {len(all_cows)}頭")
            
            # 期間フィルタリング（現時点では無効化 - すべての牛を対象）
            # TODO: 将来的には、グラフの対象項目に関連するデータがある期間でフィルタリングする
            if self.start_date or self.end_date:
                logging.info(f"期間フィルタリングは現在無効化されています: 開始日={self.start_date}, 終了日={self.end_date}")
                # 期間フィルタリングは無効化（すべての牛を対象）
            
            # アクティブなscatter artistをクリア
            if hasattr(self.canvas, "_falcon_active_scatters"):
                self.canvas._falcon_active_scatters = []
            
            # グラフ種類に応じて処理
            if self.graph_type == "棒グラフ":
                self._draw_bar_chart(all_cows)
            elif self.graph_type == "散布図":
                self._draw_scatter_chart(all_cows)
            elif self.graph_type == "箱ひげ":
                self._draw_box_plot(all_cows)
            elif self.graph_type == "円グラフ":
                self._draw_pie_chart(all_cows)
            elif self.graph_type == "空胎日数生存曲線":
                survival_data = self._compute_days_open_survival_data(all_cows)
                self._draw_survival_curve(survival_data)
            
            # グラフを描画
            self.canvas.draw()
            
            # 散布図の場合、アクティブなaxesをcanvasに保存
            if self.graph_type == "散布図":
                self.canvas._falcon_active_axes = self.ax
            
        except Exception as e:
            logging.error(f"グラフ描画エラー: {e}", exc_info=True)
            import traceback
            error_detail = traceback.format_exc()
            logging.error(f"グラフ描画エラー詳細: {error_detail}")
            messagebox.showerror("エラー", f"グラフの描画に失敗しました: {e}\n\n詳細はログを確認してください。")
    
    def _draw_bar_chart(self, cows: List[Dict[str, Any]]):
        """棒グラフを描画（DIMの平均を産次ごとに表示）"""
        # データを集計
        classification_data: Dict[str, List[float]] = {}
        
        logging.info(f"棒グラフ描画開始: 牛の数={len(cows)}, 対象項目={self.x_item}, 分類={self.classification}")
        
        processed_count = 0
        condition_filtered = 0
        classification_missing = 0
        x_value_missing = 0
        
        for cow in cows:
            # auto_idを取得（キー名を確認）
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID") or cow.get("AUTO_ID")
            if not cow_auto_id:
                logging.debug(f"auto_idが見つかりません: cow={cow}")
                continue
            
            processed_count += 1
            
            # 条件チェック
            if not self._check_condition(cow, cow_auto_id):
                condition_filtered += 1
                continue
            
            # 分類値を取得
            classification_value = self._get_classification_value(cow, cow_auto_id)
            if classification_value is None:
                classification_missing += 1
                continue
            
            # X軸項目の値を取得
            x_value = self._get_item_value(cow, cow_auto_id, self.x_item)
            if x_value is None:
                x_value_missing += 1
                continue
            
            # 分類値でグループ化
            if classification_value not in classification_data:
                classification_data[classification_value] = []
            classification_data[classification_value].append(x_value)
        
        result_msg = f"データ集計結果: 処理数={processed_count}, 条件で除外={condition_filtered}, 分類値なし={classification_missing}, 項目値なし={x_value_missing}, 有効データ={sum(len(v) for v in classification_data.values())}"
        print(result_msg)
        logging.info(result_msg)
        
        # 平均値を計算
        classification_means: Dict[str, float] = {}
        for key, values in classification_data.items():
            if values:
                classification_means[key] = np.mean(values)
        
        if not classification_means:
            self.ax.text(0.5, 0.5, "データがありません", 
                        ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 分類値をソート（産次の場合は数値順）
        categories = list(classification_means.keys())
        if self.classification and (self.classification.upper() == "LACT" or self.classification == "産次"):
            # 産次は数値順
            categories = sorted(categories, key=lambda x: int(x) if isinstance(x, int) or str(x).isdigit() else 999)
        else:
            # その他の場合は文字列順
            categories = sorted(categories)
        
        means = [classification_means[cat] for cat in categories]
        counts = [len(classification_data[cat]) for cat in categories]
        
        # 項目の表示名を取得
        x_item_display = self._get_item_display_name(self.x_item)
        classification_display = self._get_item_display_name(self.classification) if self.classification else "分類"
        
        # 棒グラフを作成（Rのようなスタイル）
        bars = self.ax.bar(range(len(categories)), means, color='steelblue', alpha=0.7, edgecolor='black', linewidth=0.5)
        
        # X軸のラベルに頭数を追加（例：1 (N=19)）
        x_labels_with_n = []
        for i, cat in enumerate(categories):
            n = counts[i] if i < len(counts) else 0
            x_labels_with_n.append(f"{cat}\n(N={n})")
        
        # X軸のティック位置とラベルを設定
        self.ax.set_xticks(range(len(categories)))
        self.ax.set_xticklabels(x_labels_with_n, rotation=0, ha='center')
        
        # ラベルとタイトルを設定
        self.ax.set_xlabel(classification_display, fontsize=11, fontweight='bold')
        self.ax.set_ylabel(f"{x_item_display}平均", fontsize=11, fontweight='bold')
        self.ax.set_title(f"{x_item_display}平均 × {classification_display}", fontsize=12, fontweight='bold', pad=15)
        
        # グリッドを追加（Rのようなスタイル）
        self.ax.grid(True, linestyle='--', alpha=0.3, axis='y')
        self.ax.set_axisbelow(True)
        
        # Y軸を常に0起点にする（少し余白を追加）
        y_max = max(means) if means else 1
        y_range = y_max
        self.ax.set_ylim(0, y_max + y_range * 0.1)
        
        # 各棒の上に値を表示
        for i, (bar, mean) in enumerate(zip(bars, means)):
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width() / 2., height,
                       f'{mean:.1f}',
                       ha='center', va='bottom', fontsize=9)
    
    def _draw_scatter_chart(self, cows: List[Dict[str, Any]], create_checkboxes: bool = True):
        """散布図を描画（分類ごとに色分け）"""
        # 分類ごとにデータをグループ化
        classification_data: Dict[str, List[Tuple[float, float]]] = {}
        # データポイントと個体の対応関係を保持（(x, y) -> cow_auto_id, cow_id）
        # 分類ごとにデータポイントと個体の対応関係を保持
        self.scatter_data_map: Dict[str, List[Tuple[float, float, int, str]]] = {}  # {classification: [(x, y, cow_auto_id, cow_id), ...]}
        
        processed_count = 0
        condition_filtered = 0
        classification_missing = 0
        x_value_missing = 0
        y_value_missing = 0
        valid_data = 0
        
        logging.info(f"散布図描画開始: 牛の数={len(cows)}, X軸={self.x_item}, Y軸={self.y_item}, 分類={self.classification}")
        
        for cow in cows:
            # auto_idを取得（キー名を確認）
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID") or cow.get("AUTO_ID")
            if not cow_auto_id:
                continue
            
            processed_count += 1
            
            # 条件チェック
            if not self._check_condition(cow, cow_auto_id):
                condition_filtered += 1
                continue
            
            # 分類値を取得
            classification_value = None
            if self.classification:
                classification_value = self._get_classification_value(cow, cow_auto_id)
                if classification_value is None:
                    classification_missing += 1
                    logging.debug(f"分類値が取得できません: cow_auto_id={cow_auto_id}, classification={self.classification}")
                    continue
            else:
                classification_value = "全体"
            
            # X軸とY軸の値を取得
            x_value = self._get_item_value(cow, cow_auto_id, self.x_item)
            y_value = self._get_item_value(cow, cow_auto_id, self.y_item)
            
            if x_value is None:
                x_value_missing += 1
                continue
            if y_value is None:
                y_value_missing += 1
                continue
            
            # データを追加
            if classification_value not in classification_data:
                classification_data[classification_value] = []
            classification_data[classification_value].append((x_value, y_value))
            # データポイントと個体の対応関係を保存（分類ごと）
            # cow_idを取得（拡大4桁ID）
            cow_id = cow.get("cow_id", "") or str(cow_auto_id)
            if classification_value not in self.scatter_data_map:
                self.scatter_data_map[classification_value] = []
            self.scatter_data_map[classification_value].append((x_value, y_value, cow_auto_id, cow_id))
            valid_data += 1
        
        result_msg = f"散布図データ集計結果: 処理数={processed_count}, 条件で除外={condition_filtered}, 分類値なし={classification_missing}, X値なし={x_value_missing}, Y値なし={y_value_missing}, 有効データ={valid_data}, 分類数={len(classification_data)}"
        print(result_msg)
        logging.info(result_msg)
        
        
        if not classification_data:
            self.ax.text(0.5, 0.5, "データがありません", 
                        ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 分類項目のチェックボックスを作成（散布図の場合のみ、初回のみ）
        if create_checkboxes and self.classification and self.classification_frame:
            # 分類データのキーを文字列に統一してからチェックボックスを作成
            classification_keys = {str(k) for k in classification_data.keys()}
            self._create_classification_checkboxes(classification_keys)
        
        # チェックボックスの状態を取得
        enabled_classifications = set()
        if self.classification_vars:
            for cat, var in self.classification_vars.items():
                if var.get():
                    enabled_classifications.add(cat)
        else:
            # チェックボックスがない場合はすべて有効
            enabled_classifications = {str(k) for k in classification_data.keys()}
        
        # 分類データのキーを文字列に統一
        classification_data_str = {str(k): v for k, v in classification_data.items()}
        
        # 分類値をソート（産次の場合は数値順、産次分類の場合は1産、2産、3産以上の順、DIM分類の場合は数値順、分娩月は数値順、空胎日数は未満→以上の順）
        def sort_key(x):
            # 産次分類の場合（1産、2産、3産以上）
            if isinstance(x, str):
                if x == "1産":
                    return (0, 1)
                elif x == "2産":
                    return (0, 2)
                elif x == "3産以上":
                    return (0, 3)
                # 分娩月の場合（1月、2月、...、12月）
                elif x.endswith("月") and x[:-1].isdigit():
                    month = int(x[:-1])
                    if 1 <= month <= 12:
                        return (1, month)
                # 空胎日数の場合（<100日、>=100日）
                elif "日" in x:
                    if x.startswith("<"):
                        # 未満の場合は先に表示
                        try:
                            threshold = int(x[1:].replace("日", ""))
                            return (2, threshold)
                        except (ValueError, TypeError):
                            pass
                    elif x.startswith(">="):
                        # 以上の場合は後に表示
                        try:
                            threshold = int(x[2:].replace("日", ""))
                            return (3, threshold)
                        except (ValueError, TypeError):
                            pass
                # DIM分類の場合（0-6, 7-13, 14-20, ...）
                elif "-" in x and x.count("-") == 1:
                    try:
                        parts = x.split("-")
                        if len(parts) == 2 and parts[0].isdigit():
                            start = int(parts[0])
                            return (4, start)  # 開始値を基準にソート
                    except (ValueError, TypeError):
                        pass
                # 数値の場合は数値順
                elif x.isdigit():
                    return (5, int(x))
            # 数値の場合は数値順
            elif isinstance(x, int):
                return (5, x)
            # その他は文字列順
            return (6, str(x))
        
        sorted_classifications = sorted(classification_data_str.keys(), key=sort_key)
        
        # 色のリストを準備（鮮やかで区別しやすい色パレットを使用）
        # Set1, Set2, Set3, Dark2, Pairedなどのカラーマップを組み合わせて使用
        num_classifications = len(sorted_classifications)
        colors = []
        
        # 複数のカラーマップから色を取得して組み合わせる
        if num_classifications <= 9:
            # Set1は9色の鮮やかな色
            colors = [plt.cm.Set1(i) for i in range(num_classifications)]
        elif num_classifications <= 18:
            # Set1とSet2を組み合わせ
            colors = [plt.cm.Set1(i) for i in range(9)] + [plt.cm.Set2(i) for i in range(num_classifications - 9)]
        elif num_classifications <= 27:
            # Set1, Set2, Set3を組み合わせ
            colors = ([plt.cm.Set1(i) for i in range(9)] + 
                     [plt.cm.Set2(i) for i in range(8)] + 
                     [plt.cm.Set3(i) for i in range(num_classifications - 17)])
        else:
            # それ以上の場合、tab20を使用
            colors = [plt.cm.tab20(i) for i in range(min(num_classifications, 20))]
            if num_classifications > 20:
                # さらに色が必要な場合、Dark2を追加
                colors.extend([plt.cm.Dark2(i % 8) for i in range(num_classifications - 20)])
        
        color_map = {}
        for i, cat in enumerate(sorted_classifications):
            color_map[cat] = colors[i] if i < len(colors) else plt.cm.tab20(i % 20)
        
        # 分類ごとに散布図を描画
        x_item_display = self._get_item_display_name(self.x_item)
        y_item_display = self._get_item_display_name(self.y_item)
        
        # X軸が日付項目かどうかをチェック（ツールチップ用に保存）
        normalized_x_item = self.x_item.upper() if self.x_item else ""
        is_date_axis = False
        is_month_axis = False  # 分娩月などの月項目かどうか
        self.is_date_axis = False  # ツールチップ用に保存
        self.is_month_axis = False  # ツールチップ用に保存
        
        # 分娩月（CALVMO）の特別処理
        if normalized_x_item == "CALVMO" or x_item_display == "分娩月":
            is_month_axis = True
        else:
            # 1. 明示的な日付項目（CLVD、BTHDなど）をチェック
            known_date_items = ["CLVD", "BTHD", "CONC", "DUED", "LAID", "LMTD", "DRYTD", "NEXTED", "PCLVD"]
            if normalized_x_item in known_date_items:
                is_date_axis = True
            else:
                # 2. item_dictionaryからdata_typeを確認
                if self.formula_engine:
                    try:
                        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
                        if normalized_x_item in item_dict:
                            item_def = item_dict[normalized_x_item]
                            data_type = item_def.get("data_type", "")
                            # data_typeが"str"で、表示名に「日」が含まれる場合は日付項目の可能性が高い
                            display_name = item_def.get("display_name", "") or item_def.get("label", "")
                            if data_type == "str" and ("日" in display_name or "date" in normalized_x_item.lower()):
                                is_date_axis = True
                            # 表示名に「月」が含まれ、数値が1-12の範囲の場合は月項目
                            elif "月" in display_name and normalized_x_item not in known_date_items:
                                is_month_axis = True
                    except Exception as e:
                        logging.debug(f"日付項目判定エラー: {e}")
                
                # 3. 実際の値が日付形式（エポック秒）かどうかを確認（フォールバック）
                # エポック秒は通常、1970-01-01以降の日付の場合、大きな正の数値（例：1000000000以上）
                # 注意: 値が1-12だからといって月項目にはしない（授精回数・産次などと混同するため。月はCALVMO/表示名「分娩月」のみ）
                if not is_date_axis and not is_month_axis and classification_data_str:
                    # 最初のデータポイントのX値を確認
                    first_data_points = next(iter(classification_data_str.values()))
                    if first_data_points:
                        first_x_value = first_data_points[0][0]
                        # エポック秒の範囲（1970-01-01から2100-01-01まで）をチェック
                        if isinstance(first_x_value, (int, float)) and first_x_value > 1000000 and first_x_value <= 4102444800:
                            # 日付として解釈できる可能性がある（小さい値は除外）
                            try:
                                test_date = datetime.fromtimestamp(first_x_value)
                                # 妥当な日付範囲内かチェック
                                if datetime(1970, 1, 1) <= test_date <= datetime(2100, 1, 1):
                                    is_date_axis = True
                            except (ValueError, OSError):
                                pass
        
        # 日付項目の判定結果を保存（ツールチップ用）
        self.is_date_axis = is_date_axis
        self.is_month_axis = is_month_axis
        
        # 分類の表示名を取得
        classification_display = self._get_classification_display_name()
        
        # アクティブなscatter artistを保持するリスト
        active_scatters = []
        
        for classification_value, data_points in classification_data_str.items():
            if classification_value not in enabled_classifications:
                continue
            
            x_values = [p[0] for p in data_points]
            y_values = [p[1] for p in data_points]
            
            # この分類のメタデータを取得（scatter_data_mapから）
            if classification_value not in self.scatter_data_map:
                logging.warning(f"分類 '{classification_value}' のメタデータが見つかりません")
                continue
            
            meta_data = self.scatter_data_map[classification_value]
            cow_ids = [m[3] for m in meta_data]  # cow_id
            cow_auto_ids = [m[2] for m in meta_data]  # cow_auto_id
            
            # データの整合性チェック
            if len(x_values) != len(cow_ids) or len(y_values) != len(cow_ids):
                logging.error(f"分類 '{classification_value}' のデータ長が一致しません: x={len(x_values)}, y={len(y_values)}, cow_ids={len(cow_ids)}")
                continue
            
            # X軸が日付項目の場合は、エポック秒を日付オブジェクトに変換
            if is_date_axis:
                x_values = [datetime.fromtimestamp(x) for x in x_values]
            # X軸が月項目（分娩月など）の場合は、数値のまま（後でフォーマッターで処理）
            
            color = color_map.get(classification_value, 'steelblue')
            
            # 分類値の表示ラベルを取得
            value_label = self._get_classification_value_label(classification_value)
            
            # 分類がある場合は「分類名：値」の形式で表示
            if self.classification:
                if classification_display:
                    label = f"{classification_display}：{value_label}"
                else:
                    label = value_label
            else:
                label = value_label
            
            # 共通関数を使用して散布図を作成
            scatter = self._build_scatter_plot(
                ax=self.ax,
                x_values=x_values,
                y_values=y_values,
                cow_ids=cow_ids,
                cow_auto_ids=cow_auto_ids,
                x_name=x_item_display,
                y_name=y_item_display,
                label=label,
                color=color,
                alpha=0.9,
                s=50,
                picker=True  # picker=Trueに変更（自動的に適切な距離を設定）
            )
            
            if scatter:
                active_scatters.append(scatter)
        
        # アクティブなscatter artistをcanvasに保存
        self.canvas._falcon_active_scatters = active_scatters
        
        # 散布図の場合、アクティブなaxesをcanvasに保存
        if self.graph_type == "散布図":
            self.canvas._falcon_active_axes = self.ax
        
        self.ax.set_xlabel(x_item_display)
        self.ax.set_ylabel(y_item_display)
        self.ax.set_title(f"{y_item_display} vs {x_item_display}")
        self.ax.grid(True, alpha=0.3)
        
        # X軸が繁殖コード（RC）の場合は、目盛りラベルに意味を追加
        normalized_x_item = self.x_item.upper() if self.x_item else ""
        if normalized_x_item == "RC" or x_item_display in ("繁殖コード", "繁殖区分"):
            def rc_formatter_x(value, pos):
                """繁殖コードのフォーマッター（X軸用）"""
                try:
                    rc_num = int(value)
                    meaning = self._get_rc_meaning(rc_num)
                    if meaning:
                        return f"{rc_num}：{meaning}"
                    else:
                        return str(rc_num)
                except (ValueError, TypeError):
                    return str(value)
            from matplotlib.ticker import FuncFormatter
            self.ax.xaxis.set_major_formatter(FuncFormatter(rc_formatter_x))
        # X軸が生まれた年（BTHYR）の場合は整数年で表示（2024, 2025, 2026）
        elif normalized_x_item == "BTHYR" or x_item_display == "生まれた年":
            from matplotlib.ticker import FuncFormatter, MaxNLocator
            self.ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: str(int(v)) if v == int(v) else str(int(round(v)))))
            self.ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        # X軸が日付項目（CLVDなど）の場合は日付形式で表示
        elif is_date_axis:
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            self.ax.xaxis.set_major_locator(mdates.AutoDateLocator())
            # X軸のラベルを回転させて見やすくする
            plt.setp(self.ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        # X軸が月項目（分娩月など）の場合は「1月」「2月」...「12月」として表示
        elif is_month_axis:
            def month_formatter(value, pos):
                """月のフォーマッター"""
                try:
                    month_num = int(value)
                    if 1 <= month_num <= 12:
                        return f"{month_num}月"
                    else:
                        return str(month_num)
                except (ValueError, TypeError):
                    return str(value)
            from matplotlib.ticker import FuncFormatter
            self.ax.xaxis.set_major_formatter(FuncFormatter(month_formatter))
            # X軸の目盛りを1-12の整数に設定
            if classification_data_str:
                # すべてのX値を収集して、1-12の範囲で目盛りを設定
                all_x_values = set()
                for data_points in classification_data_str.values():
                    for x_val, _ in data_points:
                        if isinstance(x_val, (int, float)) and 1 <= x_val <= 12:
                            all_x_values.add(int(x_val))
                if all_x_values:
                    ticks = sorted(all_x_values)
                    self.ax.set_xticks(ticks)
                    self.ax.set_xticklabels([f"{t}月" for t in ticks])
        
        # Y軸が繁殖コード（RC）の場合は、目盛りラベルに意味を追加
        normalized_y_item = self.y_item.upper() if self.y_item else ""
        if normalized_y_item == "RC" or y_item_display in ("繁殖コード", "繁殖区分"):
            def rc_formatter(value, pos):
                """繁殖コードのフォーマッター"""
                try:
                    rc_num = int(value)
                    meaning = self._get_rc_meaning(rc_num)
                    if meaning:
                        return f"{rc_num}：{meaning}"
                    else:
                        return str(rc_num)
                except (ValueError, TypeError):
                    return str(value)
            from matplotlib.ticker import FuncFormatter
            self.ax.yaxis.set_major_formatter(FuncFormatter(rc_formatter))
        
        # Y軸の範囲を設定（マイナス値がある場合はデータ範囲に合わせる）
        y_min, y_max = self.ax.get_ylim()
        if y_min < 0:
            # マイナス値がある場合は0始まりにせず、データ範囲＋余白で表示
            margin = (y_max - y_min) * 0.05 or 1
            self.ax.set_ylim(bottom=y_min - margin, top=y_max + margin)
        else:
            # すべて正の値の場合は0起点
            self.ax.set_ylim(bottom=0, top=y_max if y_max > 0 else 1)
        
        # 凡例を表示（分類がある場合のみ）
        if self.classification and len(enabled_classifications) > 1:
            self.ax.legend(loc='best', fontsize=9, title=classification_display if classification_display else None)
        
        # 近似曲線を表示（チェックON時のみ）
        if self.show_trend_line_var and self.show_trend_line_var.get():
            self._draw_trend_line(
                classification_data_str=classification_data_str,
                enabled_classifications=enabled_classifications,
                is_date_axis=is_date_axis,
                x_item_display=x_item_display,
                y_item_display=y_item_display
            )
    
    def _draw_trend_line(self, classification_data_str: Dict[str, List[Tuple[float, float]]],
                         enabled_classifications: set, is_date_axis: bool,
                         x_item_display: str, y_item_display: str):
        """近似曲線（1次回帰）を描画し、式とR²を表示する"""
        x_all: List[float] = []
        y_all: List[float] = []
        for cat in enabled_classifications:
            if cat not in classification_data_str:
                continue
            for x_val, y_val in classification_data_str[cat]:
                if isinstance(x_val, datetime):
                    x_all.append(mdates.date2num(x_val))
                else:
                    x_all.append(float(x_val))
                y_all.append(float(y_val))
        if len(x_all) < 2:
            return
        x_arr = np.array(x_all)
        y_arr = np.array(y_all)
        # 1次回帰: y = a*x + b
        coefs = np.polyfit(x_arr, y_arr, 1)
        a, b = float(coefs[0]), float(coefs[1])
        y_pred = a * x_arr + b
        ss_res = np.sum((y_arr - y_pred) ** 2)
        ss_tot = np.sum((y_arr - np.mean(y_arr)) ** 2)
        r2 = (1.0 - ss_res / ss_tot) if ss_tot != 0 else 0.0
        # 近似曲線の描画用にX範囲で補間
        x_min, x_max = float(np.min(x_arr)), float(np.max(x_arr))
        x_line = np.linspace(x_min, x_max, 100)
        y_line = a * x_line + b
        if is_date_axis:
            x_line_dates = [datetime.fromtimestamp(float(t)) for t in x_line]
            self.ax.plot(x_line_dates, y_line, 'k--', linewidth=2, label='近似曲線', zorder=5)
        else:
            self.ax.plot(x_line, y_line, 'k--', linewidth=2, label='近似曲線', zorder=5)
        # 式とR²をグラフ内に表示（左上）
        eq_text = f"y = {a:.4g} x + {b:.4g}"
        r2_text = f"R² = {r2:.4f}"
        self.ax.text(0.02, 0.98, f"{eq_text}\n{r2_text}",
                     transform=self.ax.transAxes, fontsize=10, verticalalignment='top',
                     bbox=dict(boxstyle='round,pad=0.4', facecolor='white', alpha=0.9, edgecolor='gray'))
        # 凡例に近似曲線を追加したので再描画で反映される
        classification_display = self._get_classification_display_name()
        if self.classification and len(enabled_classifications) > 1:
            self.ax.legend(loc='best', fontsize=9, title=classification_display if classification_display else None)
        else:
            self.ax.legend(loc='best', fontsize=9)
    
    def _create_classification_checkboxes(self, classification_values: set):
        """分類項目のチェックボックスを作成"""
        if not self.classification_frame:
            return
        
        # 既存のチェックボックスをクリア
        for widget in self.classification_frame.winfo_children():
            widget.destroy()
        self.classification_vars.clear()
        
        # 分類値を文字列に統一してソート（産次の場合は数値順、産次分類の場合は1産、2産、3産以上の順）
        str_values = [str(v) for v in classification_values]
        def sort_key_checkbox(x):
            # 産次分類の場合（1産、2産、3産以上）
            if x == "1産":
                return (0, 1)
            elif x == "2産":
                return (0, 2)
            elif x == "3産以上":
                return (0, 3)
            # 分娩月の場合（1月、2月、...、12月）
            elif x.endswith("月") and x[:-1].isdigit():
                month = int(x[:-1])
                if 1 <= month <= 12:
                    return (1, month)
            # 空胎日数の場合（<100日、>=100日）
            elif "日" in x:
                if x.startswith("<"):
                    # 未満の場合は先に表示
                    try:
                        threshold = int(x[1:].replace("日", ""))
                        return (2, threshold)
                    except (ValueError, TypeError):
                        pass
                elif x.startswith(">="):
                    # 以上の場合は後に表示
                    try:
                        threshold = int(x[2:].replace("日", ""))
                        return (3, threshold)
                    except (ValueError, TypeError):
                        pass
            # DIM分類の場合（0-6, 7-13, 14-20, ...）
            elif "-" in x and x.count("-") == 1:
                try:
                    parts = x.split("-")
                    if len(parts) == 2 and parts[0].isdigit():
                        start = int(parts[0])
                        return (4, start)  # 開始値を基準にソート
                except (ValueError, TypeError):
                    pass
            # 数値の場合は数値順
            elif x.isdigit():
                return (5, int(x))
            # その他は文字列順
            return (6, x)
        
        sorted_values = sorted(str_values, key=sort_key_checkbox)
        
        # チェックボックスを作成（表示ラベルは産次なら「1産」などに変換）
        for classification_value in sorted_values:
            var = tk.BooleanVar(value=True)  # デフォルトでチェック
            self.classification_vars[classification_value] = var
            display_label = self._get_classification_value_label(classification_value) if self.classification else classification_value
            checkbox = ttk.Checkbutton(
                self.classification_frame,
                text=display_label,
                variable=var,
                command=self._on_classification_checkbox_changed
            )
            checkbox.pack(side=tk.LEFT, padx=5)
    
    def _on_classification_checkbox_changed(self):
        """分類項目のチェックボックスが変更された時のコールバック"""
        self._redraw_graph()
    
    def _on_trend_line_checkbox_changed(self):
        """近似曲線チェックボックスが変更された時のコールバック"""
        self._redraw_graph()
    
    def _redraw_graph(self):
        """グラフを再描画（チェックボックス変更時など）"""
        try:
            if self.graph_type == "空胎日数生存曲線":
                all_cows = self.db.get_all_cows()
                survival_data = self._compute_days_open_survival_data(all_cows)
                self.ax.clear()
                self._draw_survival_curve(survival_data)
            else:
                all_cows = self.db.get_all_cows()
                all_cows = self._filter_existing_cows(all_cows)
                self.ax.clear()
                if hasattr(self.canvas, "_falcon_active_scatters"):
                    self.canvas._falcon_active_scatters = []
                if self.graph_type == "散布図":
                    self._draw_scatter_chart(all_cows, create_checkboxes=False)
                elif self.graph_type == "棒グラフ":
                    self._draw_bar_chart(all_cows)
                elif self.graph_type == "箱ひげ":
                    self._draw_box_plot(all_cows)
                elif self.graph_type == "円グラフ":
                    self._draw_pie_chart(all_cows)
            self.canvas.draw()
            if self.graph_type == "散布図":
                self.canvas._falcon_active_axes = self.ax
        except Exception as e:
            logging.error(f"グラフ再描画エラー: {e}", exc_info=True)
    
    def _draw_box_plot(self, cows: List[Dict[str, Any]]):
        """箱ひげ図を描画"""
        classification_data: Dict[str, List[float]] = {}
        
        for cow in cows:
            # auto_idを取得（キー名を確認）
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID") or cow.get("AUTO_ID")
            if not cow_auto_id:
                continue
            
            # 条件チェック
            if not self._check_condition(cow, cow_auto_id):
                continue
            
            # 分類値を取得
            classification_value = self._get_classification_value(cow, cow_auto_id)
            if classification_value is None:
                continue
            
            # Y軸項目の値を取得
            y_value = self._get_item_value(cow, cow_auto_id, self.y_item)
            if y_value is not None:
                if classification_value not in classification_data:
                    classification_data[classification_value] = []
                classification_data[classification_value].append(y_value)
        
        if not classification_data:
            self.ax.text(0.5, 0.5, "データがありません", 
                        ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 箱ひげ図を描画
        categories = list(classification_data.keys())
        data = [classification_data[cat] for cat in categories]
        
        y_item_display = self._get_item_display_name(self.y_item)
        classification_display = self._get_item_display_name(self.classification) if self.classification else "分類"
        
        self.ax.boxplot(data, labels=categories)
        self.ax.set_xlabel(classification_display)
        self.ax.set_ylabel(y_item_display)
        self.ax.set_title(f"{y_item_display}の分布（{classification_display}別）")
        self.ax.grid(True, alpha=0.3)
        
        # Y軸の範囲を設定（マイナス値がある場合はデータ範囲に合わせる）
        y_min, y_max = self.ax.get_ylim()
        if y_min < 0:
            margin = (y_max - y_min) * 0.05 or 1
            self.ax.set_ylim(bottom=y_min - margin, top=y_max + margin)
        else:
            self.ax.set_ylim(bottom=0, top=y_max if y_max > 0 else 1)
    
    def _draw_pie_chart(self, cows: List[Dict[str, Any]]):
        """円グラフを描画"""
        classification_counts: Dict[str, int] = {}
        
        for cow in cows:
            # auto_idを取得（キー名を確認）
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID") or cow.get("AUTO_ID")
            if not cow_auto_id:
                continue
            
            # 条件チェック
            if not self._check_condition(cow, cow_auto_id):
                continue
            
            # 分類値を取得
            classification_value = self._get_classification_value(cow, cow_auto_id)
            if classification_value is not None:
                classification_counts[classification_value] = classification_counts.get(classification_value, 0) + 1
        
        if not classification_counts:
            self.ax.text(0.5, 0.5, "データがありません", 
                        ha='center', va='center', transform=self.ax.transAxes)
            return
        
        # 円グラフを描画
        categories = list(classification_counts.keys())
        counts = list(classification_counts.values())
        
        x_item_display = self._get_item_display_name(self.x_item)
        classification_display = self._get_item_display_name(self.classification) if self.classification else "分類"
        
        self.ax.pie(counts, labels=categories, autopct='%1.1f%%', startangle=90)
        self.ax.set_title(f"{x_item_display}の分布（{classification_display}別）")
    
    def _compute_days_open_survival_data(self, cows: List[Dict[str, Any]]) -> Dict[str, List[Tuple[int, int]]]:
        """
        空胎日数生存曲線用データを計算する。
        対象：経産牛（分娩イベントが1回以上ある個体）。
        イベント：受胎＝PDP/PDP2/PAGPの日付（AI/ETのコードPになった日）。
        打ち切り：除籍（売却・死亡廃用）または繁殖中止、または次の分娩、またはDIM220日。
        
        Returns:
            classification_value -> [(dim, event), ...]  (event=1: 受胎, 0: 打ち切り). dimは0〜220でキャップ。
        """
        X_AXIS_MAX = 220
        RE = RuleEngine
        conception_events = [RE.EVENT_PDP, RE.EVENT_PDP2, RE.EVENT_PAGP]
        censor_events = [RE.EVENT_SOLD, RE.EVENT_DEAD, RE.EVENT_STOPR]
        # 「イベントの有無」で使うイベント番号（ユーザー選択。未設定なら402のまま後方互換）
        event_number_for_presence = getattr(self, 'classification_event_number', None) or 402
        
        result: Dict[str, List[Tuple[int, int]]] = {}
        item_key = (getattr(self, 'classification_item_key', None) or "").strip().upper() or None
        bin_days = getattr(self, 'classification_bin_days', None)
        threshold = getattr(self, 'classification_threshold', None)
        
        def get_cv(lact: int, event_in_spell: bool, end_dim: int, cow: Dict[str, Any], cow_auto_id: int, clvd: str, t0: datetime) -> str:
            if not self.classification:
                return "全体"
            if self.classification == "産次" or self.classification.upper() == "LACT":
                return str(lact)
            if "産次（" in self.classification and "産以上" in self.classification:
                return "1産" if lact == 1 else "2産" if lact == 2 else "3産以上"
            if self.classification == "イベントの有無":
                return "あり" if event_in_spell else "なし"
            if self.classification == "分娩月":
                try:
                    month = t0.month if hasattr(t0, 'month') else int(clvd.split("-")[1])
                    return f"{month}月"
                except (IndexError, ValueError, TypeError):
                    return "不明"
            if self.classification == "条件で分類":
                return "満たす" if self._check_condition(cow, cow_auto_id) else "満たさない"
            if self.classification == "項目で分類" and item_key:
                raw = None
                if item_key in ("LACT", "産次"):
                    raw = lact
                elif item_key in ("CALVMO", "分娩月"):
                    try:
                        raw = t0.month if hasattr(t0, 'month') else int(clvd.split("-")[1])
                    except (IndexError, ValueError, TypeError):
                        raw = None
                else:
                    try:
                        if self.formula_engine:
                            calc = self.formula_engine.calculate(cow_auto_id, item_key)
                            if calc and item_key in calc:
                                raw = calc[item_key]
                    except Exception:
                        raw = None
                if raw is None:
                    return "不明"
                # 閾値が指定されていれば「X未満」「X以上」で二分
                if threshold is not None and isinstance(raw, (int, float)):
                    try:
                        v = float(raw)
                        # 閾値の表記は整数なら整数で、小数なら必要な桁で
                        t = threshold
                        if t == int(t):
                            label = str(int(t))
                        else:
                            label = str(t)
                        return f"{label}未満" if v < threshold else f"{label}以上"
                    except (TypeError, ValueError):
                        pass
                if bin_days and isinstance(raw, (int, float)):
                    base = int(float(raw) // bin_days) * bin_days
                    return f"{base}-{base + bin_days - 1}"
                if isinstance(raw, float) and raw == int(raw):
                    raw = int(raw)
                return str(raw)
            return "全体"
        
        for cow in cows:
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID") or cow.get("AUTO_ID")
            if not cow_auto_id or not self.db:
                continue
            # 「条件で分類」のときは条件で絞らず全頭を使い、各スペルを「満たす/満たさない」で分類する
            if self.condition and self.classification != "条件で分類" and not self._check_condition(cow, cow_auto_id):
                continue
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            events.sort(key=lambda e: (e.get("event_date") or "", e.get("id") or 0))
            calvings = [e for e in events if e.get("event_number") == RE.EVENT_CALV and e.get("event_date")]
            if not calvings:
                continue
            for i, calv in enumerate(calvings):
                clvd = calv.get("event_date")
                if not clvd:
                    continue
                # 産次＝項目LACT。分娩イベントの event_lact があればそれを使い、なければ分娩順（i+1）で補完
                lact = calv.get("event_lact")
                if lact is None or (isinstance(lact, (int, float)) and lact < 1):
                    lact = i + 1
                else:
                    lact = int(lact)
                try:
                    t0 = datetime.strptime(clvd, "%Y-%m-%d")
                except (ValueError, TypeError):
                    continue
                post = [e for e in events if e.get("event_date") and e.get("event_date") > clvd]
                post.sort(key=lambda e: (e.get("event_date"), e.get("id") or 0))
                end_dim = None
                event_flag = 0
                event_in_spell = False  # 選択イベントがこのスペル内にあったか
                for e in post:
                    d = e.get("event_date")
                    if not d:
                        continue
                    try:
                        t1 = datetime.strptime(d, "%Y-%m-%d")
                        dim = (t1 - t0).days
                    except (ValueError, TypeError):
                        continue
                    if dim > X_AXIS_MAX:
                        dim = X_AXIS_MAX
                    if e.get("event_number") == event_number_for_presence:
                        event_in_spell = True
                    if e.get("event_number") in conception_events:
                        end_dim = dim
                        event_flag = 1
                        break
                    if e.get("event_number") in censor_events:
                        end_dim = dim
                        event_flag = 0
                        break
                    if e.get("event_number") == RE.EVENT_CALV:
                        end_dim = dim
                        event_flag = 0
                        break
                if end_dim is None:
                    today = datetime.now().date()
                    try:
                        t0_date = t0.date() if hasattr(t0, 'date') else t0
                    except Exception:
                        t0_date = t0
                    dim_today = (today - t0_date).days if hasattr(today, '__sub__') else (today - t0).days
                    end_dim = min(X_AXIS_MAX, dim_today)
                    event_flag = 0
                cv = get_cv(lact, event_in_spell, end_dim, cow, cow_auto_id, clvd, t0)
                if cv not in result:
                    result[cv] = []
                result[cv].append((end_dim, event_flag))
        
        return result
    
    def _draw_survival_curve(self, data: Dict[str, List[Tuple[int, int]]]):
        """空胎日数生存曲線（Kaplan-Meier）を描画。X軸=DIM(0〜220)、Y軸=未受胎の割合（生存率）。"""
        if not data or all(not v for v in data.values()):
            self.ax.text(0.5, 0.5, "データがありません", ha='center', va='center', transform=self.ax.transAxes)
            return
        # 分類チェックボックスを1回だけ作成
        if self.classification and self.classification_frame and not self.classification_vars:
            self._create_classification_checkboxes({str(k) for k in data.keys()})
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']
        self.ax.clear()
        self.ax.set_xlim(0, 220)
        self.ax.set_ylim(0, 1.02)
        self.ax.set_xlabel("DIM（分娩後日数）", fontsize=11, fontweight='bold')
        self.ax.set_ylabel("未受胎の割合（生存率）", fontsize=11, fontweight='bold')
        self.ax.set_title("空胎日数生存曲線", fontsize=12, fontweight='bold', pad=15)
        # 条件文（P値は後で追記するためここではベースのみ）
        cond = "経産牛・受胎=P、打ち切り=除籍/繁殖中止/次分娩/DIM220"
        if self.classification == "条件で分類" and self.condition and self.condition.strip():
            cond += f"  分類条件：{self.condition.strip()}"
        self.ax.grid(True, linestyle='--', alpha=0.3)
        self.ax.set_axisbelow(True)
        
        enabled = set()
        if self.classification_vars:
            for cat, var in self.classification_vars.items():
                if var.get():
                    enabled.add(cat)
        else:
            enabled = {str(k) for k in data.keys()}
        
        # Log-rank検定（2群以上かつlifelines利用可能なとき）
        p_value_text = ""
        if _HAS_LIFELINES and len(enabled) >= 2:
            durations_list = []
            groups_list = []
            events_list = []
            for cv, spells in data.items():
                if str(cv) not in enabled or not spells:
                    continue
                for dim, ev in spells:
                    durations_list.append(dim)
                    groups_list.append(str(cv))
                    events_list.append(ev)
            if len(set(groups_list)) >= 2 and len(durations_list) >= 2:
                try:
                    result = multivariate_logrank_test(
                        np.array(durations_list, dtype=float),
                        np.array(groups_list),
                        np.array(events_list, dtype=int),
                    )
                    p_val = result.p_value
                    if p_val < 0.001:
                        p_value_text = "  Log-rank P<0.001"
                    else:
                        p_value_text = f"  Log-rank P={p_val:.3f}"
                except Exception as e:
                    logging.debug("Log-rank検定の計算に失敗: %s", e)
        
        # サブタイトルにP値を追記してから描画
        if p_value_text:
            cond = cond.rstrip()
            if not cond.endswith(p_value_text):
                cond += p_value_text
        self.ax.text(0.5, 1.0, cond, transform=self.ax.transAxes, fontsize=9,
                     ha='center', va='bottom', color='#444')
        
        for idx, (cv, spells) in enumerate(data.items()):
            if str(cv) not in enabled:
                continue
            n = len(spells)
            if n == 0:
                continue
            # 同じDIMでまとめる: dim -> (num_events, num_censored)
            by_dim: Dict[int, Tuple[int, int]] = {}
            for dim, ev in spells:
                if dim not in by_dim:
                    by_dim[dim] = (0, 0)
                if ev == 1:
                    by_dim[dim] = (by_dim[dim][0] + 1, by_dim[dim][1])
                else:
                    by_dim[dim] = (by_dim[dim][0], by_dim[dim][1] + 1)
            dims_sorted = sorted(by_dim.keys())
            at_risk = n
            survival = 1.0
            step_t = [0]
            step_s = [1.0]
            for dim in dims_sorted:
                d, c = by_dim[dim]
                if at_risk <= 0:
                    break
                if d > 0:
                    survival *= (1.0 - d / at_risk)
                step_t.append(dim)
                step_s.append(survival)
                at_risk -= d + c
            c = colors[idx % len(colors)]
            label = self._get_classification_value_label(str(cv)) if hasattr(self, '_get_classification_value_label') else str(cv)
            self.ax.step(step_t, step_s, where='post', label=f"{label} (n={n})", color=c, linewidth=2)
            
            # 75％・50％・25％到達日数を算出して保持（後で一括表示）
            def first_dim_below(step_t_list, step_s_list, target):
                for i, s in enumerate(step_s_list):
                    if s <= target:
                        return step_t_list[i] if i < len(step_t_list) else None
                return None
            dim_75 = first_dim_below(step_t, step_s, 0.75)
            dim_50 = first_dim_below(step_t, step_s, 0.50)
            dim_25 = first_dim_below(step_t, step_s, 0.25)
            if not hasattr(self.ax, '_survival_curve_milestones'):
                self.ax._survival_curve_milestones = []
            self.ax._survival_curve_milestones.append((label, n, dim_75, dim_50, dim_25, c))
            
            # 75%・50%・25%の位置に赤丸マーカーを描画（小さめで視認性を保ちつつ目立ちすぎない）
            marker_style = dict(markersize=5, markeredgecolor='#c62828', markerfacecolor='white', markeredgewidth=1.2, zorder=5)
            if dim_75 is not None:
                self.ax.plot(dim_75, 0.75, 'o', **marker_style)
            if dim_50 is not None:
                self.ax.plot(dim_50, 0.50, 'o', **marker_style)
            if dim_25 is not None:
                self.ax.plot(dim_25, 0.25, 'o', **marker_style)
        
        # 75％・50％・25％到達日数をグラフ内に表示（分類ごとに区切り・視認性向上）
        if hasattr(self.ax, '_survival_curve_milestones') and self.ax._survival_curve_milestones:
            lines = []
            milestones = self.ax._survival_curve_milestones
            for i, (label, _n, d75, d50, d25, _c) in enumerate(milestones):
                s75 = f"{d75}日" if d75 is not None else "—"
                s50 = f"{d50}日" if d50 is not None else "—"
                s25 = f"{d25}日" if d25 is not None else "—"
                if len(milestones) > 1:
                    lines.append("■ " + label)  # 分類名をマーカー付きで区切り
                lines.extend([f"75%  {s75}", f"50%  {s50}", f"25%  {s25}"])
                if len(milestones) > 1 and i < len(milestones) - 1:
                    lines.append("")  # カテゴリ間に空行
            self.ax.text(0.02, 0.02, "\n".join(lines), transform=self.ax.transAxes,
                        fontsize=11, verticalalignment='bottom', horizontalalignment='left',
                        bbox=dict(boxstyle='round,pad=0.5', facecolor='white', alpha=0.9, edgecolor='gray'))
            del self.ax._survival_curve_milestones
        
        if enabled:
            self.ax.legend(loc='upper right', fontsize=9)
        self.canvas.draw()
    
    def _get_item_display_name(self, item_key: str) -> str:
        """項目の表示名を取得"""
        if not item_key:
            return ""
        
        # 「産次」を「LACT」に変換
        lookup_key = item_key
        if item_key == "産次":
            lookup_key = "LACT"
        else:
            # 項目名を大文字に変換（item_dictionaryのキーは大文字）
            lookup_key = item_key.upper() if item_key else item_key
        
        # formula_engineからitem_dictionaryを取得
        item_dict = {}
        if self.formula_engine:
            try:
                # item_dictionaryを最新の状態に更新
                if hasattr(self.formula_engine, 'reload_item_dictionary'):
                    self.formula_engine.reload_item_dictionary()
                item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
            except Exception as e:
                logging.debug(f"item_dictionary取得エラー: {e}")
        
        # item_dictionaryから表示名を取得
        if lookup_key in item_dict:
            item_def = item_dict[lookup_key]
            # label, display_name, name_jp の順で取得
            display_name = item_def.get("label") or item_def.get("display_name") or item_def.get("name_jp") or item_key
            if display_name and display_name != item_key:
                return display_name
        
        # item_dictionaryに存在しない場合、デフォルトの表示名を返す
        # よく使われる項目のデフォルト表示名
        default_display_names = {
            "CLVD": "分娩月日",
            "BTHD": "生年月日",
            "DIM": "分娩後日数",
            "RC": "繁殖区分",
            "LACT": "産次",
            "JPN10": "個体識別番号（１０桁）",
            "COW_ID": "拡大4桁ID",
            "BRD": "品種"
        }
        if lookup_key in default_display_names:
            return default_display_names[lookup_key]
        
        return item_key
    
    def _get_classification_display_name(self) -> str:
        """分類の表示名を取得"""
        if not self.classification:
            return ""
        
        # 分類の表示名を取得
        classification_display = self._get_item_display_name(self.classification)
        
        # 分類名が取得できない場合は、元の分類名を使用
        if classification_display == self.classification:
            # 一部の分類名を日本語に変換
            if self.classification == "産次":
                classification_display = "産次"
            elif "産次（" in self.classification and "産以上" in self.classification:
                classification_display = "産次"
            elif self.classification.startswith("DIM") and self.classification[3:].isdigit():
                classification_display = f"分娩後日数（{self.classification[3:]}日ごと）"
            elif "空胎日数" in self.classification:
                # 「空胎日数100日（DOPN=100）」→「空胎日数100日」
                import re
                match = re.match(r'^空胎日数\s*(\d+)\s*日', self.classification)
                if match:
                    classification_display = f"空胎日数{match.group(1)}日"
                else:
                    classification_display = "空胎日数"
            elif self.classification == "分娩月":
                classification_display = "分娩月"
        
        return classification_display
    
    def _get_rc_meaning(self, rc_value: int) -> str:
        """繁殖コードの意味を取得（日本語のみ、コード番号：日本語で表示する際に使用）"""
        rc_meanings = {
            1: "繁殖停止",
            2: "分娩後",
            3: "授精後",
            4: "空胎",
            5: "妊娠中",
            6: "乾乳中"
        }
        return rc_meanings.get(rc_value, "")
    
    def _get_classification_value_label(self, classification_value: str) -> str:
        """分類値の表示ラベルを取得"""
        # 産次の分類の場合、「1産」「2産」「3産以上」のように表示
        if "産次（" in self.classification and "産以上" in self.classification:
            if classification_value == "1産":
                return "1産"
            elif classification_value == "2産":
                return "2産"
            elif classification_value == "3産以上":
                return "3産以上"
            else:
                return str(classification_value)
        # 産次の場合、「1産」「2産」などに変換
        elif self.classification == "産次" or self.classification.upper() == "LACT":
            try:
                lact_num = int(classification_value)
                return f"{lact_num}産"
            except (ValueError, TypeError):
                return str(classification_value)
        # 分娩月の場合、「1月」「2月」などに変換
        elif self.classification == "分娩月":
            if classification_value.endswith("月"):
                return classification_value
            else:
                try:
                    month_num = int(classification_value)
                    return f"{month_num}月"
                except (ValueError, TypeError):
                    return str(classification_value)
        # 繁殖コード（RC）の場合、意味を追加
        elif self.classification.upper() == "RC" or self.classification == "繁殖区分":
            try:
                # classification_valueが数値の場合も文字列の場合も対応
                if isinstance(classification_value, (int, float)):
                    rc_num = int(classification_value)
                else:
                    rc_num = int(str(classification_value))
                meaning = self._get_rc_meaning(rc_num)
                if meaning:
                    return f"{rc_num}：{meaning}"
                else:
                    return str(classification_value)
            except (ValueError, TypeError):
                return str(classification_value)
        # その他の場合はそのまま表示
        else:
            return str(classification_value)
    
    def _get_item_value_fallback_from_cow(self, cow: Dict[str, Any], normalized_item_key: str) -> Optional[float]:
        """formula_engine に項目がない場合に cow から直接取得（BTHD/CLVD/BTHYR/BTHYM）"""
        if not cow:
            return None
        bthd = cow.get("bthd") or cow.get("BTHD") or ""
        if isinstance(bthd, str) and bthd.strip():
            try:
                date_obj = datetime.strptime(bthd.strip(), "%Y-%m-%d")
                key = normalized_item_key.upper()
                if key in ("CLVD", "BTHD"):
                    return float(date_obj.timestamp())
                if key == "BTHYR":
                    return float(date_obj.year)
                if key == "BTHYM":
                    # 年月の1日としてタイムスタンプに変換
                    ym = date_obj.strftime("%Y-%m")
                    first_day = datetime.strptime(ym, "%Y-%m")
                    return float(first_day.timestamp())
            except (ValueError, TypeError):
                pass
        if normalized_item_key.upper() in ("CLVD", "BTHD"):
            clvd = cow.get("clvd") or cow.get("CLVD") or ""
            if isinstance(clvd, str) and clvd.strip():
                try:
                    date_obj = datetime.strptime(clvd.strip(), "%Y-%m-%d")
                    return float(date_obj.timestamp())
                except (ValueError, TypeError):
                    pass
        return None
    
    def _get_item_value(self, cow: Dict[str, Any], cow_auto_id: int, item_key: str) -> Optional[float]:
        """項目の値を取得"""
        if not self.formula_engine:
            return None
        
        # 項目名を大文字に変換（item_dictionaryのキーは大文字）
        normalized_item_key = item_key.upper() if item_key else item_key
        
        try:
            calculated = self.formula_engine.calculate(cow_auto_id, normalized_item_key)
            if not calculated:
                logging.debug(f"calculatedが空: item_key={normalized_item_key}, cow_auto_id={cow_auto_id}")
                # core項目（CLVD、BTHD）または bthd から算出する項目（BTHYR、BTHYM）のフォールバック
                value = self._get_item_value_fallback_from_cow(cow, normalized_item_key)
                if value is not None:
                    return value
                return None
            if normalized_item_key not in calculated:
                logging.debug(f"項目がcalculatedに存在しない: item_key={normalized_item_key}, calculated_keys={list(calculated.keys())}, cow_auto_id={cow_auto_id}")
                value = self._get_item_value_fallback_from_cow(cow, normalized_item_key)
                if value is not None:
                    return value
                return None
            
            value = calculated[normalized_item_key]
            if value is None:
                logging.debug(f"項目値がNone: item_key={normalized_item_key}, cow_auto_id={cow_auto_id}")
                return None
            
            # 空文字列の場合はNoneを返す
            if isinstance(value, str) and value.strip() == "":
                logging.debug(f"項目値が空文字列: item_key={normalized_item_key}, cow_auto_id={cow_auto_id}")
                return None
            
            # 文字列の場合は、日付形式または数値への変換を試みる
            if isinstance(value, str):
                # 日付形式（YYYY-MM-DD）をチェック
                try:
                    date_obj = datetime.strptime(value, "%Y-%m-%d")
                    return float(date_obj.timestamp())
                except (ValueError, TypeError):
                    pass
                # 年月形式（YYYY-MM、生まれた年月 BTHYM など）をチェック
                try:
                    date_obj = datetime.strptime(value.strip(), "%Y-%m")
                    # その月の1日としてエポック秒に変換（グラフ軸で時系列順になる）
                    return float(date_obj.timestamp())
                except (ValueError, TypeError):
                    pass
                # 数値変換を試みる（生まれた年が文字列の "2020" の場合など）
                try:
                    return float(value)
                except (ValueError, TypeError) as e:
                    logging.debug(f"数値変換エラー: item_key={normalized_item_key}, value={value}, type={type(value)}, error={e}, cow_auto_id={cow_auto_id}")
                    return None
            
            # 整数・浮動小数点の場合はそのまま数値として使用（生まれた年 BTHYR など）
            if isinstance(value, (int, float)):
                try:
                    return float(value)
                except (ValueError, TypeError):
                    pass
            # その他の型も数値変換を試みる
            try:
                return float(value)
            except (ValueError, TypeError) as e:
                logging.debug(f"数値変換エラー: item_key={normalized_item_key}, value={value}, type={type(value)}, error={e}, cow_auto_id={cow_auto_id}")
                return None
        except Exception as e:
            logging.error(f"項目値取得エラー: item_key={normalized_item_key}, cow_auto_id={cow_auto_id}, error={e}")
        
        return None
    
    def _get_classification_value(self, cow: Dict[str, Any], cow_auto_id: int) -> Optional[str]:
        """分類値を取得"""
        if not self.classification:
            return "全体"
        
        # LACT（産次）の特別処理：データベースから直接取得
        classification_upper = self.classification.upper()
        if classification_upper == "LACT" or self.classification == "産次":
            if cow and 'lact' in cow:
                lact_value = cow.get('lact')
                if lact_value is not None:
                    result = str(lact_value)
                    logging.debug(f"分類値（産次）取得成功: cow_auto_id={cow_auto_id}, lact={result}")
                    return result
            logging.debug(f"分類値（産次）が取得できません: cow_auto_id={cow_auto_id}, cow={cow}")
            return None
        
        # 産次（１，２，３産以上）の特別処理：産次を取得して分類
        if "産次（" in self.classification and "産以上" in self.classification:
            if cow and 'lact' in cow:
                try:
                    lact = int(cow.get('lact'))
                    if lact == 1:
                        result = "1産"
                    elif lact == 2:
                        result = "2産"
                    else:
                        result = "3産以上"
                    logging.debug(f"分類値（産次分類）取得成功: cow_auto_id={cow_auto_id}, lact={lact}, category={result}")
                    return result
                except (ValueError, TypeError):
                    logging.debug(f"分類値（産次分類）の変換エラー: cow_auto_id={cow_auto_id}, lact={cow.get('lact')}")
            logging.debug(f"分類値（産次分類）が取得できません: cow_auto_id={cow_auto_id}, cow={cow}")
            return None
        
        # DIM分類の特別処理（DIM7、DIM14、DIM21、DIM30、DIM50など）
        if self.classification.upper().startswith("DIM") and self.classification.upper()[3:].isdigit():
            # DIMの数値を取得（例：DIM7 → 7）
            try:
                interval = int(self.classification.upper()[3:])
                # DIMの値を取得
                if not self.formula_engine:
                    logging.warning(f"FormulaEngineが設定されていません: cow_auto_id={cow_auto_id}")
                    return None
                
                calculated = self.formula_engine.calculate(cow_auto_id)
                if calculated and 'DIM' in calculated:
                    dim_value = calculated['DIM']
                    if dim_value is not None:
                        try:
                            dim_int = int(dim_value)
                            # 区間で分類（例：DIM7の場合、0-6, 7-13, 14-20, ...）
                            start = dim_int // interval * interval
                            end = (dim_int // interval + 1) * interval - 1
                            result = f"{start}-{end}"
                            logging.debug(f"分類値（DIM分類）取得成功: cow_auto_id={cow_auto_id}, DIM={dim_int}, interval={interval}, category={result}")
                            return result
                        except (ValueError, TypeError):
                            logging.debug(f"分類値（DIM分類）の変換エラー: cow_auto_id={cow_auto_id}, DIM={dim_value}")
                else:
                    logging.debug(f"DIMが取得できません: cow_auto_id={cow_auto_id}, calculated={calculated}")
            except (ValueError, TypeError):
                logging.debug(f"DIM分類のinterval取得エラー: classification={self.classification}")
        
        # 分娩月の特別処理
        if self.classification == "分娩月":
            if not self.formula_engine:
                logging.warning(f"FormulaEngineが設定されていません: cow_auto_id={cow_auto_id}")
                return None
            
            try:
                calculated = self.formula_engine.calculate(cow_auto_id)
                if calculated and 'CALVMO' in calculated:
                    calvmo_value = calculated['CALVMO']
                    if calvmo_value is not None:
                        try:
                            calvmo_int = int(calvmo_value)
                            # 分娩月は1-12の値
                            if 1 <= calvmo_int <= 12:
                                result = f"{calvmo_int}月"
                                logging.debug(f"分類値（分娩月）取得成功: cow_auto_id={cow_auto_id}, CALVMO={calvmo_int}, category={result}")
                                return result
                        except (ValueError, TypeError):
                            logging.debug(f"分類値（分娩月）の変換エラー: cow_auto_id={cow_auto_id}, CALVMO={calvmo_value}")
                else:
                    logging.debug(f"CALVMOが取得できません: cow_auto_id={cow_auto_id}, calculated={calculated}")
            except Exception as e:
                logging.debug(f"分類値（分娩月）取得エラー: cow_auto_id={cow_auto_id}, error={e}")
            return None
        
        # 空胎日数の特別処理（空胎日数100日、空胎日数150日など）
        if "空胎日数" in self.classification and "日" in self.classification:
            # 閾値を抽出（例：「空胎日数100日（DOPN=100）」→ 100、「空胎日数100日」→ 100）
            import re
            threshold = None
            # 「空胎日数100日（DOPN=100）」の形式を検出
            match_full = re.match(r'^空胎日数\s*(\d+)\s*日\s*[（(]DOPN\s*=\s*\d+[）)]', self.classification)
            if match_full:
                threshold = int(match_full.group(1))
            else:
                # 「空胎日数100日」の形式を検出
                match_simple = re.match(r'^空胎日数\s*(\d+)\s*日', self.classification)
                if match_simple:
                    threshold = int(match_simple.group(1))
            
            if threshold is not None:
                if not self.formula_engine:
                    logging.warning(f"FormulaEngineが設定されていません: cow_auto_id={cow_auto_id}")
                    return None
                
                try:
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and 'DOPN' in calculated:
                        dopn_value = calculated['DOPN']
                        if dopn_value is not None:
                            try:
                                dopn_int = int(dopn_value)
                                # 閾値に基づいて分類（未満と以上）
                                if dopn_int < threshold:
                                    result = f"<{threshold}日"
                                else:
                                    result = f">={threshold}日"
                                logging.debug(f"分類値（空胎日数）取得成功: cow_auto_id={cow_auto_id}, DOPN={dopn_int}, threshold={threshold}, category={result}")
                                return result
                            except (ValueError, TypeError):
                                logging.debug(f"分類値（空胎日数）の変換エラー: cow_auto_id={cow_auto_id}, DOPN={dopn_value}")
                    else:
                        logging.debug(f"DOPNが取得できません: cow_auto_id={cow_auto_id}, calculated={calculated}")
                except Exception as e:
                    logging.debug(f"分類値（空胎日数）取得エラー: cow_auto_id={cow_auto_id}, error={e}")
            return None
        
        # その他の分類項目はFormulaEngineで計算
        if not self.formula_engine:
            logging.warning(f"FormulaEngineが設定されていません: cow_auto_id={cow_auto_id}")
            return None
        
        try:
            calculated = self.formula_engine.calculate(cow_auto_id, self.classification)
            if calculated:
                if self.classification in calculated:
                    value = calculated[self.classification]
                    if value is not None:
                        result = str(value)
                        logging.debug(f"分類値取得成功: cow_auto_id={cow_auto_id}, classification={self.classification}, value={result}")
                        return result
                    else:
                        logging.debug(f"分類値がNone: cow_auto_id={cow_auto_id}, classification={self.classification}")
                else:
                    logging.debug(f"分類項目が計算結果に含まれていません: cow_auto_id={cow_auto_id}, classification={self.classification}, calculated_keys={list(calculated.keys())}")
            else:
                logging.debug(f"計算結果が空: cow_auto_id={cow_auto_id}, classification={self.classification}")
        except Exception as e:
            logging.error(f"分類値取得エラー: cow_auto_id={cow_auto_id}, classification={self.classification}, error={e}", exc_info=True)
        
        return None
    
    def _build_scatter_plot(self, ax, x_values: List[float], y_values: List[float], 
                           cow_ids: List[str], cow_auto_ids: List[int],
                           x_name: str, y_name: str, label: str = "",
                           color: Any = 'steelblue', alpha: float = 0.9, 
                           s: int = 50, picker: int = 5) -> Any:
        """
        共通の散布図生成関数
        
        Args:
            ax: matplotlib axes
            x_values: X軸の値のリスト
            y_values: Y軸の値のリスト
            cow_ids: 個体ID（拡大4桁ID）のリスト
            cow_auto_ids: 個体のauto_idのリスト
            x_name: X軸項目名
            y_name: Y軸項目名
            label: 凡例ラベル（任意）
            color: プロットの色
            alpha: 透明度
            s: プロットサイズ
            picker: pickerの許容距離（ピクセル）
        
        Returns:
            scatter artist
        """
        # データの整合性チェック
        n_points = len(x_values)
        if len(y_values) != n_points or len(cow_ids) != n_points or len(cow_auto_ids) != n_points:
            logging.error(f"データの長さが一致しません: x={len(x_values)}, y={len(y_values)}, cow_ids={len(cow_ids)}, cow_auto_ids={len(cow_auto_ids)}")
            raise ValueError("データの長さが一致しません")
        
        if n_points == 0:
            logging.warning("散布図のデータポイントが0です")
            return None
        
        # scatter artistを作成（pickerを有効化）
        # picker=True にすることで、matplotlibが自動的に適切な距離を設定
        scatter = ax.scatter(
            x_values, y_values, 
            alpha=alpha, s=s, label=label, color=color, 
            edgecolors='black', linewidths=0.5,
            picker=True  # pickerを有効化（Trueにすると自動的に適切な距離を設定）
        )
        
        # メタデータを付与
        scatter._falcon_meta = {
            "cow_ids": cow_ids,
            "cow_auto_ids": cow_auto_ids,
            "x": x_values,
            "y": y_values,
            "x_name": x_name,
            "y_name": y_name,
            "label": label
        }
        
        
        
        return scatter
    
    def _check_condition(self, cow: Dict[str, Any], cow_auto_id: int) -> bool:
        """条件をチェック"""
        if not self.condition:
            return True
        
        # 条件の解析とチェック（簡易版）
        # 実際の実装では、より詳細な条件解析が必要
        try:
            if "：" in self.condition or ":" in self.condition:
                # 項目名：値の形式
                if "：" in self.condition:
                    parts = self.condition.split("：", 1)
                else:
                    parts = self.condition.split(":", 1)
                item_name = parts[0].strip()
                expected_value = parts[1].strip()
                
                # 項目名を大文字に変換（item_dictionaryのキーは大文字）
                normalized_item_name = item_name.upper() if item_name else item_name
                
                calculated = self.formula_engine.calculate(cow_auto_id, normalized_item_name)
                if calculated and normalized_item_name in calculated:
                    actual_value = str(calculated[normalized_item_name])
                    return actual_value == expected_value
            else:
                # 比較演算子を含む形式（簡易版）
                import re
                match = re.match(r'^([A-Za-z_][A-Za-z0-9_]*)\s*([<>=!]+)\s*(.+)$', self.condition.strip())
                if match:
                    item_name = match.group(1)
                    operator = match.group(2)
                    threshold = match.group(3)
                    
                    # 項目名を大文字に変換（item_dictionaryのキーは大文字）
                    normalized_item_name = item_name.upper() if item_name else item_name
                    
                    calculated = self.formula_engine.calculate(cow_auto_id, normalized_item_name)
                    if calculated and normalized_item_name in calculated:
                        value = calculated[normalized_item_name]
                        if value is not None:
                            try:
                                value_float = float(value)
                                threshold_float = float(threshold)
                                if operator == "<":
                                    return value_float < threshold_float
                                elif operator == ">":
                                    return value_float > threshold_float
                                elif operator == "<=":
                                    return value_float <= threshold_float
                                elif operator == ">=":
                                    return value_float >= threshold_float
                                elif operator == "=" or operator == "==":
                                    return value_float == threshold_float
                            except (ValueError, TypeError):
                                pass
        except Exception as e:
                logging.debug(f"条件チェックエラー: {e}")
        
        return True
    
    def _on_scatter_click(self, event):
        """散布図のプロットがクリックされた時の処理（ダブルクリックで個体カードを開く）"""
        # 散布図でない場合は何もしない
        if self.graph_type != "散布図":
            logging.debug(f"グラフタイプが散布図ではありません: {self.graph_type}")
            return
        
        # クリックがグラフ領域内かチェック（アクティブなaxesを使用）
        active_ax = getattr(self.canvas, '_falcon_active_axes', self.ax)
        if event.inaxes != active_ax:
            logging.debug(f"クリックがグラフ領域外: event.inaxes={event.inaxes}, active_ax={active_ax}, self.ax={self.ax}")
            return
        
        # ダブルクリックのみ反応（event.dblclickがNoneの場合は、時間ベースの判定を行う）
        import time
        current_time = time.time()
        is_double_click = False
        
        if getattr(event, 'dblclick', False):
            is_double_click = True
            logging.info("event.dblclickがTrue")
        elif hasattr(self, 'last_click_time') and self.last_click_time > 0 and self.last_click_x is not None and self.last_click_y is not None:
            time_since_last_click = current_time - self.last_click_time
            click_x_pixel = event.x
            click_y_pixel = event.y
            position_diff = ((click_x_pixel - self.last_click_x) ** 2 + (click_y_pixel - self.last_click_y) ** 2) ** 0.5
            is_double_click = (time_since_last_click < 0.3) and (position_diff < 10)
            
            if is_double_click:
                pass
        
        if not is_double_click:
            # クリック情報を記録
            if not hasattr(self, 'last_click_time'):
                self.last_click_time = 0
            self.last_click_time = current_time
            self.last_click_x = event.x
            self.last_click_y = event.y
            return
        
        # アクティブなscatter artistを取得
        if not hasattr(self.canvas, "_falcon_active_scatters") or not self.canvas._falcon_active_scatters:
            return
        
        # すべてのアクティブなscatter artistをチェック
        for i, scatter in enumerate(self.canvas._falcon_active_scatters):
            if not hasattr(scatter, '_falcon_meta'):
                logging.warning(f"scatter[{i}]に_falcon_metaが存在しません")
                continue
            
            # scatter.contains()で当たり判定
            try:
                contains, ind = scatter.contains(event)
                if contains and ind and len(ind['ind']) > 0:
                    # ヒットした点のインデックスを取得
                    hit_indices = ind['ind']
                    hit_index = hit_indices[0]  # 最初のヒット点を使用
                    
                    meta = scatter._falcon_meta
                    if hit_index < len(meta['cow_auto_ids']):
                        cow_auto_id = meta['cow_auto_ids'][hit_index]
                        cow_id = meta['cow_ids'][hit_index] if hit_index < len(meta['cow_ids']) else str(cow_auto_id)
                        
                        # 個体カードを開く
                        self._open_cow_card(cow_auto_id)
                        return
                    else:
                        pass
            except Exception as e:
                logging.error(f"scatter[{i}] contains判定エラー: {e}", exc_info=True)
                continue
        
        pass
    
    def _on_scatter_hover(self, event):
        """散布図のホバーイベント（マウスをプロットポイントに合わせたときにデータを表示）"""
        if not event.inaxes:
            # ツールチップを非表示
            if self.tooltip_annotation:
                self.tooltip_annotation.set_visible(False)
                self.canvas.draw_idle()
            return
        
        if self.graph_type != "散布図":
            # ツールチップを非表示
            if self.tooltip_annotation:
                self.tooltip_annotation.set_visible(False)
                self.canvas.draw_idle()
            return
        
        # アクティブなaxesと一致するかチェック
        active_ax = getattr(self.canvas, '_falcon_active_axes', self.ax)
        if event.inaxes != active_ax:
            # ツールチップを非表示
            if self.tooltip_annotation:
                self.tooltip_annotation.set_visible(False)
                self.canvas.draw_idle()
            return
        
        # アクティブなscatter artistを取得
        if not hasattr(self.canvas, "_falcon_active_scatters") or not self.canvas._falcon_active_scatters:
            if self.tooltip_annotation:
                self.tooltip_annotation.set_visible(False)
                self.canvas.draw_idle()
            return
        
        # すべてのアクティブなscatter artistをチェック
        hit_scatter = None
        hit_index = None
        
        for i, scatter in enumerate(self.canvas._falcon_active_scatters):
            if not hasattr(scatter, '_falcon_meta'):
                logging.warning(f"scatter[{i}]に_falcon_metaが存在しません")
                continue
            
            # scatter.contains()で当たり判定
            try:
                contains, ind = scatter.contains(event)
                if contains and ind and len(ind['ind']) > 0:
                    # ヒットした点のインデックスを取得
                    hit_indices = ind['ind']
                    hit_index = hit_indices[0]  # 最初のヒット点を使用
                    hit_scatter = scatter
                    break
            except Exception as e:
                logging.error(f"scatter[{i}] contains判定エラー: {e}", exc_info=True)
                continue
        
        # ヒットした場合、ツールチップを表示
        if hit_scatter and hit_index is not None:
            meta = hit_scatter._falcon_meta
            
            if hit_index >= len(meta['cow_ids']) or hit_index >= len(meta['cow_auto_ids']):
                logging.warning(f"hit_index={hit_index}が範囲外です: cow_ids={len(meta['cow_ids'])}, cow_auto_ids={len(meta['cow_auto_ids'])}")
                return
            
            cow_id = meta['cow_ids'][hit_index]
            cow_auto_id = meta['cow_auto_ids'][hit_index]
            x_val = meta['x'][hit_index]
            y_val = meta['y'][hit_index]
            x_name = meta['x_name']
            y_name = meta['y_name']
            label = meta.get('label', '')
            
            # X値とY値の表示形式を決定
            if isinstance(x_val, datetime):
                x_display = x_val.strftime('%Y-%m-%d')
            elif isinstance(x_val, (int, float)):
                if self.x_item.upper() == "CALVMO" or x_name == "分娩月":
                    x_display = f"{int(x_val)}月"
                else:
                    x_display = f"{x_val:.1f}" if isinstance(x_val, float) else str(int(x_val))
            else:
                x_display = str(x_val)
            
            # Y値の表示形式を決定（RCの場合は意味も表示）
            if self.y_item.upper() == "RC":
                rc_meaning = self._get_rc_meaning(int(y_val))
                y_display = f"{int(y_val)}: {rc_meaning}"
            elif isinstance(y_val, (int, float)):
                y_display = f"{y_val:.1f}" if isinstance(y_val, float) else str(int(y_val))
            else:
                y_display = str(y_val)
            
            # ツールチップのテキストを作成（拡大4桁IDを表示）
            tooltip_text = f"ID: {cow_id}\n{x_name}: {x_display}\n{y_name}: {y_display}"
            if label:
                tooltip_text += f"\n{label}"
            
            # ツールチップの座標を決定
            tooltip_x = x_val
            tooltip_y = y_val
            
            # ツールチップを表示または更新
            if self.tooltip_annotation is None:
                self.tooltip_annotation = self.ax.annotate(
                    tooltip_text,
                    xy=(tooltip_x, tooltip_y),
                    xytext=(10, 10),
                    textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.5', facecolor='#FFFACD', alpha=0.9, edgecolor='#CCCCCC', linewidth=1),
                    arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'),
                    fontsize=9,
                    zorder=100
                )
            else:
                self.tooltip_annotation.set_text(tooltip_text)
                self.tooltip_annotation.xy = (tooltip_x, tooltip_y)
                self.tooltip_annotation.set_visible(True)
            
            self.canvas.draw_idle()
        else:
            # 近くにデータポイントがない場合はツールチップを非表示
            if self.tooltip_annotation:
                self.tooltip_annotation.set_visible(False)
                self.canvas.draw_idle()
    
    def _open_cow_card(self, cow_auto_id: int):
        """個体カードを開く（別ウィンドウとして開く）"""
        logging.info(f"個体カードを開く: cow_auto_id={cow_auto_id}, event_dictionary_path={self.event_dictionary_path is not None}")
        # グラフウィンドウからは常に別ウィンドウ（CowCardWindow）として開く
        if self.event_dictionary_path:
            try:
                from ui.cow_card_window import CowCardWindow
                from pathlib import Path
                
                cow_card_window = CowCardWindow(
                    parent=self.window,
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=Path(self.event_dictionary_path),
                    cow_auto_id=cow_auto_id
                )
                
                # show()を呼び出す（ウィンドウが既に表示されている場合でも確実に前面に表示）
                cow_card_window.show()
            except Exception as e:
                logging.error(f"個体カード表示エラー: {e}", exc_info=True)
                from tkinter import messagebox
                messagebox.showerror("エラー", f"個体カードを開くことができませんでした: {e}")
        else:
            logging.warning(f"個体カードを開くための情報が不足しています: event_dictionary_path={self.event_dictionary_path}")
            from tkinter import messagebox
            messagebox.showerror("エラー", "個体カードを開くための情報が不足しています。")