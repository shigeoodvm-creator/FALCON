"""
FALCON2 - CowCard（個体カード）
1頭の牛の情報を表示する（表示専用）
設計書 第11章・第13章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable, List
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.event_display import format_insemination_event, format_calving_event, format_reproduction_check_event, build_ai_et_event_note
from modules.app_settings_manager import get_app_settings_manager
from ui.event_input import EventInputWindow
from settings_manager import SettingsManager
from constants import FARMS_ROOT

# EventDetailWindow は後で実装（今はダミー）
try:
    from ui.event_detail_window import EventDetailWindow
except ImportError:
    # EventDetailWindow が存在しない場合はダミークラス
    class EventDetailWindow:
        def __init__(self, *args, **kwargs):
            pass
        def show(self):
            pass


class CowCard:
    """個体カード（表示専用）"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler, 
                 formula_engine: FormulaEngine,
                 rule_engine: Optional[RuleEngine] = None,
                 event_dictionary_path: Optional[Path] = None,
                 item_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス（イベント追加ボタン用）
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス
        """
        try:
            print("DEBUG: CowCard.__init__ 開始")
            
            self.db = db_handler
            self.formula_engine = formula_engine
            self.rule_engine = rule_engine  # None でもOK
            self.event_dict_path = event_dictionary_path
            self.item_dict_path = item_dictionary_path
            self.calculated_items_path = Path("config/cow_card_calculated_items.json")
            self.calc_layout_path = Path("config/cow_card_calc_layout.json")
            self.calculated_item_codes: List[str] = self._load_calculated_item_codes()
            # 各タブの並び順を保持する辞書
            self.tab_orders: Dict[str, List[str]] = self._load_tab_orders()
            self.cow_auto_id: Optional[int] = None
            self.settings_manager: Optional[SettingsManager] = None  # 後で初期化
            self.app_settings = get_app_settings_manager()  # アプリ設定マネージャー
            
            # 授精設定辞書（code -> name の形式で保持）
            self.technicians_dict: Dict[str, str] = {}  # {"1": "Sonoda"}
            self.insemination_types_dict: Dict[str, str] = {}  # {"1": "自然発情"}
            # PEN設定（code -> name）
            self.pen_settings: Dict[str, str] = {}
            
            # event_dictionary を読み込む
            self.event_dictionary: Dict[str, Dict[str, Any]] = {}
            self._load_event_dictionary()
            
            # 使用される色タグを追跡（動的にタグを設定するため）
            self._configured_color_tags: set = set()
            
            # 繁殖コード表示名：コード番号：日本語のみ（英語表記は省く）
            self.rc_names = {
                1: "1：繁殖停止",
                2: "2：分娩後",
                3: "3：授精後",
                4: "4：空胎",
                5: "5：妊娠中",
                6: "6：乾乳中"
            }
            self.rc_names_jp = {
                1: "繁殖停止",
                2: "分娩後",
                3: "授精後",
                4: "空胎",
                5: "妊娠中",
                6: "乾乳中"
            }
            
            # UI作成
            self.frame = ttk.Frame(parent)
            self._create_widgets()
            
            # イベント追加ボタン用のコールバック
            self.on_event_saved_callback: Optional[Callable] = None
            
            # 最終日付計算用のコールバック（latest_calving, latest_ai, latest_milk, latest_any を通知）
            self.on_latest_dates_calculated_callback: Optional[Callable[[Optional[str], Optional[str], Optional[str], Optional[str]], None]] = None
            
            print("DEBUG: CowCard.__init__ 完了")
            
        except Exception as e:
            import traceback
            print("ERROR: CowCard.__init__ で例外が発生しました:")
            traceback.print_exc()
            raise
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
    
    @staticmethod
    def get_latest_event_date(events: List[Dict[str, Any]], target_numbers: List[int]) -> Optional[str]:
        """
        指定されたイベント番号の中で最新のevent_dateを取得
        
        Args:
            events: イベントリスト
            target_numbers: 対象となるイベント番号のリスト
        
        Returns:
            最新のevent_date（YYYY-MM-DD形式）、存在しない場合はNone
        """
        from datetime import datetime
        
        latest_date = None
        latest_datetime = None
        
        for e in events:
            event_number = e.get("event_number")
            event_date = e.get("event_date")
            
            # event_numberが整数型でない場合は変換を試みる
            if event_number is not None:
                try:
                    event_number = int(event_number)
                except (ValueError, TypeError):
                    continue
            
            # 対象イベント番号で、event_dateが有効な場合のみ処理
            if event_number in target_numbers and event_date:
                try:
                    # 日付をdatetimeオブジェクトに変換して比較
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    if latest_datetime is None or event_dt > latest_datetime:
                        latest_datetime = event_dt
                        latest_date = event_date
                except (ValueError, TypeError):
                    # 日付形式が不正な場合はスキップ
                    continue
        
        return latest_date
    
    @staticmethod
    def get_latest_any_event_date(events: List[Dict[str, Any]]) -> Optional[str]:
        """
        全イベントの中で最新のevent_dateを取得
        
        Args:
            events: イベントリスト
        
        Returns:
            最新のevent_date（YYYY-MM-DD形式）、存在しない場合はNone
        """
        from datetime import datetime
        
        latest_date = None
        latest_datetime = None
        
        for e in events:
            event_date = e.get("event_date")
            if event_date:
                try:
                    # 日付をdatetimeオブジェクトに変換して比較
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    if latest_datetime is None or event_dt > latest_datetime:
                        latest_datetime = event_dt
                        latest_date = event_date
                except (ValueError, TypeError):
                    # 日付形式が不正な場合はスキップ
                    continue
        
        return latest_date
    
    def _get_event_name(self, event_number: int) -> str:
        """
        イベント番号からイベント名を取得
        
        Args:
            event_number: イベント番号
        
        Returns:
            イベント名（日本語、短縮版）
        """
        event_str = str(event_number)
        name_jp = None
        
        if event_str in self.event_dictionary:
            name_jp = self.event_dictionary[event_str].get('name_jp')
        
        # event_dictionary.jsonにない場合のフォールバック
        if not name_jp:
            # RuleEngineの定義に基づくデフォルト名
            default_names = {
                RuleEngine.EVENT_AI: "AI",
                RuleEngine.EVENT_ET: "ET",
                RuleEngine.EVENT_CALV: "分娩",
                RuleEngine.EVENT_DRY: "乾乳",
                RuleEngine.EVENT_STOPR: "繁殖停止",
                RuleEngine.EVENT_SOLD: "売却",
                RuleEngine.EVENT_DEAD: "死亡・淘汰",
                RuleEngine.EVENT_PDN: "妊娠鑑定マイナス",
                RuleEngine.EVENT_PDP: "妊娠鑑定プラス",
                RuleEngine.EVENT_PDP2: "妊娠鑑定プラス（検診以外）",
                RuleEngine.EVENT_ABRT: "流産",
                RuleEngine.EVENT_PAGN: "PAGマイナス",
                RuleEngine.EVENT_PAGP: "PAGプラス",
                RuleEngine.EVENT_MILK_TEST: "乳検",
                RuleEngine.EVENT_MOVE: "群変更"
            }
            name_jp = default_names.get(event_number, f'イベント{event_number}')
        
        # イベント名の短縮処理（個体カードのイベント履歴表示用）
        if event_number == RuleEngine.EVENT_PDN:
            # 妊娠鑑定マイナス → 妊鑑－
            return "妊鑑－"
        elif event_number == RuleEngine.EVENT_PDP or event_number == RuleEngine.EVENT_PDP2:
            # 妊娠鑑定プラス → 妊鑑＋
            return "妊鑑＋"
        elif event_number == 300:
            # フレッシュチェック → フレチェック
            if "フレッシュチェック" in name_jp:
                return "フレチェック"
        
        return name_jp
    
    
    def _ensure_color_tag(self, color: str, event_number: int) -> str:
        """
        色に対応するタグを確保（まだ設定されていない場合は設定する）
        
        Args:
            color: 色名（Tkinterが解釈できる色名または#RRGGBB形式）
            event_number: イベント番号（分娩イベントの場合は太字にするため）
        
        Returns:
            タグ名
        """
        # タグ名は色名を使用（#RRGGBB形式の場合はそのまま使用）
        tag_name = color
        
        # まだ設定されていない場合は設定
        if tag_name not in self._configured_color_tags:
            # イベント履歴は MS ゴシック
            font_family = "MS Gothic"
            font_size = self.app_settings.get_font_size()
            # 分娩イベントの場合は太字も設定
            if event_number == RuleEngine.EVENT_CALV:
                event_font = (font_family, font_size, 'bold')
                self.event_tree.tag_configure(tag_name, foreground=color, font=event_font)
            else:
                event_font = (font_family, font_size)
                self.event_tree.tag_configure(tag_name, foreground=color, font=event_font)
            self._configured_color_tags.add(tag_name)
        
        return tag_name
    
    def _get_event_display_color(self, event_number: int) -> str:
        """
        イベントの表示色を決定
        
        【優先順位】
        1. event_dictionary.json の display_color が指定されていればそれを使用
        2. category / outcome に応じてデフォルト色を使用
           - CALVING → #0066cc（青）
           - PREGNANCY + outcome=NEGATIVE → #cc0000（赤）
           - PREGNANCY + outcome=POSITIVE（または未指定） → #008000（緑）
           - REPRODUCTION → #000000（黒）
        3. それ以外は #000000（黒）
        
        Args:
            event_number: イベント番号
        
        Returns:
            色（#RRGGBB形式）
        """
        # イベント辞書から情報を取得
        event_str = str(event_number)
        event_dict = self.event_dictionary.get(event_str, {})
        
        # 1. display_color が指定されていればそれを使用（最優先）
        display_color = event_dict.get('display_color')
        if display_color:
            return display_color
        
        # 2. category / outcome に応じてデフォルト色を使用
        category = event_dict.get('category', '')
        
        if category == "CALVING":
            return "#0066cc"  # 青
        elif category == "PREGNANCY":
            outcome = event_dict.get('outcome', '')
            if outcome == "NEGATIVE":
                return "#cc0000"  # 赤
            else:
                return "#008000"  # 緑（POSITIVE または未指定）
        elif category == "REPRODUCTION":
            return "#000000"  # 黒
        
        # 3. それ以外は黒
        return "#000000"
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # スクロール可能なCanvasを作成
        canvas = tk.Canvas(self.frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # スクロール可能なフレームをCanvasに配置
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Canvasのリサイズ時：埋め込みウィンドウをCanvasいっぱいに（項目欄が下まで拡張される）
            if event.widget == canvas:
                try:
                    cw = canvas.winfo_width()
                    ch = canvas.winfo_height()
                    if cw > 1 and ch > 1:
                        canvas.itemconfig(canvas_window, width=cw, height=ch)
                except tk.TclError:
                    pass
            # 幅は常にCanvasに合わせる
            try:
                canvas_width = canvas.winfo_width()
                if canvas_width > 1:
                    canvas.itemconfig(canvas_window, width=canvas_width)
            except tk.TclError:
                pass
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # CanvasとScrollbarを配置
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # マウスホイールでスクロール（Canvasにフォーカスがある時のみ）
        # bind_all はウィンドウ閉鎖後も残るため、破棄済み canvas 参照で TclError が出ないよう try で保護
        def _on_mousewheel(event):
            try:
                if not canvas.winfo_exists():
                    return
                if canvas.winfo_containing(event.x_root, event.y_root):
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except tk.TclError:
                pass  # ウィンドウ破棄後は何もしない
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # メインフレーム（スクロール可能なフレーム内）
        main_frame = ttk.Frame(scrollable_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========== イベント追加ボタン ==========
        if self.rule_engine:
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.add_event_button = ttk.Button(
                button_frame,
                text="イベント追加",
                command=self._on_add_event,
                width=15
            )
            # 個体切り替えボタンの左側に配置（2個分左に移動）
            self.add_event_button.pack(side=tk.RIGHT, padx=(5, 150))
        
        # ========== 上：基本情報 ==========
        basic_frame = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        basic_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.basic_labels = {}
        basic_fields = [
            ('cow_id', '拡大4桁ID'),
            ('jpn10', '個体識別番号'),
            ('brd', '品種'),
            ('bthd', '生年月日'),
            ('clvd', '分娩月日'),
            ('lact', '産次'),
            ('rc', '繁殖区分'),
            ('pen', '群'),
            ('dim', 'DIM')
        ]
        
        # フォントを Meiryo UI に統一（サイズはアプリ設定に従う）
        font_family = "Meiryo UI"
        font_size = self.app_settings.get_font_size()
        base_font = (font_family, font_size)
        
        for i, (key, label) in enumerate(basic_fields):
            row = i // 2
            col = (i % 2) * 2
            
            ttk.Label(basic_frame, text=f"{label}:", font=base_font).grid(
                row=row, column=col, sticky=tk.W, padx=5, pady=2
            )
            value_label = ttk.Label(basic_frame, text="", foreground="blue", font=base_font)
            # 品種・群は空のときもダブルクリック/右クリックで編集できるよう最小幅を確保
            if key in ("brd", "pen"):
                value_label.config(width=16)
            value_label.grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)
            self.basic_labels[key] = value_label
            if key == "cow_id":
                value_label.bind("<Double-Button-1>", self._on_edit_cow_id)
            elif key == "brd":
                value_label.bind("<Double-Button-1>", self._on_edit_brd)
                value_label.bind("<Button-3>", lambda e, k=key: self._on_basic_right_click(k, e))
            elif key == "lact":
                value_label.bind("<Double-Button-1>", self._on_edit_lact)
            elif key == "pen":
                value_label.bind("<Double-Button-1>", self._on_edit_pen)
                value_label.bind("<Button-3>", lambda e, k=key: self._on_basic_right_click(k, e))
        
        # ========== 中：計算項目（タブ構造） ==========
        # 下部のイベント履歴を廃止したため、空いたスペースを項目タブでいっぱいに使う
        calc_frame = ttk.LabelFrame(main_frame, text="項目", padding=5)
        calc_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Notebook（タブ構造）を作成（縦方向にも拡張）
        self.calc_notebook = ttk.Notebook(calc_frame)
        self.calc_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 各タブのフレームを保持する辞書
        self.tab_frames: Dict[str, tk.Frame] = {}
        self.tab_canvases: Dict[str, tk.Canvas] = {}
        self.tab_scrollbars: Dict[str, ttk.Scrollbar] = {}
        self.tab_item_frames: Dict[str, tk.Frame] = {}
        self.tab_item_widgets: Dict[str, List[Dict[str, Any]]] = {}
        
        # タブ名称のマッピング（内部名 -> 表示名）
        self.tab_display_names = {
            "USER": "ユーザー ",
            "REPRODUCTION": "繁殖 ",
            "DHI": "乳検 ",
            "GENOMIC": "ゲノム ",
            "HEALTH": "疾病 ",
            "OTHERS": "その他 "
        }
        
        # タブスタイルを設定（Meiryo UI 統一）
        font_family = "Meiryo UI"
        font_size = self.app_settings.get_font_size()
        base_font = (font_family, font_size)
        
        style = ttk.Style()
        # 選択されたタブ：太字、背景色を変更
        style.configure("Selected.TNotebook.Tab", 
                       font=(font_family, font_size, 'bold'), 
                       padding=[10, 5],
                       background='#E0E0E0')
        # 選択されていないタブ：通常のフォント
        style.configure("TNotebook.Tab", 
                       font=base_font, 
                       padding=[10, 5])
        # タブ全体のスタイル
        style.configure("TNotebook", tabmargins=[2, 5, 2, 0])
        
        # タブを作成
        tab_names = ["USER", "REPRODUCTION", "DHI", "GENOMIC", "HEALTH", "OTHERS"]
        for tab_name in tab_names:
            self._create_tab(tab_name)
        
        # タブ選択時のイベントをバインド
        self.calc_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # 初期表示
        self.refresh_calculated_items({})
        
        # 初期選択タブのスタイルを設定
        self._on_tab_changed(None)
    
    def _create_tab(self, tab_name: str):
        """タブを作成"""
        # タブ用のフレーム
        tab_frame = ttk.Frame(self.calc_notebook)
        # 日本語名を取得（後ろにスペースが含まれている）
        display_name = self.tab_display_names.get(tab_name, tab_name + " ")
        self.calc_notebook.add(tab_frame, text=display_name)
        self.tab_frames[tab_name] = tab_frame
        
        # ヘッダーフレーム（並び替えボタン用）
        header_frame = ttk.Frame(tab_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # 並び替えボタンをヘッダーに配置
        sort_button = ttk.Button(
            header_frame,
            text="並び替え",
            command=lambda: self._open_sort_window(tab_name),
            width=10
        )
        sort_button.pack(side=tk.RIGHT)
        
        # Canvas + Scrollbar 構造（横スクロール・縦スクロール対応）
        canvas_frame = ttk.Frame(tab_frame)
        # 空いた縦方向のスペースも使用する
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas（高さはフレームのリサイズに合わせて動的に設定し、ウィンドウ下いっぱいまで拡張）
        canvas = tk.Canvas(canvas_frame, highlightthickness=0, height=300)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        # スクロール可能なフレーム
        scrollable_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # スクロールバーとCanvasの配置
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        # 縦方向にも拡張可能にする
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        # Canvasサイズ変更時の処理
        def configure_scroll_region(event):
            # スクロール領域を更新（内容のサイズに合わせる）
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 横スクロールを機能させるため、scrollable_frameの幅を内容に合わせる
            # ただし、Canvasの幅より小さい場合はCanvasの幅に合わせる（縦スクロール時の見た目を維持）
            req_width = scrollable_frame.winfo_reqwidth()
            canvas_width = canvas.winfo_width()
            if req_width > canvas_width:
                canvas.itemconfig(canvas_window, width=req_width)
            else:
                canvas.itemconfig(canvas_window, width=canvas_width)
        
        def on_canvas_configure(event):
            # Canvasのサイズ変更時もスクロール領域を更新
            canvas.configure(scrollregion=canvas.bbox("all"))
            req_width = scrollable_frame.winfo_reqwidth()
            canvas_width = event.width
            if req_width > canvas_width:
                canvas.itemconfig(canvas_window, width=req_width)
            else:
                canvas.itemconfig(canvas_window, width=canvas_width)
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # フレームのリサイズに合わせてCanvasの高さをウィンドウ下いっぱいまで拡張
        def on_canvas_frame_configure(event):
            try:
                if not canvas.winfo_exists():
                    return
                # 横スクロールバー分を除いた高さをCanvasに設定
                h = max(80, event.height - 22)
                canvas.configure(height=h)
            except tk.TclError:
                pass
        
        def _safe_resize_canvas_height():
            try:
                if canvas.winfo_exists() and canvas_frame.winfo_exists():
                    h = max(80, canvas_frame.winfo_height() - 22)
                    canvas.configure(height=h)
            except tk.TclError:
                pass
        
        canvas_frame.bind("<Configure>", on_canvas_frame_configure)
        # 初回表示時もフレーム確定後に高さを反映
        self.frame.after(150, lambda: _safe_resize_canvas_height())
        
        # マウスホイールでスクロール（Windows/Linux）
        # ウィンドウ破棄後も bind_all が残ることがあるため TclError を捕捉
        def on_mousewheel(event):
            try:
                if not canvas.winfo_exists():
                    return
                # Shiftキーが押されている場合は横スクロール、そうでなければ縦スクロール
                if event.state & 0x1:  # Shiftキーが押されている
                    canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                pass

        # Canvasにフォーカスがある時のみスクロール
        def on_enter(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)

        def on_leave(event):
            try:
                if canvas.winfo_exists():
                    canvas.unbind_all("<MouseWheel>")
            except tk.TclError:
                pass
        
        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        
        self.tab_canvases[tab_name] = canvas
        self.tab_scrollbars[tab_name] = h_scrollbar
        self.tab_item_frames[tab_name] = scrollable_frame
        self.tab_item_widgets[tab_name] = []
    
    def _on_tab_changed(self, event):
        """タブが変更されたときの処理（選択されたタブを視覚的に区別）"""
        try:
            # 現在選択されているタブのインデックスを取得
            selected_index = self.calc_notebook.index(self.calc_notebook.select())
            
            # すべてのタブのスタイルをリセット
            tab_names = ["USER", "REPRODUCTION", "DHI", "GENOMIC", "HEALTH", "OTHERS"]
            for idx, tab_name in enumerate(tab_names):
                if idx == selected_index:
                    # 選択されたタブには太字スタイルを適用
                    self.calc_notebook.tab(idx, style="Selected.TNotebook.Tab")
                else:
                    # 選択されていないタブには通常スタイルを適用
                    self.calc_notebook.tab(idx, style="TNotebook.Tab")
        except Exception as e:
            # エラーが発生しても処理を続行
            pass
    
    def load_cow(self, cow_auto_id: int):
        """
        牛の情報を読み込んで表示
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            self.cow_auto_id = cow_auto_id
            
            # cow データを取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                print(f"WARNING: cow が見つかりません: auto_id={cow_auto_id}")
                return
            
            # 授精設定辞書を必ず読み込む（毎回再読み込み）
            # これにより、異なる牛を表示した場合でも正しい設定が読み込まれる
            self._init_settings_manager(cow_auto_id)
            
            # 基本情報を表示
            self._display_basic_info(cow)
            
            # 計算項目を表示（FormulaEngine を1回だけ呼ぶ）
            try:
                calculated = self.formula_engine.calculate(cow_auto_id)
                self.refresh_calculated_items(calculated)
            except Exception as e:
                print(f"WARNING: FormulaEngine.calculate でエラー: {e}")
                import traceback
                traceback.print_exc()
            
            # 最終日付を計算してコールバックで通知
            self._calculate_and_notify_latest_dates(cow_auto_id)
            
        except Exception as e:
            import traceback
            print(f"ERROR: load_cow で例外が発生しました: {e}")
            traceback.print_exc()
            raise
    
    def _display_basic_info(self, cow: Dict[str, Any]):
        """基本情報を表示"""
        # cow_id
        self.basic_labels['cow_id'].config(text=cow.get('cow_id', ''))
        
        # jpn10（個体識別番号）- 表示のみ
        jpn10 = cow.get('jpn10', '')
        self.basic_labels['jpn10'].config(text=jpn10)
        
        # brd（品種）
        self.basic_labels['brd'].config(text=cow.get('brd', ''))
        
        # bthd（生年月日）
        self.basic_labels['bthd'].config(text=cow.get('bthd', ''))

        # clvd（分娩月日）
        self.basic_labels['clvd'].config(text=cow.get('clvd', ''))
        
        # lact（産次）
        lact = cow.get('lact')
        self.basic_labels['lact'].config(text=str(lact) if lact is not None else '')
        
        # rc（繁殖コード）：コード番号：日本語のみ表示
        rc = cow.get('rc')
        rc_text = self.rc_names.get(rc, f'{rc}：') if rc is not None else ''
        self.basic_labels['rc'].config(text=rc_text)
        
        # pen（群）
        pen_value = cow.get('pen', '')
        if pen_value:
            # 数値の場合は文字列に変換
            pen_code = str(pen_value).strip()
            if pen_code in self.pen_settings:
                pen_value = self.pen_settings.get(pen_code, pen_code)
            else:
                pen_value = pen_code
        else:
            pen_value = ''
        self.basic_labels['pen'].config(text=pen_value)
        
        # dim（分娩後日数）- 計算項目なので、後でrefresh_calculated_itemsで更新される

    def _on_edit_cow_id(self, _event=None):
        """拡大4桁ID（cow_id）を編集"""
        if not self.cow_auto_id:
            return
        
        cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
        if not cow:
            messagebox.showerror("エラー", "個体情報が取得できませんでした。")
            return
        
        current_id = cow.get("cow_id", "")
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("拡大4桁IDを修正")
        dialog.geometry("320x160")
        
        ttk.Label(dialog, text="拡大4桁ID:", font=("Meiryo UI", 10)).pack(pady=(15, 5))
        entry = ttk.Entry(dialog, width=20)
        entry.pack()
        entry.insert(0, current_id)
        entry.focus()
        
        def on_ok():
            new_id = entry.get().strip()
            if not new_id:
                messagebox.showwarning("警告", "拡大4桁IDを入力してください。")
                return
            if not new_id.isdigit():
                messagebox.showwarning("警告", "拡大4桁IDは数字のみ入力してください。")
                return
            if len(new_id) > 4:
                messagebox.showwarning("警告", "拡大4桁IDは4桁以内で入力してください。")
                return
            
            padded_id = new_id.zfill(4)
            if padded_id == current_id:
                dialog.destroy()
                return
            
            # 重複チェック（別個体が同じIDを使用している場合は不可）
            existing = self.db.get_cows_by_id(padded_id)
            for ex in existing:
                if ex.get("auto_id") != self.cow_auto_id:
                    messagebox.showerror(
                        "エラー",
                        f"拡大4桁ID「{padded_id}」は他の個体で使用されています。"
                    )
                    return
            
            try:
                self.db.update_cow(self.cow_auto_id, {"cow_id": padded_id})
                self.basic_labels['cow_id'].config(text=padded_id)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"拡大4桁IDの更新に失敗しました: {e}")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def refresh_calculated_items(self, calculated: Optional[Dict[str, Any]] = None):
        """
        計算項目表示を再描画（ウィンドウ再生成なし）
        """
        try:
            if calculated is None:
                if self.cow_auto_id is None:
                    calculated = {}
                else:
                    calculated = self.formula_engine.calculate(self.cow_auto_id)
            
            # DIMを基本情報欄に表示
            dim_value = calculated.get('DIM')
            if 'dim' in self.basic_labels:
                if dim_value is not None:
                    self.basic_labels['dim'].config(text=str(dim_value))
                else:
                    self.basic_labels['dim'].config(text='')
            
            self._render_calculated_items(calculated or {})
        except Exception as e:
            import traceback
            logging.error(f"ERROR: refresh_calculated_items で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _render_calculated_items(self, calculated: Dict[str, Any]):
        """計算項目エリアのUIを再構築（タブ構造）"""
        # 各タブのウィジェットをクリア
        for tab_name in self.tab_item_widgets:
            # 既存のウィジェットをすべて削除
            for widget_data in self.tab_item_widgets[tab_name]:
                try:
                    frame = widget_data.get("frame")
                    if frame is not None:
                        frame.destroy()
                except tk.TclError:
                    pass
                # name_labelとvalue_labelも削除
                name_label = widget_data.get("name_label")
                if name_label is not None:
                    try:
                        name_label.destroy()
                    except tk.TclError:
                        pass
                value_label = widget_data.get("value_label")
                if value_label is not None:
                    try:
                        value_label.destroy()
                    except tk.TclError:
                        pass
            self.tab_item_widgets[tab_name] = []
            
            # フレーム内のすべてのウィジェットを明示的に削除（残像を防ぐ）
            if tab_name in self.tab_item_frames:
                item_frame = self.tab_item_frames[tab_name]
                try:
                    # フレーム内のすべての子ウィジェットを削除
                    for widget in item_frame.winfo_children():
                        try:
                            widget.destroy()
                        except tk.TclError:
                            pass
                except tk.TclError:
                    pass
        
        # item_dictionaryを取得
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        
        # 生まれた年・生まれた年月を cow の bthd から補完（辞書に無くてもその他タブに表示するため）
        if self.cow_auto_id:
            cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
            if cow:
                bthd = (cow.get("bthd") or "").strip()
                if bthd:
                    try:
                        from datetime import datetime
                        dt = datetime.strptime(bthd, "%Y-%m-%d")
                        if "BTHYR" not in calculated:
                            calculated["BTHYR"] = dt.year
                        if "BTHYM" not in calculated:
                            calculated["BTHYM"] = dt.strftime("%Y-%m")
                    except (ValueError, TypeError):
                        pass
        
        # 各タブに表示する項目を分類
        tab_items: Dict[str, List[str]] = {
            "USER": [],
            "REPRODUCTION": [],
            "DHI": [],
            "GENOMIC": [],
            "HEALTH": [],
            "OTHERS": []
        }
        
        # USERタブ：ユーザーが選択した項目（並び順を保持）
        user_order = self.tab_orders.get("USER", [])
        for item_code in user_order:
            if item_code in calculated or item_code in item_dict:
                tab_items["USER"].append(item_code)
        
        # USERタブの末尾に空行を追加（項目追加用）
        tab_items["USER"].append(None)
        
        # その他のタブ：categoryに基づいて分類（保存された並び順を使用）
        # GENOME は項目辞書で使われるカテゴリ名、GENOMIC はタブ名（両方ともゲノムタブへ）
        category_tabs = {
            "REPRODUCTION": "REPRODUCTION",
            "DHI": "DHI",
            "GENOMIC": "GENOMIC",
            "GENOME": "GENOMIC",
            "HEALTH": "HEALTH"
        }
        
        # まず、categoryに基づいて分類
        categorized_items: Dict[str, List[str]] = {
            "REPRODUCTION": [],
            "DHI": [],
            "GENOMIC": [],
            "HEALTH": [],
            "OTHERS": []
        }
        
        for item_code, value in calculated.items():
            # USERタブに含まれていても、カテゴリータブにも表示する
            item_def = item_dict.get(item_code, {})
            category = item_def.get("category", "").upper()
            
            if category in category_tabs:
                categorized_items[category_tabs[category]].append(item_code)
            else:
                categorized_items["OTHERS"].append(item_code)
        
        # その他タブに必ず含める項目（生まれた年・生まれた年月）
        for code in ("BTHYR", "BTHYM"):
            if code not in categorized_items["OTHERS"]:
                categorized_items["OTHERS"].append(code)
        
        # item_dictionaryに存在するがcalculatedに含まれていない項目をカテゴリータブに追加
        # （custom項目、およびゲノムタブ用に GENOME カテゴリの source 項目も表示して並び替え可能にする）
        for item_code, item_def in item_dict.items():
            origin = item_def.get("origin", "").lower()
            category = item_def.get("category", "").upper()
            in_any_tab = any(item_code in tab_items_list for tab_items_list in tab_items.values())
            if in_any_tab:
                continue
            # custom項目で、まだどのタブにも追加されていない場合
            if origin == "custom" and item_code not in calculated:
                if category in category_tabs:
                    categorized_items[category_tabs[category]].append(item_code)
                else:
                    categorized_items["OTHERS"].append(item_code)
            # GENOME/GENOMIC カテゴリの source 項目はゲノムタブに表示（値はcalculatedにあれば表示、なければ空）
            elif category in ("GENOME", "GENOMIC") and origin == "source":
                if item_code not in categorized_items["GENOMIC"]:
                    categorized_items["GENOMIC"].append(item_code)
        
        # 保存された並び順を適用
        for tab_name in ["REPRODUCTION", "DHI", "GENOMIC", "HEALTH", "OTHERS"]:
            saved_order = self.tab_orders.get(tab_name, [])
            # 保存された順序で並べ替え（存在する項目のみ）
            ordered_items = []
            for code in saved_order:
                if code in categorized_items[tab_name]:
                    ordered_items.append(code)
            # 保存されていない新しい項目を末尾に追加
            for code in categorized_items[tab_name]:
                if code not in ordered_items:
                    ordered_items.append(code)
            tab_items[tab_name] = ordered_items
        
        # 各タブを描画
        for tab_name, item_codes in tab_items.items():
            self._render_tab_items(tab_name, item_codes, calculated, item_dict)
    
    def _render_tab_items(self, tab_name: str, item_codes: List[str], 
                          calculated: Dict[str, Any], item_dict: Dict[str, Any]):
        """タブ内の項目を描画（基本情報と同じ「項目：　数値」形式）"""
        item_frame = self.tab_item_frames[tab_name]
        widgets = []
        
        # 2列グリッドで配置（横に収まりやすく、読みやすい）
        items_per_row = 2
        
        for idx, item_code in enumerate(item_codes):
            row = idx // items_per_row
            col = (idx % items_per_row) * 2
            
            # 空行の場合（USERタブのみ）
            if item_code is None and tab_name == "USER":
                # 空行の表示（基本情報と同じ形式）
                empty_label = ttk.Label(item_frame, text="（空行）", 
                                       font=("Meiryo UI", 10), foreground="gray")
                empty_label.grid(row=row, column=col, columnspan=2, sticky=tk.W, padx=5, pady=2)
                
                # 右クリックで項目追加
                empty_label.bind(
                    '<Button-3>',
                    lambda e, idx=idx, code=None, tab=tab_name: self._on_calc_item_right_click(e, idx, code, tab)
                )
                empty_label.bind('<Enter>', lambda e, w=empty_label: w.config(cursor='hand2'))
                empty_label.bind('<Leave>', lambda e, w=empty_label: w.config(cursor=''))
                
                widgets.append({
                    "frame": None,
                    "name_label": empty_label,
                    "value_label": None,
                    "item_code": None,
                })
                continue
            
            # 通常の項目表示（基本情報と同じ形式：項目：　数値）
            name = self._get_item_display_name(item_code)
            
            # フォント（Meiryo UI 統一）
            font_family = "Meiryo UI"
            font_size = self.app_settings.get_font_size()
            base_font = (font_family, font_size)
            
            # 項目名ラベル（基本情報と同じスタイル）
            name_label = ttk.Label(item_frame, text=f"{name}:", font=base_font)
            name_label.grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            
            # 計算値を取得（item_codeが存在する場合は、calculatedから取得、なければNone）
            raw_value = calculated.get(item_code) if item_code and item_code in calculated else None
            
            # 値の表示（None/空の場合は「–」）
            if raw_value is None or raw_value == "":
                value_text = "–"
            else:
                value_text = self._format_calc_value(item_code, raw_value)
            
            # 値ラベル（基本情報と同じスタイル、青色）
            value_label = ttk.Label(item_frame, text=value_text, 
                                   foreground="blue", font=base_font)
            value_label.grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)
            
            # editableな項目の場合はダブルクリックで編集可能
            item_def = item_dict.get(item_code, {}) if item_code else {}
            if item_code and item_def.get("editable", False):
                for widget in (name_label, value_label):
                    widget.bind(
                        '<Double-Button-1>',
                        lambda e, code=item_code: self._on_edit_calc_item(code)
                    )
                    widget.bind('<Enter>', lambda e, w=widget: w.config(cursor='hand2'))
                    widget.bind('<Leave>', lambda e, w=widget: w.config(cursor=''))
            # PENはダブルクリックで変更可能
            if item_code == "PEN":
                for widget in (name_label, value_label):
                    widget.bind(
                        '<Double-Button-1>',
                        lambda e: self._on_edit_pen()
                    )
                    widget.bind('<Enter>', lambda e, w=widget: w.config(cursor='hand2'))
                    widget.bind('<Leave>', lambda e, w=widget: w.config(cursor=''))
            
            # USERタブのみ右クリックメニュー
            if tab_name == "USER":
                for widget in (name_label, value_label):
                    widget.bind(
                        '<Button-3>',
                        lambda e, idx=idx, code=item_code, tab=tab_name: self._on_calc_item_right_click(e, idx, code, tab)
                    )
            
            widgets.append({
                "frame": None,
                "name_label": name_label,
                "value_label": value_label,
                "item_code": item_code,
            })
        
        # グリッドの列の重みを設定（2列対応）
        for col_idx in range(0, 4, 2):  # 0, 2（2列分）
            item_frame.grid_columnconfigure(col_idx, weight=0)      # 項目名列
            item_frame.grid_columnconfigure(col_idx + 1, weight=1)   # 値列
        
        self.tab_item_widgets[tab_name] = widgets
    
    def _open_sort_window(self, tab_name: str):
        """並び替えウィンドウを開く"""
        try:
            # 現在表示されている項目を取得（保存された並び順を優先）
            saved_order = self.tab_orders.get(tab_name, []).copy()
            current_widgets = self.tab_item_widgets.get(tab_name, [])
            
            # 現在表示されている項目コードのリストを作成
            current_items = []
            for widget_data in current_widgets:
                item_code = widget_data.get("item_code")
                if item_code is not None:  # 空行（None）を除外
                    current_items.append(item_code)
            
            # 保存された並び順を優先し、新しい項目を末尾に追加
            display_order = []
            if saved_order:
                # USERタブの場合は空行を除外
                if tab_name == "USER":
                    saved_items = [code for code in saved_order if code is not None]
                else:
                    saved_items = saved_order
                
                # 保存された順序で存在する項目を追加
                for code in saved_items:
                    if code in current_items:
                        display_order.append(code)
                
                # 保存されていない新しい項目を末尾に追加
                for code in current_items:
                    if code not in display_order:
                        display_order.append(code)
            else:
                # 保存された順序がない場合は、現在の表示順序を使用
                display_order = current_items.copy()
            
            # それでも項目がない場合は、calculatedから直接取得を試みる
            if not display_order and self.cow_auto_id:
                calculated = self.formula_engine.calculate(self.cow_auto_id)
                item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
                
                # カテゴリに基づいて分類
                category_tabs = {
                    "REPRODUCTION": "REPRODUCTION",
                    "DHI": "DHI",
                    "GENOMIC": "GENOMIC",
                    "HEALTH": "HEALTH"
                }
                
                if tab_name == "USER":
                    # USERタブの場合は保存された順序を使用
                    user_order = self.tab_orders.get("USER", [])
                    for code in user_order:
                        if code is not None and (code in calculated or code in item_dict):
                            display_order.append(code)
                elif tab_name in category_tabs:
                    # カテゴリータブの場合はcategoryで分類（ゲノムは GENOME / GENOMIC の両方）
                    category = category_tabs[tab_name]
                    for item_code, value in calculated.items():
                        item_def = item_dict.get(item_code, {})
                        item_category = item_def.get("category", "").upper()
                        if item_category == category or (tab_name == "GENOMIC" and item_category == "GENOME"):
                            display_order.append(item_code)
                    # ゲノムタブ: item_dict の GENOME/GENOMIC 項目でまだ追加されていないもの
                    if tab_name == "GENOMIC":
                        for item_code, item_def in item_dict.items():
                            if item_code in display_order:
                                continue
                            oc = item_def.get("origin", "").lower()
                            cat = item_def.get("category", "").upper()
                            if cat in ("GENOME", "GENOMIC") and oc == "source":
                                display_order.append(item_code)
                else:
                    # OTHERSタブ
                    for item_code, value in calculated.items():
                        item_def = item_dict.get(item_code, {})
                        category = item_def.get("category", "").upper()
                        if category not in category_tabs.values():
                            display_order.append(item_code)
            
            if not display_order:
                messagebox.showinfo("情報", "並び替え可能な項目がありません。")
                return
            
            # 並び替えウィンドウを作成
            dialog = tk.Toplevel(self.frame)
            dialog.title(f"{tab_name}タブ - 項目の並び替え")
            dialog.geometry("400x500")
            
            # 説明ラベル
            ttk.Label(dialog, text="項目の順序を変更してください", font=("Meiryo UI", 10)).pack(pady=10)
            
            # リストボックスとスクロールバー
            list_frame = ttk.Frame(dialog)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            scrollbar = ttk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            listbox = tk.Listbox(list_frame, height=15, yscrollcommand=scrollbar.set)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=listbox.yview)
            
            # 項目名のリストを作成
            item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
            display_items = []
            for code in display_order:
                name = self._get_item_display_name(code)
                display_items.append((code, name))
                listbox.insert(tk.END, name)
            
            # ボタンフレーム
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            # 上へ移動ボタン
            def move_up():
                selection = listbox.curselection()
                if not selection or selection[0] == 0:
                    return
                idx = selection[0]
                # リストボックス内の項目を入れ替え
                item = listbox.get(idx)
                listbox.delete(idx)
                listbox.insert(idx - 1, item)
                listbox.selection_set(idx - 1)
                # 内部リストも入れ替え
                display_items[idx], display_items[idx - 1] = display_items[idx - 1], display_items[idx]
            
            # 下へ移動ボタン
            def move_down():
                selection = listbox.curselection()
                if not selection or selection[0] == len(display_items) - 1:
                    return
                idx = selection[0]
                # リストボックス内の項目を入れ替え
                item = listbox.get(idx)
                listbox.delete(idx)
                listbox.insert(idx + 1, item)
                listbox.selection_set(idx + 1)
                # 内部リストも入れ替え
                display_items[idx], display_items[idx + 1] = display_items[idx + 1], display_items[idx]
            
            ttk.Button(button_frame, text="↑ 上へ", command=move_up, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="↓ 下へ", command=move_down, width=10).pack(side=tk.LEFT, padx=5)
            
            # OK/キャンセルボタン
            def on_ok():
                # 新しい並び順を取得
                new_order = [item[0] for item in display_items]
                
                # USERタブの場合は空行を追加
                if tab_name == "USER":
                    new_order.append(None)
                
                # 並び順を保存
                self.tab_orders[tab_name] = new_order
                self._save_tab_orders()
                
                # 再描画
                if self.cow_auto_id:
                    calculated = self.formula_engine.calculate(self.cow_auto_id)
                    self.refresh_calculated_items(calculated)
                
                dialog.destroy()
            
            def on_cancel():
                dialog.destroy()
            
            ok_cancel_frame = ttk.Frame(dialog)
            ok_cancel_frame.pack(pady=10)
            ttk.Button(ok_cancel_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(ok_cancel_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
            
            # 初期選択
            if display_items:
                listbox.selection_set(0)
                listbox.focus_set()
            
            # ウィンドウを中央に配置
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
            y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
            dialog.geometry(f"+{x}+{y}")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _open_sort_window で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"並び替えウィンドウの表示に失敗しました: {e}")
    
    def _on_calc_item_right_click(self, event, index: int, item_code: Optional[str], tab_name: str = "USER"):
        """計算項目行の右クリックメニュー（USERタブのみ）"""
        if tab_name != "USER":
            return  # USERタブ以外では右クリックメニューを表示しない
        
        try:
            menu = tk.Menu(self.frame, tearoff=0)
            if item_code:
                # 既存項目の場合
                menu.add_command(label="項目を変更", command=lambda: self._replace_calc_item(index))
                menu.add_command(label="項目を削除", command=lambda: self._remove_calc_item(index, item_code))
                menu.add_separator()
                menu.add_command(label="項目を追加", command=lambda: self._add_calc_item_after(index))
            else:
                # 空行の場合
                menu.add_command(label="項目を追加", command=lambda: self._add_calc_item(index))
            
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_calc_item_right_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _add_calc_item(self, index: int):
        """USERタブに項目を追加（指定位置に挿入）"""
        selected = self._select_item_from_dictionary()
        if not selected:
            return
        user_order = self.tab_orders.get("USER", [])
        # 空行（None）を除外したリストでチェック
        user_order_without_none = [code for code in user_order if code is not None]
        if selected in user_order_without_none:
            messagebox.showinfo("情報", "同じ項目がすでに追加されています。")
            return
        
        # 空行の位置を考慮して挿入
        # 空行を除いた実際のインデックスを計算
        actual_index = 0
        for i, code in enumerate(user_order):
            if code is None:
                if i == index:
                    # 空行の位置に挿入
                    user_order.insert(index, selected)
                    break
            else:
                if actual_index == index:
                    # この位置に挿入
                    user_order.insert(i, selected)
                    break
                actual_index += 1
        else:
            # 末尾に追加
            # 最後のNoneを削除してから追加
            if user_order and user_order[-1] is None:
                user_order.pop()
            user_order.append(selected)
            user_order.append(None)  # 空行を再追加
        
        self.tab_orders["USER"] = user_order
        self._save_tab_orders()
        logging.info(f"CowCard USER tab item added: {selected} at index {index}")
        # 再描画
        if self.cow_auto_id:
            calculated = self.formula_engine.calculate(self.cow_auto_id)
            self.refresh_calculated_items(calculated)
    
    def _add_calc_item_after(self, index: int):
        """USERタブに項目を追加（指定項目の後に挿入）"""
        selected = self._select_item_from_dictionary()
        if not selected:
            return
        user_order = self.tab_orders.get("USER", [])
        # 空行（None）を除外したリストでチェック
        user_order_without_none = [code for code in user_order if code is not None]
        if selected in user_order_without_none:
            messagebox.showinfo("情報", "同じ項目がすでに追加されています。")
            return
        
        # 指定位置の後に挿入
        insert_index = index + 1
        # 空行を考慮して挿入位置を調整
        if insert_index < len(user_order):
            user_order.insert(insert_index, selected)
        else:
            # 末尾に追加（空行を削除してから追加）
            if user_order and user_order[-1] is None:
                user_order.pop()
            user_order.append(selected)
            user_order.append(None)  # 空行を再追加
        
        self.tab_orders["USER"] = user_order
        self._save_tab_orders()
        logging.info(f"CowCard USER tab item added after index {index}: {selected}")
        # 再描画
        if self.cow_auto_id:
            calculated = self.formula_engine.calculate(self.cow_auto_id)
            self.refresh_calculated_items(calculated)
    
    def _replace_calc_item(self, index: int):
        """USERタブの既存項目を差し替え"""
        user_order = self.tab_orders.get("USER", [])
        if index < 0 or index >= len(user_order):
            return
        current = user_order[index]
        if current is None:
            # 空行の場合は追加処理に回す
            self._add_calc_item(index)
            return
        
        selected = self._select_item_from_dictionary(initial_code=current)
        if not selected or selected == current:
            return
        
        # 空行を除外したリストでチェック
        user_order_without_none = [code for code in user_order if code is not None]
        if selected in user_order_without_none:
            messagebox.showinfo("情報", "同じ項目がすでに追加されています。")
            return
        
        user_order[index] = selected
        self.tab_orders["USER"] = user_order
        self._save_tab_orders()
        logging.info(f"CowCard USER tab item replaced: {current} -> {selected}")
        # 再描画
        if self.cow_auto_id:
            calculated = self.formula_engine.calculate(self.cow_auto_id)
            self.refresh_calculated_items(calculated)
    
    def _remove_calc_item(self, index: int, item_code: str):
        """USERタブから項目を削除"""
        user_order = self.tab_orders.get("USER", [])
        if index < 0 or index >= len(user_order):
            return
        if user_order[index] != item_code:
            return
        
        # 項目を削除
        user_order.remove(item_code)
        
        # 空行が存在しない場合は末尾に追加
        if None not in user_order:
            user_order.append(None)
        
        self.tab_orders["USER"] = user_order
        self._save_tab_orders()
        logging.info(f"CowCard USER tab item removed: {item_code}")
        # 再描画
        if self.cow_auto_id:
            calculated = self.formula_engine.calculate(self.cow_auto_id)
            self.refresh_calculated_items(calculated)
    
    def _on_edit_calc_item(self, item_code: str):
        """editableな計算項目を編集"""
        if not self.cow_auto_id:
            messagebox.showwarning("警告", "対象牛が設定されていません")
            return
        
        # 現在の値を取得
        calculated = self.formula_engine.calculate(self.cow_auto_id)
        current_value = calculated.get(item_code)
        
        # 項目定義を取得
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        item_def = item_dict.get(item_code, {})
        display_name = item_def.get("display_name") or item_def.get("label") or item_code
        data_type = item_def.get("data_type", "str")
        
        # 編集ダイアログを表示
        dialog = tk.Toplevel(self.frame)
        dialog.title(f"{display_name}を編集")
        dialog.geometry("300x150")
        
        ttk.Label(dialog, text=f"{display_name}:", font=("Meiryo UI", 10)).pack(pady=10)
        
        # 入力フィールド
        entry = ttk.Entry(dialog, width=20)
        entry.pack(pady=5)
        
        # 現在の値を設定
        if current_value is not None:
            entry.insert(0, str(current_value))
        entry.select_range(0, tk.END)
        entry.focus()
        
        def on_ok():
            try:
                value_str = entry.get().strip()
                
                # データ型に応じて変換
                if data_type == "int":
                    value = int(value_str) if value_str else None
                elif data_type == "float":
                    value = float(value_str) if value_str else None
                else:
                    value = value_str if value_str else None
                
                # item_valueテーブルに保存
                if value is not None:
                    self.db.set_item_value(self.cow_auto_id, item_code, value)
                    logging.info(f"CowCard editable calc item updated: {item_code}={value}")
                    # 授精回数（BRED）を変更した場合、直近産次のAI/ETイベントの授精回数も連動して更新
                    if item_code == "BRED" and isinstance(value, int) and value >= 1 and self.rule_engine:
                        try:
                            self.rule_engine.apply_insemination_count_from_item(self.cow_auto_id, value)
                            if self.cow_auto_id:
                                self._display_events(self.cow_auto_id)
                        except Exception as e:
                            logging.warning(f"CowCard BRED連動: AIイベント授精回数の更新に失敗: {e}")
                else:
                    # Noneの場合は削除（item_valueテーブルから削除する必要があるが、現在のDBHandlerには削除メソッドがない）
                    # とりあえず空文字列を保存
                    self.db.set_item_value(self.cow_auto_id, item_code, "")
                    logging.info(f"CowCard editable calc item cleared: {item_code}")
                
                # 表示を更新
                self.refresh_calculated_items()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("エラー", f"無効な値です。{data_type}型の値を入力してください。")
        
        def on_cancel():
            dialog.destroy()
        
        # EnterキーでOK
        entry.bind('<Return>', lambda e: on_ok())
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _select_item_from_dictionary(self, initial_code: Optional[str] = None) -> Optional[str]:
        """Item Dictionary から計算項目を選択（カテゴリー別表示・検索機能付き）"""
        selectable = self._get_selectable_items_with_category()
        if not selectable:
            messagebox.showerror("エラー", "選択可能な項目がありません。item_dictionary.json を確認してください。")
            return None
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("項目を選択")
        dialog.geometry("500x600")
        
        # カテゴリー名のマッピング
        category_names = {
            "REPRODUCTION": "繁殖",
            "DHI": "乳検",
            "GENOMIC": "ゲノム",
            "HEALTH": "疾病",
            "OTHERS": "その他",
            "USER": "ユーザー"
        }
        
        # メインフレーム
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 検索フレーム
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 左側にスペーサーを追加して検索欄を右側に配置
        ttk.Label(search_frame, text="").pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 検索欄を右側に配置するためのフレーム
        search_right_frame = ttk.Frame(search_frame)
        search_right_frame.pack(side=tk.RIGHT)
        
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_right_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.RIGHT, padx=(0, 5))
        ttk.Label(search_right_frame, text="検索:", font=("Meiryo UI", 10)).pack(side=tk.RIGHT, padx=(5, 0))
        
        # クリアボタン（後でclear_search関数のコマンドを設定）
        clear_button = ttk.Button(search_right_frame, text="クリア", width=8)
        clear_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Treeview（カテゴリー別表示）
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        tree = ttk.Treeview(tree_frame, columns=('code',), show='tree headings', height=20, yscrollcommand=scrollbar.set)
        scrollbar.config(command=tree.yview)
        tree.heading('#0', text='項目名')
        tree.heading('code', text='コード')
        tree.column('#0', width=300)
        tree.column('code', width=150)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # カテゴリー別に項目を整理
        category_items: Dict[str, List[tuple]] = {}
        for code, name, category in selectable:
            cat = category or "OTHERS"
            if cat not in category_items:
                category_items[cat] = []
            category_items[cat].append((code, name))
        
        # カテゴリーごとにソート
        for cat in category_items:
            category_items[cat].sort(key=lambda x: x[1])
        
        # カテゴリー順序（表示順）
        category_order = ["REPRODUCTION", "DHI", "GENOMIC", "HEALTH", "USER", "OTHERS"]
        
        # ロガーを取得
        logger = logging.getLogger(__name__)
        
        # Treeviewに項目を追加
        category_nodes: Dict[str, str] = {}
        # マスターリスト：node_idを含めず、(code, name, cat)のみを保持
        master_items: List[tuple] = []
        initial_selected_node = None
        
        for cat in category_order:
            if cat in category_items and category_items[cat]:
                cat_name = category_names.get(cat, cat)
                cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                category_nodes[cat] = cat_node
                
                for code, name in category_items[cat]:
                    item_node = tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                    master_items.append((code, name, cat))
                    if initial_code and code == initial_code:
                        initial_selected_node = item_node
                        tree.selection_set(item_node)
                        tree.see(item_node)
        
        # カテゴリータグのスタイル設定
        tree.tag_configure('category', font=("Meiryo UI", 10, 'bold'), background='#E0E0E0')
        tree.tag_configure('item', font=("Meiryo UI", 10))
        
        # 初期状態でカテゴリーを展開
        for cat_node in category_nodes.values():
            tree.item(cat_node, open=True)
        
        selected_code: Optional[str] = None
        
        def count_tree_nodes(tree_widget, parent=''):
            """TreeViewの全ノード数を再帰的にカウント"""
            count = 0
            try:
                children = tree_widget.get_children(parent)
                count += len(children)
                for child in children:
                    count += count_tree_nodes(tree_widget, child)
            except Exception as e:
                logger.warning(f"count_tree_nodes error: {e}")
            return count
        
        def clear_tree_all(tree_widget):
            """TreeViewの全ノードを再帰的に削除"""
            try:
                root_children = tree_widget.get_children('')
                for child in root_children:
                    _delete_recursive(tree_widget, child)
            except Exception as e:
                logger.error(f"clear_tree_all error: {e}", exc_info=True)
        
        def _delete_recursive(tree_widget, node_id):
            """再帰的にノードとその子を削除"""
            try:
                children = tree_widget.get_children(node_id)
                for child in children:
                    _delete_recursive(tree_widget, child)
                if tree_widget.exists(node_id):
                    tree_widget.delete(node_id)
            except Exception as e:
                logger.error(f"_delete_recursive error for {node_id}: {e}", exc_info=True)
        
        # filter_tree呼び出しカウンター
        filter_tree_call_count = [0]  # リストでラップしてnonlocalを回避
        
        def filter_tree(*args):
            """検索文字列でフィルタリング（マスターリストから常に再フィルタ）"""
            filter_tree_call_count[0] += 1
            call_num = filter_tree_call_count[0]
            
            try:
                search_text = search_var.get().lower().strip()
                logger.info(f"[filter_tree #{call_num}] called, search_text='{search_text}'")
                logger.info(f"[filter_tree #{call_num}] tree id={id(tree)}, exists={tree.winfo_exists()}")
                
                # 現在のノード数をカウント（削除前）
                nodes_before = count_tree_nodes(tree)
                logger.info(f"[filter_tree #{call_num}] nodes before clear: {nodes_before}")
                
                # マスターリストからフィルタリング
                if search_text:
                    filtered_items = [
                        (code, name, cat) for code, name, cat in master_items
                        if search_text in name.lower() or search_text in code.lower()
                    ]
                else:
                    filtered_items = master_items
                
                logger.info(f"[filter_tree #{call_num}] master_items={len(master_items)}, filtered_items={len(filtered_items)}")
                
                # すべてのノードを削除（カテゴリーノードも含めて完全削除）
                clear_tree_all(tree)
                category_nodes.clear()
                
                # フィルタリング結果に基づいてカテゴリーノードを再作成
                visible_categories = set()
                for code, name, cat in filtered_items:
                    visible_categories.add(cat)
                
                # カテゴリーノードを再作成
                for cat in category_order:
                    if cat in visible_categories:
                        cat_name = category_names.get(cat, cat)
                        try:
                            cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                            category_nodes[cat] = cat_node
                            tree.item(cat_node, open=True)
                        except Exception as e:
                            logger.error(f"[filter_tree #{call_num}] Failed to insert category {cat}: {e}", exc_info=True)
                
                # フィルタリング結果をTreeViewに再描画
                insert_count = 0
                insert_errors = 0
                for code, name, cat in filtered_items:
                    if cat in category_nodes:
                        try:
                            cat_node = category_nodes[cat]
                            if tree.exists(cat_node):
                                tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                                insert_count += 1
                            else:
                                logger.warning(f"[filter_tree #{call_num}] Category node {cat} does not exist")
                                insert_errors += 1
                        except Exception as e:
                            logger.error(f"[filter_tree #{call_num}] Failed to insert item {code} ({name}): {e}", exc_info=True)
                            insert_errors += 1
                
                # 削除後のノード数をカウント
                nodes_after = count_tree_nodes(tree)
                logger.info(f"[filter_tree #{call_num}] insert_count={insert_count}, insert_errors={insert_errors}, nodes after: {nodes_after}")
                
            except Exception as e:
                logger.error(f"[filter_tree #{call_num}] Unexpected error: {e}", exc_info=True)
        
        # クリアボタンのコマンドを設定（filter_tree関数の定義後に実行）
        def clear_search():
            """検索欄をクリアしてフィルタリングをリセット"""
            logger.info("clear_search called")
            search_var.set("")  # StringVar traceで自動的にfilter_treeが呼ばれる
            search_entry.focus_set()
        
        # クリアボタンのコマンドを設定
        clear_button.config(command=clear_search)
        
        # StringVarの変更を監視（これが主な更新方法）
        def on_search_changed(*args):
            """StringVar変更時のコールバック"""
            logger.info(f"on_search_changed called, search_var.get()='{search_var.get()}'")
            dialog.after_idle(filter_tree)
        
        try:
            search_var.trace_add('write', on_search_changed)
            logger.info("StringVar trace_add('write') registered")
        except AttributeError:
            # Python 3.7以前の場合はtraceを使用
            search_var.trace('w', on_search_changed)
            logger.info("StringVar trace('w') registered")
        
        def on_ok():
            nonlocal selected_code
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            
            # 選択されたノードがカテゴリーノードの場合は無視
            selected_node = sel[0]
            if tree.item(selected_node, 'tags')[0] == 'category':
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            
            # 項目ノードからコードを取得
            values = tree.item(selected_node, 'values')
            if values and values[0]:
                selected_code = values[0]
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "項目を選択してください。")
        
        def on_cancel():
            dialog.destroy()
        
        tree.bind('<Double-Button-1>', lambda e: on_ok())
        tree.bind('<Return>', lambda e: on_ok())

        def _on_tree_right_click(event):
            """右クリックで項目をコピーするコンテキストメニューを表示"""
            node = tree.identify_row(event.y)
            if not node or not tree.exists(node):
                return
            tags = tree.item(node, 'tags')
            if tags and tags[0] == 'category':
                return
            tree.selection_set(node)
            name = tree.item(node, 'text')
            values = tree.item(node, 'values')
            code = values[0] if values else ''
            menu = tk.Menu(dialog, tearoff=0)
            menu.add_command(label="コピー", command=lambda: _copy_item_to_clipboard(name, code))
            menu.tk_popup(event.x_root, event.y_root)

        def _copy_item_to_clipboard(name: str, code: str):
            """項目コードをクリップボードにコピー（コマンド・条件欄へのペースト用）"""
            text = code if code else name
            dialog.clipboard_clear()
            dialog.clipboard_append(text)
            dialog.update()

        tree.bind('<Button-3>', _on_tree_right_click)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(10, 0))
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # 検索エントリにフォーカス
        search_entry.focus()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return selected_code
    
    def _get_selectable_items(self) -> List[tuple]:
        """選択可能な計算項目の一覧を取得（後方互換性のため）"""
        items_with_cat = self._get_selectable_items_with_category()
        return [(code, name) for code, name, _ in items_with_cat]
    
    def _get_selectable_items_with_category(self) -> List[tuple]:
        """選択可能な計算項目の一覧を取得（カテゴリー情報付き）"""
        items = []
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        for code, data in item_dict.items():
            item_type = data.get("type")
            origin = data.get("origin", "").lower()
            # calculated, calc, custom の項目を含める
            if item_type not in ("calculated", "calc", "custom") and origin != "custom":
                continue
            if data.get("visible", True) is False:
                continue
            # custom項目の場合はformula/sourceがなくても可
            if origin != "custom" and not data.get("formula") and not data.get("source"):
                continue
            # label, display_name, name_jp の順で取得
            display_name = data.get("label") or data.get("display_name") or data.get("name_jp") or code
            # categoryを取得（大文字に変換）
            category = (data.get("category") or "OTHERS").upper()
            items.append((code, display_name, category))
        items.sort(key=lambda x: (x[2], x[1]))  # カテゴリー、名前の順でソート
        return items
    
    def _get_item_display_name(self, item_code: Optional[str]) -> str:
        """表示名を取得（item_dictionaryに未定義なら item_code を使用）"""
        if not item_code:
            return "（空行）"
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        if item_code in item_dict:
            data = item_dict[item_code]
            # label, display_name, name_jp の順で取得
            return data.get("label") or data.get("display_name") or data.get("name_jp") or item_code
        # 既存の固定表示名への後方互換
        fallback_names = {
            'BTHYR': '生まれた年',
            'BTHYM': '生まれた年月',
            'DIM': '分娩後日数',
            'DAI': '授精後日数',
            'DUE': '分娩予定日',
            'LAST_MILK_YIELD': '最新乳量',
            'LAST_FAT': '最新乳脂肪',
            'LAST_PROTEIN': '最新乳タンパク',
            'LAST_SNF': '最新無脂固形分',
            'LAST_SCC': '最新体細胞数',
            'LAST_LS': '最新LS',
            'LAST_MUN': '最新MUN',
            'LAST_BHB': '最新BHB',
            'LAST_DENOVO_FA': '最新DeNovo FA',
            'LAST_PREFORMED_FA': '最新Preformed FA',
            'LAST_MIXED_FA': '最新Mixed FA',
            'LAST_DENOVO_MILK': '最新DeNovo Milk'
        }
        return fallback_names.get(item_code, item_code)
    
    def _format_calc_value(self, item_code: Optional[str], value: Any) -> str:
        """
        値の表示フォーマット
        
        Args:
            item_code: 項目コード
            value: 値（None の場合は空文字列を返す、0 / 0.0 は有効な値として表示）
        
        Returns:
            フォーマットされた文字列
        """
        # None のみを空表示にする（0 / 0.0 は有効な値として表示）
        if item_code is None or value is None:
            return ""
        
        # 数値型の場合はそのまま表示（0 / 0.0 も含む）
        # 単位は表示しない（基本的に数値のみ）
        if item_code in ['SCC', 'PSCC']:
            try:
                # 整数として表示（カンマ区切り）
                return f"{int(value):,}"
            except (ValueError, TypeError):
                return str(value)
        # 繁殖コード（RC）：コード番号：日本語のみ表示
        if item_code == 'RC' or item_code == 'rc':
            try:
                rc_value = int(value)
                return self.rc_names.get(rc_value, f'{rc_value}：')
            except (ValueError, TypeError):
                return str(value)
        # PENの場合は設定の名称で表示
        if item_code == 'PEN' or item_code == 'pen':
            try:
                code = str(value).strip()
                if code and code in self.pen_settings:
                    return self.pen_settings.get(code, code)
            except Exception:
                pass
        # 受胎授精種類（CSIT）：表示名のみ表示（「自然発情」「WPG」など、コード:名称 は使わない）
        if item_code == 'CSIT' and value:
            s = str(value).strip()
            if "：" in s:
                return s.split("：", 1)[1].strip()
            if ":" in s:
                return s.split(":", 1)[1].strip()
            if self.insemination_types_dict:
                return self.insemination_types_dict.get(s, s)
            return s
        # 削除された項目（LAST_MILK_YIELD, LAST_FAT, LAST_PROTEIN, LAST_SNF, LAST_SCC等）は
        # item_dictionaryから削除されているため、通常はここに来ないが、
        # 念のため存在チェックを追加（削除された項目はそのまま文字列として表示）
        # その他の項目も 0 / 0.0 を含めて表示
        return str(value)

    def _on_basic_right_click(self, key: str, event):
        """基本情報の品種・群で右クリック時：編集・追加メニューを表示"""
        if key not in ("brd", "pen"):
            return
        menu = tk.Menu(self.frame, tearoff=0)
        if key == "brd":
            menu.add_command(label="編集・追加", command=self._on_edit_brd)
        else:
            menu.add_command(label="編集・追加", command=self._on_edit_pen)
        menu.post(event.x_root, event.y_root)
    
    def _on_edit_brd(self, _event=None):
        """品種（brd）を編集"""
        if not self.cow_auto_id:
            return
        
        cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
        if not cow:
            messagebox.showerror("エラー", "個体情報が取得できませんでした。")
            return
        
        current_brd = str(cow.get("brd") or "")
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("品種を編集")
        dialog.geometry("320x160")
        
        ttk.Label(dialog, text="品種:", font=("Meiryo UI", 10)).pack(pady=(15, 5))
        entry = ttk.Entry(dialog, width=20)
        entry.pack()
        entry.insert(0, current_brd)
        entry.focus()
        
        def on_ok():
            new_brd = entry.get().strip()
            try:
                self.db.update_cow(self.cow_auto_id, {"brd": new_brd})
                self.basic_labels['brd'].config(text=new_brd)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"品種の更新に失敗しました: {e}")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_edit_lact(self, _event=None):
        """産次（lact）を編集"""
        if not self.cow_auto_id:
            return
        
        cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
        if not cow:
            messagebox.showerror("エラー", "個体情報が取得できませんでした。")
            return
        
        current_lact = cow.get("lact")
        current_lact_str = str(current_lact) if current_lact is not None else ""
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("産次を編集")
        dialog.geometry("320x160")
        
        ttk.Label(dialog, text="産次:", font=("Meiryo UI", 10)).pack(pady=(15, 5))
        entry = ttk.Entry(dialog, width=20)
        entry.pack()
        entry.insert(0, current_lact_str)
        entry.focus()
        
        def on_ok():
            new_lact_str = entry.get().strip()
            if not new_lact_str:
                # 空の場合はNoneを設定
                new_lact = None
            else:
                try:
                    new_lact = int(new_lact_str)
                    if new_lact < 0:
                        messagebox.showwarning("警告", "産次は0以上の整数で入力してください。")
                        return
                except ValueError:
                    messagebox.showwarning("警告", "産次は整数で入力してください。")
                    return
            
            try:
                self.db.update_cow(self.cow_auto_id, {"lact": new_lact, "baseline_lact": new_lact})
                self.basic_labels['lact'].config(text=str(new_lact) if new_lact is not None else '')
                # 産次は手動変更のためRuleEngineの再計算（イベントから算出してcowを上書き）は行わないが、
                # 履歴の event_lact / event_dim は個体カードの産次に合わせて更新する（今産次＝カード、前産次＝カード−1 …）
                if self.rule_engine and new_lact is not None:
                    self.rule_engine.recalculate_events_for_cow_with_lact_override(self.cow_auto_id, new_lact)
                # 表示を更新
                self.refresh()
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"産次の更新に失敗しました: {e}")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_edit_pen(self, _event=None):
        """PEN（群）を変更"""
        if not self.cow_auto_id:
            return
        
        # PEN設定を再読込（最新の設定を反映）
        try:
            if self.settings_manager:
                self.pen_settings = self.settings_manager.load_pen_settings()
        except Exception:
            pass
        
        def sort_key(item):
            code = str(item[0])
            return int(code) if code.isdigit() else 9999
        
        display_to_code: Dict[str, str] = {}
        options = []
        for code, name in sorted(self.pen_settings.items(), key=sort_key):
            display = f"{code}：{name}" if name else str(code)
            options.append(display)
            display_to_code[display] = str(code)
        
        cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
        current_pen = cow.get('pen', '') if cow else ''
        current_display = ""
        if current_pen:
            current_code = str(current_pen)
            if current_code in self.pen_settings:
                name = self.pen_settings.get(current_code, '')
                current_display = f"{current_code}：{name}" if name else current_code
            else:
                # 名前で保存されている場合に備える
                for code, name in self.pen_settings.items():
                    if str(name) == str(current_pen):
                        current_display = f"{code}：{name}" if name else str(code)
                        break
                if not current_display:
                    current_display = str(current_pen)
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("群を変更")
        dialog.geometry("320x160")
        
        ttk.Label(dialog, text="群:", font=("Meiryo UI", 10)).pack(pady=(15, 5))
        combo = ttk.Combobox(dialog, values=options, state="readonly", width=25)
        combo.pack()
        if current_display in display_to_code:
            combo.set(current_display)
        
        def on_ok():
            selected = combo.get().strip()
            if not selected:
                messagebox.showwarning("警告", "群を選択してください。")
                return
            new_code = display_to_code.get(selected)
            if not new_code:
                # フォールバック: 文字列の先頭コードを使う
                if "：" in selected:
                    new_code = selected.split("：", 1)[0].strip()
                else:
                    new_code = selected
            try:
                self.db.update_cow(self.cow_auto_id, {"pen": new_code})
                # 表示更新
                self.refresh_calculated_items()
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    pen_value = cow.get('pen', '')
                    if isinstance(pen_value, str) and pen_value in self.pen_settings:
                        pen_value = self.pen_settings.get(pen_value, pen_value)
                    self.basic_labels['pen'].config(text=pen_value)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"群の更新に失敗しました: {e}")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _load_calculated_item_codes(self) -> List[str]:
        """表示定義リストをロード（無ければデフォルト生成）"""
        # 削除された項目（LAST_MILK_YIELD, LAST_FAT, LAST_PROTEIN, LAST_SCC等）は
        # item_dictionaryから削除されているため、デフォルトから除外
        default_codes = [
            "DIM",
            "DAI",
            "DUE",
        ]
        path = self.calculated_items_path
        try:
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list) and all(isinstance(x, str) for x in data):
                    return data
            # 無い場合はデフォルトを保存して返す
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(default_codes, f, ensure_ascii=False, indent=2)
            return default_codes
        except Exception as e:
            logging.error(f"calculated_items 読み込みエラー: {e}")
            return default_codes
    
    def _load_tab_orders(self) -> Dict[str, List[str]]:
        """各タブの並び順をロード"""
        path = self.calc_layout_path
        # calculated_item_codesは既に読み込まれている前提
        default_orders = {
            "USER": list(self.calculated_item_codes) if hasattr(self, 'calculated_item_codes') else [],
            "REPRODUCTION": [],
            "DHI": [],
            "GENOMIC": [],
            "HEALTH": [],
            "OTHERS": []
        }
        
        try:
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    # 旧形式のuser_tab_orderをサポート
                    if "user_tab_order" in data and isinstance(data["user_tab_order"], list):
                        default_orders["USER"] = data["user_tab_order"]
                    # 新形式のtab_ordersをサポート
                    if "tab_orders" in data and isinstance(data["tab_orders"], dict):
                        for tab_name in default_orders.keys():
                            if tab_name in data["tab_orders"] and isinstance(data["tab_orders"][tab_name], list):
                                default_orders[tab_name] = data["tab_orders"][tab_name]
            return default_orders
        except Exception as e:
            logging.error(f"tab_orders 読み込みエラー: {e}")
            return default_orders
    
    def _save_tab_orders(self):
        """各タブの並び順を保存"""
        try:
            self.calc_layout_path.parent.mkdir(parents=True, exist_ok=True)
            data = {
                "tab_orders": self.tab_orders
            }
            with open(self.calc_layout_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"tab_orders 保存エラー: {e}")
    
    def _save_calculated_item_codes(self):
        """表示定義リストを保存"""
        try:
            self.calculated_items_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.calculated_items_path, 'w', encoding='utf-8') as f:
                json.dump(self.calculated_item_codes, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"calculated_items 保存エラー: {e}")
    
    def _calculate_dim_at_event_date(
        self,
        events: List[Dict[str, Any]],
        event_date: str,
        cow: Optional[Dict[str, Any]] = None,
    ) -> Optional[int]:
        """
        指定されたイベント日付時点でのDIM（分娩後日数）を計算
        
        分娩日は原則としてイベント一覧中の分娩イベントから「そのイベント日以前の最新」を採用する。
        cow.clvd がイベント日以前かつイベント一覧の分娩より新しい場合のみ、cow.clvd を採用する
        （基本情報の分娩月日・DIM とイベント履歴の DIM を整合させるため）。
        
        Args:
            events: イベントリスト
            event_date: イベント日付（YYYY-MM-DD形式）
            cow: 牛データ（省略可）。clvd がある場合、上記の条件でフォールバックに使用する。
        
        Returns:
            DIM（日数）、分娩日が見つからない場合はNone
        """
        if not event_date:
            return None
        
        try:
            from datetime import datetime
            
            # イベント日付をdatetimeオブジェクトに変換
            event_dt = datetime.strptime(event_date, '%Y-%m-%d')
            
            # そのイベント日付以前の最新の分娩日を探す（イベント一覧の分娩から）
            latest_calving_date = None
            for event in events:
                if event.get('event_number') == RuleEngine.EVENT_CALV:
                    calving_date = event.get('event_date')
                    if calving_date:
                        try:
                            calving_dt = datetime.strptime(calving_date, '%Y-%m-%d')
                            if calving_dt <= event_dt:
                                if latest_calving_date is None or calving_dt > latest_calving_date:
                                    latest_calving_date = calving_dt
                        except ValueError:
                            continue
            
            # cow.clvd がある場合：イベント日以前かつ、イベント内の最新分娩より新しい場合は採用
            # （分娩イベントが未登録だが cow の分娩月日だけ更新されている場合の整合用）
            if cow and cow.get('clvd'):
                try:
                    clvd_dt = datetime.strptime(cow['clvd'], '%Y-%m-%d')
                    if clvd_dt <= event_dt and (
                        latest_calving_date is None or clvd_dt > latest_calving_date
                    ):
                        latest_calving_date = clvd_dt
                except (ValueError, TypeError):
                    pass
            
            if latest_calving_date:
                dim = (event_dt - latest_calving_date).days
                return dim if dim >= 0 else None
            
            return None
        except (ValueError, TypeError) as e:
            logging.warning(f"[CowCard] DIM計算エラー: event_date={event_date}, error={e}")
            return None
    
    def _display_events(self, cow_auto_id: int):
        """イベント履歴を表示（event_date DESC順）"""
        try:
            
            # 既存のアイテムをクリア
            for item in self.event_tree.get_children():
                self.event_tree.delete(item)
            
            # イベントを取得（既にevent_date DESC順でソート済み）
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            # 基本情報の分娩月日（cow.clvd）と整合させるため、DIM計算で参照する
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            
            # 繁殖関連イベントのみ表示フィルターが有効な場合はフィルタリング
            if self.reproduction_filter_var.get():
                events = [
                    event for event in events
                    if event.get('event_number') is not None
                    and (200 <= event.get('event_number') < 300 or 300 <= event.get('event_number') < 400)
                ]
            
            # 表示されたイベント数をカウント
            displayed_count = 0
            
            # Treeviewに追加（すべてのイベントを表示）
            for event in events:
                try:
                    event_date = event.get('event_date', '')
                    event_number = event.get('event_number')
                    
                    # event_numberがNoneの場合はスキップ
                    if event_number is None:
                        logging.warning(f"[CowCard] Skipping event with None event_number: event_id={event.get('id')}")
                        continue
                    
                    event_name = self._get_event_name(event_number)
                    # json_dataはDBから取得したものを一切加工せず、そのまま保持する
                    # get_events_by_cowで既にパース済み（辞書または空辞書）
                    json_data = event.get('json_data')
                    
                    # json_dataがNoneの場合は空辞書に変換（表示処理用）
                    if json_data is None:
                        json_data = {}
                    
                    # noteを初期化（AI/ETイベント以外は既存のnoteを使用）
                    note = ''
                    
                    # AI/ETイベント（200, 201）の場合は詳細情報を備考欄に追加
                    # 【必須ルール】event["memo"]やevent["note"]は一切使わず、必ずjson_dataから再生成する
                    # イベント入力ウィンドウと同じロジックを使用
                    if event_number in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
                        # json_dataからoutcomeを取得（P/O/R/N）
                        outcome = json_data.get('outcome')
                        
                        # 共通関数を使用して表示文字列を生成（イベント入力ウィンドウと同じ）
                        detail_text = format_insemination_event(
                            json_data,
                            self.technicians_dict,          # code -> name
                            self.insemination_types_dict,   # code -> name
                            outcome                          # 受胎ステータス（P/O/R/N）
                        )
                        
                        if detail_text:
                            note = detail_text
                            # 先頭の全角スペースを削除して左揃え（表示のため）
                            if note:
                                note = note.lstrip('　').strip()
                                # すべて削除された場合は元に戻す
                                if not note:
                                    note = detail_text.strip()
                        else:
                            note = ""
                    
                    # 分娩イベントの表示
                    elif event_number == RuleEngine.EVENT_CALV:
                        calv_def = self.event_dictionary.get(str(RuleEngine.EVENT_CALV), {})
                        diff_labels = calv_def.get("calving_difficulty", {})
                        detail_text = format_calving_event(json_data, diff_labels)
                        if detail_text:
                            note = detail_text
                        else:
                            note = ""
                    
                    # 乳検イベント（601）の場合は乳量のみを表示
                    elif event_number == RuleEngine.EVENT_MILK_TEST:
                        # 乳量のみを表示
                        milk_yield = json_data.get('milk_yield')
                        if milk_yield is not None:
                            note = f"乳量{milk_yield}kg"
                        else:
                            note = ""
                    
                    # REPRODUCTIONカテゴリイベントの場合は詳細情報を備考欄に追加
                    # イベント番号で判定（300, 301, 302）またはcategoryで判定
                    # ただし、AI/ET/分娩/乳検イベントの場合は既にnoteが設定されているのでスキップ
                    if event_number not in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET, RuleEngine.EVENT_CALV, RuleEngine.EVENT_MILK_TEST]:
                        event_dict = self.event_dictionary.get(str(event_number), {})
                        is_breeding = (event_number in [300, 301, RuleEngine.EVENT_PDN] or 
                                       event_dict.get('category') == 'REPRODUCTION')
                        
                        if is_breeding:
                            # 整列表示関数を使用
                            detail_text = format_reproduction_check_event(json_data)
                            if detail_text:
                                # 既存の備考がある場合は結合
                                if note:
                                    note = f"{detail_text} | {note}"
                                else:
                                    note = detail_text
                        # 先頭の全角スペースのみ削除して左揃え（内容は維持）
                        if note:
                            # 先頭の全角スペースだけを削除（内容が全角スペースだけの場合は削除しない）
                            original_note = note
                            while note.startswith('　') and len(note) > len('　'):
                                note = note[1:]
                            # 全角スペースだけになった場合は元に戻す
                            if not note or note.strip() == '':
                                note = original_note.strip()
                            else:
                                note = note.strip()
                    # AI/ET/繁殖検査/乳検以外のイベントは既存のnoteを使用
                    # ただし、AI/ETイベントの場合は既にnoteが設定されているのでスキップ
                    elif event_number not in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET, RuleEngine.EVENT_CALV, RuleEngine.EVENT_MILK_TEST]:
                        # REPRODUCTIONカテゴリイベントでもない場合のみ
                        event_dict = self.event_dictionary.get(str(event_number), {})
                        is_breeding = (event_number in [300, 301, RuleEngine.EVENT_PDN] or 
                                       event_dict.get('category') == 'REPRODUCTION')
                        if not is_breeding:
                            note = event.get('note', '') or ''
                        # 先頭の全角スペースのみ削除して左揃え（内容は維持）
                        if note:
                            # 先頭の全角スペースだけを削除（内容が全角スペースだけの場合は削除しない）
                            original_note = note
                            while note.startswith('　') and len(note) > len('　'):
                                note = note[1:]
                            # 全角スペースだけになった場合は元に戻す
                            if not note or note.strip() == '':
                                note = original_note.strip()
                            else:
                                note = note.strip()
                
                    # イベントの表示色を決定
                    display_color = self._get_event_display_color(event_number)
                    
                    # 色タグを動的に設定（まだ設定されていない場合）
                    color_tag = self._ensure_color_tag(display_color, event_number)
                    
                    # イベントIDを取得
                    event_id = event.get('id')
                    if event_id is None:
                        logging.warning(f"[CowCard] Skipping event with None event_id: event_number={event_number}")
                        continue
                    
                    # イベント時点でのDIMを計算（cow を渡して基本情報の分娩月日と整合）
                    dim_value = self._calculate_dim_at_event_date(events, event_date, cow)
                    dim_display = str(dim_value) if dim_value is not None else ''
                    
                    # Treeviewに追加（iid に event_id を紐づける）
                    # DB由来イベントは iid = str(event_id) とする
                    self.event_tree.insert(
                        '',
                        'end',
                        iid=str(event_id),
                        values=(event_date, dim_display, event_name, note),
                        tags=(color_tag,)
                    )
                    displayed_count += 1
                        
                except Exception as e:
                    # 個別のイベント処理でエラーが発生しても、他のイベントは表示を続ける
                    logging.error(f"[CowCard] Error processing event: {e}, event_id={event.get('id')}")
                    import traceback
                    traceback.print_exc()
                    continue
            
            logging.debug(f"[CowCard] displayed events = {displayed_count}")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _display_events で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _calculate_and_notify_latest_dates(self, cow_auto_id: int):
        """
        最終日付を計算してコールバックで通知
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            # イベントを取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            # 最終分娩日を計算（EVENT_CALV = 202、baselineも含む）
            latest_calving = self.get_latest_event_date(events, [RuleEngine.EVENT_CALV])
            
            # 最終AI日を計算（AI/ET系イベント: 200, 201, 203, 204など）
            # ユーザー要求では「200, 201, 203, 204 など」とあるが、一般的には200(AI), 201(ET)が対象
            # 203(DRY), 204(STOPR)は繁殖関連だがAI/ETではないため、200, 201のみを対象とする
            latest_ai = self.get_latest_event_date(events, [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET])
            
            # 最終乳検日を計算（EVENT_MILK_TEST = 601）
            latest_milk = self.get_latest_event_date(events, [RuleEngine.EVENT_MILK_TEST])
            
            # 最終イベント日を計算（全イベント対象）
            latest_any = self.get_latest_any_event_date(events)
            
            # コールバックがあれば呼び出す
            if self.on_latest_dates_calculated_callback:
                self.on_latest_dates_calculated_callback(latest_calving, latest_ai, latest_milk, latest_any)
                
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _calculate_and_notify_latest_dates で例外が発生しました: {e}")
            traceback.print_exc()
            # エラー時もコールバックを呼ぶ（Noneを渡す）
            if self.on_latest_dates_calculated_callback:
                self.on_latest_dates_calculated_callback(None, None, None, None)
    
    def _init_settings_manager(self, cow_auto_id: int):
        """
        SettingsManagerを初期化し、授精設定辞書・PEN設定を読み込む
        
        農場パスは (1) 牛の frm があれば C:/FARMS/{frm}、(2) なければ DB の親フォルダを使用する。
        これにより frm が未設定の個体でも、現在開いている農場の群設定が表示される。
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            farm_path = None
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                frm = cow.get('frm')
                if frm:
                    farm_path = FARMS_ROOT / frm
            
            # frm が無い場合は DB の親フォルダを農場パスとして使用（農場設定の群を表示するため）
            if farm_path is None and hasattr(self.db, 'db_path') and self.db.db_path:
                farm_path = Path(self.db.db_path).parent
            
            if farm_path and Path(farm_path).exists():
                self.settings_manager = SettingsManager(farm_path)
                self._load_insemination_settings(farm_path)
                try:
                    self.pen_settings = self.settings_manager.load_pen_settings()
                except Exception as e:
                    logging.error(f"PEN設定の読み込みに失敗しました: {e}")
                    self.pen_settings = {}
            else:
                self.technicians_dict = {}
                self.insemination_types_dict = {}
                self.pen_settings = {}
        except Exception as e:
            print(f"WARNING: SettingsManagerの初期化に失敗しました: {e}")
            self.settings_manager = None
            self.technicians_dict = {}
            self.insemination_types_dict = {}
            self.pen_settings = {}
    
    def _load_insemination_settings(self, farm_path: Path):
        """
        insemination_settings.json をロード（EventInputWindowと同じ方法）
        
        Args:
            farm_path: 農場フォルダパス
        """
        settings_file = farm_path / "insemination_settings.json"
        
        if not settings_file.exists():
            logging.warning(f"insemination_settings.json が見つかりません: {settings_file}")
            self.technicians_dict = {}
            self.insemination_types_dict = {}
            return
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            self.technicians_dict = settings.get('technicians', {})
            self.insemination_types_dict = settings.get('insemination_types', {})
            
            logging.debug(f"授精設定をロード: technicians={len(self.technicians_dict)}, insemination_types={len(self.insemination_types_dict)}")
            logging.debug(f"technicians_dict sample: {list(self.technicians_dict.items())[:3] if self.technicians_dict else 'empty'}")
            logging.debug(f"insemination_types_dict sample: {list(self.insemination_types_dict.items())[:3] if self.insemination_types_dict else 'empty'}")
        except Exception as e:
            logging.error(f"insemination_settings.json 読み込みエラー: {e}")
            self.technicians_dict = {}
            self.insemination_types_dict = {}
    
    def _on_event_double_click(self, event):
        """イベントのダブルクリック処理（編集ウィンドウを開く）"""
        try:
            selection = self.event_tree.selection()
            if selection:
                row_id = selection[0]
                # iid から直接 event_id を取得
                # DBイベント以外（baseline 等）は数値でない iid を使うため、isdigit() でチェック
                if not row_id.isdigit():
                    # ダブルクリック不可イベント（baseline 等）の場合は何もしない
                    return

                event_id = int(row_id)
                # 編集として同じ処理を実行（編集ウィンドウを開く）
                self._on_edit_event(event_id)
        except Exception as e:
            import traceback
            print(f"ERROR: _on_event_double_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_event_right_click(self, event):
        """イベントの右クリック処理（コンテキストメニュー表示）"""
        try:
            # 右クリック位置の行を取得
            row_id = self.event_tree.identify_row(event.y)
            if not row_id:
                return
            
            # 行を選択・フォーカス
            self.event_tree.selection_set(row_id)
            self.event_tree.focus(row_id)
            
            # iid から直接 event_id を取得
            # DBイベント以外（baseline 等）は数値でない iid を使うため、isdigit() でチェック
            if not row_id.isdigit():
                # 削除不可イベント（baseline 等）の場合はメニューを表示しない
                return
            
            event_id = int(row_id)
            logging.debug(f"Right-click event_id={event_id}")
            
            # イベントデータを取得して削除可否をチェック
            event_data = self.db.get_event_by_id(event_id)
            if event_data is None:
                # イベントが見つからない場合はメニューを表示しない
                logging.warning(f"右クリック対象イベントが取得できません: event_id={event_id}")
                return
            
            # json_data を安全に取得（None / str / dict のいずれでも安全に処理）
            json_data_raw = event_data.get('json_data')
            # json_data が dict でない場合は空dictとして扱う
            if not isinstance(json_data_raw, dict):
                json_data = {}
            else:
                json_data = json_data_raw
            
            event_number = event_data.get('event_number')
            
            # system_generated かつ deprecated のイベントのチェック（dict の場合のみ）
            is_system_deprecated = False
            if isinstance(json_data, dict) and json_data.get('system_generated', False):
                event_str = str(event_number)
                if event_str in self.event_dictionary:
                    if self.event_dictionary[event_str].get('deprecated', False):
                        is_system_deprecated = True
            
            # システム生成・deprecated のみメニュー非表示（基準分娩は誤入力修正のため編集可）
            if is_system_deprecated:
                return
            
            # コンテキストメニューを作成
            menu = tk.Menu(self.frame, tearoff=0)
            menu.add_command(label="編集", command=lambda: self._on_edit_event(event_id))
            menu.add_command(label="削除", command=lambda: self._on_delete_event(event_id))
            
            # メニューを表示
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
                
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_event_right_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_edit_event(self, event_id: int):
        """イベント編集処理"""
        try:
            if not self.cow_auto_id:
                messagebox.showerror("エラー", "対象牛が設定されていません")
                return
            
            if not self.rule_engine:
                messagebox.showerror("エラー", "RuleEngineが設定されていません")
                return
            
            # イベントデータを取得
            event = self.db.get_event_by_id(event_id)
            if not event:
                messagebox.showerror("エラー", "イベントが見つかりません")
                return
            
            logging.debug(f"Editing event_id={event_id}")
            
            # 農場パスを取得
            farm_path = None
            if self.cow_auto_id:
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    frm = cow.get('frm')
                    if frm:
                        farm_path = FARMS_ROOT / frm
            
            # EventInputWindowを編集モードで開く（既存があれば再利用）
            event_window = EventInputWindow.open_or_focus(
                parent=self.frame.winfo_toplevel(),
                db_handler=self.db,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path,
                cow_auto_id=self.cow_auto_id,
                on_saved=self._on_event_saved,
                farm_path=farm_path,
                edit_event_id=event_id  # 編集モード
            )
            event_window.show()
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_edit_event で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"イベントの編集に失敗しました: {e}")
    
    def _on_delete_event(self, event_id: int):
        """イベント削除処理"""
        try:
            # イベントデータを取得
            event = self.db.get_event_by_id(event_id)
            if event is None:
                messagebox.showwarning(
                    "削除不可",
                    "イベント情報が取得できません"
                )
                return
            
            # 削除制限チェック
            # json_data を安全に取得（None / str / dict のいずれでも安全に処理）
            json_data_raw = event.get('json_data')
            # json_data が dict でない場合は空dictとして扱う
            if not isinstance(json_data_raw, dict):
                json_data = {}
            else:
                json_data = json_data_raw
            
            event_number = event.get('event_number')
            
            # deprecated=true かつ system_generated=true のイベントのチェック
            # 注：event_dictionary.jsonのdeprecatedフラグとjson_dataのsystem_generatedフラグを確認
            # system_generated 判定は dict の場合のみ行う
            if isinstance(json_data, dict) and json_data.get('system_generated', False):
                # event_dictionary.jsonでdeprecatedかどうか確認
                event_str = str(event_number)
                if event_str in self.event_dictionary:
                    if self.event_dictionary[event_str].get('deprecated', False):
                        messagebox.showerror("エラー", "このイベントは削除できません（システム生成イベント）")
                        return
            
            # 確認ダイアログ（基準分娩は誤登録の修正もあるため、文言で明示）
            is_baseline_calv = (
                event_number == RuleEngine.EVENT_CALV
                and isinstance(json_data, dict)
                and json_data.get("baseline_calving", False)
            )
            if is_baseline_calv:
                confirm_msg = (
                    "この分娩は「基準分娩」として登録されています（取込・導入時の初期値など）。\n"
                    "誤入力の修正のため削除する場合は「はい」を選んでください。\n\n"
                    "このイベントを削除しますか？\n（この操作は元に戻せません）"
                )
            else:
                confirm_msg = "このイベントを削除しますか？\n（この操作は元に戻せません）"
            result = messagebox.askyesno("確認", confirm_msg)
            
            if not result:
                return
            
            logging.info(f"Deleting event_id={event_id}")
            
            # 物理削除するとイベントが取れなくなるため、先に cow_auto_id を取得
            cow_auto_id_for_recalc = event.get('cow_auto_id')
            
            # イベントを削除（物理削除）
            self.db.delete_event(event_id, soft_delete=False)
            
            try:
                from modules.activity_log import record_from_event
                record_from_event(
                    self.db,
                    cow_auto_id=cow_auto_id_for_recalc,
                    action="削除",
                    event_number=event_number,
                    event_id=event_id,
                    event_date=str(event.get("event_date") or "")[:10],
                    event_dictionary=self.event_dictionary,
                )
            except Exception:
                pass
            
            # RuleEngineを再実行（物理削除後はイベント取得不可のため cow_auto_id を渡す）
            if self.rule_engine:
                self.rule_engine.on_event_deleted(event_id, cow_auto_id=cow_auto_id_for_recalc)
            
            # 表示を更新（load_cow内で日付も再計算される）
            self.refresh()
            
            # 外部コールバックがあれば呼ぶ
            if self.on_event_saved_callback:
                self.on_event_saved_callback(self.cow_auto_id)
            
            # 削除後のイベント件数をログ出力
            events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
            logging.info(f"Event deleted: event_id={event_id}, remaining events: {len(events)}")
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_delete_event で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def _on_reproduction_filter_changed(self):
        """繁殖フィルターチェックボックスの変更時の処理"""
        if self.cow_auto_id:
            # イベント履歴のみを再表示（他の情報は更新しない）
            self._display_events(self.cow_auto_id)
    
    def refresh(self):
        """表示を更新"""
        if self.cow_auto_id:
            self.load_cow(self.cow_auto_id)
    
    def _on_add_event(self):
        """イベント追加ボタンクリック時の処理"""
        if not self.cow_auto_id:
            print("WARNING: cow_auto_id が設定されていません")
            return
        
        if not self.rule_engine:
            print("WARNING: rule_engine が設定されていません")
            return
        
        try:
            # EventInputWindow を開く（cow_auto_id を指定）
            # 農場パスを取得（新規SIRE登録ダイアログ等に必要。frm が無い場合は DB の親フォルダを使用）
            farm_path = None
            if self.cow_auto_id:
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    frm = cow.get('frm')
                    if frm:
                        farm_path = FARMS_ROOT / frm
            if farm_path is None and hasattr(self.db, 'db_path') and self.db.db_path:
                farm_path = Path(self.db.db_path).parent
            
            event_window = EventInputWindow.open_or_focus(
                parent=self.frame.winfo_toplevel(),
                db_handler=self.db,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path,
                cow_auto_id=self.cow_auto_id,
                on_saved=self._on_event_saved,
                farm_path=farm_path
            )
            event_window.show()
        except Exception as e:
            import traceback
            print(f"ERROR: _on_add_event で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_event_saved(self, cow_auto_id: int):
        """
        イベント保存後のコールバック
        
        Args:
            cow_auto_id: イベントが保存された牛の auto_id
        """
        try:
            # 個体カードを更新
            if self.cow_auto_id == cow_auto_id:
                self.refresh()
            
            # 外部コールバックがあれば呼ぶ
            if self.on_event_saved_callback:
                self.on_event_saved_callback(cow_auto_id)
        except Exception as e:
            import traceback
            print(f"ERROR: _on_event_saved で例外が発生しました: {e}")
            traceback.print_exc()
    
    def set_on_event_saved(self, callback: Callable[[int], None]):
        """
        イベント保存時のコールバックを設定
        
        Args:
            callback: コールバック関数（cow_auto_id を引数に取る）
        """
        self.on_event_saved_callback = callback
    
    def set_on_latest_dates_calculated(self, callback: Callable[[Optional[str], Optional[str], Optional[str], Optional[str]], None]):
        """
        最終日付計算時のコールバックを設定
        
        Args:
            callback: コールバック関数（latest_calving, latest_ai, latest_milk, latest_any を引数に取る）
        """
        self.on_latest_dates_calculated_callback = callback
    
    def get_widget(self) -> ttk.Frame:
        """ウィジェットを取得"""
        return self.frame

