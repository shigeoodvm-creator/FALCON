"""
FALCON2 - MainWindow（メインウィンドウ）
アプリ全体のフレーム（左：サイドメニュー、右：メイン表示領域）
設計書 第11章・第13章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Literal, List, Dict, Any, Tuple, Union
from pathlib import Path
from datetime import datetime, timedelta
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
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    logging.warning("matplotlib がインストールされていません。散布図機能は使用できません。")

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.chatgpt_client import ChatGPTClient
from modules.query_normalizer import QueryNormalizer
from modules.query_executor import QueryExecutor
from modules.query_router import QueryRouter
from settings_manager import SettingsManager
try:
    from modules.query_executor_v2 import QueryExecutorV2
    from modules.query_renderer import QueryRenderer
    from modules.query_dsl import UnhandledQuery, AmbiguousQuery
    QUERY_V2_AVAILABLE = True
except ImportError:
    QueryExecutorV2 = None
    QueryRenderer = None
    UnhandledQuery = None
    AmbiguousQuery = None
    QUERY_V2_AVAILABLE = False
    # モジュールが見つからない場合は警告を出さない（オプション機能のため）
from modules.milk_report_extras import (
    MILK_REPORT_COMMENT_CSS,
    build_comment_modal_html,
    build_comment_script,
    build_monthly_trend_section_html,
    compute_monthly_milk_trend,
    make_milk_report_comment_storage_key,
)
from modules.text_normalizer import normalize_user_input
from modules.app_settings_manager import get_app_settings_manager
from modules.activity_log import load_entries
from ui.cow_card import CowCard
from ui.event_input import EventInputWindow
from ui.batch_item_edit_window import BatchItemEditWindow
from modules.batch_update import BatchUpdate
from datetime import date, timedelta

# ========== Mixin imports ==========
from ui.main_window_dashboard import DashboardMixin
from ui.main_window_conception_rate import ConceptionRateMixin
from ui.main_window_scatter import ScatterPlotMixin
from ui.main_window_chatgpt import ChatGPTMixin
from ui.main_window_chatview import ChatViewMixin


class MainWindow(
    DashboardMixin,
    ConceptionRateMixin,
    ScatterPlotMixin,
    ChatGPTMixin,
    ChatViewMixin,
):
    """メインウィンドウ（サイドメニュー + メイン表示領域）"""
    
    # 繁殖コード（RC）表示：コード番号：日本語のみ（英語表記は省く）
    RC_NAMES = {
        RuleEngine.RC_STOPPED: "1：繁殖停止",
        RuleEngine.RC_FRESH: "2：分娩後",
        RuleEngine.RC_BRED: "3：授精後",
        RuleEngine.RC_OPEN: "4：空胎",
        RuleEngine.RC_PREGNANT: "5：妊娠中",
        RuleEngine.RC_DRY: "6：乾乳中"
    }
    
    @staticmethod
    def format_rc(rc_value: Any) -> str:
        """
        繁殖コード（RC）をフォーマットして表示（コード番号：日本語のみ）
        
        Args:
            rc_value: 繁殖コードの値（int, str, Noneなど）
        
        Returns:
            フォーマットされた文字列（例: "5：妊娠中"）
        """
        if rc_value is None:
            return ""
        try:
            rc_int = int(rc_value)
            return MainWindow.RC_NAMES.get(rc_int, f"{rc_int}：")
        except (ValueError, TypeError):
            return str(rc_value)

    def _get_pen_settings_map(self) -> Dict[str, str]:
        """PEN設定を取得（コード -> 表示名）"""
        try:
            settings_manager = SettingsManager(self.farm_path)
            return settings_manager.load_pen_settings()
        except Exception as e:
            logging.error(f"PEN設定の読み込みに失敗しました: {e}")
            return {}

    def _format_pen_value(self, value: Any, pen_settings: Optional[Dict[str, str]] = None) -> str:
        """PEN値を設定に基づいて表示名へ変換"""
        if value is None:
            return ""
        value_str = str(value).strip()
        if not value_str:
            return ""
        settings = pen_settings if pen_settings is not None else self._get_pen_settings_map()
        return settings.get(value_str, value_str)

    def _normalize_pen_values(self):
        """PENの保存値をPEN設定のコードに正規化（同一農場内）"""
        pen_settings = self._get_pen_settings_map()
        if not pen_settings:
            return
        
        # name -> code
        name_to_code: Dict[str, str] = {}
        for code, name in pen_settings.items():
            if not name:
                continue
            code_str = str(code)
            name_str = str(name)
            name_to_code[name_str] = code_str
            # スペース除去の別名も許容（例: Lactating1）
            name_no_space = name_str.replace(" ", "")
            if name_no_space and name_no_space not in name_to_code:
                name_to_code[name_no_space] = code_str
        
        # 例外的に「Lactating」だけ古い値として存在する場合
        if "Lactating" not in name_to_code:
            for code, name in pen_settings.items():
                if not name:
                    continue
                name_compact = str(name).replace(" ", "").lower()
                if name_compact == "lactating1":
                    name_to_code["Lactating"] = str(code)
                    break
        
        if not name_to_code:
            return
        
        try:
            conn = self.db.connect()
            cursor = conn.cursor()
            for old_value, new_code in name_to_code.items():
                cursor.execute(
                    "UPDATE cow SET pen = ? WHERE pen = ?",
                    (new_code, old_value)
                )
            conn.commit()
        except Exception as e:
            logging.error(f"PEN値の正規化に失敗しました: {e}")
    
    def __init__(self, root: tk.Tk, db_handler: DBHandler, 
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 farm_path: Path):
        """
        初期化
        
        Args:
            root: Tkinter ルートウィンドウ
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            farm_path: 農場フォルダのパス
        """
        self.root = root
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.farm_path = farm_path
        
        # 辞書パス（イベント・項目＝本体のみ、コマンド＝農場ごと）
        app_root = Path(__file__).parent.parent.parent
        config_default = app_root / "config_default"
        
        # event_dictionary.json：本体のみ（config_default）
        default_event_dict = config_default / "event_dictionary.json"
        if default_event_dict.exists():
            self.event_dict_path = default_event_dict
        else:
            print(f"警告: event_dictionary.json が見つかりません: {default_event_dict}")
            self.event_dict_path = None
        
        # item_dictionary.json：本体のみ（config_default）
        default_item_dict = config_default / "item_dictionary.json"
        if default_item_dict.exists():
            self.item_dict_path = default_item_dict
        else:
            print(f"警告: item_dictionary.json が見つかりません: {default_item_dict}")
            self.item_dict_path = None

        # command_dictionary.json：農場ごと（農場フォルダを優先、なければ config_default）
        farm_command_dict = farm_path / "command_dictionary.json"
        default_command_dict = config_default / "command_dictionary.json"
        if farm_command_dict.exists():
            self.command_dict_path = farm_command_dict
        elif default_command_dict.exists():
            self.command_dict_path = default_command_dict
        else:
            self.command_dict_path = None
        self.command_dictionary: Dict[str, str] = {}
        self._load_command_dictionary()

        # PEN設定に合わせて既存データのpen値を正規化
        self._normalize_pen_values()
        
        # 現在選択中の牛
        self.current_cow_auto_id: Optional[int] = None
        self.current_cow_card: Optional[CowCard] = None
        
        # 現在表示中のView
        self.current_view: Optional[tk.Widget] = None
        self.current_view_type: Optional[Literal['chat', 'cow_card', 'list', 'report']] = None
        
        # ビュー管理（再利用のため保持）
        self.views: Dict[str, tk.Widget] = {}
        
        # 個体一覧ウィンドウ参照（既に開いていれば前面に出すため）
        self._cow_list_window: Optional[Any] = None

        # SIRE一覧ウィンドウ参照（二重に開かないように1つに統一）
        self._sire_list_window: Optional[Any] = None

        # レポート作成ダイアログ・ゲノムレポート等（既に開いていれば前面に表示）
        self._report_builder_dialog: Optional[tk.Toplevel] = None
        self._genome_report_window: Optional[Any] = None
        self._dictionary_settings_dialog: Optional[tk.Toplevel] = None
        self._data_output_dialog: Optional[tk.Toplevel] = None
        
        # ChatGPTClientの初期化（エラー時はNoneのまま）
        self.chatgpt_client: Optional[ChatGPTClient] = None
        self.system_prompt: Optional[str] = None
        try:
            self.chatgpt_client = ChatGPTClient()
            # system_promptを読み込む
            system_prompt_path = Path("config/falcon_ai_system_prompt.txt")
            if system_prompt_path.exists():
                self.system_prompt = system_prompt_path.read_text(encoding='utf-8')
            else:
                logging.warning(f"system_promptファイルが見つかりません: {system_prompt_path}")
        except Exception as e:
            logging.warning(f"ChatGPTClientの初期化に失敗しました: {e}")
            # ChatGPTが使えない場合でもアプリは動作する
        
        # 分析ジョブ管理
        self.analysis_running: bool = False
        self.current_job_id: int = 0
        
        # 期間指定状態管理
        self.selected_period: Dict[str, Any] = {
            "start": None,
            "end": None,
            "source": None  # "ui" | "text" | "default"
        }
        
        # 項目辞書をロード
        self.item_dictionary: Dict[str, Any] = {}
        self._load_item_dictionary()
        
        # コマンド入力履歴（頻度計算用）
        self.command_history: Dict[str, int] = {}  # {項目名: 入力回数}
        self.command_history_path = farm_path / "command_history.json"
        
        # コマンド実行履歴（時系列、上矢印キーでさかのぼる用）
        self.command_execution_history: List[str] = []  # 実行したコマンドのリスト
        self.command_history_index: int = -1  # 現在の履歴位置（-1は最新）
        self.command_execution_history_path = farm_path / "command_execution_history.json"  # 実行履歴の保存先
        self._load_command_history()
        self._load_command_execution_history()  # 実行履歴を読み込む
        
        # コマンド入力欄のプレースホルダー関連
        self.command_has_focus: bool = False
        self.command_placeholder_shown: bool = False
        self.command_placeholder: str = "ここにコマンドを入力してください（例：リスト：ID または 集計：平均 DIM）"
        
        # 表のソート状態管理
        self.table_original_rows: List = []  # 元のデータ（ソート前の順序）
        self.table_sort_state: Dict[str, Optional[str]] = {}  # {列名: 'asc'|'desc'|None}
        
        # 現在表示中のコマンド情報
        self.current_table_command: Optional[str] = None  # 表を表示したコマンド
        self.current_graph_command: Optional[str] = None  # グラフを表示したコマンド
        
        # QueryRouter/QueryExecutor/QueryRenderer を初期化
        normalization_dir = Path(__file__).parent.parent.parent / "normalization"
        self.query_router = QueryRouter(
            normalization_dir=normalization_dir,
            item_dictionary_path=self.item_dict_path,
            event_dictionary_path=self.event_dict_path
        )
        if QUERY_V2_AVAILABLE:
            self.query_renderer = QueryRenderer(
                item_dictionary_path=str(self.item_dict_path) if self.item_dict_path else None,
                event_dictionary_path=str(self.event_dict_path) if self.event_dict_path else None
            )
        else:
            self.query_renderer = None
        
        # アプリ設定を読み込んでフォントを適用
        from modules.app_settings_manager import get_app_settings_manager
        self.app_settings = get_app_settings_manager()
        self.default_font_family = self.app_settings.get_font_family()
        self.default_font_size = self.app_settings.get_font_size()
        # メイン画面（コマンド・表・グラフ）は Meiryo UI に統一
        self._main_font_family = "Meiryo UI"
        
        # ttk.Styleを使ってデフォルトフォントを設定（Meiryo UI 統一）
        self.style = ttk.Style()
        self.style.configure('.', font=(self._main_font_family, self.default_font_size))
        self.style.configure('TLabel', font=(self._main_font_family, self.default_font_size))
        self.style.configure('TButton', font=(self._main_font_family, self.default_font_size))
        self.style.configure('TEntry', font=(self._main_font_family, self.default_font_size))
        self.style.configure('TCombobox', font=(self._main_font_family, self.default_font_size))
        self.style.configure('Treeview', font=(self._main_font_family, self.default_font_size))
        self.style.configure('Treeview.Heading', font=(self._main_font_family, self.default_font_size))
        # 結果表（表タブ）のみ MS ゴシック
        self.style.configure('ResultTable.Treeview', font=('MS Gothic', self.default_font_size))
        self.style.configure('ResultTable.Treeview.Heading', font=('MS Gothic', self.default_font_size))
        self.style.configure('TNotebook', font=(self._main_font_family, self.default_font_size))
        self.style.configure('TNotebook.Tab', font=(self._main_font_family, self.default_font_size))
        
        # ウィンドウサイズを設定
        initial_farm_name = self.farm_path.name
        try:
            settings_manager = SettingsManager(self.farm_path)
            initial_farm_name = settings_manager.get("farm_name", self.farm_path.name)
        except Exception:
            pass
        self.root.title(f"FALCON2：{initial_farm_name}")
        self.root.geometry("1200x800")
        
        # UI作成
        self._create_widgets()

        # バックアップ設定に従い、該当する場合は起動時にバックアップを実行（少し遅延してUI表示後）
        self.root.after(500, self._check_and_run_backup)
        # レポート散布図クリックで個体カードを開くキューをポーリング
        self.root.after(600, self._poll_report_cow_open_queue)
    
    def _poll_report_cow_open_queue(self):
        """レポートから届いた cow_id を処理して個体カードを開く"""
        try:
            from modules.report_cow_bridge import get_pending_cow_ids
            for cow_id in get_pending_cow_ids():
                padded = str(cow_id).strip().zfill(4)
                self._jump_to_cow_card(padded)
        except Exception as e:
            logging.debug("report_cow_open_queue: %s", e)
        self.root.after(500, self._poll_report_cow_open_queue)
    
    def _check_and_run_backup(self):
        """バックアップ設定に従い、該当する場合に農場フォルダをバックアップする"""
        try:
            from backup_manager import run_backup_if_due
            backup_dest = run_backup_if_due(self.farm_path)
            if backup_dest is not None:
                messagebox.showinfo(
                    "バックアップ完了",
                    f"今月のバックアップを実行しました。\n{backup_dest}"
                )
        except Exception as e:
            logging.warning(f"バックアップ確認中にエラー: {e}")
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインPanedWindow（左右分割）
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左カラム：サイドメニュー（ダークネイビーでスタイリッシュに） ==========
        _df = "Meiryo UI"
        side_bg = "#1e2a3a"  # ダークネイビー
        side_fg = "#ffffff"  # 白文字
        side_hover_bg = "#2d3e50"  # ホバー時の背景
        side_accent = "#5c9ce6"  # アクセントカラー（淡いブルー）
        try:
            _sm = SettingsManager(self.farm_path)
            _farm_display_name = _sm.get("farm_name", self.farm_path.name)
        except Exception:
            _farm_display_name = self.farm_path.name
        
        left_container = tk.Frame(main_paned, bg=side_bg, width=220)
        left_container.pack_propagate(False)
        main_paned.add(left_container, weight=0)
        
        # サイドメニュー・ヘッダー（FALCONテキストのみ）
        side_header = tk.Frame(left_container, bg=side_bg, pady=18, padx=16)
        side_header.pack(fill=tk.X)
        tk.Label(side_header, text="FALCON", font=(_df, 14, "bold"), bg=side_bg, fg=side_fg).pack(side=tk.LEFT, anchor=tk.W)
        
        # セパレーター
        separator = tk.Frame(left_container, bg="#3a4a5a", height=1)
        separator.pack(fill=tk.X, padx=16)
        
        # メニューボタン（ダーク背景・白文字）
        menu_buttons = [
            ("🐄 個体カード", self._on_cow_card),
            ("🩺 繁殖検診", self._on_reproduction_checkup),
            ("📝 イベント入力", self._on_event_input),
            ("📋 レポート作成", self._on_report_builder),
            ("📊 ダッシュボード", self._on_dashboard),
            ("📥 データ吸い込み", self._on_milk_test_import),
            ("📤 データ出力", self._on_data_output),
            ("📚 辞書・設定", self._on_dictionary_settings),
            ("🏡 農場管理", self._on_farm_management),
            ("⚙️ アプリ設定", self._on_app_settings),
        ]
        
        menu_inner = tk.Frame(left_container, bg=side_bg, padx=8, pady=12)
        menu_inner.pack(fill=tk.BOTH, expand=True)
        
        def on_enter(e, btn):
            btn.config(bg=side_hover_bg)
        def on_leave(e, btn):
            btn.config(bg=side_bg)
        
        for text, command in menu_buttons:
            btn = tk.Button(
                menu_inner, text=text, command=command,
                font=(_df, 11), bg=side_bg, fg=side_fg,
                activebackground=side_hover_bg, activeforeground=side_fg,
                relief=tk.FLAT, bd=0, highlightthickness=0,
                cursor="hand2", anchor=tk.W, padx=14, pady=10
            )
            btn.pack(fill=tk.X, pady=2)
            btn.bind("<Enter>", lambda e, b=btn: on_enter(e, b))
            btn.bind("<Leave>", lambda e, b=btn: on_leave(e, b))
        
        # 従来の SideMenu.TButton を参照している他箇所用にスタイルは維持（ttk は右カラムで使用）
        self.style.configure('SideMenu.TButton', padding=(8, 10))
        
        # ========== 右カラム：メイン表示領域（クリーンな白背景） ==========
        _df = "Meiryo UI"
        bg = "#ffffff"  # 白背景
        right_container = tk.Frame(main_paned, bg=bg)
        main_paned.add(right_container, weight=1)
        
        # ヘッダー（農場名を表示・サイドバーと同じ濃い藍色）
        header = tk.Frame(right_container, bg=side_bg, pady=12, padx=24)
        header.pack(fill=tk.X)
        title_frame = tk.Frame(header, bg=side_bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._main_farm_name_label = tk.Label(title_frame, text=_farm_display_name, font=(_df, 16, "bold"), bg=side_bg, fg=side_fg)
        self._main_farm_name_label.pack(anchor=tk.W)
        tk.Label(title_frame, text="コマンドを入力してリスト・集計・グラフを表示します", font=(_df, 10), bg=side_bg, fg="#90a4ae").pack(anchor=tk.W)
        
        right_frame = ttk.Frame(right_container)
        right_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 10))
        
        # 右上：コマンド入力欄
        command_frame = ttk.Frame(right_frame)
        command_frame.pack(fill=tk.X, padx=0, pady=5)
        
        # コマンドタイプ選択（Combobox）
        command_type_frame = ttk.Frame(command_frame)
        command_type_frame.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(command_type_frame, text="タイプ:", font=(self._main_font_family, self.default_font_size)).pack(side=tk.LEFT, padx=(0, 3))
        self.command_type_var = tk.StringVar(value="")
        command_type_combo = ttk.Combobox(
            command_type_frame,
            textvariable=self.command_type_var,
            values=["", "受胎率（経産）", "リスト", "イベントリスト", "イベント集計", "集計（経産牛）", "グラフ", "項目一括変更"],
            width=12,
            state="readonly",
            font=(self._main_font_family, self.default_font_size)
        )
        command_type_combo.pack(side=tk.LEFT)
        command_type_combo.bind("<<ComboboxSelected>>", self._on_command_type_selected)
        
        ttk.Label(command_frame, text="コマンド:", font=(self._main_font_family, self.default_font_size)).pack(side=tk.LEFT, padx=(10, 5))
        # コマンド入力欄をComboboxに変更（プルダウンで履歴を表示）
        self.command_entry = ttk.Combobox(command_frame, width=50, state="normal")
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        # 履歴をプルダウンに設定（最新が上に来るように逆順）
        self._update_command_history_dropdown()
        self.command_entry.bind('<Return>', self._on_command_enter)
        self.command_entry.bind('<KeyPress>', self._on_command_key_press_entry)
        self.command_entry.bind('<FocusIn>', self._on_command_focus_in_entry)
        self.command_entry.bind('<Down>', self._on_command_dropdown_open)  # 下矢印キーでプルダウンを開く
        # プルダウンから選択された時の処理
        self.command_entry.bind('<<ComboboxSelected>>', self._on_command_history_selected)
        
        command_btn = ttk.Button(command_frame, text="実行", command=self._on_command_execute)
        command_btn.pack(side=tk.LEFT, padx=5)
        
        # 項目一覧を表示ボタン
        item_list_btn = ttk.Button(command_frame, text="項目一覧を表示", command=self._on_show_item_list)
        item_list_btn.pack(side=tk.LEFT, padx=5)
        
        # イベント選択フレーム（イベントリストタイプ選択時のみ表示）
        self.event_selection_frame = ttk.Frame(right_frame)
        # 初期状態では非表示
        
        # イベント選択ラベル
        self.event_selection_label = ttk.Label(self.event_selection_frame, text="イベント選択:", font=(self._main_font_family, self.default_font_size))
        self.event_selection_label.pack(side=tk.LEFT, padx=(10, 5))
        
        # イベント選択用のコンボボックス（プルダウン）
        # イベント辞書からイベント一覧を取得
        event_dict_path = Path(__file__).parent.parent.parent / "config_default" / "event_dictionary.json"
        event_options = []
        self.event_number_to_name = {}
        if event_dict_path.exists():
            with open(event_dict_path, 'r', encoding='utf-8') as f:
                event_dict_data = json.load(f)
                for event_num_str, event_data in event_dict_data.items():
                    # deprecatedでないイベントのみ追加
                    if not event_data.get('deprecated', False):
                        event_number = int(event_num_str)
                        event_name = event_data.get('name_jp', event_data.get('alias', str(event_number)))
                        event_options.append(f"{event_number}: {event_name}")
                        self.event_number_to_name[event_number] = event_name
        
        # イベント選択用のチェックボックスリスト（プルダウン風）
        # ボタンで開閉できるフレームを作成
        self.event_selection_button_frame = ttk.Frame(self.event_selection_frame)
        self.event_selection_button_frame.pack(side=tk.LEFT, padx=5)
        
        self.event_selection_button = ttk.Button(
            self.event_selection_button_frame,
            text="イベントを選択",
            command=self._toggle_event_selection_popup,
            width=15
        )
        self.event_selection_button.pack(side=tk.LEFT)
        
        # 選択されたイベント数を表示するラベル
        self.event_selection_count_label = ttk.Label(
            self.event_selection_button_frame,
            text="(未選択)",
            foreground="gray"
        )
        self.event_selection_count_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # 部分一致入力欄
        self.note_partial_match_label = ttk.Label(
            self.event_selection_frame,
            text="部分一致:",
            font=(self._main_font_family, self.default_font_size)
        )
        self.note_partial_match_label.pack(side=tk.LEFT, padx=(10, 5))
        
        self.note_partial_match_entry = ttk.Entry(self.event_selection_frame, width=16)
        self.note_partial_match_entry.pack(side=tk.LEFT, padx=5)
        self.note_partial_match_entry.bind('<Return>', self._on_condition_enter)  # エンターキーで実行
        
        # イベント選択用のポップアップウィンドウ（初期状態では非表示）
        self.event_selection_popup = None
        self.event_selection_checkboxes = {}  # {event_number: checkbox_var}
        self.selected_event_numbers = set()  # 選択されたイベント番号のセット
        
        # 条件入力欄（リストタイプ・集計タイプ選択時のみ表示）
        self.condition_frame = ttk.Frame(right_frame)
        # 初期状態では非表示
        # self.condition_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 分類入力欄（集計タイプ選択時のみ表示）- プルダウンに変更
        self.classification_label = ttk.Label(self.condition_frame, text="分類:", font=(self._main_font_family, self.default_font_size))
        self.classification_label.pack(side=tk.LEFT, padx=(10, 5))
        # 分類パターンの定義
        self.classification_options = [
            "",  # 空欄（分類なし、単純平均用）
            "産次", "産次（１産、２産、３産以上）", "DIM7", "DIM14", "DIM21", "DIM30", "DIM50", 
            "DIM：産次", "空胎日数100日（DOPN=100）", "空胎日数150日（DOPN=150）", "分娩月"
        ]
        self.classification_entry = ttk.Combobox(self.condition_frame, width=14, values=self.classification_options, state="readonly")
        self.classification_entry.pack(side=tk.LEFT, padx=5)
        self.classification_entry.set("")  # 初期値は空
        self.classification_entry.bind('<Return>', self._on_condition_enter)  # エンターキーで実行
        self.classification_placeholder = ""
        self.classification_placeholder_shown = False
        
        # 条件入力欄
        self.condition_label = ttk.Label(self.condition_frame, text="条件:", font=(self._main_font_family, self.default_font_size))
        self.condition_label.pack(side=tk.LEFT, padx=(10, 5))
        self.condition_entry = ttk.Entry(self.condition_frame, width=50)
        self.condition_entry.pack(side=tk.LEFT, padx=5)
        self.condition_entry.insert(0, "例）DIM<150 または 産次：初産")
        self.condition_entry.config(foreground="gray")
        self.condition_entry.bind('<FocusIn>', self._on_condition_focus_in)
        self.condition_entry.bind('<FocusOut>', self._on_condition_focus_out)
        self.condition_entry.bind('<KeyPress>', self._on_condition_key_press)
        self.condition_entry.bind('<Return>', self._on_condition_enter)  # エンターキーで実行
        self.condition_has_focus = False
        self.condition_placeholder = "例）DIM<150 または 産次：初産"
        self.condition_placeholder_shown = True

        # 内訳入力欄（集計タイプ選択時のみ表示）
        self.breakdown_label = ttk.Label(self.condition_frame, text="内訳:", font=(self._main_font_family, self.default_font_size))
        self.breakdown_label.pack(side=tk.LEFT, padx=(10, 5))
        self.breakdown_entry = ttk.Entry(self.condition_frame, width=12)
        self.breakdown_entry.pack(side=tk.LEFT, padx=5)
        self.breakdown_entry.insert(0, "例）150, 200")
        self.breakdown_entry.config(foreground="gray")
        self.breakdown_entry.bind('<FocusIn>', self._on_breakdown_focus_in)
        self.breakdown_entry.bind('<FocusOut>', self._on_breakdown_focus_out)
        self.breakdown_entry.bind('<KeyPress>', self._on_breakdown_key_press)
        self.breakdown_entry.bind('<Return>', self._on_condition_enter)
        self.breakdown_placeholder = "例）150, 200"
        self.breakdown_placeholder_shown = True
        
        # 並び順（リストタイプ時のみ表示）
        self.sort_label = ttk.Label(self.condition_frame, text="並び順:", font=(self._main_font_family, self.default_font_size))
        self.sort_column_entry = ttk.Entry(self.condition_frame, width=10)
        self.sort_order_var = tk.StringVar(value="昇順")
        self.sort_order_combo = ttk.Combobox(
            self.condition_frame, width=14,
            values=["昇順", "降順"],
            state="readonly", textvariable=self.sort_order_var
        )
        self.sort_order_combo.set("昇順")
        
        # 受胎率フレーム（受胎率タイプ選択時のみ表示）
        self.conception_rate_frame = ttk.Frame(right_frame)
        # 初期状態では非表示
        
        # 受胎率の種類選択（プルダウン）
        self.conception_rate_type_label = ttk.Label(self.conception_rate_frame, text="受胎率の種類:", font=(self._main_font_family, self.default_font_size))
        self.conception_rate_type_label.pack(side=tk.LEFT, padx=(10, 5))
        conception_rate_type_options = ["月", "産次", "授精回数", "DIMサイクル", "授精種類", "授精間隔", "授精師", "SIRE"]
        self.conception_rate_type_var = tk.StringVar(value="")
        self.conception_rate_type_entry = ttk.Combobox(self.conception_rate_frame, width=14, values=conception_rate_type_options, state="readonly", textvariable=self.conception_rate_type_var)
        # 選択肢を明示的に設定（念のため）
        self.conception_rate_type_entry['values'] = conception_rate_type_options
        self.conception_rate_type_entry.pack(side=tk.LEFT, padx=5)
        self.conception_rate_type_entry.set("")  # 初期値は空
        self.conception_rate_type_entry.bind('<Return>', self._on_condition_enter)  # エンターキーで実行
        
        # 受胎率用の条件入力欄
        self.conception_rate_condition_label = ttk.Label(self.conception_rate_frame, text="条件:", font=(self._main_font_family, self.default_font_size))
        self.conception_rate_condition_label.pack(side=tk.LEFT, padx=(10, 5))
        self.conception_rate_condition_entry = ttk.Entry(self.conception_rate_frame, width=50)
        self.conception_rate_condition_entry.pack(side=tk.LEFT, padx=5)
        self.conception_rate_condition_entry.insert(0, "例）DIM<150 または 産次：初産")
        self.conception_rate_condition_entry.config(foreground="gray")
        self.conception_rate_condition_entry.bind('<FocusIn>', self._on_conception_rate_condition_focus_in)
        self.conception_rate_condition_entry.bind('<FocusOut>', self._on_conception_rate_condition_focus_out)
        self.conception_rate_condition_entry.bind('<KeyPress>', self._on_conception_rate_condition_key_press)
        self.conception_rate_condition_entry.bind('<Return>', self._on_condition_enter)  # エンターキーで実行
        self.conception_rate_condition_placeholder = "例）DIM<150 または 産次：初産"
        self.conception_rate_condition_placeholder_shown = True
        
        # グラフ入力フレーム（グラフタイプ選択時のみ表示）
        self.graph_input_frame = ttk.Frame(right_frame)
        # 初期状態では非表示
        
        # グラフ種類を最初に配置
        self.graph_type_label = ttk.Label(self.graph_input_frame, text="グラフ種類:", font=(self._main_font_family, self.default_font_size))
        self.graph_type_label.pack(side=tk.LEFT, padx=(10, 5))
        graph_type_options = ["散布図", "空胎日数生存曲線"]
        self.graph_type_entry = ttk.Combobox(self.graph_input_frame, width=10, values=graph_type_options, state="readonly")
        self.graph_type_entry.pack(side=tk.LEFT, padx=5)
        self.graph_type_entry.set("")
        self.graph_type_entry.bind("<<ComboboxSelected>>", self._on_graph_type_selected)
        
        # Y軸（初期状態では非表示）
        self.graph_y_label = ttk.Label(self.graph_input_frame, text="Y軸:", font=(self._main_font_family, self.default_font_size))
        self.graph_y_entry = ttk.Entry(self.graph_input_frame, width=12)
        # Y軸用の項目一覧ボタン（初期状態では非表示）
        self.graph_y_item_list_btn = ttk.Button(self.graph_input_frame, text="項目一覧を表示", command=self._on_show_item_list_for_graph_y)
        
        # X軸（初期状態では非表示）
        self.graph_x_label = ttk.Label(self.graph_input_frame, text="X軸:", font=(self._main_font_family, self.default_font_size))
        self.graph_x_entry = ttk.Entry(self.graph_input_frame, width=12)
        # 対象項目用の項目一覧ボタン（初期状態では非表示）
        self.graph_x_item_list_btn = ttk.Button(self.graph_input_frame, text="項目一覧を表示", command=self._on_show_item_list_for_graph_x)
        
        # 分類（初期状態では非表示、Y軸・X軸の後に表示）
        # グラフ専用の分類オプション（DIM：産次は除外）
        self.graph_classification_options = [
            "",  # 空欄（分類なし、単純平均用）
            "産次", "産次（１産、２産、３産以上）", "イベントの有無",
            "DIM7", "DIM14", "DIM21", "DIM30", "DIM50", 
            "空胎日数100日（DOPN=100）", "空胎日数150日（DOPN=150）", "分娩月"
        ]
        # 空胎日数生存曲線用：産次・分娩月・項目・条件（DIM帯は廃止）
        self.graph_classification_options_survival = [
            "", "産次", "産次（１産、２産、３産以上）",
            "分娩月", "項目で分類", "条件で分類"
        ]
        self.graph_classification_label = ttk.Label(self.graph_input_frame, text="分類:", font=(self._main_font_family, self.default_font_size))
        self.graph_classification_entry = ttk.Combobox(self.graph_input_frame, width=14, values=self.graph_classification_options, state="readonly")
        self.graph_classification_entry.set("")
        self.graph_classification_entry.bind('<Return>', self._on_condition_enter)
        self.graph_classification_entry.bind('<<ComboboxSelected>>', self._on_graph_classification_selected)
        # 生存曲線「イベントの有無」用
        self.graph_classification_event_label = ttk.Label(self.graph_input_frame, text="対象イベント:", font=(self._main_font_family, self.default_font_size))
        self.graph_classification_event_entry = ttk.Combobox(self.graph_input_frame, width=18, state="readonly")
        self.graph_classification_event_entry.set("")
        # 生存曲線「項目で分類」用（対象項目・閾値）
        self.graph_classification_item_label = ttk.Label(self.graph_input_frame, text="対象項目:", font=(self._main_font_family, self.default_font_size))
        self.graph_classification_item_entry = ttk.Combobox(self.graph_input_frame, width=14, state="readonly")
        self.graph_classification_item_entry.set("")
        self.graph_classification_item_list_btn = ttk.Button(self.graph_input_frame, text="項目一覧を表示", command=self._on_show_item_list_for_survival_classification_item)
        self.graph_classification_bin_label = ttk.Label(self.graph_input_frame, text="閾値:", font=(self._main_font_family, self.default_font_size))
        self.graph_classification_bin_entry = ttk.Entry(self.graph_input_frame, width=10)
        self.graph_classification_bin_entry.insert(0, "例) 0.1")
        self.graph_classification_bin_entry.config(foreground="gray")
        self.graph_classification_bin_placeholder = "例) 0.1"
        self.graph_classification_bin_placeholder_shown = True
        self.graph_classification_bin_entry.bind('<FocusIn>', self._on_graph_bin_focus_in)
        self.graph_classification_bin_entry.bind('<FocusOut>', self._on_graph_bin_focus_out)
        
        # 条件（初期状態では非表示、分類の後に表示）
        self.graph_condition_label = ttk.Label(self.graph_input_frame, text="条件:", font=(self._main_font_family, self.default_font_size))
        self.graph_condition_entry = ttk.Entry(self.graph_input_frame, width=50)
        self.graph_condition_entry.insert(0, "例）DIM<150 または 産次：初産")
        self.graph_condition_entry.config(foreground="gray")
        self.graph_condition_entry.bind('<FocusIn>', self._on_graph_condition_focus_in)
        self.graph_condition_entry.bind('<FocusOut>', self._on_graph_condition_focus_out)
        self.graph_condition_entry.bind('<KeyPress>', self._on_graph_condition_key_press)
        self.graph_condition_entry.bind('<Return>', self._on_condition_enter)
        self.graph_condition_placeholder = "例）DIM<150 または 産次：初産"
        self.graph_condition_placeholder_shown = True
        
        # 期間設定フレーム（イベントリストタイプ選択時のみ表示）
        self.period_frame = ttk.Frame(right_frame)
        # 初期状態では非表示
        
        # 開始日
        self.start_date_label = ttk.Label(self.period_frame, text="開始日:", font=(self._main_font_family, self.default_font_size))
        self.start_date_label.pack(side=tk.LEFT, padx=(10, 5))
        self.start_date_entry = ttk.Entry(self.period_frame, width=12)
        self.start_date_entry.pack(side=tk.LEFT, padx=5)
        self.start_date_entry.insert(0, "YYYY-MM-DD")
        self.start_date_entry.config(foreground="gray")
        self.start_date_placeholder = "YYYY-MM-DD"
        self.start_date_placeholder_shown = True
        self.start_date_entry.bind('<FocusIn>', lambda e: self._on_date_entry_focus_in(self.start_date_entry, self.start_date_placeholder))
        self.start_date_entry.bind('<FocusOut>', lambda e: self._on_date_entry_focus_out(self.start_date_entry, self.start_date_placeholder))
        
        # 開始日カレンダーボタン
        self.start_date_calendar_btn = ttk.Button(
            self.period_frame,
            text="📅",
            command=lambda: self._show_calendar_dialog(self.start_date_entry),
            width=3
        )
        self.start_date_calendar_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 終了日
        self.end_date_label = ttk.Label(self.period_frame, text="終了日:", font=(self._main_font_family, self.default_font_size))
        self.end_date_label.pack(side=tk.LEFT, padx=(10, 5))
        self.end_date_entry = ttk.Entry(self.period_frame, width=12)
        self.end_date_entry.pack(side=tk.LEFT, padx=5)
        self.end_date_entry.insert(0, "YYYY-MM-DD")
        self.end_date_entry.config(foreground="gray")
        self.end_date_placeholder = "YYYY-MM-DD"
        self.end_date_placeholder_shown = True
        self.end_date_entry.bind('<FocusIn>', lambda e: self._on_date_entry_focus_in(self.end_date_entry, self.end_date_placeholder))
        self.end_date_entry.bind('<FocusOut>', lambda e: self._on_date_entry_focus_out(self.end_date_entry, self.end_date_placeholder))
        
        # 終了日カレンダーボタン
        self.end_date_calendar_btn = ttk.Button(
            self.period_frame,
            text="📅",
            command=lambda: self._show_calendar_dialog(self.end_date_entry),
            width=3
        )
        self.end_date_calendar_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 期間リセットボタン
        self.period_reset_btn = ttk.Button(
            self.period_frame,
            text="期間リセット",
            command=self._reset_period_entries,
            width=10
        )
        self.period_reset_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # コマンド入力欄の直下：最終日付表示
        self.latest_dates_frame = ttk.Frame(right_frame)
        self.latest_dates_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        self.latest_dates_label = ttk.Label(
            self.latest_dates_frame,
            text="最終分娩：—　最終AI/ET：—　最終乳検：—　最終イベント：—",
            foreground="black"
        )
        self.latest_dates_label.pack(side=tk.LEFT, padx=5)
        
        # コマンド結果表示エリア（表・グラフタブ）
        result_notebook = ttk.Notebook(right_frame)
        result_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 表タブ
        table_frame = ttk.Frame(result_notebook)
        result_notebook.add(table_frame, text="表")
        
        # 表タブの上部：コマンド表示とボタンエリア
        table_header_frame = ttk.Frame(table_frame)
        table_header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # コマンド表示テキスト（選択・コピー可能）
        # ttk.Frameの背景色を取得（システムのデフォルト色を使用）
        try:
            style = ttk.Style()
            bg_color = style.lookup("TFrame", "background")
            if not bg_color:
                bg_color = "SystemButtonFace"  # Windowsのデフォルト背景色
        except:
            bg_color = "SystemButtonFace"
        
        self.table_command_text = tk.Text(
            table_header_frame,
            height=1,
            font=(self._main_font_family, 9),
            foreground="black",
            wrap=tk.NONE,
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            state=tk.DISABLED,
            cursor="ibeam",
            bg=bg_color
        )
        self.table_command_text.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # ボタンエリア（初期状態では非表示）
        table_button_frame = ttk.Frame(table_header_frame)
        table_button_frame.pack(side=tk.RIGHT, padx=5)
        
        # エクスポートボタン
        self.table_export_btn = ttk.Button(
            table_button_frame,
            text="エクセルにエクスポート",
            command=self._on_export_table_to_excel,
            state=tk.DISABLED
        )
        self.table_export_btn.pack(side=tk.LEFT, padx=2)
        
        # プリントボタン
        self.table_print_btn = ttk.Button(
            table_button_frame,
            text="プリント",
            command=self._on_print_table,
            state=tk.DISABLED
        )
        self.table_print_btn.pack(side=tk.LEFT, padx=2)
        
        # 表用のTreeview（グリッド形式）
        table_container = ttk.Frame(table_frame)
        table_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 横スクロールバー用のフレーム
        table_h_scroll_frame = ttk.Frame(table_container)
        table_h_scroll_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        table_scrollbar_x = ttk.Scrollbar(table_h_scroll_frame, orient=tk.HORIZONTAL)
        table_scrollbar_x.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 縦スクロールバーとTreeview用のフレーム
        table_v_scroll_frame = ttk.Frame(table_container)
        table_v_scroll_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        table_scrollbar_y = ttk.Scrollbar(table_v_scroll_frame, orient=tk.VERTICAL)
        table_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.result_treeview = ttk.Treeview(
            table_v_scroll_frame,
            yscrollcommand=table_scrollbar_y.set,
            xscrollcommand=table_scrollbar_x.set,
            show='headings',
            style='ResultTable.Treeview'
        )
        self.result_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        table_scrollbar_y.config(command=self.result_treeview.yview)
        table_scrollbar_x.config(command=self.result_treeview.xview)
        
        # グラフタブ
        graph_frame = ttk.Frame(result_notebook)
        result_notebook.add(graph_frame, text="グラフ")
        
        # グラフタブの上部：コマンド表示エリア
        graph_header_frame = ttk.Frame(graph_frame)
        graph_header_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # コマンド表示テキスト（選択・コピー可能）
        self.graph_command_text = tk.Text(
            graph_header_frame,
            height=1,
            font=(self._main_font_family, 9),
            foreground="black",
            wrap=tk.NONE,
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            state=tk.DISABLED,
            cursor="ibeam"
        )
        self.graph_command_text.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # グラフ表示エリア（matplotlib）
        if MATPLOTLIB_AVAILABLE:
            self.graph_figure = Figure(figsize=(10, 6), dpi=100)
            self.graph_canvas = FigureCanvasTkAgg(self.graph_figure, graph_frame)
            self.graph_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        else:
            self.graph_figure = None
            self.graph_canvas = None
            no_graph_label = tk.Label(
                graph_frame,
                text="matplotlib がインストールされていません。グラフ機能は使用できません。",
                foreground="red"
            )
            no_graph_label.pack(expand=True)
        
        self.graph_frame = graph_frame
        
        # 操作ログタブ（イベント登録・更新・削除の履歴、config/activity_log.json）
        activity_log_frame = ttk.Frame(result_notebook)
        result_notebook.add(activity_log_frame, text="操作ログ")
        al_top = ttk.Frame(activity_log_frame)
        al_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            al_top,
            text="直近のイベント登録・更新・削除（このPCの記録）。コマンド実行とは独立しています。",
            font=(self._main_font_family, 9),
            foreground="#546e7a",
        ).pack(side=tk.LEFT)
        ttk.Button(
            al_top,
            text="更新",
            command=self._refresh_main_activity_log,
            width=8,
        ).pack(side=tk.RIGHT)
        
        al_tree_frame = ttk.Frame(activity_log_frame)
        al_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        al_columns = ("ts", "farm", "action", "cow_id", "cow_auto_id", "event", "event_date", "event_id")
        self.activity_log_tree = ttk.Treeview(
            al_tree_frame,
            columns=al_columns,
            show="headings",
            height=18,
        )
        self.activity_log_tree.heading("ts", text="日時")
        self.activity_log_tree.heading("farm", text="農場")
        self.activity_log_tree.heading("action", text="種別")
        self.activity_log_tree.heading("cow_id", text="個体ID")
        self.activity_log_tree.heading("cow_auto_id", text="auto_id")
        self.activity_log_tree.heading("event", text="イベント")
        self.activity_log_tree.heading("event_date", text="イベント日")
        self.activity_log_tree.heading("event_id", text="event_id")
        
        self.activity_log_tree.column("ts", width=150, anchor="w")
        self.activity_log_tree.column("farm", width=90, anchor="w")
        self.activity_log_tree.column("action", width=48, anchor="center")
        self.activity_log_tree.column("cow_id", width=64, anchor="w")
        self.activity_log_tree.column("cow_auto_id", width=56, anchor="e")
        self.activity_log_tree.column("event", width=160, anchor="w")
        self.activity_log_tree.column("event_date", width=88, anchor="w")
        self.activity_log_tree.column("event_id", width=64, anchor="e")
        
        al_vsb = ttk.Scrollbar(al_tree_frame, orient=tk.VERTICAL, command=self.activity_log_tree.yview)
        al_hsb = ttk.Scrollbar(activity_log_frame, orient=tk.HORIZONTAL, command=self.activity_log_tree.xview)
        self.activity_log_tree.configure(yscrollcommand=al_vsb.set, xscrollcommand=al_hsb.set)
        self.activity_log_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        al_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        al_hsb.pack(fill=tk.X)
        
        result_notebook.bind("<<NotebookTabChanged>>", self._on_main_result_notebook_tab_changed)
        self._refresh_main_activity_log()
        
        self.result_notebook = result_notebook
        
        # 右側：メイン表示領域（ChatGPT / View切替）- 結果表示エリアの下に配置
        self.main_content_frame = ttk.Frame(right_frame)
        # 初期表示時は非表示（表・グラフタブのエリアを拡張するため）
        # self.main_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 初期表示：ChatGPT画面（非表示のまま）
        # self._show_chat_view()
        
        # 農場全体の最終日付を計算して表示
        self._calculate_and_update_farm_latest_dates()
    
    
    def show_view(self, view_widget: tk.Widget, view_type: Literal['chat', 'cow_card', 'list', 'report']):
        """
        Viewを表示（切替機構）
        
        Args:
            view_widget: 表示するウィジェット
            view_type: Viewの種類
        """
        # デバッグログ：親フレームの存在確認
        try:
            parent_exists = self.main_content_frame.winfo_exists()
            logging.debug(f"show_view: parent exists={parent_exists}, view_type={view_type}")
        except tk.TclError as e:
            logging.error(f"show_view: parent check failed: {e}")
            return
        
        # 既存のビューを pack_forget() で非表示にする（destroy しない）
        for v in self.views.values():
            try:
                if v.winfo_exists():
                    v.pack_forget()
            except tk.TclError:
                # 既に破棄されている場合は無視
                pass
        
        # 新しいViewを配置
        try:
            view_widget.pack(fill=tk.BOTH, expand=True)
        except tk.TclError as e:
            logging.error(f"show_view: pack failed: {e}, view_type={view_type}")
            return
        
        # ビューを辞書に保存（再利用のため）
        self.views[view_type] = view_widget
        
        # 現在のViewを記録
        self.current_view = view_widget
        self.current_view_type = view_type
    
    def _show_chat_view(self):
        """ChatGPT画面を表示"""
        # 既存のchat_viewがあれば再利用、なければ新規作成
        if 'chat' in self.views:
            chat_frame = self.views['chat']
        else:
            # 新しいchat_frameを作成（parent は main_content_frame）
            chat_frame = ttk.Frame(self.main_content_frame)
            
            # Canvas + Scrollbar でチャット履歴表示エリアを作成
            # CanvasとScrollbarを配置するフレーム
            canvas_container = ttk.Frame(chat_frame)
            canvas_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Scrollbar
            scrollbar = ttk.Scrollbar(canvas_container, orient=tk.VERTICAL)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Canvas
            canvas = tk.Canvas(
                canvas_container,
                bg="#FFFFFF",
                yscrollcommand=scrollbar.set,
                highlightthickness=0
            )
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=canvas.yview)
            
            # メッセージを配置するフレーム（Canvas内に配置）
            messages_frame = ttk.Frame(canvas)
            canvas_window = canvas.create_window(
                (0, 0),
                window=messages_frame,
                anchor=tk.NW
            )
            
            # Canvasのサイズ変更時にmessages_frameの幅を調整
            def configure_canvas_width(event):
                canvas_width = event.width
                canvas.itemconfig(canvas_window, width=canvas_width)
                # messages_frameの各メッセージカードの幅も調整
                for widget in messages_frame.winfo_children():
                    if isinstance(widget, tk.Frame):
                        widget.config(width=canvas_width)
            
            canvas.bind('<Configure>', configure_canvas_width)
            
            # messages_frameのサイズ変更時にCanvasのスクロール領域を更新
            def configure_messages_frame(event):
                canvas.update_idletasks()
                canvas.config(scrollregion=canvas.bbox("all"))
            
            messages_frame.bind('<Configure>', configure_messages_frame)
            
            # マウスホイールでスクロール（Windows用）
            def on_mousewheel(event):
                # Windowsでは event.delta が使用可能
                if event.delta:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                # Linux/Macでは event.num を使用
                elif event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")
            
            # Windows用のマウスホイールイベント
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            # Linux用のマウスホイールイベント
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)
            
            # Canvas内のマウスホイールも処理
            def on_canvas_mousewheel(event):
                canvas.focus_set()
                on_mousewheel(event)
            
            canvas.bind("<MouseWheel>", on_canvas_mousewheel)
            
            # チャット関連のUI要素を保存
            self.chat_canvas = canvas
            self.chat_messages_frame = messages_frame
            self.chat_scrollbar = scrollbar
            self.chat_canvas_window = canvas_window
            
            # 散布図表示用のFrame（初期状態では非表示）
            self.scatter_plot_frame = ttk.Frame(chat_frame)
            # 初期状態ではpackしない（散布図が表示される時にpackする）
        
        # View切替（show_view 内で既存のViewを pack_forget）
        self.show_view(chat_frame, 'chat')
        
        # 初期メッセージは表示しない（日本語クエリUIとして即座に使えるように）
        if not hasattr(self, '_chat_initialized'):
            self._chat_initialized = True
    
    def _on_cow_card(self):
        """個体カードメニューをクリック（個体一覧ウィンドウを開く／既に開いていれば前面に出す）"""
        from ui.cow_list_window import CowListWindow
        
        # 既に個体一覧が開いていれば前面に出す
        if self._cow_list_window is not None and self._cow_list_window.window.winfo_exists():
            self._cow_list_window.window.lift()
            self._cow_list_window.window.focus_set()
            return
        # 新規に個体一覧ウィンドウを開く
        def _clear_cow_list_ref():
            self._cow_list_window = None
        list_window = CowListWindow(
            parent=self.root,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=self.item_dict_path,
            on_close=_clear_cow_list_ref
        )
        self._cow_list_window = list_window
        list_window.show()
    
    def _show_cow_card_view(self, cow_auto_id: int):
        """
        個体カードViewを表示
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 既存のCowCardがあれば再利用、なければ新規作成
        if 'cow_card' in self.views and self.current_cow_card:
            cow_card = self.current_cow_card
            cow_card_widget = cow_card.get_widget()
        else:
            # CowCard を作成（parent は必ず main_content_frame）
            cow_card = CowCard(
                parent=self.main_content_frame,
                db_handler=self.db,
                formula_engine=self.formula_engine,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path
            )
            cow_card_widget = cow_card.get_widget()
            # イベント保存時のコールバックを設定
            cow_card.set_on_event_saved(self._on_event_saved)
            # 現在の牛を記録
            self.current_cow_card = cow_card
        
        # 牛の情報を読み込んで表示（既存のCowCardでも再読み込み）
        cow_card.load_cow(cow_auto_id)
        
        # View切替（show_view 内で既存のViewを pack_forget）
        self.show_view(cow_card_widget, 'cow_card')
        
        # 現在の牛を記録
        self.current_cow_auto_id = cow_auto_id
        
        # 群全体の最終日付を再計算して表示を更新（個体カード表示時も更新）
        self._calculate_and_update_farm_latest_dates()
    
    def _on_reproduction_checkup(self):
        """繁殖検診メニューをクリック"""
        from ui.reproduction_checkup_flow_window import ReproductionCheckupFlowWindow
        
        repro_flow_window = ReproductionCheckupFlowWindow(
            parent=self.root,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path,
            event_dictionary_path=self.event_dict_path
        )
        repro_flow_window.show()
    
    def _on_event_input(self):
        """
        イベント入力メニューをクリック
        
        【重要】左メニューからのイベント入力は常に汎用モード（ID入力から開始）
        メイン画面の状態（個体カード表示有無）に依存しない
        """
        # event_dictionary_path が None の場合はエラー
        if self.event_dict_path is None:
            messagebox.showerror(
                "エラー",
                "event_dictionary.json が見つかりません"
            )
            return
        
        # EventInputWindow を生成または再利用して表示
        # 左メニューからの起動は常に cow_auto_id=None で汎用モード
        # 必ずID入力欄を表示状態で起動する
        event_input_window = EventInputWindow.open_or_focus(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 常にNone（汎用モード、ID入力から開始）
            on_saved=self._on_event_saved,
            farm_path=self.farm_path
        )
        event_input_window.show()
    
    def _on_milk_test_import(self):
        """データ吸い込みメニューをクリック"""
        # データ吸い込みタイプ選択ダイアログを表示
        self._show_data_import_dialog()
    
    def _show_data_import_dialog(self):
        """データ吸い込みタイプ選択ダイアログ（デザイン統一・スマート表示）"""
        _df = "Meiryo UI"
        dialog = tk.Toplevel(self.root)
        dialog.title("データ吸い込み")
        # 閉じるボタンまで余裕を持ってすべてが表示されるように高さを少し大きめに確保
        dialog.geometry("640x680")
        dialog.minsize(620, 640)
        dialog.configure(bg="#f5f5f5")

        # ttk スタイル（背景色は環境で白抜きになるためフォント・余白のみ指定し、実行も一括削除も同じ見た目で文字を確実に表示）
        style = ttk.Style()
        style.configure("DataImport.Exec.TButton", font=(_df, 10), padding=(14, 8))
        style.configure("DataImport.Secondary.TButton", font=(_df, 10), padding=(14, 8))

        main = ttk.Frame(dialog, padding=24)
        main.pack(fill=tk.BOTH, expand=True)

        # ヘッダー（タイトル・説明をすっきり）
        header_frame = ttk.Frame(main)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header_frame, text="データ吸い込み", font=(_df, 16, "bold")).pack(anchor=tk.W)
        ttk.Label(header_frame, text="データ吸い込みタイプを選択し、実行または一括削除を押してください。", font=(_df, 9)).pack(anchor=tk.W)

        rows = [
            ("AI", "人工授精（AI）データを読み込みます", "AI", "AIキャンセル"),
            ("乳検", "乳検データを読み込みます", "乳検", "乳検キャンセル"),
            ("ゲノム", "ゲノム関連データを読み込みます", "ゲノム", "ゲノムキャンセル"),
            ("新規一括導入", "新規個体を一括で登録します", "新規一括導入", "一括導入キャンセル"),
            ("DC305からイベント吸い込み", "DC305形式Excelからイベントを取り込みます", "DC305", "DC305キャンセル"),
        ]

        content = ttk.Frame(main)
        content.pack(fill=tk.BOTH, expand=True)
        for i, (title, desc, exec_type, cancel_type) in enumerate(rows):
            card = ttk.LabelFrame(content, text=title, padding=12)
            card.pack(fill=tk.X, pady=6)
            row_inner = ttk.Frame(card)
            row_inner.pack(fill=tk.X)
            ttk.Label(row_inner, text=desc, font=(_df, 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
            ttk.Button(
                row_inner, text="実行", style="DataImport.Exec.TButton", width=10,
                command=lambda e=exec_type: self._on_data_import_selected(e, dialog),
            ).pack(side=tk.RIGHT, padx=(4, 0))
            ttk.Button(
                row_inner, text="一括削除", style="DataImport.Secondary.TButton", width=10,
                command=lambda c=cancel_type: self._on_data_import_selected(c, dialog),
            ).pack(side=tk.RIGHT)

        # フッター（閉じるはセカンダリで統一）
        footer = ttk.Frame(main)
        footer.pack(fill=tk.X, pady=(20, 0))
        ttk.Separator(footer, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 12))
        ttk.Button(footer, text="閉じる", style="DataImport.Secondary.TButton", width=12, command=dialog.destroy).pack(side=tk.LEFT)

    def _unified_dialog_icon_font(self):
        """記号が表示されるアイコン用フォント。Meiryo UI に統一（Segoe UI Symbol は Tk で未描画時に '--' になることがあるため）。"""
        return ("Meiryo UI", 22)

    def _build_unified_dialog_header(self, dialog: tk.Toplevel, icon_char: str, title_text: str, subtitle_text: str):
        """吸い込みウィンドウと同一デザインのヘッダー（アイコン・タイトル・サブタイトル）を追加"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        header = tk.Frame(dialog, bg=bg, pady=20, padx=24)
        header.pack(fill=tk.X)
        icon_font = self._unified_dialog_icon_font()
        tk.Label(header, text=icon_char, font=(icon_font[0], 24), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text=title_text, font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text=subtitle_text, font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)

    def _build_unified_dialog_card(self, parent: tk.Frame, icon_char: str, title_text: str, desc_text: str,
                                   btn_label: str, command):
        """吸い込みウィンドウと同一デザインのカード（アイコン・タイトル・説明・1ボタン）を追加"""
        _df = "Meiryo UI"
        card_bg = "#f5f5f5"
        card_pad = 16
        btn_primary_bg = "#3949ab"
        btn_primary_fg = "#ffffff"
        icon_cell_width = 40  # 全カードでアイコン幅を統一して左揃え
        card = tk.Frame(parent, bg=card_bg, padx=card_pad, pady=card_pad,
                        highlightbackground="#e0e7ef", highlightthickness=1)
        card.pack(fill=tk.X, pady=6, padx=24)
        icon_frame = tk.Frame(card, bg=card_bg, width=icon_cell_width)
        icon_frame.pack(side=tk.LEFT, padx=(0, 14), pady=4)
        icon_frame.pack_propagate(False)
        icon_font = self._unified_dialog_icon_font()
        tk.Label(icon_frame, text=icon_char, font=icon_font, bg=card_bg, fg="#3949ab").pack(anchor=tk.W)
        text_frame = tk.Frame(card, bg=card_bg)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(text_frame, text=title_text, font=(_df, 11, "bold"), bg=card_bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(text_frame, text=desc_text, font=(_df, 9), bg=card_bg, fg="#78909c").pack(anchor=tk.W)
        btn = tk.Button(card, text=btn_label, font=(_df, 10), bg=btn_primary_bg, fg=btn_primary_fg,
                       activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
                       padx=16, pady=8, cursor="hand2", command=command)
        btn.pack(side=tk.RIGHT, padx=(10, 0))

    def _build_milk_test_report_card(self, parent: tk.Frame, use_ai_var: tk.BooleanVar):
        """乳検レポート用カード（AI使用チェックは乳検レポート専用としてカード内に配置）"""
        _df = "Meiryo UI"
        card_bg = "#f5f5f5"
        card_pad = 16
        btn_primary_bg = "#3949ab"
        btn_primary_fg = "#ffffff"
        icon_cell_width = 40
        card = tk.Frame(parent, bg=card_bg, padx=card_pad, pady=card_pad,
                        highlightbackground="#e0e7ef", highlightthickness=1)
        card.pack(fill=tk.X, pady=6, padx=24)
        icon_frame = tk.Frame(card, bg=card_bg, width=icon_cell_width)
        icon_frame.pack(side=tk.LEFT, padx=(0, 14), pady=4)
        icon_frame.pack_propagate(False)
        icon_font = self._unified_dialog_icon_font()
        tk.Label(icon_frame, text="", font=icon_font, bg=card_bg, fg="#3949ab").pack(anchor=tk.W)
        text_frame = tk.Frame(card, bg=card_bg)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(text_frame, text="乳検レポート", font=(_df, 11, "bold"), bg=card_bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(text_frame, text="乳検日を選択して乳検レポートを作成します", font=(_df, 9), bg=card_bg, fg="#78909c").pack(anchor=tk.W)
        # AI使用（乳検レポートの前月比較のみ対象）
        ai_frame = tk.Frame(text_frame, bg=card_bg)
        ai_frame.pack(anchor=tk.W, pady=(6, 0))
        tk.Checkbutton(
            ai_frame, text="AI使用（前月との比較を生成）",
            variable=use_ai_var, font=(_df, 10), bg=card_bg, fg="#37474f",
            activebackground=card_bg, activeforeground="#37474f", selectcolor=card_bg
        ).pack(anchor=tk.W)
        tk.Label(ai_frame, text="オフにするとAPIに依存せずレポートのみ出力します。", font=(_df, 9), bg=card_bg, fg="#78909c").pack(anchor=tk.W)
        btn = tk.Button(card, text="実行", font=(_df, 10), bg=btn_primary_bg, fg=btn_primary_fg,
                       activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
                       padx=16, pady=8, cursor="hand2",
                       command=lambda: self._on_milk_test_report_builder(use_ai=use_ai_var.get()))
        btn.pack(side=tk.RIGHT, padx=(10, 0))

    def _build_unified_dialog_footer(self, dialog: tk.Toplevel, on_close=None):
        """吸い込みウィンドウと同一デザインのフッター（閉じるボタン）を追加。on_close 指定時はそのコールバックを使用（閉じる処理を呼び側で行う）。"""
        _df = "Meiryo UI"
        btn_secondary_bg = "#fafafa"
        btn_secondary_fg = "#546e7a"
        btn_secondary_bd = "#b0bec5"
        footer = tk.Frame(dialog, bg="#f5f5f5", pady=20)
        footer.pack(fill=tk.X)
        cmd = on_close if on_close else dialog.destroy
        tk.Button(footer, text="閉じる", font=(_df, 10), bg=btn_secondary_bg, fg=btn_secondary_fg,
                 activebackground="#eceff1", relief=tk.FLAT, padx=24, pady=10,
                 highlightbackground=btn_secondary_bd, highlightthickness=1, cursor="hand2",
                 command=cmd).pack()
    
    def _on_data_import_selected(self, import_type: str, dialog: tk.Toplevel):
        """データ吸い込みタイプが選択されたときの処理"""
        dialog.destroy()  # ダイアログを閉じる
        
        if import_type == "AI":
            self._on_ai_import()
        elif import_type == "AIキャンセル":
            self._on_ai_import_cancel()
        elif import_type == "乳検":
            self._on_milk_test_import_window()
        elif import_type == "乳検キャンセル":
            self._on_milk_test_import_cancel()
        elif import_type == "ゲノム":
            self._on_genome_import()
        elif import_type == "ゲノムキャンセル":
            self._on_genome_import_cancel()
        elif import_type == "新規一括導入":
            self._on_cow_registration_import_window()
        elif import_type == "一括導入キャンセル":
            self._on_cow_registration_cancel()
        elif import_type == "DC305":
            self._on_dc305_import()
        elif import_type == "DC305キャンセル":
            self._on_dc305_import_cancel()
    
    def _on_milk_test_import_window(self):
        """乳検吸い込みウィンドウを開く"""
        from ui.milk_test_import_window import MilkTestImportWindow
        
        milk_test_import_window = MilkTestImportWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine
        )
        milk_test_import_window.show()
    
    def _on_milk_test_import_cancel(self):
        """乳検キャンセル：一括吸い込みした乳検日を選択し、その日の乳検イベントを一括削除"""
        from modules.rule_engine import RuleEngine
        from collections import defaultdict
        
        # 乳検イベント（EVENT_MILK_TEST）を取得し、複数件吸い込みがあった日付のみを取得
        milk_events = self.db.get_events_by_number(RuleEngine.EVENT_MILK_TEST, include_deleted=False)
        if not milk_events:
            messagebox.showinfo("乳検キャンセル", "乳検データの吸い込み日付はありません。")
            return
        
        # 日付ごとの乳検イベント件数を集計し、2件以上ある日付のみに絞る
        date_counts = defaultdict(list)  # date -> list of event (for id)
        for ev in milk_events:
            d = ev.get("event_date")
            if d:
                date_counts[d].append(ev)
        import_dates = sorted(
            [d for d, evs in date_counts.items() if len(evs) >= 2],
            reverse=True
        )
        if not import_dates:
            messagebox.showinfo("乳検キャンセル", "複数件を一括吸い込みした乳検日はありません。")
            return
        
        # 日付選択ダイアログ
        select_dialog = tk.Toplevel(self.root)
        select_dialog.title("乳検キャンセル")
        select_dialog.geometry("380x320")
        
        ttk.Label(select_dialog, text="一括吸い込みした乳検日を選択してください", font=("", 11)).pack(pady=15, padx=15)
        
        list_frame = ttk.Frame(select_dialog)
        list_frame.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, height=10, font=("", 11), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for d in import_dates:
            listbox.insert(tk.END, d)
        if import_dates:
            listbox.selection_set(0)
        
        selected_date = [None]
        
        def on_ok():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("乳検キャンセル", "乳検日を選択してください。")
                return
            selected_date[0] = import_dates[sel[0]]
            select_dialog.destroy()
        
        def on_cancel():
            select_dialog.destroy()
        
        btn_frame = ttk.Frame(select_dialog)
        btn_frame.pack(pady=15, padx=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        
        select_dialog.wait_window()
        
        if selected_date[0] is None:
            return
        
        target_date = selected_date[0]
        
        # その日の乳検イベントを取得
        events_on_date = self.db.get_events_by_number_and_period(
            RuleEngine.EVENT_MILK_TEST, target_date, target_date, include_deleted=False
        )
        count = len(events_on_date)
        
        if count == 0:
            messagebox.showinfo("乳検キャンセル", f"{target_date} の乳検データはありません。")
            return
        
        if not messagebox.askyesno(
            "乳検キャンセル",
            f"{target_date} に吸い込んだ乳検データ {count} 件を削除します。\nよろしいですか？"
        ):
            return
        
        try:
            affected_cows = set()
            for ev in events_on_date:
                event_id = ev.get("id")
                cow_auto_id = ev.get("cow_auto_id")
                if event_id:
                    self.db.delete_event(event_id, soft_delete=True)
                if cow_auto_id:
                    affected_cows.add(cow_auto_id)
            for cow_auto_id in affected_cows:
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
            messagebox.showinfo("乳検キャンセル", f"{count} 件の乳検データを削除しました。")
            if hasattr(self, "_refresh_table_display"):
                self._refresh_table_display()
        except Exception as e:
            messagebox.showerror("乳検キャンセル", f"削除中にエラーが発生しました:\n{e}")
    
    def _on_cow_registration_import_window(self):
        """新規一括導入ウィンドウを開く"""
        from ui.cow_registration_import_window import CowRegistrationImportWindow
        
        cow_registration_import_window = CowRegistrationImportWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path
        )
        cow_registration_import_window.show()
    
    def _on_cow_registration_cancel(self):
        """一括導入キャンセル：一括導入日付を選択し、その日付で導入した個体を一括削除"""
        from modules.rule_engine import RuleEngine
        
        # 導入イベント（EVENT_IN）を取得し、複数頭が導入された日付のみを取得
        intro_events = self.db.get_events_by_number(RuleEngine.EVENT_IN, include_deleted=False)
        if not intro_events:
            messagebox.showinfo("一括導入キャンセル", "一括導入した日付はありません。")
            return
        
        # 日付ごとの導入頭数を集計し、2頭以上導入された日付のみに絞る
        from collections import defaultdict
        date_counts = defaultdict(set)  # date -> set of cow_auto_id
        for ev in intro_events:
            d = ev.get("event_date")
            if d:
                date_counts[d].add(ev.get("cow_auto_id"))
        import_dates = sorted(
            [d for d, ids in date_counts.items() if len(ids) >= 2],
            reverse=True
        )
        if not import_dates:
            messagebox.showinfo("一括導入キャンセル", "複数頭を一括導入した日付はありません。")
            return
        
        # 日付選択ダイアログ
        select_dialog = tk.Toplevel(self.root)
        select_dialog.title("一括導入キャンセル")
        select_dialog.geometry("380x320")
        
        ttk.Label(select_dialog, text="一括導入した日付を選択してください", font=("", 11)).pack(pady=15, padx=15)
        
        list_frame = ttk.Frame(select_dialog)
        list_frame.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, height=10, font=("", 11), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for d in import_dates:
            listbox.insert(tk.END, d)
        if import_dates:
            listbox.selection_set(0)
        
        selected_date = [None]  # クロージャで保持
        
        def on_ok():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("一括導入キャンセル", "日付を選択してください。")
                return
            selected_date[0] = import_dates[sel[0]]
            select_dialog.destroy()
        
        def on_cancel():
            select_dialog.destroy()
        
        btn_frame = ttk.Frame(select_dialog)
        btn_frame.pack(pady=15, padx=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        
        select_dialog.wait_window()
        
        if selected_date[0] is None:
            return
        
        target_date = selected_date[0]
        
        # その日付の導入イベントに紐づく cow_auto_id を取得
        events_on_date = self.db.get_events_by_number_and_period(
            RuleEngine.EVENT_IN, target_date, target_date, include_deleted=False
        )
        cow_auto_ids = list({ev["cow_auto_id"] for ev in events_on_date})
        count = len(cow_auto_ids)
        
        if count == 0:
            messagebox.showinfo("一括導入キャンセル", f"{target_date} に一括導入した個体はありません。")
            return
        
        if not messagebox.askyesno(
            "一括導入キャンセル",
            f"{target_date} に一括導入した {count} 頭を削除します。\nよろしいですか？\n\n※関連するイベントもすべて削除されます。"
        ):
            return
        
        try:
            for auto_id in cow_auto_ids:
                self.db.delete_cow(auto_id)
            messagebox.showinfo("一括導入キャンセル", f"{count} 頭を削除しました。")
            if hasattr(self, "_refresh_table_display"):
                self._refresh_table_display()
        except Exception as e:
            messagebox.showerror("一括導入キャンセル", f"削除中にエラーが発生しました:\n{e}")
    
    def _on_dc305_import(self):
        """DC305からイベント吸い込みウィンドウを開く"""
        from ui.dc305_import_window import DC305ImportWindow
        dc305_window = DC305ImportWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path,
        )
        dc305_window.show()
    
    def _on_dc305_import_cancel(self):
        """DC305キャンセル：吸い込み日を選択し、その日のDC305取込イベントを一括削除"""
        from ui.dc305_import_window import DC305_NOTE_PREFIX
        events = self.db.get_events_by_note_prefix(DC305_NOTE_PREFIX, include_deleted=False)
        if not events:
            messagebox.showinfo("DC305キャンセル", "DC305で取り込んだイベントはありません。")
            return
        import_notes = sorted(set(ev.get("note") or "" for ev in events if (ev.get("note") or "").startswith(DC305_NOTE_PREFIX)), reverse=True)
        if not import_notes:
            messagebox.showinfo("DC305キャンセル", "DC305で取り込んだイベントはありません。")
            return
        select_dialog = tk.Toplevel(self.root)
        select_dialog.title("DC305キャンセル")
        select_dialog.geometry("400x320")
        ttk.Label(select_dialog, text="削除する吸い込み日を選択してください（note: DC305取込:日付）", font=("", 10)).pack(pady=15, padx=15)
        list_frame = ttk.Frame(select_dialog)
        list_frame.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox = tk.Listbox(list_frame, height=10, font=("", 11), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        for note in import_notes:
            listbox.insert(tk.END, note)
        if import_notes:
            listbox.selection_set(0)
        selected_note = [None]
        def on_ok():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("DC305キャンセル", "日付を選択してください。")
                return
            selected_note[0] = import_notes[sel[0]]
            select_dialog.destroy()
        def on_cancel():
            select_dialog.destroy()
        btn_frame = ttk.Frame(select_dialog)
        btn_frame.pack(pady=15, padx=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(side=tk.LEFT)
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        select_dialog.wait_window()
        if selected_note[0] is None:
            return
        target_note = selected_note[0]
        to_delete = [ev for ev in events if ev.get("note") == target_note]
        count = len(to_delete)
        if count == 0:
            messagebox.showinfo("DC305キャンセル", "該当するイベントはありません。")
            return
        if not messagebox.askyesno("DC305キャンセル", f"{target_note} で取り込んだイベント {count} 件を削除します。\nよろしいですか？"):
            return
        try:
            affected_cows = set()
            for ev in to_delete:
                eid = ev.get("id")
                if eid:
                    self.db.delete_event(eid, soft_delete=True)
                if ev.get("cow_auto_id"):
                    affected_cows.add(ev["cow_auto_id"])
            for cow_auto_id in affected_cows:
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
            messagebox.showinfo("DC305キャンセル", f"{count} 件を削除しました。")
            if hasattr(self, "_refresh_table_display"):
                self._refresh_table_display()
        except Exception as e:
            messagebox.showerror("DC305キャンセル", f"削除中にエラーが発生しました:\n{e}")
    
    def _on_ai_import(self):
        """AI吸い込みウィンドウを開く"""
        from ui.ai_import_window import AIImportWindow
        ai_import_window = AIImportWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path,
        )
        ai_import_window.show()
    
    def _on_ai_import_cancel(self):
        """AIキャンセル：一括吸い込みしたAIの日付を選択し、その日のAIイベントを一括削除"""
        from modules.rule_engine import RuleEngine
        from collections import defaultdict
        
        # AIイベント（EVENT_AI）を取得し、複数件吸い込みがあった日付のみを取得
        ai_events = self.db.get_events_by_number(RuleEngine.EVENT_AI, include_deleted=False)
        if not ai_events:
            messagebox.showinfo("AIキャンセル", "AIデータの吸い込み日付はありません。")
            return
        
        # 日付ごとのAIイベント件数を集計し、2件以上ある日付のみに絞る
        date_counts = defaultdict(list)
        for ev in ai_events:
            d = ev.get("event_date")
            if d:
                date_counts[d].append(ev)
        import_dates = sorted(
            [d for d, evs in date_counts.items() if len(evs) >= 2],
            reverse=True
        )
        if not import_dates:
            messagebox.showinfo("AIキャンセル", "複数件を一括吸い込みしたAIの日付はありません。")
            return
        
        # 日付選択ダイアログ
        select_dialog = tk.Toplevel(self.root)
        select_dialog.title("AIキャンセル")
        select_dialog.geometry("380x320")
        
        ttk.Label(select_dialog, text="一括吸い込みしたAIの日付を選択してください", font=("", 11)).pack(pady=15, padx=15)
        
        list_frame = ttk.Frame(select_dialog)
        list_frame.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, height=10, font=("", 11), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for d in import_dates:
            listbox.insert(tk.END, d)
        if import_dates:
            listbox.selection_set(0)
        
        selected_date = [None]
        
        def on_ok():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("AIキャンセル", "日付を選択してください。")
                return
            selected_date[0] = import_dates[sel[0]]
            select_dialog.destroy()
        
        def on_cancel():
            select_dialog.destroy()
        
        btn_frame = ttk.Frame(select_dialog)
        btn_frame.pack(pady=15, padx=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        
        select_dialog.wait_window()
        
        if selected_date[0] is None:
            return
        
        target_date = selected_date[0]
        
        # その日のAIイベントを取得
        events_on_date = self.db.get_events_by_number_and_period(
            RuleEngine.EVENT_AI, target_date, target_date, include_deleted=False
        )
        count = len(events_on_date)
        
        if count == 0:
            messagebox.showinfo("AIキャンセル", f"{target_date} のAIデータはありません。")
            return
        
        if not messagebox.askyesno(
            "AIキャンセル",
            f"{target_date} に吸い込んだAIデータ {count} 件を削除します。\nよろしいですか？"
        ):
            return
        
        try:
            affected_cows = set()
            for ev in events_on_date:
                event_id = ev.get("id")
                cow_auto_id = ev.get("cow_auto_id")
                if event_id:
                    self.db.delete_event(event_id, soft_delete=True)
                if cow_auto_id:
                    affected_cows.add(cow_auto_id)
            for cow_auto_id in affected_cows:
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
            messagebox.showinfo("AIキャンセル", f"{count} 件のAIデータを削除しました。")
            if hasattr(self, "_refresh_table_display"):
                self._refresh_table_display()
        except Exception as e:
            messagebox.showerror("AIキャンセル", f"削除中にエラーが発生しました:\n{e}")
    
    def _on_genome_import(self):
        """ゲノム吸い込みウィンドウを開く"""
        from ui.genome_import_window import GenomeImportWindow

        genome_import_window = GenomeImportWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            item_dict_path=self.item_dict_path,
        )
        genome_import_window.show()
        # ゲノム項目を辞書に追加した場合に備え、項目辞書を再読み込み（個体カードのゲノムタブ・リスト/グラフで反映）
        try:
            self.formula_engine.reload_item_dictionary()
            self._load_item_dictionary()
        except Exception as e:
            logging.debug(f"項目辞書再読み込み: {e}")
    
    def _on_genome_import_cancel(self):
        """ゲノムキャンセル：一括吸い込みしたゲノム日を選択し、その日のゲノムイベントを一括削除"""
        from modules.rule_engine import RuleEngine
        from collections import defaultdict
        
        # ゲノムイベント（EVENT_GENOMIC）を取得し、複数件吸い込みがあった日付のみを取得
        genome_events = self.db.get_events_by_number(RuleEngine.EVENT_GENOMIC, include_deleted=False)
        if not genome_events:
            messagebox.showinfo("ゲノムキャンセル", "ゲノムデータの吸い込み日付はありません。")
            return
        
        # 日付ごとのゲノムイベント件数を集計し、2件以上ある日付のみに絞る
        date_counts = defaultdict(list)
        for ev in genome_events:
            d = ev.get("event_date")
            if d:
                date_counts[d].append(ev)
        import_dates = sorted(
            [d for d, evs in date_counts.items() if len(evs) >= 2],
            reverse=True
        )
        if not import_dates:
            messagebox.showinfo("ゲノムキャンセル", "複数件を一括吸い込みしたゲノムの日付はありません。")
            return
        
        # 日付選択ダイアログ
        select_dialog = tk.Toplevel(self.root)
        select_dialog.title("ゲノムキャンセル")
        select_dialog.geometry("380x320")
        
        ttk.Label(select_dialog, text="一括吸い込みしたゲノムの日付を選択してください", font=("", 11)).pack(pady=15, padx=15)
        
        list_frame = ttk.Frame(select_dialog)
        list_frame.pack(pady=10, padx=15, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, height=10, font=("", 11), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for d in import_dates:
            listbox.insert(tk.END, d)
        if import_dates:
            listbox.selection_set(0)
        
        selected_date = [None]
        
        def on_ok():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("ゲノムキャンセル", "日付を選択してください。")
                return
            selected_date[0] = import_dates[sel[0]]
            select_dialog.destroy()
        
        def on_cancel():
            select_dialog.destroy()
        
        btn_frame = ttk.Frame(select_dialog)
        btn_frame.pack(pady=15, padx=15)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        listbox.bind("<Double-Button-1>", lambda e: on_ok())
        
        select_dialog.wait_window()
        
        if selected_date[0] is None:
            return
        
        target_date = selected_date[0]
        
        # その日のゲノムイベントを取得
        events_on_date = self.db.get_events_by_number_and_period(
            RuleEngine.EVENT_GENOMIC, target_date, target_date, include_deleted=False
        )
        count = len(events_on_date)
        
        if count == 0:
            messagebox.showinfo("ゲノムキャンセル", f"{target_date} のゲノムデータはありません。")
            return
        
        if not messagebox.askyesno(
            "ゲノムキャンセル",
            f"{target_date} に吸い込んだゲノムデータ {count} 件を削除します。\nよろしいですか？"
        ):
            return
        
        try:
            affected_cows = set()
            for ev in events_on_date:
                event_id = ev.get("id")
                cow_auto_id = ev.get("cow_auto_id")
                if event_id:
                    self.db.delete_event(event_id, soft_delete=True)
                if cow_auto_id:
                    affected_cows.add(cow_auto_id)
            for cow_auto_id in affected_cows:
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
            messagebox.showinfo("ゲノムキャンセル", f"{count} 件のゲノムデータを削除しました。")
            if hasattr(self, "_refresh_table_display"):
                self._refresh_table_display()
        except Exception as e:
            messagebox.showerror("ゲノムキャンセル", f"削除中にエラーが発生しました:\n{e}")
    
    def _on_data_output(self):
        """データ出力メニューをクリック（吸い込みウィンドウと同一デザイン）。既に開いていれば前面に表示。"""
        if self._data_output_dialog is not None and self._data_output_dialog.winfo_exists():
            self._data_output_dialog.lift()
            self._data_output_dialog.focus_force()
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("データ出力")
        dialog.geometry("520x380")
        dialog.configure(bg="#f5f5f5")
        dialog.minsize(480, 320)
        self._data_output_dialog = dialog

        def _close_data_output():
            self._data_output_dialog = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close_data_output)

        self._build_unified_dialog_header(dialog, "\u2b06\ufe0f", "データ出力", "出力するレポートを選択してください")
        content = tk.Frame(dialog, bg="#f5f5f5", padx=0, pady=8)
        content.pack(fill=tk.BOTH, expand=True)
        self._build_unified_dialog_card(content, "\U0001f95b", "乳検", "乳検レポートを出力します", "実行",
                                        lambda: self._on_milk_test_report(dialog))
        self._build_unified_dialog_footer(dialog, on_close=_close_data_output)

    def _on_report_builder(self):
        """レポート作成メニューをクリック（吸い込みウィンドウと同一デザイン）。既に開いていれば前面に表示。"""
        if self._report_builder_dialog is not None and self._report_builder_dialog.winfo_exists():
            self._report_builder_dialog.lift()
            self._report_builder_dialog.focus_force()
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("レポート作成")
        dialog.geometry("600x540")
        dialog.configure(bg="#f5f5f5")
        dialog.minsize(560, 440)
        self._report_builder_dialog = dialog

        def _close_report_builder():
            self._report_builder_dialog = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close_report_builder)

        self._build_unified_dialog_header(dialog, "\U0001F4CB", "レポート作成", "作成するレポートを選択してください")  # 📋
        content = tk.Frame(dialog, bg="#f5f5f5", padx=0, pady=8)
        content.pack(fill=tk.BOTH, expand=True)
        _df = "Meiryo UI"
        # 乳検レポート用カード（AI使用は乳検レポート専用なのでカード内に配置）
        dialog.use_ai_var = tk.BooleanVar(value=True)
        self._build_milk_test_report_card(content, dialog.use_ai_var)
        # ゲノムレポート
        self._build_unified_dialog_card(
            content, "", "ゲノムレポート", "", "実行",
            self._on_genome_report_builder
        )
        # 牛群動態予測レポート（乳用種♀出生見込みから頭数変動を予測。実装はあとから）
        self._build_unified_dialog_card(
            content, "", "牛群動態予測レポート", "乳用種♀の出生見込みから、今後の頭数変動を予測します（実装予定）", "実行",
            self._on_herd_dynamics_report_builder
        )
        self._build_unified_dialog_footer(dialog, on_close=_close_report_builder)
    
    def _on_genome_report_builder(self):
        """ゲノムレポート：設定ウィンドウを開き、DBのゲノムデータでレポート生成。既に開いていれば前面に表示。"""
        if not self.db:
            messagebox.showerror("ゲノムレポート", "データベースが選択されていません。")
            return
        if self._genome_report_window is not None and self._genome_report_window.window.winfo_exists():
            self._genome_report_window.window.lift()
            self._genome_report_window.window.focus_force()
            return
        try:
            from ui.genome_report_window import GenomeReportWindow
            farm_name = ""
            if self.farm_path:
                farm_name = Path(self.farm_path).name if isinstance(self.farm_path, (str, Path)) else getattr(self.farm_path, "name", "")
            genome_report_window = GenomeReportWindow(
                parent=self.root,
                db=self.db,
                item_dict_path=self.item_dict_path,
                farm_name=farm_name,
                on_close=lambda: setattr(self, "_genome_report_window", None),
            )
            self._genome_report_window = genome_report_window
            genome_report_window.show()
        except Exception as e:
            logging.error(f"ゲノムレポート起動エラー: {e}", exc_info=True)
            messagebox.showerror("ゲノムレポート", f"起動に失敗しました:\n{e}")

    def _on_herd_dynamics_report_builder(self):
        """牛群動態予測レポート：分娩予定月×予定産子種類の表とグラフをブラウザで表示"""
        if not self.farm_path:
            messagebox.showerror("エラー", "農場が選択されていません")
            return
        try:
            from modules.herd_dynamics_report import build_herd_dynamics_data
            data = build_herd_dynamics_data(self.db, self.formula_engine, self.farm_path)
            table_rows = data["table_rows"]
            months = data["months"]
            if not table_rows:
                messagebox.showinfo("牛群動態予測レポート", "分娩予定が決まっている個体がありません。")
                return
            html_content = self._build_herd_dynamics_report_html(data)
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", suffix=".html", delete=False) as f:
                f.write(html_content)
                temp_file_path = f.name
            webbrowser.open(f"file://{temp_file_path}")
            if self.app_settings.get("dont_show_herd_dynamics_report_info"):
                return
            self._show_herd_dynamics_report_info_dialog()
        except Exception as e:
            logging.error(f"牛群動態予測レポートエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"牛群動態予測レポートの作成中にエラーが発生しました:\n{e}")

    def _show_herd_dynamics_report_info_dialog(self):
        """牛群動態予測レポートの案内ダイアログ（次回から表示しないチェック付き）"""
        dialog = tk.Toplevel(self.root)
        dialog.title("牛群動態予測レポート")
        dialog.configure(bg="#f5f5f5")
        dialog.transient(self.root)
        dialog.grab_set()
        # メッセージ
        msg_frame = tk.Frame(dialog, bg="#f5f5f5", padx=16, pady=12)
        msg_frame.pack(fill=tk.X)
        info_icon = tk.Label(msg_frame, text="ℹ", font=("Segoe UI Symbol", 14), fg="#2196F3", bg="#f5f5f5")
        info_icon.pack(side=tk.LEFT, padx=(0, 8))
        msg_text = "ブラウザでレポートが開かれました。\n印刷する場合はブラウザのメニューから「印刷」を選択してください。"
        msg_label = tk.Label(msg_frame, text=msg_text, font=("Meiryo UI", 9), bg="#f5f5f5", justify=tk.LEFT)
        msg_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # チェックボックスとOK
        bottom = tk.Frame(dialog, bg="#f5f5f5", padx=16, pady=8)
        bottom.pack(fill=tk.X)
        dont_show_var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(bottom, text="次回から表示しない", variable=dont_show_var)
        cb.pack(side=tk.LEFT)
        def _on_ok():
            if dont_show_var.get():
                self.app_settings.set("dont_show_herd_dynamics_report_info", True)
            dialog.destroy()
        ok_btn = ttk.Button(bottom, text="OK", command=_on_ok)
        ok_btn.pack(side=tk.RIGHT)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_reqwidth() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_reqheight() // 2)
        dialog.geometry(f"+{x}+{y}")
        ok_btn.focus_set()
        dialog.wait_window()

    def _build_herd_dynamics_chart_section(self, data):
        """分娩予定月別の産子種類内訳グラフのHTMLフラグメントを返す（ダッシュボード用）。全種類（不明含む）表示・1行表示でコンパクトに。"""
        import html as html_module
        if not data or not data.get("table_rows"):
            return '<div class="subnote">分娩予定が決まっている個体がありません。</div>'
        table_rows = data["table_rows"]
        calf_types = data["calf_types"]
        colors = {
            "乳用種メス": "#E91E63",
            "乳用種": "#9C27B0",
            "F1": "#2196F3",
            "黒毛和種": "#4CAF50",
            "不明": "#9E9E9E",
        }
        max_total = max((r.get("合計", 0) for r in table_rows), default=1)
        bar_height = 22
        row_height = 30
        chart_width = 400
        year_month_width = 100   # 左端：年-月・総頭数
        chart_start_x = year_month_width + 8
        label_x = chart_start_x + chart_width + 8
        detail_width = 320
        svg_total_width = year_month_width + 8 + chart_width + 8 + detail_width

        def _detail_tspans(r):
            """内訳をグラフ色のtspanで返す（総頭数は含めない）。F1など末尾が数字のラベルは頭数と区切る"""
            labels = {"乳用種メス": "乳用種♀", "乳用種": "乳用種", "F1": "F1", "黒毛和種": "黒毛和種", "不明": "不明"}
            tspan_parts = []
            for ct in calf_types:
                n = r.get(ct, 0) or 0
                if n > 0:
                    color = colors.get(ct, "#333")
                    label = labels.get(ct, ct)
                    # F1＋1頭→「F11頭」と誤読されないよう、ラベル末尾が数字のときはスペースを挟む
                    sep = " " if (label and label[-1].isdigit()) else ""
                    tspan_parts.append(f'<tspan fill="{color}">{html_module.escape(label)}{sep}{n}頭</tspan>')
            if not tspan_parts:
                return ""
            return '<tspan fill="#666"> </tspan>' + '<tspan fill="#999">、</tspan>'.join(tspan_parts)

        svg_parts = []
        for i, row in enumerate(table_rows):
            due_ym = row.get("due_ym", "")
            total = row.get("合計", 0)
            y = i * row_height + 12
            # 左端：年-月・総頭数
            svg_parts.append(
                f'<text x="8" y="{y + bar_height / 2 + 3}" font-size="11" fill="#333" text-anchor="start">{html_module.escape(due_ym)}（{total}頭）</text>'
            )
            x_offset = chart_start_x
            for ct in calf_types:
                cnt = row.get(ct, 0)
                if cnt <= 0:
                    continue
                w = (cnt / max_total * chart_width) if max_total else 0
                fill = colors.get(ct, "#9E9E9E")
                svg_parts.append(
                    f'<rect x="{x_offset}" y="{y}" width="{w:.1f}" height="{bar_height - 2}" fill="{fill}" stroke="#fff" stroke-width="1"/>'
                )
                x_offset += w
            # 棒の右隣に内訳テキスト
            detail_tspans = _detail_tspans(row)
            if detail_tspans:
                svg_parts.append(
                    f'<text x="{label_x}" y="{y + bar_height / 2 + 3}" font-size="11">{detail_tspans}</text>'
                )
        svg_content = "\n".join(svg_parts)
        legend = "".join(
            f'<span class="legend-swatch" style="background-color:{colors.get(ct, "#999")};"></span><span class="legend-label">{html_module.escape(ct)}</span> '
            for ct in calf_types
        )
        return f'''<div class="section-title">分娩予定月別の産子種類内訳</div>
<div class="herd-dynamics-chart"><svg width="{svg_total_width}" height="{len(table_rows) * row_height + 24}" xmlns="http://www.w3.org/2000/svg">
{svg_content}
</svg></div>
<div class="legend herd-dynamics-legend">{legend}</div>'''

    def _build_herd_dynamics_report_html(self, data: Dict[str, Any]) -> str:
        """分娩予定月×予定産子種類の表とグラフのHTMLを組み立てる。全種類（不明含む）表示・1行でコンパクトに。"""
        import html as html_module
        import math
        table_rows = data["table_rows"]
        months = data["months"]
        calf_types = data["calf_types"]
        total_by_type = data.get("total_by_type") or {}
        unknown_details = data.get("unknown_details") or []
        colors = {
            "乳用種メス": "#E91E63",
            "乳用種": "#9C27B0",
            "F1": "#2196F3",
            "黒毛和種": "#4CAF50",
            "不明": "#9E9E9E",
        }
        max_total = max((r.get("合計", 0) for r in table_rows), default=1) if table_rows else 1
        bar_height = 16
        row_height = 22
        chart_width = 400
        year_month_width = 100   # 左端：年-月・総頭数
        chart_start_x = year_month_width + 8
        label_x = chart_start_x + chart_width + 8
        labels_short = {"乳用種メス": "乳用種♀", "乳用種": "乳用種", "F1": "F1", "黒毛和種": "黒毛和種", "不明": "不明"}
        detail_width = 320
        svg_total_width = year_month_width + 8 + chart_width + 8 + detail_width

        def _detail_tspans(r: Dict[str, Any]) -> str:
            """内訳をグラフ色のtspanで返す（総頭数は含めない）。F1など末尾が数字のラベルは頭数と区切る"""
            tspan_parts = []
            for ct in calf_types:
                n = r.get(ct, 0) or 0
                if n > 0:
                    color = colors.get(ct, "#333")
                    label = labels_short.get(ct, ct)
                    # F1＋1頭→「F11頭」と誤読されないよう、ラベル末尾が数字のときはスペースを挟む
                    sep = " " if (label and label[-1].isdigit()) else ""
                    tspan_parts.append(f'<tspan fill="{color}">{html_module.escape(label)}{sep}{n}頭</tspan>')
            if not tspan_parts:
                return ""
            return '<tspan fill="#666"> </tspan>' + '<tspan fill="#999">、</tspan>'.join(tspan_parts)

        svg_parts = []
        for i, row in enumerate(table_rows):
            due_ym = row.get("due_ym", "")
            total = row.get("合計", 0)
            y = i * row_height + 12
            # 左端：年-月・総頭数
            svg_parts.append(
                f'<text x="8" y="{y + bar_height / 2 + 2}" font-size="9" fill="#333" text-anchor="start">{html_module.escape(due_ym)}（{total}頭）</text>'
            )
            x_offset = chart_start_x
            for ct in calf_types:
                cnt = row.get(ct, 0)
                if cnt <= 0:
                    continue
                w = (cnt / max_total * chart_width) if max_total else 0
                fill = colors.get(ct, "#9E9E9E")
                svg_parts.append(
                    f'<rect x="{x_offset}" y="{y}" width="{w:.1f}" height="{bar_height - 2}" fill="{fill}" stroke="#fff" stroke-width="1"/>'
                )
                x_offset += w
            # 棒の右隣に内訳テキスト
            detail_tspans = _detail_tspans(row)
            if detail_tspans:
                svg_parts.append(
                    f'<text x="{label_x}" y="{y + bar_height / 2 + 2}" font-size="9">{detail_tspans}</text>'
                )
        svg_content = "\n".join(svg_parts)
        table_body = "".join(
            f"<tr><td>{html_module.escape(r.get('due_ym', ''))}</td>"
            + "".join(f"<td class=\"num\">{r.get(ct, 0)}</td>" for ct in calf_types)
            + f"<td class=\"num\"><strong>{r.get('合計', 0)}</strong></td></tr>"
            for r in table_rows
        )
        total_row = (
            "<tr><td><strong>合計</strong></td>"
            + "".join(f"<td class=\"num\"><strong>{total_by_type.get(ct, 0)}</strong></td>" for ct in calf_types)
            + f"<td class=\"num\"><strong>{sum(total_by_type.get(ct, 0) for ct in calf_types)}</strong></td></tr>"
        )
        th_cells = "".join(f"<th>{html_module.escape(ct)}</th>" for ct in calf_types)
        # 列幅を均一にする（分娩予定月 + 産子種類列 + 合計）
        ncols = 2 + len(calf_types)
        col_pct = 100.0 / ncols
        colgroup = '<colgroup>' + ''.join(f'<col style="width:{col_pct:.2f}%"/>' for _ in range(ncols)) + '</colgroup>'
        # 凡例：全種類（不明含む）
        legend = "".join(
            f'<span class="legend-swatch" style="background-color:{colors.get(ct, "#999")};"></span><span class="legend-label">{html_module.escape(ct)}</span> '
            for ct in calf_types
        )
        # ドーナツグラフ（全予定産子の割合）
        def _build_donut_svg() -> str:
            total = sum(total_by_type.get(ct, 0) for ct in calf_types)
            if total <= 0:
                return '<div class="donut-note">分娩予定が決まっている個体がありません。</div>'
            cx, cy = 80, 80
            r_outer = 60
            r_inner = 35
            start_angle = -math.pi / 2  # 上から時計回りに描画
            paths = []

            def polar_to_cartesian(cx, cy, r, angle):
                return cx + r * math.cos(angle), cy + r * math.sin(angle)

            for ct in calf_types:
                value = total_by_type.get(ct, 0)
                if value <= 0:
                    continue
                fraction = value / total
                angle = fraction * 2 * math.pi
                end_angle = start_angle + angle
                x0, y0 = polar_to_cartesian(cx, cy, r_outer, start_angle)
                x1, y1 = polar_to_cartesian(cx, cy, r_outer, end_angle)
                x2, y2 = polar_to_cartesian(cx, cy, r_inner, end_angle)
                x3, y3 = polar_to_cartesian(cx, cy, r_inner, start_angle)
                large_arc = 1 if angle > math.pi else 0
                fill = colors.get(ct, "#9E9E9E")
                path_d = (
                    f"M {x0:.2f} {y0:.2f} "
                    f"A {r_outer} {r_outer} 0 {large_arc} 1 {x1:.2f} {y1:.2f} "
                    f"L {x2:.2f} {y2:.2f} "
                    f"A {r_inner} {r_inner} 0 {large_arc} 0 {x3:.2f} {y3:.2f} Z"
                )
                paths.append(f'<path d="{path_d}" fill="{fill}" stroke="#fff" stroke-width="1"/>')
                start_angle = end_angle
            paths_svg = "\n".join(paths)
            total_label = f"総頭数 {total}頭"
            return f'''
<svg class="donut-chart" width="160" height="160" viewBox="0 0 160 160" xmlns="http://www.w3.org/2000/svg">
<g>
{paths_svg}
</g>
<text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="middle" font-size="11" fill="#37474F">{html_module.escape(total_label)}</text>
</svg>
'''

        donut_svg = _build_donut_svg()
        donut_html = f'''
<div class="donut-container">
  <div class="section-title-small">全予定産子の割合</div>
  {donut_svg}
  <div class="legend donut-legend">{legend}</div>
</div>
'''
        unknown_section = ""
        if unknown_details:
            unknown_rows = "".join(
                f"<tr><td>{html_module.escape(u.get('cow_id', ''))}</td><td>{html_module.escape(str(u.get('sire', '')))}</td><td>{html_module.escape(u.get('due_ym', ''))}</td></tr>"
                for u in unknown_details
            )
            unknown_section = f"""
<div class="section">
<div class="section-title">不明の内訳（{len(unknown_details)}頭）・SIRE一覧で登録すると産子種類に反映されます</div>
<table>
<thead><tr><th>個体ID</th><th>SIRE</th><th>分娩予定月</th></tr></thead>
<tbody>
{unknown_rows}
</tbody>
</table>
</div>"""
        return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>牛群動態予測レポート</title>
<style>
body {{ font-family: "Meiryo UI", sans-serif; margin: 12px; background: #f5f5f5; color: #263238; font-size: 12px; }}
h1 {{ font-size: 14px; margin-bottom: 6px; }}
.section {{ background: #fff; padding: 10px 12px; margin-bottom: 10px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
.section-main .section-title {{ margin-bottom: 6px; }}
.section-main-inner {{ display: flex; gap: 12px; align-items: flex-start; }}
.section-main-left {{ flex: 0 0 66.666%; max-width: 66.666%; }}
.section-main-right {{ flex: 1; min-width: 0; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; }}
.section-title {{ font-size: 12px; font-weight: bold; margin-bottom: 8px; }}
table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
table.dynamics-table {{ table-layout: fixed; width: 100%; }}
table.dynamics-table th, table.dynamics-table td {{ border: 1px solid #e0e0e0; padding: 4px 6px; text-align: left; font-size: 10px; }}
table.dynamics-table th {{ background: #eceff1; font-weight: bold; }}
td.num {{ text-align: right; }}
.donut-container {{ width: 100%; text-align: center; }}
.donut-chart {{ display: block; margin: 4px auto; }}
.section-title-small {{ font-size: 11px; font-weight: bold; margin-bottom: 4px; }}
.legend {{ margin-top: 8px; font-size: 10px; color: #546e7a; display: flex; flex-wrap: wrap; align-items: center; gap: 4px 12px; }}
.legend-swatch {{ display: inline-block; width: 12px; height: 12px; border: 1px solid #666; vertical-align: middle; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.legend-label {{ vertical-align: middle; }}
.chart-section {{ page-break-inside: avoid; }}
svg {{ font-family: "Meiryo UI", sans-serif; }}
@media print {{
  @page {{ size: A4 portrait; margin: 1cm; }}
  body {{ margin: 0; padding: 0; font-size: 11px; }}
  .section {{ padding: 8px 10px; margin-bottom: 8px; }}
  .section-title {{ font-size: 11px; margin-bottom: 6px; }}
  table.dynamics-table th, table.dynamics-table td {{ padding: 3px 5px; font-size: 9px; }}
  .chart-section {{ page-break-inside: avoid; break-inside: avoid; }}
  .section-main-inner {{ display: flex; gap: 8px; align-items: flex-start; }}
  .section-main-left {{ flex: 0 0 66%; max-width: 66%; }}
  .section-main-right {{ flex: 1; }}
  .section {{ page-break-inside: avoid; break-inside: avoid; }}
  .legend {{ font-size: 9px; margin-top: 6px; }}
  .legend-swatch {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  svg rect {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
}}
</style>
</head>
<body>
<h1>牛群動態予測レポート（分娩予定月×予定産子種類）</h1>
<div class="section section-main">
<div class="section-title">表：分娩予定月別・産子種類別頭数</div>
<div class="section-main-inner">
<div class="section-main-left">
<table class="dynamics-table">
{colgroup}
<thead><tr><th>分娩予定月</th>{th_cells}<th>合計</th></tr></thead>
<tbody>
{table_body}
{total_row}
</tbody>
</table>
</div>
<div class="section-main-right">
{donut_html}
</div>
</div>
</div>
<div class="section chart-section">
<div class="section-title">分娩予定月別の産子種類内訳</div>
<svg width="{svg_total_width}" height="{len(table_rows) * row_height + 24}" xmlns="http://www.w3.org/2000/svg">
{svg_content}
</svg>
<div class="legend">{legend}</div>
</div>
{unknown_section}
</body>
</html>"""

    def _on_dictionary_settings(self):
        """辞書設定メニューをクリック（吸い込みウィンドウと同一デザイン）。既に開いていれば前面に表示。"""
        if self._dictionary_settings_dialog is not None and self._dictionary_settings_dialog.winfo_exists():
            self._dictionary_settings_dialog.lift()
            self._dictionary_settings_dialog.focus_force()
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("辞書・設定")
        dialog.geometry("520x620")
        dialog.configure(bg="#f5f5f5")
        dialog.minsize(480, 560)
        self._dictionary_settings_dialog = dialog

        def _close_dictionary_settings():
            self._dictionary_settings_dialog = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close_dictionary_settings)

        self._build_unified_dialog_header(dialog, "\u2605", "辞書・設定", "操作を選択してください")  # ★ BMPで表示確実
        # スクロール可能なコンテナ（ウィンドウサイズが小さくても全項目を表示）
        container = tk.Frame(dialog, bg="#f5f5f5")
        container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container, bg="#f5f5f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container)
        content = tk.Frame(canvas, bg="#f5f5f5", padx=0, pady=8)
        content_window = canvas.create_window((0, 0), window=content, anchor=tk.NW)

        def _on_content_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(content_window, width=event.width)

        content.bind("<Configure>", _on_content_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=canvas.yview)

        def _on_mousewheel(event):
            if getattr(event, "delta", None):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)
        content.bind("<MouseWheel>", _on_mousewheel)
        content.bind("<Button-4>", _on_mousewheel)
        content.bind("<Button-5>", _on_mousewheel)
        canvas.bind("<Enter>", lambda e: (canvas.bind_all("<MouseWheel>", _on_mousewheel),
                                          canvas.bind_all("<Button-4>", _on_mousewheel),
                                          canvas.bind_all("<Button-5>", _on_mousewheel)))
        canvas.bind("<Leave>", lambda e: (canvas.unbind_all("<MouseWheel>"),
                                          canvas.unbind_all("<Button-4>"),
                                          canvas.unbind_all("<Button-5>")))

        def _on_dialog_destroy(_event):
            try:
                canvas.unbind_all("<MouseWheel>")
                canvas.unbind_all("<Button-4>")
                canvas.unbind_all("<Button-5>")
            except tk.TclError:
                pass
        dialog.bind("<Destroy>", _on_dialog_destroy)

        # アイコンは白紙（表示しない）
        self._build_unified_dialog_card(content, "", "イベント辞書", "イベント種別の辞書を編集します", "開く",
                                        lambda: self._on_event_dictionary(dialog))
        self._build_unified_dialog_card(content, "", "コマンド辞書", "コマンドの辞書を編集します", "開く",
                                        lambda: self._on_command_dictionary(dialog))
        self._build_unified_dialog_card(content, "", "項目辞書", "項目の辞書を編集します", "開く",
                                        lambda: self._on_item_dictionary(dialog))
        self._build_unified_dialog_card(content, "", "SIRE一覧", "SIRE一覧と種別を設定します", "開く",
                                        lambda: self._on_sire_list(dialog))
        self._build_unified_dialog_card(content, "", "農場設定", "授精設定・ペン設定など農場の設定を編集します", "開く",
                                        lambda: self._on_farm_settings(dialog))
        self._build_unified_dialog_card(content, "", "バックアップ設定", "バックアップの間隔と保存先を設定します", "開く",
                                        lambda: self._on_backup_settings(dialog))
        self._build_unified_dialog_footer(dialog, on_close=_close_dictionary_settings)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

    def _on_milk_test_report(self, parent_dialog: Optional[tk.Toplevel] = None):
        """乳検データ出力（乳検日を選択して表示）"""
        try:
            selected_date = self._select_milk_test_event_date()
            if not selected_date:
                return
            self._show_milk_test_report_for_date(selected_date, parent_dialog)
        except Exception as e:
            logging.error(f"乳検データ出力エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"乳検データ出力に失敗しました: {e}")

    def _on_milk_test_report_builder(self, use_ai: bool = True):
        """レポート作成＞乳検レポート（乳検日選択まで）。use_ai=False のときは前月比較（AI）をスキップしAPIに依存しない。"""
        try:
            selected_date = self._select_milk_test_event_date()
            if not selected_date:
                return

            # 乳検イベントを対象に集計
            events = self.db.get_events_by_number(self.rule_engine.EVENT_MILK_TEST, include_deleted=False)
            target_events = [e for e in events if e.get("event_date") == selected_date]
            if not target_events:
                messagebox.showinfo("情報", "対象日の乳検イベントがありません。")
                return

            cow_ids = set()
            dim_values: List[int] = []
            milk_values: List[float] = []
            milk_stats: List[Tuple[float, str]] = []
            scc_values: List[float] = []
            ls_values: List[float] = []
            parity_values: List[int] = []
            parity_counts = {"1": 0, "2": 0, "3+": 0}  # 3つのカテゴリ（1産、2産、3産以上）
            milk_scatter_points: List[Tuple[int, float, int, str, str]] = []
            ls_scatter_points: List[Tuple[int, float, int, str, str]] = []

            def calc_dim_on_date(clvd_date: Optional[str], event_date: str) -> Optional[int]:
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

            selected_dt = None
            try:
                selected_dt = datetime.strptime(selected_date, "%Y-%m-%d")
            except (ValueError, TypeError):
                selected_dt = None

            # 分娩イベント：検定日以前で最新の分娩日（検定時点の泌乳 DIM 用。産次更新後も再現する）
            calv_events = self.db.get_events_by_number(self.rule_engine.EVENT_CALV, include_deleted=False)
            calv_dates_by_cow: Dict[int, List[str]] = {}
            for ce in calv_events:
                caid = ce.get("cow_auto_id")
                ed = ce.get("event_date")
                if not caid or not ed:
                    continue
                try:
                    aid = int(caid)
                except (ValueError, TypeError):
                    continue
                ed_s = str(ed).strip()[:10]
                calv_dates_by_cow.setdefault(aid, []).append(ed_s)
            for aid in calv_dates_by_cow:
                calv_dates_by_cow[aid].sort()

            def clvd_for_milk_dim(cow_auto_id: Any, cow: Optional[Dict[str, Any]]) -> Optional[str]:
                """検定日時点の直近分娩日。イベントが無い場合は cow.clvd にフォールバック。"""
                try:
                    aid = int(cow_auto_id)
                except (ValueError, TypeError):
                    return cow.get("clvd") if cow else None
                dates = calv_dates_by_cow.get(aid, [])
                on_or_before = [d for d in dates if d <= selected_date]
                if on_or_before:
                    return max(on_or_before)
                if cow:
                    return cow.get("clvd")
                return None

            # 乳検イベント履歴（前月SCC取得用）
            event_history: Dict[int, List[Tuple[datetime, Optional[float]]]] = {}
            milk_history: Dict[int, List[Tuple[datetime, Optional[float]]]] = {}
            for event in events:
                cow_auto_id = event.get("cow_auto_id")
                event_date = event.get("event_date")
                if not cow_auto_id or not event_date:
                    continue
                try:
                    event_dt = datetime.strptime(event_date, "%Y-%m-%d")
                except (ValueError, TypeError):
                    continue
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                scc_value = json_data.get("scc")
                scc_num = None
                if scc_value is not None and scc_value != "":
                    try:
                        scc_num = float(scc_value)
                    except (ValueError, TypeError):
                        scc_num = None
                event_history.setdefault(cow_auto_id, []).append((event_dt, scc_num))
                milk_value = json_data.get("milk_yield")
                milk_num = None
                if milk_value is not None and milk_value != "":
                    try:
                        milk_num = float(milk_value)
                    except (ValueError, TypeError):
                        milk_num = None
                milk_history.setdefault(cow_auto_id, []).append((event_dt, milk_num))
            for history in event_history.values():
                history.sort(key=lambda x: x[0])
            for history in milk_history.values():
                history.sort(key=lambda x: x[0])

            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                cow_ids.add(cow_auto_id)
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                clvd = clvd_for_milk_dim(cow_auto_id, cow)
                dim = calc_dim_on_date(clvd, selected_date)
                if dim is not None:
                    dim_values.append(dim)
                if cow and cow.get("lact") is not None:
                    try:
                        lact_val = int(cow.get("lact"))
                        parity_values.append(lact_val)
                        # 産次を3つのカテゴリに分類（1産、2産、3産以上）
                        if lact_val == 1:
                            parity_counts["1"] += 1
                        elif lact_val == 2:
                            parity_counts["2"] += 1
                        elif lact_val >= 3:
                            parity_counts["3+"] += 1
                    except (ValueError, TypeError):
                        pass
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                milk_value = json_data.get("milk_yield")
                if milk_value is not None and milk_value != "":
                    try:
                        milk_value_num = float(milk_value)
                        milk_values.append(milk_value_num)
                        cow_id = None
                        if cow:
                            cow_id = cow.get("cow_id") or cow.get("jpn10")
                        display_id = str(cow_id) if cow_id else str(cow_auto_id)
                        milk_stats.append((milk_value_num, display_id))
                        if dim is not None:
                            # 産次を取得して3つのカテゴリに分類（1産、2産、3産以上）
                            parity = 0
                            if cow and cow.get("lact") is not None:
                                try:
                                    lact_val = int(cow.get("lact"))
                                    if lact_val == 1:
                                        parity = 1
                                    elif lact_val == 2:
                                        parity = 2
                                    elif lact_val >= 3:
                                        parity = 3
                                except (ValueError, TypeError):
                                    pass
                            cow_id = str(cow.get("cow_id") or "") if cow else ""
                            jpn10 = str(cow.get("jpn10") or "") if cow else ""
                            milk_scatter_points.append((dim, milk_value_num, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass
                scc_value = json_data.get("scc")
                if scc_value is not None and scc_value != "":
                    try:
                        scc_num = float(scc_value)
                        if scc_num > 0:
                            scc_values.append(scc_num)
                    except (ValueError, TypeError):
                        pass
                ls_value = json_data.get("ls")
                if ls_value is not None and ls_value != "":
                    try:
                        ls_value_num = float(ls_value)
                        ls_values.append(ls_value_num)
                        if dim is not None:
                            # 産次を取得して3つのカテゴリに分類（1産、2産、3産以上）
                            parity = 0
                            if cow and cow.get("lact") is not None:
                                try:
                                    lact_val = int(cow.get("lact"))
                                    if lact_val == 1:
                                        parity = 1
                                    elif lact_val == 2:
                                        parity = 2
                                    elif lact_val >= 3:
                                        parity = 3
                                except (ValueError, TypeError):
                                    pass
                            cow_id = str(cow.get("cow_id") or "") if cow else ""
                            jpn10 = str(cow.get("jpn10") or "") if cow else ""
                            ls_scatter_points.append((dim, ls_value_num, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass

            cow_count = len(cow_ids)
            avg_dim = round(sum(dim_values) / len(dim_values), 1) if dim_values else None
            avg_milk = round(sum(milk_values) / len(milk_values), 1) if milk_values else None
            max_milk = None
            max_milk_id = None
            min_milk = None
            min_milk_id = None
            if milk_stats:
                max_milk, max_milk_id = max(milk_stats, key=lambda x: x[0])
                min_milk, min_milk_id = min(milk_stats, key=lambda x: x[0])
            avg_scc = round(sum(scc_values) / len(scc_values), 1) if scc_values else None
            avg_ls = round(sum(ls_values) / len(ls_values), 1) if ls_values else None
            ls_total = len(ls_values)
            ls_under_2 = len([v for v in ls_values if v <= 2.0])
            ls_under_2_rate = round((ls_under_2 / ls_total) * 100, 1) if ls_total > 0 else None
            avg_parity = round(sum(parity_values) / len(parity_values), 1) if parity_values else None
            total_parity = sum(parity_counts.values())
            parity_segments = []
            if total_parity > 0:
                import math
                # 3つのカテゴリに対応する色のみを使用（1産、2産、3産以上）
                parity_colors = [
                    ("1", "#4ECDC4", "1産"),
                    ("2", "#45B7D1", "2産"),
                    ("3+", "#FF6B6B", "3産以上"),
                ]
                circumference = 2 * math.pi * 50
                offset = 0.0
                for key, color, label in parity_colors:
                    count = parity_counts.get(key, 0)
                    if count <= 0:
                        continue  # カウントが0の場合はスキップ
                    percent = (count / total_parity) * 100
                    length = (percent / 100) * circumference
                    parity_segments.append({
                        "color": color,
                        "length": length,
                        "offset": -offset,
                        "label": label,
                        "percent": round(percent, 1),
                        "count": count,
                    })
                    offset += length

            # 乳量分類（<20, 20-29.9, 30-39.9, 40-49.9, 50+）
            milk_bins = {
                "<20": 0,
                "20-29": 0,
                "30-39": 0,
                "40-49": 0,
                "50+": 0,
            }
            for value in milk_values:
                if value < 20:
                    milk_bins["<20"] += 1
                elif value < 30:
                    milk_bins["20-29"] += 1
                elif value < 40:
                    milk_bins["30-39"] += 1
                elif value < 50:
                    milk_bins["40-49"] += 1
                else:
                    milk_bins["50+"] += 1
            milk_total = sum(milk_bins.values())
            milk_segments = []
            if milk_total > 0:
                import math
                milk_colors = [
                    ("<20", "#f1c40f", "~20kg"),
                    ("20-29", "#2ecc71", "20kg台"),
                    ("30-39", "#16a085", "30kg台"),
                    ("40-49", "#3498db", "40kg台"),
                    ("50+", "#f39c12", "~50kg"),
                ]
                circumference = 2 * math.pi * 50
                offset = 0.0
                for key, color, label in milk_colors:
                    count = milk_bins.get(key, 0)
                    percent = (count / milk_total) * 100
                    length = (percent / 100) * circumference
                    milk_segments.append({
                        "color": color,
                        "length": length,
                        "offset": -offset,
                        "label": label,
                        "percent": round(percent, 1),
                        "count": count,
                    })
                    offset += length

            # リニアスコア分類（<=2, 3-4, >=5）
            ls_bins = {
                "ls_le2": 0,
                "ls_3_4": 0,
                "ls_ge5": 0,
            }
            for value in ls_values:
                if value <= 2:
                    ls_bins["ls_le2"] += 1
                elif value < 5:
                    ls_bins["ls_3_4"] += 1
                else:
                    ls_bins["ls_ge5"] += 1
            ls_total = sum(ls_bins.values())
            ls_segments = []
            if ls_total > 0:
                import math
                ls_colors = [
                    ("ls_le2", "#00b5ff", "LS2以下"),
                    ("ls_3_4", "#f5b700", "LS3-4"),
                    ("ls_ge5", "#e74c3c", "LS5以上"),
                ]
                circumference = 2 * math.pi * 50
                offset = 0.0
                for key, color, label in ls_colors:
                    count = ls_bins.get(key, 0)
                    percent = (count / ls_total) * 100
                    length = (percent / 100) * circumference
                    ls_segments.append({
                        "color": color,
                        "length": length,
                        "offset": -offset,
                        "label": label,
                        "percent": round(percent, 1),
                        "count": count,
                    })
                    offset += length

            avg_scc_display = None
            if avg_scc is not None:
                import math
                avg_scc_display = int(math.floor(avg_scc + 0.5))

            # 前回（対象日より前で最新）の乳検平均乳量
            previous_dates = sorted(
                {e.get("event_date") for e in events if e.get("event_date") and e.get("event_date") < selected_date},
                reverse=True
            )
            prev_avg_milk = None
            if previous_dates:
                prev_date = previous_dates[0]
                prev_events = [e for e in events if e.get("event_date") == prev_date]
                prev_values: List[float] = []
                for event in prev_events:
                    json_data = event.get("json_data") or {}
                    if isinstance(json_data, str):
                        try:
                            json_data = json.loads(json_data)
                        except Exception:
                            json_data = {}
                    milk_value = json_data.get("milk_yield")
                    if milk_value is not None and milk_value != "":
                        try:
                            prev_values.append(float(milk_value))
                        except (ValueError, TypeError):
                            pass
                if prev_values:
                    prev_avg_milk = round(sum(prev_values) / len(prev_values), 1)

            # 農場名と目標設定
            settings = SettingsManager(self.farm_path)
            farm_name = settings.get("farm_name", self.farm_path.name)
            farm_goals = settings.get("farm_goals", {}) or {}
            scc_threshold = farm_goals.get("individual_scc")
            if scc_threshold is None:
                scc_threshold = 200
            heifer_milk_target = farm_goals.get("milk_100d_heifer")
            if heifer_milk_target is None:
                heifer_milk_target = 30
            parous_milk_target = farm_goals.get("milk_100d_parous")
            if parous_milk_target is None:
                parous_milk_target = 40

            # 前月の乳検日（前月のカレンダー月で最新の検定日）と前月集計
            prev_month_date: Optional[str] = None
            prev_month_cow_count: Optional[int] = None
            prev_month_avg_milk: Optional[float] = None
            prev_month_avg_scc: Optional[float] = None
            prev_month_avg_ls: Optional[float] = None
            prev_month_ls_under_2_rate: Optional[float] = None
            prev_month_scc_over_count: Optional[int] = None
            if selected_dt:
                prev_month_first = (selected_dt.replace(day=1) - timedelta(days=1))
                prev_month_str = prev_month_first.strftime("%Y-%m")
                prev_month_dates = sorted(
                    [e.get("event_date") for e in events if e.get("event_date") and str(e.get("event_date", "")).startswith(prev_month_str)],
                    reverse=True
                )
                if prev_month_dates:
                    prev_month_date = prev_month_dates[0]
                    prev_month_events = [e for e in events if e.get("event_date") == prev_month_date]
                    prev_month_cow_count = len(prev_month_events)
                    prev_milk_vals: List[float] = []
                    prev_scc_vals: List[float] = []
                    prev_ls_vals: List[float] = []
                    prev_scc_over = 0
                    for e in prev_month_events:
                        j = e.get("json_data") or {}
                        if isinstance(j, str):
                            try:
                                j = json.loads(j)
                            except Exception:
                                j = {}
                        mv = j.get("milk_yield")
                        if mv not in ("", None):
                            try:
                                prev_milk_vals.append(float(mv))
                            except (ValueError, TypeError):
                                pass
                        sv = j.get("scc")
                        if sv not in ("", None):
                            try:
                                s = float(sv)
                                prev_scc_vals.append(s)
                                if scc_threshold is not None and s > scc_threshold:
                                    prev_scc_over += 1
                            except (ValueError, TypeError):
                                pass
                        lv = j.get("ls")
                        if lv not in ("", None):
                            try:
                                prev_ls_vals.append(float(lv))
                            except (ValueError, TypeError):
                                pass
                    if prev_milk_vals:
                        prev_month_avg_milk = round(sum(prev_milk_vals) / len(prev_milk_vals), 1)
                    if prev_scc_vals:
                        prev_month_avg_scc = round(sum(prev_scc_vals) / len(prev_scc_vals), 1)
                    if prev_ls_vals:
                        prev_month_avg_ls = round(sum(prev_ls_vals) / len(prev_ls_vals), 1)
                        prev_ls_under_2 = len([v for v in prev_ls_vals if v <= 2.0])
                        prev_month_ls_under_2_rate = round((prev_ls_under_2 / len(prev_ls_vals)) * 100, 1)
                    prev_month_scc_over_count = prev_scc_over

            # 乳質（SCC目標超え）リスト
            scc_over_rows = []
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                scc_value = json_data.get("scc")
                try:
                    scc_num = float(scc_value) if scc_value not in ("", None) else None
                except (ValueError, TypeError):
                    scc_num = None
                if scc_num is None or scc_num <= scc_threshold:
                    continue

                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                cow_id = cow.get("cow_id", "") if cow else ""
                jpn10 = cow.get("jpn10", "") if cow else ""
                lact = cow.get("lact", "") if cow else ""
                clvd = clvd_for_milk_dim(cow_auto_id, cow)
                dim = calc_dim_on_date(clvd, selected_date)
                milk_value = json_data.get("milk_yield")
                try:
                    milk_num = float(milk_value) if milk_value not in ("", None) else None
                except (ValueError, TypeError):
                    milk_num = None

                # 前月SCC（前回乳検イベント）
                prev_scc = None
                history = event_history.get(cow_auto_id, [])
                for event_dt, hist_scc in reversed(history):
                    if event_dt.strftime("%Y-%m-%d") < selected_date:
                        prev_scc = hist_scc
                        break

                def fmt_value(val):
                    if val is None or val == "":
                        return "-"
                    if isinstance(val, float) and val.is_integer():
                        return str(int(val))
                    return str(val)
                
                def fmt_milk_fat(val):
                    """乳量・乳脂率を小数点以下1桁で統一表示（40 → 40.0）"""
                    if val is None or val == "":
                        return "-"
                    try:
                        num = float(val)
                        return f"{num:.1f}"
                    except (ValueError, TypeError):
                        return str(val) if val != "" else "-"

                scc_class = " scc-alert" if scc_num is not None and scc_num >= 200 else ""
                prev_scc_class = " scc-alert" if prev_scc is not None and prev_scc >= 200 else ""
                rc_val = cow.get("rc") if cow else None
                id_pregnant = " id-pregnant" if (rc_val == 5 or rc_val == "5") else ""
                scc_over_rows.append({
                    "cow_id": fmt_value(cow_id),
                    "jpn10": fmt_value(jpn10),
                    "lact": fmt_value(lact),
                    "dim": fmt_value(dim),
                    "milk": fmt_milk_fat(milk_num),
                    "scc": fmt_value(scc_num),
                    "prev_scc": fmt_value(prev_scc),
                    "scc_class": scc_class,
                    "prev_scc_class": prev_scc_class,
                    "id_cell_class": id_pregnant,
                })

            # 初産成績（DIM100以内 & 目標乳量以下）
            heifer_under_rows = []
            heifer_scatter_points: List[Tuple[int, float]] = []
            parous_under_rows = []
            parous_scatter_points: List[Tuple[int, float]] = []
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                lact = None
                if cow and cow.get("lact") is not None:
                    try:
                        lact = int(cow.get("lact"))
                    except (ValueError, TypeError):
                        lact = None
                clvd = clvd_for_milk_dim(cow_auto_id, cow)
                dim = calc_dim_on_date(clvd, selected_date)
                if dim is None:
                    continue
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                milk_value = json_data.get("milk_yield")
                try:
                    milk_num = float(milk_value) if milk_value not in ("", None) else None
                except (ValueError, TypeError):
                    milk_num = None
                if milk_num is None:
                    continue
                if lact == 1:
                    heifer_scatter_points.append((dim, milk_num))
                elif lact is not None and lact >= 2:
                    parous_scatter_points.append((dim, milk_num))

                if dim > 100:
                    continue

                cow_id = cow.get("cow_id", "") if cow else ""
                jpn10 = cow.get("jpn10", "") if cow else ""
                scc_value = json_data.get("scc")
                fat_value = json_data.get("fat")
                bhb_value = json_data.get("bhb")
                denovo_value = json_data.get("denovo_fa")
                try:
                    scc_num = float(scc_value) if scc_value not in ("", None) else None
                except (ValueError, TypeError):
                    scc_num = None
                try:
                    fat_num = float(fat_value) if fat_value not in ("", None) else None
                except (ValueError, TypeError):
                    fat_num = None
                try:
                    bhb_num = float(bhb_value) if bhb_value not in ("", None) else None
                except (ValueError, TypeError):
                    bhb_num = None
                try:
                    denovo_num = float(denovo_value) if denovo_value not in ("", None) else None
                except (ValueError, TypeError):
                    denovo_num = None
                rc_val = cow.get("rc") if cow else None
                id_pregnant = " id-pregnant" if (rc_val == 5 or rc_val == "5") else ""
                if lact == 1:
                    if milk_num > heifer_milk_target:
                        continue
                    heifer_under_rows.append({
                        "cow_id": fmt_value(cow_id),
                        "jpn10": fmt_value(jpn10),
                        "lact": fmt_value(lact),
                        "dim": fmt_value(dim),
                        "milk": fmt_milk_fat(milk_num),
                        "scc": fmt_value(scc_num),
                        "fat": fmt_milk_fat(fat_num),
                        "bhb": fmt_value(bhb_num),
                        "bhb_class": " bhb-alert" if bhb_num is not None and bhb_num >= 0.13 else "",
                        "scc_class": " scc-alert" if scc_num is not None and scc_num >= 200 else "",
                        "denovo": fmt_value(denovo_num),
                        "id_cell_class": id_pregnant,
                    })
                elif lact is not None and lact >= 2:
                    if milk_num > parous_milk_target:
                        continue
                    parous_under_rows.append({
                        "cow_id": fmt_value(cow_id),
                        "jpn10": fmt_value(jpn10),
                        "lact": fmt_value(lact),
                        "dim": fmt_value(dim),
                        "milk": fmt_milk_fat(milk_num),
                        "scc": fmt_value(scc_num),
                        "fat": fmt_milk_fat(fat_num),
                        "bhb": fmt_value(bhb_num),
                        "bhb_class": " bhb-alert" if bhb_num is not None and bhb_num >= 0.13 else "",
                        "scc_class": " scc-alert" if scc_num is not None and scc_num >= 200 else "",
                        "denovo": fmt_value(denovo_num),
                        "id_cell_class": id_pregnant,
                    })

            # 異常牛検知分析
            milk_drop_rows = []
            bhb_over_rows = []
            denovo_under_rows = []
            denovo_under_61_rows = []
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                lact = None
                if cow and cow.get("lact") is not None:
                    try:
                        lact = int(cow.get("lact"))
                    except (ValueError, TypeError):
                        lact = None
                clvd = clvd_for_milk_dim(cow_auto_id, cow)
                dim = calc_dim_on_date(clvd, selected_date)
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                milk_value = json_data.get("milk_yield")
                try:
                    milk_num = float(milk_value) if milk_value not in ("", None) else None
                except (ValueError, TypeError):
                    milk_num = None
                scc_value = json_data.get("scc")
                fat_value = json_data.get("fat")
                bhb_value = json_data.get("bhb")
                denovo_value = json_data.get("denovo_fa")
                try:
                    scc_num = float(scc_value) if scc_value not in ("", None) else None
                except (ValueError, TypeError):
                    scc_num = None
                try:
                    fat_num = float(fat_value) if fat_value not in ("", None) else None
                except (ValueError, TypeError):
                    fat_num = None
                try:
                    bhb_num = float(bhb_value) if bhb_value not in ("", None) else None
                except (ValueError, TypeError):
                    bhb_num = None
                try:
                    denovo_num = float(denovo_value) if denovo_value not in ("", None) else None
                except (ValueError, TypeError):
                    denovo_num = None

                cow_id = cow.get("cow_id", "") if cow else ""
                jpn10 = cow.get("jpn10", "") if cow else ""
                rc_val = cow.get("rc") if cow else None
                id_pregnant = " id-pregnant" if (rc_val == 5 or rc_val == "5") else ""

                bhb_class = " bhb-alert" if bhb_num is not None and bhb_num >= 0.13 else ""

                # 前月から15%以上乳量が減少
                prev_milk = None
                history = milk_history.get(cow_auto_id, [])
                if selected_dt and history:
                    for event_dt, hist_milk in reversed(history):
                        if event_dt < selected_dt:
                            prev_milk = hist_milk
                            break
                scc_class = " scc-alert" if scc_num is not None and scc_num >= 200 else ""
                if milk_num is not None and prev_milk is not None and prev_milk > 0:
                    if milk_num <= prev_milk * 0.85:
                        milk_drop_rows.append({
                            "cow_id": fmt_value(cow_id),
                            "jpn10": fmt_value(jpn10),
                            "lact": fmt_value(lact),
                            "dim": fmt_value(dim),
                            "milk": fmt_milk_fat(milk_num),
                            "prev_milk": fmt_milk_fat(prev_milk),
                            "scc": fmt_value(scc_num),
                            "fat": fmt_milk_fat(fat_num),
                            "bhb": fmt_value(bhb_num),
                            "bhb_class": bhb_class,
                            "scc_class": scc_class,
                            "denovo": fmt_value(denovo_num),
                            "id_cell_class": id_pregnant,
                        })

                # BHBが0.13以上
                if bhb_num is not None and bhb_num >= 0.13:
                    bhb_over_rows.append({
                        "cow_id": fmt_value(cow_id),
                        "jpn10": fmt_value(jpn10),
                        "lact": fmt_value(lact),
                        "dim": fmt_value(dim),
                        "milk": fmt_milk_fat(milk_num),
                        "prev_milk": fmt_milk_fat(prev_milk),
                        "scc": fmt_value(scc_num),
                        "fat": fmt_milk_fat(fat_num),
                        "bhb": fmt_value(bhb_num),
                        "bhb_class": bhb_class,
                        "scc_class": scc_class,
                        "denovo": fmt_value(denovo_num),
                        "id_cell_class": id_pregnant,
                    })

                # DIMが60日以内かつデノボFAが22%未満
                if dim is not None and dim <= 60 and denovo_num is not None and denovo_num < 22:
                    denovo_under_rows.append({
                        "cow_id": fmt_value(cow_id),
                        "jpn10": fmt_value(jpn10),
                        "lact": fmt_value(lact),
                        "dim": fmt_value(dim),
                        "milk": fmt_milk_fat(milk_num),
                        "prev_milk": fmt_milk_fat(prev_milk),
                        "scc": fmt_value(scc_num),
                        "fat": fmt_milk_fat(fat_num),
                        "bhb": fmt_value(bhb_num),
                        "bhb_class": bhb_class,
                        "scc_class": scc_class,
                        "denovo": fmt_value(denovo_num),
                        "id_cell_class": id_pregnant,
                    })

                # DIMが61日以上かつデノボFAが28%未満
                if dim is not None and dim >= 61 and denovo_num is not None and denovo_num < 28:
                    denovo_under_61_rows.append({
                        "cow_id": fmt_value(cow_id),
                        "jpn10": fmt_value(jpn10),
                        "lact": fmt_value(lact),
                        "dim": fmt_value(dim),
                        "milk": fmt_milk_fat(milk_num),
                        "prev_milk": fmt_milk_fat(prev_milk),
                        "scc": fmt_value(scc_num),
                        "fat": fmt_milk_fat(fat_num),
                        "bhb": fmt_value(bhb_num),
                        "bhb_class": bhb_class,
                        "scc_class": scc_class,
                        "denovo": fmt_value(denovo_num),
                        "id_cell_class": id_pregnant,
                    })

            # 前月の個体別該当リスト（AI用：個体ごとの前月→今月コメント用）
            prev_month_scc_over_cow_ids: List[str] = []
            prev_month_bhb_over_cow_ids: List[str] = []
            if prev_month_date and prev_month_events:
                for e in prev_month_events:
                    cow_auto_id = e.get("cow_auto_id")
                    if not cow_auto_id:
                        continue
                    cow = self.db.get_cow_by_auto_id(cow_auto_id)
                    cow_id = (cow.get("cow_id") or "") if cow else ""
                    if not cow_id:
                        continue
                    j = e.get("json_data") or {}
                    if isinstance(j, str):
                        try:
                            j = json.loads(j)
                        except Exception:
                            j = {}
                    scc_val = j.get("scc")
                    if scc_val not in ("", None):
                        try:
                            if float(scc_val) > scc_threshold:
                                prev_month_scc_over_cow_ids.append(str(cow_id))
                        except (ValueError, TypeError):
                            pass
                    bhb_val = j.get("bhb")
                    if bhb_val not in ("", None):
                        try:
                            if float(bhb_val) >= 0.13:
                                prev_month_bhb_over_cow_ids.append(str(cow_id))
                        except (ValueError, TypeError):
                            pass

            this_month_all_cow_ids = set()
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                cid = (cow.get("cow_id") or "") if cow else ""
                if cid:
                    this_month_all_cow_ids.add(str(cid))
            this_month_scc_ids = [r["cow_id"] for r in scc_over_rows]
            this_month_bhb_ids = [r["cow_id"] for r in bhb_over_rows]
            this_month_milk_drop_ids = [r["cow_id"] for r in milk_drop_rows]
            set_prev_scc = set(prev_month_scc_over_cow_ids)
            set_prev_bhb = set(prev_month_bhb_over_cow_ids)
            set_this_scc = set(this_month_scc_ids)
            set_this_bhb = set(this_month_bhb_ids)
            scc_improved_ids = sorted(set_prev_scc & this_month_all_cow_ids - set_this_scc)
            scc_worsened_ids = sorted(set_this_scc - set_prev_scc)
            bhb_improved_ids = sorted(set_prev_bhb & this_month_all_cow_ids - set_this_bhb)
            bhb_worsened_ids = sorted(set_this_bhb - set_prev_bhb)

            # 個体別前月→今月の変化テキスト（AIプロンプト用・数値付き）
            individual_comparison_lines: List[str] = []
            if prev_month_date:
                if scc_improved_ids:
                    individual_comparison_lines.append(f"・SCC目標超え：前月該当→今月は該当外（改善）: ID {', '.join(scc_improved_ids)}")
                if scc_worsened_ids:
                    scc_parts = [f"ID {r['cow_id']}（SCC={r['scc']}）" for r in scc_over_rows if r["cow_id"] in scc_worsened_ids]
                    individual_comparison_lines.append(f"・SCC目標超え：今月該当（前月は該当しなかった）: {', '.join(scc_parts)}")
                if bhb_improved_ids:
                    individual_comparison_lines.append(f"・BHB0.13以上：前月該当→今月は該当外（改善）: ID {', '.join(bhb_improved_ids)}")
                if bhb_worsened_ids:
                    bhb_parts = [f"ID {r['cow_id']}（BHB={r['bhb']}）" for r in bhb_over_rows if r["cow_id"] in bhb_worsened_ids]
                    individual_comparison_lines.append(f"・BHB0.13以上：今月該当（前月は該当しなかった）: {', '.join(bhb_parts)}")
                if this_month_milk_drop_ids:
                    individual_comparison_lines.append(f"・乳量15%減：今月該当: ID {', '.join(this_month_milk_drop_ids)}")

            # リニアスコア分布（初産/経産 × DIM区分）
            def classify_dim_stage(dim_value: Optional[int]) -> Optional[str]:
                if dim_value is None:
                    return None
                if dim_value < 60:
                    return "lt60"
                if dim_value <= 120:
                    return "60_120"
                return "gt120"

            def classify_ls(ls_value: Optional[float]) -> Optional[str]:
                if ls_value is None:
                    return None
                if ls_value <= 2:
                    return "ls_le2"
                if ls_value < 5:
                    return "ls_3_4"
                return "ls_ge5"

            ls_stage_buckets = {
                "heifer": {"lt60": {"classes": [], "values": []}, "60_120": {"classes": [], "values": []}, "gt120": {"classes": [], "values": []}},
                "parous": {"lt60": {"classes": [], "values": []}, "60_120": {"classes": [], "values": []}, "gt120": {"classes": [], "values": []}},
            }
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                if not cow_auto_id:
                    continue
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                lact = None
                if cow and cow.get("lact") is not None:
                    try:
                        lact = int(cow.get("lact"))
                    except (ValueError, TypeError):
                        lact = None
                clvd = clvd_for_milk_dim(cow_auto_id, cow)
                dim = calc_dim_on_date(clvd, selected_date)
                stage = classify_dim_stage(dim)
                if stage is None or lact is None:
                    continue
                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}
                ls_value = json_data.get("ls")
                try:
                    ls_num = float(ls_value) if ls_value not in ("", None) else None
                except (ValueError, TypeError):
                    ls_num = None
                ls_class = classify_ls(ls_num)
                if ls_class is None:
                    continue
                bucket_key = "heifer" if lact == 1 else "parous"
                if bucket_key not in ls_stage_buckets:
                    continue
                ls_stage_buckets[bucket_key][stage]["classes"].append(ls_class)
                if ls_num is not None:
                    ls_stage_buckets[bucket_key][stage]["values"].append(ls_num)

            def build_ls_segments(classes: List[str], values: List[float]) -> Tuple[List[Dict[str, Any]], int, Optional[float]]:
                total = len(classes)
                counts = {
                    "ls_le2": classes.count("ls_le2"),
                    "ls_3_4": classes.count("ls_3_4"),
                    "ls_ge5": classes.count("ls_ge5"),
                }
                segments = []
                avg_ls = None
                if total > 0:
                    import math
                    colors = [
                        ("ls_le2", "#bdbdbd", "LS2以下"),
                        ("ls_3_4", "#f5b700", "LS3-4"),
                        ("ls_ge5", "#e74c3c", "LS5以上"),
                    ]
                    circumference = 2 * math.pi * 50
                    offset = 0.0
                    for key, color, label in colors:
                        count = counts.get(key, 0)
                        percent = (count / total) * 100 if total > 0 else 0
                        length = (percent / 100) * circumference
                        segments.append({
                            "color": color,
                            "length": length,
                            "offset": -offset,
                            "label": label,
                            "percent": round(percent, 1),
                            "count": count,
                        })
                        offset += length
                    # 平均値を計算
                    if values:
                        avg_ls = sum(values) / len(values)
                return segments, total, avg_ls

            ls_donut_data = {
                "heifer": {
                    "lt60": build_ls_segments(ls_stage_buckets["heifer"]["lt60"]["classes"], ls_stage_buckets["heifer"]["lt60"]["values"]),
                    "60_120": build_ls_segments(ls_stage_buckets["heifer"]["60_120"]["classes"], ls_stage_buckets["heifer"]["60_120"]["values"]),
                    "gt120": build_ls_segments(ls_stage_buckets["heifer"]["gt120"]["classes"], ls_stage_buckets["heifer"]["gt120"]["values"]),
                },
                "parous": {
                    "lt60": build_ls_segments(ls_stage_buckets["parous"]["lt60"]["classes"], ls_stage_buckets["parous"]["lt60"]["values"]),
                    "60_120": build_ls_segments(ls_stage_buckets["parous"]["60_120"]["classes"], ls_stage_buckets["parous"]["60_120"]["values"]),
                    "gt120": build_ls_segments(ls_stage_buckets["parous"]["gt120"]["classes"], ls_stage_buckets["parous"]["gt120"]["values"]),
                },
            }

            # 生成AIで前月比較を取得（use_ai=True かつ API利用可能時のみ）
            ai_comparison_bullets = ""
            if use_ai:
                try:
                    client = ChatGPTClient()
                    cur = [
                        f"検定日: {selected_date}",
                        f"検定牛頭数: {cow_count} 頭",
                        f"平均搾乳日数(DIM): {avg_dim if avg_dim is not None else '-'}",
                        f"平均乳量: {f'{avg_milk:.1f}' if avg_milk is not None else '-'} kg（前回: {f'{prev_avg_milk:.1f}' if prev_avg_milk is not None else '-'} kg）",
                        f"最高乳量: {f'{max_milk:.1f}' if max_milk is not None else '-'} kg、最低: {f'{min_milk:.1f}' if min_milk is not None else '-'} kg",
                        f"平均体細胞: {avg_scc_display if avg_scc_display is not None else '-'} 千、平均リニアスコア: {avg_ls if avg_ls is not None else '-'}、LS2以下割合: {ls_under_2_rate if ls_under_2_rate is not None else '-'}%",
                        f"SCC目標超え: {len(scc_over_rows)} 頭、初産DIM100以内目標未満: {len(heifer_under_rows)} 頭、経産同: {len(parous_under_rows)} 頭",
                        f"乳量15%減: {len(milk_drop_rows)} 頭、BHB0.13以上: {len(bhb_over_rows)} 頭、デノボ不足(60日以内22%未満): {len(denovo_under_rows)} 頭、61日以上28%未満: {len(denovo_under_61_rows)} 頭",
                    ]
                    prev_lines: List[str] = []
                    if prev_month_date and prev_month_cow_count is not None:
                        prev_lines = [
                            f"前月検定日: {prev_month_date}",
                            f"前月検定頭数: {prev_month_cow_count} 頭",
                            f"前月平均乳量: {prev_month_avg_milk if prev_month_avg_milk is not None else '-'} kg",
                            f"前月平均体細胞: {prev_month_avg_scc if prev_month_avg_scc is not None else '-'} 千、平均LS: {prev_month_avg_ls if prev_month_avg_ls is not None else '-'}、LS2以下割合: {prev_month_ls_under_2_rate if prev_month_ls_under_2_rate is not None else '-'}%",
                            f"前月SCC目標超え: {prev_month_scc_over_count if prev_month_scc_over_count is not None else '-'} 頭",
                        ]
                    user_text = "【今月の乳検サマリ】\n" + "\n".join(cur)
                    if prev_lines:
                        user_text += "\n\n【前月の乳検サマリ】\n" + "\n".join(prev_lines)
                    else:
                        user_text += "\n\n【前月】データなし"
                    if prev_lines and individual_comparison_lines:
                        user_text += "\n\n【個体別の前月→今月の変化（参考）】\n" + "\n".join(individual_comparison_lines)
                    system_prompt = (
                        "あなたは酪農・乳検データの分析専門家です。"
                        "指示に従い、日本語で箇条書きのみを出力してください。"
                        "文体は「です」「ます」の敬体で統一してください。"
                        "番号や見出しは不要です。各項目は「・」で始める短い文にしてください。"
                    )
                    if prev_lines:
                        # 前月との比較：集計の変化＋個体ごとのコメント（ID付きで前月→今月の状態を書く）
                        comparison_prompt = (
                            "以下は今月と前月の乳検サマリと、個体別の前月→今月の変化（該当する場合）です。\n\n"
                            "「前月との比較」を箇条書きで書いてください。文末は「です」「ます」で統一してください。\n"
                            "（1）集計の変化：検定頭数・平均乳量・平均体細胞などの前月からの増減を、前月の数値と対比して書いてください。\n"
                            "（2）個体ごとのコメント：必ず含めてください。以下のような形で、個体IDを明示し、前月と今月の状態の違いが分かるように書いてください。\n"
                            "　例：「前月はID100、200、201がSCC200以上でしたが、今月は200以下になりました。」\n"
                            "　例：「ID200は前月はBHBで引っかかりませんでしたが、今月は引っかかっています（BHB＝2.0）。」\n"
                            "　例：「ID50は前月は乳量15%減で該当していましたが、今月は該当外になりました。」\n"
                            "改善した個体（前月該当→今月該当外）と、悪化した個体（前月は該当しなかったが今月該当）の両方を、SCC・乳量・BHBごとに箇条書きで書いてください。項目数は８項目以内を目安にしてください。\n\n" + user_text
                        )
                        ai_comparison_bullets = client.ask(system_prompt, comparison_prompt)
                    else:
                        ai_comparison_bullets = "・前月の乳検データがないため比較できません。"
                except Exception as ai_err:
                    logging.debug("乳検レポートAI要約スキップ: %s", ai_err)

            # 前月との比較セクション（use_ai 時のみ出力）
            ai_section_html = ""
            if use_ai:
                ai_bullets_text = html.escape(ai_comparison_bullets) if ai_comparison_bullets else "前月データがないため比較できません。"
                ai_section_html = (
                    '<div class="report-layout page-break-before ai-summary-page">'
                    '<div style="width: 100%;">'
                    '<div class="section-title">前月との比較</div>'
                    '<div class="ai-summary-card">'
                    f'<p class="ai-bullets">{ai_bullets_text}</p>'
                    '</div></div></div>'
                )

            # 乳量・リニアスコア散布図（Plotly でホバーID表示・2グラフ連動）
            milk_report_scatter_html = ""
            try:
                from modules.genome_report_html import _get_plotly_scatter_script as _report_plotly_script
            except ImportError:
                _report_plotly_script = None
            if _report_plotly_script and (milk_scatter_points or ls_scatter_points):
                def _unpack_pt(p):
                    if len(p) >= 5:
                        return (p[0], p[1], p[2], str(p[3]), str(p[4]))
                    return (p[0], p[1], p[2], "", "")
                configs = []
                if milk_scatter_points:
                    m_pts = [_unpack_pt(p) for p in milk_scatter_points]
                    configs.append({
                        "id": "scatter-report-milk",
                        "data": {
                            "x": [p[0] for p in m_pts], "y": [p[1] for p in m_pts],
                            "parity": [p[2] for p in m_pts], "cow_ids": [p[3] for p in m_pts], "jpn10s": [p[4] for p in m_pts],
                            "title": "乳量", "xlabel": "DIM", "ylabel": "乳量", "year_min": 0, "year_max": 400,
                            "y_start_zero": True,
                        },
                    })
                if ls_scatter_points:
                    l_pts = [_unpack_pt(p) for p in ls_scatter_points]
                    configs.append({
                        "id": "scatter-report-ls",
                        "data": {
                            "x": [p[0] for p in l_pts], "y": [p[1] for p in l_pts],
                            "parity": [p[2] for p in l_pts], "cow_ids": [p[3] for p in l_pts], "jpn10s": [p[4] for p in l_pts],
                            "title": "リニアスコア", "xlabel": "DIM", "ylabel": "リニアスコア", "year_min": 0, "year_max": 400,
                            "y_start_zero": True,
                        },
                    })
                if configs:
                    config_json = json.dumps(configs, ensure_ascii=False).replace("</", "<\\/")
                    script = _report_plotly_script(config_json)
                    parts = []
                    for c in configs:
                        tid = c["id"]
                        title = "乳量" if "milk" in tid else "リニアスコア"
                        parts.append(
                            f'<div class="scatter-card"><div class="chart-title">{html.escape(title)}</div>'
                            f'<div class="scatter-chart"><div id="{html.escape(tid)}" class="plotly-scatter-div" style="min-height:260px;"></div></div></div>'
                        )
                    milk_report_scatter_html = '<div class="scatter-stack">' + "".join(parts) + '</div>\n' + script
            if not milk_report_scatter_html:
                milk_report_scatter_html = (
                    '<div class="scatter-stack">'
                    '<div class="scatter-card"><div class="chart-title">乳量</div><div class="scatter-chart">'
                    + self._build_milk_scatter_svg([(p[0], p[1], p[2]) for p in milk_scatter_points])
                    + '</div></div><div class="scatter-card"><div class="chart-title">リニアスコア</div><div class="scatter-chart">'
                    + self._build_ls_scatter_svg([(p[0], p[1], p[2]) for p in ls_scatter_points])
                    + '</div></div></div>'
                )

            trend_rows = compute_monthly_milk_trend(events, selected_date, 5)
            monthly_trend_html = build_monthly_trend_section_html(trend_rows)
            _comment_storage_key = make_milk_report_comment_storage_key(farm_name, selected_date)
            _comment_storage_key_js = json.dumps(_comment_storage_key, ensure_ascii=False)
            milk_report_comment_modal_html = build_comment_modal_html()
            milk_report_comment_script = build_comment_script(_comment_storage_key_js)

            # HTML生成
            try:
                from modules.genome_report_html import REPORT_BASE_CSS
            except ImportError:
                REPORT_BASE_CSS = ""
            milk_report_extra_css = """
        .report-layout { display: flex; gap: 24px; align-items: flex-start; flex-wrap: wrap; }
        .chart-stack { display: flex; flex-direction: column; gap: 16px; }
        .scatter-stack { display: flex; flex-direction: column; gap: 16px; margin-top: 16px; }
        .scatter-chart { margin-top: 4px; }
        .summary-table { width: 420px; }
        .summary-table th { width: 200px; text-align: right; font-weight: 600; }
        .spacer td { border-bottom: none; padding: 8px 0; }
        .donut-wrap { display: flex; flex-direction: column; align-items: center; }
        .legend { margin-top: 8px; font-size: 11px; }
        .legend-row { display: grid; grid-template-columns: 1fr 110px; gap: 8px; align-items: center; }
        .legend-row span:last-child { text-align: right; }
        .legend-value { display: grid; grid-template-columns: 52px 1fr; column-gap: 6px; }
        .legend-count, .legend-percent { text-align: right; font-variant-numeric: tabular-nums; }
        .legend-dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 6px; }
        .donut-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px; }
        .donut-card .chart-title { margin-bottom: 6px; font-size: 12px; }
        .page-break-after { break-after: page; page-break-after: always; }
        .page-break-before { break-before: page; page-break-before: always; }
        .summary-table thead { display: table-header-group; }
        .summary-table tbody { display: table-row-group; }
        .ai-summary-page { max-width: 600px; margin-top: 24px; }
        .ai-summary-card { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; }
        .ai-summary-card h3 { font-size: 14px; font-weight: bold; color: #212529; margin: 0 0 10px 0; padding-bottom: 6px; border-bottom: 1px solid #dee2e6; }
        .ai-bullets { white-space: pre-line; font-size: 12px; line-height: 1.6; color: #495057; margin: 0; }
        .ai-bullets.subnote { font-size: 11px; color: #6c757d; }
        /* 乳検レポート1ページ目：概要・散布図・ドーナツを1ページに収める */
        .milk-report-first-page { page-break-inside: avoid; }
        .milk-report-first-page.report-layout { gap: 12px; }
        .milk-report-first-page .summary-table { width: 360px; font-size: 10px; }
        .milk-report-first-page .summary-table th,
        .milk-report-first-page .summary-table td { padding: 4px 6px; }
        .milk-report-first-page .scatter-stack { gap: 8px; margin-top: 8px; }
        .milk-report-first-page .scatter-card { padding: 8px; }
        .milk-report-first-page .chart-stack { gap: 8px; }
        .milk-report-first-page .chart-card { padding: 8px; }
        .milk-report-first-page .donut-wrap svg { width: 100px; height: 100px; }
        .milk-report-first-page .chart-title { font-size: 11px; margin-bottom: 4px; }
        .milk-report-first-page .legend { margin-top: 4px; font-size: 10px; }
        @media print {
          @page { size: A4 portrait; margin: 10mm 11mm; }
          body { padding: 0 !important; background: #fff !important; font-size: 10px !important; }
          .report-container { padding: 8px 10px !important; max-width: none !important; box-shadow: none !important;
            border-radius: 0 !important; }
          .report-header { margin-bottom: 8px !important; padding-bottom: 8px !important; }
          .report-brand { margin-top: 2px !important; }
          .report-container > p.subnote { margin: 0 0 6px !important; }
          .summary-table { margin-bottom: 10px !important; }
          .summary-table th, .summary-table td { padding: 3px 5px !important; }
          .section-title { margin: 10px 0 5px !important; font-size: 0.95rem !important; padding-bottom: 2px !important; }
          .subheader { margin-bottom: 5px !important; }
          .report-layout { gap: 10px !important; }
          .donut-grid { gap: 6px !important; margin-top: 6px !important; }
          .chart-stack { gap: 5px !important; }
          .scatter-stack { gap: 5px !important; margin-top: 6px !important; }
          .scatter-card, .chart-card, .donut-card { padding: 6px !important; }
          .plotly-scatter-div { min-height: 200px !important; }
          .ai-summary-page { margin-top: 8px !important; max-width: none !important; }
          .ai-summary-card { padding: 8px 10px !important; margin-bottom: 8px !important; }
          .ai-bullets { font-size: 10px !important; line-height: 1.45 !important; }
          .milk-report-first-page .summary-table { font-size: 9px; }
          .milk-report-first-page .summary-table th,
          .milk-report-first-page .summary-table td { padding: 2px 4px; }
          .milk-report-first-page .donut-wrap svg { width: 88px; height: 88px; }
        }
            """ + MILK_REPORT_COMMENT_CSS
            milk_report_css = (REPORT_BASE_CSS + milk_report_extra_css) if REPORT_BASE_CSS else "body{font-family:'Meiryo','Yu Gothic',sans-serif;font-size:12px;}.report-container{max-width:1200px;margin:0 auto;}.report-header{border-bottom:2px solid #0d6efd;padding-bottom:16px;}.section-title{color:#0d6efd;}.summary-table th{background:#e9ecef;border:1px solid #dee2e6;}"

            try:
                from modules.report_cow_bridge import DEFAULT_PORT as _report_open_cow_port
            except ImportError:
                _report_open_cow_port = 51985
            html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>乳検分析レポート</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Yomogi&family=Hachi+Maru+Pop&family=Klee+One&family=Zen+Kurenaido&display=swap" rel="stylesheet">
    <script>var FALCON_OPEN_COW_PORT = {_report_open_cow_port};</script>
    <style>
        {milk_report_css}
    </style>
</head>
<body>
<div class="report-container">
<header class="report-header">
  <dl>
    <dt>農場名</dt><dd>{html.escape(str(farm_name))}</dd>
    <dt>検定日</dt><dd>{html.escape(str(selected_date))}</dd>
  </dl>
  <p class="report-brand">FALCON 乳検分析レポート</p>
</header>
<p class="subnote" style="margin: 0 0 12px 0;">IDが<span style="background:#d4edda; padding: 0 4px;">薄い緑</span>の個体は妊娠中（RC=5）です。治療するか・廃用予定にするかの判断の参考にしてください。</p>
<div id="main-content">
<div class="milk-report-toolbar">
  <button type="button" class="btn-ghost btn-comment-mode" onclick="toggleCommentMode()">✏ コメントモード</button>
</div>
<div id="s-overview" class="report-layout milk-report-first-page page-break-after">
        <div>
            <table class="summary-table">
        <tr>
            <th>検定牛頭数</th>
            <td>{cow_count} 頭</td>
        </tr>
        <tr>
            <th>平均搾乳日数（DIM）</th>
            <td>{html.escape(str(avg_dim)) if avg_dim is not None else "-"}</td>
        </tr>
        <tr>
            <th>平均乳量</th>
            <td>
                {html.escape(f"{avg_milk:.1f}") if avg_milk is not None else "-"} kg<br>
                <span class="subnote">（前回乳量：{html.escape(f"{prev_avg_milk:.1f}") if prev_avg_milk is not None else "-"} kg）</span>
            </td>
        </tr>
        <tr>
            <th>最高乳量</th>
            <td>{html.escape(f"{max_milk:.1f}") if max_milk is not None else "-"} kg（{html.escape(str(max_milk_id)) if max_milk_id is not None else "-"}）</td>
        </tr>
        <tr>
            <th>最低乳量</th>
            <td>{html.escape(f"{min_milk:.1f}") if min_milk is not None else "-"} kg（{html.escape(str(min_milk_id)) if min_milk_id is not None else "-"}）</td>
        </tr>
        <tr class="spacer">
            <td colspan="2"></td>
        </tr>
        <tr>
            <th>平均体細胞</th>
            <td>{html.escape(str(avg_scc_display)) if avg_scc_display is not None else "-"} 千</td>
        </tr>
        <tr>
            <th>平均ﾘﾆｱｽｺｱ</th>
            <td>{html.escape(str(avg_ls)) if avg_ls is not None else "-"}</td>
        </tr>
        <tr>
            <th>ﾘﾆｱｽｺｱ2以下の割合</th>
            <td>{html.escape(str(ls_under_2_rate)) if ls_under_2_rate is not None else "-"} %</td>
        </tr>
            </table>
            {milk_report_scatter_html}
        </div>
        <div class="chart-stack">
            <div class="chart-card">
                <div class="chart-title">産次</div>
                <div class="donut-wrap">
                    {self._build_parity_donut_svg(parity_segments, avg_parity, total_parity)}
                    {self._build_parity_legend(parity_segments, total_parity)}
                </div>
            </div>
            <div class="chart-card">
                <div class="chart-title">乳量</div>
                <div class="donut-wrap">
                    {self._build_milk_donut_svg(milk_segments, avg_milk, milk_total)}
                    {self._build_milk_legend(milk_segments, milk_total)}
                </div>
            </div>
            <div class="chart-card">
                <div class="chart-title">リニアスコア</div>
                <div class="donut-wrap">
                    {self._build_ls_donut_svg(ls_segments, avg_ls, ls_total)}
                    {self._build_ls_legend(ls_segments, ls_total)}
                </div>
            </div>
        </div>
    </div>
    <div id="s-milkquality" class="report-layout page-break-before page-break-after">
        <div style="width: 100%;">
            <div class="section-title">乳質</div>
            <div class="subheader">体細胞 {html.escape(str(scc_threshold))}（千）以上の個体リスト</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>検定時DIM</th>
                        <th>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>体細胞（前月）</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num{r["prev_scc_class"]}">{html.escape(r["prev_scc"])}</td>'
                        f'</tr>'
                        for r in scc_over_rows
                    ) if scc_over_rows else '<tr><td colspan="7" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>
            <div class="donut-grid">
                <div class="donut-card">
                    <div class="chart-title">初産DIM&lt;60</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["heifer"]["lt60"][0], ls_donut_data["heifer"]["lt60"][2], ls_donut_data["heifer"]["lt60"][1])}
                        {self._build_ls_legend(ls_donut_data["heifer"]["lt60"][0], ls_donut_data["heifer"]["lt60"][1])}
                    </div>
                </div>
                <div class="donut-card">
                    <div class="chart-title">初産DIM60〜120</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["heifer"]["60_120"][0], ls_donut_data["heifer"]["60_120"][2], ls_donut_data["heifer"]["60_120"][1])}
                        {self._build_ls_legend(ls_donut_data["heifer"]["60_120"][0], ls_donut_data["heifer"]["60_120"][1])}
                    </div>
                </div>
                <div class="donut-card">
                    <div class="chart-title">初産DIM&gt;=121</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["heifer"]["gt120"][0], ls_donut_data["heifer"]["gt120"][2], ls_donut_data["heifer"]["gt120"][1])}
                        {self._build_ls_legend(ls_donut_data["heifer"]["gt120"][0], ls_donut_data["heifer"]["gt120"][1])}
                    </div>
                </div>
                <div class="donut-card">
                    <div class="chart-title">2産以上DIM&lt;60</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["parous"]["lt60"][0], ls_donut_data["parous"]["lt60"][2], ls_donut_data["parous"]["lt60"][1])}
                        {self._build_ls_legend(ls_donut_data["parous"]["lt60"][0], ls_donut_data["parous"]["lt60"][1])}
                    </div>
                </div>
                <div class="donut-card">
                    <div class="chart-title">2産以上DIM60〜120</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["parous"]["60_120"][0], ls_donut_data["parous"]["60_120"][2], ls_donut_data["parous"]["60_120"][1])}
                        {self._build_ls_legend(ls_donut_data["parous"]["60_120"][0], ls_donut_data["parous"]["60_120"][1])}
                    </div>
                </div>
                <div class="donut-card">
                    <div class="chart-title">2産以上DIM&gt;=121</div>
                    <div class="donut-wrap">
                        {self._build_ls_donut_svg(ls_donut_data["parous"]["gt120"][0], ls_donut_data["parous"]["gt120"][2], ls_donut_data["parous"]["gt120"][1])}
                        {self._build_ls_legend(ls_donut_data["parous"]["gt120"][0], ls_donut_data["parous"]["gt120"][1])}
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div id="s-firstparity" class="report-layout page-break-before page-break-after">
        <div style="width: 100%;">
            <div class="section-title">初産成績</div>
            <div class="subheader">初産で分娩後100日以内かつ乳量{html.escape(str(heifer_milk_target))}kg以下の個体リスト</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>検定時DIM</th>
                        <th>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in heifer_under_rows
                    ) if heifer_under_rows else '<tr><td colspan="9" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>
            <div class="scatter-card" style="margin-top: 16px;">
                <div class="chart-title">初産乳量</div>
                {self._build_heifer_milk_scatter_svg(heifer_scatter_points, heifer_milk_target)}
            </div>
        </div>
    </div>
    <div id="s-multiparity" class="report-layout page-break-before page-break-after">
        <div style="width: 100%;">
            <div class="section-title">２産以上成績</div>
            <div class="subheader">経産で分娩後100日以内かつ乳量{html.escape(str(parous_milk_target))}kg以下の個体リスト</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>検定時DIM</th>
                        <th>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in parous_under_rows
                    ) if parous_under_rows else '<tr><td colspan="9" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>
            <div class="scatter-card" style="margin-top: 16px;">
                <div class="chart-title">経産乳量</div>
                {self._build_heifer_milk_scatter_svg(parous_scatter_points, parous_milk_target)}
            </div>
        </div>
    </div>
    <div id="s-anomaly" class="report-layout page-break-before">
        <div style="width: 100%;">
            <div class="section-title">異常牛検知分析</div>
            <div class="subheader">前月から乳量15%減の個体（なぜ？）</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>DIM</th>
                        <th>今月<br>乳量</th>
                        <th>前月<br>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num">{html.escape(r["prev_milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in milk_drop_rows
                    ) if milk_drop_rows else '<tr><td colspan="10" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>

            <div class="subheader" style="margin-top: 16px;">BHB0.13以上の個体（ケトーシス）</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>DIM</th>
                        <th>今月<br>乳量</th>
                        <th>前月<br>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num">{html.escape(r["prev_milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in bhb_over_rows
                    ) if bhb_over_rows else '<tr><td colspan="10" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>

            <div class="subheader" style="margin-top: 16px;">分娩後60日以内デノボ22%未満（食べてる？）</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>DIM</th>
                        <th>今月<br>乳量</th>
                        <th>前月<br>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num">{html.escape(r["prev_milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in denovo_under_rows
                    ) if denovo_under_rows else '<tr><td colspan="10" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>

            <div class="subheader" style="margin-top: 16px;">分娩後61日以上デノボ28%未満（食べてる？）</div>
            <table class="summary-table" style="width: 100%;">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>JPN10</th>
                        <th>産次</th>
                        <th>DIM</th>
                        <th>今月<br>乳量</th>
                        <th>前月<br>乳量</th>
                        <th>体細胞（SCC）</th>
                        <th>乳脂率</th>
                        <th>BHB</th>
                        <th>デノボFA</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(
                        f'<tr class="data-row">'
                        f'<td class="num{r.get("id_cell_class", "")}"><a href="#" class="report-cow-link" data-cow-id="{html.escape(r["cow_id"])}">{html.escape(r["cow_id"])}</a></td>'
                        f'<td class="num">{html.escape(r["jpn10"])}</td>'
                        f'<td class="num">{html.escape(r["lact"])}</td>'
                        f'<td class="num">{html.escape(r["dim"])}</td>'
                        f'<td class="num">{html.escape(r["milk"])}</td>'
                        f'<td class="num">{html.escape(r["prev_milk"])}</td>'
                        f'<td class="num{r["scc_class"]}">{html.escape(r["scc"])}</td>'
                        f'<td class="num">{html.escape(r["fat"])}</td>'
                        f'<td class="num{r["bhb_class"]}">{html.escape(r["bhb"])}</td>'
                        f'<td class="num">{html.escape(r["denovo"])}</td>'
                        f'</tr>'
                        for r in denovo_under_61_rows
                    ) if denovo_under_61_rows else '<tr><td colspan="10" class="subnote">該当なし</td></tr>'}
                </tbody>
            </table>
        </div>
    </div>
    {monthly_trend_html}
    <div id="s-cowlist" class="report-section-anchor" style="height:1px;margin:0;padding:0;border:0;overflow:hidden" aria-hidden="true"></div>
    {ai_section_html}
</div>
{milk_report_comment_modal_html}
    <script>
    (function() {{
      if (typeof FALCON_OPEN_COW_PORT === 'undefined') return;
      function attachReportCowLinks() {{
        var links = document.querySelectorAll('.summary-table .report-cow-link');
        links.forEach(function(link) {{
          link.addEventListener('click', function(ev) {{
            ev.preventDefault();
            var cowId = link.getAttribute('data-cow-id');
            if (!cowId) return;
            var url = 'http://127.0.0.1:' + FALCON_OPEN_COW_PORT + '/open_cow?cow_id=' + encodeURIComponent(cowId);
            fetch(url).catch(function() {{}});
          }});
        }});
      }}
      if (document.readyState === 'loading') {{
        document.addEventListener('DOMContentLoaded', attachReportCowLinks);
      }} else {{
        attachReportCowLinks();
      }}
    }})();
    </script>
{milk_report_comment_script}
</div>
</body>
</html>"""

            # 一時HTMLファイルを作成
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.html', delete=False) as f:
                f.write(html_content)
                temp_file_path = f.name

            # ブラウザで開く
            webbrowser.open(f'file://{temp_file_path}')

        except Exception as e:
            logging.error(f"乳検レポート選択エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"乳検レポートの準備中にエラーが発生しました: {e}")

    def _build_parity_donut_svg(self, segments: List[Dict[str, Any]], avg_parity: Optional[float], total_parity: int) -> str:
        if total_parity <= 0:
            return '<div class="subnote">データなし</div>'

        radius = 50
        circumference = 2 * 3.141592653589793 * radius
        
        # セグメントを順番にソート（1産、2産、3産以上の順）
        parity_order = ["1産", "2産", "3産以上"]
        sorted_segments = []
        for label in parity_order:
            for seg in segments:
                if seg.get("label") == label and seg.get("count", 0) > 0:
                    sorted_segments.append(seg)
                    break
        
        svg_parts = [
            '<svg width="140" height="140" viewBox="0 0 120 120">',
            # 背景の円を描画（完全な円を形成するため）
            f'<circle cx="60" cy="60" r="{radius}" fill="none" stroke="#e9ecef" stroke-width="14"></circle>'
        ]
        
        for seg in sorted_segments:
            if seg["count"] <= 0:
                continue
            # offsetは既に負の値として保存されているので、そのまま使用
            offset_value = seg.get("offset", 0)
            svg_parts.append(
                f'<circle cx="60" cy="60" r="{radius}" fill="none" '
                f'stroke="{seg["color"]}" stroke-width="14" '
                f'stroke-dasharray="{seg["length"]} {circumference - seg["length"]}" '
                f'stroke-dashoffset="{offset_value}" '
                f'transform="rotate(-90 60 60)"></circle>'
            )
        if avg_parity is not None:
            avg_text = f"{avg_parity:.1f}"
            svg_parts.append(f'<text x="60" y="55" text-anchor="middle" class="donut-center">平均産次</text>')
            svg_parts.append(f'<text x="60" y="75" text-anchor="middle" class="donut-center">{html.escape(avg_text)}</text>')
        else:
            svg_parts.append(f'<text x="60" y="58" text-anchor="middle" class="donut-center">平均</text>')
            svg_parts.append(f'<text x="60" y="78" text-anchor="middle" class="donut-center">-</text>')
        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _build_parity_legend(self, segments: List[Dict[str, Any]], total_parity: int) -> str:
        if total_parity <= 0:
            return ""
        rows = []
        for seg in segments:
            if seg["count"] <= 0:
                continue
            label = html.escape(str(seg["label"]))
            rows.append(
                '<div class="legend-row">'
                f'<span><span class="legend-dot" style="background:{seg["color"]};"></span>{label}</span>'
                f'<span class="legend-value"><span class="legend-count">{seg["count"]}頭</span>'
                f'<span class="legend-percent">{seg["percent"]}%</span></span>'
                '</div>'
            )
        return f'<div class="legend">{"".join(rows)}</div>'

    def _build_milk_donut_svg(self, segments: List[Dict[str, Any]], avg_milk: Optional[float], total_count: int) -> str:
        if total_count <= 0:
            return '<div class="subnote">データなし</div>'

        radius = 50
        circumference = 2 * 3.141592653589793 * radius
        
        # セグメントを順番にソート（~20kg, 20kg台, 30kg台, 40kg台, ~50kgの順）
        milk_order = ["~20kg", "20kg台", "30kg台", "40kg台", "~50kg"]
        sorted_segments = []
        for label in milk_order:
            for seg in segments:
                if seg.get("label") == label and seg.get("count", 0) > 0:
                    sorted_segments.append(seg)
                    break
        
        svg_parts = [
            '<svg width="140" height="140" viewBox="0 0 120 120">',
            # 背景の円を描画（完全な円を形成するため）
            f'<circle cx="60" cy="60" r="{radius}" fill="none" stroke="#e9ecef" stroke-width="14"></circle>'
        ]
        
        for seg in sorted_segments:
            if seg["count"] <= 0:
                continue
            # offsetは既に負の値として保存されているので、そのまま使用
            offset_value = seg.get("offset", 0)
            svg_parts.append(
                f'<circle cx="60" cy="60" r="{radius}" fill="none" '
                f'stroke="{seg["color"]}" stroke-width="14" '
                f'stroke-dasharray="{seg["length"]} {circumference - seg["length"]}" '
                f'stroke-dashoffset="{offset_value}" '
                f'transform="rotate(-90 60 60)"></circle>'
            )
        
        avg_text = "-" if avg_milk is None else f"{avg_milk:.1f}"
        svg_parts.append(f'<text x="60" y="58" text-anchor="middle" class="donut-center">平均</text>')
        svg_parts.append(f'<text x="60" y="78" text-anchor="middle" class="donut-center">{html.escape(avg_text)}</text>')
        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _build_milk_legend(self, segments: List[Dict[str, Any]], total_count: int) -> str:
        if total_count <= 0:
            return ""
        rows = []
        for seg in segments:
            if seg["count"] <= 0:
                continue
            label = html.escape(str(seg["label"]))
            rows.append(
                '<div class="legend-row">'
                f'<span><span class="legend-dot" style="background:{seg["color"]};"></span>{label}</span>'
                f'<span class="legend-value"><span class="legend-count">{seg["count"]}頭</span>'
                f'<span class="legend-percent">{seg["percent"]}%</span></span>'
                '</div>'
            )
        return f'<div class="legend">{"".join(rows)}</div>'

    def _build_ls_donut_svg(self, segments: List[Dict[str, Any]], avg_ls: Optional[float], total_count: int) -> str:
        if total_count <= 0:
            return '<div class="subnote">データなし</div>'

        radius = 50
        circumference = 2 * 3.141592653589793 * radius
        svg_parts = [
            '<svg width="140" height="140" viewBox="0 0 120 120">',
            f'<circle cx="60" cy="60" r="{radius}" fill="none" stroke="#e9ecef" stroke-width="14"></circle>'
        ]
        for seg in segments:
            if seg["count"] <= 0:
                continue
            svg_parts.append(
                f'<circle cx="60" cy="60" r="{radius}" fill="none" '
                f'stroke="{seg["color"]}" stroke-width="14" '
                f'stroke-dasharray="{seg["length"]} {circumference - seg["length"]}" '
                f'stroke-dashoffset="{seg["offset"]}" '
                f'transform="rotate(-90 60 60)"></circle>'
            )
        if avg_ls is not None:
            avg_text = f"{avg_ls:.1f}"
            svg_parts.append(f'<text x="60" y="58" text-anchor="middle" class="donut-center">平均</text>')
            svg_parts.append(f'<text x="60" y="78" text-anchor="middle" class="donut-center">{html.escape(avg_text)}</text>')
        else:
            svg_parts.append(f'<text x="60" y="58" text-anchor="middle" class="donut-center">平均</text>')
            svg_parts.append(f'<text x="60" y="78" text-anchor="middle" class="donut-center">-</text>')
        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _build_ls_legend(self, segments: List[Dict[str, Any]], total_count: int) -> str:
        if total_count <= 0:
            return ""
        rows = []
        for seg in segments:
            if seg["count"] <= 0:
                continue
            label = html.escape(str(seg["label"]))
            rows.append(
                '<div class="legend-row">'
                f'<span><span class="legend-dot" style="background:{seg["color"]};"></span>{label}</span>'
                f'<span class="legend-value"><span class="legend-count">{seg["count"]}頭</span>'
                f'<span class="legend-percent">{seg["percent"]}%</span></span>'
                '</div>'
            )
        return f'<div class="legend">{"".join(rows)}</div>'

    def _build_milk_scatter_svg(self, points: List[Tuple[int, float, int]]) -> str:
        if not points:
            return '<div class="subnote">データなし</div>'

        width = 400
        height = 220
        margin_left = 36
        margin_right = 10
        margin_top = 10
        margin_bottom = 26
        plot_width = width - margin_left - margin_right
        plot_height = height - margin_top - margin_bottom
        max_dim = 400
        max_milk = 60

        svg_parts = [f'<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}">']

        # 背景
        svg_parts.append(
            f'<rect x="{margin_left}" y="{margin_top}" width="{plot_width}" height="{plot_height}" '
            f'fill="white" stroke="#ddd"></rect>'
        )

        # 縦グリッド（DIM 50日ごと）
        for x_val in range(0, max_dim + 1, 50):
            x = margin_left + (x_val / max_dim) * plot_width
            svg_parts.append(
                f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{margin_top + plot_height}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{x}" y="{margin_top + plot_height + 16}" font-size="10" '
                f'text-anchor="middle" fill="#666">{x_val}</text>'
            )

        # 横グリッド（乳量 10ごと）
        for y_val in range(0, max_milk + 1, 10):
            y = margin_top + plot_height - (y_val / max_milk) * plot_height
            svg_parts.append(
                f'<line x1="{margin_left}" y1="{y}" x2="{margin_left + plot_width}" y2="{y}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{margin_left - 6}" y="{y + 3}" font-size="10" '
                f'text-anchor="end" fill="#666">{y_val}</text>'
            )

        # 産次ごとの色定義（1産、2産、3産以上）- 濃い色に変更
        parity_colors = {
            1: "#E91E63",  # 1産: 濃いピンク
            2: "#00ACC1",  # 2産: 濃いターコイズ
            3: "#26A69A"   # 3産以上: 濃いミントグリーン
        }
        
        # データ点（産次で色分け）
        for dim, milk, parity in points:
            if dim < 0 or milk < 0:
                continue
            dim_clamped = min(dim, max_dim)
            milk_clamped = min(milk, max_milk)
            x = margin_left + (dim_clamped / max_dim) * plot_width
            y = margin_top + plot_height - (milk_clamped / max_milk) * plot_height
            color = parity_colors.get(parity, "#1f77b4")
            svg_parts.append(f'<circle cx="{x}" cy="{y}" r="3" fill="{color}" stroke="{color}" stroke-width="0.5" opacity="0.8"></circle>')

        # 軸ラベル
        svg_parts.append(
            f'<text x="12" y="{margin_top + plot_height / 2}" font-size="11" '
            f'text-anchor="middle" fill="#333" transform="rotate(-90 12 {margin_top + plot_height / 2})">乳量</text>'
        )

        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _build_ls_scatter_svg(self, points: List[Tuple[int, float, int]]) -> str:
        if not points:
            return '<div class="subnote">データなし</div>'

        width = 400
        height = 220
        margin_left = 36
        margin_right = 10
        margin_top = 10
        margin_bottom = 26
        plot_width = width - margin_left - margin_right
        plot_height = height - margin_top - margin_bottom
        max_dim = 400
        max_ls = 10

        svg_parts = [f'<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}">']

        # 背景
        svg_parts.append(
            f'<rect x="{margin_left}" y="{margin_top}" width="{plot_width}" height="{plot_height}" '
            f'fill="white" stroke="#ddd"></rect>'
        )

        # 縦グリッド（DIM 50日ごと）
        for x_val in range(0, max_dim + 1, 50):
            x = margin_left + (x_val / max_dim) * plot_width
            svg_parts.append(
                f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{margin_top + plot_height}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{x}" y="{margin_top + plot_height + 16}" font-size="10" '
                f'text-anchor="middle" fill="#666">{x_val}</text>'
            )

        # 横グリッド（リニアスコア 1ごと）
        for y_val in range(0, max_ls + 1, 1):
            y = margin_top + plot_height - (y_val / max_ls) * plot_height
            svg_parts.append(
                f'<line x1="{margin_left}" y1="{y}" x2="{margin_left + plot_width}" y2="{y}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{margin_left - 6}" y="{y + 3}" font-size="10" '
                f'text-anchor="end" fill="#666">{y_val}</text>'
            )

        # 産次ごとの色定義（1産、2産、3産以上）- 濃い色に変更
        parity_colors = {
            1: "#E91E63",  # 1産: 濃いピンク
            2: "#00ACC1",  # 2産: 濃いターコイズ
            3: "#26A69A"   # 3産以上: 濃いミントグリーン
        }
        
        # データ点（産次で色分け）
        for dim, ls, parity in points:
            if dim < 0 or ls < 0:
                continue
            dim_clamped = min(dim, max_dim)
            ls_clamped = min(ls, max_ls)
            x = margin_left + (dim_clamped / max_dim) * plot_width
            y = margin_top + plot_height - (ls_clamped / max_ls) * plot_height
            color = parity_colors.get(parity, "#1f77b4")
            svg_parts.append(f'<circle cx="{x}" cy="{y}" r="3" fill="{color}" stroke="{color}" stroke-width="0.5" opacity="0.8"></circle>')

        # 軸ラベル
        svg_parts.append(
            f'<text x="12" y="{margin_top + plot_height / 2}" font-size="11" '
            f'text-anchor="middle" fill="#333" transform="rotate(-90 12 {margin_top + plot_height / 2})">リニアスコア</text>'
        )

        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _build_heifer_milk_scatter_svg(self, points: List[Tuple[int, float]], target_milk: float) -> str:
        if not points:
            return '<div class="subnote">データなし</div>'

        width = 400
        height = 220
        margin_left = 36
        margin_right = 10
        margin_top = 10
        margin_bottom = 26
        plot_width = width - margin_left - margin_right
        plot_height = height - margin_top - margin_bottom
        max_dim = 400
        max_milk = 60

        svg_parts = [f'<svg width="{width}" height="{height}" viewBox="0 0 {width} {height}">']

        # 背景
        svg_parts.append(
            f'<rect x="{margin_left}" y="{margin_top}" width="{plot_width}" height="{plot_height}" '
            f'fill="white" stroke="#ddd"></rect>'
        )

        # 縦グリッド（DIM 50日ごと）
        for x_val in range(0, max_dim + 1, 50):
            x = margin_left + (x_val / max_dim) * plot_width
            svg_parts.append(
                f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{margin_top + plot_height}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{x}" y="{margin_top + plot_height + 16}" font-size="10" '
                f'text-anchor="middle" fill="#666">{x_val}</text>'
            )

        # 横グリッド（乳量 10ごと）
        for y_val in range(0, max_milk + 1, 10):
            y = margin_top + plot_height - (y_val / max_milk) * plot_height
            svg_parts.append(
                f'<line x1="{margin_left}" y1="{y}" x2="{margin_left + plot_width}" y2="{y}" '
                f'stroke="#eee" />'
            )
            svg_parts.append(
                f'<text x="{margin_left - 6}" y="{y + 3}" font-size="10" '
                f'text-anchor="end" fill="#666">{y_val}</text>'
            )

        # 目標ボックス（DIM100日、乳量目標）
        target_dim = 100
        x_target = margin_left + (target_dim / max_dim) * plot_width
        target_milk_clamped = min(max(target_milk, 0), max_milk)
        y_target = margin_top + plot_height - (target_milk_clamped / max_milk) * plot_height
        y_bottom = margin_top + plot_height
        svg_parts.append(
            f'<line x1="{margin_left}" y1="{y_target}" x2="{x_target}" y2="{y_target}" '
            f'stroke="#e74c3c" stroke-width="2" />'
        )
        svg_parts.append(
            f'<line x1="{x_target}" y1="{y_target}" x2="{x_target}" y2="{y_bottom}" '
            f'stroke="#e74c3c" stroke-width="2" />'
        )

        # データ点
        for dim, milk in points:
            if dim < 0 or milk < 0:
                continue
            dim_clamped = min(dim, max_dim)
            milk_clamped = min(milk, max_milk)
            x = margin_left + (dim_clamped / max_dim) * plot_width
            y = margin_top + plot_height - (milk_clamped / max_milk) * plot_height
            svg_parts.append(f'<circle cx="{x}" cy="{y}" r="2.5" fill="#1f77b4"></circle>')

        # 軸ラベル
        svg_parts.append(
            f'<text x="12" y="{margin_top + plot_height / 2}" font-size="11" '
            f'text-anchor="middle" fill="#333" transform="rotate(-90 12 {margin_top + plot_height / 2})">乳量</text>'
        )

        svg_parts.append("</svg>")
        return "".join(svg_parts)

    def _select_milk_test_event_date(self) -> Optional[str]:
        """乳検イベント日付を選択するダイアログ"""
        # 乳検イベント日付リストを取得
        milk_events = self.db.get_events_by_number(self.rule_engine.EVENT_MILK_TEST, include_deleted=False)
        if not milk_events:
            messagebox.showinfo("情報", "乳検イベントがありません。")
            return None
        unique_dates = sorted({e.get("event_date") for e in milk_events if e.get("event_date")}, reverse=True)
        if not unique_dates:
            messagebox.showinfo("情報", "乳検イベントの日付が取得できません。")
            return None

        selected_date: Optional[str] = None

        # 日付選択ダイアログ
        date_dialog = tk.Toplevel(self.root)
        date_dialog.title("乳検日を選択")
        date_dialog.geometry("320x380")

        ttk.Label(date_dialog, text="乳検日を選択してください", font=("", 10, "bold")).pack(pady=(15, 8))
        list_frame = ttk.Frame(date_dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        date_list = tk.Listbox(list_frame, height=12, yscrollcommand=scrollbar.set)
        date_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=date_list.yview)

        for d in unique_dates:
            date_list.insert(tk.END, d)
        date_list.selection_set(0)

        def on_ok():
            nonlocal selected_date
            selection = date_list.curselection()
            if not selection:
                messagebox.showwarning("警告", "乳検日を選択してください。")
                return
            selected_date = date_list.get(selection[0])
            date_dialog.destroy()

        def on_cancel():
            date_dialog.destroy()

        btn_frame = ttk.Frame(date_dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

        date_dialog.wait_window()
        return selected_date

    def _show_milk_test_report_for_date(self, selected_date: str, parent_dialog: Optional[tk.Toplevel] = None):
        """選択された乳検日のデータを表示"""
        try:
            from ui.report_table_window import ReportTableWindow

            # 乳検イベントに格納される項目のみ出力（保存順）
            milk_test_fields = [
                ("milk_yield", "乳量"),
                ("fat", "乳脂率"),
                ("protein", "乳蛋白率"),
                ("mun", "MUN"),
                ("scc", "SCC"),
                ("ls", "リニアスコア"),
                ("bhb", "BHB"),
                ("denovo_fa", "デノボFA"),
                ("preformed_fa", "プレフォームFA"),
                ("mixed_fa", "ミックスFA")
            ]
            columns = ["ID", "JPN10"] + [label for _, label in milk_test_fields]

            rows: List[List[Any]] = []
            events = self.db.get_events_by_number(self.rule_engine.EVENT_MILK_TEST, include_deleted=False)
            target_events = [e for e in events if e.get("event_date") == selected_date]
            for event in target_events:
                cow_auto_id = event.get("cow_auto_id")
                cow = self.db.get_cow_by_auto_id(cow_auto_id) if cow_auto_id else None
                row = [
                    cow.get("cow_id", "") if cow else "",
                    cow.get("jpn10", "") if cow else ""
                ]

                json_data = event.get("json_data") or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except Exception:
                        json_data = {}

                for item_key, _label in milk_test_fields:
                    val = json_data.get(item_key)
                    row.append(val if val is not None else "")
                rows.append(row)

            if not rows:
                messagebox.showinfo("情報", "乳検データがありません。")
                return

            ReportTableWindow(
                parent=self.root,
                report_title=f"乳検データ出力（{selected_date}）",
                columns=columns,
                rows=rows,
                conditions=f"乳検日：{selected_date}"
            )

            if parent_dialog:
                parent_dialog.destroy()

        except Exception as e:
            logging.error(f"乳検データ出力表示エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"乳検データ出力の表示に失敗しました: {e}")
        
    
    def _on_event_dictionary(self, parent_dialog):
        """イベント辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        from ui.event_dictionary_window import EventDictionaryWindow
        
        event_dict_window = EventDictionaryWindow(
            parent=self.root,
            event_dictionary_path=self.event_dict_path
        )
        event_dict_window.show()
    
    def _on_command_dictionary(self, parent_dialog):
        """コマンド辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        if not self.command_dict_path:
            messagebox.showerror("エラー", "command_dictionary.json のパスが設定されていません")
            return
        from ui.command_dictionary_window import CommandDictionaryWindow
        CommandDictionaryWindow(
            parent=self.root,
            command_dictionary_path=self.command_dict_path,
            on_updated=self._load_command_dictionary,
        ).show()

    def _on_item_dictionary(self, parent_dialog):
        """項目辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        if not self.item_dict_path:
            messagebox.showerror("エラー", "item_dictionary.json のパスが設定されていません")
            return
        
        from ui.item_dictionary_window import ItemDictionaryWindow
        
        item_window = ItemDictionaryWindow(
            parent=self.root,
            item_dictionary_path=self.item_dict_path,
            on_item_updated=self._on_item_dictionary_changed,
            formula_engine=self.formula_engine,
        )
        item_window.show()

    def _on_sire_list(self, parent_dialog):
        """SIRE一覧ボタンをクリック（既に開いていればそのウィンドウを前面に表示）"""
        parent_dialog.destroy()
        if not self.farm_path:
            messagebox.showerror("エラー", "農場が選択されていません")
            return
        if self._sire_list_window is not None and self._sire_list_window.window.winfo_exists():
            self._sire_list_window.window.lift()
            self._sire_list_window.window.focus_force()
            return
        from ui.sire_list_window import SireListWindow
        sire_list_window = SireListWindow(
            parent=self.root,
            db_handler=self.db,
            farm_path=self.farm_path,
            on_saved=self._on_sire_list_saved,
        )
        self._sire_list_window = sire_list_window
        def _on_destroy(e):
            if self._sire_list_window is not None and getattr(self._sire_list_window, "window", None) is e.widget:
                self._sire_list_window = None
        sire_list_window.window.bind("<Destroy>", _on_destroy)
        sire_list_window.show()

    def _on_sire_list_saved(self) -> None:
        """SIRE一覧の設定保存後：胎子乳用メスフラグ等はsire_listに依存するため、表示中の個体カードの計算項目を再描画する"""
        if self.current_view_type == "cow_card" and self.current_cow_card and self.current_cow_auto_id:
            try:
                self.current_cow_card.refresh_calculated_items()
            except Exception as e:
                logging.warning(f"SIRE保存後 個体カード再描画でエラー: {e}")

    def _load_item_dictionary(self):
        """item_dictionary.json を読み込む"""
        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
                logging.info(f"item_dictionary.json を読み込みました: {len(self.item_dictionary)} 項目")
            except Exception as e:
                logging.error(f"item_dictionary.json 読み込みエラー: {e}")
                self.item_dictionary = {}
        else:
            self.item_dictionary = {}

    def _load_command_dictionary(self):
        """command_dictionary.json を読み込む（短縮名 → 展開コマンド）"""
        if self.command_dict_path and self.command_dict_path.exists():
            try:
                with open(self.command_dict_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.command_dictionary = {str(k): str(v) for k, v in (data or {}).items()}
                logging.info(f"command_dictionary.json を読み込みました: {len(self.command_dictionary)} 件")
            except Exception as e:
                logging.error(f"command_dictionary.json 読み込みエラー: {e}")
                self.command_dictionary = {}
        else:
            self.command_dictionary = {}
    
    def _on_item_dictionary_changed(self):
        """項目辞書更新後のハンドラ（FormulaEngineを再読込し、表示を更新）"""
        try:
            # item_dictionaryを再読み込み
            self._load_item_dictionary()
            
            if self.formula_engine:
                self.formula_engine.reload_item_dictionary()
        except Exception as e:
            logging.error(f"item_dictionary reload failed: {e}")
        
        # CowCardを再描画
        if self.current_view_type == 'cow_card' and self.current_cow_card and self.current_cow_auto_id:
            try:
                self.current_cow_card.load_cow(self.current_cow_auto_id)
                self._add_chat_message("システム", "項目辞書の更新を反映しました。")
            except Exception as e:
                logging.error(f"CowCard refresh failed after item_dictionary update: {e}")
    
    def _on_farm_settings(self, parent_dialog):
        """農場設定ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        from ui.farm_settings_window import FarmSettingsWindow
        
        farm_settings_window = FarmSettingsWindow(
            parent=self.root,
            farm_path=self.farm_path
        )
        farm_settings_window.show()

    def _on_backup_settings(self, parent_dialog):
        """バックアップ設定ボタンをクリック"""
        parent_dialog.destroy()
        from ui.backup_settings_window import BackupSettingsWindow
        win = BackupSettingsWindow(parent=self.root, farm_path=self.farm_path)
        win.show()

    def _on_farm_management(self):
        """農場管理メニューをクリック"""
        from ui.farm_management_window import FarmManagementWindow
        
        def on_farm_changed(farm_path: Path):
            """農場が変更された時のコールバック"""
            # 農場を自動的に切り替え
            self._switch_farm(farm_path)
        
        # 農場管理ウィンドウを表示
        farm_management_window = FarmManagementWindow(
            parent=self.root,
            current_farm_path=self.farm_path,
            on_farm_changed=on_farm_changed
        )
        farm_management_window.show()
    
    def _switch_farm(self, new_farm_path: Path):
        """
        農場を切り替える（自動再初期化）
        
        Args:
            new_farm_path: 新しい農場フォルダのパス
        """
        try:
            from db.db_handler import DBHandler
            from modules.formula_engine import FormulaEngine
            from modules.rule_engine import RuleEngine
            from settings_manager import SettingsManager
            
            logging.info(f"農場を切り替え中: {self.farm_path.name} -> {new_farm_path.name}")
            
            # 現在のDB接続を閉じる
            if self.db:
                self.db.close()
            
            # 新しい農場のパスに更新
            self.farm_path = new_farm_path
            
            # 設定ファイルをロード
            settings_manager = SettingsManager(new_farm_path)
            settings_manager.load()
            
            # データベースパス
            db_path = new_farm_path / "farm.db"
            
            # 新しい農場のDBに接続（ここで self.db を差し替える）
            self.db = DBHandler(db_path)
            
            # 項目辞書・イベント辞書は本体のみのためパスは変更しない
            # コマンド辞書は農場ごとのため、切り替え先農場のパスに更新
            app_root = Path(__file__).parent.parent.parent
            config_default = app_root / "config_default"
            farm_command_dict = new_farm_path / "command_dictionary.json"
            default_command_dict = config_default / "command_dictionary.json"
            if farm_command_dict.exists():
                self.command_dict_path = farm_command_dict
            elif default_command_dict.exists():
                self.command_dict_path = default_command_dict
            else:
                self.command_dict_path = None
            self._load_command_dictionary()
            
            # FormulaEngineを再初期化（項目辞書は本体のまま）
            self.formula_engine = FormulaEngine(self.db, self.item_dict_path)
            
            # RuleEngineを再初期化
            self.rule_engine = RuleEngine(self.db)
            
            # コマンド履歴のパスを更新
            self.command_history_path = new_farm_path / "command_history.json"
            self.command_execution_history_path = new_farm_path / "command_execution_history.json"
            self._load_command_history()
            self._load_command_execution_history()
            
            # 項目辞書を再読み込み
            self._load_item_dictionary()
            
            # 現在の表示をクリア
            self._clear_result_display()
            
            # 現在の牛情報をクリア
            self.current_cow_auto_id = None
            self.current_cow_card = None
            self.current_view = None
            self.current_view_type = None
            
            # QueryRouter/QueryRendererを再初期化
            normalization_dir = Path(__file__).parent.parent.parent / "normalization"
            from modules.query_router import QueryRouter
            
            self.query_router = QueryRouter(
                normalization_dir=normalization_dir,
                item_dictionary_path=self.item_dict_path,
                event_dictionary_path=self.event_dict_path
            )
            if QUERY_V2_AVAILABLE:
                from modules.query_renderer import QueryRenderer
                self.query_renderer = QueryRenderer(
                    item_dictionary_path=str(self.item_dict_path) if self.item_dict_path else None,
                    event_dictionary_path=str(self.event_dict_path) if self.event_dict_path else None
                )
            else:
                self.query_renderer = None
            
            # ウィンドウタイトルとメイン画面の農場名ラベルを更新
            farm_name = settings_manager.get("farm_name", new_farm_path.name)
            self.root.title(f"FALCON2：{farm_name}")
            if getattr(self, "_main_farm_name_label", None) is not None:
                self._main_farm_name_label.config(text=farm_name)
            
            # 切り替え先農場の最終日付（最終分娩・最終AI・最終乳検・最終イベント）を再計算して表示を更新
            self._calculate_and_update_farm_latest_dates()
            
            logging.info(f"農場切り替え完了: {new_farm_path.name}")
            
            from tkinter import messagebox
            messagebox.showinfo(
                "農場切り替え完了",
                f"農場を「{new_farm_path.name}」に切り替えました。"
            )
            # 切り替え先農場でバックアップ該当なら実行（少し遅延してから）
            self.root.after(500, self._check_and_run_backup)
            
        except Exception as e:
            import traceback
            logging.error(f"農場切り替え中にエラーが発生しました: {e}")
            traceback.print_exc()
            from tkinter import messagebox
            messagebox.showerror(
                "エラー",
                f"農場切り替え中にエラーが発生しました:\n{e}"
            )
    
    def _on_app_settings(self):
        """アプリ設定メニューをクリック"""
        from ui.app_settings_window import AppSettingsWindow
        
        app_settings_window = AppSettingsWindow(
            parent=self.root,
            on_settings_changed=self._on_app_settings_changed
        )
        app_settings_window.show()
    
    def _on_app_settings_changed(self):
        """アプリ設定変更時のコールバック"""
        # 設定が変更された場合の処理（必要に応じて実装）
        messagebox.showinfo(
            "設定変更",
            "フォント設定を変更しました。\n完全に反映するにはアプリを再起動してください。"
        )
    
    def _on_event_saved(self, cow_auto_id: int):
        """
        イベント保存後に現在のViewを更新するコールバック
        
        Args:
            cow_auto_id: イベントが保存された牛の auto_id
        """
        # 現在CowCardが表示されている場合は更新
        if self.current_view_type == 'cow_card' and self.current_cow_card:
            if self.current_cow_auto_id == cow_auto_id:
                self.current_cow_card.load_cow(cow_auto_id)
                self._add_chat_message("システム", f"個体カードを更新しました: {cow_auto_id}")
        else:
            # 他のViewが表示されている場合はチャットにメッセージを追加
            self._add_chat_message("システム", f"イベントを保存しました: 個体ID {cow_auto_id}")
        
        # 群全体の最終日付を再計算して表示を更新
        self._calculate_and_update_farm_latest_dates()
    
    def _calculate_and_update_farm_latest_dates(self):
        """
        群全体（農場全体）の最終日付を計算して表示を更新
        """
        # latest_dates_labelが初期化されていない場合はスキップ
        if not hasattr(self, 'latest_dates_label') or self.latest_dates_label is None:
            return
        
        try:
            # 農場全体のイベントを取得
            events = self.db.get_all_events(include_deleted=False)
            
            logging.debug(f"[最終日付計算] 取得したイベント数: {len(events)}")
            
            # 最終分娩日を計算（EVENT_CALV = 202）
            latest_calving = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_CALV])
            logging.debug(f"[最終日付計算] 最終分娩日: {latest_calving}")
            
            # 最終AI/ET日を計算（AI/ET系イベント: 200, 201）
            # ※農場にAI(200)/ET(201)イベントが1件もない場合、latest_aiはNoneとなり「最終AI/ET：—」と表示される
            latest_ai = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET])
            logging.debug(f"[最終日付計算] 最終AI/ET日: {latest_ai}")
            
            # 最終乳検日を計算（EVENT_MILK_TEST = 601）
            latest_milk = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_MILK_TEST])
            logging.debug(f"[最終日付計算] 最終乳検日: {latest_milk}")
            
            # 最終イベント日を計算（全イベント対象）
            latest_any = CowCard.get_latest_any_event_date(events)
            logging.debug(f"[最終日付計算] 最終イベント日: {latest_any}")
            
            # 日付を表示用文字列に変換（Noneの場合は"—"）
            calving_str = latest_calving if latest_calving else "—"
            ai_str = latest_ai if latest_ai else "—"
            milk_str = latest_milk if latest_milk else "—"
            any_str = latest_any if latest_any else "—"
            
            # 表示文字列を作成
            display_text = f"最終分娩：{calving_str}　最終AI/ET：{ai_str}　最終乳検：{milk_str}　最終イベント：{any_str}"
            
            # Labelを更新
            self.latest_dates_label.config(text=display_text)
            logging.debug(f"[最終日付計算] 表示文字列を更新: {display_text}")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _calculate_and_update_farm_latest_dates で例外が発生しました: {e}")
            traceback.print_exc()
            # エラー時は"—"を表示
            display_text = "最終分娩：—　最終AI/ET：—　最終乳検：—　最終イベント：—"
            self.latest_dates_label.config(text=display_text)
    
    def _on_command_enter(self, event):
        """コマンド入力欄でEnterキー押下時（新しい仕様：Enterで実行）"""
        # Shift+Enterの場合は改行、通常のEnterは実行
        if event.state & 0x1:  # Shiftキーが押されている
            return  # 改行を許可
        self._on_command_execute()
        return "break"  # デフォルトの改行を防止
    
    def _on_command_key_press_entry(self, event):
        """コマンド入力欄（Entry）のキー入力イベント"""
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
    
    def _on_command_focus_in_entry(self, event):
        """コマンド入力欄（Entry）にフォーカスが入った時"""
        self.command_has_focus = True
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
        # 履歴インデックスをリセット（最新位置）
        self.command_history_index = -1
    
    def _on_command_history_up(self, event):
        """上矢印キーでコマンド履歴をさかのぼる"""
        if not self.command_execution_history:
            return "break"
        
        # 履歴インデックスを更新（-1は最新、0は最古）
        if self.command_history_index == -1:
            # 現在の入力を一時保存（履歴に追加しない）
            current_text = self.command_entry.get().strip()
            if current_text and current_text not in self.command_execution_history:
                # 現在の入力が履歴にない場合は、一時的に保存
                self.command_entry_temp_text = current_text
            else:
                self.command_entry_temp_text = None
            # 最新の履歴を表示
            self.command_history_index = len(self.command_execution_history) - 1
        else:
            # 前の履歴を表示
            if self.command_history_index > 0:
                self.command_history_index -= 1
        
        # 履歴からコマンドを取得して表示
        if 0 <= self.command_history_index < len(self.command_execution_history):
            self.command_entry.set(self.command_execution_history[self.command_history_index])
            self.command_placeholder_shown = False
        
        return "break"
    
    def _on_command_history_down(self, event):
        """下矢印キーでコマンド履歴を進む"""
        if not self.command_execution_history:
            return "break"
        
        # 履歴インデックスを更新
        if self.command_history_index == -1:
            return "break"  # 既に最新位置
        
        if self.command_history_index < len(self.command_execution_history) - 1:
            # 次の履歴を表示
            self.command_history_index += 1
            self.command_entry.set(self.command_execution_history[self.command_history_index])
            self.command_placeholder_shown = False
        else:
            # 最新位置に戻る
            self.command_history_index = -1
            # 一時保存したテキストがあれば復元
            if hasattr(self, 'command_entry_temp_text') and self.command_entry_temp_text:
                self.command_entry.set(self.command_entry_temp_text)
                self.command_placeholder_shown = False
            else:
                # タイプに応じてプレースホルダーを表示
                selected_type = self.command_type_var.get() if hasattr(self, 'command_type_var') else ""
                if selected_type == "リスト":
                    self.command_entry.set("項目１ 項目２ ・・・")
                    self.command_entry.config(foreground="gray")
                    self.command_placeholder_shown = True
                elif selected_type == "集計（経産牛）":
                    self.command_entry.set("集計したい項目を選んでください")
                    self.command_entry.config(foreground="gray")
                    self.command_placeholder_shown = True
                elif selected_type == "受胎率（経産）":
                    self.command_entry.set("")
                    self.command_entry.config(foreground="black")
                    self.command_placeholder_shown = False
                elif selected_type == "イベントリスト":
                    self.command_entry.set("all")
                    self.command_entry.config(foreground="gray")
                    self.command_placeholder_shown = True
                elif selected_type == "イベント集計":
                    self.command_entry.set("all")
                    self.command_entry.config(foreground="black")
                    self.command_placeholder_shown = False
                else:
                    self.command_entry.set("")
        
        return "break"
    
    def _update_command_history_dropdown(self):
        """コマンド入力欄のプルダウンに履歴を設定"""
        if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
            # 履歴を逆順で表示（最新が上）
            history_values = list(reversed(self.command_execution_history))
            self.command_entry['values'] = history_values
    
    def _on_command_history_selected(self, event=None):
        """プルダウンから履歴が選択された時の処理"""
        selected_value = self.command_entry.get()
        if selected_value:
            self.command_entry.set(selected_value)
            self.command_placeholder_shown = False
    
    def _on_command_dropdown_open(self, event):
        """下矢印キーでプルダウンを開く（Comboboxのデフォルト動作を利用）"""
        # Comboboxのデフォルト動作でプルダウンが開くため、特別な処理は不要
        pass
    
    def _on_command_key_press(self, event):
        """
        コマンド入力欄のキー入力イベント
        """
        if not hasattr(self, 'command_text'):
            return
        
        # プレースホルダーを非表示
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
    
    def _on_command_focus_in(self, event):
        """
        コマンド入力欄にフォーカスが入った時
        """
        self.command_has_focus = True
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
    
    def _on_command_focus_out(self, event):
        """
        コマンド入力欄からフォーカスが外れた時
        """
        if not hasattr(self, 'command_text'):
            return
        
        self.command_has_focus = False
        current_text = self.command_text.get("1.0", tk.END).strip()
        if current_text == "":
            self._show_command_placeholder()
    
    def _show_command_placeholder(self):
        """
        コマンド入力欄にプレースホルダーを表示
        """
        if not hasattr(self, 'command_text'):
            return
        
        current_text = self.command_text.get("1.0", tk.END).strip()
        if not self.command_has_focus and current_text == "":
            self.command_text.insert("1.0", self.command_placeholder)
            self.command_text.config(foreground="gray")
            self.command_placeholder_shown = True
    
    def _hide_command_placeholder(self):
        """
        コマンド入力欄のプレースホルダーを非表示
        """
        if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
            if self.command_placeholder_shown:
                current_value = self.command_entry.get()
                # プレースホルダーの場合はクリア
                if current_value in ["項目１ 項目２ ・・・", "集計したい項目を選んでください", self.command_placeholder]:
                    self.command_entry.set("")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
        elif hasattr(self, 'command_text') and self.command_text.winfo_exists():
            if self.command_placeholder_shown:
                self.command_text.delete("1.0", tk.END)
                self.command_text.config(foreground="black")
                self.command_placeholder_shown = False
    
    def _on_condition_key_press(self, event):
        """条件入力欄のキー入力イベント"""
        if self.condition_placeholder_shown:
            self.condition_entry.delete(0, tk.END)
            self.condition_entry.config(foreground="black")
            self.condition_placeholder_shown = False
    
    def _on_condition_focus_in(self, event):
        """条件入力欄にフォーカスが入った時"""
        self.condition_has_focus = True
        if self.condition_placeholder_shown:
            self.condition_entry.delete(0, tk.END)
            self.condition_entry.config(foreground="black")
            self.condition_placeholder_shown = False
    
    def _on_condition_focus_out(self, event):
        """条件入力欄からフォーカスが外れた時"""
        self.condition_has_focus = False
        current_text = self.condition_entry.get().strip()
        if not current_text:
            self.condition_entry.insert(0, self.condition_placeholder)
            self.condition_entry.config(foreground="gray")
            self.condition_placeholder_shown = True

    def _on_breakdown_key_press(self, event):
        """内訳入力欄のキー入力イベント"""
        if self.breakdown_placeholder_shown:
            self.breakdown_entry.delete(0, tk.END)
            self.breakdown_entry.config(foreground="black")
            self.breakdown_placeholder_shown = False

    def _on_breakdown_focus_in(self, event):
        """内訳入力欄にフォーカスが入った時"""
        if self.breakdown_placeholder_shown:
            self.breakdown_entry.delete(0, tk.END)
            self.breakdown_entry.config(foreground="black")
            self.breakdown_placeholder_shown = False

    def _on_breakdown_focus_out(self, event):
        """内訳入力欄からフォーカスが外れた時"""
        current_text = self.breakdown_entry.get().strip()
        if not current_text:
            self.breakdown_entry.insert(0, self.breakdown_placeholder)
            self.breakdown_entry.config(foreground="gray")
            self.breakdown_placeholder_shown = True

    def _parse_breakdown_thresholds(self, text: str) -> Optional[List[int]]:
        """内訳入力欄の閾値を解析（例: '100,150' -> [100, 150]）"""
        text = text.strip()
        if not text:
            return []
        import re
        parts = [p for p in re.split(r'[,\s]+', text) if p]
        thresholds = []
        for part in parts:
            if not part.isdigit():
                self.add_message(role="system", text="内訳は数字をカンマ区切りで入力してください。")
                return None
            thresholds.append(int(part))
        # 重複除去して昇順
        return sorted(set(thresholds))

    def _on_graph_condition_key_press(self, event):
        """グラフ条件入力欄のキー入力イベント"""
        if self.graph_condition_placeholder_shown:
            self.graph_condition_entry.delete(0, tk.END)
            self.graph_condition_entry.config(foreground="black")
            self.graph_condition_placeholder_shown = False

    def _on_graph_condition_focus_in(self, event):
        """グラフ条件入力欄にフォーカスが入った時"""
        if self.graph_condition_placeholder_shown:
            self.graph_condition_entry.delete(0, tk.END)
            self.graph_condition_entry.config(foreground="black")
            self.graph_condition_placeholder_shown = False

    def _on_graph_condition_focus_out(self, event):
        """グラフ条件入力欄からフォーカスが外れた時"""
        current_text = self.graph_condition_entry.get().strip()
        if not current_text:
            self.graph_condition_entry.insert(0, self.graph_condition_placeholder)
            self.graph_condition_entry.config(foreground="gray")
            self.graph_condition_placeholder_shown = True
    
    def _on_graph_bin_focus_in(self, event):
        """生存曲線「閾値」入力欄にフォーカスが入った時"""
        if getattr(self, 'graph_classification_bin_placeholder_shown', True):
            if hasattr(self, 'graph_classification_bin_entry'):
                self.graph_classification_bin_entry.delete(0, tk.END)
                self.graph_classification_bin_entry.config(foreground="black")
            self.graph_classification_bin_placeholder_shown = False
    
    def _on_graph_bin_focus_out(self, event):
        """生存曲線「閾値」入力欄からフォーカスが外れた時"""
        if not hasattr(self, 'graph_classification_bin_entry'):
            return
        current = self.graph_classification_bin_entry.get().strip()
        if not current:
            self.graph_classification_bin_entry.insert(0, getattr(self, 'graph_classification_bin_placeholder', '例) 0.1'))
            self.graph_classification_bin_entry.config(foreground="gray")
            self.graph_classification_bin_placeholder_shown = True
    
    def _on_classification_key_press(self, event):
        """分類入力欄のキー入力イベント（プルダウンのため使用しない）"""
        pass
    
    def _on_classification_focus_in(self, event):
        """分類入力欄にフォーカスが入った時（プルダウンのため使用しない）"""
        pass
    
    def _on_classification_focus_out(self, event):
        """分類入力欄からフォーカスが外れた時（プルダウンのため使用しない）"""
        pass
    
    def _on_conception_rate_condition_key_press(self, event):
        """受胎率条件入力欄のキー入力イベント"""
        if self.conception_rate_condition_placeholder_shown:
            self.conception_rate_condition_entry.delete(0, tk.END)
            self.conception_rate_condition_entry.config(foreground="black")
            self.conception_rate_condition_placeholder_shown = False
    
    def _on_conception_rate_condition_focus_in(self, event):
        """受胎率条件入力欄にフォーカスが入った時"""
        if self.conception_rate_condition_placeholder_shown:
            self.conception_rate_condition_entry.delete(0, tk.END)
            self.conception_rate_condition_entry.config(foreground="black")
            self.conception_rate_condition_placeholder_shown = False
    
    def _on_conception_rate_condition_focus_out(self, event):
        """受胎率条件入力欄からフォーカスが外れた時"""
        current_text = self.conception_rate_condition_entry.get().strip()
        if not current_text:
            self.conception_rate_condition_entry.insert(0, self.conception_rate_condition_placeholder)
            self.conception_rate_condition_entry.config(foreground="gray")
            self.conception_rate_condition_placeholder_shown = True
    
    def _toggle_event_selection_popup(self):
        """イベント選択ポップアップを開閉"""
        if self.event_selection_popup is None or not self.event_selection_popup.winfo_exists():
            # ポップアップを開く
            self._show_event_selection_popup()
        else:
            # ポップアップを閉じる
            self._close_event_selection_popup()
    
    def _show_event_selection_popup(self):
        """イベント選択ポップアップを表示"""
        # 既存のポップアップがあれば閉じる
        if self.event_selection_popup is not None and self.event_selection_popup.winfo_exists():
            self._close_event_selection_popup()
        
        # ポップアップウィンドウを作成
        self.event_selection_popup = tk.Toplevel(self.root)
        self.event_selection_popup.title("イベント選択")
        self.event_selection_popup.geometry("400x500")
        
        # ポップアップが閉じられた時に処理
        self.event_selection_popup.protocol("WM_DELETE_WINDOW", self._close_event_selection_popup)
        
        # スクロール可能なフレームを作成
        canvas = tk.Canvas(self.event_selection_popup)
        scrollbar = ttk.Scrollbar(self.event_selection_popup, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # イベント辞書からイベント一覧を取得
        event_dict_path = Path(__file__).parent.parent.parent / "config_default" / "event_dictionary.json"
        events_list = []
        if event_dict_path.exists():
            with open(event_dict_path, 'r', encoding='utf-8') as f:
                event_dict_data = json.load(f)
                for event_num_str, event_data in event_dict_data.items():
                    # deprecatedでないイベントのみ追加
                    if not event_data.get('deprecated', False):
                        event_number = int(event_num_str)
                        event_name = event_data.get('name_jp', event_data.get('alias', str(event_number)))
                        events_list.append((event_number, event_name))
        
        # イベント番号でソート
        events_list.sort(key=lambda x: x[0])
        
        # チェックボックスを作成
        self.event_selection_checkboxes = {}
        for event_number, event_name in events_list:
            var = tk.BooleanVar()
            # 既に選択されている場合はチェックを入れる
            if event_number in self.selected_event_numbers:
                var.set(True)
            
            checkbox = ttk.Checkbutton(
                scrollable_frame,
                text=f"{event_number}: {event_name}",
                variable=var,
                command=lambda en=event_number, v=var: self._on_event_checkbox_changed(en, v)
            )
            checkbox.pack(anchor=tk.W, padx=10, pady=2)
            self.event_selection_checkboxes[event_number] = var
        
        # 全選択/全解除ボタン
        button_frame = ttk.Frame(self.event_selection_popup)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        select_all_btn = ttk.Button(
            button_frame,
            text="全選択",
            command=self._select_all_events
        )
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        deselect_all_btn = ttk.Button(
            button_frame,
            text="全解除",
            command=self._deselect_all_events
        )
        deselect_all_btn.pack(side=tk.LEFT, padx=5)
        
        close_btn = ttk.Button(
            button_frame,
            text="閉じる",
            command=self._close_event_selection_popup
        )
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        # キャンバスとスクロールバーを配置
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def _close_event_selection_popup(self):
        """イベント選択ポップアップを閉じる"""
        if self.event_selection_popup is not None and self.event_selection_popup.winfo_exists():
            self.event_selection_popup.destroy()
        self.event_selection_popup = None
        self._update_event_selection_count()
    
    def _on_event_checkbox_changed(self, event_number, var):
        """イベントチェックボックスの状態が変更された時"""
        if var.get():
            self.selected_event_numbers.add(event_number)
        else:
            self.selected_event_numbers.discard(event_number)
        self._update_event_selection_count()
    
    def _select_all_events(self):
        """すべてのイベントを選択"""
        for event_number, var in self.event_selection_checkboxes.items():
            var.set(True)
            self.selected_event_numbers.add(event_number)
        self._update_event_selection_count()
    
    def _deselect_all_events(self):
        """すべてのイベントを解除"""
        for event_number, var in self.event_selection_checkboxes.items():
            var.set(False)
            self.selected_event_numbers.clear()
        self._update_event_selection_count()
    
    def _update_event_selection_count(self):
        """選択されたイベント数を更新"""
        count = len(self.selected_event_numbers)
        if count == 0:
            self.event_selection_count_label.config(text="(未選択)", foreground="gray")
        else:
            self.event_selection_count_label.config(text=f"({count}件選択)", foreground="black")
    
    def _on_date_entry_focus_in(self, entry, placeholder):
        """日付入力欄にフォーカスが入った時"""
        current_text = entry.get().strip()
        if current_text == placeholder:
            entry.delete(0, tk.END)
            entry.config(foreground="black")
    
    def _on_date_entry_focus_out(self, entry, placeholder):
        """日付入力欄からフォーカスが外れた時"""
        current_text = entry.get().strip()
        if not current_text:
            entry.insert(0, placeholder)
            entry.config(foreground="gray")
        else:
            # 日付形式を検証（YYYY-MM-DD）
            try:
                datetime.strptime(current_text, '%Y-%m-%d')
                entry.config(foreground="black")
            except ValueError:
                # 無効な日付形式の場合は赤色で表示
                entry.config(foreground="red")
    
    def _reset_period_entries(self):
        """期間設定（開始日・終了日）をプレースホルダーにリセットする"""
        if not hasattr(self, 'start_date_entry') or not self.start_date_entry.winfo_exists():
            return
        if not hasattr(self, 'end_date_entry') or not self.end_date_entry.winfo_exists():
            return
        ph = "YYYY-MM-DD"
        for entry, attr in [
            (self.start_date_entry, "start_date_placeholder_shown"),
            (self.end_date_entry, "end_date_placeholder_shown"),
        ]:
            entry.delete(0, tk.END)
            entry.insert(0, ph)
            entry.config(foreground="gray")
            setattr(self, attr, True)
    
    def _show_calendar_dialog(self, target_entry):
        """カレンダーダイアログを表示（統一コンポーネントを使用）"""
        from ui.date_picker_window import DatePickerWindow
        
        # 現在の入力値を初期値として使用
        initial_date = None
        if target_entry and target_entry.get():
            try:
                datetime.strptime(target_entry.get(), '%Y-%m-%d')
                initial_date = target_entry.get()
            except:
                pass
        
        def on_date_selected(date_str: str):
            """日付選択時のコールバック"""
            if target_entry:
                target_entry.delete(0, tk.END)
                target_entry.insert(0, date_str)
                target_entry.config(foreground="black")
        
        date_picker = DatePickerWindow(
            parent=self.root,
            initial_date=initial_date,
            on_date_selected=on_date_selected
        )
        date_picker.show()
    
    def _on_condition_enter(self, event):
        """条件入力欄でEnterキー押下時（コマンド実行）"""
        self._on_command_execute()
        return "break"  # デフォルトの動作を防止
    
    def _is_known_command(self, text: str) -> bool:
        """
        既存コマンドに一致するか判定
        
        Args:
            text: 入力文字列
        
        Returns:
            既存コマンドに一致する場合はTrue
        """
        known_commands = [
            "個体カード",
            "繁殖検診",
            "イベント入力",
            "データ吸い込み",
            "データ出力",
            "辞書・設定",
            "農場管理",
            "アプリ設定"
        ]
        return text in known_commands
    
    def _execute_known_command(self, text: str):
        """
        既存コマンドを実行
        
        Args:
            text: コマンド文字列
        """
        if text == "個体カード":
            self._on_cow_card()
        elif text == "繁殖検診":
            self._on_reproduction_checkup()
        elif text == "イベント入力":
            self._on_event_input()
        elif text == "データ吸い込み":
            self._on_milk_test_import()
        elif text == "データ出力":
            self._on_data_output()
        elif text == "辞書・設定":
            self._on_dictionary_settings()
        elif text == "農場管理":
            self._on_farm_management()
        elif text == "アプリ設定":
            self._on_app_settings()
    
    def _is_db_aggregation_query(self, text: str) -> bool:
        """
        DB集計系クエリかどうかを判定（項目名＋集計語）
        
        Args:
            text: ユーザー入力
        
        Returns:
            DB集計系クエリの場合はTrue
        """
        # 集計語のキーワード
        aggregation_keywords = ["平均", "割合", "率", "分布", "表", "グラフ", "散布図", "集計", "統計", "合計", "最大", "最小", "何日", "何頭", "%", "％"]
        
        # 項目辞書にマッチする項目が含まれているかチェック
        matched_items = self._match_items_from_query(text)
        
        # 集計語と項目名の両方が含まれている場合
        has_aggregation = any(keyword in text for keyword in aggregation_keywords)
        
        return has_aggregation and len(matched_items) > 0
    
    def _should_use_db_only(self, text: str) -> bool:
        """
        AI送信禁止ルール：必ずDB処理を実行すべきか判定
        
        Args:
            text: ユーザー入力
        
        Returns:
            DB処理のみを実行すべき場合はTrue
        """
        # イベント名のキーワード
        event_keywords = ["分娩", "AI", "ET", "妊娠鑑定"]
        
        # 期間・日付のキーワード
        date_keywords = ["月", "期間", "日付", "日に", "年", "日"]
        
        # 個体・一覧のキーワード
        list_keywords = ["個体", "一覧", "教えて"]
        
        # イベント名が含まれている
        has_event = any(keyword in text for keyword in event_keywords)
        
        # 期間/日付または個体/一覧が含まれている
        has_date_or_list = any(keyword in text for keyword in date_keywords) or \
                          any(keyword in text for keyword in list_keywords)
        
        # イベント名＋（期間/日付 または 個体/一覧）の場合は必ずDB処理
        return has_event and has_date_or_list
    
    def _normalize_query(self, query: str) -> str:
        """
        入力文字列を正規化（表記ゆれ吸収）
        
        Args:
            query: ユーザー入力
        
        Returns:
            正規化されたクエリ
        """
        # 全角・半角の統一
        normalized = query
        # 全角数字を半角に
        normalized = normalized.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
        # 全角英字を半角に
        normalized = normalized.translate(str.maketrans('ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
        normalized = normalized.translate(str.maketrans('ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ', 'abcdefghijklmnopqrstuvwxyz'))
        # 全角スペースを半角に
        normalized = normalized.replace('　', ' ')
        # 連続するスペースを1つに
        normalized = ' '.join(normalized.split())
        
        return normalized.strip()
    
    def _on_command_execute(self):
        """コマンド実行（ルールエンジン版：AI不使用）"""
        # コマンド入力欄からテキストを取得（command_entry または command_text）
        raw_input = ""
        selected_type = self.command_type_var.get() if hasattr(self, 'command_type_var') else ""
        if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
            raw_input = self.command_entry.get().strip()
            # イベントリストタイプの場合、「all」プレースホルダーはそのまま使用
            if selected_type == "イベントリスト" and self.command_placeholder_shown and raw_input == "all":
                # 「all」プレースホルダーの場合はそのまま使用（空文字列にしない）
                pass
            elif self.command_placeholder_shown:
                # その他のプレースホルダーの場合は空文字列として扱う
                raw_input = ""
        elif hasattr(self, 'command_text') and self.command_text.winfo_exists():
            raw_input = self.command_text.get("1.0", tk.END).strip()
            # プレースホルダーが表示されている場合は空文字列として扱う
            if self.command_placeholder_shown:
                raw_input = ""
        
        # コマンド辞書に登録されていれば展開（表示用に元入力も保持）
        user_input_for_display = raw_input
        if raw_input and self.command_dictionary and raw_input in self.command_dictionary:
            raw_input = self.command_dictionary[raw_input]
        
        # 集計でコマンド空欄の場合は表示用に「頭数」または「頭数：産次」を使う
        if selected_type == "集計（経産牛）":
            effective_for_display = raw_input.strip()
            if not effective_for_display or effective_for_display == "集計したい項目を選んでください":
                _cls = ""
                if hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                    _cls = self.classification_entry.get().strip()
                user_input_for_display = "頭数" + ("：" + _cls if _cls else "")
        
        # 受胎率・グラフ・集計（経産牛）の場合は、コマンド入力欄が空でも処理を続行（集計は空欄時は頭数を表示）
        if not raw_input and selected_type not in ["受胎率（経産）", "グラフ", "集計（経産牛）"]:
            return
        
        # タイプが選択されている場合、自動的にプレフィックスを付与
        selected_type = self.command_type_var.get() if hasattr(self, 'command_type_var') else ""
        actual_command = raw_input  # 実際に実行されるコマンド（表タブに表示される）
        
        # リストタイプの場合、条件入力欄の内容を統合
        if selected_type == "リスト":
            # コマンド欄が既に「リスト：」で始まっている場合は除去（二重プレフィックス防止）
            _raw = raw_input.strip()
            while _raw.startswith("リスト：") or _raw.startswith("リスト:"):
                _raw = (_raw[4:].strip() if _raw.startswith("リスト：") else _raw[4:].strip())
            raw_input = _raw or raw_input
            # 条件入力欄から条件を取得
            condition_text = ""
            if hasattr(self, 'condition_entry') and self.condition_entry.winfo_exists():
                condition_text = self.condition_entry.get().strip()
                # プレースホルダーの場合は無視
                if condition_text == self.condition_placeholder:
                    condition_text = ""
            
            # 条件がある場合、条件を解析して適切に処理
            if condition_text:
                # キーワード「受胎」「空胎」を RC 条件に展開（コマンド解析でそのまま使えるように）
                cond_stripped = condition_text.strip()
                if cond_stripped == "受胎":
                    condition_text = "RC=5,6"
                elif cond_stripped == "空胎":
                    condition_text = "RC=1,2,3,4"
                # 条件を解析（例：DIM>150, 産次：初産, RC>=5）
                # 条件を項目名と値に分離
                # 比較演算子（<, >, <=, >=, =）を含む場合と、項目名：値の形式を処理
                if "：" in condition_text or ":" in condition_text:
                    # 項目名：値の形式（例：産次：初産）
                    if "：" in condition_text:
                        cond_parts = condition_text.split("：", 1)
                    else:
                        cond_parts = condition_text.split(":", 1)
                    cond_item = cond_parts[0].strip()
                    cond_value = cond_parts[1].strip()
                    # 項目リストに条件項目を追加（まだ含まれていない場合）
                    parts = raw_input.split()
                    if cond_item not in parts:
                        parts.append(cond_item)
                    # 条件を付与
                    parts[-1] = f"{cond_item}：{cond_value}"
                    raw_input = " ".join(parts)
                else:
                    # 比較演算子を含む形式（例：DIM>150, AGE>30）
                    # 条件を解析して項目名を抽出
                    import re
                    # 比較演算子パターン: <, >, <=, >=, =, !=
                    match = re.match(r'^([A-Za-z_][A-Za-z0-9_]*)\s*([<>=!]+)\s*(.+)$', condition_text.strip())
                    if match:
                        cond_item = match.group(1)
                        # 項目リストに条件項目を追加（まだ含まれていない場合）
                        parts = raw_input.split()
                        if cond_item not in parts:
                            parts.append(cond_item)
                        # 条件を「：条件」の形式で付与（例：ID DIM RC AGE：AGE>30）
                        raw_input = " ".join(parts) + f"：{condition_text}"
                    else:
                        # 解析できない場合は、そのまま最後に「：条件」を付与
                        parts = raw_input.split()
                        if parts:
                            raw_input = " ".join(parts) + f"：{condition_text}"
                        else:
                            raw_input = f"：{condition_text}"
                _list_input = raw_input.strip()
                while _list_input.startswith("リスト：") or _list_input.startswith("リスト:"):
                    _list_input = (_list_input[4:].strip() if _list_input.startswith("リスト：") else _list_input[4:].strip())
                actual_command = "リスト：" + _list_input
            else:
                # 条件がない場合
                _list_input = raw_input.strip()
                while _list_input.startswith("リスト：") or _list_input.startswith("リスト:"):
                    _list_input = (_list_input[4:].strip() if _list_input.startswith("リスト：") else _list_input[4:].strip())
                actual_command = "リスト：" + _list_input if _list_input else raw_input
            # 並び順が指定されていればコマンドに付与（リスト：～～～～：条件：昇順/降順 項目名）
            if selected_type == "リスト" and hasattr(self, 'sort_order_combo') and hasattr(self, 'sort_column_entry'):
                sort_order = (self.sort_order_var.get() or "").strip()
                sort_column = (self.sort_column_entry.get() or "").strip()
                if sort_column and sort_order in ("昇順", "降順"):
                    actual_command = actual_command + "：" + sort_order + " " + sort_column
        elif selected_type == "イベントリスト":
            # イベントリストタイプの場合、条件入力欄は使用しない
            # コマンド文字列に条件を含めない
            if not (raw_input.startswith("イベントリスト：") or raw_input.startswith("イベントリスト:")):
                actual_command = "イベントリスト：" + raw_input
            else:
                actual_command = raw_input
        elif selected_type == "イベント集計":
            # イベント集計タイプの場合
            if not (raw_input.startswith("イベント集計：") or raw_input.startswith("イベント集計:")):
                actual_command = "イベント集計：" + raw_input
            else:
                actual_command = raw_input
            # 分類欄の値をコマンドに追加
            classification = ""
            if hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                classification = self.classification_entry.get().strip()
            if classification:
                actual_command = actual_command + "：" + classification
        elif selected_type:
            prefix_map = {
                "受胎率（経産）": "受胎率：",
                "集計（経産牛）": "集計：",
                "グラフ": "グラフ："
            }
            prefix = prefix_map.get(selected_type, "")
            if prefix:
                # 既にプレフィックスが付いている場合はそのまま、付いていない場合は付与
                if not (raw_input.startswith("リスト：") or raw_input.startswith("リスト:") or
                        raw_input.startswith("イベントリスト：") or raw_input.startswith("イベントリスト:") or
                        raw_input.startswith("受胎率：") or raw_input.startswith("受胎率:") or
                        raw_input.startswith("集計：") or raw_input.startswith("集計:") or
                        raw_input.startswith("グラフ：") or raw_input.startswith("グラフ:") or
                        raw_input.startswith("散布図：") or raw_input.startswith("散布図:")):
                    actual_command = prefix + raw_input
                else:
                    # 既にプレフィックスが付いている場合はそのまま使用
                    actual_command = raw_input
            else:
                actual_command = raw_input
            
            # 集計タイプの場合、分類欄と条件欄の値をコマンドに追加
            if selected_type == "集計（経産牛）":
                classification = ""
                if hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                    classification = self.classification_entry.get().strip()
                
                condition_text = ""
                if hasattr(self, 'condition_entry') and self.condition_entry.winfo_exists():
                    condition_text = self.condition_entry.get().strip()
                    # プレースホルダーの場合は無視
                    if condition_text == self.condition_placeholder:
                        condition_text = ""
                
                # コマンド欄が空またはプレースホルダーの場合は「頭数」として実行（経産牛頭数 or 産次別頭数）
                effective_input = raw_input.strip()
                if not effective_input or effective_input == "集計したい項目を選んでください":
                    effective_input = "頭数"
                
                if not (effective_input.startswith("集計：") or effective_input.startswith("集計:")):
                    actual_command = "集計：" + effective_input
                else:
                    actual_command = effective_input
                
                if classification:
                    actual_command = actual_command + "：" + classification
                
                if condition_text:
                    actual_command = actual_command + "：" + condition_text
        
        # デバッグ補助
        print("COMMAND:", actual_command)
        
        # コマンド実行履歴に追加（ユーザーが入力した短縮名を保存し、履歴で再実行時に再度展開される）
        if not self.command_execution_history or self.command_execution_history[-1] != user_input_for_display:
            self.command_execution_history.append(user_input_for_display)
            # 履歴は最大50件まで保持
            if len(self.command_execution_history) > 50:
                self.command_execution_history.pop(0)
            # 履歴をファイルに保存
            self._save_command_execution_history()
            # プルダウンの履歴を更新
            self._update_command_history_dropdown()
        # 履歴インデックスをリセット（最新位置）
        self.command_history_index = -1
        
        # 内訳入力の値を保持（入力欄はこの後クリアされるため）
        self._pending_breakdown_text = ""
        if selected_type == "集計（経産牛）" and hasattr(self, 'breakdown_entry') and self.breakdown_entry.winfo_exists():
            breakdown_text = self.breakdown_entry.get().strip()
            if breakdown_text == self.breakdown_placeholder:
                breakdown_text = ""
            self._pending_breakdown_text = breakdown_text
        
        # 前回の結果をクリア
        self._clear_result_display()
        
        # コマンドをチャット履歴に追加（ユーザーが入力した内容を表示・辞書展開前の入力で表示）
        self.add_message(role="user", text=user_input_for_display)
        
        # コマンドをクリア（集計・リスト・イベントリスト・イベント集計の場合は保持）
        if selected_type not in ["集計（経産牛）", "リスト", "イベントリスト", "イベント集計"]:
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.set("")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
            elif hasattr(self, 'command_text') and self.command_text.winfo_exists():
                self._hide_command_placeholder()
                self.command_text.delete("1.0", tk.END)
        
        # 条件入力欄をクリア（集計は保持）
        if selected_type == "リスト" and hasattr(self, 'condition_entry') and self.condition_entry.winfo_exists():
            self.condition_entry.delete(0, tk.END)
            self.condition_entry.insert(0, self.condition_placeholder)
            self.condition_entry.config(foreground="gray")
            self.condition_placeholder_shown = True
        # 内訳入力欄・分類入力欄は集計の場合保持
        
        # 入力文字列を正規化（表記ゆれ吸収）
        normalized_input = self._normalize_query(actual_command)
        
        # コマンド入力履歴を記録（項目名として使用可能な場合）
        self._record_command_history(actual_command, normalized_input)
        
        # 判定ロジック（優先順位順、厳守）
        # ① 受胎率コマンド（受胎率タイプが選択されている場合）
        if selected_type == "受胎率（経産）":
            logging.info(f"受胎率コマンドを実行します")
            self._execute_conception_rate_command()
            return
        
        # ① 項目一括変更コマンド（項目一括変更タイプが選択されている場合）
        if selected_type == "項目一括変更":
            logging.info(f"項目一括変更コマンドを実行します: '{raw_input}'")
            self._execute_batch_item_edit_command(raw_input)
            return
        
        # ② イベント集計コマンド（「イベント集計：」で始まる）
        if actual_command.startswith("イベント集計：") or actual_command.startswith("イベント集計:"):
            logging.info(f"イベント集計コマンドを実行します: '{actual_command}'")
            self._execute_event_aggregate_command(actual_command)
            return
        
        # ② イベントリストコマンド（「イベントリスト：」で始まる）
        logging.info(f"コマンド実行判定: actual_command='{actual_command}'")
        if actual_command.startswith("イベントリスト：") or actual_command.startswith("イベントリスト:"):
            logging.info(f"イベントリストコマンドを実行します: '{actual_command}'")
            self._execute_event_list_command(actual_command)
            return
        
        # ② リストコマンド（「リスト：」で始まる）
        if actual_command.startswith("リスト：") or actual_command.startswith("リスト:"):
            self._execute_list_command(actual_command)
            return
        
        # ③ 集計コマンド（「集計：」で始まる）
        if actual_command.startswith("集計：") or actual_command.startswith("集計:"):
            # 表タブを選択してから処理
            self.result_notebook.select(0)
            self._execute_aggregate_command(actual_command)
            return
        
        # ③ グラフコマンド（グラフタイプが選択されている場合）
        if selected_type == "グラフ":
            self._execute_graph_command()
            return
        
        # ③ グラフコマンド（「グラフ：」で始まる）
        if actual_command.startswith("グラフ：") or actual_command.startswith("グラフ:") or actual_command.startswith("散布図：") or actual_command.startswith("散布図:"):
            # グラフタブを選択してから処理
            self.result_notebook.select(1)  # 1番目のタブ（グラフタブ）
            if not self._prompt_monthly_milk_test_months_for_command(actual_command):
                self.add_message(role="system", text="月の指定がキャンセルされました。")
                return
            self._execute_query_router(actual_command)
            return
        
        # ④ 個体ID（数値のみ）- タイプが選択されていない場合のみ
        if not selected_type:
            raw_normalized = self._normalize_query(raw_input)
            if raw_normalized.isdigit():
                padded_id = raw_normalized.zfill(4)
                print("PADDED ID:", padded_id)
                self._jump_to_cow_card(padded_id)
                return
        
        # ⑤ 既存コマンドに一致する場合 - タイプが選択されていない場合のみ
        if not selected_type:
            raw_normalized = self._normalize_query(raw_input)
            if self._is_known_command(raw_normalized):
                self._execute_known_command(raw_normalized)
                return
        
        # ⑥ イベント名で検索（alias または name_jp で前方一致）- タイプが選択されていない場合のみ
        if not selected_type:
            raw_normalized = self._normalize_query(raw_input)
            event_number = self._find_event_by_name(raw_normalized)
            if event_number is not None:
                self._open_event_input_for_event(event_number)
                return
        
        # ⑦ 新しいQueryRouterシステムで処理（デフォルトは表タブ）
        # 期間指定の表示/非表示は_execute_query_router内で判定
        self.result_notebook.select(0)
        self._execute_query_router(actual_command)

    def _load_command_history(self):
        """
        コマンド履歴をファイルから読み込む
        """
        if self.command_history_path.exists():
            try:
                with open(self.command_history_path, 'r', encoding='utf-8') as f:
                    self.command_history = json.load(f)
                logging.info(f"コマンド履歴を読み込みました: {len(self.command_history)}件")
            except Exception as e:
                logging.warning(f"コマンド履歴の読み込みに失敗しました: {e}")
                self.command_history = {}
        else:
            self.command_history = {}
    
    def _save_command_history(self):
        """
        コマンド履歴をファイルに保存
        """
        try:
            # 農場フォルダが存在することを確認
            self.command_history_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.command_history_path, 'w', encoding='utf-8') as f:
                json.dump(self.command_history, f, ensure_ascii=False, indent=2)
            logging.debug(f"コマンド履歴を保存しました: {len(self.command_history)}件")
        except Exception as e:
            logging.error(f"コマンド履歴の保存に失敗しました: {e}")
    
    def _load_command_execution_history(self):
        """
        コマンド実行履歴をファイルから読み込む
        """
        if self.command_execution_history_path.exists():
            try:
                with open(self.command_execution_history_path, 'r', encoding='utf-8') as f:
                    loaded_history = json.load(f)
                    if isinstance(loaded_history, list):
                        self.command_execution_history = loaded_history
                        # 履歴は最大50件まで保持
                        if len(self.command_execution_history) > 50:
                            self.command_execution_history = self.command_execution_history[-50:]
                        logging.info(f"コマンド実行履歴を読み込みました: {len(self.command_execution_history)}件")
                    else:
                        logging.warning("コマンド実行履歴の形式が不正です。")
                        self.command_execution_history = []
            except Exception as e:
                logging.warning(f"コマンド実行履歴の読み込みに失敗しました: {e}")
                self.command_execution_history = []
        else:
            self.command_execution_history = []
    
    def _save_command_execution_history(self):
        """
        コマンド実行履歴をファイルに保存
        """
        try:
            # 農場フォルダが存在することを確認
            self.command_execution_history_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.command_execution_history_path, 'w', encoding='utf-8') as f:
                json.dump(self.command_execution_history, f, ensure_ascii=False, indent=2)
            logging.debug(f"コマンド実行履歴を保存しました: {len(self.command_execution_history)}件")
        except Exception as e:
            logging.error(f"コマンド実行履歴の保存に失敗しました: {e}")
    
    def _record_command_history(self, raw_input: str, normalized_input: str):
        """
        コマンド入力履歴を記録
        
        Args:
            raw_input: 元の入力文字列
            normalized_input: 正規化された入力文字列
        """
        history_updated = False
        # 項目名として使用可能かチェック（item_dictionaryに存在するか）
        # スペースで分割して、各単語をチェック
        words = normalized_input.split()
        for word in words:
            # item_dictionaryに存在する項目名かチェック
            if word in self.item_dictionary:
                self.command_history[word] = self.command_history.get(word, 0) + 1
                history_updated = True
            # イベント名としてもチェック（event_dictionaryから）
            elif self._find_event_by_name(word) is not None:
                self.command_history[word] = self.command_history.get(word, 0) + 1
                history_updated = True
        
        # 履歴が更新された場合は保存
        if history_updated:
            self._save_command_history()
    
    def _on_command_type_selected(self, event=None):
        """コマンドタイプが選択された時（command_entry用）"""
        selected_type = self.command_type_var.get()
        
        if selected_type != "イベントリスト":
            # イベントリスト固定の「all」設定を解除し、通常の履歴プルダウンに戻す
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self._update_command_history_dropdown()
        
        # 受胎率タイプが選択された場合
        if selected_type == "受胎率（経産）":
            # イベント選択フレームを非表示
            if hasattr(self, 'event_selection_frame'):
                self.event_selection_frame.pack_forget()
            # グラフ入力フレームを非表示
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            # 条件入力欄を非表示
            if hasattr(self, 'condition_frame'):
                self.condition_frame.pack_forget()
            
            # コマンド入力欄を無効化（記入できないようにする）
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.set("")
                self.command_entry.config(state="disabled")
                self.command_placeholder_shown = False
            
            # 受胎率フレームを表示
            if hasattr(self, 'conception_rate_frame'):
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.conception_rate_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.conception_rate_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 受胎率の種類の選択肢を明示的に再設定
            if hasattr(self, 'conception_rate_type_entry'):
                conception_rate_type_options = ["月", "産次", "授精回数", "DIMサイクル", "授精種類", "授精間隔", "授精師", "SIRE"]
                self.conception_rate_type_entry['values'] = conception_rate_type_options
            
            # 期間設定フレームを表示
            if hasattr(self, 'period_frame'):
                # conception_rate_frameの後に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # デフォルト期間を設定（30日前からさかのぼって1年間）
            if hasattr(self, 'start_date_entry') and hasattr(self, 'end_date_entry'):
                from datetime import datetime, timedelta
                today = datetime.now()
                # 終了日：30日前
                end_date_dt = today - timedelta(days=30)
                end_date = end_date_dt.strftime('%Y-%m-%d')
                # 開始日：終了日から1年前
                start_date_dt = end_date_dt - timedelta(days=365)
                start_date = start_date_dt.strftime('%Y-%m-%d')
                self.start_date_entry.delete(0, tk.END)
                self.start_date_entry.insert(0, start_date)
                self.start_date_entry.config(foreground="black")
                self.start_date_placeholder_shown = False
                self.end_date_entry.delete(0, tk.END)
                self.end_date_entry.insert(0, end_date)
                self.end_date_entry.config(foreground="black")
                self.end_date_placeholder_shown = False
        # リストタイプが選択された場合
        elif selected_type == "リスト":
            # 受胎率フレームを非表示
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            # グラフ入力フレームを非表示
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            # イベント選択フレームを非表示
            if hasattr(self, 'event_selection_frame'):
                self.event_selection_frame.pack_forget()
            # コマンド入力欄を有効化
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.config(state="normal")
            # 分類ラベルと分類入力欄を非表示、条件ラベルと条件入力欄のみ表示
            if hasattr(self, 'classification_label'):
                self.classification_label.pack_forget()
            if hasattr(self, 'classification_entry'):
                self.classification_entry.pack_forget()
            if hasattr(self, 'breakdown_label'):
                self.breakdown_label.pack_forget()
            if hasattr(self, 'breakdown_entry'):
                self.breakdown_entry.pack_forget()
            # 条件ラベルと条件入力欄を表示
            if hasattr(self, 'condition_label'):
                try:
                    self.condition_label.pack_info()
                except:
                    self.condition_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'condition_entry'):
                try:
                    self.condition_entry.pack_info()
                except:
                    self.condition_entry.pack(side=tk.LEFT, padx=5)
            # 並び順（リスト時のみ）
            if hasattr(self, 'sort_label'):
                self.sort_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'sort_column_entry'):
                self.sort_column_entry.pack(side=tk.LEFT, padx=2)
            if hasattr(self, 'sort_order_combo'):
                self.sort_order_combo.pack(side=tk.LEFT, padx=2)
            
            # 条件入力欄のフレームを表示
            if hasattr(self, 'condition_frame'):
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 期間設定フレームを非表示
            if hasattr(self, 'period_frame'):
                self.period_frame.pack_forget()
            
            # コマンドラインに「リスト：ID」を自動入力
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.set("リスト：ID")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
        elif selected_type == "イベントリスト":
            # 受胎率フレームを非表示
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            # グラフ入力フレームを非表示
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            # イベントリストタイプが選択された場合
            # イベント選択フレームを表示
            if hasattr(self, 'event_selection_frame'):
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.event_selection_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.event_selection_frame.pack(fill=tk.X, padx=10, pady=5)
            # コマンド入力欄を有効化
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.config(state="readonly")
                self.command_entry['values'] = ["all"]
                self.command_entry.set("all")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
            
            # 分類ラベルと分類入力欄を非表示
            if hasattr(self, 'classification_label'):
                self.classification_label.pack_forget()
            if hasattr(self, 'classification_entry'):
                self.classification_entry.pack_forget()
            if hasattr(self, 'breakdown_label'):
                self.breakdown_label.pack_forget()
            if hasattr(self, 'breakdown_entry'):
                self.breakdown_entry.pack_forget()
            
            # 条件入力欄のフレームを非表示
            if hasattr(self, 'condition_frame'):
                self.condition_frame.pack_forget()
            
            # 期間設定フレームを表示
            if hasattr(self, 'period_frame'):
                # event_selection_frameの後に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # イベントリストはコマンド固定（all）
        elif selected_type == "イベント集計":
            # 受胎率フレームを非表示
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            # グラフ入力フレームを非表示
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            # イベント集計タイプが選択された場合
            # イベント選択フレームを表示（部分一致は非表示）
            if hasattr(self, 'event_selection_frame'):
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.event_selection_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.event_selection_frame.pack(fill=tk.X, padx=10, pady=5)
            # 部分一致入力欄を非表示
            if hasattr(self, 'note_partial_match_label'):
                self.note_partial_match_label.pack_forget()
            if hasattr(self, 'note_partial_match_entry'):
                self.note_partial_match_entry.pack_forget()
            
            # 期間設定フレームを表示
            if hasattr(self, 'period_frame'):
                # event_selection_frameの後に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 分類入力欄を表示（イベント集計専用の分類オプションを使用）
            # イベント集計専用の分類オプションを設定
            if hasattr(self, 'classification_entry'):
                event_aggregate_classification_options = [
                    "",  # 空欄（分類なし）
                    "月", "産次", "産次（１産、２産、３産以上）", "DIM30"
                ]
                self.classification_entry['values'] = event_aggregate_classification_options
                self.classification_entry.set("")  # 初期値は空
            
            # 条件入力欄のフレームを表示
            if hasattr(self, 'condition_frame'):
                # 分類ラベルと分類入力欄を表示
                if hasattr(self, 'classification_label'):
                    self.classification_label.pack(side=tk.LEFT, padx=(10, 5))
                if hasattr(self, 'classification_entry'):
                    self.classification_entry.pack(side=tk.LEFT, padx=5)
                # 条件ラベルと条件入力欄は非表示
                if hasattr(self, 'condition_label'):
                    self.condition_label.pack_forget()
                if hasattr(self, 'condition_entry'):
                    self.condition_entry.pack_forget()
                # 内訳ラベルと内訳入力欄は非表示
                if hasattr(self, 'breakdown_label'):
                    self.breakdown_label.pack_forget()
                if hasattr(self, 'breakdown_entry'):
                    self.breakdown_entry.pack_forget()
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # コマンド入力欄を有効化（固定値なし）
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.config(state="normal")
                self.command_entry.set("all")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
        elif selected_type == "集計（経産牛）":
            # 受胎率フレームを非表示
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            # イベント選択フレームを非表示
            if hasattr(self, 'event_selection_frame'):
                self.event_selection_frame.pack_forget()
            # 期間設定フレームを非表示
            if hasattr(self, 'period_frame'):
                self.period_frame.pack_forget()
            # 集計（経産牛）タイプが選択された場合、全てのウィジェットを正しい順序で再配置
            # まず、全てのウィジェットを一時的に非表示にしてから、正しい順序で再配置
            # コマンド入力欄を有効化
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.config(state="normal")
            
            # 分類オプションを元に戻す（集計用の全オプション）
            if hasattr(self, 'classification_entry'):
                self.classification_entry['values'] = self.classification_options
                self.classification_entry.set("")  # 初期値は空
            if hasattr(self, 'classification_label'):
                try:
                    self.classification_label.pack_forget()
                except:
                    pass
            if hasattr(self, 'classification_entry'):
                try:
                    self.classification_entry.pack_forget()
                except:
                    pass
            if hasattr(self, 'condition_label'):
                try:
                    self.condition_label.pack_forget()
                except:
                    pass
            if hasattr(self, 'condition_entry'):
                try:
                    self.condition_entry.pack_forget()
                except:
                    pass
            if hasattr(self, 'breakdown_label'):
                try:
                    self.breakdown_label.pack_forget()
                except:
                    pass
            if hasattr(self, 'breakdown_entry'):
                try:
                    self.breakdown_entry.pack_forget()
                except:
                    pass
            
            # 並び順は集計時は非表示
            if hasattr(self, 'sort_label'):
                self.sort_label.pack_forget()
            if hasattr(self, 'sort_column_entry'):
                self.sort_column_entry.pack_forget()
            if hasattr(self, 'sort_order_combo'):
                self.sort_order_combo.pack_forget()
            # 正しい順序で再配置：分類ラベル → 分類入力欄 → 条件ラベル → 条件入力欄 → 内訳ラベル → 内訳入力欄
            if hasattr(self, 'classification_label'):
                self.classification_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'classification_entry'):
                self.classification_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'condition_label'):
                self.condition_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'condition_entry'):
                self.condition_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'breakdown_label'):
                self.breakdown_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'breakdown_entry'):
                self.breakdown_entry.pack(side=tk.LEFT, padx=5)
            
            # 条件入力欄のフレームを表示
            if hasattr(self, 'condition_frame'):
                # latest_dates_frameの前に配置
                if hasattr(self, 'latest_dates_frame'):
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.condition_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # コマンドラインにプレースホルダーを表示
            current_text = self.command_entry.get().strip()
            if not current_text or current_text == self.command_placeholder:
                self.command_entry.set("集計したい項目を選んでください")
                self.command_entry.config(foreground="gray")
                self.command_placeholder_shown = True
        elif selected_type == "グラフ":
            # 受胎率フレームを非表示
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            # イベント選択フレームを非表示
            if hasattr(self, 'event_selection_frame'):
                self.event_selection_frame.pack_forget()
            # 条件入力欄（集計用）を非表示
            if hasattr(self, 'condition_frame'):
                self.condition_frame.pack_forget()
            
            # コマンド入力欄を無効化
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.set("")
                self.command_entry.config(state="disabled")
                self.command_placeholder_shown = False
            
            # グラフ入力フレームを表示
            if hasattr(self, 'graph_input_frame'):
                if hasattr(self, 'latest_dates_frame'):
                    self.graph_input_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.graph_input_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # グラフ種類をリセットし、Y軸、X軸、分類、条件を非表示
            if hasattr(self, 'graph_type_entry'):
                self.graph_type_entry.set("")
            if hasattr(self, 'graph_y_label'):
                self.graph_y_label.pack_forget()
            if hasattr(self, 'graph_y_entry'):
                self.graph_y_entry.delete(0, tk.END)
                self.graph_y_entry.pack_forget()
            if hasattr(self, 'graph_y_item_list_btn'):
                self.graph_y_item_list_btn.pack_forget()
            if hasattr(self, 'graph_x_label'):
                self.graph_x_label.pack_forget()
            if hasattr(self, 'graph_x_entry'):
                self.graph_x_entry.delete(0, tk.END)
                self.graph_x_entry.pack_forget()
            if hasattr(self, 'graph_x_item_list_btn'):
                self.graph_x_item_list_btn.pack_forget()
            if hasattr(self, 'graph_classification_label'):
                self.graph_classification_label.pack_forget()
            if hasattr(self, 'graph_classification_entry'):
                self.graph_classification_entry.set("")
                self.graph_classification_entry.pack_forget()
            if hasattr(self, 'graph_condition_label'):
                self.graph_condition_label.pack_forget()
            if hasattr(self, 'graph_condition_entry'):
                self.graph_condition_entry.delete(0, tk.END)
                self.graph_condition_entry.insert(0, self.graph_condition_placeholder)
                self.graph_condition_entry.config(foreground="gray")
                self.graph_condition_placeholder_shown = True
                self.graph_condition_entry.pack_forget()
            
            # 期間設定フレームを表示
            if hasattr(self, 'period_frame'):
                if hasattr(self, 'latest_dates_frame'):
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5, before=self.latest_dates_frame)
                else:
                    self.period_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # デフォルト期間を設定（30日前からさかのぼって1年間）
            if hasattr(self, 'start_date_entry') and hasattr(self, 'end_date_entry'):
                from datetime import datetime, timedelta
                today = datetime.now()
                end_date_dt = today - timedelta(days=30)
                end_date = end_date_dt.strftime('%Y-%m-%d')
                start_date_dt = end_date_dt - timedelta(days=365)
                start_date = start_date_dt.strftime('%Y-%m-%d')
                self.start_date_entry.delete(0, tk.END)
                self.start_date_entry.insert(0, start_date)
                self.start_date_entry.config(foreground="black")
                self.start_date_placeholder_shown = False
                self.end_date_entry.delete(0, tk.END)
                self.end_date_entry.insert(0, end_date)
                self.end_date_entry.config(foreground="black")
                self.end_date_placeholder_shown = False
        else:
            # タイプが空欄・項目一括変更など：コマンド入力を有効化
            if hasattr(self, 'conception_rate_frame'):
                self.conception_rate_frame.pack_forget()
            if hasattr(self, 'condition_frame'):
                self.condition_frame.pack_forget()
            if hasattr(self, 'period_frame'):
                self.period_frame.pack_forget()
            if hasattr(self, 'graph_input_frame'):
                self.graph_input_frame.pack_forget()
            if hasattr(self, 'event_selection_frame'):
                self.event_selection_frame.pack_forget()
            # コマンド入力欄を有効化（履歴プルダウンを復元）
            if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
                self.command_entry.config(state="normal")
                self._update_command_history_dropdown()
            
            # プレースホルダーをクリア
            if self.command_placeholder_shown:
                current_value = self.command_entry.get()
                # プレースホルダーの場合はクリア
                if current_value in ["項目１ 項目２ ・・・", "集計したい項目を選んでください", self.command_placeholder]:
                    self.command_entry.set("")
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
    
    def _on_command_type_selected_text(self, event=None):
        """コマンドタイプが選択された時（command_text用）"""
        # タイプ選択時はプレフィックスを表示しない（実行時に自動付与）
        # ユーザーは直接項目を入力できる
        pass
    
    def _on_graph_type_selected(self, event=None):
        """グラフ種類が選択された時にY軸とX軸の表示を制御"""
        selected_type = self.graph_type_entry.get()
        
        # 一旦すべての要素を非表示
        if hasattr(self, 'graph_y_label'):
            self.graph_y_label.pack_forget()
        if hasattr(self, 'graph_y_entry'):
            self.graph_y_entry.pack_forget()
        if hasattr(self, 'graph_y_item_list_btn'):
            self.graph_y_item_list_btn.pack_forget()
        if hasattr(self, 'graph_x_label'):
            self.graph_x_label.pack_forget()
        if hasattr(self, 'graph_x_entry'):
            self.graph_x_entry.pack_forget()
        if hasattr(self, 'graph_x_item_list_btn'):
            self.graph_x_item_list_btn.pack_forget()
        if hasattr(self, 'graph_classification_label'):
            self.graph_classification_label.pack_forget()
        if hasattr(self, 'graph_classification_entry'):
            self.graph_classification_entry.pack_forget()
        if hasattr(self, 'graph_condition_label'):
            self.graph_condition_label.pack_forget()
        if hasattr(self, 'graph_condition_entry'):
            self.graph_condition_entry.pack_forget()
        if hasattr(self, 'graph_classification_event_label'):
            self.graph_classification_event_label.pack_forget()
        if hasattr(self, 'graph_classification_event_entry'):
            self.graph_classification_event_entry.pack_forget()
        if hasattr(self, 'graph_classification_item_label'):
            self.graph_classification_item_label.pack_forget()
        if hasattr(self, 'graph_classification_item_entry'):
            self.graph_classification_item_entry.pack_forget()
        if hasattr(self, 'graph_classification_item_list_btn'):
            self.graph_classification_item_list_btn.pack_forget()
        if hasattr(self, 'graph_classification_bin_label'):
            self.graph_classification_bin_label.pack_forget()
        if hasattr(self, 'graph_classification_bin_entry'):
            self.graph_classification_bin_entry.pack_forget()
        
        # グラフ種類に応じて表示を制御（散布図・空胎日数生存曲線）
        if selected_type == "散布図":
            # 散布図：全分類オプションを復元
            if hasattr(self, 'graph_classification_entry'):
                self.graph_classification_entry['values'] = self.graph_classification_options
            # 散布図：Y軸、X軸、分類、条件の順に表示
            if hasattr(self, 'graph_y_label'):
                self.graph_y_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_y_entry'):
                self.graph_y_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_y_item_list_btn'):
                self.graph_y_item_list_btn.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_x_label'):
                self.graph_x_label.config(text="X軸:")
                self.graph_x_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_x_entry'):
                self.graph_x_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_x_item_list_btn'):
                self.graph_x_item_list_btn.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_classification_label'):
                self.graph_classification_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_classification_entry'):
                self.graph_classification_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_condition_label'):
                self.graph_condition_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_condition_entry'):
                self.graph_condition_entry.pack(side=tk.LEFT, padx=5)
        elif selected_type == "空胎日数生存曲線":
            # 空胎日数生存曲線：分類は産次・産次（１産、２産、３産以上）のみ
            if hasattr(self, 'graph_classification_entry'):
                self.graph_classification_entry['values'] = self.graph_classification_options_survival
                self.graph_classification_entry.set("")
            if hasattr(self, 'graph_classification_label'):
                self.graph_classification_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_classification_entry'):
                self.graph_classification_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_condition_label'):
                self.graph_condition_label.pack(side=tk.LEFT, padx=(10, 5))
            if hasattr(self, 'graph_condition_entry'):
                self.graph_condition_entry.pack(side=tk.LEFT, padx=5)
            self._on_graph_classification_selected()
        else:
            # グラフ種類が選択されていない場合、すべて非表示
            pass
    
    def _on_graph_classification_selected(self, event=None):
        """グラフの分類が選択されたとき、生存曲線で「イベントの有無」なら対象イベント、「項目で分類」なら対象項目+区分を表示"""
        if not hasattr(self, 'graph_type_entry') or self.graph_type_entry.get() != "空胎日数生存曲線":
            for attr in ('graph_classification_event_label', 'graph_classification_event_entry',
                         'graph_classification_item_label', 'graph_classification_item_entry',
                         'graph_classification_item_list_btn',
                         'graph_classification_bin_label', 'graph_classification_bin_entry'):
                if hasattr(self, attr):
                    getattr(self, attr).pack_forget()
            return
        cl = hasattr(self, 'graph_classification_entry') and self.graph_classification_entry.get()
        # イベントの有無
        if cl == "イベントの有無":
            if hasattr(self, 'graph_classification_event_label') and hasattr(self, 'graph_classification_event_entry'):
                event_options = self._get_graph_event_options()
                self.graph_classification_event_entry['values'] = event_options
                if event_options and not self.graph_classification_event_entry.get():
                    self.graph_classification_event_entry.set(event_options[0])
                self.graph_classification_event_label.pack(side=tk.LEFT, padx=(10, 5))
                self.graph_classification_event_entry.pack(side=tk.LEFT, padx=5)
            for attr in ('graph_classification_item_label', 'graph_classification_item_entry',
                         'graph_classification_item_list_btn',
                         'graph_classification_bin_label', 'graph_classification_bin_entry'):
                if hasattr(self, attr):
                    getattr(self, attr).pack_forget()
            return
        # 項目で分類
        if cl == "項目で分類":
            if hasattr(self, 'graph_classification_item_label') and hasattr(self, 'graph_classification_item_entry'):
                item_options = self._get_survival_classification_item_options()
                self.graph_classification_item_entry['values'] = item_options
                if item_options and not self.graph_classification_item_entry.get():
                    self.graph_classification_item_entry.set(item_options[0])
                self.graph_classification_item_label.pack(side=tk.LEFT, padx=(10, 5))
                self.graph_classification_item_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_classification_item_list_btn'):
                self.graph_classification_item_list_btn.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_classification_bin_label') and hasattr(self, 'graph_classification_bin_entry'):
                self.graph_classification_bin_label.pack(side=tk.LEFT, padx=(10, 5))
                self.graph_classification_bin_entry.pack(side=tk.LEFT, padx=5)
            if hasattr(self, 'graph_classification_event_label'):
                self.graph_classification_event_label.pack_forget()
            if hasattr(self, 'graph_classification_event_entry'):
                self.graph_classification_event_entry.pack_forget()
            return
        # その他
        for attr in ('graph_classification_event_label', 'graph_classification_event_entry',
                     'graph_classification_item_label', 'graph_classification_item_entry',
                     'graph_classification_item_list_btn',
                     'graph_classification_bin_label', 'graph_classification_bin_entry'):
            if hasattr(self, attr):
                getattr(self, attr).pack_forget()
    
    def _get_survival_classification_item_options(self):
        """生存曲線「項目で分類」用の項目一覧（key: 表示名）。item_dictionary から取得。"""
        options = []
        try:
            selectable = self._get_selectable_items_with_category()
            if selectable:
                seen = set()
                for code, display_name, _cat in selectable:
                    if not code or code in seen:
                        continue
                    seen.add(code)
                    options.append(f"{code}: {display_name}")
        except Exception as e:
            logging.debug(f"生存曲線用項目一覧取得エラー: {e}")
        if not options:
            options = ["LACT: 産次", "DIM: 分娩後日数", "CALVMO: 分娩月", "BRD: 品種", "PEN: 群"]
        return options
    
    def _get_graph_event_options(self):
        """グラフ用のイベント一覧（番号: 名前）を取得。event_dictionary から読み込む。"""
        event_options = []
        paths_to_try = []
        if hasattr(self, 'farm_path') and self.farm_path:
            farm_config = Path(self.farm_path) / "config" / "event_dictionary.json"
            if farm_config.exists():
                paths_to_try.append(farm_config)
        if hasattr(self, 'event_dict_path') and self.event_dict_path:
            paths_to_try.append(Path(self.event_dict_path) if not isinstance(self.event_dict_path, Path) else self.event_dict_path)
        if not paths_to_try:
            paths_to_try.append(Path(__file__).resolve().parent.parent / "config_default" / "event_dictionary.json")
        for path in paths_to_try:
            path = Path(path) if not isinstance(path, Path) else path
            if not path.exists():
                continue
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for event_num_str, event_data in data.items():
                    if event_data.get('deprecated', False):
                        continue
                    try:
                        event_number = int(event_num_str)
                    except (ValueError, TypeError):
                        continue
                    name = event_data.get('name_jp', event_data.get('alias', str(event_number)))
                    event_options.append(f"{event_number}: {name}")
                if event_options:
                    break
            except Exception as e:
                logging.debug(f"イベント辞書読み込みスキップ {path}: {e}")
        return sorted(event_options, key=lambda s: int(s.split(":", 1)[0].strip()) if ":" in s else 0)
    
    def _on_show_item_list(self):
        """項目一覧ウィンドウを表示（個体カードと同じダイアログを使用）"""
        # 個体カードと同じ項目選択ダイアログを表示
        self._show_item_selection_dialog_for_command()
    
    def _on_show_item_list_for_graph_x(self):
        """グラフX軸用の項目一覧ウィンドウを表示"""
        self._show_item_selection_dialog_for_graph_field("x")
    
    def _on_show_item_list_for_graph_y(self):
        """グラフY軸用の項目一覧ウィンドウを表示"""
        self._show_item_selection_dialog_for_graph_field("y")
    
    def _on_show_item_list_for_survival_classification_item(self):
        """生存曲線「項目で分類」の対象項目を項目一覧ダイアログで選択"""
        self._show_item_selection_dialog_for_graph_field("survival_classification_item")
    
    def _show_item_selection_dialog_for_graph_field(self, field_type: str):
        """グラフ用の項目選択ダイアログ（X軸またはY軸）"""
        from tkinter import messagebox
        
        # 選択可能な項目を取得
        selectable = self._get_selectable_items_with_category()
        if not selectable:
            messagebox.showerror("エラー", "選択可能な項目がありません。item_dictionary.json を確認してください。")
            return
        
        dialog = tk.Toplevel(self.root)
        if field_type == "survival_classification_item":
            dialog.title("分類の対象項目を選択")
        else:
            dialog.title("項目を選択")
        dialog.geometry("500x600")
        
        # カテゴリー名のマッピング
        category_names = {
            "REPRODUCTION": "繁殖",
            "DHI": "乳検",
            "GENOMIC": "ゲノム",
            "GENOME": "ゲノム",
            "HEALTH": "疾病",
            "CORE": "基本",
            "OTHERS": "その他",
            "USER": "ユーザー"
        }
        
        # メインフレーム
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 検索フレーム
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="").pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_right_frame = ttk.Frame(search_frame)
        search_right_frame.pack(side=tk.RIGHT)
        
        # クリアボタン（後でclear_search関数のコマンドを設定）
        clear_button = ttk.Button(search_right_frame, text="クリア", width=8)
        clear_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_right_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.RIGHT, padx=(0, 5))
        ttk.Label(search_right_frame, text="検索:", font=('', 10)).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree = ttk.Treeview(
            tree_frame,
            columns=('code',),
            show='tree headings',
            height=20,
            yscrollcommand=scrollbar.set,
            selectmode='extended'
        )
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
        
        for cat in category_items:
            category_items[cat].sort(key=lambda x: x[1])
        
        category_order = ["REPRODUCTION", "DHI", "GENOMIC", "GENOME", "HEALTH", "CORE", "USER", "OTHERS"]
        extra_categories = sorted([c for c in category_items.keys() if c not in category_order])
        category_order.extend(extra_categories)
        
        # ロガーを取得
        logger = logging.getLogger(__name__)
        
        category_nodes: Dict[str, str] = {}
        # マスターリスト：node_idを含めず、(code, name, cat)のみを保持
        master_items: List[tuple] = []
        
        for cat in category_order:
            if cat in category_items and category_items[cat]:
                cat_name = category_names.get(cat, cat)
                cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                category_nodes[cat] = cat_node
                
                for code, name in category_items[cat]:
                    item_node = tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                    master_items.append((code, name, cat))
        
        tree.tag_configure('category', font=('', 10, 'bold'), background='#E0E0E0')
        tree.tag_configure('item', font=('', 10))
        
        for cat_node in category_nodes.values():
            tree.item(cat_node, open=True)
        
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
        
        def _get_selected_item_code():
            sel = tree.selection()
            if not sel:
                return None
            node_id = sel[0]
            tags = tree.item(node_id, 'tags')
            if tags and tags[0] == 'category':
                return None
            values = tree.item(node_id, 'values')
            if values and values[0]:
                return values[0]
            return None
        
        def _get_selected_item_code_and_name():
            """選択された項目の (code, display_name) を返す。未選択なら None。"""
            sel = tree.selection()
            if not sel:
                return None
            node_id = sel[0]
            tags = tree.item(node_id, 'tags')
            if tags and tags[0] == 'category':
                return None
            values = tree.item(node_id, 'values')
            if not values or not values[0]:
                return None
            code = values[0]
            name = tree.item(node_id, 'text') or code
            return (code, name)
        
        def on_item_selected(event=None):
            if field_type == "survival_classification_item":
                pair = _get_selected_item_code_and_name()
                if pair:
                    code, name = pair
                    if hasattr(self, 'graph_classification_item_entry'):
                        self.graph_classification_item_entry.set(f"{code}: {name}")
                    dialog.destroy()
                return
            code = _get_selected_item_code()
            if code:
                if field_type == "x":
                    self.graph_x_entry.delete(0, tk.END)
                    self.graph_x_entry.insert(0, code)
                elif field_type == "y":
                    self.graph_y_entry.delete(0, tk.END)
                    self.graph_y_entry.insert(0, code)
                dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        def on_double_click(event):
            """ダブルクリック時の処理"""
            # クリックされた行を取得
            item = tree.identify_row(event.y)
            if item:
                # 行を選択
                tree.selection_set(item)
                tree.focus(item)
                # 項目選択処理を実行
                on_item_selected()
        
        tree.bind('<Double-Button-1>', on_double_click)
        tree.bind('<Return>', lambda e: on_item_selected())
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(10, 0))
        ttk.Button(btn_frame, text="OK", command=on_item_selected, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        search_entry.focus()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
    
    def _show_item_selection_dialog_for_command(self, initial_code: Optional[str] = None) -> None:
        """Item Dictionary から計算項目を選択（カテゴリー別表示・検索機能付き）"""
        from tkinter import messagebox
        
        # 選択可能な項目を取得
        selectable = self._get_selectable_items_with_category()
        if not selectable:
            messagebox.showerror("エラー", "選択可能な項目がありません。item_dictionary.json を確認してください。")
            return None
        
        dialog = tk.Toplevel(self.root)
        dialog.title("項目を選択")
        dialog.geometry("500x600")
        
        # カテゴリー名のマッピング
        category_names = {
            "REPRODUCTION": "繁殖",
            "DHI": "乳検",
            "GENOMIC": "ゲノム",
            "GENOME": "ゲノム",
            "HEALTH": "疾病",
            "CORE": "基本",
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
        ttk.Label(search_right_frame, text="検索:", font=('', 10)).pack(side=tk.RIGHT, padx=(5, 0))
        
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
        tree = ttk.Treeview(
            tree_frame,
            columns=('code',),
            show='tree headings',
            height=20,
            yscrollcommand=scrollbar.set,
            selectmode='extended'
        )
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
        category_order = ["REPRODUCTION", "DHI", "GENOMIC", "GENOME", "HEALTH", "CORE", "USER", "OTHERS"]
        extra_categories = sorted([c for c in category_items.keys() if c not in category_order])
        category_order.extend(extra_categories)
        
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
        tree.tag_configure('category', font=('', 10, 'bold'), background='#E0E0E0')
        tree.tag_configure('item', font=('', 10))
        
        # 初期状態でカテゴリーを展開
        for cat_node in category_nodes.values():
            tree.item(cat_node, open=True)
        
        selected_codes: List[str] = []
        
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
        search_entry.bind('<KeyRelease>', lambda e: filter_tree())
        
        def _get_selected_item_codes() -> List[str]:
            sel = tree.selection()
            selected = []
            for node_id in sel:
                tags = tree.item(node_id, 'tags')
                if tags and tags[0] == 'category':
                    continue
                values = tree.item(node_id, 'values')
                if values and values[0]:
                    selected.append(values[0])
            return selected

        def _add_selected_items(close_dialog: bool = False):
            selected = _get_selected_item_codes()
            if not selected:
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            for code in selected:
                self._on_item_selected_from_list(code)
            selected_codes.clear()
            selected_codes.extend(selected)
            if close_dialog:
                dialog.destroy()
        
        def on_ok():
            _add_selected_items(close_dialog=True)
        
        def on_cancel():
            dialog.destroy()
        
        tree.bind('<Double-Button-1>', lambda e: _add_selected_items(close_dialog=False))
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
        return None
    
    def _get_selectable_items_with_category(self) -> List[tuple]:
        """項目辞書の全項目を取得（カテゴリー情報付き）"""
        # 常に最新の辞書を反映
        self._load_item_dictionary()
        item_dict = self.item_dictionary or {}
        if not item_dict:
            item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}

        items: List[tuple] = []
        for code, data in item_dict.items():
            if not isinstance(data, dict):
                data = {}
            if data.get("visible", True) is False:
                continue
            # label, display_name, name_jp の順で取得
            display_name = data.get("label") or data.get("display_name") or data.get("name_jp") or code
            # categoryを取得（大文字に変換）
            category = (data.get("category") or "OTHERS").upper()
            items.append((code, display_name, category))
        items.sort(key=lambda x: (x[2], x[1]))  # カテゴリー、名前の順でソート
        return items
    
    def _on_item_selected_from_list(self, item_key: str):
        """
        項目一覧から項目が選択された時のコールバック。
        選択された項目は日本語名（display_name）でコマンド欄に転記し、
        略称で直接入力した場合も既存の解析で両方とも正しく解釈される。
        
        Args:
            item_key: 選択された項目のキー（略称）
        """
        # 表示用には日本語名を使用（慣れているユーザーは略称を直接打ち込む想定で、解析側で略称・日本語名両対応）
        item_def = self.item_dictionary.get(item_key, {})
        display_name = (
            item_def.get("display_name")
            or item_def.get("label")
            or item_def.get("name_jp")
            or item_key
        )
        text_to_insert = display_name

        # コマンド入力欄に転記
        if hasattr(self, 'command_entry') and self.command_entry.winfo_exists():
            # 単一行Comboboxの場合
            current_text = self.command_entry.get().strip()
            # プレースホルダーが表示中なら解除
            if self.command_placeholder_shown:
                current_text = ""
                self.command_entry.config(foreground="black")
                self.command_placeholder_shown = False
            existing_items = current_text.split() if current_text else []
            # 既に同じ項目が入っていれば追加しない（略称・日本語名のどちらで入っていても判定）
            if item_key in existing_items or display_name in existing_items:
                self.command_entry.focus_set()
                return
            if current_text and current_text not in ["項目１ 項目２ ・・・", "集計したい項目を選んでください", self.command_placeholder]:
                self.command_entry.set(current_text + " " + text_to_insert)
            else:
                self.command_entry.set(text_to_insert)
            self.command_entry.focus_set()
        elif hasattr(self, 'command_text') and self.command_text.winfo_exists():
            # 複数行Textの場合
            current_text = self.command_text.get("1.0", tk.END).strip()
            if self.command_placeholder_shown:
                current_text = ""
                self._hide_command_placeholder()
                self.command_placeholder_shown = False
            existing_items = current_text.split() if current_text else []
            if item_key in existing_items or display_name in existing_items:
                self.command_text.focus_set()
                return
            if current_text and not self.command_placeholder_shown:
                self.command_text.insert(tk.END, " " + text_to_insert)
            else:
                self._hide_command_placeholder()
                self.command_text.insert("1.0", text_to_insert)
            self.command_text.focus_set()
    
    def _ask_month_input(self, prompt: str) -> Optional[str]:
        """
        月の入力を求めるダイアログを表示
        
        Args:
            prompt: プロンプトメッセージ
        
        Returns:
            入力された月（文字列）、キャンセルされた場合はNone
        """
        dialog = tk.Toplevel(self.root)
        dialog.title("月の指定")
        dialog.geometry("400x150")
        
        # プロンプト
        prompt_label = ttk.Label(dialog, text=prompt)
        prompt_label.pack(pady=10)
        
        # 入力フィールド
        month_var = tk.StringVar()
        month_entry = ttk.Entry(dialog, textvariable=month_var, width=10)
        month_entry.pack(pady=5)
        month_entry.focus_set()
        
        result = [None]  # 結果を格納するリスト（クロージャで変更可能にするため）
        
        def on_ok():
            result[0] = month_var.get().strip()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # EnterキーでOK
        month_entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())
        
        # ダイアログが閉じられるまで待機
        dialog.wait_window()
        
        return result[0]

    def _prompt_monthly_milk_test_months(self, need_month_1: bool, need_month_2: bool) -> bool:
        """任意月の乳検データ用の月を毎回指定してもらう"""
        settings = get_app_settings_manager()
        if need_month_1:
            month_input = self._ask_month_input("任意月の乳検データ１の月を指定してください（1-12）")
            if month_input is None:
                return False
            try:
                month_1 = int(month_input)
                if month_1 < 1 or month_1 > 12:
                    self.add_message(role="system", text="月は1-12の範囲で指定してください。")
                    return False
                settings.set_selected_month_1(month_1)
            except ValueError:
                self.add_message(role="system", text="有効な月を入力してください（1-12）。")
                return False
        if need_month_2:
            month_input = self._ask_month_input("任意月の乳検データ２の月を指定してください（1-12）")
            if month_input is None:
                return False
            try:
                month_2 = int(month_input)
                if month_2 < 1 or month_2 > 12:
                    self.add_message(role="system", text="月は1-12の範囲で指定してください。")
                    return False
                settings.set_selected_month_2(month_2)
            except ValueError:
                self.add_message(role="system", text="有効な月を入力してください（1-12）。")
                return False
        return True

    def _detect_monthly_milk_test_items_in_text(self, text: str) -> Tuple[bool, bool]:
        """コマンド文字列から任意月の乳検項目の有無を判定"""
        monthly_keys_1 = {'MILK1M', 'SCC1M', 'LS1M', 'MUN1M', 'PROT1M', 'FAT1M', 'DENFA1M'}
        monthly_keys_2 = {'MILK2M', 'SCC2M', 'LS2M', 'MUN2M', 'PROT2M', 'FAT2M', 'DENFA2M'}
        text_lower = text.lower()

        need_month_1 = any(k.lower() in text_lower for k in monthly_keys_1)
        need_month_2 = any(k.lower() in text_lower for k in monthly_keys_2)

        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        for key in monthly_keys_1:
            name = item_dict.get(key, {}).get("display_name")
            if name and name in text:
                need_month_1 = True
                break
        for key in monthly_keys_2:
            name = item_dict.get(key, {}).get("display_name")
            if name and name in text:
                need_month_2 = True
                break

        return need_month_1, need_month_2

    def _prompt_monthly_milk_test_months_for_command(self, command_text: str) -> bool:
        """コマンド内に任意月の乳検項目があれば月指定を求める"""
        need_month_1, need_month_2 = self._detect_monthly_milk_test_items_in_text(command_text)
        if not (need_month_1 or need_month_2):
            return True
        return self._prompt_monthly_milk_test_months(need_month_1=need_month_1, need_month_2=need_month_2)
    
    def _parse_condition_text_to_or_groups(self, condition_text: str):
        """
        条件文字列をパースし、OR グループのリストを返す。
        - 「 または 」または「 OR 」で OR 区切り。各グループ内は AND（スペース区切り）。
        - 項目名は日本語可（例：フレッシュチェックNOTE）。＝は=に正規化。
        - 完全一致: 項目=値 または 項目：値。部分一致（文字列に含む）: 項目>=値。
        
        Returns:
            List[List[dict]]: 各要素は AND 条件のリスト。いずれか 1 グループが成立すれば OR で真。
        """
        import re
        if not condition_text or not condition_text.strip():
            return []
        # 全角＝を半角に、全角スペースを半角に
        s = condition_text.replace('＝', '=').replace('　', ' ')
        # OR で分割（ または / OR ）
        or_parts = re.split(r'\s+または\s+|\s+OR\s+', s, flags=re.IGNORECASE)
        or_groups = []
        for group_str in or_parts:
            group_str = group_str.strip()
            if not group_str:
                continue
            tokens = group_str.split()
            group_conditions = []
            for p in tokens:
                p = p.strip()
                if not p:
                    continue
                if p == "受胎":
                    group_conditions.append({'item_name': 'RC', 'operator': '=', 'value': [5, 6]})
                    continue
                if p == "空胎":
                    group_conditions.append({'item_name': 'RC', 'operator': '=', 'value': [1, 2, 3, 4]})
                    continue
                # 項目 演算子 値（項目名は日本語可）
                match = re.match(r'^(.+?)\s*(>=|<=|!=|==|=|<|>)\s*(.*)$', p)
                if match:
                    item_name = match.group(1).strip()
                    operator = match.group(2)
                    value_str = match.group(3).strip()
                    # 範囲指定（例：DIM=100-150）
                    if (operator == "=" or operator == "==") and '-' in value_str:
                        range_parts = value_str.split('-', 1)
                        if len(range_parts) == 2:
                            try:
                                min_value = float(range_parts[0].strip())
                                max_value = float(range_parts[1].strip())
                                group_conditions.append({
                                    'item_name': item_name,
                                    'operator': '>=',
                                    'value': str(min_value),
                                    'is_range': True,
                                    'range_max': str(max_value)
                                })
                            except (ValueError, TypeError):
                                group_conditions.append({'item_name': item_name, 'operator': operator, 'value': value_str})
                        else:
                            group_conditions.append({'item_name': item_name, 'operator': operator, 'value': value_str})
                    elif ',' in value_str and (operator == "=" or operator == "=="):
                        values = [v.strip() for v in value_str.split(',')]
                        group_conditions.append({'item_name': item_name, 'operator': operator, 'value': values})
                    else:
                        group_conditions.append({'item_name': item_name, 'operator': operator, 'value': value_str})
                    continue
                # 項目：値（完全一致）
                if "：" in p or ":" in p:
                    sep = "：" if "：" in p else ":"
                    parts_colon = p.split(sep, 1)
                    if len(parts_colon) == 2:
                        group_conditions.append({
                            'item_name': parts_colon[0].strip(),
                            'operator': '=',
                            'value': parts_colon[1].strip()
                        })
            if group_conditions:
                or_groups.append(group_conditions)
        return or_groups
    
    def _list_row_matches_condition_group(self, row_dict, group, display_name_to_item_key, item_key_to_display_name,
                                          item_key_lower_to_item_key, display_name_lower_to_display_name, calc_item_keys):
        """1行が1グループの条件をすべて満たすか（AND）。文字列の>=は部分一致。"""
        import re
        for filter_condition in group:
            item_name = filter_condition['item_name']
            operator = filter_condition['operator']
            value_str = filter_condition['value']
            item_name_lower = item_name.lower()
            normalized_item_name = None
            item_key = None
            if item_name_lower in display_name_lower_to_display_name:
                normalized_item_name = display_name_lower_to_display_name[item_name_lower]
                item_key = display_name_to_item_key.get(normalized_item_name)
            elif item_name_lower in item_key_lower_to_item_key:
                item_key = item_key_lower_to_item_key[item_name_lower]
                if item_key in item_key_to_display_name:
                    normalized_item_name = item_key_to_display_name[item_key]
            row_value = None
            is_rc_condition = (item_key == "RC" or normalized_item_name in ("繁殖コード", "繁殖区分"))
            if is_rc_condition:
                cow_auto_id = row_dict.get("auto_id")
                if cow_auto_id and self.formula_engine:
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and "RC" in calculated:
                        row_value = calculated["RC"]
            else:
                if normalized_item_name:
                    for key in row_dict.keys():
                        if key.lower() == normalized_item_name.lower():
                            row_value = row_dict[key]
                            break
                if row_value is None and item_key and item_key in row_dict:
                    row_value = row_dict[item_key]
                if row_value is None:
                    cow_auto_id = row_dict.get("auto_id")
                    if cow_auto_id and self.formula_engine:
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and item_key and item_key in calculated:
                            row_value = calculated[item_key]
            if row_value is None:
                return False
            if is_rc_condition and isinstance(row_value, str):
                m = re.match(r'^(\d+)', str(row_value))
                if m:
                    row_value = int(m.group(1))
            if filter_condition.get('is_range', False):
                try:
                    row_value_num = float(row_value)
                    min_v = float(value_str)
                    max_v = float(filter_condition.get('range_max', value_str))
                    if not (row_value_num >= min_v and row_value_num <= max_v):
                        return False
                except (ValueError, TypeError):
                    return False
            elif isinstance(value_str, list):
                match = False
                for val in value_str:
                    try:
                        if float(row_value) == float(val):
                            match = True
                            break
                    except (ValueError, TypeError):
                        if str(row_value) == str(val):
                            match = True
                            break
                if not match:
                    return False
            else:
                try:
                    row_value_num = float(row_value)
                    condition_value_num = float(value_str)
                    if operator == "<" and not (row_value_num < condition_value_num): return False
                    if operator == ">" and not (row_value_num > condition_value_num): return False
                    if operator == "<=" and not (row_value_num <= condition_value_num): return False
                    if operator == ">=" and not (row_value_num >= condition_value_num): return False
                    if operator in ("=", "==") and not (row_value_num == condition_value_num): return False
                    if operator in ("!=", "<>") and not (row_value_num != condition_value_num): return False
                except (ValueError, TypeError):
                    if operator in ("=", "==") and str(row_value).strip() != str(value_str).strip():
                        return False
                    if operator in ("!=", "<>") and str(row_value).strip() == str(value_str).strip():
                        return False
                    if operator == ">=" and str(value_str).strip() not in str(row_value):
                        return False
        return True
    
    def _aggregate_cow_matches_condition_group(self, cow_row, group, display_name_to_item_key, item_key_to_display_name,
                                               item_key_lower_to_item_key, display_name_lower_to_display_name, calc_item_keys):
        """1頭が1グループの条件をすべて満たすか（集計用）。文字列の>=は部分一致。"""
        for filter_condition in group:
            condition_item_name = filter_condition['item_name']
            operator = filter_condition['operator']
            value_str = filter_condition['value']
            condition_item_name_lower = condition_item_name.lower()
            condition_normalized_item_name = None
            condition_item_key = None
            if condition_item_name_lower in display_name_lower_to_display_name:
                condition_normalized_item_name = display_name_lower_to_display_name[condition_item_name_lower]
                condition_item_key = display_name_to_item_key.get(condition_normalized_item_name)
            elif condition_item_name_lower in item_key_lower_to_item_key:
                condition_item_key = item_key_lower_to_item_key[condition_item_name_lower]
                if condition_item_key in item_key_to_display_name:
                    condition_normalized_item_name = item_key_to_display_name[condition_item_key]
            is_lact_condition = (condition_item_name_lower == "lact" or condition_item_key == "LACT")
            if is_lact_condition:
                condition_item_key = "lact"
                condition_normalized_item_name = "産次"
            row_value = None
            is_rc_condition = (condition_item_key == "RC" or condition_normalized_item_name in ("繁殖コード", "繁殖区分"))
            cow_auto_id = cow_row['auto_id']
            if is_rc_condition:
                try:
                    state = self.rule_engine.apply_events(cow_auto_id)
                    if state and 'rc' in state:
                        row_value = state['rc']
                except Exception:
                    pass
            if row_value is None and is_lact_condition:
                cow = self.db.get_cow_by_id(cow_row['ID'])
                if cow and 'lact' in cow:
                    row_value = cow['lact']
            elif row_value is None and condition_item_key and condition_item_key in calc_item_keys:
                calculated = self.formula_engine.calculate(cow_auto_id)
                if calculated and condition_item_key in calculated:
                    row_value = calculated[condition_item_key]
            elif row_value is None and condition_item_key:
                cow = self.db.get_cow_by_id(cow_row['ID'])
                if cow and condition_item_key in cow:
                    row_value = cow[condition_item_key]
                elif is_rc_condition and cow and 'rc' in cow:
                    row_value = cow['rc']
            if row_value is None:
                return False
            if filter_condition.get('is_range', False):
                try:
                    row_value_num = float(row_value)
                    min_v = float(value_str)
                    max_v = float(filter_condition.get('range_max', value_str))
                    if not (row_value_num >= min_v and row_value_num <= max_v):
                        return False
                except (ValueError, TypeError):
                    return False
            elif isinstance(value_str, list):
                match = False
                for val in value_str:
                    try:
                        if float(row_value) == float(val):
                            match = True
                            break
                    except (ValueError, TypeError):
                        if str(row_value) == str(val):
                            match = True
                            break
                if not match:
                    return False
            else:
                try:
                    row_value_num = float(row_value)
                    condition_value_num = float(value_str)
                    if operator == "<" and not (row_value_num < condition_value_num): return False
                    if operator == ">" and not (row_value_num > condition_value_num): return False
                    if operator == "<=" and not (row_value_num <= condition_value_num): return False
                    if operator == ">=" and not (row_value_num >= condition_value_num): return False
                    if operator in ("=", "==") and not (row_value_num == condition_value_num): return False
                    if operator in ("!=", "<>") and not (row_value_num != condition_value_num): return False
                except (ValueError, TypeError):
                    if operator in ("=", "==") and str(row_value).strip() != str(value_str).strip():
                        return False
                    if operator in ("!=", "<>") and str(row_value).strip() == str(value_str).strip():
                        return False
                    if operator == ">=" and str(value_str).strip() not in str(row_value):
                        return False
        return True
    
    def _execute_list_command(self, command: str):
        """
        リストコマンドを実行
        
        Args:
            command: リストコマンド（例：「リスト：ID」「リスト：ID 分娩後日数 産次」）
        """
        try:
            # 元のコマンドを保存（ID追加後に更新するため）
            original_command = command
            
            # 「リスト：」または「リスト:」をすべて除去（二重プレフィックス対策）
            list_text = command.strip()
            while list_text.startswith("リスト：") or list_text.startswith("リスト:"):
                list_text = (list_text[4:].strip() if list_text.startswith("リスト：") else list_text[4:].strip())
            
            if not list_text:
                self.add_message(role="system", text="リストコマンドに項目を指定してください。\n例：リスト：ID または リスト：ID 分娩後日数 産次")
                return
            
            # 項目名・条件・並び順を分離（リスト：項目：条件：昇順/降順 項目名、最大3部分）
            import re
            sep = "：" if "：" in list_text else ":"
            parts = [p.strip() for p in list_text.split(sep, 2)]  # 最大3部分に分割
            items_text = parts[0] if len(parts) >= 1 else ""
            condition_text = parts[1] if len(parts) >= 2 else ""
            sort_text = parts[2] if len(parts) >= 3 else ""
            # 2部分のみの場合、2番目が「昇順 項目名」「降順 項目名」なら並び順として扱う（条件なし）
            if len(parts) == 2 and parts[1]:
                p1 = parts[1]
                if p1.startswith("昇順 ") or p1.startswith("降順 "):
                    condition_text = ""
                    sort_text = p1
                    sort_text = sort_text.strip()
            
            # 並び順の解釈（昇順 項目名 / 降順 項目名）
            sort_order_asc = True  # デフォルトはID昇順
            sort_column_display_name = None
            if sort_text:
                if sort_text.startswith("昇順 "):
                    sort_order_asc = True
                    sort_column_display_name = sort_text[2:].strip()
                elif sort_text.startswith("降順 "):
                    sort_order_asc = False
                    sort_column_display_name = sort_text[2:].strip()
                if not sort_column_display_name:
                    sort_column_display_name = None
            
            # 項目リストを分割
            columns = []
            if items_text:
                columns = items_text.replace("　", " ").split()  # 全角スペースを半角に変換してから分割
            
            # 条件を解析（OR 対応・日本語項目名・＝正規化は _parse_condition_text_to_or_groups で統一）
            or_groups = self._parse_condition_text_to_or_groups(condition_text) if condition_text else []
            # SQL/計算用には OR が1グループのときだけそのグループを使用。複数グループは後で Python で OR フィルタ
            filter_conditions = or_groups[0] if len(or_groups) == 1 else []
            
            # 既存のconditions辞書は空に（後で使わない）
            conditions = {}
            
            # item_dictionary.jsonを読み込んで、表示名から項目キーへのマッピングを作成
            display_name_to_item_key = {}
            item_key_to_display_name = {}  # 項目キーから表示名へのマッピング
            item_key_lower_to_item_key = {}  # 小文字の項目キー -> 元の項目キー（大文字小文字を区別しない検索用）
            display_name_lower_to_display_name = {}  # 小文字の表示名 -> 元の表示名（大文字小文字を区別しない検索用）
            calc_items = set()  # 計算項目の表示名セット
            calc_item_keys = set()  # 計算項目の項目キーセット
            source_items = set()  # ソース項目（乳検・ゲノム等イベント由来）の表示名セット
            source_item_keys = set()  # ソース項目の項目キーセット
            
            if self.item_dict_path and self.item_dict_path.exists():
                try:
                    with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                        item_dict = json.load(f)
                    for item_key, item_def in item_dict.items():
                        display_name = item_def.get("display_name", "")
                        if display_name:
                            display_name_to_item_key[display_name] = item_key
                            item_key_to_display_name[item_key] = display_name
                            item_key_lower_to_item_key[item_key.lower()] = item_key
                            display_name_lower_to_display_name[display_name.lower()] = display_name
                            # 別名（alias）でも同じ項目にヒットするようにする（例: 最高乳量 ← ピーク乳量）
                            alias = item_def.get("alias", "")
                            if alias:
                                display_name_to_item_key[alias] = item_key
                                display_name_lower_to_display_name[alias.lower()] = display_name
                            # 計算項目・ソース項目を識別（ソース＝乳検・ゲノム等、SQLでは取れないのでformula_engineで取得）
                            origin = (item_def.get("origin") or item_def.get("type", "")).strip()
                            if origin == "calc":
                                calc_items.add(display_name)
                                calc_item_keys.add(item_key)
                            elif origin == "source":
                                source_items.add(display_name)
                                source_item_keys.add(item_key)
                except Exception as e:
                    logging.warning(f"item_dictionary.json読み込みエラー: {e}")
            
            # 入力された項目名を正規化（項目キーを表示名に変換、大文字小文字を区別しない）
            normalized_columns = []
            for col in columns:
                col_lower = col.lower()
                # 項目キーが直接指定された場合（大文字小文字を区別しない）、表示名に変換
                if col_lower in item_key_lower_to_item_key:
                    original_item_key = item_key_lower_to_item_key[col_lower]
                    if original_item_key in item_key_to_display_name:
                        normalized_columns.append(item_key_to_display_name[original_item_key])
                    else:
                        normalized_columns.append(col)
                # 表示名が指定された場合（大文字小文字を区別しない）、元の表示名を使用
                elif col_lower in display_name_lower_to_display_name:
                    original_display_name = display_name_lower_to_display_name[col_lower]
                    normalized_columns.append(original_display_name)
                else:
                    normalized_columns.append(col)
            
            # リストコマンドの場合、常にIDを含める（IDが含まれていない場合は先頭に追加）
            id_column_names = {'ID', '個体ID', '拡大4桁ID', 'COW_ID', 'id', '個体id', '拡大4桁id', 'cow_id'}
            has_id = any(col in id_column_names or col.lower() in {n.lower() for n in id_column_names} for col in normalized_columns)
            if not has_id:
                # IDを先頭に追加
                normalized_columns.insert(0, "ID")
                # コマンド表示テキストも更新（IDを追加したことを反映）
                if original_command:
                    # コマンドにIDが含まれていない場合は追加
                    if "ID" not in original_command and "id" not in original_command.lower():
                        # 「リスト：」の後にIDを追加
                        if original_command.startswith("リスト："):
                            command = f"リスト：ID {original_command[len('リスト：'):].strip()}"
                        elif original_command.startswith("リスト:"):
                            command = f"リスト:ID {original_command[len('リスト:'):].strip()}"
                        else:
                            command = f"リスト：ID {original_command.strip()}"
                    else:
                        command = original_command
                else:
                    command = original_command
            
            # 任意月の乳検項目を検出（MILK1M, MILK2M, SCC1M, SCC2M, LS1M, LS2M, MUN1M, MUN2M, PROT1M, PROT2M, FAT1M, FAT2M, DENFA1M, DENFA2M）
            # 項目キーから判定
            monthly_milk_test_item_keys_1 = {'MILK1M', 'SCC1M', 'LS1M', 'MUN1M', 'PROT1M', 'FAT1M', 'DENFA1M'}
            monthly_milk_test_item_keys_2 = {'MILK2M', 'SCC2M', 'LS2M', 'MUN2M', 'PROT2M', 'FAT2M', 'DENFA2M'}
            
            # 指定された任意月の乳検項目を検出
            detected_monthly_items_1 = set()  # 月1設定を使用する項目
            detected_monthly_items_2 = set()  # 月2設定を使用する項目
            
            for col in normalized_columns:
                col_lower = col.lower()
                # 項目キーで検索（大文字小文字を区別しない）
                if col_lower in {k.lower() for k in monthly_milk_test_item_keys_1}:
                    detected_monthly_items_1.add(col)
                elif col_lower in {k.lower() for k in monthly_milk_test_item_keys_2}:
                    detected_monthly_items_2.add(col)
                else:
                    # 表示名で検索（「任意月の乳検」で始まり、「１」または「２」で終わる）
                    if '任意月の乳検' in col and ('１' in col or '1' in col):
                        detected_monthly_items_1.add(col)
                    elif '任意月の乳検' in col and ('２' in col or '2' in col):
                        detected_monthly_items_2.add(col)
            
            # 任意月の乳検項目が指定された場合、毎回月を聞く
            if detected_monthly_items_1 or detected_monthly_items_2:
                if not self._prompt_monthly_milk_test_months(
                    need_month_1=bool(detected_monthly_items_1),
                    need_month_2=bool(detected_monthly_items_2)
                ):
                        self.add_message(role="system", text="月の指定がキャンセルされました。")
                        return
            
            # 計算項目・ソース項目（乳検・ゲノム等）とSQLで取得できる項目を分離（大文字小文字を区別しない）
            # ソース項目もformula_engineで取得するため calc_columns に含める
            sql_columns = []
            calc_columns = []
            calc_items_lower = {item.lower() for item in calc_items}
            source_items_lower = {item.lower() for item in source_items}
            for col in normalized_columns:
                col_lower = col.lower()
                if col_lower in calc_items_lower:
                    for calc_item in calc_items:
                        if calc_item.lower() == col_lower:
                            calc_columns.append(calc_item)
                            break
                elif col_lower in source_items_lower:
                    for src_item in source_items:
                        if src_item.lower() == col_lower:
                            calc_columns.append(src_item)
                            break
                else:
                    sql_columns.append(col)
            
            # 条件をSQL条件とフィルタリング条件に分離
            sql_conditions = {}  # データベースカラムの条件（SQLに含める）
            filter_conditions_for_calc = []  # 計算項目の条件（後でフィルタリング）
            
            if filter_conditions:
                for filter_condition in filter_conditions:
                    item_name = filter_condition['item_name']
                    operator = filter_condition['operator']
                    value_str = filter_condition['value']
                    
                    # 条件項目名を正規化
                    item_name_lower = item_name.lower()
                    normalized_item_name = None
                    item_key = None
                    
                    # 表示名で検索
                    if item_name_lower in display_name_lower_to_display_name:
                        normalized_item_name = display_name_lower_to_display_name[item_name_lower]
                        item_key = display_name_to_item_key.get(normalized_item_name)
                    # 項目キーで検索
                    elif item_name_lower in item_key_lower_to_item_key:
                        item_key = item_key_lower_to_item_key[item_name_lower]
                        if item_key in item_key_to_display_name:
                            normalized_item_name = item_key_to_display_name[item_key]
                    
                    # 計算項目かどうかを判定
                    is_calc_item = False
                    if normalized_item_name and normalized_item_name.lower() in calc_items_lower:
                        is_calc_item = True
                    elif item_key and item_key in calc_item_keys:
                        is_calc_item = True
                    
                    if is_calc_item:
                        # 計算項目の条件は後でフィルタリング
                        filter_conditions_for_calc.append(filter_condition)
                    else:
                        # データベースカラムの条件はSQLに含める
                        # 範囲指定のチェック（例：DIM=100-150）
                        if filter_condition.get('is_range', False):
                            # 範囲指定の場合は、>=min AND <=max の2つの条件として扱う
                            # ただし、_generate_list_sqlは1つの条件しか処理できないため、
                            # 範囲指定の場合はフィルタリング条件として扱う
                            filter_conditions_for_calc.append(filter_condition)
                        elif isinstance(value_str, list):
                            # 複数値指定（例：RC=4,5）の場合は、IN句として扱う
                            # ただし、現在の_generate_list_sqlはIN句に対応していないため、
                            # 一旦フィルタリング条件として扱う
                            filter_conditions_for_calc.append(filter_condition)
                        else:
                            # 単一値の場合
                            condition_value = operator + str(value_str)
                            # 正規化された項目名または元の項目名を使用
                            condition_key = normalized_item_name if normalized_item_name else item_name
                            sql_conditions[condition_key] = condition_value
            
            # SQLを生成（SQL条件を含める）
            sql = None
            rows = []
            if sql_columns:
                sql = self._generate_list_sql(sql_columns, sql_conditions)
                if sql:
                    try:
                        # SQLを実行
                        conn = self.db.connect()
                        cursor = conn.cursor()
                        cursor.execute(sql)
                        rows = cursor.fetchall()
                    except Exception as e:
                        logging.error(f"SQL実行エラー: {e}", exc_info=True)
                        self.add_message(role="system", text=f"データの取得中にエラーが発生しました: {e}")
                        # エラーが発生しても処理を続行（結果は空になる）
                        rows = []
                else:
                    # SQL生成に失敗した場合でも、計算項目があれば処理を続行
                    if not calc_columns:
                        self.add_message(role="system", text=f"指定された項目が認識できませんでした: {list_text}")
                        return
            elif calc_columns:
                # 計算項目のみの場合、全牛を取得（IDも含める）
                # COW_IDが含まれていない場合は自動的に追加
                if "ID" not in sql_columns and "個体ID" not in sql_columns and "拡大4桁ID" not in sql_columns and "COW_ID" not in sql_columns:
                    sql_columns.append("ID")  # IDを追加してSQLを生成
                    sql = self._generate_list_sql(sql_columns, conditions)
                    if sql:
                        conn = self.db.connect()
                        cursor = conn.cursor()
                        cursor.execute(sql)
                        rows = cursor.fetchall()
                    else:
                        # SQL生成に失敗した場合は、シンプルなSQLを実行
                        conn = self.db.connect()
                        cursor = conn.cursor()
                        cursor.execute("""
                            SELECT c.auto_id AS auto_id, c.cow_id AS ID
                            FROM cow c
                            ORDER BY c.cow_id
                        """)
                        rows = cursor.fetchall()
                else:
                    # IDが既に含まれている場合
                    conn = self.db.connect()
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT c.auto_id AS auto_id, c.cow_id AS ID
                        FROM cow c
                        ORDER BY c.cow_id
                    """)
                    rows = cursor.fetchall()
            
            # 現存牛のみをフィルタリング（除籍牛を除外）
            rows = self._filter_existing_cows(rows)
            
            # 計算項目がある場合、各牛について計算
            if calc_columns and rows:
                # 受胎授精種類のコード→表示名変換用（リストに含まれる場合のみ読み込み）
                insemination_type_names = {}
                if "受胎授精種類" in calc_columns and hasattr(self, 'farm_path') and self.farm_path:
                    try:
                        settings_path = self.farm_path / "insemination_settings.json"
                        if settings_path.exists():
                            with open(settings_path, 'r', encoding='utf-8') as f:
                                insemination_type_names = json.load(f).get('insemination_types', {})
                    except Exception:
                        pass
                # 各行について計算項目を追加
                enhanced_rows = []
                for row in rows:
                    # rowを辞書に変換
                    row_dict = dict(row)
                    
                    # SQLで取得した繁殖コード／繁殖区分もフォーマット
                    if "繁殖コード" in row_dict:
                        row_dict["繁殖コード"] = self.format_rc(row_dict.get("繁殖コード"))
                    if "繁殖区分" in row_dict:
                        row_dict["繁殖区分"] = self.format_rc(row_dict.get("繁殖区分"))
                    
                    # cow_auto_idを取得
                    cow_auto_id = row_dict.get("auto_id")
                    if not cow_auto_id:
                        cow_id = row_dict.get("ID") or row_dict.get("個体ID")
                        if cow_id:
                            # cow_idからauto_idを取得
                            cow = self.db.get_cow_by_id(cow_id)
                            if cow:
                                cow_auto_id = cow.get('auto_id')
                    
                    if cow_auto_id:
                        # FormulaEngineで計算項目を計算
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        
                        # 分娩予定〇日前：分娩予定日から算出するフォールバック用マッピング
                        duem_display_to_days = {
                            "分娩予定60日前": 60, "分娩予定40日前": 40, "分娩予定30日前": 30,
                            "分娩予定21日前": 21, "分娩予定14日前": 14, "分娩予定7日前": 7,
                        }
                        # 計算項目を追加
                        for col in calc_columns:
                            item_key = display_name_to_item_key.get(col)
                            if item_key and item_key in calculated:
                                val = calculated[item_key]
                                # 繁殖コードの場合はフォーマット
                                if col == "繁殖コード" or col == "繁殖区分" or item_key == "RC":
                                    val = self.format_rc(val)
                                # 分娩予定〇日前がNoneで分娩予定日がある場合は分娩予定日から算出
                                if val is None and col in duem_display_to_days:
                                    due_val = calculated.get("DUE") or row_dict.get("分娩予定日")
                                    if due_val and isinstance(due_val, str):
                                        try:
                                            from datetime import datetime, timedelta
                                            due_dt = datetime.strptime(due_val.strip()[:10], "%Y-%m-%d")
                                            target_dt = due_dt - timedelta(days=duem_display_to_days[col])
                                            val = target_dt.strftime("%Y-%m-%d")
                                        except (ValueError, TypeError):
                                            pass
                                # 受胎授精種類：「自然発情」「WPG」など表示名のみ表示（コード:名称 形式は使わない）
                                if col == "受胎授精種類" and val and isinstance(val, str):
                                    if "：" in val:
                                        row_dict[col] = val.split("：", 1)[1].strip()
                                    elif ":" in val:
                                        row_dict[col] = val.split(":", 1)[1].strip()
                                    elif insemination_type_names:
                                        code = val.strip()
                                        row_dict[col] = insemination_type_names.get(str(code), val)
                                    else:
                                        row_dict[col] = val
                                else:
                                    row_dict[col] = val
                            else:
                                row_dict[col] = None
                                # 同上：分娩予定〇日前のフォールバック
                                if col in duem_display_to_days:
                                    due_val = calculated.get("DUE") if (calculated and cow_auto_id) else None
                                    if not due_val and row_dict.get("分娩予定日"):
                                        due_val = row_dict.get("分娩予定日")
                                    if due_val and isinstance(due_val, str):
                                        try:
                                            from datetime import datetime, timedelta
                                            due_dt = datetime.strptime(due_val.strip()[:10], "%Y-%m-%d")
                                            target_dt = due_dt - timedelta(days=duem_display_to_days[col])
                                            row_dict[col] = target_dt.strftime("%Y-%m-%d")
                                        except (ValueError, TypeError):
                                            pass
                    
                    enhanced_rows.append(row_dict)
                rows = enhanced_rows
            else:
                # SQLで取得したデータのみの場合も、繁殖コードをフォーマット
                enhanced_rows = []
                for row in rows:
                    row_dict = dict(row)
                    # SQLで取得した繁殖コード／繁殖区分もフォーマット
                    if "繁殖コード" in row_dict:
                        row_dict["繁殖コード"] = self.format_rc(row_dict.get("繁殖コード"))
                    if "繁殖区分" in row_dict:
                        row_dict["繁殖区分"] = self.format_rc(row_dict.get("繁殖区分"))
                    enhanced_rows.append(row_dict)
                rows = enhanced_rows
            
            # 条件でフィルタリング（計算項目の条件のみ、複数条件に対応：すべての条件を満たす行のみを残す）
            if filter_conditions_for_calc and rows:
                filtered_rows = []
                
                # 各行をフィルタリング
                for row_dict in rows:
                    match_all_conditions = True
                    
                    # すべての条件をチェック（計算項目の条件のみ）
                    for filter_condition in filter_conditions_for_calc:
                        item_name = filter_condition['item_name']
                        operator = filter_condition['operator']
                        value_str = filter_condition['value']
                        
                        # 条件項目名を正規化
                        item_name_lower = item_name.lower()
                        normalized_item_name = None
                        item_key = None
                        
                        # 表示名で検索
                        if item_name_lower in display_name_lower_to_display_name:
                            normalized_item_name = display_name_lower_to_display_name[item_name_lower]
                            item_key = display_name_to_item_key.get(normalized_item_name)
                        # 項目キーで検索
                        elif item_name_lower in item_key_lower_to_item_key:
                            item_key = item_key_lower_to_item_key[item_name_lower]
                            if item_key in item_key_to_display_name:
                                normalized_item_name = item_key_to_display_name[item_key]
                        
                        # 行から値を取得
                        row_value = None
                        is_rc_condition = (item_key == "RC" or normalized_item_name in ("繁殖コード", "繁殖区分"))
                        
                        # RCの場合は、フォーマット前の数値を取得する必要がある
                        if is_rc_condition:
                            cow_auto_id = row_dict.get("auto_id")
                            if cow_auto_id:
                                calculated = self.formula_engine.calculate(cow_auto_id)
                                if "RC" in calculated:
                                    row_value = calculated["RC"]  # フォーマット前の数値
                        else:
                            # 1. row_dictから値を取得（表示名で検索）
                            if normalized_item_name:
                                for key in row_dict.keys():
                                    if key.lower() == normalized_item_name.lower():
                                        row_value = row_dict[key]
                                        break
                            
                            # 2. row_dictから値を取得（項目キーで検索）
                            if row_value is None and item_key:
                                if item_key in row_dict:
                                    row_value = row_dict[item_key]
                            
                            # 3. 計算項目がまだ計算されていない場合は計算
                            if row_value is None:
                                cow_auto_id = row_dict.get("auto_id")
                                if cow_auto_id:
                                    calculated = self.formula_engine.calculate(cow_auto_id)
                                    if item_key and item_key in calculated:
                                        row_value = calculated[item_key]
                        
                        # 条件を評価
                        if row_value is None:
                            match_all_conditions = False
                            break  # 値が取得できない場合は条件を満たさない
                        
                        # RCの場合は、フォーマットされた文字列から数値を抽出する必要がある場合がある
                        if is_rc_condition and isinstance(row_value, str):
                            # フォーマットされた文字列（例："4: Dry（乾乳中）"）から数値を抽出
                            import re
                            match = re.match(r'^(\d+)', str(row_value))
                            if match:
                                row_value = int(match.group(1))
                        
                        # 範囲指定のチェック（例：DIM=100-150）
                        if filter_condition.get('is_range', False):
                            # 範囲指定の場合、>=min AND <=max の両方を満たす必要がある
                            try:
                                row_value_num = float(row_value)
                                min_value = float(value_str)  # 最小値
                                max_value = float(filter_condition.get('range_max', value_str))  # 最大値
                                condition_match = (row_value_num >= min_value and row_value_num <= max_value)
                                
                                if not condition_match:
                                    match_all_conditions = False
                                    break  # 1つの条件でも満たさない場合はスキップ
                            except (ValueError, TypeError):
                                # 数値変換に失敗した場合は条件を満たさない
                                match_all_conditions = False
                                break
                        # 値がリストの場合（複数値指定、例：RC=4,5）
                        elif isinstance(value_str, list):
                            # 複数値指定の場合、いずれかの値に等しいかチェック
                            condition_match = False
                            try:
                                row_value_num = float(row_value)
                                for val in value_str:
                                    try:
                                        condition_value_num = float(val)
                                        if row_value_num == condition_value_num:
                                            condition_match = True
                                            break
                                    except (ValueError, TypeError):
                                        # 数値変換に失敗した場合は文字列比較
                                        if str(row_value) == str(val):
                                            condition_match = True
                                            break
                            except (ValueError, TypeError):
                                # row_valueが数値でない場合は文字列比較
                                for val in value_str:
                                    if str(row_value) == str(val):
                                        condition_match = True
                                        break
                            
                            if not condition_match:
                                match_all_conditions = False
                                break  # 1つの条件でも満たさない場合はスキップ
                        else:
                            # 単一値の場合
                            try:
                                row_value_num = float(row_value)
                                condition_value_num = float(value_str)
                                
                                # 比較演算子で評価
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
                                    match_all_conditions = False
                                    break  # 1つの条件でも満たさない場合はスキップ
                            except (ValueError, TypeError):
                                # 数値変換に失敗した場合は文字列比較（＝完全一致、>=部分一致）
                                condition_match = False
                                if operator == "=" or operator == "==":
                                    condition_match = (str(row_value).strip() == str(value_str).strip())
                                elif operator == "!=" or operator == "<>":
                                    condition_match = (str(row_value).strip() != str(value_str).strip())
                                elif operator == ">=":
                                    # 文字列の部分一致（値が項目の文字列に含まれる）
                                    condition_match = (str(value_str).strip() in str(row_value))
                                
                                if not condition_match:
                                    match_all_conditions = False
                                    break  # 1つの条件でも満たさない場合はスキップ
                    
                    # すべての条件を満たす場合のみ追加
                    if match_all_conditions:
                        filtered_rows.append(row_dict)
                
                rows = filtered_rows
            elif or_groups and len(or_groups) > 1 and rows:
                # 複数 OR グループ：いずれか 1 グループの条件をすべて満たす行を残す
                filtered_rows = []
                for row_dict in rows:
                    if any(
                        self._list_row_matches_condition_group(
                            row_dict, g,
                            display_name_to_item_key, item_key_to_display_name,
                            item_key_lower_to_item_key, display_name_lower_to_display_name,
                            calc_item_keys
                        )
                        for g in or_groups
                    ):
                        filtered_rows.append(row_dict)
                rows = filtered_rows
            
            # 並び順を適用（指定がなければID昇順）
            if rows:
                sort_key_col = sort_column_display_name if sort_column_display_name else "ID"
                # 行のキーは表示名なので、項目キー（dim等）の場合は表示名に解決する
                if sort_key_col and sort_key_col != "ID":
                    sort_lower = sort_key_col.strip().lower()
                    if sort_lower in item_key_lower_to_item_key:
                        item_key = item_key_lower_to_item_key[sort_lower]
                        if item_key in item_key_to_display_name:
                            sort_key_col = item_key_to_display_name[item_key]
                    elif sort_lower in display_name_lower_to_display_name:
                        sort_key_col = display_name_lower_to_display_name[sort_lower]
                    elif (normalized_columns or []) and sort_key_col not in normalized_columns:
                        for col in normalized_columns:
                            if col and (col.lower() == sort_lower or (sort_column_display_name and sort_column_display_name in col)):
                                sort_key_col = col
                                break
                def _list_sort_key(row):
                    r = row if isinstance(row, dict) else dict(row)
                    val = r.get(sort_key_col)
                    # 数値でソートできる場合は数値として比較（降順で正しく並ぶように）
                    if val is None or val == "":
                        return (1, 0)
                    try:
                        n = float(val)
                        return (0, n)
                    except (TypeError, ValueError):
                        return (0, str(val))
                try:
                    rows = sorted(rows, key=_list_sort_key, reverse=not sort_order_asc)
                except Exception:
                    pass
            
            # 結果を表タブに表示（auto_idカラムは除外）
            if rows:
                # auto_idカラムを除外して表示（正規化された列名を使用）
                display_columns = [col for col in normalized_columns if col != "auto_id"]
                try:
                    # コマンド情報を渡す
                    self._display_list_result_in_table(display_columns, rows, command=command)
                    # 表タブを選択
                    self.result_notebook.select(0)  # 0番目のタブ（表タブ）
                except Exception as e:
                    logging.error(f"結果表示エラー: {e}", exc_info=True)
                    self.add_message(role="system", text=f"結果の表示中にエラーが発生しました: {e}")
            else:
                self.add_message(role="system", text="該当データがありません。")
                
        except Exception as e:
            logging.error(f"リストコマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"リストコマンドの実行中にエラーが発生しました: {e}")
            # エラーが発生しても処理を続行できるように、結果表示エリアをクリア
            try:
                self._clear_result_display()
            except Exception:
                pass  # クリア処理でエラーが発生しても無視
    
    def _execute_event_list_command(self, command: str):
        """
        イベントリストコマンドを実行
        
        Args:
            command: イベントリストコマンド（例：「イベントリスト：」「イベントリスト：0980」）
        """
        logging.info(f"イベントリストコマンド実行開始: command='{command}'")
        try:
            # 「イベントリスト：」または「イベントリスト:」を除去
            if command.startswith("イベントリスト："):
                event_list_text = command[len("イベントリスト："):].strip()
            elif command.startswith("イベントリスト:"):
                event_list_text = command[len("イベントリスト:"):].strip()
            else:
                event_list_text = command.strip()
            
            # コマンド文字列から条件部分を除去（「：」で区切られている場合）
            # 例：「all：WPG」→「all」、条件は条件入力欄から取得
            if "：" in event_list_text:
                parts = event_list_text.split("：", 1)
                event_list_text = parts[0].strip()
            elif ":" in event_list_text:
                parts = event_list_text.split(":", 1)
                event_list_text = parts[0].strip()
            
            logging.info(f"イベントリスト: event_list_text='{event_list_text}' (条件部分は除去済み)")
            
            # 期間設定を取得
            start_date = None
            end_date = None
            if hasattr(self, 'start_date_entry') and self.start_date_entry.winfo_exists():
                start_date_str = self.start_date_entry.get().strip()
                if start_date_str and start_date_str != self.start_date_placeholder:
                    try:
                        datetime.strptime(start_date_str, '%Y-%m-%d')
                        start_date = start_date_str
                    except ValueError:
                        pass
            
            if hasattr(self, 'end_date_entry') and self.end_date_entry.winfo_exists():
                end_date_str = self.end_date_entry.get().strip()
                if end_date_str and end_date_str != self.end_date_placeholder:
                    try:
                        datetime.strptime(end_date_str, '%Y-%m-%d')
                        end_date = end_date_str
                    except ValueError:
                        pass
            
            # 期間設定が空欄の場合は直近一年間
            if not start_date or not end_date:
                today = datetime.now()
                end_date = today.strftime('%Y-%m-%d')
                one_year_ago = today - timedelta(days=365)
                start_date = one_year_ago.strftime('%Y-%m-%d')
            
            # 条件設定は使用しない（イベントリストでは条件欄を削除）
            condition_text = ""
            logging.info(f"イベントリスト: 条件フィルタリングは使用しません")
            
            # 個体IDを取得（番号が入力されている場合）
            cow_id = None
            if event_list_text and event_list_text.isdigit():
                cow_id = event_list_text.zfill(4)
            
            # イベントを取得
            if cow_id:
                events = self.db.get_events_by_cow_id_and_period(cow_id, start_date, end_date)
                logging.info(f"イベントリスト: 個体ID={cow_id}, 期間={start_date}～{end_date}, 取得件数={len(events)}")
            else:
                events = self.db.get_events_by_period(start_date, end_date)
                logging.info(f"イベントリスト: 全個体, 期間={start_date}～{end_date}, 取得件数={len(events)}")
            
            # 選択されたイベントでフィルタリング
            if hasattr(self, 'selected_event_numbers') and self.selected_event_numbers:
                filtered_events = []
                for event in events:
                    event_number = event.get('event_number')
                    if self._event_number_in_selection(event_number):
                        filtered_events.append(event)
                events = filtered_events
                logging.info(f"イベントリスト: イベント選択フィルタリング後件数={len(events)}")
            
            # 部分一致でフィルタリング（NOTE内で部分一致）
            if hasattr(self, 'note_partial_match_entry') and self.note_partial_match_entry.winfo_exists():
                partial_match_text = self.note_partial_match_entry.get().strip()
                if partial_match_text:
                    filtered_events = []
                    partial_match_lower = partial_match_text.lower()
                    
                    for event in events:
                        # 表示用のNOTE（NOTE + json_data）を作成（実際に表示される内容と同じ）
                        note = event.get('note', '') or ''
                        json_data = event.get('json_data', {})
                        
                        # 表示用のNOTEを作成（表示ロジックと同じ）
                        display_note = note
                        if json_data:
                            if isinstance(json_data, dict):
                                json_display = json.dumps(json_data, ensure_ascii=False)
                            else:
                                json_display = str(json_data)
                            if note:
                                display_note = f"{note} | {json_display}"
                            else:
                                display_note = json_display
                        
                        # 表示用NOTEに部分一致文字列が含まれているかチェック（大文字小文字を区別しない）
                        display_note_lower = display_note.lower()
                        if partial_match_lower in display_note_lower:
                            filtered_events.append(event)
                    
                    events = filtered_events
                    logging.info(f"イベントリスト: 部分一致フィルタリング後件数={len(events)}, 検索文字列='{partial_match_text}'")
            
            # 妊娠鑑定プラス等は同一個体×産次で初回のみ表示
            events = self._filter_events_unique_per_cow_per_lact(events)
            logging.info(f"イベントリスト: ユニーク化後件数={len(events)}")
            
            # イベント辞書を読み込んでイベントタイプ名を取得
            event_dict_path = Path(__file__).parent.parent.parent / "config_default" / "event_dictionary.json"
            event_dict = {}
            if event_dict_path.exists():
                with open(event_dict_path, 'r', encoding='utf-8') as f:
                    event_dict_data = json.load(f)
                    for event_num_str, event_data in event_dict_data.items():
                        event_dict[int(event_num_str)] = event_data.get('name_jp', event_data.get('alias', str(event_num_str)))
            
            # 結果を表形式で表示
            logging.info(f"イベントリスト: フィルタリング後件数={len(events)}")
            if events:
                # 列を定義
                columns = ['ID', '日付', 'DIM', 'イベントタイプ', 'NOTE']
                
                # データを準備
                rows = []
                for event in events:
                    cow_auto_id = event.get('cow_auto_id')
                    cow = None
                    cow_id_display = ""
                    if cow_auto_id:
                        cow = self.db.get_cow_by_auto_id(cow_auto_id)
                        if cow:
                            cow_id_display = cow.get('cow_id', '')
                    
                    event_date = event.get('event_date', '')
                    event_number = event.get('event_number')
                    event_type = event_dict.get(event_number, f"イベント{event_number}")
                    note = event.get('note', '') or ''
                    
                    # json_dataの全情報をNOTEに追加（NOTEがある場合は結合）
                    json_data = event.get('json_data', {})
                    if json_data:
                        json_str = json.dumps(json_data, ensure_ascii=False)
                        if note:
                            note = f"{note} | {json_str}"
                        else:
                            note = json_str
                    
                    # DIMを計算（イベント日付時点での分娩後日数）
                    # 重要：イベント日付時点での「現在の産次」の分娩日を基準にする
                    # 例：12月30日のイベントに対しては、その時点での最新の分娩日（2月1日）を基準にする
                    # その後の分娩（1月1日）は考慮しない
                    dim_display = ""
                    if cow_auto_id and event_date:
                        try:
                            event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                            
                            # その牛のすべてのイベントを取得
                            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                            
                            # イベント日付以前の分娩イベント（event_number=202）を抽出
                            calv_events_before = []
                            for ev in all_events:
                                ev_date_str = ev.get('event_date', '')
                                ev_number = ev.get('event_number')
                                if ev_number == 202 and ev_date_str:  # CALVイベント
                                    try:
                                        ev_dt = datetime.strptime(ev_date_str, '%Y-%m-%d')
                                        if ev_dt <= event_dt:  # イベント日付以前の分娩のみ
                                            calv_events_before.append((ev_dt, ev_date_str))
                                    except ValueError:
                                        continue
                            
                            # イベント日付以前の分娩イベントを日付の降順でソート（最新が最初）
                            if calv_events_before:
                                calv_events_before.sort(reverse=True)  # 日付の降順
                                # 最新の分娩日を取得（イベント日付時点での「現在の産次」の分娩日）
                                calv_date = calv_events_before[0][1]
                                
                                # DIMを計算
                                calv_dt = datetime.strptime(calv_date, '%Y-%m-%d')
                                dim = (event_dt - calv_dt).days
                                if dim >= 0:
                                    dim_display = str(dim)
                        except Exception as e:
                            logging.debug(f"DIM計算エラー (cow_auto_id={cow_auto_id}, event_date={event_date}): {e}")
                            dim_display = ""
                    
                    rows.append([cow_id_display, event_date, dim_display, event_type, note])
                
                logging.info(f"イベントリスト: 表示行数={len(rows)}")
                # 期間情報を文字列化
                period_text = f"{start_date} ～ {end_date}"
                # 表タブに表示
                self._display_list_result_in_table(columns, rows, command=command, period=period_text)
                self.result_notebook.select(0)  # 表タブを選択
            else:
                self.add_message(role="system", text="該当するイベントがありません。")
                
        except Exception as e:
            logging.error(f"イベントリストコマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"イベントリストコマンドの実行中にエラーが発生しました: {e}")
            try:
                self._clear_result_display()
            except Exception:
                pass
    
    # 同一産次で1頭あたり1回とカウントするイベント番号（妊娠鑑定プラス等：頭数ベースの集計）
    EVENT_NUMBERS_UNIQUE_PER_COW_PER_LACT = {303, 304}  # 妊娠鑑定プラス、妊娠鑑定プラス（直近以外）

    def _event_number_in_selection(self, event_number: Any) -> bool:
        """イベントリスト／集計のチェックボックス選択と DB の event_number を照合（int/str の混在を吸収）。"""
        if not hasattr(self, "selected_event_numbers") or not self.selected_event_numbers:
            return True
        try:
            n = int(event_number)
        except (TypeError, ValueError):
            n = event_number
        for x in self.selected_event_numbers:
            try:
                if int(x) == n:
                    return True
            except (TypeError, ValueError):
                if x == n:
                    return True
        return False

    def _filter_events_unique_per_cow_per_lact(self, events: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        妊娠鑑定プラス等のイベントを「個体×産次」単位で初回のみ残す
        """
        if not events or not self.EVENT_NUMBERS_UNIQUE_PER_COW_PER_LACT:
            return events
        
        earliest_event_tokens: Dict[Tuple[Any, Any], Tuple[Tuple[Any, ...], str]] = {}
        for idx, event in enumerate(events):
            event_number = event.get('event_number')
            if event_number not in self.EVENT_NUMBERS_UNIQUE_PER_COW_PER_LACT:
                continue
            
            event_date_str = event.get('event_date')
            if not event_date_str:
                continue
            try:
                event_dt = datetime.strptime(event_date_str, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            lact = self._get_event_lact(event, event_date_str)
            if lact is None:
                continue
            
            key = (cow_auto_id, lact)
            tie_breaker = (
                event_dt,
                event_number if event_number is not None else 0,
                idx,
            )
            token = self._build_event_token(event, idx)
            if key not in earliest_event_tokens or tie_breaker < earliest_event_tokens[key][0]:
                earliest_event_tokens[key] = (tie_breaker, token)
        
        if not earliest_event_tokens:
            return events
        
        tokens_to_keep = {info[1] for info in earliest_event_tokens.values()}
        deduped_events: List[Dict[str, Any]] = []
        for idx, event in enumerate(events):
            event_number = event.get('event_number')
            if event_number in self.EVENT_NUMBERS_UNIQUE_PER_COW_PER_LACT:
                token = self._build_event_token(event, idx)
                if token not in tokens_to_keep:
                    continue
            deduped_events.append(event)
        
        return deduped_events
    
    def _build_event_token(self, event: Dict[str, Any], fallback_index: int) -> str:
        """
        イベントを一意に識別するためのトークンを生成
        """
        event_id = event.get('id')
        if event_id is not None:
            return f"id:{event_id}"
        cow_auto_id = event.get('cow_auto_id', '')
        event_number = event.get('event_number', '')
        event_date = event.get('event_date', '')
        return f"idx:{fallback_index}:{cow_auto_id}:{event_number}:{event_date}"
    
    def _execute_event_aggregate_command(self, command: str):
        """
        イベント集計コマンドを実行
        
        Args:
            command: イベント集計コマンド（例：「イベント集計：all：月」）
        """
        logging.info(f"イベント集計コマンド実行開始: command='{command}'")
        try:
            # 「イベント集計：」を除去
            if command.startswith("イベント集計："):
                command_body = command[len("イベント集計："):].strip()
            elif command.startswith("イベント集計:"):
                command_body = command[len("イベント集計:"):].strip()
            else:
                command_body = command.strip()
            
            # コマンドから分類を抽出
            classification = ""
            if "：" in command_body:
                parts = command_body.split("：", 1)
                command_body = parts[0].strip()
                if len(parts) > 1:
                    classification = parts[1].strip()
            elif ":" in command_body:
                parts = command_body.split(":", 1)
                command_body = parts[0].strip()
                if len(parts) > 1:
                    classification = parts[1].strip()
            
            # 分類が指定されていない場合はエラー
            if not classification:
                self.add_message(role="system", text="分類を指定してください。")
                return
            
            logging.info(f"イベント集計: command_body='{command_body}', classification='{classification}'")
            
            # 期間設定を取得
            start_date = None
            end_date = None
            if hasattr(self, 'start_date_entry') and self.start_date_entry.winfo_exists():
                start_date_str = self.start_date_entry.get().strip()
                if start_date_str and start_date_str != self.start_date_placeholder:
                    try:
                        datetime.strptime(start_date_str, '%Y-%m-%d')
                        start_date = start_date_str
                    except ValueError:
                        pass
            
            if hasattr(self, 'end_date_entry') and self.end_date_entry.winfo_exists():
                end_date_str = self.end_date_entry.get().strip()
                if end_date_str and end_date_str != self.end_date_placeholder:
                    try:
                        datetime.strptime(end_date_str, '%Y-%m-%d')
                        end_date = end_date_str
                    except ValueError:
                        pass
            
            # 期間設定が空欄の場合は直近一年間
            if not start_date or not end_date:
                today = datetime.now()
                end_date = today.strftime('%Y-%m-%d')
                one_year_ago = today - timedelta(days=365)
                start_date = one_year_ago.strftime('%Y-%m-%d')
            
            # イベントを取得
            events = self.db.get_events_by_period(start_date, end_date)
            logging.info(f"イベント集計: 期間={start_date}～{end_date}, 取得件数={len(events)}")
            
            # 選択されたイベントでフィルタリング
            if hasattr(self, 'selected_event_numbers') and self.selected_event_numbers:
                filtered_events = []
                for event in events:
                    event_number = event.get('event_number')
                    if self._event_number_in_selection(event_number):
                        filtered_events.append(event)
                events = filtered_events
                logging.info(f"イベント集計: イベント選択フィルタリング後件数={len(events)}")
            else:
                # イベントが選択されていない場合は全イベント
                logging.info(f"イベント集計: イベントが選択されていないため全イベントを使用")
            
            # 妊娠鑑定プラス等は同一個体×産次で初回のみカウント対象にする
            events = self._filter_events_unique_per_cow_per_lact(events)
            logging.info(f"イベント集計: ユニーク化後件数={len(events)}")
            
            # イベントが一つもない場合は表示しない
            if not events:
                self.add_message(role="system", text="該当するイベントがありません。")
                return
            
            # イベント辞書を読み込んでイベントタイプ名を取得
            event_dict_path = Path(__file__).parent.parent.parent / "config_default" / "event_dictionary.json"
            event_dict = {}
            if event_dict_path.exists():
                with open(event_dict_path, 'r', encoding='utf-8') as f:
                    event_dict_data = json.load(f)
                    for event_num_str, event_data in event_dict_data.items():
                        event_dict[int(event_num_str)] = event_data.get('name_jp', event_data.get('alias', str(event_num_str)))
            
            # イベント×分類で集計
            # 集計結果: {event_number: {classification_value: count}}
            aggregate_result = {}
            # 同一産次1頭1回用: {event_number: {classification_value: set((cow_auto_id, lact), ...)}}
            unique_per_cow_per_lact = {}
            
            for event in events:
                event_number = event.get('event_number')
                event_date = event.get('event_date', '')
                
                if not event_date:
                    continue
                
                # 分類値を取得
                classification_value = self._get_event_classification_value(event, classification, event_date)
                if classification_value is None:
                    continue
                
                # 妊娠鑑定プラス等：同一産次で1頭1回カウント（頭数ベース）
                if event_number in self.EVENT_NUMBERS_UNIQUE_PER_COW_PER_LACT:
                    lact = self._get_event_lact(event, event_date)
                    if lact is None:
                        continue
                    cow_auto_id = event.get('cow_auto_id')
                    if not cow_auto_id:
                        continue
                    if event_number not in unique_per_cow_per_lact:
                        unique_per_cow_per_lact[event_number] = {}
                    if classification_value not in unique_per_cow_per_lact[event_number]:
                        unique_per_cow_per_lact[event_number][classification_value] = set()
                    unique_per_cow_per_lact[event_number][classification_value].add((cow_auto_id, lact))
                    continue
                
                # 通常の発生回数集計
                if event_number not in aggregate_result:
                    aggregate_result[event_number] = {}
                if classification_value not in aggregate_result[event_number]:
                    aggregate_result[event_number][classification_value] = 0
                aggregate_result[event_number][classification_value] += 1
            
            # ユニーク集計を件数に変換して aggregate_result にマージ（表示・メタデータで同じ構造を使う）
            for event_number, by_cv in unique_per_cow_per_lact.items():
                if event_number not in aggregate_result:
                    aggregate_result[event_number] = {}
                for classification_value, cow_lact_set in by_cv.items():
                    aggregate_result[event_number][classification_value] = len(cow_lact_set)
            
            # 集計結果が空の場合は表示しない
            if not aggregate_result:
                self.add_message(role="system", text="該当するイベントがありません。")
                return
            
            # 結果を表形式で表示
            # 列を定義: 合計、分類値1、分類値2、...
            # すべての分類値を収集
            all_classification_values = set()
            for event_data in aggregate_result.values():
                all_classification_values.update(event_data.keys())
            
            # 分類値をソート
            sorted_classification_values = sorted(all_classification_values, key=lambda x: self._sort_classification_value(x, classification))
            
            # 列を定義
            columns = ['イベント'] + ['合計'] + sorted_classification_values
            
            # データを準備
            rows = []
            for event_number in sorted(aggregate_result.keys()):
                event_name = event_dict.get(event_number, f"イベント{event_number}")
                event_data = aggregate_result[event_number]
                
                # 合計を計算
                total = sum(event_data.values())
                
                # 行データを作成
                row = [event_name, total]
                for classification_value in sorted_classification_values:
                    count = event_data.get(classification_value, 0)
                    row.append(count)
                
                rows.append(row)
            
            logging.info(f"イベント集計: 表示行数={len(rows)}")
            # 期間情報を文字列化
            period_text = f"{start_date} ～ {end_date}"
            
            # イベント集計のメタデータを保存（セルダブルクリック時に使用）
            self.event_aggregate_metadata = {
                'start_date': start_date,
                'end_date': end_date,
                'classification': classification,
                'selected_event_numbers': list(self.selected_event_numbers) if hasattr(self, 'selected_event_numbers') and self.selected_event_numbers else None,
                'event_dict': event_dict,
                'aggregate_result': aggregate_result,
                'sorted_classification_values': sorted_classification_values
            }
            
            # 表タブに表示
            self._display_list_result_in_table(columns, rows, command=command, period=period_text)
            self.result_notebook.select(0)  # 表タブを選択
                
        except Exception as e:
            logging.error(f"イベント集計コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"イベント集計コマンドの実行中にエラーが発生しました: {e}")
            try:
                self._clear_result_display()
            except Exception:
                pass
    
    def _get_event_lact(self, event: Dict[str, Any], event_date: str) -> Optional[int]:
        """
        イベント日付時点での産次（LACT）を取得。同一産次1頭1回カウント用。
        
        Args:
            event: イベントデータ（cow_auto_id を含むこと）
            event_date: イベント日付
        
        Returns:
            産次（整数）、取得できない場合はNone
        """
        cow_auto_id = event.get('cow_auto_id')
        if not cow_auto_id or not self.rule_engine:
            return None
        try:
            state = self.rule_engine.apply_events_until_date(cow_auto_id, event_date)
            lact = state.get('lact', 0)
            return int(lact) if lact else None
        except Exception as e:
            logging.debug(f"産次取得エラー cow_auto_id={cow_auto_id} event_date={event_date}: {e}")
            return None

    def _get_event_classification_value(self, event: Dict[str, Any], classification: str, event_date: str) -> Optional[str]:
        """
        イベントの分類値を取得
        
        Args:
            event: イベントデータ
            classification: 分類名（例：「月」「産次」）
            event_date: イベント日付
        
        Returns:
            分類値（文字列）、取得できない場合はNone
        """
        try:
            # 分類が「月」の場合
            if classification == "月":
                try:
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    month = event_dt.month
                    return f"{event_dt.year}-{month:02d}"
                except (ValueError, TypeError):
                    return None
            
            # 分類が「産次」の場合
            if classification == "産次":
                cow_auto_id = event.get('cow_auto_id')
                if not cow_auto_id:
                    return None
                
                # イベント日付時点での産次（LACT）を取得
                # 産次とは、項目「産次（LACT）」のこと
                # RuleEngineを使って、イベント日付時点での状態を計算し、その時点での産次を取得
                if not self.rule_engine:
                    return None
                
                try:
                    # RuleEngineを使って、イベント日付時点での状態を計算
                    state = self.rule_engine.apply_events_until_date(cow_auto_id, event_date)
                    
                    # イベント日付時点での産次を取得
                    lact = state.get('lact', 0)
                    return str(lact) if lact else None
                    
                except Exception as e:
                    logging.error(f"イベント日付時点での産次計算エラー: {e}", exc_info=True)
                    return None
            
            # 分類が「産次（１産、２産、３産以上）」の場合
            if classification == "産次（１産、２産、３産以上）":
                cow_auto_id = event.get('cow_auto_id')
                if not cow_auto_id:
                    return None
                
                # イベント日付時点での産次（LACT）を取得
                # 産次とは、項目「産次（LACT）」のこと
                # RuleEngineを使って、イベント日付時点での状態を計算し、その時点での産次を取得
                if not self.rule_engine:
                    return None
                
                try:
                    # RuleEngineを使って、イベント日付時点での状態を計算
                    state = self.rule_engine.apply_events_until_date(cow_auto_id, event_date)
                    
                    # イベント日付時点での産次を取得
                    lact = state.get('lact', 0)
                    if not lact:
                        return None
                    
                    # 産次を1産、2産、3産以上に分類
                    try:
                        lact_int = int(lact)
                        if lact_int == 1:
                            return "1産"
                        elif lact_int == 2:
                            return "2産"
                        else:
                            return "3産以上"
                    except (ValueError, TypeError):
                        return None
                    
                except Exception as e:
                    logging.error(f"イベント日付時点での産次計算エラー: {e}", exc_info=True)
                    return None
            
            # 分類が「DIM30」の場合
            if classification == "DIM30":
                cow_auto_id = event.get('cow_auto_id')
                if not cow_auto_id:
                    return None
                
                # イベント日付時点でのDIMを取得
                if not self.rule_engine:
                    return None
                
                try:
                    # RuleEngineを使って、イベント日付時点での状態を計算
                    state = self.rule_engine.apply_events_until_date(cow_auto_id, event_date)
                    
                    # イベント日付時点での分娩日を取得
                    clvd = state.get('clvd')
                    if not clvd:
                        return None
                    
                    # DIMを計算（イベント日付 - 分娩日）
                    try:
                        clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
                        event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        dim = (event_dt - clvd_dt).days
                        
                        if dim < 0:
                            return None
                        
                        # 30日ごとに分類（0-30、31-60、61-90、...）
                        dim_range_start = (dim // 30) * 30
                        dim_range_end = dim_range_start + 30
                        return f"{dim_range_start}-{dim_range_end}日"
                    except (ValueError, TypeError) as e:
                        logging.error(f"DIM計算エラー: {e}")
                        return None
                    
                except Exception as e:
                    logging.error(f"イベント日付時点でのDIM計算エラー: {e}", exc_info=True)
                    return None
            
            # その他の分類は未対応（将来拡張可能）
            return None
            
        except Exception as e:
            logging.error(f"分類値取得エラー: {e}", exc_info=True)
            return None
    
    def _sort_classification_value(self, value: str, classification: str) -> tuple:
        """
        分類値のソートキーを取得
        
        Args:
            value: 分類値
            classification: 分類名
        
        Returns:
            ソートキー（タプル）
        """
        # 分類が「月」の場合、年月でソート
        if classification == "月":
            try:
                # "2024-12" 形式を想定
                if "-" in value:
                    parts = value.split("-")
                    if len(parts) == 2:
                        year = int(parts[0])
                        month = int(parts[1])
                        return (0, year, month)
            except (ValueError, TypeError):
                pass
        
        # 分類が「産次」の場合、数値でソート
        if classification == "産次":
            try:
                return (1, int(value))
            except (ValueError, TypeError):
                pass
        
        # 分類が「産次（１産、２産、３産以上）」の場合、1産、2産、3産以上の順でソート
        if classification == "産次（１産、２産、３産以上）":
            if value == "1産":
                return (2, 1)
            elif value == "2産":
                return (2, 2)
            elif value == "3産以上":
                return (2, 3)
        
        # 分類が「DIM30」の場合、数値でソート（0-30、31-60、...）
        if classification == "DIM30":
            try:
                # "0-30日"形式を想定
                if value.endswith("日"):
                    range_str = value[:-1]  # "0-30"を取得
                    if "-" in range_str:
                        parts = range_str.split("-")
                        if len(parts) == 2:
                            start = int(parts[0])
                            return (3, start)
            except (ValueError, TypeError):
                pass
        
        # その他の場合は文字列順
        return (4, str(value))
    
    def _execute_head_count_command(self, classification: str, condition_text: str, command: str):
        """
        頭数専用コマンドを実行（コマンド欄空欄時のデフォルト）
        - 分類なし: 経産牛頭数を1行で表示
        - 分類＝産次: 産次別の頭数を表示
        - condition_text が指定されている場合は条件を満たす牛のみ集計
        """
        try:
            self.current_table_command = command
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                self._update_command_text(self.table_command_text, f"コマンド: {command}")
            
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT auto_id, cow_id AS ID FROM cow WHERE COALESCE(lact, 0) >= 1 ORDER BY cow_id"
            )
            all_cows = cursor.fetchall()
            conn.close()
            
            all_cows = self._filter_existing_cows(all_cows)
            
            # 条件が指定されている場合は条件でフィルタ（集計と同じロジック）
            if condition_text and condition_text.strip() and all_cows:
                or_groups = self._parse_condition_text_to_or_groups(condition_text.strip())
                if or_groups:
                    # item_dictionary から表示名・項目キーのマッピングを取得
                    display_name_to_item_key = {}
                    item_key_to_display_name = {}
                    item_key_lower_to_item_key = {}
                    display_name_lower_to_display_name = {}
                    calc_item_keys = set()
                    if self.item_dict_path and self.item_dict_path.exists():
                        try:
                            with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                                item_dict = json.load(f)
                            for item_key, item_def in item_dict.items():
                                display_name = item_def.get("display_name", "")
                                if display_name:
                                    display_name_to_item_key[display_name] = item_key
                                    item_key_to_display_name[item_key] = display_name
                                    item_key_lower_to_item_key[item_key.lower()] = item_key
                                    display_name_lower_to_display_name[display_name.lower()] = display_name
                                    alias = item_def.get("alias", "")
                                    if alias:
                                        display_name_to_item_key[alias] = item_key
                                        display_name_lower_to_display_name[alias.lower()] = display_name
                                    origin = item_def.get("origin", "")
                                    if origin == "calc":
                                        calc_item_keys.add(item_key)
                        except Exception as e:
                            logging.error(f"item_dictionary.json読み込みエラー: {e}")
                    if len(or_groups) == 1:
                        filtered_cows = [
                            cow_row for cow_row in all_cows
                            if self._aggregate_cow_matches_condition_group(
                                cow_row, or_groups[0],
                                display_name_to_item_key, item_key_to_display_name,
                                item_key_lower_to_item_key, display_name_lower_to_display_name,
                                calc_item_keys
                            )
                        ]
                    else:
                        filtered_cows = [
                            cow_row for cow_row in all_cows
                            if any(
                                self._aggregate_cow_matches_condition_group(
                                    cow_row, g,
                                    display_name_to_item_key, item_key_to_display_name,
                                    item_key_lower_to_item_key, display_name_lower_to_display_name,
                                    calc_item_keys
                                )
                                for g in or_groups
                            )
                        ]
                    all_cows = filtered_cows
            
            if not classification or classification.strip() == "":
                # 経産牛頭数のみ
                count = len(all_cows)
                result_rows = [{"頭数": str(count)}]
                display_columns = ["頭数"]
                self._display_list_result_in_table(display_columns, result_rows, command=command)
                self.result_notebook.select(0)
                return
            
            if classification.strip() == "産次":
                # 産次別頭数
                from collections import defaultdict
                lact_to_cows = defaultdict(list)
                for cow_row in all_cows:
                    cow = self.db.get_cow_by_id(cow_row['ID'])
                    lact = cow.get('lact') if cow else None
                    if lact is not None:
                        try:
                            lact_int = int(lact)
                            lact_to_cows[lact_int].append(cow_row['auto_id'])
                        except (ValueError, TypeError):
                            pass
                
                total = sum(len(cows) for cows in lact_to_cows.values())
                result_rows = []
                aggregate_metadata = {'item_key': 'LACT', 'item_name': '産次', 'is_rc': False, 'value_to_cows': {}}
                
                for lact in sorted(lact_to_cows.keys()):
                    cows = lact_to_cows[lact]
                    count = len(cows)
                    pct = (count / total * 100) if total > 0 else 0
                    label = f"{lact}産"
                    result_rows.append({
                        "項目": label,
                        "頭数": str(count),
                        "％": f"{pct:.1f}",
                        "_aggregate_value": lact
                    })
                    aggregate_metadata['value_to_cows'][str(lact)] = cows
                
                result_rows.append({
                    "項目": "合計",
                    "頭数": str(total),
                    "％": "100.0",
                    "_aggregate_value": None
                })
                display_columns = ["項目", "頭数", "％"]
                self._display_aggregate_result_in_table(display_columns, result_rows, aggregate_metadata, command=command)
                self.result_notebook.select(0)
                return
            
            # 分類が「産次」以外の場合は経産牛頭数のみ表示（将来拡張可能）
            count = len(all_cows)
            result_rows = [{"頭数": str(count)}]
            display_columns = ["頭数"]
            self._display_list_result_in_table(display_columns, result_rows, command=command)
            self.result_notebook.select(0)
        except Exception as e:
            logging.error(f"頭数コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"頭数の集計中にエラーが発生しました: {e}")
    
    def _execute_aggregate_command(self, command: str):
        """
        集計コマンドを実行
        
        Args:
            command: 集計コマンド（例：「集計：DIM」）
        """
        try:
            # コマンドから項目名、分類、条件を抽出
            # 「集計：DIM：分類：条件：DIM>150」のような形式に対応
            # まず「集計：」を除去
            command_body = ""
            if command.startswith("集計："):
                command_body = command[len("集計："):].strip()
            elif command.startswith("集計:"):
                command_body = command[len("集計:"):].strip()
            else:
                self.add_message(role="system", text="集計コマンドの形式が正しくありません。")
                return
            
            if not command_body:
                self.add_message(role="system", text="集計コマンドの形式が正しくありません。")
                return
            
            # 項目名、分類、条件を抽出
            # 形式：「項目：分類：条件」または「項目：分類」または「項目：条件」または「項目」
            # 条件は比較演算子（<, >, <=, >=, =, !=）を含む文字列で判定
            command_classification = ""
            condition_text_from_command = ""
            
            # 「：」で分割（最大3つまで：項目、分類、条件）
            if "：" in command_body:
                parts = command_body.split("：")
            elif ":" in command_body:
                parts = command_body.split(":")
            else:
                parts = [command_body]
            
            if len(parts) == 0:
                self.add_message(role="system", text="集計コマンドの形式が正しくありません。")
                return
            
            # 最初の部分をスペースで分割して複数の項目名を取得
            item_names_str = parts[0].strip()
            item_names = [name.strip() for name in item_names_str.split() if name.strip()]
            
            if not item_names:
                self.add_message(role="system", text="集計する項目を指定してください。")
                return
            
            # 残りの部分から分類と条件を判定
            # 条件は比較演算子を含む文字列で判定（ただし、括弧内の=は分類の一部として扱う）
            condition_pattern = r'[<>=!]'
            
            if len(parts) >= 2:
                # 2つ目以降の部分をチェック
                remaining_parts = [p.strip() for p in parts[1:]]
                
                # 条件部分を検出（比較演算子を含む最初の部分）
                # ただし、括弧内の=は分類の一部として扱う（例：「空胎日数100日（DOPN=100）」）
                condition_index = None
                for i, part in enumerate(remaining_parts):
                    # 括弧内の=を除外して条件パターンをチェック
                    # 括弧で囲まれた部分を一時的に削除してチェック
                    part_without_brackets = re.sub(r'[（(].*?[）)]', '', part)
                    if re.search(condition_pattern, part_without_brackets):
                        condition_index = i
                        break
                
                if condition_index is not None:
                    # 条件が見つかった場合
                    # 条件より前は分類、条件以降はすべて条件として結合
                    if condition_index > 0:
                        command_classification = "：".join(remaining_parts[:condition_index])
                    condition_text_from_command = " ".join(remaining_parts[condition_index:])
                else:
                    # 条件が見つからない場合、すべて分類として扱う
                    command_classification = "：".join(remaining_parts)
            
            # コマンドが「頭数」のみの場合は経産牛頭数（分類なし）または産次別頭数（分類＝産次）を表示
            if len(item_names) == 1 and item_names[0].strip() == "頭数":
                classification = command_classification.strip()
                if hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                    c = self.classification_entry.get().strip()
                    if c:
                        classification = c
                self._execute_head_count_command(classification, condition_text_from_command, command)
                return
            
            # 複数項目と分類の組み合わせは対応可能（制限を削除）
            
            # 集計対象の項目名を保存（条件解析で上書きされないように）
            # 複数項目の場合は最初の項目名を表示用に使用
            aggregate_item_name = item_names[0] if len(item_names) == 1 else " ".join(item_names)
            
            # コマンド情報を保存
            self.current_table_command = command
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                self._update_command_text(self.table_command_text, f"コマンド: {command}")
            
            # item_dictionary.jsonを読み込んで、項目情報を取得
            display_name_to_item_key = {}
            item_key_to_display_name = {}
            item_key_lower_to_item_key = {}
            display_name_lower_to_display_name = {}
            calc_items = set()
            calc_item_keys = set()
            
            if self.item_dict_path and self.item_dict_path.exists():
                try:
                    with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                        item_dict = json.load(f)
                    for item_key, item_def in item_dict.items():
                        display_name = item_def.get("display_name", "")
                        if display_name:
                            display_name_to_item_key[display_name] = item_key
                            item_key_to_display_name[item_key] = display_name
                            item_key_lower_to_item_key[item_key.lower()] = item_key
                            display_name_lower_to_display_name[display_name.lower()] = display_name
                            alias = item_def.get("alias", "")
                            if alias:
                                display_name_to_item_key[alias] = item_key
                                display_name_lower_to_display_name[alias.lower()] = display_name
                            origin = item_def.get("origin", "")
                            if origin == "calc":
                                calc_items.add(display_name)
                                calc_item_keys.add(item_key)
                except Exception as e:
                    logging.error(f"item_dictionary.json読み込みエラー: {e}")
            
            # 任意月の乳検項目を検出（リストコマンドと同じロジック）
            monthly_milk_test_item_keys_1 = {'MILK1M', 'SCC1M', 'LS1M', 'MUN1M', 'PROT1M', 'FAT1M', 'DENFA1M'}
            monthly_milk_test_item_keys_2 = {'MILK2M', 'SCC2M', 'LS2M', 'MUN2M', 'PROT2M', 'FAT2M', 'DENFA2M'}
            
            # 指定された任意月の乳検項目を検出
            detected_monthly_items_1 = set()  # 月1設定を使用する項目
            detected_monthly_items_2 = set()  # 月2設定を使用する項目
            
            for item_name in item_names:
                item_name_lower = item_name.lower()
                # 項目キーで検索（大文字小文字を区別しない）
                if item_name_lower in {k.lower() for k in monthly_milk_test_item_keys_1}:
                    detected_monthly_items_1.add(item_name)
                elif item_name_lower in {k.lower() for k in monthly_milk_test_item_keys_2}:
                    detected_monthly_items_2.add(item_name)
                else:
                    # 表示名で検索（「任意月の乳検」で始まり、「１」または「２」で終わる）
                    if '任意月の乳検' in item_name and ('１' in item_name or '1' in item_name):
                        detected_monthly_items_1.add(item_name)
                    elif '任意月の乳検' in item_name and ('２' in item_name or '2' in item_name):
                        detected_monthly_items_2.add(item_name)
            
            # 任意月の乳検項目が指定された場合、毎回月を聞く
            if detected_monthly_items_1 or detected_monthly_items_2:
                if not self._prompt_monthly_milk_test_months(
                    need_month_1=bool(detected_monthly_items_1),
                    need_month_2=bool(detected_monthly_items_2)
                ):
                        self.add_message(role="system", text="月の指定がキャンセルされました。")
                        return
            
            # 複数項目の場合は、項目名の正規化のみ先に実行（all_cows取得後に処理）
            item_info_list = None
            if len(item_names) > 1:
                # 各項目名を正規化してitem_keyを取得
                item_info_list = []  # [{item_name, normalized_name, item_key, is_lact, is_rc}, ...]
                
                for item_name in item_names:
                    item_name_lower = item_name.lower()
                    normalized_name = None
                    item_key = None
                    
                    # 表示名で検索
                    if item_name_lower in display_name_lower_to_display_name:
                        normalized_name = display_name_lower_to_display_name[item_name_lower]
                        item_key = display_name_to_item_key.get(normalized_name)
                    # 項目キーで検索
                    elif item_name_lower in item_key_lower_to_item_key:
                        item_key = item_key_lower_to_item_key[item_name_lower]
                        if item_key in item_key_to_display_name:
                            normalized_name = item_key_to_display_name[item_key]
                    
                    if not item_key:
                        self.add_message(role="system", text=f"項目「{item_name}」が見つかりません。")
                        return
                    
                    is_lact = (item_key == "LACT" or normalized_name == "産次")
                    is_rc = (item_key == "RC" or normalized_name in ("繁殖コード", "繁殖区分"))
                    
                    item_info_list.append({
                        'item_name': item_name,
                        'normalized_name': normalized_name or item_name,
                        'item_key': item_key,
                        'is_lact': is_lact,
                        'is_rc': is_rc
                    })
            
            # 複数項目の場合は単一項目の処理をスキップ（all_cows取得後に処理される）
            if item_info_list is not None:
                # 単一項目の処理をスキップ（all_cows取得後に複数項目処理が実行される）
                # 変数を初期化（後続の処理でエラーにならないように）
                normalized_item_name = None
                item_key = None
                # 複数項目の場合も分類を取得
                classification = ""
                if command_classification:
                    classification = command_classification
                    logging.debug(f"[集計] コマンドから分類を取得: '{classification}'")
                elif hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                    classification = self.classification_entry.get().strip()
                    logging.debug(f"[集計] 分類入力値: '{classification}'")
                # 条件を取得
                condition_text = ""
                if condition_text_from_command:
                    condition_text = condition_text_from_command
                    logging.debug(f"[集計] コマンドから条件を取得: '{condition_text}'")
                elif hasattr(self, 'condition_entry') and self.condition_entry.winfo_exists():
                    condition_text = self.condition_entry.get().strip()
                    # プレースホルダーの場合は無視
                    if condition_text == self.condition_placeholder:
                        condition_text = ""
            else:
                # 単一項目の処理（既存のロジック）
                # 集計対象の項目名を正規化
                aggregate_item_name_lower = aggregate_item_name.lower()
                normalized_item_name = None
                item_key = None
                
                # 表示名で検索
                if aggregate_item_name_lower in display_name_lower_to_display_name:
                    normalized_item_name = display_name_lower_to_display_name[aggregate_item_name_lower]
                    item_key = display_name_to_item_key.get(normalized_item_name)
                # 項目キーで検索
                elif aggregate_item_name_lower in item_key_lower_to_item_key:
                    item_key = item_key_lower_to_item_key[aggregate_item_name_lower]
                    if item_key in item_key_to_display_name:
                        normalized_item_name = item_key_to_display_name[item_key]
                
                if not item_key:
                    self.add_message(role="system", text=f"項目「{aggregate_item_name}」が見つかりません。")
                    return
                
                # 分類入力欄から分類を取得（コマンドに含まれている場合は優先）
                classification = ""
                if command_classification:
                    classification = command_classification
                    logging.debug(f"[集計] コマンドから分類を取得: '{classification}'")
                elif hasattr(self, 'classification_entry') and self.classification_entry.winfo_exists():
                    classification = self.classification_entry.get().strip()
                    logging.debug(f"[集計] 分類入力値: '{classification}'")
                
                # 条件入力欄から条件を取得（コマンドに含まれている場合は優先）
                condition_text = ""
                if condition_text_from_command:
                    condition_text = condition_text_from_command
                    logging.debug(f"[集計] コマンドから条件を取得: '{condition_text}'")
                elif hasattr(self, 'condition_entry') and self.condition_entry.winfo_exists():
                    condition_text = self.condition_entry.get().strip()
                    # プレースホルダーの場合は無視
                    if condition_text == self.condition_placeholder:
                        condition_text = ""
            
            # 条件を解析（OR 対応・日本語項目名・＝正規化は _parse_condition_text_to_or_groups で統一）
            # 複数項目の場合は条件解析をスキップ（_execute_multi_item_aggregate内で処理）
            if item_info_list is not None:
                or_groups = []
                filter_conditions = []
            else:
                or_groups = self._parse_condition_text_to_or_groups(condition_text) if condition_text else []
                filter_conditions = or_groups[0] if len(or_groups) == 1 else []
            
            # 全牛のデータを取得（集計（経産牛）のため LACT>=1 のみ対象）
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT auto_id, cow_id AS ID FROM cow WHERE COALESCE(lact, 0) >= 1 ORDER BY cow_id"
            )
            all_cows = cursor.fetchall()
            
            # 現存牛のみをフィルタリング（除籍牛を除外）
            all_cows = self._filter_existing_cows(all_cows)
            
            # 複数項目の場合は特別処理（all_cows取得後に実行）
            if item_info_list is not None:
                # 複数項目の場合の分類は既に取得済み（2667行目で取得）
                self._execute_multi_item_aggregate(
                    all_cows, item_info_list, classification, condition_text,
                    display_name_to_item_key, item_key_to_display_name,
                    item_key_lower_to_item_key, display_name_lower_to_display_name,
                    calc_item_keys, command
                )
                return
            
            # 条件に基づいてフィルタリング（リストと同じロジック）
            if filter_conditions and all_cows:
                filtered_cows = []
                
                # calc_items_lowerを取得（条件解析用）
                calc_items_lower = set()
                if display_name_to_item_key:
                    for display_name in display_name_to_item_key.keys():
                        calc_items_lower.add(display_name.lower())
                
                for cow_row in all_cows:
                    cow_auto_id = cow_row['auto_id']
                    match_all_conditions = True
                    
                    # すべての条件をチェック
                    for filter_condition in filter_conditions:
                        condition_item_name = filter_condition['item_name']  # 条件の項目名（集計対象とは別）
                        operator = filter_condition['operator']
                        value_str = filter_condition['value']
                        
                        # 条件項目名を正規化
                        condition_item_name_lower = condition_item_name.lower()
                        condition_normalized_item_name = None
                        condition_item_key = None
                        
                        # 表示名で検索
                        if condition_item_name_lower in display_name_lower_to_display_name:
                            condition_normalized_item_name = display_name_lower_to_display_name[condition_item_name_lower]
                            condition_item_key = display_name_to_item_key.get(condition_normalized_item_name)
                        # 項目キーで検索
                        elif condition_item_name_lower in item_key_lower_to_item_key:
                            condition_item_key = item_key_lower_to_item_key[condition_item_name_lower]
                            if condition_item_key in item_key_to_display_name:
                                condition_normalized_item_name = item_key_to_display_name[condition_item_key]
                        
                        # LACT（産次）の特別処理：LACTはデータベースのlactカラムに対応
                        is_lact_condition = (condition_item_name_lower == "lact" or condition_item_key == "LACT")
                        if is_lact_condition:
                            condition_item_key = "lact"  # データベースカラム名（小文字）
                            condition_normalized_item_name = "産次"
                        
                        # 値を取得
                        row_value = None
                        is_rc_condition = (condition_item_key == "RC" or condition_normalized_item_name in ("繁殖コード", "繁殖区分"))
                        
                        # RCの場合は、RuleEngineから取得
                        if is_rc_condition:
                            try:
                                state = self.rule_engine.apply_events(cow_auto_id)
                                if state and 'rc' in state:
                                    row_value = state['rc']
                            except Exception:
                                pass
                        
                        # LACT（産次）の場合は、データベースから直接取得
                        if row_value is None and is_lact_condition:
                            cow = self.db.get_cow_by_id(cow_row['ID'])
                            if cow and 'lact' in cow:
                                row_value = cow['lact']
                        # 計算項目の場合
                        elif row_value is None and condition_item_key and condition_item_key in calc_item_keys:
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            if calculated and condition_item_key in calculated:
                                row_value = calculated[condition_item_key]
                        # データベースカラムの場合
                        elif row_value is None and condition_item_key:
                            cow = self.db.get_cow_by_id(cow_row['ID'])
                            if cow and condition_item_key in cow:
                                row_value = cow[condition_item_key]
                            elif is_rc_condition and cow and 'rc' in cow:
                                row_value = cow['rc']
                        
                        # 条件を評価
                        if row_value is None:
                            match_all_conditions = False
                            break
                        
                        # 範囲指定のチェック
                        if filter_condition.get('is_range', False):
                            try:
                                row_value_num = float(row_value)
                                min_value = float(value_str)
                                max_value = float(filter_condition.get('range_max', value_str))
                                condition_match = (row_value_num >= min_value and row_value_num <= max_value)
                                if not condition_match:
                                    match_all_conditions = False
                                    break
                            except (ValueError, TypeError):
                                match_all_conditions = False
                                break
                        # 値がリストの場合（複数値指定）
                        elif isinstance(value_str, list):
                            condition_match = False
                            try:
                                row_value_num = float(row_value)
                                for val in value_str:
                                    try:
                                        condition_value_num = float(val)
                                        if row_value_num == condition_value_num:
                                            condition_match = True
                                            break
                                    except (ValueError, TypeError):
                                        if str(row_value) == str(val):
                                            condition_match = True
                                            break
                            except (ValueError, TypeError):
                                for val in value_str:
                                    if str(row_value) == str(val):
                                        condition_match = True
                                        break
                            if not condition_match:
                                match_all_conditions = False
                                break
                        else:
                            # 単一値の場合
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
                                    match_all_conditions = False
                                    break
                            except (ValueError, TypeError):
                                condition_match = False
                                if operator == "=" or operator == "==":
                                    condition_match = (str(row_value).strip() == str(value_str).strip())
                                elif operator == "!=" or operator == "<>":
                                    condition_match = (str(row_value).strip() != str(value_str).strip())
                                elif operator == ">=":
                                    condition_match = (str(value_str).strip() in str(row_value))
                                if not condition_match:
                                    match_all_conditions = False
                                    break
                    
                    if match_all_conditions:
                        filtered_cows.append(cow_row)
                
                all_cows = filtered_cows
            elif or_groups and len(or_groups) > 1 and all_cows:
                # 複数 OR グループ：いずれか 1 グループの条件をすべて満たす牛を残す
                filtered_cows = [
                    cow_row for cow_row in all_cows
                    if any(
                        self._aggregate_cow_matches_condition_group(
                            cow_row, g,
                            display_name_to_item_key, item_key_to_display_name,
                            item_key_lower_to_item_key, display_name_lower_to_display_name,
                            calc_item_keys
                        )
                        for g in or_groups
                    )
                ]
                all_cows = filtered_cows
            
            if not all_cows:
                self.add_message(role="system", text="データがありません。")
                return
            
            # 項目の値を取得して集計
            values = []
            is_numeric = True
            is_rc = (item_key == "RC" or normalized_item_name in ("繁殖コード", "繁殖区分"))
            is_lact = (item_key == "LACT" or normalized_item_name == "産次")
            
            logging.debug(f"[集計] 集計対象項目名: {aggregate_item_name}, 項目キー: {item_key}, 計算項目: {item_key in calc_item_keys}, RC: {is_rc}, LACT: {is_lact}")
            
            for cow_row in all_cows:
                cow_auto_id = cow_row['auto_id']
                value = None
                
                # 計算項目の場合
                if item_key in calc_item_keys:
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    logging.debug(f"[集計] cow_auto_id={cow_auto_id}, calculated keys={list(calculated.keys()) if calculated else []}")
                    if calculated and item_key in calculated:
                        value = calculated[item_key]
                        logging.debug(f"[集計] cow_auto_id={cow_auto_id}, {item_key}={value}")
                        # RCの場合は数値として扱う
                        if is_rc:
                            try:
                                value = int(value) if value is not None else None
                            except (ValueError, TypeError):
                                value = None
                    else:
                        # RCの場合は、RuleEngineから直接取得を試みる
                        if is_rc:
                            try:
                                state = self.rule_engine.apply_events(cow_auto_id)
                                if state and 'rc' in state:
                                    value = state['rc']
                                    logging.debug(f"[集計] cow_auto_id={cow_auto_id}, RC from RuleEngine={value}")
                            except Exception as e:
                                logging.debug(f"[集計] RuleEngine取得エラー: {e}")
                else:
                    # データベースカラムの場合
                    cow = self.db.get_cow_by_id(cow_row['ID'])
                    # LACT（産次）の特別処理：LACTはデータベースのlactカラムに対応
                    if is_lact:
                        if cow and 'lact' in cow:
                            value = cow['lact']
                            logging.debug(f"[集計] cow_id={cow_row['ID']}, LACT from cow table={value}")
                    elif cow and item_key in cow:
                        value = cow[item_key]
                        logging.debug(f"[集計] cow_id={cow_row['ID']}, {item_key}={value}")
                    # RCの場合は、cowテーブルのrcカラムも確認
                    elif is_rc:
                        if cow and 'rc' in cow:
                            value = cow['rc']
                            logging.debug(f"[集計] cow_id={cow_row['ID']}, RC from cow table={value}")
                
                # 値が取得できた場合のみ追加
                if value is not None:
                    # 数値かどうかを判定
                    try:
                        float(value)
                        values.append(float(value))
                    except (ValueError, TypeError):
                        is_numeric = False
                        values.append(str(value))
            
            logging.debug(f"[集計] 取得した値の数: {len(values)}, 数値項目: {is_numeric}")
            
            if not values:
                self.add_message(role="system", text="該当データがありません。")
                return
            
            # RCの場合は常にカテゴリ項目として扱う
            if is_rc:
                is_numeric = False
                # RCの値を文字列に変換（集計用）
                values = [str(int(v)) if isinstance(v, (int, float)) else str(v) for v in values]
            
            # 分類が指定されている場合、分類ごとに集計
            logging.debug(f"[集計] 分類チェック: classification='{classification}', 空でない={bool(classification)}")
            if classification:
                logging.debug(f"[集計] 分類集計を実行: 分類={classification}, 集計対象項目={aggregate_item_name}")
                breakdown_thresholds = []
                breakdown_text = ""
                if hasattr(self, '_pending_breakdown_text') and self._pending_breakdown_text:
                    breakdown_text = self._pending_breakdown_text
                elif hasattr(self, 'breakdown_entry') and self.breakdown_entry.winfo_exists():
                    breakdown_text = self.breakdown_entry.get().strip()
                    if breakdown_text == self.breakdown_placeholder:
                        breakdown_text = ""
                self._pending_breakdown_text = ""
                if breakdown_text:
                        parsed = self._parse_breakdown_thresholds(breakdown_text)
                        if parsed is None:
                            return
                        breakdown_thresholds = parsed
                self._execute_aggregate_with_classification(
                    all_cows, item_key, normalized_item_name, aggregate_item_name, 
                    is_numeric, is_rc, classification, condition_text, 
                    display_name_to_item_key, item_key_to_display_name,
                    item_key_lower_to_item_key, display_name_lower_to_display_name,
                    calc_item_keys, command, breakdown_thresholds
                )
                return
            else:
                logging.debug(f"[集計] 分類が指定されていないため、通常の集計を実行")
            
            # 数値項目の場合：平均値と頭数を表示
            if is_numeric:
                avg_value = sum(values) / len(values)
                count = len(values)
                result_rows = [{"平均": f"{avg_value:.2f}", "頭数": str(count)}]
                display_columns = ["平均", "頭数"]
                self._display_list_result_in_table(display_columns, result_rows, command=command)
            else:
                # 非数値項目（カテゴリ項目）の場合：内訳を集計
                from collections import Counter
                counter = Counter(values)
                total = len(values)
                
                # 結果行を作成
                result_rows = []
                # RCの場合は数値順にソート、それ以外は値順にソート
                if is_rc:
                    sorted_items = sorted(counter.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 999)
                else:
                    sorted_items = sorted(counter.items())
                
                # 集計結果のメタデータを保存（個体一覧表示用）
                aggregate_metadata = {
                    'item_key': item_key,
                    'item_name': normalized_item_name or item_name,
                    'is_rc': is_rc,
                    'value_to_cows': {}  # {value: [cow_auto_id, ...]}
                }
                
                # 各牛の値を記録
                for idx, cow_row in enumerate(all_cows):
                    cow_auto_id = cow_row['auto_id']
                    value = None
                    
                    # 計算項目の場合
                    if item_key in calc_item_keys:
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and item_key in calculated:
                            value = calculated[item_key]
                            if is_rc:
                                try:
                                    value = int(value) if value is not None else None
                                except (ValueError, TypeError):
                                    value = None
                    else:
                        # データベースカラムの場合
                        cow = self.db.get_cow_by_id(cow_row['ID'])
                        if cow and item_key in cow:
                            value = cow[item_key]
                        elif is_rc:
                            if cow and 'rc' in cow:
                                value = cow['rc']
                    
                    if value is not None:
                        # 値の文字列表現をキーとして使用
                        # RCの場合は数値として扱う
                        if is_rc:
                            try:
                                value_key = str(int(value)) if isinstance(value, (int, float)) else str(value)
                            except (ValueError, TypeError):
                                value_key = str(value)
                        else:
                            value_key = str(int(value)) if isinstance(value, (int, float)) else str(value)
                        
                        if value_key not in aggregate_metadata['value_to_cows']:
                            aggregate_metadata['value_to_cows'][value_key] = []
                        aggregate_metadata['value_to_cows'][value_key].append(cow_auto_id)
                        logging.debug(f"[集計] value_key={value_key}, cow_auto_id={cow_auto_id}")
                
                for value, count in sorted_items:
                    percentage = (count / total * 100) if total > 0 else 0
                    # RCの場合はフォーマット
                    if is_rc:
                        try:
                            rc_value = int(value) if str(value).isdigit() else value
                            display_value = self.format_rc(rc_value)
                        except (ValueError, TypeError):
                            display_value = str(value)
                    else:
                        display_value = str(value)
                    
                    result_rows.append({
                        "項目": display_value,
                        "頭数": str(count),
                        "％": f"{percentage:.1f}",
                        "_aggregate_value": value  # 元の値を保存（表示されない）
                    })
                
                # 合計行を追加
                result_rows.append({
                    "項目": "合計",
                    "頭数": str(total),
                    "％": "100.0",
                    "_aggregate_value": None  # 合計行には値がない
                })
                
                display_columns = ["項目", "頭数", "％"]
                self._display_aggregate_result_in_table(display_columns, result_rows, aggregate_metadata, command=command)
                
                # RCの場合はドーナツ円グラフを別ウィンドウで表示
                if is_rc:
                    self._display_rc_donut_chart(result_rows, command=command)
            
            # 表タブを選択
            self.result_notebook.select(0)
            
        except Exception as e:
            logging.error(f"集計コマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"集計コマンドの実行中にエラーが発生しました: {e}")
            try:
                self._clear_result_display()
            except Exception:
                pass
    
    def _execute_aggregate_with_classification(
        self, all_cows, item_key, normalized_item_name, item_name,
        is_numeric, is_rc, classification, condition_text,
        display_name_to_item_key, item_key_to_display_name,
        item_key_lower_to_item_key, display_name_lower_to_display_name,
        calc_item_keys, command, breakdown_thresholds: Optional[List[int]] = None
    ):
        """
        分類を指定した集計を実行
        
        Args:
            all_cows: 全牛のデータ
            item_key: 集計項目のキー
            normalized_item_name: 正規化された項目名
            item_name: 項目名
            is_numeric: 数値項目かどうか
            is_rc: RC項目かどうか
            classification: 分類（例：LACT, DIM7, DIM14など）
            condition_text: 条件テキスト
            display_name_to_item_key: 表示名から項目キーへのマッピング
            item_key_to_display_name: 項目キーから表示名へのマッピング
            item_key_lower_to_item_key: 小文字の項目キーから項目キーへのマッピング
            display_name_lower_to_display_name: 小文字の表示名から表示名へのマッピング
            calc_item_keys: 計算項目のキーセット
            command: 実行コマンド
        """
        try:
            # LACT（産次）の特別処理フラグ
            is_lact = (item_key == "LACT" or normalized_item_name == "産次")
            
            # 分類パターンの定義
            classification_patterns = {
                "産次": {"type": "lact", "label": "産次"},
                "産次（１産、２産、３産以上）": {"type": "lact_category", "label": "産次（１産、２産、３産以上）"},
                "DIM7": {"type": "dim_range", "label": "DIM7日ごと", "interval": 7},
                "DIM14": {"type": "dim_range", "label": "DIM14日ごと", "interval": 14},
                "DIM21": {"type": "dim_range", "label": "DIM21日ごと", "interval": 21},
                "DIM30": {"type": "dim_range", "label": "DIM30日ごと", "interval": 30},
                "DIM50": {"type": "dim_range", "label": "DIM50日ごと", "interval": 50},
                "DIM：産次": {"type": "dim_lact_cross", "label": "産次×DIM"},
                "DIM:産次": {"type": "dim_lact_cross", "label": "産次×DIM"},
                "空胎日数100日（DOPN=100）": {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [100]},
                "空胎日数150日（DOPN=150）": {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [150]},
                "分娩月": {"type": "calvmo", "label": "分娩月"},
            }
            
            # DIM=150, DOPN=150などの形式に対応
            # 例：DIM=150 → {"type": "dim_threshold", "thresholds": [150]}
            # 例：DIM=100,150 → {"type": "dim_threshold", "thresholds": [100, 150]}
            # 例：DOPN=150 → {"type": "dopn_threshold", "thresholds": [150]}
            import re
            pattern = None
            dim_match = re.match(r'^DIM\s*=\s*(.+)$', classification.upper().strip())
            if dim_match:
                thresholds_str = dim_match.group(1)
                try:
                    thresholds = [int(x.strip()) for x in thresholds_str.split(',')]
                    thresholds = sorted(thresholds)
                    pattern = {"type": "dim_threshold", "label": "DIM", "thresholds": thresholds}
                except ValueError:
                    pattern = None
            else:
                dopn_match = re.match(r'^DOPN\s*=\s*(.+)$', classification.upper().strip())
                if dopn_match:
                    thresholds_str = dopn_match.group(1)
                    try:
                        thresholds = [int(x.strip()) for x in thresholds_str.split(',')]
                        thresholds = sorted(thresholds)
                        pattern = {"type": "dopn_threshold", "label": "空胎日数", "thresholds": thresholds}
                    except ValueError:
                        pattern = None
            
            # 分類パターンを取得（DIM=, DOPN=形式で見つからなかった場合）
            if pattern is None:
                # まず完全一致で検索
                pattern = classification_patterns.get(classification)
                logging.debug(f"[集計] 完全一致検索: classification='{classification}', pattern={pattern}")
                
                if not pattern:
                    # 大文字小文字を区別せずに検索
                    classification_upper = classification.upper()
                    for key, value in classification_patterns.items():
                        if key.upper() == classification_upper:
                            pattern = value
                            logging.debug(f"[集計] 大文字小文字を区別しない検索でマッチ: key='{key}'")
                            break
                
                # まだ見つからない場合、「空胎日数100日（DOPN=100）」や「空胎日数100日」のような形式を検出
                if not pattern:
                    # 「空胎日数100日（DOPN=100）」の形式を検出
                    dopn_match_full = re.match(r'^空胎日数\s*(\d+)\s*日\s*[（(]DOPN\s*=\s*\d+[）)]', classification)
                    if dopn_match_full:
                        threshold_str = dopn_match_full.group(1)
                        try:
                            threshold = int(threshold_str)
                            pattern = {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [threshold]}
                            logging.debug(f"[集計] 空胎日数分類（完全形式）を検出: {threshold}日")
                        except ValueError:
                            pass
                    
                    # 「空胎日数100日」の形式を検出
                    if not pattern:
                        dopn_match = re.match(r'^空胎日数\s*(\d+)\s*日', classification)
                        if dopn_match:
                            threshold_str = dopn_match.group(1)
                            try:
                                threshold = int(threshold_str)
                                pattern = {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [threshold]}
                                logging.debug(f"[集計] 空胎日数分類（簡易形式）を検出: {threshold}日")
                            except ValueError:
                                pass
            
            # 分類が空欄の場合は分類なしで単純平均を計算
            if not classification or classification.strip() == "":
                # 分類なしで単純平均を計算
                valid_values = []
                valid_count = 0
                valid_cow_ids = []
                
                for cow_row in all_cows:
                    cow_auto_id = cow_row['auto_id']
                    cow_id = cow_row['ID']
                    
                    # 集計項目の値を取得
                    item_value = None
                    if item_key in calc_item_keys:
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and item_key in calculated:
                            item_value = calculated[item_key]
                            if is_rc:
                                try:
                                    item_value = int(item_value) if item_value is not None else None
                                except (ValueError, TypeError):
                                    item_value = None
                    else:
                        cow = self.db.get_cow_by_id(cow_id)
                        # LACT（産次）の特別処理：LACTはデータベースのlactカラムに対応
                        if is_lact:
                            if cow and 'lact' in cow:
                                item_value = cow['lact']
                        elif cow and item_key in cow:
                            item_value = cow[item_key]
                        elif is_rc:
                            if cow and 'rc' in cow:
                                item_value = cow['rc']
                    
                    # 条件チェック（あれば）
                    if condition_text and condition_text.strip():
                        if not self._check_condition(cow_row, condition_text, display_name_to_item_key, item_key_to_display_name, item_key_lower_to_item_key, display_name_lower_to_display_name, calc_item_keys):
                            continue
                    
                    # 値が有効な場合のみカウント
                    if item_value is not None:
                        if is_numeric:
                            try:
                                float_val = float(item_value)
                                valid_values.append(float_val)
                                valid_count += 1
                                valid_cow_ids.append(cow_auto_id)
                            except (ValueError, TypeError):
                                pass
                        else:
                            valid_count += 1
                            valid_cow_ids.append(cow_auto_id)
                
                # 結果を表示（既存の形式に合わせる）
                result_rows = []
                aggregate_metadata = {
                    'item_key': item_key,
                    'item_name': normalized_item_name or item_name,
                    'is_rc': is_rc,
                    'value_to_cows': {"全件": valid_cow_ids}
                }
                
                if is_numeric and valid_values:
                    avg_value = sum(valid_values) / len(valid_values)
                    result_rows.append({
                        "": "全件",
                        f"{normalized_item_name or item_name}平均": f"{avg_value:.1f}",
                        "頭数": str(valid_count)
                    })
                    display_columns = ["", f"{normalized_item_name or item_name}平均", "頭数"]
                elif not is_numeric:
                    result_rows.append({
                        "": "全件",
                        "頭数": str(valid_count)
                    })
                    display_columns = ["", "頭数"]
                else:
                    result_rows.append({
                        "": "全件",
                        f"{normalized_item_name or item_name}平均": "-",
                        "頭数": "0"
                    })
                    display_columns = ["", f"{normalized_item_name or item_name}平均", "頭数"]
                
                self._display_aggregate_result_in_table(display_columns, result_rows, aggregate_metadata, command=command)
                return
            
            if not pattern:
                self.add_message(role="system", text=f"分類「{classification}」は定義されていません。")
                return
            
            # DIM：LACTの2次元集計の場合は特別処理
            if pattern["type"] == "dim_lact_cross":
                self._execute_dim_lact_cross_aggregation(
                    all_cows, item_key, normalized_item_name, item_name,
                    is_numeric, condition_text, command
                )
                return
            
            # 各牛のデータを取得（集計項目と分類項目の両方）
            cow_data_list = []  # [{cow_auto_id, item_value, classification_value, cow_id}, ...]
            
            for cow_row in all_cows:
                cow_auto_id = cow_row['auto_id']
                cow_id = cow_row['ID']
                
                # 集計項目の値を取得
                item_value = None
                if item_key in calc_item_keys:
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and item_key in calculated:
                        item_value = calculated[item_key]
                        if is_rc:
                            try:
                                item_value = int(item_value) if item_value is not None else None
                            except (ValueError, TypeError):
                                item_value = None
                else:
                    cow = self.db.get_cow_by_id(cow_id)
                    # LACT（産次）の特別処理：LACTはデータベースのlactカラムに対応
                    if is_lact:
                        if cow and 'lact' in cow:
                            item_value = cow['lact']
                    elif cow and item_key in cow:
                        item_value = cow[item_key]
                    elif is_rc:
                        if cow and 'rc' in cow:
                            item_value = cow['rc']
                
                # 分類項目の値を取得
                classification_value = None
                if pattern["type"] == "month":
                    # 月を取得：集計対象項目に関連する月を取得
                    # まず、集計対象項目が分娩関連（DIM、DIMFAI、DAIなど）の場合は分娩月を使用
                    # それ以外の場合は最新のイベント日付から年月を抽出
                    cow = self.db.get_cow_by_auto_id(cow_auto_id)
                    if cow:
                        # 分娩日（clvd）がある場合は分娩月を使用
                        clvd = cow.get('clvd')
                        if clvd:
                            try:
                                clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
                                classification_value = f"{clvd_dt.year}-{clvd_dt.month:02d}"
                            except (ValueError, TypeError):
                                # 分娩日が無効な場合は最新イベント日付を使用
                                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                                if events:
                                    sorted_events = sorted(events, key=lambda e: (e.get('event_date', ''), e.get('id', 0)), reverse=True)
                                    for event in sorted_events:
                                        event_date = event.get('event_date', '')
                                        if event_date:
                                            try:
                                                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                                                classification_value = f"{event_dt.year}-{event_dt.month:02d}"
                                                break
                                            except (ValueError, TypeError):
                                                continue
                        else:
                            # 分娩日がない場合は最新イベント日付を使用
                            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                            if events:
                                sorted_events = sorted(events, key=lambda e: (e.get('event_date', ''), e.get('id', 0)), reverse=True)
                                for event in sorted_events:
                                    event_date = event.get('event_date', '')
                                    if event_date:
                                        try:
                                            event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                                            classification_value = f"{event_dt.year}-{event_dt.month:02d}"
                                            break
                                        except (ValueError, TypeError):
                                            continue
                elif pattern["type"] == "lact":
                    # 産次を取得
                    cow = self.db.get_cow_by_id(cow_id)
                    if cow and 'lact' in cow:
                        classification_value = cow['lact']
                elif pattern["type"] == "lact_category":
                    # 産次分類：1, 2, 3産以上
                    cow = self.db.get_cow_by_id(cow_id)
                    if cow and 'lact' in cow:
                        try:
                            lact = int(cow['lact'])
                            if lact == 1:
                                classification_value = 1
                            elif lact == 2:
                                classification_value = 2
                            else:
                                classification_value = 3  # 3産以上
                        except (ValueError, TypeError):
                            pass
                elif pattern["type"] == "dim_range":
                    # DIMを取得
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and 'DIM' in calculated:
                        dim_value = calculated['DIM']
                        if dim_value is not None:
                            try:
                                dim_int = int(dim_value)
                                # 区間で分類（例：DIM7の場合、0-6, 7-13, 14-20, ...）
                                interval = pattern["interval"]
                                classification_value = f"{dim_int // interval * interval}-{(dim_int // interval + 1) * interval - 1}"
                            except (ValueError, TypeError):
                                pass
                elif pattern["type"] == "dim_threshold":
                    # DIMの閾値による分類（例：DIM=150 → 150未満と150以上）
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and 'DIM' in calculated:
                        dim_value = calculated['DIM']
                        if dim_value is not None:
                            try:
                                dim_int = int(dim_value)
                                thresholds = pattern["thresholds"]
                                # 閾値に基づいて分類
                                if len(thresholds) == 1:
                                    # 1つの閾値の場合：未満と以上
                                    if dim_int < thresholds[0]:
                                        classification_value = f"<{thresholds[0]}"
                                    else:
                                        classification_value = f">={thresholds[0]}"
                                else:
                                    # 複数の閾値の場合：複数の範囲に分類
                                    if dim_int < thresholds[0]:
                                        classification_value = f"<{thresholds[0]}"
                                    else:
                                        found = False
                                        for i in range(len(thresholds) - 1):
                                            if thresholds[i] <= dim_int < thresholds[i + 1]:
                                                classification_value = f"{thresholds[i]}-{thresholds[i + 1] - 1}"
                                                found = True
                                                break
                                        if not found:
                                            classification_value = f">={thresholds[-1]}"
                            except (ValueError, TypeError):
                                pass
                elif pattern["type"] == "dopn_threshold":
                    # DOPNの閾値による分類（例：DOPN=150 → 150未満と150以上）
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and 'DOPN' in calculated:
                        dopn_value = calculated['DOPN']
                        if dopn_value is not None:
                            try:
                                dopn_int = int(dopn_value)
                                thresholds = pattern["thresholds"]
                                # 閾値に基づいて分類
                                if len(thresholds) == 1:
                                    # 1つの閾値の場合：未満と以上
                                    if dopn_int < thresholds[0]:
                                        classification_value = f"<{thresholds[0]}"
                                    else:
                                        classification_value = f">={thresholds[0]}"
                                else:
                                    # 複数の閾値の場合：複数の範囲に分類
                                    if dopn_int < thresholds[0]:
                                        classification_value = f"<{thresholds[0]}"
                                    else:
                                        found = False
                                        for i in range(len(thresholds) - 1):
                                            if thresholds[i] <= dopn_int < thresholds[i + 1]:
                                                classification_value = f"{thresholds[i]}-{thresholds[i + 1] - 1}"
                                                found = True
                                                break
                                        if not found:
                                            classification_value = f">={thresholds[-1]}"
                            except (ValueError, TypeError):
                                pass
                elif pattern["type"] == "calvmo":
                    # 分娩月を取得（CALVMO項目）
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    if calculated and 'CALVMO' in calculated:
                        calvmo_value = calculated['CALVMO']
                        if calvmo_value is not None:
                            try:
                                calvmo_int = int(calvmo_value)
                                # 分娩月は1-12の値
                                if 1 <= calvmo_int <= 12:
                                    classification_value = calvmo_int
                            except (ValueError, TypeError):
                                pass
                elif pattern["type"] == "dim_lact_cross":
                    # DIM：LACTの2次元集計の場合は、この分類処理は使わない
                    # 別のメソッドで処理するため、ここではスキップ
                    pass
                
                # 値が取得できた場合のみ追加（dim_lact_crossの場合はスキップ）
                if pattern["type"] != "dim_lact_cross" and item_value is not None and classification_value is not None:
                    cow_data_list.append({
                        'cow_auto_id': cow_auto_id,
                        'cow_id': cow_id,
                        'item_value': item_value,
                        'classification_value': classification_value
                    })
            
            if not cow_data_list:
                self.add_message(role="system", text="該当データがありません。")
                return
            
            # 分類ごとに集計
            from collections import defaultdict
            classification_groups = defaultdict(list)  # {classification_value: [cow_data, ...]}
            
            for cow_data in cow_data_list:
                classification_groups[cow_data['classification_value']].append(cow_data)
            
            # 結果行を作成
            result_rows = []
            classification_label = pattern["label"]
            
            # 分類値をソート
            if pattern["type"] == "month":
                # 月は年月順（YYYY-MM形式）
                sorted_classifications = sorted(classification_groups.keys(), key=lambda x: self._sort_classification_value(x, "月"))
            elif pattern["type"] == "lact":
                # 産次は数値順
                sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x) if isinstance(x, int) or str(x).isdigit() else 999)
            elif pattern["type"] == "lact_category":
                # 産次分類：1, 2, 3産以上（既にグループ化されている）
                sorted_classifications = sorted(classification_groups.keys())
            elif pattern["type"] == "calvmo":
                # 分娩月は1-12の数値順
                sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x) if isinstance(x, int) or str(x).isdigit() else 999)
            elif pattern["type"] == "dim_range":
                # DIM範囲は開始値でソート
                sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x.split('-')[0]) if isinstance(x, str) and '-' in x else 999)
            elif pattern["type"] == "dim_threshold" or pattern["type"] == "dopn_threshold":
                # 閾値による分類のソート：<100, 100-149, >=150 の順
                def threshold_sort_key(x):
                    if isinstance(x, str):
                        if x.startswith('<'):
                            return (0, int(x[1:]))
                        elif x.startswith('>='):
                            return (2, int(x[2:]))
                        elif '-' in x:
                            parts = x.split('-')
                            return (1, int(parts[0]))
                    return (999, 0)
                sorted_classifications = sorted(classification_groups.keys(), key=threshold_sort_key)
            else:
                sorted_classifications = sorted(classification_groups.keys())
            
            total_count = 0
            total_sum = 0.0
            breakdown_thresholds = breakdown_thresholds or []
            breakdown_labels = []
            if is_numeric and breakdown_thresholds:
                breakdown_thresholds = sorted(set(breakdown_thresholds))
                if len(breakdown_thresholds) == 1:
                    t = breakdown_thresholds[0]
                    breakdown_labels = [f"{t}未満", f"{t}以上"]
                else:
                    breakdown_labels = [f"{breakdown_thresholds[0]}未満"]
                    for i in range(len(breakdown_thresholds) - 1):
                        breakdown_labels.append(f"{breakdown_thresholds[i]}-{breakdown_thresholds[i + 1]}")
                    breakdown_labels.append(f"{breakdown_thresholds[-1]}以上")
            
            def _count_breakdown(values_list: List[float]) -> List[int]:
                counts = [0] * (len(breakdown_thresholds) + 1)
                for val in values_list:
                    placed = False
                    for i, threshold in enumerate(breakdown_thresholds):
                        if val < threshold:
                            counts[i] += 1
                            placed = True
                            break
                    if not placed:
                        counts[-1] += 1
                return counts
            
            total_breakdown_counts = [0] * (len(breakdown_thresholds) + 1) if breakdown_labels else []
            
            for classification_value in sorted_classifications:
                cow_data_list_in_group = classification_groups[classification_value]
                count = len(cow_data_list_in_group)
                total_count += count
                
                # 集計項目の平均を計算
                if is_numeric:
                    values_in_group = [float(cow_data['item_value']) for cow_data in cow_data_list_in_group if cow_data['item_value'] is not None]
                    if values_in_group:
                        avg_value = sum(values_in_group) / len(values_in_group)
                        total_sum += sum(values_in_group)
                        breakdown_counts = []
                        if breakdown_labels:
                            breakdown_counts = _count_breakdown(values_in_group)
                            total_breakdown_counts = [
                                total_breakdown_counts[i] + breakdown_counts[i]
                                for i in range(len(breakdown_counts))
                            ]
                        # 分類値の表示
                        if pattern["type"] == "month":
                            # 月の表示（YYYY-MM形式）
                            classification_display = str(classification_value)
                        elif pattern["type"] == "lact_category":
                            if classification_value == 1:
                                classification_display = "１産"
                            elif classification_value == 2:
                                classification_display = "２産"
                            else:
                                classification_display = "３産以上"
                        elif pattern["type"] == "calvmo":
                            # 分娩月の表示（1-12月）
                            if isinstance(classification_value, int):
                                classification_display = f"{classification_value}月"
                            else:
                                classification_display = str(classification_value)
                        elif pattern["type"] == "dim_range":
                            classification_display = f"{classification_value}日"
                        elif pattern["type"] == "dim_threshold" or pattern["type"] == "dopn_threshold":
                            # 閾値による分類の表示：<100 → "100日未満", >=150 → "150日以上", 100-149 → "100日以上150日未満"
                            if isinstance(classification_value, str):
                                if classification_value.startswith('<'):
                                    threshold = classification_value[1:]
                                    classification_display = f"{threshold}日未満"
                                elif classification_value.startswith('>='):
                                    threshold = classification_value[2:]
                                    classification_display = f"{threshold}日以上"
                                elif '-' in classification_value:
                                    parts = classification_value.split('-')
                                    classification_display = f"{parts[0]}日以上{parts[1]}日未満"
                                else:
                                    classification_display = classification_value
                            else:
                                classification_display = str(classification_value)
                        else:
                            classification_display = str(classification_value)
                        
                        row = {
                            classification_label: classification_display,
                            f"{normalized_item_name or item_name}平均": f"{avg_value:.1f}",
                            "頭数": str(count)
                        }
                        if breakdown_labels:
                            for label, cnt in zip(breakdown_labels, breakdown_counts):
                                row[label] = str(cnt)
                        result_rows.append(row)
                else:
                    # 非数値項目の場合は、分類ごとの内訳を表示（簡易版）
                    if pattern["type"] == "lact_category":
                        if classification_value == 1:
                            classification_display = "１産"
                        elif classification_value == 2:
                            classification_display = "２産"
                        else:
                            classification_display = "３産以上"
                    elif pattern["type"] == "calvmo":
                        # 分娩月の表示（1-12月）
                        if isinstance(classification_value, int):
                            classification_display = f"{classification_value}月"
                        else:
                            classification_display = str(classification_value)
                    elif pattern["type"] == "dim_range":
                        classification_display = f"{classification_value}日"
                    elif pattern["type"] == "dim_threshold" or pattern["type"] == "dopn_threshold":
                        # 閾値による分類の表示：<100 → "100日未満", >=150 → "150日以上", 100-149 → "100日以上150日未満"
                        if isinstance(classification_value, str):
                            if classification_value.startswith('<'):
                                threshold = classification_value[1:]
                                classification_display = f"{threshold}日未満"
                            elif classification_value.startswith('>='):
                                threshold = classification_value[2:]
                                classification_display = f"{threshold}日以上"
                            elif '-' in classification_value:
                                parts = classification_value.split('-')
                                classification_display = f"{parts[0]}日以上{parts[1]}日未満"
                            else:
                                classification_display = classification_value
                        else:
                            classification_display = str(classification_value)
                    else:
                        classification_display = str(classification_value)
                    
                    result_rows.append({
                        classification_label: classification_display,
                        "頭数": str(count)
                    })
            
            # 群（全体）の行を追加
            if is_numeric and total_count > 0:
                total_avg = total_sum / total_count
                total_row = {
                    classification_label: "群",
                    f"{normalized_item_name or item_name}平均": f"{total_avg:.1f}",
                    "頭数": str(total_count)
                }
                if breakdown_labels:
                    for label, cnt in zip(breakdown_labels, total_breakdown_counts):
                        total_row[label] = str(cnt)
                result_rows.append(total_row)
            else:
                result_rows.append({
                    classification_label: "群",
                    "頭数": str(total_count)
                })
            
            # 表示カラムを決定
            if is_numeric:
                display_columns = [classification_label, f"{normalized_item_name or item_name}平均", "頭数"]
                if breakdown_labels:
                    display_columns.extend(breakdown_labels)
            else:
                display_columns = [classification_label, "頭数"]
            
            # 集計メタデータを作成（個体一覧表示用）
            aggregate_metadata = {
                'item_key': item_key,
                'item_name': normalized_item_name or item_name,
                'is_rc': is_rc,
                'value_to_cows': {}  # {classification_value: [cow_auto_id, ...]}
            }
            
            # 分類値ごとの個体リストをメタデータに追加
            for classification_value, cow_data_list_in_group in classification_groups.items():
                cow_auto_ids = [cow_data['cow_auto_id'] for cow_data in cow_data_list_in_group]
                # classification_valueを文字列キーに変換
                classification_key = str(classification_value)
                aggregate_metadata['value_to_cows'][classification_key] = cow_auto_ids
            
            # 結果を表示（集計結果として表示）
            # 結果行に分類値を追加（_aggregate_valueとして）
            for row in result_rows:
                # 分類値を取得（表示値から逆引き）
                classification_display = row[classification_label]
                # 表示値から元の分類値を取得
                for classification_value, cow_data_list_in_group in classification_groups.items():
                    # 分類値の表示を生成して比較
                    if pattern["type"] == "lact_category":
                        if classification_value == 1:
                            display = "１産"
                        elif classification_value == 2:
                            display = "２産"
                        else:
                            display = "３産以上"
                    elif pattern["type"] == "calvmo":
                        if isinstance(classification_value, int):
                            display = f"{classification_value}月"
                        else:
                            display = str(classification_value)
                    elif pattern["type"] == "dim_range":
                        display = f"{classification_value}日"
                    elif pattern["type"] == "dim_threshold" or pattern["type"] == "dopn_threshold":
                        if isinstance(classification_value, str):
                            if classification_value.startswith('<'):
                                threshold = classification_value[1:]
                                display = f"{threshold}日未満"
                            elif classification_value.startswith('>='):
                                threshold = classification_value[2:]
                                display = f"{threshold}日以上"
                            elif '-' in classification_value:
                                parts = classification_value.split('-')
                                display = f"{parts[0]}日以上{parts[1]}日未満"
                            else:
                                display = classification_value
                        else:
                            display = str(classification_value)
                    else:
                        display = str(classification_value)
                    
                    if display == classification_display:
                        row['_aggregate_value'] = classification_value
                        break
                # "群"行の場合は_aggregate_valueを追加しない（合計行として扱う）
                if classification_display == "群":
                    row['_aggregate_value'] = None
            
            self._display_aggregate_result_in_table(display_columns, result_rows, aggregate_metadata, command=command)
            
            # 分類が指定されている場合、グラフも作成（別ウィンドウで表示）
            logging.debug(f"[集計] グラフ作成チェック: classification='{classification}', is_numeric={is_numeric}")
            if classification and is_numeric:
                logging.debug(f"[集計] グラフを作成します: result_rows={len(result_rows)}, display_columns={display_columns}")
                self._display_aggregate_graph(
                    result_rows, display_columns, normalized_item_name or item_name,
                    classification, command
                )
            else:
                logging.debug(f"[集計] グラフを作成しません: classification={classification}, is_numeric={is_numeric}")
            
            # グラフは別ウィンドウで表示されるため、メインUIは常に表タブを選択
            self.result_notebook.select(0)  # 表タブ
            
        except Exception as e:
            logging.error(f"分類集計エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"分類集計の実行中にエラーが発生しました: {e}")
    
    def _execute_dim_lact_cross_aggregation(
        self, all_cows, item_key, normalized_item_name, item_name,
        is_numeric, condition_text, command
    ):
        """
        DIM：LACTの2次元集計を実行（産次×DIMのクロス集計表）
        
        Args:
            all_cows: 全牛のデータ
            item_key: 集計項目のキー（通常はDIM）
            normalized_item_name: 正規化された項目名
            item_name: 項目名
            is_numeric: 数値項目かどうか
            condition_text: 条件テキスト
            command: 実行コマンド
        """
        try:
            from collections import defaultdict
            
            # 各牛の産次とDIMを取得
            cow_data_list = []  # [{cow_auto_id, lact, dim, cow_id}, ...]
            
            for cow_row in all_cows:
                cow_auto_id = cow_row['auto_id']
                cow_id = cow_row['ID']
                
                # 産次を取得
                cow = self.db.get_cow_by_id(cow_id)
                if not cow:
                    continue
                
                lact = cow.get('lact')
                if lact is None:
                    continue
                
                try:
                    lact = int(lact)
                except (ValueError, TypeError):
                    continue
                
                # DIMを取得
                calculated = self.formula_engine.calculate(cow_auto_id)
                dim_value = None
                if calculated and 'DIM' in calculated:
                    dim_value = calculated['DIM']
                
                if dim_value is None:
                    continue
                
                try:
                    dim_int = int(dim_value)
                    # DIMを10日ごとにグループ化（例：0-9, 10-19, 20-29, ...）
                    # 表示用には範囲の開始値を使用（例：0, 10, 20, ...）
                    dim_group = (dim_int // 10) * 10
                    dim_display = dim_group  # 範囲の開始値
                except (ValueError, TypeError):
                    continue
                
                cow_data_list.append({
                    'cow_auto_id': cow_auto_id,
                    'cow_id': cow_id,
                    'lact': lact,
                    'dim': dim_int,
                    'dim_display': dim_display
                })
            
            if not cow_data_list:
                self.add_message(role="system", text="該当データがありません。")
                return
            
            # 産次×DIMの2次元集計
            # cross_table[lact][dim_display] = count
            cross_table = defaultdict(lambda: defaultdict(int))
            lact_set = set()
            dim_display_set = set()
            
            for cow_data in cow_data_list:
                lact = cow_data['lact']
                dim_display = cow_data['dim_display']
                cross_table[lact][dim_display] += 1
                lact_set.add(lact)
                dim_display_set.add(dim_display)
            
            # 産次とDIMをソート
            sorted_lacts = sorted(lact_set)
            sorted_dim_displays = sorted(dim_display_set)
            
            # 結果行を作成
            result_rows = []
            
            # 各産次について行を作成
            for lact in sorted_lacts:
                row = {"産次": str(lact)}
                total_for_lact = 0
                for dim_display in sorted_dim_displays:
                    count = cross_table[lact][dim_display]
                    row[f"DIM{dim_display}"] = str(count) if count > 0 else ""
                    total_for_lact += count
                row["頭数"] = str(total_for_lact)
                result_rows.append(row)
            
            # 群（全体）の行を追加
            total_row = {"産次": "群"}
            total_all = 0
            for dim_display in sorted_dim_displays:
                count = sum(cross_table[lact][dim_display] for lact in sorted_lacts)
                total_row[f"DIM{dim_display}"] = str(count) if count > 0 else ""
                total_all += count
            total_row["頭数"] = str(total_all)
            result_rows.append(total_row)
            
            # 表示カラムを決定
            display_columns = ["産次"] + [f"DIM{dim_display}" for dim_display in sorted_dim_displays] + ["頭数"]
            
            # 結果を表示
            self._display_list_result_in_table(display_columns, result_rows, command=command)
            self.result_notebook.select(0)
            
        except Exception as e:
            logging.error(f"DIM：LACT集計エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"DIM：LACT集計の実行中にエラーが発生しました: {e}")
    
    def _execute_graph_command(self):
        """グラフコマンドを実行（グラフ入力フレームから情報を取得）"""
        try:
            print("グラフコマンド実行開始")
            logging.info("グラフコマンド実行開始")
            from tkinter import messagebox
            from datetime import datetime
            from ui.graph_window import GraphWindow
            
            # グラフ種類を取得
            graph_type = self.graph_type_entry.get() if hasattr(self, 'graph_type_entry') else ""
            logging.info(f"グラフ種類: {graph_type}")
            if not graph_type:
                messagebox.showwarning("警告", "グラフ種類を選択してください。")
                return
            
            # グラフ種類に応じて必要な情報を取得
            x_item = ""
            y_item = ""
            classification = ""
            condition = ""
            
            if graph_type == "棒グラフ" or graph_type == "円グラフ":
                # 棒グラフ・円グラフ：X軸（対象項目）のみ
                x_item = self.graph_x_entry.get().strip() if hasattr(self, 'graph_x_entry') else ""
                if not x_item:
                    messagebox.showwarning("警告", "対象項目を入力してください。")
                    return
                
                # 任意月の乳検項目を検出（X軸をチェック）
                need_month_1, need_month_2 = self._detect_monthly_milk_test_items_in_text(x_item)
                if need_month_1 or need_month_2:
                    if not self._prompt_monthly_milk_test_months(
                        need_month_1=need_month_1,
                        need_month_2=need_month_2
                    ):
                        messagebox.showinfo("情報", "月の指定がキャンセルされました。")
                        return
            elif graph_type == "散布図" or graph_type == "箱ひげ":
                # 散布図・箱ひげ：Y軸とX軸
                y_item = self.graph_y_entry.get().strip() if hasattr(self, 'graph_y_entry') else ""
                x_item = self.graph_x_entry.get().strip() if hasattr(self, 'graph_x_entry') else ""
                if not y_item or not x_item:
                    messagebox.showwarning("警告", "Y軸とX軸の両方を入力してください。")
                    return
                
                # 任意月の乳検項目を検出（Y軸とX軸の両方をチェック）
                items_text = f"{y_item} {x_item}"
                need_month_1, need_month_2 = self._detect_monthly_milk_test_items_in_text(items_text)
                if need_month_1 or need_month_2:
                    if not self._prompt_monthly_milk_test_months(
                        need_month_1=need_month_1,
                        need_month_2=need_month_2
                    ):
                        messagebox.showinfo("情報", "月の指定がキャンセルされました。")
                        return
            elif graph_type == "空胎日数生存曲線":
                # 空胎日数生存曲線：X軸=DIM・Y軸=生存率は固定。分類・条件のみ使用
                x_item = "DIM"
                y_item = ""
            
            # 分類を取得
            if hasattr(self, 'graph_classification_entry'):
                classification = self.graph_classification_entry.get().strip()
            
            # 生存曲線で「イベントの有無」のとき対象イベント番号を取得
            classification_event_number = None
            if graph_type == "空胎日数生存曲線" and classification == "イベントの有無":
                if hasattr(self, 'graph_classification_event_entry'):
                    ev_val = self.graph_classification_event_entry.get().strip()
                    if ev_val and ":" in ev_val:
                        try:
                            classification_event_number = int(ev_val.split(":", 1)[0].strip())
                        except (ValueError, TypeError):
                            pass
                if classification_event_number is None:
                    messagebox.showwarning("警告", "分類「イベントの有無」では対象イベントを選択してください。")
                    return
            # 生存曲線で「項目で分類」のとき対象項目と閾値を取得
            classification_item_key = None
            classification_bin_days = None
            classification_threshold = None
            if graph_type == "空胎日数生存曲線" and classification == "項目で分類":
                if hasattr(self, 'graph_classification_item_entry'):
                    item_val = self.graph_classification_item_entry.get().strip()
                    if item_val and ":" in item_val:
                        classification_item_key = item_val.split(":", 1)[0].strip()
                if not classification_item_key:
                    messagebox.showwarning("警告", "分類「項目で分類」では対象項目を選択してください。")
                    return
                if hasattr(self, 'graph_classification_bin_entry'):
                    bin_val = self.graph_classification_bin_entry.get().strip()
                    if bin_val and bin_val != getattr(self, 'graph_classification_bin_placeholder', '例) 0.1'):
                        try:
                            classification_threshold = float(bin_val)
                        except ValueError:
                            messagebox.showwarning("警告", "閾値は数値で入力してください（例: 0.1）。")
                            return
            
            # 条件を取得
            if hasattr(self, 'graph_condition_entry'):
                condition = self.graph_condition_entry.get().strip()
                if condition == self.graph_condition_placeholder:
                    condition = ""
            
            # 期間を取得
            start_date = ""
            end_date = ""
            if hasattr(self, 'start_date_entry') and hasattr(self, 'end_date_entry'):
                start_date = self.start_date_entry.get().strip()
                end_date = self.end_date_entry.get().strip()
                if start_date == "YYYY-MM-DD":
                    start_date = ""
                if end_date == "YYYY-MM-DD":
                    end_date = ""
            
            # グラフウィンドウを表示（参照を保持してGCを防ぐ）
            logging.info(f"グラフウィンドウ作成: graph_type={graph_type}, x_item={x_item}, y_item={y_item}, classification={classification}")
            
            # GraphWindow参照を保持（GCを防ぐため）
            if not hasattr(self, 'graph_windows'):
                self.graph_windows = []
            
            graph_window = GraphWindow(
                parent=self.root,
                graph_type=graph_type,
                x_item=x_item,
                y_item=y_item,
                classification=classification,
                condition=condition,
                start_date=start_date,
                end_date=end_date,
                db=self.db,
                formula_engine=self.formula_engine,
                rule_engine=self.rule_engine,
                farm_path=self.farm_path,
                event_dictionary_path=self.event_dict_path,
                on_cow_card_requested=self._show_cow_card_view,
                classification_event_number=classification_event_number if graph_type == "空胎日数生存曲線" else None,
                classification_item_key=classification_item_key if graph_type == "空胎日数生存曲線" else None,
                classification_bin_days=classification_bin_days if graph_type == "空胎日数生存曲線" else None,
                classification_threshold=classification_threshold if graph_type == "空胎日数生存曲線" else None
            )
            
            # 参照をリストに追加（複数開く場合に対応）
            self.graph_windows.append(graph_window)
            
            # ウィンドウが閉じられたときに参照を削除
            def on_window_close():
                if graph_window in self.graph_windows:
                    self.graph_windows.remove(graph_window)
                graph_window.window.destroy()
            
            graph_window.window.protocol("WM_DELETE_WINDOW", on_window_close)
            logging.info("グラフウィンドウ作成完了")
            
        except Exception as e:
            logging.error(f"グラフコマンド実行エラー: {e}", exc_info=True)
            from tkinter import messagebox
            messagebox.showerror("エラー", f"グラフの表示に失敗しました: {e}")
    
    def _update_command_text(self, text_widget: tk.Text, text: str):
        """
        コマンド表示テキストを更新（選択・コピー可能な状態を維持）
        
        Args:
            text_widget: 更新するTextウィジェット
            text: 表示するテキスト
        """
        if not text_widget.winfo_exists():
            return
        text_widget.config(state=tk.NORMAL)
        text_widget.delete("1.0", tk.END)
        if text:
            text_widget.insert("1.0", text)
        text_widget.config(state=tk.DISABLED)
    
    def _refresh_main_activity_log(self):
        """メイン画面の操作ログ Treeview を再読込"""
        if not hasattr(self, "activity_log_tree"):
            return
        try:
            for item in self.activity_log_tree.get_children():
                self.activity_log_tree.delete(item)
            for row in load_entries(limit=200):
                ts = str(row.get("ts") or "")
                farm = str(row.get("farm") or "")
                action = str(row.get("action") or "")
                cow_id = str(row.get("cow_id") or "")
                ca = row.get("cow_auto_id")
                cow_auto = "" if ca is None else str(ca)
                ev = str(row.get("event") or "")
                ed = str(row.get("event_date") or "")
                eid = row.get("event_id")
                eid_s = "" if eid is None else str(eid)
                self.activity_log_tree.insert(
                    "",
                    tk.END,
                    values=(ts, farm, action, cow_id, cow_auto, ev, ed, eid_s),
                )
        except Exception as e:
            logging.error("操作ログ表示エラー: %s", e, exc_info=True)
    
    def _on_main_result_notebook_tab_changed(self, event=None):
        """操作ログタブ選択時に最新化（インデックス: 0=表 1=グラフ 2=操作ログ）"""
        try:
            idx = self.result_notebook.index(self.result_notebook.select())
            if idx == 2:
                self._refresh_main_activity_log()
        except Exception:
            pass
    
    def _clear_result_display(self):
        """
        結果表示エリアをクリア（表タブとグラフタブの両方）
        """
        # 表タブをクリア（データのみ、カラムは保持）
        if hasattr(self, 'result_treeview'):
            for item in self.result_treeview.get_children():
                self.result_treeview.delete(item)
            # データとソート状態をクリア（カラムはクリアしない）
            self.table_original_rows = []
            self.table_sort_state = {}
        
        # コマンド情報をクリア
        self.current_table_command = None
        self.current_graph_command = None
        if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
            self._update_command_text(self.table_command_text, "")
        if hasattr(self, 'graph_command_text') and self.graph_command_text.winfo_exists():
            self._update_command_text(self.graph_command_text, "")
        if hasattr(self, 'table_export_btn') and self.table_export_btn.winfo_exists():
            self.table_export_btn.config(state=tk.DISABLED)
        if hasattr(self, 'table_print_btn') and self.table_print_btn.winfo_exists():
            self.table_print_btn.config(state=tk.DISABLED)
        
        # グラフタブをクリア
        if hasattr(self, 'graph_frame'):
            # graph_frame内のすべてのウィジェットを削除
            for widget in self.graph_frame.winfo_children():
                widget.destroy()
        
        # 散布図フレームもクリア（チャット画面内の散布図用）
        if hasattr(self, 'scatter_plot_frame'):
            for widget in self.scatter_plot_frame.winfo_children():
                widget.destroy()
    
    def _display_list_result_in_table(self, columns: List[str], rows: List, command: Optional[str] = None, period: Optional[str] = None):
        """
        リストコマンドの結果を表タブのTreeviewに表示
        
        Args:
            columns: カラム名のリスト
            rows: データ行のリスト（sqlite3.Rowオブジェクト）
            command: 実行したコマンド（オプション）
            period: 対象期間（オプション、例："2024-01-01 ～ 2024-12-31"）
        """
        # 集計メタデータをリセット（リスト表示では使用しない）
        self.aggregate_metadata = None

        # 期間情報を保存
        if period:
            self.current_table_period = period
        else:
            self.current_table_period = None
        # result_treeviewが初期化されているか確認
        if not hasattr(self, 'result_treeview'):
            logging.error("result_treeviewが初期化されていません")
            return
        
        # コマンド情報を保存
        if command:
            self.current_table_command = command
            # コマンド表示テキストを更新
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                # 期間情報を取得
                period_text = ""
                if hasattr(self, 'current_table_period') and self.current_table_period:
                    period_text = f" | 対象期間：{self.current_table_period}"
                
                # イベントリストの場合は件数も表示
                if command.startswith("イベントリスト：") or command.startswith("イベントリスト:"):
                    count_text = f"コマンド: {command} | イベント総数：{len(rows)}件{period_text}"
                    self._update_command_text(self.table_command_text, count_text)
                # リストの場合は件数も表示
                elif command.startswith("リスト：") or command.startswith("リスト:"):
                    count_text = f"コマンド: {command} | リスト総数：{len(rows)}件{period_text}"
                    self._update_command_text(self.table_command_text, count_text)
                else:
                    command_text = f"コマンド: {command}{period_text}"
                    self._update_command_text(self.table_command_text, command_text)
            # ボタンを有効化
            if hasattr(self, 'table_export_btn') and self.table_export_btn.winfo_exists():
                self.table_export_btn.config(state=tk.NORMAL)
            if hasattr(self, 'table_print_btn') and self.table_print_btn.winfo_exists():
                self.table_print_btn.config(state=tk.NORMAL)
        else:
            # コマンド情報がない場合は非表示
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                self._update_command_text(self.table_command_text, "")
            # データがない場合はボタンを無効化
            if not rows:
                if hasattr(self, 'table_export_btn') and self.table_export_btn.winfo_exists():
                    self.table_export_btn.config(state=tk.DISABLED)
                if hasattr(self, 'table_print_btn') and self.table_print_btn.winfo_exists():
                    self.table_print_btn.config(state=tk.DISABLED)
        
        # 元のデータを保存（リスト形式に変換）
        self.table_original_rows = []
        for row_index, row in enumerate(rows):
            row_dict = {}
            # rowが辞書形式の場合
            if isinstance(row, dict):
                for col in columns:
                    row_dict[col] = row.get(col, None)
                # auto_idも保存（表示されていなくても保存）
                if 'auto_id' in row:
                    row_dict['auto_id'] = row['auto_id']
            elif isinstance(row, (list, tuple)):
                # rowがリストまたはタプルの場合（インデックスでアクセス）
                for col_index, col in enumerate(columns):
                    if col_index < len(row):
                        row_dict[col] = row[col_index]
                    else:
                        row_dict[col] = None
                # auto_idが含まれていない場合、IDから取得を試みる
                if 'auto_id' not in row_dict:
                    cow_id = row_dict.get('ID')
                    if cow_id:
                        try:
                            cow = self.db.get_cow_by_id(str(cow_id).zfill(4))
                            if cow:
                                row_dict['auto_id'] = cow.get('auto_id')
                        except Exception as e:
                            logging.debug(f"auto_id取得エラー (cow_id={cow_id}): {e}")
            else:
                # rowがsqlite3.Rowなどの場合
                for col in columns:
                    try:
                        val = row[col] if hasattr(row, 'keys') and col in row.keys() else (row[col] if hasattr(row, '__getitem__') else None)
                        row_dict[col] = val
                    except (KeyError, IndexError, TypeError):
                        row_dict[col] = None
                # auto_idも保存（表示されていなくても保存）
                if hasattr(row, 'keys') and 'auto_id' in row.keys():
                    row_dict['auto_id'] = row['auto_id']
                elif hasattr(row, '__getitem__'):
                    try:
                        if 'auto_id' in row:
                            row_dict['auto_id'] = row['auto_id']
                    except (KeyError, TypeError):
                        pass
                # auto_idが含まれていない場合、IDから取得を試みる
                if 'auto_id' not in row_dict:
                    cow_id = row_dict.get('ID')
                    if cow_id:
                        try:
                            cow = self.db.get_cow_by_id(str(cow_id).zfill(4))
                            if cow:
                                row_dict['auto_id'] = cow.get('auto_id')
                        except Exception as e:
                            logging.debug(f"auto_id取得エラー (cow_id={cow_id}): {e}")
            # 行のインデックスを保存（ダブルクリック時に使用）
            row_dict['_row_index'] = row_index
            self.table_original_rows.append(row_dict)
        
        # ソート状態をリセット
        self.table_sort_state = {col: None for col in columns}
        
        # 既存のデータをクリア
        for item in self.result_treeview.get_children():
            self.result_treeview.delete(item)
        
        # カラムを設定
        self.result_treeview['columns'] = columns
        self.result_treeview['show'] = 'headings'
        
        # 各カラムの設定（クリック時にソートするコマンドを設定）
        for col in columns:
            # 列ヘッダークリック時にソートするコマンドを設定
            self.result_treeview.heading(col, text=col, command=lambda c=col: self._on_column_header_click(c))
            # カラム幅は後で調整（初期値として設定、stretch=Falseで拡大しない）
            # 数値列は右揃え、それ以外は左揃え
            # 受胎率の表の数値列を判定
            numeric_columns = ["受胎率", "受胎", "不受胎", "その他", "総数", "％全授精", "合計"]
            # イベント集計の表の場合、分類値の列（数値）も右揃えにする
            is_event_aggregate = hasattr(self, 'event_aggregate_metadata') and self.event_aggregate_metadata is not None
            if is_event_aggregate:
                # イベント集計の場合、「イベント」列以外は数値列として扱う
                if col != "イベント":
                    anchor = tk.E  # 右揃え
                else:
                    anchor = tk.W  # 左揃え
            elif col in numeric_columns:
                anchor = tk.E  # 右揃え
            else:
                anchor = tk.W  # 左揃え
            self.result_treeview.column(col, width=120, anchor=anchor, minwidth=80, stretch=False)
        
        # データを挿入
        self._refresh_table_display()
        
        # データ挿入後に列幅を再調整（実際の表示内容に基づく）
        self._adjust_column_widths(columns, rows)
        
        # ダブルクリックイベントをバインド
        # 複数項目と分類の集計結果の場合は列（セル）のダブルクリックで個体リストを表示
        # イベント集計の場合は列（セル）のダブルクリックで個体リストを表示
        # それ以外の場合は行のダブルクリックで個体カードを開く
        self.result_treeview.unbind("<Double-Button-1>")
        self.result_treeview.unbind("<Button-1>")
        # イベント集計かどうかを判定
        is_event_aggregate = hasattr(self, 'event_aggregate_metadata') and self.event_aggregate_metadata is not None
        # 複数項目と分類の集計結果かどうかを判定（列名に「平均」が含まれているか、または「分類」列があるか）
        is_multi_item_aggregate = any("平均" in col for col in columns) and "分類" in columns
        if is_event_aggregate:
            self.result_treeview.bind("<Double-Button-1>", self._on_event_aggregate_cell_double_click)
        elif is_multi_item_aggregate:
            self.result_treeview.bind("<Double-Button-1>", self._on_multi_item_aggregate_cell_double_click)
        else:
            self.result_treeview.bind("<Double-Button-1>", self._on_list_row_double_click)
            # リストの行をクリック（シングルクリック）でも個体カードを開く（add='+' で既定の行選択も維持）
            self.result_treeview.bind("<Button-1>", self._on_list_row_click, add='+')
    
    def _adjust_column_widths(self, columns: List[str], rows: List):
        """
        列幅をデータの内容に基づいて自動調整
        
        Args:
            columns: カラム名のリスト
            rows: データ行のリスト（辞書、sqlite3.Row、リストなど）
        """
        if not hasattr(self, 'result_treeview'):
            return
        
        try:
            import tkinter.font as tkfont
            # フォントを取得（列幅計算用）
            font = tkfont.nametofont("TkDefaultFont")
            
            for col in columns:
                # ヘッダーの幅を計算
                header_width = font.measure(str(col)) + 20  # パディング追加
                
                # データの最大幅を計算
                max_data_width = 0
                for row in rows:
                    value = ""
                    # rowが辞書形式の場合
                    if isinstance(row, dict):
                        value = str(row.get(col, ""))
                    elif isinstance(row, (list, tuple)):
                        # リストの場合（インデックスで取得）
                        col_index = columns.index(col) if col in columns else -1
                        if col_index >= 0 and col_index < len(row):
                            value = str(row[col_index] if row[col_index] is not None else "")
                    elif hasattr(row, 'keys') and col in row.keys():
                        # sqlite3.Rowなどの場合
                        value = str(row[col] if row[col] is not None else "")
                    else:
                        # その他の場合
                        try:
                            if hasattr(row, '__getitem__'):
                                value = str(row[col] if row[col] is not None else "")
                        except (KeyError, IndexError, TypeError):
                            value = ""
                    
                    data_width = font.measure(value) + 20  # パディング追加
                    max_data_width = max(max_data_width, data_width)
                
                # 列幅を決定（ヘッダーとデータの最大値、最小値80、最大値300）
                column_width = max(80, min(300, max(header_width, max_data_width)))
                
                # 列幅を設定
                self.result_treeview.column(col, width=int(column_width))
                
        except Exception as e:
            logging.warning(f"列幅の自動調整でエラーが発生しました: {e}")
            # エラーが発生した場合はデフォルト幅（120）を使用
            for col in columns:
                self.result_treeview.column(col, width=120)
    
    def _display_aggregate_result_in_table(self, columns: List[str], rows: List, aggregate_metadata: Dict, command: Optional[str] = None):
        """
        集計結果を表タブのTreeviewに表示（頭数列のダブルクリックで個体一覧を表示）
        
        Args:
            columns: カラム名のリスト
            rows: データ行のリスト
            aggregate_metadata: 集計結果のメタデータ（個体一覧表示用）
            command: 実行したコマンド（オプション）
        """
        # result_treeviewが初期化されているか確認
        if not hasattr(self, 'result_treeview'):
            logging.error("result_treeviewが初期化されていません")
            return
        
        # コマンド情報を保存
        if command:
            self.current_table_command = command
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                self._update_command_text(self.table_command_text, f"コマンド: {command}")
            if hasattr(self, 'table_export_btn') and self.table_export_btn.winfo_exists():
                self.table_export_btn.config(state=tk.NORMAL)
            if hasattr(self, 'table_print_btn') and self.table_print_btn.winfo_exists():
                self.table_print_btn.config(state=tk.NORMAL)
        
        # 集計メタデータを保存
        self.aggregate_metadata = aggregate_metadata
        
        # 元のデータを保存
        self.table_original_rows = []
        for row in rows:
            row_dict = {}
            for col in columns:
                val = row.get(col, None)
                row_dict[col] = val
            # メタデータも保存
            if "_aggregate_value" in row:
                row_dict["_aggregate_value"] = row["_aggregate_value"]
            self.table_original_rows.append(row_dict)
        
        # ソート状態をリセット
        self.table_sort_state = {col: None for col in columns}
        
        # 既存のデータをクリア
        for item in self.result_treeview.get_children():
            self.result_treeview.delete(item)
        
        # カラムを設定
        self.result_treeview['columns'] = columns
        self.result_treeview['show'] = 'headings'
        
        # 各カラムの設定
        for col in columns:
            self.result_treeview.heading(col, text=col, command=lambda c=col: self._on_column_header_click(c))
            # カラム幅は後で調整（初期値として設定）
            self.result_treeview.column(col, width=120, anchor=tk.W, minwidth=80, stretch=False)
        
        # 既存のデータをクリア
        for item in self.result_treeview.get_children():
            self.result_treeview.delete(item)
        
        # メタデータを保存（個体一覧表示用）
        self.aggregate_row_metadata = {}
        
        # データを挿入（メタデータを保持）
        pen_settings = self._get_pen_settings_map() if aggregate_metadata.get("item_key") == "PEN" else None
        row_index = 0
        for row in rows:
            values = []
            for col in columns:
                val = row.get(col, "")
                if pen_settings and col == "項目":
                    val = self._format_pen_value(val, pen_settings)
                values.append(str(val))
            item_id = self.result_treeview.insert('', tk.END, values=values)
            # メタデータを保存（個体一覧表示用）
            if "_aggregate_value" in row:
                self.aggregate_row_metadata[item_id] = {
                    "aggregate_value": row["_aggregate_value"],
                    "row_index": row_index
                }
            row_index += 1
        
        # データ挿入後に列幅を再調整（実際の表示内容に基づく）
        self._adjust_column_widths(columns, rows)
        
        # 頭数列のダブルクリックイベントをバインド（頭数列がある場合のみ）
        # リスト結果のダブルクリックイベントを解除してから、集計結果のイベントをバインド
        self.result_treeview.unbind("<Double-Button-1>")
        # 頭数列がある、または集計結果の場合（aggregate_metadataが存在する場合）はダブルクリックを有効化
        if "頭数" in columns or (hasattr(self, 'aggregate_metadata') and self.aggregate_metadata):
            self.result_treeview.bind("<Double-Button-1>", self._on_aggregate_count_double_click)
            logging.debug("[集計] ダブルクリックイベントをバインドしました")
    
    def _on_aggregate_count_double_click(self, event):
        """集計結果の頭数列をダブルクリックした時の処理（個体一覧を表示）"""
        logging.debug("[集計] ダブルクリックイベント発生")
        
        # クリックされた行を取得
        selection = self.result_treeview.selection()
        logging.debug(f"[集計] 選択された行: {selection}")
        
        if not selection:
            logging.debug("[集計] 行が選択されていません")
            return
        
        item_id = selection[0]
        logging.debug(f"[集計] item_id: {item_id}")
        
        # 行のデータを取得
        values = self.result_treeview.item(item_id, 'values')
        columns = list(self.result_treeview['columns'])
        logging.debug(f"[集計] columns: {columns}, values: {values}")
        
        # 頭数列のインデックスを取得（頭数列がない場合もあるため、オプショナル）
        count_idx = None
        if "頭数" in columns:
            count_idx = columns.index("頭数")
            if count_idx >= len(values):
                logging.debug(f"[集計] 頭数列のインデックスが範囲外: {count_idx} >= {len(values)}")
                return
        
        # 分類列または項目列のインデックスを取得（分類集計の場合は最初の列が分類列、通常集計の場合は項目列）
        classification_idx = None
        if "項目" in columns:
            classification_idx = columns.index("項目")
        elif len(columns) > 0:
            # 分類集計の場合は最初の列が分類列
            classification_idx = 0
        
        if classification_idx is None or classification_idx >= len(values):
            logging.debug(f"[集計] 分類/項目列のインデックスが範囲外: {classification_idx} >= {len(values) if values else 0}")
            return
        
        # 合計行または群行の場合は処理しない
        classification_display_value = values[classification_idx]
        logging.debug(f"[集計] 分類/項目表示値: {classification_display_value}")
        if classification_display_value == "合計" or classification_display_value == "群":
            logging.debug("[集計] 合計/群行のため処理をスキップ")
            return
        
        # メタデータから元の値を取得（aggregate_row_metadataがない場合は、表示値から取得）
        aggregate_value = None
        
        if hasattr(self, 'aggregate_row_metadata') and item_id in self.aggregate_row_metadata:
            row_metadata = self.aggregate_row_metadata[item_id]
            aggregate_value = row_metadata.get("aggregate_value")
            logging.debug(f"[集計] aggregate_row_metadataから取得: {aggregate_value}")
        
        # aggregate_row_metadataがない場合、またはaggregate_valueがNoneの場合は、表示値から推測
        if aggregate_value is None:
            # 表示値から分類値を推測（分類集計の場合）
            aggregate_value = classification_display_value
            logging.debug(f"[集計] 表示値から分類値を取得: {aggregate_value}")
        
        # 該当する個体のauto_idリストを取得
        if not hasattr(self, 'aggregate_metadata'):
            logging.debug("[集計] aggregate_metadataが存在しません")
            return
        
        metadata = self.aggregate_metadata
        value_to_cows = metadata.get('value_to_cows', {})
        logging.debug(f"[集計] value_to_cowsのキー: {list(value_to_cows.keys())}")
        logging.debug(f"[集計] aggregate_valueの型: {type(aggregate_value)}, 値: {aggregate_value}")
        
        # aggregate_valueを文字列に変換して検索
        # すべての可能なキー形式を試す
        possible_keys = []
        try:
            if isinstance(aggregate_value, (int, float)):
                possible_keys.append(str(int(aggregate_value)))
                possible_keys.append(str(float(aggregate_value)))
            possible_keys.append(str(aggregate_value))
        except Exception as e:
            logging.error(f"[集計] 値の変換エラー: {e}")
            possible_keys = [str(aggregate_value)]
        
        cow_auto_ids = []
        for key in possible_keys:
            if key in value_to_cows:
                cow_auto_ids = value_to_cows[key]
                logging.debug(f"[集計] キー '{key}' でマッチ: {len(cow_auto_ids)}頭")
                break
        
        logging.debug(f"[集計] 該当する個体数: {len(cow_auto_ids)}")
        if not cow_auto_ids:
            logging.debug(f"[集計] マッチするキーが見つかりません。試したキー: {possible_keys}, value_to_cowsのキー: {list(value_to_cows.keys())}")
        
        if not cow_auto_ids:
            logging.debug("[集計] 該当する個体がありません")
            return
        
        # 個体一覧ウィンドウを表示
        logging.debug(f"[集計] 個体一覧ウィンドウを表示: {len(cow_auto_ids)}頭")
        self._show_aggregate_cow_list(cow_auto_ids, classification_display_value, metadata)
    
    def _display_aggregate_graph(
        self, result_rows: List[Dict], display_columns: List[str],
        aggregate_item_name: str, classification: str, command: Optional[str] = None
    ):
        """
        集計結果をグラフで表示（分類が指定されている場合のみ）
        別ウィンドウでグラフを表示する
        
        Args:
            result_rows: 集計結果の行データ
            display_columns: 表示カラム名のリスト
            aggregate_item_name: 集計項目名
            classification: 分類名
            command: 実行したコマンド（オプション）
        """
        logging.info(f"[グラフ] _display_aggregate_graph呼び出し開始: result_rows={len(result_rows)}, display_columns={display_columns}")
        
        if not MATPLOTLIB_AVAILABLE:
            logging.error("[グラフ] matplotlibが利用できません")
            return False
        
        try:
            # 新しいウィンドウを作成
            graph_window = tk.Toplevel(self.root)
            graph_window.title(f"グラフ: {aggregate_item_name} × {classification}")
            graph_window.geometry("800x600")
            
            # コマンド表示
            if command:
                command_label = tk.Label(
                    graph_window,
                    text=f"コマンド: {command}",
                    font=(self._main_font_family, 9),
                    anchor=tk.W
                )
                command_label.pack(fill=tk.X, padx=10, pady=5)
            
            # matplotlib FigureとCanvasを作成
            figure = Figure(figsize=(10, 6), dpi=100)
            canvas = FigureCanvasTkAgg(figure, graph_window)
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # コマンド情報を保存
            if command:
                self.current_graph_command = command
                if hasattr(self, 'graph_command_text') and self.graph_command_text.winfo_exists():
                    self._update_command_text(self.graph_command_text, f"コマンド: {command}")
            
            # 分類列と平均値列、頭数列を取得
            classification_col = display_columns[0] if len(display_columns) > 0 else None
            avg_col = None
            count_col = None
            
            for col in display_columns:
                if "平均" in col or "平" in col:
                    avg_col = col
                elif "頭数" in col:
                    count_col = col
            
            if not classification_col or not avg_col or not count_col:
                # 必要な列がない場合はグラフを作成しない
                logging.warning(f"[グラフ] 必要な列が見つかりません: classification_col={classification_col}, avg_col={avg_col}, count_col={count_col}")
                return False
            
            logging.debug(f"[グラフ] データ抽出開始: classification_col={classification_col}, avg_col={avg_col}, count_col={count_col}")
            
            # データを抽出（「群」行は除外）
            x_labels = []
            y_values = []
            n_values = []
            
            logging.debug(f"[グラフ] result_rowsの数: {len(result_rows)}")
            for idx, row in enumerate(result_rows):
                logging.debug(f"[グラフ] 行{idx}: {row}")
                classification_value = row.get(classification_col, "")
                if classification_value == "群" or classification_value == "合計":
                    logging.debug(f"[グラフ] 行{idx}をスキップ（群/合計行）")
                    continue
                
                avg_value_str = row.get(avg_col, "")
                count_value_str = row.get(count_col, "")
                
                logging.debug(f"[グラフ] 行{idx}のデータ: classification_value={classification_value}, avg_value_str={avg_value_str}, count_value_str={count_value_str}")
                
                try:
                    avg_value = float(avg_value_str)
                    count_value = int(count_value_str)
                    
                    x_labels.append(str(classification_value))
                    y_values.append(avg_value)
                    n_values.append(count_value)
                    logging.debug(f"[グラフ] データ追加成功: {classification_value}, {avg_value}, {count_value}")
                except (ValueError, TypeError) as e:
                    logging.warning(f"[グラフ] データ変換エラー: {e}, classification_value={classification_value}, avg_value_str={avg_value_str}, count_value_str={count_value_str}")
                    continue
            
            if not x_labels or not y_values:
                logging.warning(f"[グラフ] データがありません: x_labels={len(x_labels)}, y_values={len(y_values)}")
                return False
            
            logging.debug(f"[グラフ] データ抽出完了: {len(x_labels)}件")
            
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
            
            # 棒グラフを作成（Rのようなスタイル）
            bars = ax.bar(range(len(x_labels)), y_values, color='steelblue', alpha=0.7, edgecolor='black', linewidth=0.5)
            
            # X軸のラベルに頭数を追加（例：1 (N=19)）
            x_labels_with_n = []
            for i, label in enumerate(x_labels):
                n = n_values[i] if i < len(n_values) else 0
                x_labels_with_n.append(f"{label}\n(N={n})")
            
            # X軸のティック位置とラベルを設定
            ax.set_xticks(range(len(x_labels)))
            ax.set_xticklabels(x_labels_with_n, rotation=0, ha='center')
            
            # ラベルとタイトルを設定
            ax.set_xlabel(classification, fontsize=11, fontweight='bold')
            ax.set_ylabel(f"{aggregate_item_name}平均", fontsize=11, fontweight='bold')
            ax.set_title(f"{aggregate_item_name}平均 × {classification}", fontsize=12, fontweight='bold', pad=15)
            
            # グリッドを追加（Rのようなスタイル）
            ax.grid(True, linestyle='--', alpha=0.3, axis='y')
            ax.set_axisbelow(True)
            
            # Y軸の範囲を調整（マイナス値がある場合はデータ範囲に合わせる）
            y_min = min(y_values) if y_values else 0
            y_max = max(y_values) if y_values else 1
            y_range = y_max - y_min or 1
            if y_min < 0:
                ax.set_ylim(y_min - y_range * 0.1, y_max + y_range * 0.1)
            else:
                ax.set_ylim(max(0, y_min - y_range * 0.1), y_max + y_range * 0.1)
            
            # 各棒の上に値を表示（オプション）
            for i, (bar, value) in enumerate(zip(bars, y_values)):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2., height,
                       f'{value:.1f}',
                       ha='center', va='bottom', fontsize=9)
            
            # レイアウトを調整
            figure.tight_layout()
            canvas.draw()
            
            logging.info(f"[グラフ] グラフ描画完了: {len(x_labels)}件のデータを表示")
            
            # グラフが正常に作成されたことを返す
            return True
            
        except Exception as e:
            logging.error(f"[グラフ] グラフ表示エラー: {e}", exc_info=True)
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
                return False
            except Exception as e2:
                logging.error(f"[グラフ] エラー表示も失敗: {e2}")
                return False
    
    def _display_rc_donut_chart(self, result_rows: List[Dict], command: Optional[str] = None) -> bool:
        """
        RCの集計結果をドーナツ円グラフで表示（別ウィンドウ）
        
        Args:
            result_rows: 集計結果の行データ（項目、頭数、％を含む）
            command: 実行したコマンド（オプション）
        
        Returns:
            成功した場合はTrue、失敗した場合はFalse
        """
        logging.info(f"[グラフ] RCドーナツ円グラフ表示開始: result_rows={len(result_rows)}")
        
        if not MATPLOTLIB_AVAILABLE:
            logging.error("[グラフ] matplotlibが利用できません")
            return False
        
        try:
            # 合計行を除外してデータを抽出し、RCの値でソート（時計回りにRC=1→RC=2→...）
            chart_data = []
            
            for row in result_rows:
                if row.get("項目") == "合計":
                    continue
                
                item = row.get("項目", "")
                count_str = row.get("頭数", "0")
                percentage_str = row.get("％", "0.0")
                
                try:
                    count = int(count_str)
                    percentage = float(percentage_str)
                    
                    # 項目名からRCの値を抽出（例："1: Stopped（繁殖停止）" → 1）
                    rc_value = None
                    import re
                    match = re.match(r'^(\d+):', item)
                    if match:
                        rc_value = int(match.group(1))
                    else:
                        # フォールバック：項目名から直接RC値を推測
                        if "Stopped" in item or "繁殖停止" in item:
                            rc_value = 1
                        elif "Fresh" in item or "分娩後" in item:
                            rc_value = 2
                        elif "Bred" in item or "授精後" in item:
                            rc_value = 3
                        elif "Open" in item or "空胎" in item:
                            rc_value = 4
                        elif "Pregnant" in item or "妊娠中" in item:
                            rc_value = 5
                        elif "Dry" in item or "乾乳中" in item:
                            rc_value = 6
                    
                    if rc_value is not None:
                        chart_data.append({
                            'rc_value': rc_value,
                            'item': item,
                            'count': count,
                            'percentage': percentage
                        })
                except (ValueError, TypeError):
                    continue
            
            if not chart_data:
                logging.warning("[グラフ] グラフ表示用のデータがありません")
                return False
            
            # RCの値でソート（時計回りにRC=1→RC=2→...の順序）
            chart_data.sort(key=lambda x: x['rc_value'])
            
            # RC値から日本語名へのマッピング
            rc_japanese_names = {
                1: "繁殖停止",
                2: "分娩後",
                3: "授精後",
                4: "空胎",
                5: "妊娠中",
                6: "乾乳中"
            }
            
            # ソート済みデータからラベルとカウントを抽出
            chart_labels = []  # グラフ内に表示するラベル（日本語のみ）
            chart_counts = []
            rc_values = []
            label_data = []  # 各セグメントのラベル情報を保存
            
            for data in chart_data:
                rc_value = data['rc_value']
                japanese_name = rc_japanese_names.get(rc_value, "")
                # 日本語名のみのラベル（グラフ内に配置）
                label = f"{rc_value}：{japanese_name}\n(N={data['count']} {data['percentage']}%)"
                chart_labels.append(label)
                chart_counts.append(data['count'])
                rc_values.append(rc_value)
                label_data.append({
                    'rc_value': rc_value,
                    'japanese_name': japanese_name,
                    'count': data['count'],
                    'percentage': data['percentage']
                })
            
            # 新しいウィンドウを作成
            graph_window = tk.Toplevel(self.root)
            graph_window.title(f"グラフ: 繁殖区分")
            graph_window.geometry("800x600")
            
            # コマンド表示
            if command:
                command_label = tk.Label(
                    graph_window,
                    text=f"コマンド: {command}",
                    font=(self._main_font_family, 9),
                    anchor=tk.W
                )
                command_label.pack(fill=tk.X, padx=10, pady=5)
            
            # matplotlib FigureとCanvasを作成
            figure = Figure(figsize=(10, 8), dpi=100)
            canvas = FigureCanvasTkAgg(figure, graph_window)
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 日本語フォントの設定
            try:
                import matplotlib.pyplot as plt
                import matplotlib.font_manager as fm
                
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
            
            # ドーナツ円グラフを作成
            ax = figure.add_subplot(111)
            
            # RCの値に応じた色を割り当て（スタイリッシュでスマートな色配分）
            # RC=1 (Stopped): 黒
            # RC=2 (Fresh): 明るい青
            # RC=3 (Bred): ターコイズグリーン
            # RC=4 (Open): オレンジ
            # RC=5 (Pregnant): 赤（ポジティブ、Bredと区別）
            # RC=6 (Dry): グレー（別ステージ）
            rc_color_map = {
                1: '#2C2C2C',  # Stopped: ダークグレー/黒
                2: '#3498DB',  # Fresh: 明るい青
                3: '#16A085',  # Bred: ターコイズグリーン
                4: '#F39C12',  # Open: オレンジ
                5: '#E74C3C',  # Pregnant: 赤（ポジティブ、Bredと区別）
                6: '#95A5A6'   # Dry: グレー（別ステージ）
            }
            
            # RCの値に応じて色を割り当て
            colors = [rc_color_map.get(rc_val, '#CCCCCC') for rc_val in rc_values]
            
            # ドーナツ円グラフを描画（wedgepropsでドーナツの幅を設定）
            # startangle=90で上から開始、時計回りにRC=1→RC=2→...の順序
            # counterclock=Falseで時計回りに設定
            # labelsは空にして、後で各セグメント内に手動で配置
            wedges, texts, autotexts = ax.pie(
                chart_counts,
                labels=None,  # ラベルは後で手動配置
                autopct='',
                startangle=90,
                counterclock=False,  # 時計回りに設定
                colors=colors,
                wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
            )
            
            # 各セグメント内にラベルを配置（グラフと被らせる）
            import math
            for i, (wedge, label_info) in enumerate(zip(wedges, label_data)):
                # セグメントの中心角度を計算
                angle = (wedge.theta2 + wedge.theta1) / 2
                # セグメントの中心位置を計算（ドーナツの内側、半径0.7の位置）
                angle_rad = math.radians(angle)
                x = 0.7 * math.cos(angle_rad)
                y = 0.7 * math.sin(angle_rad)
                
                # ラベルテキスト（日本語のみ、シンプルに）
                label_text = f"{label_info['rc_value']}：{label_info['japanese_name']}\n(N={label_info['count']} {label_info['percentage']}%)"
                
                # テキストを配置（背景色を設定して見やすく）
                ax.text(x, y, label_text, 
                       ha='center', va='center',
                       fontsize=9, fontweight='bold',
                       bbox=dict(boxstyle='round,pad=0.3', 
                                facecolor='white', 
                                alpha=0.85,
                                edgecolor='gray',
                                linewidth=0.5))
            
            # パーセンテージを中央に表示（ドーナツの内側）
            total = sum(chart_counts)
            center_text = f"合計\n{total}頭"
            ax.text(0, 0, center_text, ha='center', va='center', 
                   fontsize=14, fontweight='bold', 
                   bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))
            
            # タイトルを設定
            ax.set_title('繁殖区分分布', fontsize=14, fontweight='bold', pad=20)
            
            # レイアウトを調整
            figure.tight_layout()
            canvas.draw()
            
            logging.info(f"[グラフ] RCドーナツ円グラフ描画完了: {len(chart_data)}件のデータを表示")
            return True
            
        except Exception as e:
            logging.error(f"[グラフ] RCドーナツ円グラフ表示エラー: {e}", exc_info=True)
            # エラー時も新しいウィンドウにエラーメッセージを表示
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
                return False
            except Exception as e2:
                logging.error(f"[グラフ] エラー表示も失敗: {e2}")
            return False
    
    def _execute_multi_item_aggregate(
        self, all_cows, item_info_list: List[Dict], classification: str, condition_text: str,
        display_name_to_item_key: Dict, item_key_to_display_name: Dict,
        item_key_lower_to_item_key: Dict, display_name_lower_to_display_name: Dict,
        calc_item_keys: set, command: str
    ):
        """
        複数項目の集計を実行
        
        Args:
            all_cows: 全牛のデータ
            item_info_list: 項目情報のリスト [{item_name, normalized_name, item_key, is_lact, is_rc}, ...]
            classification: 分類（空文字列の場合は分類なし）
            condition_text: 条件テキスト
            display_name_to_item_key: 表示名から項目キーへのマッピング
            item_key_to_display_name: 項目キーから表示名へのマッピング
            item_key_lower_to_item_key: 小文字の項目キーから項目キーへのマッピング
            display_name_lower_to_display_name: 小文字の表示名から表示名へのマッピング
            calc_item_keys: 計算項目のキーセット
            command: 実行コマンド
        """
        try:
            # 条件を解析（既存のロジックを使用。受胎/空胎キーワード対応）
            filter_conditions = []
            if condition_text:
                condition_parts = condition_text.replace("　", " ").split()
                for part in condition_parts:
                    p = part.strip()
                    if p == "受胎":
                        filter_conditions.append({'item_name': 'RC', 'operator': '=', 'value': [5, 6]})
                        continue
                    if p == "空胎":
                        filter_conditions.append({'item_name': 'RC', 'operator': '=', 'value': [1, 2, 3, 4]})
                        continue
                    match = re.match(r'^([A-Za-z_][A-Za-z0-9_]*)\s*([<>=!]+)\s*(.+)$', p)
                    if match:
                        value_str = match.group(3).strip()
                        operator = match.group(2)
                        
                        if (operator == "=" or operator == "==") and '-' in value_str:
                            range_parts = value_str.split('-', 1)
                            if len(range_parts) == 2:
                                try:
                                    min_value = float(range_parts[0].strip())
                                    max_value = float(range_parts[1].strip())
                                    filter_conditions.append({
                                        'item_name': match.group(1),
                                        'operator': '>=',
                                        'value': str(min_value),
                                        'is_range': True,
                                        'range_max': str(max_value)
                                    })
                                except (ValueError, TypeError):
                                    filter_conditions.append({
                                        'item_name': match.group(1),
                                        'operator': operator,
                                        'value': value_str
                                    })
                        elif ',' in value_str and (operator == "=" or operator == "=="):
                            values = [v.strip() for v in value_str.split(',')]
                            filter_conditions.append({
                                'item_name': match.group(1),
                                'operator': operator,
                                'value': values
                            })
                        else:
                            filter_conditions.append({
                                'item_name': match.group(1),
                                'operator': operator,
                                'value': value_str
                            })
            
            # 条件に基づいてフィルタリング（既存のロジックを使用）
            if filter_conditions and all_cows:
                filtered_cows = []
                calc_items_lower = set()
                if display_name_to_item_key:
                    for display_name in display_name_to_item_key.keys():
                        calc_items_lower.add(display_name.lower())
                
                for cow_row in all_cows:
                    cow_auto_id = cow_row['auto_id']
                    match_all_conditions = True
                    
                    for filter_condition in filter_conditions:
                        condition_item_name = filter_condition['item_name']
                        operator = filter_condition['operator']
                        value_str = filter_condition['value']
                        
                        condition_item_name_lower = condition_item_name.lower()
                        condition_normalized_item_name = None
                        condition_item_key = None
                        
                        if condition_item_name_lower in display_name_lower_to_display_name:
                            condition_normalized_item_name = display_name_lower_to_display_name[condition_item_name_lower]
                            condition_item_key = display_name_to_item_key.get(condition_normalized_item_name)
                        elif condition_item_name_lower in item_key_lower_to_item_key:
                            condition_item_key = item_key_lower_to_item_key[condition_item_name_lower]
                            if condition_item_key in item_key_to_display_name:
                                condition_normalized_item_name = item_key_to_display_name[condition_item_key]
                        
                        if condition_item_name_lower == "lact" and condition_item_key is None:
                            condition_item_key = "lact"
                            condition_normalized_item_name = "産次"
                        
                        row_value = None
                        is_rc_condition = (condition_item_key == "RC" or condition_normalized_item_name in ("繁殖コード", "繁殖区分"))
                        
                        if is_rc_condition:
                            try:
                                state = self.rule_engine.apply_events(cow_auto_id)
                                if state and 'rc' in state:
                                    row_value = state['rc']
                            except Exception:
                                pass
                        
                        if row_value is None and condition_item_key and condition_item_key in calc_item_keys:
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            if calculated and condition_item_key in calculated:
                                row_value = calculated[condition_item_key]
                        elif row_value is None and condition_item_key:
                            cow = self.db.get_cow_by_id(cow_row['ID'])
                            if cow and condition_item_key in cow:
                                row_value = cow[condition_item_key]
                            elif is_rc_condition and cow and 'rc' in cow:
                                row_value = cow['rc']
                        
                        if row_value is None:
                            match_all_conditions = False
                            break
                        
                        # 条件評価（既存のロジックを使用）
                        if filter_condition.get('is_range', False):
                            try:
                                row_value_num = float(row_value)
                                min_value = float(value_str)
                                max_value = float(filter_condition.get('range_max', value_str))
                                condition_match = (row_value_num >= min_value and row_value_num <= max_value)
                                if not condition_match:
                                    match_all_conditions = False
                                    break
                            except (ValueError, TypeError):
                                match_all_conditions = False
                                break
                        elif isinstance(value_str, list):
                            condition_match = False
                            try:
                                row_value_num = float(row_value)
                                for val in value_str:
                                    try:
                                        condition_value_num = float(val)
                                        if row_value_num == condition_value_num:
                                            condition_match = True
                                            break
                                    except (ValueError, TypeError):
                                        if str(row_value) == str(val):
                                            condition_match = True
                                            break
                            except (ValueError, TypeError):
                                for val in value_str:
                                    if str(row_value) == str(val):
                                        condition_match = True
                                        break
                            if not condition_match:
                                match_all_conditions = False
                                break
                        else:
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
                                    match_all_conditions = False
                                    break
                            except (ValueError, TypeError):
                                condition_match = False
                                if operator == "=" or operator == "==":
                                    condition_match = (str(row_value).strip() == str(value_str).strip())
                                elif operator == "!=" or operator == "<>":
                                    condition_match = (str(row_value).strip() != str(value_str).strip())
                                elif operator == ">=":
                                    condition_match = (str(value_str).strip() in str(row_value))
                                if not condition_match:
                                    match_all_conditions = False
                                    break
                    
                    if match_all_conditions:
                        filtered_cows.append(cow_row)
                
                all_cows = filtered_cows
            
            if not all_cows:
                self.add_message(role="system", text="データがありません。")
                return
            
            # 分類がある場合は分類ごとに集計、ない場合は全体で集計
            if classification:
                # 分類ごとに集計（_execute_aggregate_with_classificationのロジックを参考）
                # 分類パターンの定義（_execute_aggregate_with_classificationと同じ）
                classification_patterns = {
                    "月": {"type": "month", "label": "月"},
                    "産次": {"type": "lact", "label": "産次"},
                    "産次（１産、２産、３産以上）": {"type": "lact_category", "label": "産次（１産、２産、３産以上）"},
                    "DIM7": {"type": "dim_range", "label": "DIM7日ごと", "interval": 7},
                    "DIM14": {"type": "dim_range", "label": "DIM14日ごと", "interval": 14},
                    "DIM21": {"type": "dim_range", "label": "DIM21日ごと", "interval": 21},
                    "DIM30": {"type": "dim_range", "label": "DIM30日ごと", "interval": 30},
                    "DIM50": {"type": "dim_range", "label": "DIM50日ごと", "interval": 50},
                    "DIM：産次": {"type": "dim_lact_cross", "label": "産次×DIM"},
                    "DIM:産次": {"type": "dim_lact_cross", "label": "産次×DIM"},
                    "空胎日数100日（DOPN=100）": {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [100]},
                    "空胎日数150日（DOPN=150）": {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [150]},
                    "分娩月": {"type": "calvmo", "label": "分娩月"},
                }
                
                import re
                pattern = None
                dim_match = re.match(r'^DIM\s*=\s*(.+)$', classification.upper().strip())
                if dim_match:
                    thresholds_str = dim_match.group(1)
                    try:
                        thresholds = [int(x.strip()) for x in thresholds_str.split(',')]
                        thresholds = sorted(thresholds)
                        pattern = {"type": "dim_threshold", "label": "DIM", "thresholds": thresholds}
                    except ValueError:
                        pattern = None
                else:
                    dopn_match = re.match(r'^DOPN\s*=\s*(.+)$', classification.upper().strip())
                    if dopn_match:
                        thresholds_str = dopn_match.group(1)
                        try:
                            thresholds = [int(x.strip()) for x in thresholds_str.split(',')]
                            thresholds = sorted(thresholds)
                            pattern = {"type": "dopn_threshold", "label": "空胎日数", "thresholds": thresholds}
                        except ValueError:
                            pattern = None
                
                if pattern is None:
                    pattern = classification_patterns.get(classification)
                    if not pattern:
                        classification_upper = classification.upper()
                        for key, value in classification_patterns.items():
                            if key.upper() == classification_upper:
                                pattern = value
                                break
                    
                    if not pattern:
                        dopn_match = re.match(r'^空胎日数\s*(\d+)\s*日', classification)
                        if dopn_match:
                            threshold_str = dopn_match.group(1)
                            try:
                                threshold = int(threshold_str)
                                pattern = {"type": "dopn_threshold", "label": "空胎日数", "thresholds": [threshold]}
                            except ValueError:
                                pass
                
                if not pattern:
                    self.add_message(role="system", text=f"分類「{classification}」は定義されていません。")
                    return
                
                # 各牛のデータを取得（各項目の値と分類値）
                cow_data_list = []  # [{cow_auto_id, item_values: {item_key: value}, classification_value}, ...]
                
                for cow_row in all_cows:
                    cow_auto_id = cow_row['auto_id']
                    cow_id = cow_row['ID']
                    
                    # 各項目の値を取得
                    item_values = {}
                    for item_info in item_info_list:
                        item_key = item_info['item_key']
                        is_lact = item_info['is_lact']
                        is_rc = item_info['is_rc']
                        
                        value = None
                        if item_key in calc_item_keys:
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            if calculated and item_key in calculated:
                                value = calculated[item_key]
                                if is_rc:
                                    try:
                                        value = int(value) if value is not None else None
                                    except (ValueError, TypeError):
                                        value = None
                        else:
                            cow = self.db.get_cow_by_id(cow_id)
                            if is_lact:
                                if cow and 'lact' in cow:
                                    value = cow['lact']
                            elif cow and item_key in cow:
                                value = cow[item_key]
                            elif is_rc:
                                if cow and 'rc' in cow:
                                    value = cow['rc']
                        
                        if value is not None:
                            try:
                                float(value)
                                item_values[item_key] = float(value)
                            except (ValueError, TypeError):
                                pass
                    
                    # 分類値を取得（_execute_aggregate_with_classificationと同じロジック）
                    classification_value = None
                    if pattern["type"] == "month":
                        # 月を取得：集計対象項目に関連する月を取得
                        # まず、集計対象項目が分娩関連（DIM、DIMFAI、DAIなど）の場合は分娩月を使用
                        # それ以外の場合は最新のイベント日付から年月を抽出
                        cow = self.db.get_cow_by_auto_id(cow_auto_id)
                        if cow:
                            # 分娩日（clvd）がある場合は分娩月を使用
                            clvd = cow.get('clvd')
                            if clvd:
                                try:
                                    clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
                                    classification_value = f"{clvd_dt.year}-{clvd_dt.month:02d}"
                                except (ValueError, TypeError):
                                    # 分娩日が無効な場合は最新イベント日付を使用
                                    events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                                    if events:
                                        sorted_events = sorted(events, key=lambda e: (e.get('event_date', ''), e.get('id', 0)), reverse=True)
                                        for event in sorted_events:
                                            event_date = event.get('event_date', '')
                                            if event_date:
                                                try:
                                                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                                                    classification_value = f"{event_dt.year}-{event_dt.month:02d}"
                                                    break
                                                except (ValueError, TypeError):
                                                    continue
                            else:
                                # 分娩日がない場合は最新イベント日付を使用
                                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                                if events:
                                    sorted_events = sorted(events, key=lambda e: (e.get('event_date', ''), e.get('id', 0)), reverse=True)
                                    for event in sorted_events:
                                        event_date = event.get('event_date', '')
                                        if event_date:
                                            try:
                                                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                                                classification_value = f"{event_dt.year}-{event_dt.month:02d}"
                                                break
                                            except (ValueError, TypeError):
                                                continue
                    elif pattern["type"] == "lact":
                        cow = self.db.get_cow_by_id(cow_id)
                        if cow and 'lact' in cow:
                            classification_value = cow['lact']
                    elif pattern["type"] == "lact_category":
                        cow = self.db.get_cow_by_id(cow_id)
                        if cow and 'lact' in cow:
                            try:
                                lact = int(cow['lact'])
                                if lact == 1:
                                    classification_value = 1
                                elif lact == 2:
                                    classification_value = 2
                                else:
                                    classification_value = 3
                            except (ValueError, TypeError):
                                pass
                    elif pattern["type"] == "dim_range":
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and 'DIM' in calculated:
                            dim_value = calculated['DIM']
                            if dim_value is not None:
                                try:
                                    dim_int = int(dim_value)
                                    interval = pattern["interval"]
                                    classification_value = f"{dim_int // interval * interval}-{(dim_int // interval + 1) * interval - 1}"
                                except (ValueError, TypeError):
                                    pass
                    elif pattern["type"] == "dim_threshold":
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and 'DIM' in calculated:
                            dim_value = calculated['DIM']
                            if dim_value is not None:
                                try:
                                    dim_int = int(dim_value)
                                    thresholds = pattern["thresholds"]
                                    if len(thresholds) == 1:
                                        if dim_int < thresholds[0]:
                                            classification_value = f"<{thresholds[0]}"
                                        else:
                                            classification_value = f">={thresholds[0]}"
                                    else:
                                        if dim_int < thresholds[0]:
                                            classification_value = f"<{thresholds[0]}"
                                        else:
                                            found = False
                                            for i in range(len(thresholds) - 1):
                                                if thresholds[i] <= dim_int < thresholds[i + 1]:
                                                    classification_value = f"{thresholds[i]}-{thresholds[i + 1] - 1}"
                                                    found = True
                                                    break
                                            if not found:
                                                classification_value = f">={thresholds[-1]}"
                                except (ValueError, TypeError):
                                    pass
                    elif pattern["type"] == "dopn_threshold":
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and 'DOPN' in calculated:
                            dopn_value = calculated['DOPN']
                            if dopn_value is not None:
                                try:
                                    dopn_int = int(dopn_value)
                                    thresholds = pattern["thresholds"]
                                    if len(thresholds) == 1:
                                        if dopn_int < thresholds[0]:
                                            classification_value = f"<{thresholds[0]}"
                                        else:
                                            classification_value = f">={thresholds[0]}"
                                    else:
                                        if dopn_int < thresholds[0]:
                                            classification_value = f"<{thresholds[0]}"
                                        else:
                                            found = False
                                            for i in range(len(thresholds) - 1):
                                                if thresholds[i] <= dopn_int < thresholds[i + 1]:
                                                    classification_value = f"{thresholds[i]}-{thresholds[i + 1] - 1}"
                                                    found = True
                                                    break
                                            if not found:
                                                classification_value = f">={thresholds[-1]}"
                                except (ValueError, TypeError):
                                    pass
                    elif pattern["type"] == "calvmo":
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        if calculated and 'CALVMO' in calculated:
                            calvmo_value = calculated['CALVMO']
                            if calvmo_value is not None:
                                try:
                                    calvmo_int = int(calvmo_value)
                                    if 1 <= calvmo_int <= 12:
                                        classification_value = calvmo_int
                                except (ValueError, TypeError):
                                    pass
                    
                    # すべての項目の値と分類値が取得できた場合のみ追加
                    if item_values and classification_value is not None:
                        cow_data_list.append({
                            'cow_auto_id': cow_auto_id,
                            'item_values': item_values,
                            'classification_value': classification_value
                        })
                
                if not cow_data_list:
                    self.add_message(role="system", text="該当データがありません。")
                    return
                
                # 分類ごとに集計
                from collections import defaultdict
                classification_groups = defaultdict(list)  # {classification_value: [cow_data, ...]}
                
                for cow_data in cow_data_list:
                    classification_groups[cow_data['classification_value']].append(cow_data)
                
                # 分類値をソート
                if pattern["type"] == "month":
                    # 月は年月順（YYYY-MM形式）
                    sorted_classifications = sorted(classification_groups.keys(), key=lambda x: self._sort_classification_value(x, "月"))
                elif pattern["type"] == "lact":
                    sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x) if isinstance(x, int) or str(x).isdigit() else 999)
                elif pattern["type"] == "lact_category":
                    sorted_classifications = sorted(classification_groups.keys())
                elif pattern["type"] == "calvmo":
                    sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x) if isinstance(x, int) or str(x).isdigit() else 999)
                elif pattern["type"] == "dim_range":
                    sorted_classifications = sorted(classification_groups.keys(), key=lambda x: int(x.split('-')[0]) if isinstance(x, str) and '-' in x else 999)
                elif pattern["type"] == "dim_threshold" or pattern["type"] == "dopn_threshold":
                    # 閾値による分類は文字列としてソート（<100, >=100など）
                    sorted_classifications = sorted(classification_groups.keys())
                else:
                    sorted_classifications = sorted(classification_groups.keys())
                
                # 結果行を作成
                result_rows = []
                classification_label = pattern["label"]
                
                for classification_value in sorted_classifications:
                    cow_data_group = classification_groups[classification_value]
                    
                    # 各項目の平均を計算
                    result_row = {}
                    # 分類列には分類値を表示（分類名は列ヘッダーに表示）
                    result_row[classification_label] = str(classification_value)
                    
                    # 各項目の平均を計算
                    for item_info in item_info_list:
                        item_key = item_info['item_key']
                        normalized_name = item_info['normalized_name']
                        
                        values = []
                        cow_auto_ids_for_cell = []  # このセルに対応する個体のauto_idリスト
                        for cow_data in cow_data_group:
                            if item_key in cow_data['item_values']:
                                values.append(cow_data['item_values'][item_key])
                                cow_auto_ids_for_cell.append(cow_data['cow_auto_id'])
                        
                        if values:
                            avg_value = sum(values) / len(values)
                            col_name = f"{normalized_name}平均"
                            result_row[col_name] = f"{avg_value:.2f}"
                            # セルに対応する個体のauto_idリストを保存（メタデータとして）
                            result_row[f"_{col_name}_cow_ids"] = cow_auto_ids_for_cell
                    
                    # 頭数を追加
                    result_row["頭数"] = str(len(cow_data_group))
                    # 頭数列に対応する個体のauto_idリストも保存
                    cow_auto_ids_for_count = [cow_data['cow_auto_id'] for cow_data in cow_data_group]
                    result_row["_頭数_cow_ids"] = cow_auto_ids_for_count
                    result_rows.append(result_row)
                
                # 最終行に合計行を追加（群の平均と合計頭数）
                total_row = {}
                total_row[classification_label] = "合計"
                
                # 各項目の全体平均を計算（全分類を合わせた平均）
                for item_info in item_info_list:
                    item_key = item_info['item_key']
                    normalized_name = item_info['normalized_name']
                    
                    # 全データから該当項目の値を取得
                    all_values = []
                    for cow_data in cow_data_list:
                        if item_key in cow_data['item_values']:
                            all_values.append(cow_data['item_values'][item_key])
                    
                    if all_values:
                        total_avg = sum(all_values) / len(all_values)
                        col_name = f"{normalized_name}平均"
                        total_row[col_name] = f"{total_avg:.2f}"
                
                # 合計頭数を追加
                total_row["頭数"] = str(len(cow_data_list))
                # 合計行の各セルに対応する個体のauto_idリストも保存（全データ）
                for item_info in item_info_list:
                    item_key = item_info['item_key']
                    normalized_name = item_info['normalized_name']
                    col_name = f"{normalized_name}平均"
                    # 全データから該当項目の値がある個体のauto_idリスト
                    cow_auto_ids_for_total = [
                        cow_data['cow_auto_id'] for cow_data in cow_data_list
                        if item_key in cow_data['item_values']
                    ]
                    total_row[f"_{col_name}_cow_ids"] = cow_auto_ids_for_total
                total_row["_頭数_cow_ids"] = [cow_data['cow_auto_id'] for cow_data in cow_data_list]
                result_rows.append(total_row)
                
                # 表示列を設定（分類列のヘッダー名を分類名に変更）
                display_columns = [classification_label]
                for item_info in item_info_list:
                    normalized_name = item_info['normalized_name']
                    display_columns.append(f"{normalized_name}平均")
                display_columns.append("頭数")
                
                # 集計メタデータを作成（単一項目と分類の集計結果と同じ形式）
                aggregate_metadata = {
                    'item_key': None,  # 複数項目の場合はNone
                    'item_name': "複数項目",
                    'is_rc': False,
                    'value_to_cows': {}  # {classification_value: [cow_auto_id, ...]}
                }
                
                # 分類ごとの個体リストを保存
                for classification_value in sorted_classifications:
                    cow_data_group = classification_groups[classification_value]
                    cow_auto_ids = [cow_data['cow_auto_id'] for cow_data in cow_data_group]
                    # 分類値を文字列に変換してキーとして使用
                    classification_key = str(classification_value)
                    aggregate_metadata['value_to_cows'][classification_key] = cow_auto_ids
                
                # 合計行の個体リストも保存（"合計"をキーとして使用）
                all_cow_auto_ids = [cow_data['cow_auto_id'] for cow_data in cow_data_list]
                aggregate_metadata['value_to_cows']["合計"] = all_cow_auto_ids
                
                # _display_aggregate_result_in_tableを使用して表示（単一項目と分類と同じ方法）
                self._display_aggregate_result_in_table(display_columns, result_rows, aggregate_metadata, command=command)
                self.result_notebook.select(0)
                return
            else:
                # 分類がない場合は全体で集計（既存のロジック）
                item_results = []  # [{normalized_name, avg_value, count}, ...]
                
                for item_info in item_info_list:
                    item_key = item_info['item_key']
                    normalized_name = item_info['normalized_name']
                    is_lact = item_info['is_lact']
                    is_rc = item_info['is_rc']
                    
                    values = []
                    
                    for cow_row in all_cows:
                        cow_auto_id = cow_row['auto_id']
                        value = None
                        
                        # 計算項目の場合
                        if item_key in calc_item_keys:
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            if calculated and item_key in calculated:
                                value = calculated[item_key]
                                if is_rc:
                                    try:
                                        value = int(value) if value is not None else None
                                    except (ValueError, TypeError):
                                        value = None
                        else:
                            cow = self.db.get_cow_by_id(cow_row['ID'])
                            # LACT（産次）の特別処理
                            if is_lact:
                                if cow and 'lact' in cow:
                                    value = cow['lact']
                            elif cow and item_key in cow:
                                value = cow[item_key]
                            elif is_rc:
                                if cow and 'rc' in cow:
                                    value = cow['rc']
                        
                        if value is not None:
                            try:
                                float(value)
                                values.append(float(value))
                            except (ValueError, TypeError):
                                # 数値でない場合はスキップ（複数項目の場合は数値項目のみ対応）
                                pass
                    
                    if values:
                        avg_value = sum(values) / len(values)
                        count = len(values)
                        item_results.append({
                            'normalized_name': normalized_name,
                            'avg_value': avg_value,
                            'count': count
                        })
                
                if not item_results:
                    self.add_message(role="system", text="該当データがありません。")
                    return
                
                # 結果を表形式で表示
                display_columns = []
                result_row = {}
                
                for item_result in item_results:
                    col_name = f"{item_result['normalized_name']}平均"
                    display_columns.append(col_name)
                    result_row[col_name] = f"{item_result['avg_value']:.2f}"
                
                # 頭数列を追加
                display_columns.append("頭数")
                # すべての項目で同じ頭数のはず（念のため最小値を取る）
                min_count = min(item_result['count'] for item_result in item_results)
                result_row["頭数"] = str(min_count)
                
                result_rows = [result_row]
            
            # コマンド情報を保存
            self.current_table_command = command
            if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                self._update_command_text(self.table_command_text, f"コマンド: {command}")
            
            # 表に表示
            self._display_list_result_in_table(display_columns, result_rows, command=command)
            self.result_notebook.select(0)
            
        except Exception as e:
            logging.error(f"複数項目集計エラー: {e}", exc_info=True)
            self.add_message(role="system", text=f"複数項目集計の実行中にエラーが発生しました: {e}")
    
    def _show_aggregate_cow_list(self, cow_auto_ids: List[int], category_value: str, metadata: Dict):
        """
        集計結果の個体一覧を表示
        
        Args:
            cow_auto_ids: 個体のauto_idリスト
            category_value: カテゴリ値（表示用）
            metadata: 集計メタデータ
        """
        # 個体一覧ウィンドウを作成
        list_window = tk.Toplevel(self.root)
        list_window.title(f"個体一覧: {category_value}")
        list_window.geometry("800x500")
        
        # フレームとスクロールバー
        frame = ttk.Frame(list_window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 集計項目名を取得（metadataから）
        aggregate_item_name = metadata.get('item_name', '項目')
        aggregate_item_key = metadata.get('item_key', '')
        
        # Treeview（ID、産次、繁殖区分、集計項目）
        columns = ["ID", "産次", "RC", aggregate_item_name]
        # 列の表示名（RCは「繁殖区分」で統一）
        def _col_display(col):
            return "繁殖区分" if col == "RC" else col
        # auto_idカラムを追加（非表示、個体カード表示用）
        all_columns = columns + ["auto_id"]
        treeview = ttk.Treeview(frame, columns=all_columns, show="headings", yscrollcommand=scrollbar.set)
        treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=treeview.yview)
        
        # カラム設定
        for col in columns:
            treeview.heading(col, text=_col_display(col))
            treeview.column(col, width=120, anchor=tk.W)
        
        # auto_idカラムは非表示
        treeview.column("auto_id", width=0, minwidth=0)
        treeview.heading("auto_id", text="")
        
        # ソート状態を管理（列名 -> クリック回数: 0=ソートなし, 1=昇順, 2=降順）
        sort_state = {col: 0 for col in columns}
        # 元のデータを保持（ソート解除用）
        original_data = []
        
        # 個体データを取得して表示
        for cow_auto_id in cow_auto_ids:
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            cow_id = cow.get('cow_id', '')
            
            # 計算項目を取得
            calculated = self.formula_engine.calculate(cow_auto_id)
            
            # 産次（LACT）
            lact = calculated.get('LACT', '') if calculated else ''
            if lact is None:
                lact = ''
            
            # RC
            rc = calculated.get('RC', '') if calculated else ''
            if rc is not None:
                rc_display = self.format_rc(rc)
            else:
                rc_display = ''
            
            # 集計項目の値
            aggregate_value = ''
            if aggregate_item_key and calculated:
                aggregate_value = calculated.get(aggregate_item_key, '')
                if aggregate_value is None:
                    aggregate_value = ''
            
            # 行を挿入（auto_idも含める）
            values = [str(cow_id), str(lact), rc_display, str(aggregate_value), str(cow_auto_id)]
            item_id = treeview.insert('', tk.END, values=values)
            original_data.append(values)
        
        # 列ヘッダーをクリックしたときのソート処理
        def on_column_click(event):
            # クリックされた列を特定
            region = treeview.identify_region(event.x, event.y)
            if region != "heading":
                return
            
            column = treeview.identify_column(event.x)
            if not column:
                return
            
            # 列インデックスを取得（#1, #2, ... の形式）
            col_index = int(column.replace('#', '')) - 1
            if col_index < 0 or col_index >= len(columns):
                return
            
            col_name = columns[col_index]
            
            # クリック回数を更新（0 -> 1 -> 2 -> 0 のサイクル）
            sort_state[col_name] = (sort_state[col_name] + 1) % 3
            
            # 他の列のソート状態をリセット
            for other_col in columns:
                if other_col != col_name:
                    sort_state[other_col] = 0
            
            # データをソート
            if sort_state[col_name] == 0:
                # ソートなし（元の順序に戻す）
                # 既存のアイテムを削除
                for item in treeview.get_children():
                    treeview.delete(item)
                # 元のデータを再挿入
                for values in original_data:
                    treeview.insert('', tk.END, values=values)
                # ヘッダーをリセット
                treeview.heading(col_name, text=_col_display(col_name))
            elif sort_state[col_name] == 1:
                # 昇順ソート
                items = [(treeview.item(item, 'values'), item) for item in treeview.get_children()]
                # ソートキーを取得（数値として比較できる場合は数値で比較）
                def sort_key(item_data):
                    value = item_data[0][col_index] if item_data[0] else ''
                    # 数値として比較を試みる
                    try:
                        return (0, float(value))  # 数値の場合は (0, 数値)
                    except (ValueError, TypeError):
                        return (1, str(value))  # 文字列の場合は (1, 文字列)
                
                items.sort(key=sort_key)
                # 既存のアイテムを削除
                for item in treeview.get_children():
                    treeview.delete(item)
                # ソート済みデータを再挿入
                for values, _ in items:
                    treeview.insert('', tk.END, values=values)
                # ヘッダーに昇順マークを追加
                treeview.heading(col_name, text=f"{_col_display(col_name)} ↑")
            else:  # sort_state[col_name] == 2
                # 降順ソート
                items = [(treeview.item(item, 'values'), item) for item in treeview.get_children()]
                # ソートキーを取得（数値として比較できる場合は数値で比較）
                def sort_key(item_data):
                    value = item_data[0][col_index] if item_data[0] else ''
                    # 数値として比較を試みる
                    try:
                        return (0, float(value))  # 数値の場合は (0, 数値)
                    except (ValueError, TypeError):
                        return (1, str(value))  # 文字列の場合は (1, 文字列)
                
                items.sort(key=sort_key, reverse=True)
                # 既存のアイテムを削除
                for item in treeview.get_children():
                    treeview.delete(item)
                # ソート済みデータを再挿入
                for values, _ in items:
                    treeview.insert('', tk.END, values=values)
                # ヘッダーに降順マークを追加
                treeview.heading(col_name, text=f"{_col_display(col_name)} ↓")
        
        # 列ヘッダーのクリックイベントをバインド
        treeview.bind("<Button-1>", on_column_click)
        
        # ダブルクリックで個体カードを開く
        def on_cow_double_click(event):
            selection = treeview.selection()
            if not selection:
                return
            
            item_id = selection[0]
            # valuesからauto_idを取得（最後のカラム）
            values = treeview.item(item_id, 'values')
            if values and len(values) > 0:
                try:
                    # auto_idは最後のカラム（非表示）
                    auto_id_str = values[-1] if len(values) > len(columns) else None
                    if auto_id_str:
                        auto_id = int(auto_id_str)
                        cow = self.db.get_cow_by_auto_id(auto_id)
                        if cow:
                            cow_id = cow.get('cow_id', '')
                            self._jump_to_cow_card(cow_id)
                            list_window.destroy()
                except (ValueError, TypeError, IndexError) as e:
                    logging.error(f"[集計] 個体カード表示エラー: {e}")
        
        treeview.bind("<Double-Button-1>", on_cow_double_click)
        
        # 閉じるボタン
        close_btn = ttk.Button(list_window, text="閉じる", command=list_window.destroy)
        close_btn.pack(pady=5)
    
    def _on_multi_item_aggregate_cell_double_click(self, event):
        """
        複数項目と分類の集計結果のセルをダブルクリックした時の処理（個体リストを表示）
        """
        # クリックされた位置から列と行を取得
        region = self.result_treeview.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        # クリックされたアイテム（行）を取得
        item = self.result_treeview.identify_row(event.y)
        if not item:
            return
        
        # クリックされた列を取得
        column = self.result_treeview.identify_column(event.x)
        if not column:
            return
        
        # 列インデックスを取得（#0はツリー列、#1以降がデータ列）
        try:
            col_index = int(column.replace('#', '')) - 1
            if col_index < 0:
                return
        except (ValueError, TypeError):
            return
        
        # 現在の列名を取得
        columns = list(self.result_treeview['columns'])
        if col_index >= len(columns):
            return
        
        col_name = columns[col_index]
        
        # 「分類」列の場合は処理しない
        if col_name == "分類":
            return
        
        # クリックされた行のインデックスを取得
        item_index = None
        for idx, child in enumerate(self.result_treeview.get_children()):
            if child == item:
                item_index = idx
                break
        
        if item_index is None or item_index >= len(self.table_original_rows):
            return
        
        # 元のデータ行を取得
        original_row = self.table_original_rows[item_index]
        
        # セルに対応する個体のauto_idリストを取得
        cow_ids_key = f"_{col_name}_cow_ids"
        if cow_ids_key not in original_row:
            return
        
        cow_auto_ids = original_row[cow_ids_key]
        if not cow_auto_ids:
            return
        
        # 分類値と列名から個体リストのタイトルを生成
        classification_value = original_row.get("分類", "")
        if classification_value == "合計":
            title = f"個体一覧: {col_name}（全体）"
        else:
            title = f"個体一覧: {col_name}（分類={classification_value}）"
        
        # 個体リストを表示
        self._show_aggregate_cow_list(cow_auto_ids, title, {})
    
    def _on_column_header_click(self, column: str):
        """
        列ヘッダークリック時の処理（ソート状態を循環させる）
        
        Args:
            column: クリックされた列名
        """
        # 現在のソート状態を確認
        current_state = self.table_sort_state.get(column, None)
        
        # 状態を循環させる: None → 'asc' → 'desc' → None
        if current_state is None:
            # 他の列のソート状態をリセット
            for col in self.table_sort_state:
                self.table_sort_state[col] = None
            self.table_sort_state[column] = 'asc'
        elif current_state == 'asc':
            self.table_sort_state[column] = 'desc'
        else:  # 'desc'
            self.table_sort_state[column] = None
        
        # テーブルを再表示
        self._refresh_table_display()
    
    def _refresh_table_display(self):
        """
        ソート状態に基づいてテーブルの表示を更新
        """
        # 既存のデータをクリア
        for item in self.result_treeview.get_children():
            self.result_treeview.delete(item)
        
        # ソートする列と方向を取得
        sort_column = None
        sort_direction = None
        for col, state in self.table_sort_state.items():
            if state is not None:
                sort_column = col
                sort_direction = state
                break
        
        # データをソート
        if sort_column and sort_direction:
            sorted_rows = sorted(
                self.table_original_rows,
                key=lambda row: self._get_sort_key(row.get(sort_column)),
                reverse=(sort_direction == 'desc')
            )
        else:
            # ソートなし（元の順序）
            sorted_rows = self.table_original_rows
        
        # 列名を取得（Treeviewのcolumnsから）
        columns = list(self.result_treeview['columns'])

        # PEN表示用の設定を取得（必要な場合のみ）
        pen_settings = None
        if "群" in columns or "PEN" in columns:
            pen_settings = self._get_pen_settings_map()
        elif hasattr(self, 'aggregate_metadata') and self.aggregate_metadata:
            if self.aggregate_metadata.get("item_key") == "PEN" and "項目" in columns:
                pen_settings = self._get_pen_settings_map()
        
        # データを挿入
        for row_index, row in enumerate(sorted_rows):
            values = []
            for col in columns:
                val = row.get(col)
                # 繁殖区分の場合はフォーマット
                if col == "繁殖コード" or col == "繁殖区分" or col == "RC":
                    val = self.format_rc(val)
                if pen_settings:
                    if col in ("群", "PEN") or (col == "項目" and getattr(self, 'aggregate_metadata', {}).get("item_key") == "PEN"):
                        val = self._format_pen_value(val, pen_settings)
                values.append(str(val) if val is not None else "")
            # 行のインデックスをタグとして保存（ダブルクリック時に元の行を特定するため）
            original_index = row.get('_row_index', row_index)
            item = self.result_treeview.insert('', tk.END, values=values, tags=(f"row_{original_index}",))
    
    def _on_event_aggregate_cell_double_click(self, event):
        """
        イベント集計のセルをダブルクリックした時の処理（個体リストを表示）
        """
        # クリックされた位置から列と行を取得
        region = self.result_treeview.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        # クリックされたアイテム（行）を取得
        item = self.result_treeview.identify_row(event.y)
        if not item:
            return
        
        # クリックされた列を取得
        column = self.result_treeview.identify_column(event.x)
        if not column:
            return
        
        # 列インデックスを取得（#0はツリー列、#1以降がデータ列）
        try:
            col_index = int(column.replace('#', '')) - 1
            if col_index < 0:
                return
        except (ValueError, TypeError):
            return
        
        # 現在の列名を取得
        columns = list(self.result_treeview['columns'])
        if col_index >= len(columns):
            return
        
        col_name = columns[col_index]
        
        # 「イベント」列または「合計」列の場合は処理しない
        if col_name == "イベント" or col_name == "合計":
            return
        
        # クリックされた行のインデックスを取得
        item_index = None
        for idx, child in enumerate(self.result_treeview.get_children()):
            if child == item:
                item_index = idx
                break
        
        if item_index is None or item_index >= len(self.table_original_rows):
            return
        
        # 元のデータ行を取得
        original_row = self.table_original_rows[item_index]
        
        # イベント集計のメタデータを取得
        if not hasattr(self, 'event_aggregate_metadata') or not self.event_aggregate_metadata:
            return
        
        metadata = self.event_aggregate_metadata
        start_date = metadata['start_date']
        end_date = metadata['end_date']
        classification = metadata['classification']
        selected_event_numbers = metadata['selected_event_numbers']
        event_dict = metadata['event_dict']
        
        # 行からイベント名を取得
        event_name = original_row.get('イベント', '')
        if not event_name:
            return
        
        # イベント名からイベント番号を取得
        event_number = None
        for ev_num, ev_name in event_dict.items():
            if ev_name == event_name:
                event_number = ev_num
                break
        
        if event_number is None:
            return
        
        # 分類値を取得（列名が分類値）
        classification_value = col_name
        
        # 該当するイベントを持つ個体を取得
        # 期間内で、指定されたイベント番号で、指定された分類値に該当するイベントを持つ個体を取得
        events = self.db.get_events_by_period(start_date, end_date)
        
        # イベント番号でフィルタリング
        filtered_events = []
        for ev in events:
            if ev.get('event_number') == event_number:
                filtered_events.append(ev)
        
        filtered_events = self._filter_events_unique_per_cow_per_lact(filtered_events)
        
        # 分類値でフィルタリング
        matching_cow_auto_ids = set()
        for ev in filtered_events:
            event_date = ev.get('event_date', '')
            if not event_date:
                continue
            
            # このイベントの分類値を取得
            ev_classification_value = self._get_event_classification_value(ev, classification, event_date)
            # デバッグログ（必要に応じて有効化）
            # logging.debug(f"個体一覧フィルタ: event_id={ev.get('id')}, event_number={ev.get('event_number')}, event_date={event_date}, ev_classification_value={ev_classification_value}, target_classification_value={classification_value}")
            if ev_classification_value == classification_value:
                cow_auto_id = ev.get('cow_auto_id')
                if cow_auto_id:
                    matching_cow_auto_ids.add(cow_auto_id)
        
        if not matching_cow_auto_ids:
            self.add_message(role="system", text="該当する個体がありません。")
            return
        
        # 個体リストのタイトルを生成
        title = f"個体一覧: {event_name}（{classification}={classification_value}）"
        
        # 個体リストを表示
        self._show_aggregate_cow_list(list(matching_cow_auto_ids), title, {})
    
    def _open_cow_card_for_list_item(self, item: str) -> bool:
        """
        リスト結果の行（treeview item）から個体カードを開く。
        values の ID 列 → タグ＋table_original_rows → 行インデックス＋table_original_rows の順で試す。
        
        Returns:
            個体カードを開けた場合 True、開けなかった場合 False
        """
        if not item or not hasattr(self, 'result_treeview') or not self.result_treeview.exists(item):
            return False
        tree = self.result_treeview
        values = tree.item(item, 'values')
        columns = list(tree['columns'])
        # まず表示されている ID/個体ID 列から取得（単一列「リスト: ID」で確実に開く）
        for id_col in ('ID', '個体ID'):
            if id_col in columns:
                id_idx = columns.index(id_col)
                if id_idx < len(values) and values[id_idx] not in (None, ''):
                    try:
                        cow_id = str(values[id_idx]).strip()
                        if cow_id.isdigit() or (cow_id and cow_id.isalnum()):
                            self._jump_to_cow_card(cow_id.zfill(4) if cow_id.isdigit() else cow_id)
                            return True
                    except Exception:
                        pass
                    break
        # タグから table_original_rows の行を特定
        tags = tree.item(item, 'tags')
        if tags and hasattr(self, 'table_original_rows') and self.table_original_rows:
            try:
                tag_str = tags[0] if isinstance(tags[0], str) else str(tags[0])
                row_index_str = tag_str.replace("row_", "").strip()
                if row_index_str.isdigit():
                    row_index = int(row_index_str)
                    if 0 <= row_index < len(self.table_original_rows):
                        row_data = self.table_original_rows[row_index]
                        auto_id = row_data.get('auto_id')
                        cow_id = row_data.get('ID') or row_data.get('個体ID')
                        if auto_id:
                            self._show_cow_card_view(auto_id)
                            return True
                        if cow_id not in (None, ''):
                            self._jump_to_cow_card(str(cow_id).zfill(4))
                            return True
            except (ValueError, IndexError, KeyError):
                pass
        # 行インデックスで table_original_rows から取得
        if hasattr(self, 'table_original_rows') and self.table_original_rows:
            children = tree.get_children()
            try:
                row_index = children.index(item)
                if 0 <= row_index < len(self.table_original_rows):
                    row_data = self.table_original_rows[row_index]
                    auto_id = row_data.get('auto_id')
                    cow_id = row_data.get('ID') or row_data.get('個体ID')
                    if auto_id:
                        self._show_cow_card_view(auto_id)
                        return True
                    if cow_id not in (None, ''):
                        self._jump_to_cow_card(str(cow_id).zfill(4))
                        return True
            except (ValueError, IndexError):
                pass
        return False
    
    def _on_list_row_double_click(self, event):
        """リスト結果の行をダブルクリックした時の処理（個体カードを開く）"""
        item = self.result_treeview.identify_row(event.y)
        if not item:
            return
        self.result_treeview.selection_set(item)
        self.result_treeview.focus(item)
        self._open_cow_card_for_list_item(item)
    
    def _on_list_row_click(self, event):
        """リスト結果の行をクリックした時の処理（セルクリック時のみ個体カードを開く）"""
        region = self.result_treeview.identify_region(event.x, event.y)
        # ヘッダー以外（セルまたはその他）のクリックで個体カードを開く
        if region == "heading":
            return
        item = self.result_treeview.identify_row(event.y)
        if not item:
            return
        self.result_treeview.selection_set(item)
        self.result_treeview.focus(item)
        self._open_cow_card_for_list_item(item)
    
    def _get_sort_key(self, value: Any) -> Any:
        """
        ソート用のキーを取得（数値と文字列が混在する列でも比較可能なタプルを返す）
        
        Args:
            value: ソート対象の値
        
        Returns:
            ソートキー (型順序, 値)。型順序により数値同士・文字列同士で比較し、str/float の直接比較を避ける。
        """
        if value is None:
            return (0, float('-inf'))  # None は先頭に
        
        # 数値として扱えるか試す
        if isinstance(value, (int, float)):
            return (0, float(value))
        
        value_str = str(value).strip()
        if value_str == "":
            return (0, float('-inf'))
        
        # 日付形式（YYYY-MM-DD）かチェック → 文字列としてソート
        if re.match(r'^\d{4}-\d{2}-\d{2}$', value_str):
            return (1, value_str)
        
        # 数値として解析できるか試す
        try:
            if '.' in value_str:
                return (0, float(value_str))
            else:
                return (0, int(value_str))
        except (ValueError, TypeError):
            return (1, value_str)
    
    def _generate_list_sql(self, columns: List[str], conditions: Dict[str, str] = None) -> Optional[str]:
        """
        リストコマンド用のSQLを生成
        
        Args:
            columns: 表示する項目名のリスト
            conditions: 条件の辞書 {項目名: 条件値}
        
        Returns:
            SQL文（生成できない場合はNone）
        """
        if conditions is None:
            conditions = {}
        
        # item_dictionary.jsonから動的にマッピングを生成
        column_mapping = {}  # 元の項目名 -> SQLカラム
        column_mapping_lower = {}  # 小文字の項目名 -> 元の項目名（大文字小文字を区別しない検索用）
        
        # 既存の固定マッピング
        fixed_mappings = {
            "ID": "c.cow_id",
            "個体ID": "c.cow_id",
            "産次": "c.lact",
            "繁殖コード": "c.rc",
            "繁殖区分": "c.rc",
            "RC": "c.rc",
            "分娩日": "e.event_date",
            "検診日": "e.event_date",
            "検診内容": "e.note"
        }
        column_mapping.update(fixed_mappings)
        for key, value in fixed_mappings.items():
            column_mapping_lower[key.lower()] = key
        
        # item_dictionary.jsonからマッピングを追加
        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                    item_dict = json.load(f)
                for item_key, item_def in item_dict.items():
                    origin = item_def.get("origin") or item_def.get("type", "")
                    display_name = item_def.get("display_name", "")
                    
                    # core項目で、formulaがcow.get('xxx')の形式の場合
                    if origin == "core":
                        formula = item_def.get("formula", "")
                        if formula.startswith("cow.get('") and formula.endswith("')"):
                            # cow.get('cow_id') -> cow_id
                            db_column = formula[9:-2]  # 'cow.get(' と ')' を除去
                            # 項目キーと表示名の両方をマッピング
                            column_mapping[item_key] = f"c.{db_column}"
                            column_mapping_lower[item_key.lower()] = item_key
                            if display_name:
                                column_mapping[display_name] = f"c.{db_column}"
                                column_mapping_lower[display_name.lower()] = display_name
            except Exception as e:
                logging.warning(f"item_dictionary.json読み込みエラー（SQL生成時）: {e}")
        
        # SELECT句を生成（大文字小文字を区別しない検索）
        select_parts = []
        # auto_idも取得（計算項目用）
        select_parts.append("c.auto_id AS auto_id")
        uses_event_table = False  # eventテーブルのカラムを使用しているか
        for col in columns:
            # 大文字小文字を区別しない検索
            col_lower = col.lower()
            if col_lower in column_mapping_lower:
                original_key = column_mapping_lower[col_lower]
                sql_col = column_mapping[original_key]
                select_parts.append(f"{sql_col} AS {original_key}")
                # eventテーブルのカラムを使用しているかチェック
                if sql_col.startswith("e."):
                    uses_event_table = True
            else:
                # マッピングにない場合はエラー
                return None
        
        if not select_parts:
            return None
        
        # WHERE句を生成（大文字小文字を区別しない検索）
        where_parts = []
        # eventテーブルのカラムを使用している場合のみ、deletedをチェック
        if uses_event_table:
            where_parts.append("(e.deleted = 0 OR e.deleted IS NULL)")
        
        for col_name, condition_value in conditions.items():
            # 大文字小文字を区別しない検索
            col_name_lower = col_name.lower()
            if col_name_lower not in column_mapping_lower:
                continue
            
            original_col_name = column_mapping_lower[col_name_lower]
            sql_col = column_mapping[original_col_name]
            # 条件でeventテーブルのカラムを使用している場合もフラグを立てる
            if sql_col.startswith("e."):
                uses_event_table = True
            
            # 条件値の解釈（比較演算子を含む場合を処理）
            import re
            # 条件値は既に「>30」のような形式で保存されている
            condition_value_clean = condition_value.strip()
            # 比較演算子パターン: <, >, <=, >=, =, !=
            operator_match = re.match(r'^([<>=!]+)(.+)$', condition_value_clean)
            if operator_match:
                operator = operator_match.group(1)
                value_str = operator_match.group(2).strip()
                # 数値かどうかチェック
                try:
                    num_value = float(value_str)
                    where_parts.append(f"{sql_col} {operator} {num_value}")
                except ValueError:
                    # 文字列の場合
                    where_parts.append(f"{sql_col} {operator} '{value_str}'")
            elif original_col_name == "産次":
                if condition_value == "初産":
                    where_parts.append(f"{sql_col} = 1")
                elif condition_value == "経産":
                    where_parts.append(f"{sql_col} > 1")
                elif condition_value.isdigit():
                    where_parts.append(f"{sql_col} = {condition_value}")
                else:
                    # その他の場合は文字列として扱う
                    where_parts.append(f"{sql_col} = '{condition_value}'")
            else:
                # その他の項目は文字列として扱う
                if condition_value.isdigit():
                    where_parts.append(f"{sql_col} = {condition_value}")
                else:
                    where_parts.append(f"{sql_col} = '{condition_value}'")
        
        # SQLを生成
        where_clause = " AND ".join(where_parts) if where_parts else "1=1"
        # eventテーブルを使用する場合のみJOIN
        if uses_event_table:
            join_clause = "LEFT JOIN event e ON c.auto_id = e.cow_auto_id AND e.deleted = 0"
        else:
            join_clause = ""
        
        sql = f"""
        SELECT {', '.join(select_parts)}
        FROM cow c
        {join_clause}
        WHERE {where_clause}
        ORDER BY c.cow_id
        LIMIT 100
        """
        
        return sql.strip()
    
    def _execute_query_router(self, query: str):
        """
        新しいQueryRouterシステムでクエリを処理
        
        Args:
            query: ユーザー入力
        """
        # 分析処理中の場合は新規分析を受け付けない
        if self.analysis_running:
            return
        
        # 分析ジョブを開始
        self.current_job_id += 1
        job_id = self.current_job_id
        self.analysis_running = True
        logging.info(f"[QueryRouter] クエリ処理開始: job_id={job_id}, query=\"{query}\"")
        
        # バックグラウンドスレッドで処理
        thread = threading.Thread(
            target=self._run_query_router_in_thread,
            args=(query, job_id),
            daemon=True
        )
        thread.start()
    
    def _run_query_router_in_thread(self, query: str, job_id: int):
        """
        バックグラウンドスレッドでQueryRouter処理を実行
        
        Args:
            query: ユーザー入力
            job_id: ジョブID
        """
        try:
            # QueryRouterでDSLを取得（UI指定期間を渡す）
            ui_period = None
            if self.selected_period.get("source") == "ui":
                ui_period = self.selected_period
            
            dsl_query = self.query_router.route(query, ui_period=ui_period)
            
            # AmbiguousQueryの場合は候補選択ダイアログを表示
            if QUERY_V2_AVAILABLE and isinstance(dsl_query, AmbiguousQuery):
                # メインスレッドで候補選択ダイアログを表示
                self.root.after(0, lambda: self._handle_ambiguous_query(dsl_query, query, ui_period, job_id))
                return
            
            # 使用された期間を取得（DSLから）
            used_period = None
            used_period_source = None
            if hasattr(dsl_query, 'start') and hasattr(dsl_query, 'end'):
                if dsl_query.start and dsl_query.end:
                    used_period = {
                        "start": dsl_query.start,
                        "end": dsl_query.end
                    }
                    used_period_source = ui_period.get("source") if ui_period else "text"
            elif hasattr(dsl_query, 'as_of_date'):
                if dsl_query.as_of_date:
                    used_period = {
                        "start": None,
                        "end": dsl_query.as_of_date
                    }
                    used_period_source = ui_period.get("source") if ui_period else "text"
            
            # UnhandledQueryの場合は未対応メッセージを返す
            result = None
            if QUERY_V2_AVAILABLE and isinstance(dsl_query, UnhandledQuery):
                if "個体ID" in dsl_query.reason:
                    # 個体IDの場合は従来の処理に委譲（ここでは処理しない）
                    result_text = None
                else:
                    # 未対応クエリの場合、正規化後文字列とマッチ情報をログ出力
                    normalized_input = normalize_user_input(query)
                    normalized_text = normalized_input["normalized"]
                    nospace_text = normalized_input["nospace"]
                    
                    # 正規化結果からマッチ情報を取得（デバッグ用）
                    normalized_result = self.query_router.normalizer.normalize_query(query)
                    matched_item = normalized_result.get("item")
                    matched_event = normalized_result.get("event")
                    matched_term = normalized_result.get("term")
                    
                    logging.warning(f"[QueryRouter] 未対応クエリ - "
                                  f"元のクエリ: '{query}', "
                                  f"正規化後: '{normalized_text}', "
                                  f"スペース除去: '{nospace_text}', "
                                  f"マッチ情報 - item={matched_item}, event={matched_event}, term={matched_term}, "
                                  f"理由: {dsl_query.reason}")
                    
                    # 未対応クエリの定型メッセージ
                    result_text = self._get_unhandled_message()
            elif QUERY_V2_AVAILABLE:
                # QueryExecutorでDB実行
                db_path = self.farm_path / "farm.db"
                executor = QueryExecutorV2(
                    db_path=db_path,
                    item_dictionary_path=self.item_dict_path,
                    event_dictionary_path=self.event_dict_path
                )
                result = executor.execute(dsl_query)
                
                # QueryRendererでテキスト化
                if self.query_renderer:
                    result_text = self.query_renderer.render(result, query_type=getattr(dsl_query, "type", None))
                else:
                    result_text = "QueryRendererが利用できません。"
            else:
                # QueryV2が利用できない場合は未対応メッセージ
                result_text = self._get_unhandled_message()
                
                # 集計期間を表示（安心設計）
                if used_period:
                    period_display = self._format_period_display(used_period, used_period_source)
                    if period_display:
                        result_text = period_display + "\n\n" + result_text
            
            # 結果をメインスレッドで表示（result_dataも渡す）
            # lambda式でresultをキャプチャするため、ローカル変数に保存
            result_for_callback = result
            # 期間情報も渡す
            period_info = {
                "period": used_period,
                "source": used_period_source
            }
            # クエリ情報も渡す
            self.root.after(0, lambda: self._handle_query_router_result(result_text, job_id, result_for_callback, period_info, query))
            
        except Exception as e:
            logging.error(f"[QueryRouter] エラー: {e}")
            import traceback
            traceback.print_exc()
            error_text = f"エラーが発生しました: {e}"
            self.root.after(0, lambda: self._handle_query_router_result(error_text, job_id, None))
    
    def _handle_ambiguous_query(self, ambiguous_query: AmbiguousQuery, 
                                original_query: str, ui_period: Optional[Dict[str, Any]], 
                                job_id: int):
        """
        曖昧なクエリ（複数候補）を処理（メインスレッドで実行）
        
        Args:
            ambiguous_query: AmbiguousQueryオブジェクト
            original_query: 元のクエリ文字列
            ui_period: UIで指定された期間
            job_id: ジョブID
        """
        try:
            # ジョブIDが古い場合は無視
            if job_id != self.current_job_id:
                logging.info(f"[QueryRouter] 古いジョブの結果を無視: job_id={job_id}")
                return
        
            candidates = ambiguous_query.candidates or []
            if not candidates:
                # 候補がない場合はエラー
                error_text = "候補が見つかりませんでした"
                self._handle_query_router_result(error_text, job_id, None, None, original_query)
                return
            
            # 候補選択ダイアログを表示
            selected_item_key = self._show_item_selection_dialog(
                original_query=ambiguous_query.original_query,
                candidates=candidates
            )
            
            if not selected_item_key:
                # キャンセルされた場合
                logging.info(f"[QueryRouter] 候補選択がキャンセルされました: クエリ='{original_query}'")
                return
            
            # 選択されたitem_keyで再ルーティング
            logging.info(f"[QueryRouter] 候補が選択されました: item_key={selected_item_key}, クエリ='{original_query}'")
            
            # バックグラウンドスレッドで再ルーティング
            thread = threading.Thread(
                target=self._run_query_router_in_thread_with_item,
                args=(original_query, selected_item_key, ui_period, job_id),
                daemon=True
            )
            thread.start()
            
        except Exception as e:
            logging.error(f"[QueryRouter] 曖昧クエリ処理エラー: {e}")
            import traceback
            traceback.print_exc()
            error_text = f"エラーが発生しました: {e}"
            self._handle_query_router_result(error_text, job_id, None, None, original_query)
    
    def _show_item_selection_dialog(self, original_query: str, 
                                    candidates: List[Dict[str, str]]) -> Optional[str]:
        """
        項目選択ダイアログを表示
        
        Args:
            original_query: 元のクエリ文字列
            candidates: 候補リスト [{"item_key": "...", "label": "...", "description": "..."}, ...]
        
        Returns:
            選択されたitem_key または None（キャンセル時）
        """
        # ダイアログウィンドウを作成（個体カードの項目選択と同じ仕様）
        dialog = tk.Toplevel(self.root)
        dialog.title("項目を選択")
        dialog.geometry("500x600")
        
        # カテゴリー名のマッピング
        category_names = {
            "REPRODUCTION": "繁殖",
            "DHI": "乳検",
            "GENOMIC": "ゲノム",
            "GENOME": "ゲノム",
            "HEALTH": "疾病",
            "CORE": "基本",
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
        ttk.Label(search_right_frame, text="検索:", font=('', 10)).pack(side=tk.RIGHT, padx=(5, 0))
        
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
        
        # 候補をカテゴリー別に整理
        item_dict = self.item_dictionary or {}
        category_items: Dict[str, List[tuple]] = {}
        for candidate in candidates:
            item_key = candidate.get("item_key") or ""
            if not item_key:
                continue
            label = candidate.get("label") or candidate.get("display_name") or candidate.get("name_jp")
            if not label:
                item_data = item_dict.get(item_key, {})
                label = item_data.get("label") or item_data.get("display_name") or item_data.get("name_jp") or item_key
            category = candidate.get("category")
            if not category:
                category = item_dict.get(item_key, {}).get("category")
            category = (category or "OTHERS").upper()
            if category not in category_items:
                category_items[category] = []
            category_items[category].append((item_key, label))
        
        # カテゴリーごとにソート
        for cat in category_items:
            category_items[cat].sort(key=lambda x: x[1])
        
        # カテゴリー順序（表示順）
        category_order = ["REPRODUCTION", "DHI", "GENOMIC", "GENOME", "HEALTH", "CORE", "USER", "OTHERS"]
        extra_categories = sorted([c for c in category_items.keys() if c not in category_order])
        category_order.extend(extra_categories)
        
        # Treeviewに項目を追加
        category_nodes: Dict[str, str] = {}
        all_items: List[tuple] = []  # (code, name, category, node_id)
        
        for cat in category_order:
            if cat in category_items and category_items[cat]:
                cat_name = category_names.get(cat, cat)
                cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                category_nodes[cat] = cat_node
                
                for code, name in category_items[cat]:
                    item_node = tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                    all_items.append((code, name, cat, item_node))
        
        # カテゴリータグのスタイル設定
        tree.tag_configure('category', font=('', 10, 'bold'), background='#E0E0E0')
        tree.tag_configure('item', font=('', 10))
        
        # 初期状態でカテゴリーを展開
        for cat_node in category_nodes.values():
            tree.item(cat_node, open=True)
        
        selected_code: Optional[str] = None
        
        def filter_tree(*args):
            """検索文字列でフィルタリング"""
            search_text = search_var.get().lower()
            
            # すべての項目を一旦非表示
            for code, name, cat, node_id in all_items:
                tree.detach(node_id)
            
            # 検索に一致する項目のみ表示
            visible_categories = set()
            if search_text:
                for code, name, cat, node_id in all_items:
                    if search_text in name.lower() or search_text in code.lower():
                        if cat in category_nodes:
                            tree.move(node_id, category_nodes[cat], 'end')
                            visible_categories.add(cat)
            else:
                # 検索が空の場合はすべて表示
                for code, name, cat, node_id in all_items:
                    if cat in category_nodes:
                        tree.move(node_id, category_nodes[cat], 'end')
                        visible_categories.add(cat)
            
            # カテゴリーノードの表示/非表示を制御
            for cat, cat_node in category_nodes.items():
                if cat in visible_categories:
                    if cat_node not in tree.get_children(''):
                        tree.move(cat_node, '', 'end')
                    tree.item(cat_node, open=True)
                else:
                    if cat_node in tree.get_children(''):
                        tree.detach(cat_node)
        
        search_var.trace('w', filter_tree)
        search_entry.bind('<KeyRelease>', lambda e: filter_tree())
        
        def on_ok():
            nonlocal selected_code
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            
            selected_node = sel[0]
            if tree.item(selected_node, 'tags') and tree.item(selected_node, 'tags')[0] == 'category':
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            
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
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return selected_code
    
    def _run_query_router_in_thread_with_item(self, query: str, selected_item_key: str,
                                              ui_period: Optional[Dict[str, Any]], job_id: int):
        """
        選択されたitem_keyで再ルーティング（バックグラウンドスレッド）
        
        Args:
            query: 元のクエリ文字列
            selected_item_key: 選択されたitem_key
            ui_period: UIで指定された期間
            job_id: ジョブID
        """
        try:
            # QueryRouterでDSLを取得（selected_item_keyを指定）
            dsl_query = self.query_router.route(query, ui_period=ui_period, 
                                                selected_item_key=selected_item_key)
            
            # 使用された期間を取得（DSLから）
            used_period = None
            used_period_source = None
            if hasattr(dsl_query, 'start') and hasattr(dsl_query, 'end'):
                if dsl_query.start and dsl_query.end:
                    used_period = {
                        "start": dsl_query.start,
                        "end": dsl_query.end
                    }
                    used_period_source = ui_period.get("source") if ui_period else "text"
            elif hasattr(dsl_query, 'as_of_date'):
                if dsl_query.as_of_date:
                    used_period = {
                        "start": None,
                        "end": dsl_query.as_of_date
                    }
                    used_period_source = ui_period.get("source") if ui_period else "text"
            
            # UnhandledQueryの場合は未対応メッセージを返す
            result = None
            if QUERY_V2_AVAILABLE and isinstance(dsl_query, UnhandledQuery):
                result_text = self._get_unhandled_message()
            elif QUERY_V2_AVAILABLE:
                # QueryExecutorでDB実行
                db_path = self.farm_path / "farm.db"
                executor = QueryExecutorV2(
                    db_path=db_path,
                    item_dictionary_path=self.item_dict_path,
                    event_dictionary_path=self.event_dict_path
                )
                result = executor.execute(dsl_query)
                
                # QueryRendererでテキスト化
                if self.query_renderer:
                    result_text = self.query_renderer.render(result, query_type=getattr(dsl_query, "type", None))
                else:
                    result_text = "QueryRendererが利用できません。"
            else:
                # QueryV2が利用できない場合は未対応メッセージ
                result_text = self._get_unhandled_message()
                
                # 集計期間を表示（安心設計）
                if used_period:
                    period_display = self._format_period_display(used_period, used_period_source)
                    if period_display:
                        result_text = period_display + "\n\n" + result_text
            
            # 結果をメインスレッドで表示
            result_for_callback = result
            period_info = {
                "period": used_period,
                "source": used_period_source
            }
            self.root.after(0, lambda: self._handle_query_router_result(result_text, job_id, result_for_callback, period_info, query))
        
        except Exception as e:
            logging.error(f"[QueryRouter] 再ルーティングエラー: {e}")
            import traceback
            traceback.print_exc()
            error_text = f"エラーが発生しました: {e}"
            self.root.after(0, lambda: self._handle_query_router_result(error_text, job_id, None, None, query))
    
    def _handle_query_router_result(self, result_text: Optional[str], job_id: int, 
                                   result_data: Optional[Dict[str, Any]] = None,
                                   period_info: Optional[Dict[str, Any]] = None,
                                   query: Optional[str] = None):
        """
        QueryRouter結果を処理（メインスレッドで実行）
        
        Args:
            result_text: 結果テキスト（Noneの場合は処理しない）
            job_id: ジョブID
            result_data: 結果データ（scatter/tableの場合に使用）
            period_info: 期間情報 {"period": {...}, "source": "ui"|"text"|"default"}
            query: 元のクエリ文字列
        """
        try:
            # ジョブIDが古い場合は無視
            if job_id != self.current_job_id:
                logging.info(f"[QueryRouter] 古いジョブの結果を無視: job_id={job_id}")
            return
        
            # scatter結果の場合は散布図を表示
            if result_data and result_data.get("kind") == "scatter":
                self._display_scatter_plot_from_result(result_data, command=query)
                # CSVデータもテキストとして表示
                if result_text:
                    self.add_message(role="ai", text=result_text)
            elif result_data and result_data.get("kind") == "table":
                # 表結果の場合はテキスト表示
                if result_text:
                    self.add_message(role="ai", text=result_text)
                # 表結果の場合もコマンド情報を保存
                if query:
                    self.current_table_command = query
                    if hasattr(self, 'table_command_text') and self.table_command_text.winfo_exists():
                        self._update_command_text(self.table_command_text, f"コマンド: {query}")
                    if hasattr(self, 'table_export_btn') and self.table_export_btn.winfo_exists():
                        self.table_export_btn.config(state=tk.NORMAL)
                    if hasattr(self, 'table_print_btn') and self.table_print_btn.winfo_exists():
                        self.table_print_btn.config(state=tk.NORMAL)
            elif result_text:
                self.add_message(role="ai", text=result_text)
        
        finally:
            # 分析ロックを解除
            if job_id == self.current_job_id:
                self.analysis_running = False
    
    def _get_unhandled_message(self) -> str:
        """未対応クエリの定型メッセージを返す"""
        return """この指示は現在のFALCON2では解釈できません。
対応例：
・平均 初回授精日数
・10月 分娩 頭数
・10月 分娩 個体 一覧
・縦軸 初回授精日数 横軸 DIM 散布図"""
    
    def _format_period_display(self, period: Dict[str, Any], source: Optional[str]) -> str:
        """
        集計期間の表示文字列を生成
        
        Args:
            period: 期間情報 {"start": date|str, "end": date|str}
            source: 期間のソース ("ui" | "text" | "default")
        
        Returns:
            表示文字列
        """
        if not period:
            return ""
        
        start = period.get("start")
        end = period.get("end")
        
        if not start and not end:
            return ""
        
        # 日付を文字列に変換
        if isinstance(start, date):
            start_str = start.strftime("%Y-%m-%d")
        elif isinstance(start, str):
            start_str = start
        else:
            start_str = "—"
        
        if isinstance(end, date):
            end_str = end.strftime("%Y-%m-%d")
        elif isinstance(end, str):
            end_str = end
        else:
            end_str = "—"
        
        # ソースに応じた表示
        if source == "ui":
            source_label = "UI指定"
        elif source == "text":
            source_label = "コマンド指定"
        else:
            source_label = "デフォルト"
        
        if start_str != "—" and end_str != "—":
            return f"集計期間：{start_str} ～ {end_str}（{source_label}）"
        elif end_str != "—":
            return f"基準日：{end_str}（{source_label}）"
        else:
            return ""
    
    def _execute_db_event_extraction(self, query: str):
        """
        DBイベント抽出を実行（即座に実行、AI送信なし）
        
        Args:
            query: ユーザー入力
        """
        # 分析処理中の場合は新規分析を受け付けない（メッセージは表示しない）
        if self.analysis_running:
            return
        
        # 分析ジョブを開始
        self.current_job_id += 1
        job_id = self.current_job_id
        self.analysis_running = True
        logging.info(f"[DB処理] イベント抽出開始: job_id={job_id}, query=\"{query}\"")
        
        # queryを保存（次の処理に進むため）
        self._pending_query_for_fallback = query
        
        # バックグラウンドスレッドで処理
        thread = threading.Thread(
            target=self._run_db_event_extraction_in_thread,
            args=(query, job_id),
            daemon=True
        )
        thread.start()
    
    def _run_db_event_extraction_in_thread(self, query: str, job_id: int):
        """
        バックグラウンドスレッドでDBイベント抽出を実行
        
        Args:
            query: ユーザー入力
            job_id: ジョブID
        """
        # ワーカースレッド内で新しいDBHandlerを生成（SQLiteスレッド違反を回避）
        db_path = self.farm_path / "farm.db"
        db = None
        try:
            from db.db_handler import DBHandler
            db = DBHandler(db_path)
            
            # イベント抽出処理を実行（新規DBHandlerを使用）
            result = self._process_event_extraction_query(query, db)
            logging.info(f"[DB処理] イベント抽出正常終了: job_id={job_id}, result={'あり' if result else 'なし'}")
            
            # 結果をメインスレッドで表示
            self.root.after(0, lambda: self._handle_db_event_extraction_result(query, result, job_id))
            
        except Exception as e:
            logging.error(f"[DB処理] イベント抽出エラー: job_id={job_id}, error={e}")
            import traceback
            traceback.print_exc()
            # エラー時も次の処理に進む（エラーメッセージは表示しない）
            self.root.after(0, lambda: self._handle_db_event_extraction_fallback(query, job_id))
        finally:
            # DBHandlerをクローズ
            if db is not None:
                try:
                    db.close()
                except:
                    pass
    
    def _handle_db_event_extraction_result(self, query: str, result: Optional[str], job_id: int):
        """
        DBイベント抽出結果を処理（メインスレッドで実行）
        
        Args:
            query: ユーザー入力
            result: 抽出結果
            job_id: ジョブID
        """
        try:
            if job_id != self.current_job_id:
                logging.info(f"[DB処理] 古いジョブの結果を無視: job_id={job_id}")
                return
            
            if result:
                # 成功時は結果を表示
                self.add_message(role="ai", text=result)
            else:
                # 失敗時はエラー表示せず、次の処理に進む
                logging.info(f"[DB処理] イベント抽出失敗（次の処理に進む）: job_id={job_id}, query=\"{query}\"")
                self._handle_db_event_extraction_fallback(query, job_id)
        
        finally:
            if job_id == self.current_job_id and result:
                # 成功時のみロック解除（失敗時は次の処理で解除）
                self.analysis_running = False
                logging.info(f"[DB処理] 分析ロック解除: job_id={job_id}")
    
    def _handle_db_event_extraction_fallback(self, query: str, job_id: int):
        """
        イベント抽出失敗時の次の処理に進む（エラー表示なし）
        
        Args:
            query: ユーザー入力
            job_id: ジョブID
        """
        try:
            if job_id != self.current_job_id:
                logging.info(f"[DB処理] 古いジョブのフォールバックを無視: job_id={job_id}")
                return
            
            logging.info(f"[DB処理] イベント抽出フォールバック開始: query=\"{query}\"")
            
            # 次の処理を判定して実行
            # a. 集計語（平均・合計・散布図・表）が含まれる → 項目集計処理
            if self._is_db_aggregation_query(query):
                logging.info(f"[DB処理] フォールバック: 項目集計処理に進む")
                self._execute_db_aggregation(query)
                return
            
            # b. 個体IDに該当 → 個体カード表示
            if query.isdigit():
                padded_id = query.zfill(4)
                logging.info(f"[DB処理] フォールバック: 個体カード表示に進む: {padded_id}")
                self._jump_to_cow_card(padded_id)
                return
            
            # c. 上記に該当しない → 未対応として返す（AIは使用しない）
            logging.info(f"[DB処理] フォールバック: 未対応のクエリ")
            self.add_message(role="system", text="未対応のクエリです。")
        
        finally:
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[DB処理] 分析ロック解除（フォールバック後）: job_id={job_id}")
    
    def _execute_db_aggregation(self, query: str):
        """
        DB集計を実行（即座に実行、AI送信なし）
        
        Args:
            query: ユーザー入力
        """
        # 分析処理中の場合は新規分析を受け付けない（メッセージは表示しない）
        if self.analysis_running:
            return
        
        # 分析ジョブを開始
        self.current_job_id += 1
        job_id = self.current_job_id
        self.analysis_running = True
        logging.info(f"[DB処理] DB集計開始: job_id={job_id}, query=\"{query}\"")
        
        # バックグラウンドスレッドで処理
        thread = threading.Thread(
            target=self._run_db_aggregation_in_thread,
            args=(query, job_id),
            daemon=True
        )
        thread.start()
    
    def _run_db_aggregation_in_thread(self, query: str, job_id: int):
        """
        バックグラウンドスレッドでDB集計を実行
        
        Args:
            query: ユーザー入力
            job_id: ジョブID
        """
        # ワーカースレッド内で新しいDBHandlerを生成（SQLiteスレッド違反を回避）
        db_path = self.farm_path / "farm.db"
        db = None
        try:
            from db.db_handler import DBHandler
            db = DBHandler(db_path)
            
            # 分析処理を実行（DB集計、新規DBHandlerを使用）
            calculated_data = self._process_analysis_query(query, db)
            logging.info(f"[DB処理] DB集計正常終了: job_id={job_id}")
            
            # 結果をメインスレッドで表示
            self.root.after(0, lambda: self._handle_db_aggregation_result(query, calculated_data, job_id))
            
        except Exception as e:
            logging.error(f"[DB処理] DB集計エラー: job_id={job_id}, error={e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: self._handle_db_processing_error(job_id, str(e)))
        finally:
            # DBHandlerをクローズ
            if db is not None:
                try:
                    db.close()
                except:
                    pass
    
    def _handle_db_aggregation_result(self, query: str, calculated_data: Any, job_id: int):
        """
        DB集計結果を処理（メインスレッドで実行）
        
        Args:
            query: ユーザー入力
            calculated_data: 計算済みデータ
            job_id: ジョブID
        """
        try:
            if job_id != self.current_job_id:
                logging.info(f"[DB処理] 古いジョブの結果を無視: job_id={job_id}")
                return
            
            if calculated_data is None:
                # 集計失敗時はエラー表示せず、次の処理に進む
                logging.info(f"[DB処理] DB集計失敗（次の処理に進む）: job_id={job_id}, query=\"{query}\"")
                self._handle_db_aggregation_fallback(query, job_id)
                return
            
            # 数値系クエリの場合は数値のみを直接表示
            if isinstance(calculated_data, str):
                self.add_message(role="ai", text=calculated_data)
                return
            
            # 散布図クエリの場合はグラフを表示
            if isinstance(calculated_data, dict) and 'x_data' in calculated_data:
                try:
                    self._display_scatter_plot(calculated_data)
                except Exception as e:
                    logging.error(f"[DB処理] 散布図表示エラー: {e}")
                    self.add_message(role="system", text=f"散布図の表示に失敗しました: {e}")
                return
            
            # 表レポートの場合はReportTableWindowを開く
            if isinstance(calculated_data, dict) and calculated_data.get('type') == 'table_report':
                try:
                    self._display_table_report(calculated_data)
                except Exception as e:
                    logging.error(f"[DB処理] 表レポート表示エラー: {e}")
                    self.add_message(role="system", text=f"表レポートの表示に失敗しました: {e}")
                return
        
        finally:
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[DB処理] 分析ロック解除: job_id={job_id}")
    
    def _handle_db_aggregation_fallback(self, query: str, job_id: int):
        """
        DB集計失敗時の次の処理に進む（エラー表示なし）
        
        Args:
            query: ユーザー入力
            job_id: ジョブID
        """
        try:
            if job_id != self.current_job_id:
                logging.info(f"[DB処理] 古いジョブのフォールバックを無視: job_id={job_id}")
                return
            
            logging.info(f"[DB処理] DB集計フォールバック開始: query=\"{query}\"")
            
            # 次の処理を判定して実行
            # a. 個体IDに該当 → 個体カード表示
            if query.isdigit():
                padded_id = query.zfill(4)
                logging.info(f"[DB処理] フォールバック: 個体カード表示に進む: {padded_id}")
                self._jump_to_cow_card(padded_id)
                return
            
            # b. 上記に該当しない → 未対応として返す（AIは使用しない）
            logging.info(f"[DB処理] フォールバック: 未対応のクエリ")
            self.add_message(role="system", text="未対応のクエリです。")
        
        finally:
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[DB処理] 分析ロック解除（フォールバック後）: job_id={job_id}")
    
    def _handle_db_processing_error(self, job_id: int, error_message: str):
        """
        DB処理エラーを処理（メインスレッドで実行）
        
        Args:
            job_id: ジョブID
            error_message: エラーメッセージ
        """
        try:
            if job_id != self.current_job_id:
                return
            
            # エラーはログにのみ出力（ユーザーには表示しない）
            logging.error(f"[DB処理] DB処理例外発生: job_id={job_id}, error={error_message}")
            
            # 保存されたクエリがあれば、次の処理に進む
            if hasattr(self, '_pending_query_for_fallback') and self._pending_query_for_fallback:
                query = self._pending_query_for_fallback
                delattr(self, '_pending_query_for_fallback')
                logging.info(f"[DB処理] エラー後フォールバック: query=\"{query}\"")
                self._handle_db_event_extraction_fallback(query, job_id)
            else:
                # クエリがない場合はロックのみ解除
                self.analysis_running = False
        
        except Exception as e:
            logging.error(f"[DB処理] エラー処理中に例外発生: {e}")
            if job_id == self.current_job_id:
                self.analysis_running = False
        finally:
            if job_id == self.current_job_id and self.analysis_running:
                self.analysis_running = False
                logging.info(f"[DB処理] 分析ロック解除（エラー時）: job_id={job_id}")
    
    def _try_db_processing(self, query: str):
        """
        DB処理を試行（AI送信禁止ルールに該当する場合）
        
        Args:
            query: ユーザー入力
        """
        # イベント抽出を試行
        if self._is_event_extraction_query(query):
            self._execute_db_event_extraction(query)
            return
        
        # DB集計を試行
        if self._is_db_aggregation_query(query):
            self._execute_db_aggregation(query)
            return
        
        # どちらも該当しない場合はエラーメッセージ
        self.add_message(role="system", text="該当する処理が見つかりませんでした。")
    
    def _is_event_extraction_query(self, text: str) -> bool:
        """
        イベント抽出系クエリかどうかを判定
        
        Args:
            text: ユーザー入力
        
        Returns:
            イベント抽出系クエリの場合はTrue
        """
        # イベント名のキーワード
        event_keywords = ["分娩", "AI", "ET", "妊娠鑑定", "授精", "乾乳", "乳検", "BCS", "WGT"]
        
        # 期間・日付のキーワード
        date_keywords = ["月", "期間", "日付", "日に", "年", "日"]
        
        # 個体・一覧のキーワード
        list_keywords = ["個体", "一覧", "教えて", "リスト", "見せて", "表示"]
        
        # イベント名と期間と個体/一覧が含まれている場合
        has_event = any(keyword in text for keyword in event_keywords)
        has_date = any(keyword in text for keyword in date_keywords)
        has_list = any(keyword in text for keyword in list_keywords)
        
        return has_event and (has_date or has_list)
    
    def _classify_query(self, text: str) -> str:
        """
        クエリを分類する
        
        Args:
            text: ユーザー入力
        
        Returns:
            'nav' (ナビ・ヘルプ系), 'analysis' (分析・集計系), 'consultation' (相談・解釈系)
        """
        # イベント抽出系クエリの場合は分析処理に分類
        if self._is_event_extraction_query(text):
            return 'analysis'
        
        # 分析・集計系のキーワード
        analysis_keywords = ["平均", "割合", "率", "分布", "表", "グラフ", "散布図", "集計", "統計", "合計", "最大", "最小", "何日", "何頭", "%", "％"]
        
        # ナビ・ヘルプ系のキーワード
        nav_keywords = ["どこで", "どうやって", "見る", "確認", "方法", "手順", "使い方"]
        
        text_lower = text.lower()
        
        # 分析・集計系のキーワードが含まれる場合
        if any(keyword in text for keyword in analysis_keywords):
            return 'analysis'
        
        # ナビ・ヘルプ系のキーワードが含まれる場合
        if any(keyword in text_lower for keyword in nav_keywords):
            return 'nav'
        
        # デフォルトは相談・解釈系
        return 'consultation'
    
    def _is_numeric_query(self, text: str) -> bool:
        """
        数値系クエリかどうかを判定
        
        Args:
            text: ユーザー入力
        
        Returns:
            数値系クエリの場合はTrue
        """
        numeric_keywords = ["平均", "合計", "最大", "最小", "何日", "何頭", "率", "%", "％"]
        return any(keyword in text for keyword in numeric_keywords)
    
    def _match_items_from_query(self, query: str) -> List[Dict[str, Any]]:
        """
        ユーザー入力から項目辞書のdisplay_nameを検索
        
        Args:
            query: ユーザー入力
        
        Returns:
            マッチした項目のリスト [{"display_name": "...", "item_key": "...", ...}, ...]
        """
        matched_items = []
        
        # 項目辞書の各項目をチェック（display_name および alias でマッチ）
        for item_key, item_def in self.item_dictionary.items():
            display_name = item_def.get('display_name', '')
            if not display_name:
                continue
            alias = item_def.get('alias', '')
            # display_name または alias が入力文に含まれていればマッチ
            if display_name in query or (alias and alias in query):
                matched_items.append({
                    'display_name': display_name,
                    'item_key': item_key,
                    'category': item_def.get('category', ''),
                    'origin': item_def.get('origin', ''),
                    'data_type': item_def.get('data_type', '')
                })
        
        return matched_items
    
    def _process_analysis_query(self, query: str, db: DBHandler) -> Optional[Union[str, Dict[str, Any]]]:
        """
        分析・集計系クエリを処理してデータを取得・計算
        
        Args:
            query: ユーザー入力
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            計算済みデータ（文字列または辞書）、処理できない場合はNone
            散布図クエリの場合は辞書を返す
        """
        # イベント抽出系クエリの処理（項目辞書に依存しない、最優先で処理）
        if self._is_event_extraction_query(query):
            return self._process_event_extraction_query(query, db)
        
        # 項目辞書からマッチする項目を検索
        matched_items = self._match_items_from_query(query)
        
        # 分析コンテキストを作成
        analysis_context = {
            'matched_items': matched_items,
            'item_keys': [item['item_key'] for item in matched_items]
        }
        
        # 項目が一つも見つからない場合はNoneを返す（エラーメッセージ表示）
        if not matched_items:
            return None
        
        query_lower = query.lower()
        
        # 散布図判定：項目が2つ以上かつ「×」「と」「vs」「対」などが含まれている場合
        if len(matched_items) >= 2 and any(keyword in query for keyword in ["×", "と", "vs", "対", "散布図"]):
            return self._process_scatter_plot_query(query, analysis_context, db)
        
        # 表レポート判定
        if self._is_table_report_query(query):
            return self._process_table_report_query(query, db)
        
        # 数値系クエリの判定（項目が1つの場合、または平均/合計などのキーワードがある場合）
        if self._is_numeric_query(query):
            # 平均分娩後日数（DIM）
            if "平均" in query and any(item['item_key'] == 'DIM' for item in matched_items):
                return self._calculate_average_dim(db)
            
            # 平均分娩間隔
            if "平均" in query and any(item['item_key'] == 'CINT' for item in matched_items):
                return self._calculate_average_calving_interval(db)
            
            # 産次別受胎率
            if "産次" in query and ("受胎率" in query or "受胎" in query):
                return self._calculate_conception_rate_by_lact(db)
        
        # その他の分析クエリは未実装としてNoneを返す
        # ただし、項目がマッチしている場合はエラーメッセージを表示しない
        return None
    
    def _is_table_report_query(self, query: str) -> bool:
        """
        表レポートクエリかどうかを判定
        
        Args:
            query: ユーザー入力
        
        Returns:
            表レポートクエリの場合はTrue
        """
        table_keywords = ["表", "一覧", "クロス", "産次×", "月別", "集計"]
        return any(keyword in query for keyword in table_keywords)
    
    def _process_event_extraction_query(self, query: str, db: DBHandler) -> Optional[Union[str, Dict[str, Any]]]:
        """
        イベント抽出系クエリを処理（DBからデータを抽出）
        
        Args:
            query: ユーザー入力
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            抽出結果（文字列または辞書）、処理できない場合はNone
        """
        try:
            # イベント名を特定
            event_number = None
            event_name = None
            
            if "分娩" in query:
                event_number = self.rule_engine.EVENT_CALV
                event_name = "分娩"
            elif "AI" in query or "授精" in query:
                event_number = self.rule_engine.EVENT_AI
                event_name = "AI"
            elif "ET" in query or "胚移植" in query:
                event_number = self.rule_engine.EVENT_ET
                event_name = "ET"
            elif "妊娠鑑定" in query or "妊鑑" in query:
                # プラスを優先
                if "マイナス" in query or "陰性" in query:
                    event_number = self.rule_engine.EVENT_PDN
                    event_name = "妊娠鑑定マイナス"
                else:
                    event_number = self.rule_engine.EVENT_PDP
                    event_name = "妊娠鑑定プラス"
            
            if event_number is None:
                logging.warning(f"イベント抽出クエリでイベント名が特定できません: {query}")
                return None
            
            # 期間を抽出
            import re
            year = None
            month = None
            
            # 年の抽出（例：2024年、2024）
            year_match = re.search(r'(\d{4})年?', query)
            if year_match:
                year = int(year_match.group(1))
            else:
                year = datetime.now().year  # デフォルトは現在年
            
            # 月の抽出（例：10月、10）
            month_match = re.search(r'(\d{1,2})月', query)
            if month_match:
                month = int(month_match.group(1))
            
            # 期間の開始日・終了日を計算
            if month:
                start_date = datetime(year, month, 1)
                if month == 12:
                    end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
                else:
                    end_date = datetime(year, month + 1, 1) - timedelta(days=1)
            else:
                # 月が指定されていない場合は全期間
                start_date = None
                end_date = None
            
            # DBからイベントを抽出（引数で受け取ったDBHandlerを使用）
            cows = db.get_all_cows()
            results = []
            
            for cow in cows:
                events = self.db.get_events_by_cow(cow.get('auto_id'), include_deleted=False)
                
                for event in events:
                    if event.get('event_number') != event_number:
                        continue
                    
                    event_date_str = event.get('event_date')
                    if not event_date_str:
                        continue
                    
                    try:
                        event_date = datetime.strptime(event_date_str, '%Y-%m-%d')
                        
                        # 期間でフィルタ
                        if start_date and end_date:
                            if not (start_date <= event_date <= end_date):
                                continue
                        elif start_date:
                            if event_date < start_date:
                                continue
                        elif end_date:
                            if event_date > end_date:
                                continue
                        
                        results.append({
                            'cow_id': cow.get('cow_id', ''),
                            'jpn10': cow.get('jpn10', ''),
                            'event_date': event_date_str,
                            'cow_auto_id': cow.get('auto_id')
                        })
                    except:
                        continue
            
            if not results:
                # データがない場合も簡潔に表示
                if start_date and end_date:
                    period_str = f"{start_date.strftime('%Y-%m-%d')} ～ {end_date.strftime('%Y-%m-%d')}"
                elif month:
                    period_str = f"{year}-{month:02d}-01 ～ {year}-{month:02d}-{end_date.day:02d}" if month and end_date else f"{year}-{month:02d}"
                else:
                    period_str = f"{year}"
                return f"{event_name}頭数：0 頭\n期間：{period_str}"
            
            # 日付順にソート
            results.sort(key=lambda x: (x['event_date'], x['cow_id']))
            
            # 経産牛・未経産牛をカウント
            parous_count = 0
            nulliparous_count = 0
            for result in results:
                cow_auto_id = result['cow_auto_id']
                cow = next((c for c in cows if c.get('auto_id') == cow_auto_id), None)
                if cow:
                    lact = cow.get('lact', 0)
                    if lact > 0:
                        parous_count += 1
                    else:
                        nulliparous_count += 1
            
            # 期間文字列を生成（YYYY-MM-DD形式）
            if start_date and end_date:
                period_str = f"{start_date.strftime('%Y-%m-%d')} ～ {end_date.strftime('%Y-%m-%d')}"
            elif month:
                # 月の終了日を計算
                if month == 12:
                    month_end = datetime(year + 1, 1, 1) - timedelta(days=1)
                else:
                    month_end = datetime(year, month + 1, 1) - timedelta(days=1)
                period_str = f"{year}-{month:02d}-01 ～ {month_end.strftime('%Y-%m-%d')}"
            else:
                period_str = f"{year}-01-01 ～ {year}-12-31"
            
            # 結果を簡潔にフォーマット（説明文なし）
            lines = []
            lines.append(f"{event_name}頭数：{len(results)} 頭")
            lines.append(f"対象：経産牛 {parous_count} 頭 / 未経産牛 {nulliparous_count} 頭")
            lines.append(f"期間：{period_str}")
            
            return "\n".join(lines)
            
        except Exception as e:
            logging.error(f"イベント抽出処理エラー: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _process_table_report_query(self, query: str, db: DBHandler) -> Optional[Dict[str, Any]]:
        """
        表レポートクエリを処理
        
        Args:
            query: ユーザー入力
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            表レポートデータの辞書、処理できない場合はNone
        """
        query_lower = query.lower()
        
        # 産次×月別分娩頭数
        if ("産次" in query or "産次×" in query) and ("月別" in query or "月" in query) and ("分娩" in query or "分娩頭数" in query):
            return self._calculate_calving_by_lact_and_month(db)
        
        # その他の表レポートは未対応
        logging.warning(f"未対応の表レポートクエリ: {query}")
        return None
    
    def _calculate_calving_by_lact_and_month(self, db: DBHandler) -> Optional[Dict[str, Any]]:
        """
        産次×月別分娩頭数を計算
        
        Args:
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        
        Returns:
            表レポートデータの辞書
        """
        try:
            cows = db.get_all_cows()
            today = datetime.now()
            
            # 最大産次を取得
            max_lact = 0
            for cow in cows:
                lact = cow.get('lact', 0)
                if lact > max_lact:
                    max_lact = lact
            
            # 集計用辞書 {産次: {月: 頭数}}
            lact_month_counts = {}
            for lact in range(1, max_lact + 1):
                lact_month_counts[lact] = {month: 0 for month in range(1, 13)}
            
            # 分娩イベントを集計
            for cow in cows:
                events = db.get_events_by_cow(cow.get('auto_id'), include_deleted=False)
                calving_events = [e for e in events if e.get('event_number') == self.rule_engine.EVENT_CALV]
                
                # 日付順にソート
                calving_events_sorted = sorted(
                    calving_events,
                    key=lambda e: (e.get('event_date', ''), e.get('id', 0))
                )
                
                # 各分娩イベントを処理（baseline_calvingフラグがないもののみ）
                lact_counter = 0
                for event in calving_events_sorted:
                    # baseline_calvingフラグをチェック
                    json_data = event.get('json_data') or {}
                    if isinstance(json_data, str):
                        try:
                            json_data = json.loads(json_data)
                        except:
                            json_data = {}
                    
                    # baseline_calvingフラグがある場合はスキップ
                    if json_data.get('baseline_calving', False):
                        continue
                    
                    lact_counter += 1
                    event_date_str = event.get('event_date')
                    if not event_date_str:
                        continue
                    
                    try:
                        event_date = datetime.strptime(event_date_str, '%Y-%m-%d')
                        month = event_date.month
                        
                        if lact_counter <= max_lact:
                            lact_month_counts[lact_counter][month] += 1
                    except:
                        continue
            
            # 表データを生成
            columns = ["産次", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
            rows = []
            
            for lact in sorted(lact_month_counts.keys()):
                row = [lact]
                for month in range(1, 13):
                    row.append(lact_month_counts[lact][month])
                rows.append(row)
            
            # 期間を計算（全データの範囲）
            all_calving_dates = []
            for cow in cows:
                events = db.get_events_by_cow(cow.get('auto_id'), include_deleted=False)
                for event in events:
                    if event.get('event_number') == self.rule_engine.EVENT_CALV:
                        event_date_str = event.get('event_date')
                        if event_date_str:
                            try:
                                event_date = datetime.strptime(event_date_str, '%Y-%m-%d')
                                all_calving_dates.append(event_date)
                            except:
                                pass
            
            if all_calving_dates:
                min_date = min(all_calving_dates)
                max_date = max(all_calving_dates)
                period = f"{min_date.strftime('%Y-%m-%d')} ～ {max_date.strftime('%Y-%m-%d')}"
            else:
                period = f"{today.strftime('%Y-%m-%d')} ～ {today.strftime('%Y-%m-%d')}"
            
            return {
                'type': 'table_report',
                'title': '産次×月別 分娩頭数',
                'columns': columns,
                'rows': rows,
                'conditions': f'全牛 {len(cows)} 頭',
                'period': period
            }
            
        except Exception as e:
            logging.error(f"産次×月別分娩頭数計算エラー: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _calculate_average_dim(self, db: DBHandler) -> str:
        """
        平均分娩後日数（DIM）を計算
        
        Args:
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        """
        try:
            cows = db.get_all_cows()
            dim_values = []
            
            for cow in cows:
                # 搾乳中（RC=2: Fresh）の牛のみ対象
                if cow.get('rc') != 2:  # RC_FRESH
                    continue
                
                clvd = cow.get('clvd')
                if not clvd:
                    continue
                
                # DIMを計算
                try:
                    clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
                    today = datetime.now()
                    dim = (today - clvd_date).days
                    if dim >= 0:
                        dim_values.append(dim)
                except:
                    continue
            
            if not dim_values:
                return None
            
            avg_dim = sum(dim_values) / len(dim_values)
            n_cows = len(dim_values)
            today = datetime.now()
            base_date_str = today.strftime('%Y-%m-%d')
            
            # 簡潔なフォーマット（説明文なし）
            lines = []
            lines.append(f"平均分娩後日数：{avg_dim:.1f} 日")
            lines.append(f"対象：搾乳牛 {n_cows} 頭")
            lines.append(f"基準日：{base_date_str}")
            
            return "\n".join(lines)
            
        except Exception as e:
            logging.error(f"平均DIM計算エラー: {e}")
            return None
    
    def _calculate_average_calving_interval(self, db: DBHandler) -> str:
        """
        平均分娩間隔を計算
        
        Args:
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        """
        try:
            cows = db.get_all_cows()
            intervals = []
            
            for cow in cows:
                events = db.get_events_by_cow(cow.get('auto_id'), include_deleted=False)
                calving_events = [e for e in events if e.get('event_number') == self.rule_engine.EVENT_CALV]
                calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
                
                # 2回以上の分娩がある場合のみ計算
                if len(calving_events) < 2:
                    continue
                
                # 連続する分娩間隔を計算
                for i in range(1, len(calving_events)):
                    try:
                        date1 = datetime.strptime(calving_events[i-1].get('event_date'), '%Y-%m-%d')
                        date2 = datetime.strptime(calving_events[i].get('event_date'), '%Y-%m-%d')
                        interval = (date2 - date1).days
                        if interval > 0:
                            intervals.append(interval)
                    except:
                        continue
            
            if not intervals:
                return None
            
            avg_interval = sum(intervals) / len(intervals)
            n_intervals = len(intervals)
            today = datetime.now()
            base_date_str = today.strftime('%Y-%m-%d')
            
            # 簡潔なフォーマット（説明文なし）
            lines = []
            lines.append(f"平均分娩間隔：{avg_interval:.1f} 日")
            lines.append(f"対象：{n_intervals} 分娩間隔")
            lines.append(f"基準日：{base_date_str}")
            
            return "\n".join(lines)
            
        except Exception as e:
            logging.error(f"平均分娩間隔計算エラー: {e}")
            return None
    
    def _calculate_conception_rate_by_lact(self, db: DBHandler) -> str:
        """
        産次別受胎率を計算
        
        Args:
            db: DBHandlerインスタンス（ワーカースレッド内で生成されたもの）
        """
        try:
            cows = db.get_all_cows()
            
            # 産次別の集計
            lact_stats = {}  # {lact: {'inseminated': count, 'conceived': count}}
            
            for cow in cows:
                events = db.get_events_by_cow(cow.get('auto_id'), include_deleted=False)
                
                # 産次ごとに処理
                lact = cow.get('lact') or 0
                if lact == 0:
                    continue
                
                # 該当産次の分娩日を取得
                calving_events = [e for e in events if e.get('event_number') == self.rule_engine.EVENT_CALV]
                calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
                
                if lact > len(calving_events):
                    continue
                
                # 該当産次の分娩日
                target_calving = calving_events[lact - 1] if lact > 0 else None
                if not target_calving:
                    continue
                
                calving_date = target_calving.get('event_date')
                
                # 該当産次での授精・受胎をカウント
                inseminated = False
                conceived = False
                
                for event in events:
                    event_date = event.get('event_date', '')
                    if event_date < calving_date:
                        continue
                    
                    event_number = event.get('event_number')
                    if event_number in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET]:
                        inseminated = True
                    elif event_number in [self.rule_engine.EVENT_PDP, self.rule_engine.EVENT_PDP2]:
                        conceived = True
                        break
                
                if inseminated:
                    if lact not in lact_stats:
                        lact_stats[lact] = {'inseminated': 0, 'conceived': 0}
                    lact_stats[lact]['inseminated'] += 1
                    if conceived:
                        lact_stats[lact]['conceived'] += 1
            
            if not lact_stats:
                return None
            
            # CSV形式で出力（説明文なし）
            csv_lines = ["産次,授精頭数,受胎頭数,受胎率(%)"]
            for lact in sorted(lact_stats.keys()):
                stats = lact_stats[lact]
                inseminated = stats['inseminated']
                conceived = stats['conceived']
                rate = (conceived / inseminated * 100) if inseminated > 0 else 0
                csv_lines.append(f"{lact},{inseminated},{conceived},{rate:.1f}")
            
            total_inseminated = sum(s['inseminated'] for s in lact_stats.values())
            total_conceived = sum(s['conceived'] for s in lact_stats.values())
            total_rate = (total_conceived / total_inseminated * 100) if total_inseminated > 0 else 0
            csv_lines.append(f"合計,{total_inseminated},{total_conceived},{total_rate:.1f}")
            
            # CSV形式のみを返す（説明文なし）
            return chr(10).join(csv_lines)
            
        except Exception as e:
            logging.error(f"産次別受胎率計算エラー: {e}")
            return None
    
    def _execute_batch_item_edit_command(self, raw_input: str):
        """項目一括変更コマンドを実行
        
        - 「項目=値：条件」形式（例: PEN=100：LACT=0）の場合は条件に合う個体を一括更新
        - 項目名のみの場合は項目一括編集ウィンドウを開く（従来どおりID入力で1頭ずつ編集）
        
        Args:
            raw_input: 項目名、または「項目=値：条件」
        """
        try:
            raw_stripped = raw_input.strip()
            if not raw_stripped:
                self.add_message(role="system", text="項目名を入力するか、「項目=値：条件」の形式で入力してください")
                return
            
            # 「項目=値：条件」形式か判定（= があり、かつ ： または : で区切られている）
            normalized_input = raw_stripped.replace('：', ':')
            is_condition_format = '=' in normalized_input and ':' in normalized_input
            
            if is_condition_format:
                parts = normalized_input.split(':', 1)
                update_part = (parts[0] or '').strip()
                condition_part = (parts[1] or '').strip()
                if not update_part or '=' not in update_part:
                    self.add_message(role="system", text="「項目=値：条件」の形式で入力してください。例: PEN=100：LACT=0")
                    return
                if not condition_part:
                    self.add_message(role="system", text="条件を入力してください。例: PEN=100：LACT=0")
                    return
                item_name = update_part.split('=', 1)[0].strip().upper()
                value_str = update_part.split('=', 1)[1].strip()
            else:
                item_name = raw_stripped.upper()
                value_str = ''
                condition_part = ''
            
            # item_dictionaryから項目情報を取得
            if item_name not in self.item_dictionary:
                self.add_message(
                    role="system",
                    text=f"項目 '{item_name}' が見つかりません。項目一覧を確認してください。"
                )
                return
            
            item_info = self.item_dictionary[item_name]
            item_type = item_info.get('type', '')
            origin = item_info.get('origin', item_type)
            display_name = item_info.get('display_name', item_name)
            
            # 計算・ソース項目は編集不可
            if item_type == 'calc' or origin == 'calc':
                formula = item_info.get('formula', '（計算式なし）')
                self.add_message(
                    role="system",
                    text=f"この項目は計算値です。編集できません。\n\n項目: {display_name} ({item_name})\n計算式: {formula}"
                )
                return
            if item_type == 'source' or origin == 'source':
                source = item_info.get('source', '')
                self.add_message(
                    role="system",
                    text=f"この項目はイベントデータから取得されます。編集するには対象イベントを編集してください。\n\n項目: {display_name} ({item_name})\nソース: {source}"
                )
                return
            
            # 条件付き一括更新
            if is_condition_format:
                self._execute_batch_item_edit_by_condition(
                    item_name=item_name,
                    item_info=item_info,
                    display_name=display_name,
                    value_str=value_str,
                    condition_part=condition_part,
                )
                return
            
            # 従来: 項目一括編集ウィンドウを開く
            self.add_message(
                role="system",
                text=f"項目一括編集を開始します: {display_name} ({item_name})"
            )
            def on_closed():
                pass
            edit_window = BatchItemEditWindow(
                parent=self.root,
                db_handler=self.db,
                item_name=item_name,
                item_info=item_info,
                on_closed=on_closed,
                formula_engine=self.formula_engine
            )
            edit_window.show()
            
        except Exception as e:
            logging.error(f"項目一括変更エラー: {e}")
            import traceback
            traceback.print_exc()
            self.add_message(role="system", text=f"エラーが発生しました: {e}")
    
    def _execute_batch_item_edit_by_condition(self, item_name: str, item_info: Dict[str, Any],
                                             display_name: str, value_str: str, condition_part: str):
        """条件に合う個体に対して項目を一括更新する（PEN=100：LACT=0 形式）"""
        db_column = BatchItemEditWindow.CORE_ITEM_TO_DB_COLUMN.get(item_name)
        origin = (item_info.get('origin') or item_info.get('type', '')).lower()
        is_custom = origin == 'custom' or (not db_column and origin != 'core')
        
        # 値のバリデーション・変換（BatchItemEditWindow と同様）
        date_items = ['BTHD', 'ENTR', 'CLVD']
        int_items = ['LACT', 'RC']
        if item_name in date_items and value_str:
            try:
                datetime.strptime(value_str, '%Y-%m-%d')
            except ValueError:
                self.add_message(role="system", text=f"{display_name}の形式が正しくありません (YYYY-MM-DD)")
                return
        is_int_item = item_name in int_items or (
            is_custom and (item_info.get('data_type') or '').lower() == 'int'
        )
        if is_int_item and value_str:
            try:
                int(value_str)
            except ValueError:
                self.add_message(role="system", text=f"{display_name}は数値である必要があります")
                return
        
        def convert_value(v: str):
            if not v:
                return None
            if item_name in int_items:
                return int(v)
            if is_custom and (item_info.get('data_type') or '').lower() == 'int':
                return int(v)
            return v
        
        converted_value = convert_value(value_str)
        
        batch_update = BatchUpdate(self.db, self.formula_engine)
        matching_cows = batch_update.get_matching_cows(condition_part)
        
        if not matching_cows:
            self.add_message(role="system", text=f"条件「{condition_part}」に該当する個体はありません。")
            return
        
        count = len(matching_cows)
        if not messagebox.askyesno(
            "確認",
            f"条件「{condition_part}」に該当する {count} 頭を {display_name} に「{value_str}」で更新します。\nよろしいですか？"
        ):
            return
        
        updated = 0
        for cow in matching_cows:
            auto_id = cow.get('auto_id')
            if not auto_id:
                continue
            try:
                if db_column:
                    self.db.update_cow(auto_id, {db_column: converted_value})
                else:
                    self.db.set_item_value(auto_id, item_name, converted_value)
                updated += 1
            except Exception as e:
                logging.error(f"一括更新エラー cow auto_id={auto_id}: {e}")
        
        self.add_message(role="system", text=f"条件付き一括更新を完了しました。{updated} 頭の {display_name} を「{value_str}」に更新しました。")
    
