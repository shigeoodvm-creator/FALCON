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

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    logging.warning("matplotlib がインストールされていません。散布図機能は使用できません。")

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.chatgpt_client import ChatGPTClient
from modules.query_router import QueryRouter
from modules.analysis_sql_templates import AnalysisSQLTemplates
from modules.analysis_exporter import AnalysisResultExporter
from modules.analysis_reports import AnalysisReports
from modules.query_router_v2 import QueryRouterV2
from modules.executor_v2 import ExecutorV2
from app.settings_manager import SettingsManager
from ui.cow_card import CowCard
from ui.event_input import EventInputWindow
from ui.analysis_ui import AnalysisUI
from datetime import date, timedelta
from calendar import monthrange


class MainWindow:
    """メインウィンドウ（サイドメニュー + メイン表示領域）"""
    
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
        
        # 設定管理（Phase E-2）
        self.settings_manager = SettingsManager(farm_path)
        
        # event_dictionary.json のパス（農場フォルダ側を優先）
        # 同期処理で作成されるため、農場フォルダ側のみを参照
        self.event_dict_path = farm_path / "event_dictionary.json"
        if not self.event_dict_path.exists():
            # 同期処理で作成されるはずだが、念のため警告
            print(f"警告: event_dictionary.json が見つかりません: {self.event_dict_path}")
            self.event_dict_path = None
        
        # item_dictionary.json のパス（農場フォルダ側を優先）
        # 同期処理で作成されるため、農場フォルダ側のみを参照
        self.item_dict_path = farm_path / "item_dictionary.json"
        if not self.item_dict_path.exists():
            # 同期処理で作成されるはずだが、念のため警告
            print(f"警告: item_dictionary.json が見つかりません: {self.item_dict_path}")
            self.item_dict_path = None
        
        # 現在選択中の牛
        self.current_cow_auto_id: Optional[int] = None
        self.current_cow_card: Optional[CowCard] = None
        
        # 現在表示中のView
        self.current_view: Optional[tk.Widget] = None
        self.current_view_type: Optional[Literal['chat', 'cow_card', 'list', 'report', 'analysis_ui']] = None
        
        # ビュー管理（再利用のため保持）
        self.views: Dict[str, tk.Widget] = {}
        
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
        
        # QueryRouter（最小構成）
        self.query_router = QueryRouter()
        
        # QueryRouterV2 / ExecutorV2（AnalysisUI用）
        # 初期化は遅延実行（メニューから開かれたときに初期化）
        self.query_router_v2: Optional[QueryRouterV2] = None
        self.executor_v2: Optional[ExecutorV2] = None
        self.analysis_ui: Optional[AnalysisUI] = None
        
        # 分析モードフラグ（Phase D - Step 2）
        self.analysis_mode: bool = False
        
        # SQLテンプレート管理（Phase D - Step 3）
        self.sql_templates = AnalysisSQLTemplates()
        
        # 分析結果エクスポーター（Phase D - Step 4）
        self.result_exporter = AnalysisResultExporter()
        
        # 定型レポート管理（Phase E-1）
        self.analysis_reports = AnalysisReports()
        
        # 出力形式フラグ（Phase D - Step 4）
        self.export_csv: bool = False
        self.export_excel: bool = False
        
        # コマンド入力状態管理（Phase E-2）
        self.command_has_focus: bool = False
        self.command_placeholder_shown: bool = True
        
        # ウィンドウサイズを設定
        self.root.title("FALCON2")
        self.root.geometry("1200x800")
        
        # UI作成
        self._create_widgets()
        
        # 初回起動時のガイド表示（Phase E-2）
        self._show_quick_guide_if_first_time()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインPanedWindow（左右分割）
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左カラム：サイドメニュー ==========
        menu_frame = ttk.Frame(main_paned, width=200)
        main_paned.add(menu_frame, weight=0)
        
        # メニューボタン
        menu_buttons = [
            ("個体カード", self._on_cow_card),
            ("繁殖検診", self._on_reproduction_checkup),
            ("イベント入力", self._on_event_input),
            ("分析", self._on_analysis_ui),
            ("データ出力", self._on_data_output),
            ("辞書・設定", self._on_dictionary_settings),
            ("農場管理", self._on_farm_management),
        ]
        
        for text, command in menu_buttons:
            btn = ttk.Button(menu_frame, text=text, command=command, width=20)
            btn.pack(pady=5, padx=10)
        
        # ========== 右カラム：メイン表示領域 ==========
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        # 右上：コマンド入力欄（Phase E-2）
        command_frame = ttk.Frame(right_frame)
        command_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # モード状態表示ラベル（Phase E-2）
        mode_frame = ttk.Frame(command_frame)
        mode_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        
        self.mode_label = ttk.Label(
            mode_frame,
            text="通常モード",
            foreground="gray",
            font=("", 9)
        )
        self.mode_label.pack(side=tk.LEFT)
        
        ttk.Label(command_frame, text="コマンド:").pack(side=tk.LEFT, padx=5)
        
        # コマンド入力欄（プレースホルダー付き、Phase E-2）
        self.command_entry = ttk.Entry(command_frame, width=50)
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.command_entry.bind('<Return>', self._on_command_enter)
        self.command_entry.bind('<KeyPress>', self._on_command_key_press)
        self.command_entry.bind('<FocusIn>', self._on_command_focus_in)
        self.command_entry.bind('<FocusOut>', self._on_command_focus_out)
        
        # プレースホルダーテキスト（Phase E-2）
        self.command_placeholder = "例）101 または レポート：月別分娩頭数 2024年 CSV または 分析：最近6か月で2産の受胎率が下がった理由"
        self._show_command_placeholder()
        
        command_btn = ttk.Button(command_frame, text="実行", command=self._on_command_execute)
        command_btn.pack(side=tk.LEFT, padx=5)
        
        # 入力補助ドロップダウン（Phase E-2）
        self.report_dropdown_frame = None
        self.report_dropdown_listbox = None
        
        # 分析モードチェックボックス（Phase D - Step 2）
        self.analysis_mode_var = tk.BooleanVar(value=False)
        analysis_checkbox = ttk.Checkbutton(
            command_frame,
            text="分析モード",
            variable=self.analysis_mode_var,
            command=self._on_analysis_mode_toggle
        )
        analysis_checkbox.pack(side=tk.LEFT, padx=5)
        
        # 出力形式チェックボックス（Phase D - Step 4）
        self.export_csv_var = tk.BooleanVar(value=False)
        csv_checkbox = ttk.Checkbutton(
            command_frame,
            text="CSV出力",
            variable=self.export_csv_var
        )
        csv_checkbox.pack(side=tk.LEFT, padx=5)
        
        self.export_excel_var = tk.BooleanVar(value=False)
        excel_checkbox = ttk.Checkbutton(
            command_frame,
            text="Excel出力",
            variable=self.export_excel_var
        )
        excel_checkbox.pack(side=tk.LEFT, padx=5)
        
        # コマンド入力欄の直下：期間指定エリア
        self._create_period_selector(right_frame)
        
        # コマンド入力欄の直下：最終日付表示
        latest_dates_frame = ttk.Frame(right_frame)
        latest_dates_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        self.latest_dates_label = ttk.Label(
            latest_dates_frame,
            text="最終分娩：—　最終AI：—　最終乳検：—　最終イベント：—",
            foreground="gray"
        )
        self.latest_dates_label.pack(side=tk.LEFT, padx=5)
        
        # 右側：メイン表示領域（ChatGPT / View切替）
        self.main_content_frame = ttk.Frame(right_frame)
        self.main_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 初期表示：ChatGPT画面
        self._show_chat_view()
    
    def _create_period_selector(self, parent: tk.Widget):
        """
        期間指定エリアを作成
        
        Args:
            parent: 親ウィジェット
        """
        period_frame = ttk.LabelFrame(parent, text="期間指定", padding=5)
        period_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        # 日付入力欄
        date_input_frame = ttk.Frame(period_frame)
        date_input_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(date_input_frame, text="開始日:").pack(side=tk.LEFT, padx=5)
        self.period_start_entry = ttk.Entry(date_input_frame, width=12)
        self.period_start_entry.pack(side=tk.LEFT, padx=5)
        self.period_start_entry.insert(0, "YYYY-MM-DD")
        self.period_start_entry.bind('<FocusIn>', lambda e: self._on_period_entry_focus_in(self.period_start_entry))
        self.period_start_entry.bind('<FocusOut>', lambda e: self._on_period_entry_focus_out(self.period_start_entry))
        
        ttk.Label(date_input_frame, text="終了日:").pack(side=tk.LEFT, padx=5)
        self.period_end_entry = ttk.Entry(date_input_frame, width=12)
        self.period_end_entry.pack(side=tk.LEFT, padx=5)
        self.period_end_entry.insert(0, "YYYY-MM-DD")
        self.period_end_entry.bind('<FocusIn>', lambda e: self._on_period_entry_focus_in(self.period_end_entry))
        self.period_end_entry.bind('<FocusOut>', lambda e: self._on_period_entry_focus_out(self.period_end_entry))
        
        # クイックボタン群
        quick_buttons_frame = ttk.Frame(period_frame)
        quick_buttons_frame.pack(fill=tk.X)
        
        quick_buttons = [
            ("今年", self._on_period_this_year),
            ("昨年", self._on_period_last_year),
            ("直近1年", self._on_period_last_1year),
            ("今月", self._on_period_this_month),
            ("先月", self._on_period_last_month),
            ("クリア", self._on_period_clear)
        ]
        
        for text, command in quick_buttons:
            btn = ttk.Button(quick_buttons_frame, text=text, command=command, width=10)
            btn.pack(side=tk.LEFT, padx=2)
    
    def _on_period_entry_focus_in(self, entry: ttk.Entry):
        """期間入力欄にフォーカスが入ったとき"""
        if entry.get() == "YYYY-MM-DD":
            entry.delete(0, tk.END)
    
    def _on_period_entry_focus_out(self, entry: ttk.Entry):
        """期間入力欄からフォーカスが外れたとき"""
        text = entry.get().strip()
        if not text:
            entry.insert(0, "YYYY-MM-DD")
        else:
            # 日付形式を検証
            try:
                date.fromisoformat(text)
                # 有効な日付の場合は selected_period を更新
                self._update_period_from_entries()
            except ValueError:
                # 無効な日付の場合は警告
                messagebox.showwarning("警告", f"無効な日付形式です: {text}\nYYYY-MM-DD 形式で入力してください。")
                entry.delete(0, tk.END)
                entry.insert(0, "YYYY-MM-DD")
    
    def _update_period_from_entries(self):
        """Entry から期間を読み取って selected_period を更新"""
        start_text = self.period_start_entry.get().strip()
        end_text = self.period_end_entry.get().strip()
        
        if start_text == "YYYY-MM-DD" or end_text == "YYYY-MM-DD":
            # どちらかが未入力の場合はクリア
            self.selected_period = {
                "start": None,
                "end": None,
                "source": None
            }
            return
        
        try:
            start_date = date.fromisoformat(start_text)
            end_date = date.fromisoformat(end_text)
            
            if start_date > end_date:
                messagebox.showwarning("警告", "開始日が終了日より後です。")
                return
            
            self.selected_period = {
                "start": start_date,
                "end": end_date,
                "source": "ui"
            }
        except ValueError:
            # 日付解析エラーは _on_period_entry_focus_out で処理済み
            pass
    
    def _on_period_this_year(self):
        """今年ボタンの動作"""
        today = date.today()
        start_date = date(today.year, 1, 1)
        end_date = date(today.year, 12, 31)
        self._set_period(start_date, end_date)
    
    def _on_period_last_year(self):
        """昨年ボタンの動作"""
        today = date.today()
        start_date = date(today.year - 1, 1, 1)
        end_date = date(today.year - 1, 12, 31)
        self._set_period(start_date, end_date)
    
    def _on_period_last_1year(self):
        """直近1年ボタンの動作"""
        today = date.today()
        start_date = today - timedelta(days=365)
        end_date = today
        self._set_period(start_date, end_date)
    
    def _on_period_this_month(self):
        """今月ボタンの動作"""
        today = date.today()
        start_date = date(today.year, today.month, 1)
        _, last_day = monthrange(today.year, today.month)
        end_date = date(today.year, today.month, last_day)
        self._set_period(start_date, end_date)
    
    def _on_period_last_month(self):
        """先月ボタンの動作"""
        today = date.today()
        if today.month == 1:
            last_month = 12
            last_year = today.year - 1
        else:
            last_month = today.month - 1
            last_year = today.year
        
        start_date = date(last_year, last_month, 1)
        _, last_day = monthrange(last_year, last_month)
        end_date = date(last_year, last_month, last_day)
        self._set_period(start_date, end_date)
    
    def _on_period_clear(self):
        """クリアボタンの動作"""
        self.selected_period = {
            "start": None,
            "end": None,
            "source": None
        }
        self.period_start_entry.delete(0, tk.END)
        self.period_start_entry.insert(0, "YYYY-MM-DD")
        self.period_end_entry.delete(0, tk.END)
        self.period_end_entry.insert(0, "YYYY-MM-DD")
    
    def _set_period(self, start_date: date, end_date: date):
        """
        期間を設定
        
        Args:
            start_date: 開始日
            end_date: 終了日
        """
        self.selected_period = {
            "start": start_date,
            "end": end_date,
            "source": "ui"
        }
        self.period_start_entry.delete(0, tk.END)
        self.period_start_entry.insert(0, start_date.strftime("%Y-%m-%d"))
        self.period_end_entry.delete(0, tk.END)
        self.period_end_entry.insert(0, end_date.strftime("%Y-%m-%d"))
        
        # 群全体の最終日付を計算して表示
        self._calculate_and_update_farm_latest_dates()
    
    def show_view(self, view_widget: tk.Widget, view_type: Literal['chat', 'cow_card', 'list', 'report', 'analysis_ui']):
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
        """個体カードメニューをクリック（個体一覧ウィンドウを開く）"""
        from ui.cow_list_window import CowListWindow
        
        # 個体一覧ウィンドウを開く
        list_window = CowListWindow(
            parent=self.root,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=self.item_dict_path
        )
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
        
        # EventInputWindow を生成して表示
        # 左メニューからの起動は常に cow_auto_id=None で汎用モード
        # 必ずID入力欄を表示状態で起動する
        event_input_window = EventInputWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 常にNone（汎用モード、ID入力から開始）
            on_saved=self._on_event_saved,
            farm_path=self.farm_path
        )
        event_input_window.show()
    
    def _on_analysis_ui(self):
        """分析UIメニューをクリック"""
        try:
            # QueryRouterV2 / ExecutorV2が初期化されていない場合は初期化
            if self.query_router_v2 is None or self.executor_v2 is None:
                self._init_analysis_ui_components()
            
            # 初期化に失敗した場合はエラー
            if self.query_router_v2 is None or self.executor_v2 is None:
                messagebox.showerror(
                    "エラー",
                    "分析UIの初期化に失敗しました。辞書ファイルを確認してください。"
                )
                return
            
            # AnalysisUIが既に存在する場合は再利用、なければ新規作成
            if self.analysis_ui is None:
                # AnalysisUIを新規作成
                self.analysis_ui = AnalysisUI(
                    parent=self.main_content_frame,
                    farm_path=self.farm_path,
                    query_router=self.query_router_v2,
                    executor=self.executor_v2,
                    on_execute=self._on_analysis_result
                )
            
            # AnalysisUIを表示
            self.show_view(self.analysis_ui.frame, 'analysis_ui')
        
        except Exception as e:
            logging.error(f"AnalysisUI表示エラー: {e}", exc_info=True)
            messagebox.showerror(
                "エラー",
                f"分析UIの表示に失敗しました: {str(e)}"
            )
    
    def _init_analysis_ui_components(self):
        """AnalysisUI用のコンポーネントを初期化"""
        try:
            # QueryRouterV2を初期化
            if self.query_router_v2 is None:
                self.query_router_v2 = QueryRouterV2(
                    item_dictionary_path=self.item_dict_path,
                    event_dictionary_path=self.event_dict_path
                )
            
            # ExecutorV2を初期化
            if self.executor_v2 is None:
                self.executor_v2 = ExecutorV2(
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    item_dictionary_path=self.item_dict_path,
                    event_dictionary_path=self.event_dict_path
                )
        except Exception as e:
            logging.error(f"AnalysisUIコンポーネントの初期化エラー: {e}", exc_info=True)
            messagebox.showerror(
                "エラー",
                f"分析UIの初期化に失敗しました: {str(e)}"
            )
    
    def _on_analysis_result(self, result: Dict[str, Any]):
        """AnalysisUIの実行結果を受け取るコールバック"""
        # 結果はAnalysisUI内で表示されるため、ここではログ出力のみ
        if result.get("success"):
            logging.info("分析が正常に完了しました")
        else:
            errors = result.get("errors", [])
            if errors:
                logging.warning(f"分析エラー: {errors}")
    
    def _on_data_output(self):
        """データ出力メニューをクリック"""
        print("データ出力メニューをクリック")
        # TODO: データ出力画面を実装
    
    def _on_dictionary_settings(self):
        """辞書設定メニューをクリック"""
        # 辞書設定選択ダイアログを表示
        dialog = tk.Toplevel(self.root)
        dialog.title("辞書・設定")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # タイトル
        title_label = ttk.Label(
            dialog,
            text="辞書・設定",
            font=("", 12, "bold")
        )
        title_label.pack(pady=20)
        
        # 説明
        desc_label = ttk.Label(
            dialog,
            text="操作を選択してください",
            font=("", 9)
        )
        desc_label.pack(pady=(0, 20))
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        # イベント辞書ボタン
        event_dict_btn = ttk.Button(
            button_frame,
            text="イベント辞書",
            command=lambda: self._on_event_dictionary(dialog),
            width=20
        )
        event_dict_btn.pack(pady=5, padx=20)
        
        # 項目辞書ボタン
        item_dict_btn = ttk.Button(
            button_frame,
            text="項目辞書",
            command=lambda: self._on_item_dictionary(dialog),
            width=20
        )
        item_dict_btn.pack(pady=5, padx=20)
        
        # 農場設定ボタン
        farm_settings_btn = ttk.Button(
            button_frame,
            text="農場設定",
            command=lambda: self._on_farm_settings(dialog),
            width=20
        )
        farm_settings_btn.pack(pady=5, padx=20)
        
        # キャンセルボタン
        cancel_btn = ttk.Button(
            button_frame,
            text="キャンセル",
            command=dialog.destroy,
            width=20
        )
        cancel_btn.pack(pady=10, padx=20)
        
        # ウィンドウを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_event_dictionary(self, parent_dialog):
        """イベント辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        from ui.event_dictionary_window import EventDictionaryWindow
        
        event_dict_window = EventDictionaryWindow(
            parent=self.root,
            event_dictionary_path=self.event_dict_path
        )
        event_dict_window.show()
    
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
    
    def _on_farm_management(self):
        """農場管理メニューをクリック"""
        print("農場管理メニューをクリック")
        # TODO: 農場管理画面を実装
    
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
        try:
            # 農場全体のイベントを取得
            events = self.db.get_all_events(include_deleted=False)
            
            # デバッグ: イベント数とイベント番号を確認
            logging.debug(f"[_calculate_and_update_farm_latest_dates] 取得したイベント数: {len(events)}")
            if events:
                event_numbers = [e.get('event_number') for e in events[:10]]  # 最初の10件
                logging.debug(f"[_calculate_and_update_farm_latest_dates] 最初の10件のイベント番号: {event_numbers}")
            
            # 最終分娩日を計算（EVENT_CALV = 202、baselineも含む）
            latest_calving = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_CALV])
            
            # 最終AI日を計算（AI/ET系イベント: 200, 201）
            latest_ai = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET])
            
            # 最終乳検日を計算（EVENT_MILK_TEST = 601）
            latest_milk = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_MILK_TEST])
            
            # 最終イベント日を計算（全イベント対象）
            latest_any = CowCard.get_latest_any_event_date(events)
            
            # デバッグ: 計算結果を確認
            logging.debug(f"[_calculate_and_update_farm_latest_dates] latest_calving={latest_calving}, latest_ai={latest_ai}, latest_milk={latest_milk}, latest_any={latest_any}")
            
            # 日付を表示用文字列に変換（Noneの場合は"—"）
            calving_str = latest_calving if latest_calving else "—"
            ai_str = latest_ai if latest_ai else "—"
            milk_str = latest_milk if latest_milk else "—"
            any_str = latest_any if latest_any else "—"
            
            # 表示文字列を作成
            display_text = f"最終分娩：{calving_str}　最終AI：{ai_str}　最終乳検：{milk_str}　最終イベント：{any_str}"
            
            # Labelを更新
            self.latest_dates_label.config(text=display_text)
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _calculate_and_update_farm_latest_dates で例外が発生しました: {e}")
            traceback.print_exc()
            # エラー時は"—"を表示
            display_text = "最終分娩：—　最終AI：—　最終乳検：—　最終イベント：—"
            self.latest_dates_label.config(text=display_text)
    
    def _on_command_enter(self, event):
        """コマンド入力欄でEnterキー押下時"""
        self._on_command_execute()
    
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
            "データ出力",
            "辞書・設定",
            "農場管理"
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
        elif text == "データ出力":
            self._on_data_output()
        elif text == "辞書・設定":
            self._on_dictionary_settings()
        elif text == "農場管理":
            self._on_farm_management()
    
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
    
    def _on_analysis_mode_toggle(self):
        """
        分析モードのトグル処理（Phase D - Step 2）
        """
        self.analysis_mode = self.analysis_mode_var.get()
        if self.analysis_mode:
            logging.info("分析モードが有効になりました")
        else:
            logging.info("分析モードが無効になりました")
    
    def classify_command(self, text: str) -> Dict[str, Any]:
        """
        コマンドを分類して適切なモードにルーティング
        
        【判定順序（厳守）】
        1. 数字のみ → 個体カードを開く
        2. 「レポート：」「報告：」 → 定型レポート強制
        3. 「分析：」「解説：」「理由：」 → 分析モード強制
        4. 定型レポート名に一致 → 定型レポート
        5. 「一覧」「個体」「牛」「誰」「抽出」「表示」を含む → 分析（抽出）
        6. 「なぜ」「理由」「原因」「傾向」「下がった」「高い」「低い」等を含む → 分析（解釈）
        7. どれにも該当しない場合 → ガイド文を表示
        
        Args:
            text: ユーザー入力
        
        Returns:
            {
                "mode": "cow_card" | "report" | "analysis" | "guide",
                "original_text": str,
                "clean_text": str (プレフィックス除去後),
                "report_info": Dict (reportモードの場合)
            }
        """
        text = text.strip()
        if not text:
            return {"mode": "guide", "original_text": text, "clean_text": text}
        
        # 1. 「リスト」で始まる → リストコマンド
        if text.startswith("リスト") or text.startswith("リスト　"):
            return {
                "mode": "list",
                "original_text": text,
                "clean_text": text
            }
        
        # 2. 数字のみ → 個体カードを開く
        if text.isdigit():
            return {
                "mode": "cow_card",
                "original_text": text,
                "clean_text": text
            }
        
        # 3. 「レポート：」「報告：」 → 定型レポート強制
        report_prefixes = ["レポート：", "報告："]
        for prefix in report_prefixes:
            if text.startswith(prefix):
                report_text = text[len(prefix):].strip()
                report_info = self._check_report_command(text)  # 元のテキストでチェック
                if report_info:
                    return {
                        "mode": "report",
                        "original_text": text,
                        "clean_text": report_text,
                        "report_info": report_info
                    }
                # プレフィックスはあるがレポートが見つからない場合もreportモードとして扱う
                return {
                    "mode": "report",
                    "original_text": text,
                    "clean_text": report_text,
                    "report_info": None
                }
        
        # 4. 「分析：」「解説：」「理由：」 → 分析モード強制
        analysis_keywords = ["分析：", "解説：", "理由："]
        for keyword in analysis_keywords:
            if text.startswith(keyword):
                clean_text = text[len(keyword):].strip()
                return {
                    "mode": "analysis",
                    "original_text": text,
                    "clean_text": clean_text
                }
        
        # 5. 定型レポート名に一致 → 定型レポート
        # 定型レポート名のリストを取得
        reports = self.analysis_reports.list_reports()
        for report_key, report in reports.items():
            display_name = report.get("display_name", "")
            # 表示名が含まれているかチェック（部分一致）
            if display_name in text or text in display_name:
                report_info = {
                    "report_key": report_key,
                    "report": report,
                    "input_text": text
                }
                return {
                    "mode": "report",
                    "original_text": text,
                    "clean_text": text,
                    "report_info": report_info
                }
        
        # 6. 「一覧」「個体」「牛」「誰」「抽出」「表示」を含む → 分析（抽出）
        extraction_keywords = ["一覧", "個体", "牛", "誰", "抽出", "表示", "リスト", "検索"]
        if any(keyword in text for keyword in extraction_keywords):
            return {
                "mode": "analysis",
                "original_text": text,
                "clean_text": text
            }
        
        # 7. 「なぜ」「理由」「原因」「傾向」「下がった」「高い」「低い」等を含む → 分析（解釈）
        interpretation_keywords = [
            "なぜ", "理由", "原因", "傾向", "下がった", "上がった", "高い", "低い",
            "改善", "悪化", "変化", "比較", "違い", "差", "要因", "影響"
        ]
        if any(keyword in text for keyword in interpretation_keywords):
            return {
                "mode": "analysis",
                "original_text": text,
                "clean_text": text
            }
        
        # 8. どれにも該当しない場合 → ガイド文を表示
        return {
            "mode": "guide",
            "original_text": text,
            "clean_text": text
        }
    
    def _check_analysis_mode_activation(self, raw_input: str) -> bool:
        """
        分析モードの有効化条件をチェック（Phase D - Step 2）
        
        Args:
            raw_input: ユーザー入力
        
        Returns:
            分析モードを有効化する場合はTrue
        """
        # UIで「分析モード」がONの場合
        if self.analysis_mode_var.get():
            return True
        
        # 入力文の先頭が特定のキーワードで始まる場合
        analysis_keywords = ["分析：", "解説：", "理由："]
        for keyword in analysis_keywords:
            if raw_input.startswith(keyword):
                return True
        
        return False
    
    def _show_guide_message(self, text: str):
        """
        ガイドメッセージを表示
        
        Args:
            text: ユーザー入力
        """
        guide_text = f"""この質問は『条件付きの個体抽出』として実行できます。

例：
• 「10月に分娩した個体を一覧にして」
• 「2産の個体を表示」
• 「受胎率が高い個体を抽出」

または、定型レポートを使用する場合：
• 「月別分娩頭数」
• 「月別受胎率（産次別）」

分析を依頼する場合：
• 「分析：最近6か月で2産の受胎率が下がった理由」
• 「なぜ受胎率が低いのか」"""
        
        self.add_message(role="system", text=guide_text)
    
    def _show_guide_for_report(self, text: str):
        """
        レポートが見つからない場合のガイドを表示
        
        Args:
            text: ユーザー入力
        """
        # 定型レポート一覧を取得
        reports = self.analysis_reports.list_reports()
        report_names = [report.get("display_name", "") for report in reports.values()]
        
        guide_text = f"""レポート「{text}」が見つかりませんでした。

利用可能な定型レポート：
{chr(10).join(f"• {name}" for name in report_names)}

例：
• 「レポート：月別分娩頭数（産次別）」
• 「月別受胎率（産次別）」"""
        
        self.add_message(role="system", text=guide_text)
    
    def _check_report_command(self, raw_input: str) -> Optional[Dict[str, Any]]:
        """
        定型レポートコマンドをチェック（Phase E-1）
        
        Args:
            raw_input: ユーザー入力
        
        Returns:
            レポート情報（該当する場合）、該当しない場合はNone
        """
        # 「レポート：」「報告：」で始まる場合
        report_prefixes = ["レポート：", "報告："]
        report_text = None
        
        for prefix in report_prefixes:
            if raw_input.startswith(prefix):
                report_text = raw_input[len(prefix):].strip()
                break
        
        if not report_text:
            return None
        
        # 表示名からレポートキーを検索
        report_key = self.analysis_reports.find_report_by_display_name(report_text)
        if not report_key:
            return None
        
        # レポート定義を取得
        report = self.analysis_reports.get_report(report_key)
        if not report:
            return None
        
        return {
            "report_key": report_key,
            "report": report,
            "input_text": report_text
        }
    
    def _execute_report(self, report_info: Dict[str, Any], raw_input: str):
        """
        定型レポートを実行（Phase E-1）
        
        Args:
            report_info: レポート情報
            raw_input: ユーザー入力
        """
        report = report_info["report"]
        template_name = report.get("template_name")
        
        if not template_name:
            self.add_message(role="system", text="レポート定義にテンプレート名がありません。")
            return
        
        # 期間を決定（優先順位：入力文 > UI期間設定 > デフォルト）
        period_type = report.get("default_params", {}).get("period", "last_12_months")
        
        # 1. 入力文から期間を解析（Phase E-1）
        parsed_period = self.analysis_reports.parse_period_from_text(raw_input)
        if parsed_period:
            # 日本語で期間が指定された場合は上書き
            period_params = parsed_period
        else:
            # 2. UIの期間設定をチェック
            ui_period = self._get_period_from_ui()
            if ui_period:
                period_params = ui_period
            else:
                # 3. デフォルト期間を使用
                period_params = self.analysis_reports.calculate_period(period_type)
        
        # 出力形式を決定
        default_output = report.get("default_output", ["screen"])
        export_csv, export_excel = self._check_export_format(raw_input)
        
        # デフォルト出力を上書き
        if export_csv or export_excel:
            # ユーザー指定があれば上書き
            pass
        else:
            # デフォルト出力に従う
            export_csv = "csv" in default_output
            export_excel = "excel" in default_output
        
        # 定型レポート実行時は analysis_mode を有効化しない
        # （定型レポートのSQLは安全なSELECT文のみのため、危険SQLチェックをスキップする）
        self.export_csv = export_csv
        self.export_excel = export_excel
        
        # テンプレートを展開
        sql = self.sql_templates.expand_template(template_name, period_params)
        
        if not sql:
            # Phase E-3.3: 業務向け文言
            self.add_message(role="system", text="処理を実行できませんでした。レポート定義に問題があります。")
            return
        
        # SQLを実行（定型レポート用：危険SQLチェックをスキップ）
        # 定型レポートのSQLは analysis_sql_templates.py に定義された安全なSELECT文のみ
        sql_result = self._execute_sql_safely(sql, skip_dangerous_check=True)
        
        if sql_result is None:
            # Phase E-3.3: 業務向け文言
            self.add_message(role="system", text="予期しないエラーが発生しました。操作をやり直してください")
            return
        
        # 結果を表示・出力
        self._display_report_result(
            report_info,
            sql,
            sql_result,
            period_params
        )
    
    def _display_report_result(
        self,
        report_info: Dict[str, Any],
        sql: str,
        sql_result: List[Dict[str, Any]],
        period_params: Dict[str, str]
    ):
        """
        定型レポートの結果を表示（Phase E-1）
        
        Args:
            report_info: レポート情報
            sql: 実行されたSQL
            sql_result: SQL実行結果
            period_params: 期間パラメータ
        """
        report = report_info["report"]
        template_name = report.get("template_name")
        display_name = report.get("display_name", "")
        
        output_parts = []
        
        # レポート情報
        output_parts.append(f"【定型レポート】{display_name}")
        output_parts.append(f"期間: {period_params.get('start')} ～ {period_params.get('end')}")
        output_parts.append("")
        
        # 使用したSQL
        output_parts.append("【使用したSQL】")
        output_parts.append(f"テンプレート: {template_name}")
        output_parts.append("```sql")
        output_parts.append(sql)
        output_parts.append("```")
        output_parts.append("")
        
        # SQL結果（Phase E-3.1: 0件でも安全に処理）
        output_parts.append("【SQL結果】")
        formatted_result = self._format_sql_result(sql_result)
        if formatted_result == "該当データがありません":
            output_parts.append("該当データがありません")
        else:
            output_parts.append(formatted_result)
        output_parts.append("")
        
        # ファイル出力
        if self.export_csv or self.export_excel:
            export_info = self._export_analysis_results(
                sql_result,
                template_name,
                period_params
            )
            if export_info:
                output_parts.append("【出力ファイル】")
                for info in export_info:
                    output_parts.append(info)
        
        # 結果を表示
        self.add_message(role="ai", text="\n".join(output_parts))
    
    def _get_period_from_ui(self) -> Optional[Dict[str, str]]:
        """
        UIの期間設定を取得
        
        Returns:
            期間パラメータ（{"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"}）またはNone
        """
        # 期間設定エントリから最新の値を取得
        self._update_period_from_entries()
        
        # selected_periodをチェック
        if self.selected_period.get("start") and self.selected_period.get("end"):
            start_date = self.selected_period["start"]
            end_date = self.selected_period["end"]
            
            # dateオブジェクトの場合は文字列に変換
            if isinstance(start_date, date):
                start_str = start_date.strftime("%Y-%m-%d")
            else:
                start_str = str(start_date)
            
            if isinstance(end_date, date):
                end_str = end_date.strftime("%Y-%m-%d")
            else:
                end_str = str(end_date)
            
            return {
                "start": start_str,
                "end": end_str
            }
        
        return None
    
    def _check_export_format(self, raw_input: str) -> Tuple[bool, bool]:
        """
        出力形式をチェック（Phase D - Step 4）
        
        Args:
            raw_input: ユーザー入力
        
        Returns:
            (csv出力フラグ, excel出力フラグ)
        """
        # UIチェックボックスから取得
        export_csv = self.export_csv_var.get()
        export_excel = self.export_excel_var.get()
        
        # 入力文の接頭辞から取得
        input_upper = raw_input.upper()
        if "CSV：" in raw_input or raw_input.startswith("CSV："):
            export_csv = True
        if "EXCEL：" in input_upper or raw_input.startswith("Excel：") or raw_input.startswith("EXCEL："):
            export_excel = True
        if "CSV+EXCEL：" in input_upper or "CSV+Excel：" in raw_input:
            export_csv = True
            export_excel = True
        
        return (export_csv, export_excel)
    
    def _on_command_execute(self):
        """
        コマンド実行（自動モード判定対応）
        
        処理フロー:
        1. コマンドラインに入力された文字列を受け取る
        2. classify_command()でモードを判定
        3. 判定されたモードに応じて処理を実行
        """
        raw_input = self.command_entry.get().strip()
        if not raw_input:
            return
        
        # プレースホルダーを非表示（Phase E-2）
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
        
        # コマンドをクリア
        self.command_entry.delete(0, tk.END)
        
        # モード状態を更新（Phase E-2）
        self._update_mode_label("通常モード", "gray")
        
        # コマンドを分類
        classification = self.classify_command(raw_input)
        mode = classification["mode"]
        clean_text = classification.get("clean_text", raw_input)
        
        # モードに応じて処理を実行
        if mode == "cow_card":
            # 個体カードを開く
            padded_id = str(clean_text).zfill(4)
            self._jump_to_cow_card(padded_id)
        
        elif mode == "report":
            # 定型レポートを実行
            report_info = classification.get("report_info")
            if report_info:
                self._update_mode_label("定型レポート実行中", "green")
                self._execute_report(report_info, raw_input)
            else:
                # レポートが見つからない場合のガイド
                self._show_guide_for_report(raw_input)
        
        elif mode == "analysis":
            # 分析モード：分析処理を実行
            self.analysis_mode = True
            self._update_mode_label("分析モード（DB限定）", "blue")
            
            # 出力形式をチェック（Phase D - Step 4）
            self.export_csv, self.export_excel = self._check_export_format(raw_input)
            
            # 分析モードのキーワードを除去してから実行
            analysis_input = clean_text
            for keyword in ["分析：", "解説：", "理由："]:
                if raw_input.startswith(keyword):
                    analysis_input = raw_input[len(keyword):].strip()
                    break
            
            self._execute_analysis_mode(analysis_input if analysis_input else raw_input)
        
        elif mode == "list":
            # リストコマンドを実行
            self._update_mode_label("リスト表示中", "purple")
            self._execute_list_command(raw_input)
        
        elif mode == "guide":
            # ガイド文を表示
            self._show_guide_message(raw_input)
    
    def _get_analysis_mode_system_prompt(self) -> str:
        """
        分析モード時のシステムプロンプトを取得（Phase D - Step 2/3）
        
        Returns:
            分析モード用のシステムプロンプト
        """
        base_prompt = self.system_prompt or ""
        
        # テンプレート一覧を取得
        template_list = self.sql_templates.get_template_list_for_ai()
        
        analysis_prompt = f"""
---
【分析モード専用ルール】

あなたは FALCON2 の「分析AI」です。

・あなたは必ずデータベース（SQLite）から取得された結果のみを根拠にする
・仮定・一般論・経験則のみでの結論は禁止
・INSERT / UPDATE / DELETE / ALTER は一切禁止
・SQL結果が無い場合は「データ不足」と明示する
・event テーブルの event_lact / event_dim を唯一の産次・DIM定義として扱う
・cow.lact は使用禁止

【クエリタイプの判定（重要）】

ユーザーのクエリを以下の2つに分類すること：

1. **個体一覧・抽出クエリ**（「一覧」「個体」「誰」「抽出」「表示」等を含む）
   → 集計ではなく、条件に合致する個体のリストを返す
   → 例：「3月分娩一覧」→ 3月に分娩した個体のID、分娩日、産次をリスト表示
   → 例：「2産の個体を表示」→ 2産の個体のID、産次等をリスト表示

2. **集計・分析クエリ**（「月別」「平均」「合計」「率」等を含む）
   → 集計結果を返す
   → 例：「月別分娩頭数」→ 月ごとの分娩頭数を集計

【SQLテンプレート使用ルール（最重要）】

原則として SQL は以下のテンプレートから選択すること。

{template_list}

テンプレートを使用する場合は、以下の形式で宣言すること：

【使用テンプレート】
テンプレート名: <template_name>
使用理由: <理由>
パラメータ:
  - start: <開始日 YYYY-MM-DD>
  - end: <終了日 YYYY-MM-DD>

**個体一覧・抽出クエリの場合**：
- 既存テンプレートが集計用の場合は、新規SQLを生成すること
- 新規SQLは以下の形式で生成：
  ```sql
  SELECT 
    c.cow_id AS ID,
    e.event_date AS 分娩月日,
    e.event_lact AS 産次
  FROM event e
  JOIN cow c ON e.cow_auto_id = c.auto_id
  WHERE e.event_number = 202  -- 分娩イベント
    AND e.event_date >= 'YYYY-MM-DD'
    AND e.event_date <= 'YYYY-MM-DD'
    AND e.deleted = 0
  ORDER BY e.event_date, c.cow_id
  ```

新規SQLを生成する場合は、「既存テンプレートでは対応できない理由」を明示すること。

【出力形式（必須）】
1. 使用したSQL（テンプレート名または生成したSQL）
2. SQL結果（表 or CSV）
3. 結論（短く・DB結果に基づくもののみ）

「推測」「可能性」「一般論」は禁止
数値が出ない場合は「該当データなし」と明示
---
"""
        return base_prompt + "\n" + analysis_prompt
    
    def _execute_analysis_mode(self, user_input: str):
        """
        分析モードの実行処理（Phase D - Step 2）
        
        Args:
            user_input: ユーザー入力（「分析：」などのプレフィックスを含む）
        """
        # ChatGPTClientが利用可能かチェック
        if not self.chatgpt_client:
            self.add_message(role="system", text="分析モードにはChatGPT APIが必要です。")
            return
        
        # 分析モードのキーワードを除去
        clean_input = user_input
        for keyword in ["分析：", "解説：", "理由："]:
            if clean_input.startswith(keyword):
                clean_input = clean_input[len(keyword):].strip()
                break
        
        if not clean_input:
            self.add_message(role="system", text="分析内容を入力してください。")
            return
        
        # UIの期間設定を取得して、ユーザー入力に追加
        ui_period = self._get_period_from_ui()
        if ui_period:
            period_note = f"\n\n【期間指定】{ui_period['start']} ～ {ui_period['end']} のデータを対象としてください。"
            clean_input = clean_input + period_note
        
        # ユーザーメッセージを追加
        self.add_message(role="user", text=user_input)
        
        # 分析処理をバックグラウンドで実行
        self.current_job_id += 1
        job_id = self.current_job_id
        self.analysis_running = True
        
        # 期間情報もスレッドに渡す
        thread = threading.Thread(
            target=self._run_analysis_mode_in_thread,
            args=(clean_input, job_id, ui_period),
            daemon=True
        )
        thread.start()
    
    def _run_analysis_mode_in_thread(self, user_input: str, job_id: int, ui_period: Optional[Dict[str, str]] = None):
        """
        バックグラウンドスレッドで分析モードを実行（Phase D - Step 2）
        
        Args:
            user_input: ユーザー入力（プレフィックス除去済み、期間情報含む）
            job_id: ジョブID
            ui_period: UIの期間設定（オプション）
        """
        try:
            # 分析モード用のシステムプロンプトを取得
            system_prompt = self._get_analysis_mode_system_prompt()
            
            # AIに問い合わせ（SQL生成を期待）
            ai_response = self.chatgpt_client.ask(system_prompt, user_input)
            
            # メインスレッドで結果を処理
            self.root.after(0, self._handle_analysis_mode_result, user_input, ai_response, job_id)
            
        except Exception as e:
            # Phase E-3.3: エラーは内部ログにのみ出力
            logging.error(f"分析モード実行エラー: {e}", exc_info=True)
            self.root.after(0, self._handle_analysis_mode_error, job_id)
    
    def _extract_template_info_from_response(self, response: str) -> Optional[Dict[str, Any]]:
        """
        AI応答からテンプレート情報を抽出（Phase D - Step 3）
        
        Args:
            response: AI応答テキスト
        
        Returns:
            テンプレート情報（template_name, params）またはNone
        """
        # テンプレート名を抽出
        template_pattern = r'テンプレート名:\s*(\w+)'
        match = re.search(template_pattern, response, re.IGNORECASE)
        if not match:
            return None
        
        template_name = match.group(1).strip()
        
        # パラメータを抽出
        params = {}
        
        # start パラメータ
        start_pattern = r'start:\s*(\d{4}-\d{2}-\d{2})'
        match = re.search(start_pattern, response, re.IGNORECASE)
        if match:
            params['start'] = match.group(1)
        
        # end パラメータ
        end_pattern = r'end:\s*(\d{4}-\d{2}-\d{2})'
        match = re.search(end_pattern, response, re.IGNORECASE)
        if match:
            params['end'] = match.group(1)
        
        return {
            'template_name': template_name,
            'params': params
        }
    
    def _extract_sql_from_response(self, response: str) -> Optional[str]:
        """
        AI応答からSQL文を抽出（Phase D - Step 2/3）
        
        Args:
            response: AI応答テキスト
        
        Returns:
            抽出されたSQL文（見つからない場合はNone）
        """
        # まずテンプレート情報を抽出（Phase D - Step 3）
        template_info = self._extract_template_info_from_response(response)
        if template_info:
            template_name = template_info['template_name']
            params = template_info['params']
            
            # テンプレートを展開
            sql = self.sql_templates.expand_template(template_name, params)
            if sql:
                logging.info(f"[分析モード] テンプレート使用: {template_name}, params={params}")
                return sql
        
        # テンプレートが見つからない場合、直接SQLを抽出（後方互換性）
        # ```sql ... ``` 形式
        sql_pattern1 = r'```sql\s*(.*?)\s*```'
        match = re.search(sql_pattern1, response, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        # ``` ... ``` 形式（SQLと明示されていない場合）
        sql_pattern2 = r'```\s*(SELECT.*?)\s*```'
        match = re.search(sql_pattern2, response, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        # SELECT文が直接書かれている場合
        sql_pattern3 = r'(SELECT\s+.*?;)'
        match = re.search(sql_pattern3, response, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        return None
    
    def _execute_sql_safely(self, sql: str, skip_dangerous_check: bool = False) -> Optional[List[Dict[str, Any]]]:
        """
        SQLを安全に実行（SELECT文のみ、Phase D - Step 2）
        
        【重要】
        - 危険SQLチェックは「sql引数」に対してのみ行う
        - AIレスポンス全文・ログ・テンプレート宣言文は一切チェック対象にしない
        
        Args:
            sql: SQL文（SELECT文のみ）
            skip_dangerous_check: Trueの場合、危険SQLチェックをスキップ（定型レポート用）
        
        Returns:
            実行結果（行のリスト）、エラー時はNone
        """
        # 1. sql が None / 空文字 の場合は即 return（チェックしない）
        if not sql or not sql.strip():
            return None
        
        # 危険SQLチェック（定型レポート実行時はスキップ）
        if not skip_dangerous_check:
            # 2. チェック対象は re.sub などで整形した SQL 本体のみ
            # SQL文を整形：前後の空白を除去、複数の空白を1つに
            sql_cleaned = re.sub(r'\s+', ' ', sql.strip())
            
            # 3. "DELETE", "UPDATE", "INSERT", "ALTER", "DROP" は
            #    SQL文の先頭キーワードとしてのみ検出する
            # 先頭のキーワードを抽出（大文字小文字を区別しない）
            dangerous_keywords = ['DELETE', 'UPDATE', 'INSERT', 'ALTER', 'DROP', 'CREATE', 'TRUNCATE']
            sql_upper = sql_cleaned.upper()
            
            # SQL文の先頭キーワードをチェック（正規表現で先頭のみマッチ）
            for keyword in dangerous_keywords:
                # 先頭にキーワードが来るパターン（空白やコメントを除く）
                pattern = rf'^\s*{keyword}\s+'
                if re.match(pattern, sql_upper, re.IGNORECASE):
                    logging.error(f"危険なSQL文が検出されました: {keyword}")
                    return None
            
            # SELECT文のみ許可（先頭キーワードとして、またはWITH句で始まるSELECT文）
            # WITH RECURSIVE ... SELECT 形式も許可
            if not (re.match(r'^\s*SELECT\s+', sql_upper, re.IGNORECASE) or 
                    re.match(r'^\s*WITH\s+', sql_upper, re.IGNORECASE)):
                logging.error("SELECT文以外は実行できません")
                return None
        
        try:
            # ワーカースレッド内で新しいDBHandlerを生成（SQLiteスレッド違反を回避）
            db_path = self.farm_path / "farm.db"
            from db.db_handler import DBHandler
            db = DBHandler(db_path)
            conn = db.connect()
            cursor = conn.cursor()
            
            # SQLを実行
            cursor.execute(sql)
            rows = cursor.fetchall()
            
            # 結果を辞書形式に変換
            result = []
            if rows:
                # カラム名を取得
                columns = [description[0] for description in cursor.description]
                for row in rows:
                    result.append(dict(zip(columns, row)))
            
            conn.close()
            db.close()
            
            return result
            
        except Exception as e:
            # Phase E-3.3: エラーは内部ログにのみ出力
            logging.error(f"SQL実行エラー: {e}", exc_info=True)
            return []  # Phase E-3.1: エラー時は空リストを返す（0件扱い）
        
        finally:
            # 接続を確実に閉じる（Phase E-3.2）
            try:
                if 'conn' in locals():
                    conn.close()
                if 'db' in locals():
                    db.close()
            except:
                pass
    
    def _format_sql_result(self, result: List[Dict[str, Any]]) -> str:
        """
        SQL結果を表形式の文字列にフォーマット（Phase D - Step 2 / E-3.1）
        
        Args:
            result: SQL実行結果（0件でもOK）
        
        Returns:
            フォーマットされた文字列（0件の場合は"該当データがありません"）
        """
        if not result:
            return "該当データがありません"
        
        # カラム名を取得
        columns = list(result[0].keys())
        
        # ヘッダー行
        header = " | ".join(str(col) for col in columns)
        separator = "-" * len(header)
        
        # データ行
        rows = []
        for row in result:
            values = [str(row.get(col, '')) if row.get(col) is not None else 'NULL' for col in columns]
            rows.append(" | ".join(values))
        
        # 合計行を追加（ピボット形式の結果の場合）
        # lact1, lact2, lact3plus, total などの数値列がある場合に合計を計算
        numeric_columns = []
        for col in columns:
            # ym列以外で、数値として扱える列を特定
            if col != 'ym':
                # 最初の行で数値かどうかをチェック
                first_value = result[0].get(col)
                if first_value is not None:
                    # 数値型（int, float）または数値文字列をチェック
                    if isinstance(first_value, (int, float)):
                        numeric_columns.append(col)
                    else:
                        try:
                            # 文字列の場合、数値に変換できるかチェック
                            float(str(first_value))
                            numeric_columns.append(col)
                        except (ValueError, TypeError):
                            pass
        
        # 合計行を追加
        if numeric_columns:
            total_row = []
            for col in columns:
                if col == 'ym':
                    total_row.append("合計")
                elif col in numeric_columns:
                    # 数値列の合計を計算
                    total = 0
                    for row in result:
                        val = row.get(col)
                        if val is not None:
                            try:
                                if isinstance(val, (int, float)):
                                    total += val
                                else:
                                    total += float(str(val))
                            except (ValueError, TypeError):
                                pass
                    total_row.append(str(total))
                else:
                    total_row.append("")  # 非数値列は空
            
            rows.append("-" * len(header))  # 区切り線
            rows.append(" | ".join(total_row))
        
        return "\n".join([header, separator] + rows)
    
    def _parse_list_command(self, command: str) -> Dict[str, Any]:
        """
        リストコマンドをパース
        
        Args:
            command: リストコマンド（例：「リスト　ID　分娩後日数　産次　繁殖コード　検診日　検診内容　：経産牛のみ：分娩後150日以上」）
        
        Returns:
            {
                "columns": List[str],  # 列名リスト
                "conditions": List[str]  # 条件リスト
            }
        """
        # 「リスト」を除去
        text = command.replace("リスト", "").strip()
        
        # 「：」で条件部分を分離
        parts = text.split("：")
        column_part = parts[0].strip() if parts else ""
        conditions = [c.strip() for c in parts[1:] if c.strip()] if len(parts) > 1 else []
        
        # 列名を抽出（スペースまたは全角スペースで分割）
        columns = []
        if column_part:
            # 全角スペースと半角スペースの両方に対応
            import re
            columns = re.split(r'[\s　]+', column_part)
            columns = [c.strip() for c in columns if c.strip()]
        
        return {
            "columns": columns,
            "conditions": conditions
        }
    
    def _build_list_sql(self, columns: List[str], conditions: List[str]) -> Optional[str]:
        """
        リストコマンド用のSQLを生成
        
        Args:
            columns: 列名リスト
            conditions: 条件リスト
        
        Returns:
            SQL文（生成できない場合はNone）
        """
        # 項目名マッピング
        column_mapping = {
            "ID": "c.cow_id",
            "分娩後日数": "CASE WHEN c.clvd IS NOT NULL THEN CAST((julianday('now') - julianday(c.clvd)) AS INTEGER) ELSE NULL END",
            "産次": "c.lact",
            "繁殖コード": "c.rc",
            "検診日": """(
                SELECT MAX(e2.event_date)
                FROM event e2
                WHERE e2.cow_auto_id = c.auto_id
                  AND e2.event_number IN (300, 301, 302, 303, 304, 305, 306, 307)
                  AND e2.deleted = 0
            )""",
            "検診内容": """(
                SELECT 
                    CASE e3.event_number
                        WHEN 300 THEN '繁殖検診'
                        WHEN 301 THEN '繁殖検診'
                        WHEN 302 THEN '妊娠鑑定マイナス'
                        WHEN 303 THEN '妊娠鑑定プラス'
                        WHEN 304 THEN '妊娠鑑定プラス（検診以外）'
                        WHEN 305 THEN '流産'
                        WHEN 306 THEN 'PAGマイナス'
                        WHEN 307 THEN 'PAGプラス'
                        ELSE ''
                    END
                FROM event e3
                WHERE e3.cow_auto_id = c.auto_id
                  AND e3.event_number IN (300, 301, 302, 303, 304, 305, 306, 307)
                  AND e3.deleted = 0
                ORDER BY e3.event_date DESC
                LIMIT 1
            )""",
        }
        
        # SELECT句を構築
        select_parts = []
        for col in columns:
            if col in column_mapping:
                select_parts.append(f"{column_mapping[col]} AS `{col}`")
            else:
                # 未知の列名はそのまま使用（エスケープ）
                select_parts.append(f"NULL AS `{col}`")
        
        if not select_parts:
            return None
        
        select_clause = "SELECT " + ", ".join(select_parts)
        
        # FROM句
        from_clause = "FROM cow c"
        
        # WHERE句を構築
        where_parts = ["c.auto_id IS NOT NULL"]  # 基本条件
        
        # 条件をパースしてWHERE句に追加
        for condition in conditions:
            condition_lower = condition.lower()
            
            # 経産牛のみ
            if "経産牛" in condition and ("のみ" in condition or "だけ" in condition):
                where_parts.append("c.lact >= 2")
            
            # 初産牛のみ
            elif "初産" in condition and ("のみ" in condition or "だけ" in condition):
                where_parts.append("c.lact = 1")
            
            # 分娩後日数
            elif "分娩後" in condition:
                # 数値を抽出
                import re
                numbers = re.findall(r'\d+', condition)
                if numbers:
                    days = int(numbers[0])
                    if "以上" in condition or "超" in condition:
                        where_parts.append(f"CAST((julianday('now') - julianday(c.clvd)) AS INTEGER) >= {days}")
                    elif "未満" in condition or "以下" in condition:
                        where_parts.append(f"CAST((julianday('now') - julianday(c.clvd)) AS INTEGER) < {days}")
                    elif "日" in condition:
                        # 単純に「分娩後150日」の場合は「以上」と解釈
                        where_parts.append(f"CAST((julianday('now') - julianday(c.clvd)) AS INTEGER) >= {days}")
            
            # 繁殖コード
            elif "繁殖コード" in condition or "RC" in condition:
                rc_mapping = {
                    "Fresh": 1, "分娩後": 1,
                    "Bred": 2, "授精後": 2,
                    "Pregnant": 3, "妊娠中": 3,
                    "Dry": 4, "乾乳": 4,
                    "Open": 5, "空胎": 5,
                    "Stopped": 6, "繁殖停止": 6
                }
                for key, value in rc_mapping.items():
                    if key in condition:
                        where_parts.append(f"c.rc = {value}")
                        break
        
        where_clause = "WHERE " + " AND ".join(where_parts) if where_parts else ""
        
        # ORDER BY句（ID順）
        order_clause = "ORDER BY c.cow_id"
        
        # SQLを組み立て
        sql = f"{select_clause}\n{from_clause}\n{where_clause}\n{order_clause}"
        
        return sql
    
    def _execute_list_command(self, command: str):
        """
        リストコマンドを実行
        
        Args:
            command: リストコマンド
        """
        try:
            # コマンドをパース
            parsed = self._parse_list_command(command)
            columns = parsed["columns"]
            conditions = parsed["conditions"]
            
            if not columns:
                self.add_message(role="system", text="列が指定されていません。例：リスト　ID　産次")
                return
            
            # SQLを生成
            sql = self._build_list_sql(columns, conditions)
            if not sql:
                self.add_message(role="system", text="SQLの生成に失敗しました。")
                return
            
            # SQLを実行
            sql_result = self._execute_sql_safely(sql, skip_dangerous_check=True)
            
            if sql_result is None:
                self.add_message(role="system", text="予期しないエラーが発生しました。操作をやり直してください")
                return
            
            # 結果を表示
            self._display_list_result(columns, sql_result, sql, command)
            
        except Exception as e:
            logging.error(f"リストコマンド実行エラー: {e}", exc_info=True)
            self.add_message(role="system", text="予期しないエラーが発生しました。操作をやり直してください")
    
    def _display_list_result(
        self,
        columns: List[str],
        sql_result: List[Dict[str, Any]],
        sql: str,
        command: str
    ):
        """
        リスト結果を表示（エクセル出力・プリントアウトボタン付き）
        
        Args:
            columns: 列名リスト
            sql_result: SQL実行結果
            sql: 実行されたSQL
            command: 元のコマンド
        """
        output_parts = []
        
        # ヘッダー
        output_parts.append("【リスト表示】")
        output_parts.append("")
        
        # SQL結果
        if sql_result:
            formatted_result = self._format_sql_result(sql_result)
            output_parts.append(formatted_result)
        else:
            output_parts.append("該当データがありません")
        
        output_parts.append("")
        
        # 使用したSQL（折りたたみ可能）
        output_parts.append("【使用したSQL】")
        output_parts.append("```sql")
        output_parts.append(sql)
        output_parts.append("```")
        output_parts.append("")
        
        # 結果を表示
        result_text = "\n".join(output_parts)
        self.add_message(role="ai", text=result_text)
        
        # エクセル出力・プリントアウトボタンを追加
        self._add_list_action_buttons(columns, sql_result, command)
    
    def _add_list_action_buttons(
        self,
        columns: List[str],
        sql_result: List[Dict[str, Any]],
        command: str
    ):
        """
        リスト結果にエクセル出力・プリントアウトボタンを追加
        
        Args:
            columns: 列名リスト
            sql_result: SQL実行結果
            command: 元のコマンド
        """
        # 最新のメッセージカードを取得
        if not hasattr(self, 'chat_messages_frame') or self.chat_messages_frame is None:
            return
        
        # ボタンフレームを作成
        button_frame = tk.Frame(self.chat_messages_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # エクセル出力ボタン
        excel_button = tk.Button(
            button_frame,
            text="📊 Excelに出力",
            command=lambda: self._export_list_to_excel(columns, sql_result, command),
            bg="#4CAF50",
            fg="white",
            font=("", 10, "bold"),
            padx=10,
            pady=5
        )
        excel_button.pack(side=tk.LEFT, padx=5)
        
        # プリントアウトボタン
        print_button = tk.Button(
            button_frame,
            text="🖨️ プリントアウト",
            command=lambda: self._print_list_result(columns, sql_result, command),
            bg="#2196F3",
            fg="white",
            font=("", 10, "bold"),
            padx=10,
            pady=5
        )
        print_button.pack(side=tk.LEFT, padx=5)
    
    def _export_list_to_excel(
        self,
        columns: List[str],
        sql_result: List[Dict[str, Any]],
        command: str
    ):
        """
        リスト結果をExcelに出力
        
        Args:
            columns: 列名リスト
            sql_result: SQL実行結果
            command: 元のコマンド
        """
        try:
            # 保存先ディレクトリ
            export_dir = self.farm_path / "exports" / "list"
            export_dir.mkdir(parents=True, exist_ok=True)
            
            # ファイル名を生成
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"list_{timestamp}.xlsx"
            filepath = export_dir / filename
            
            # Excelに出力
            success, error_msg = self.result_exporter.export_to_excel(sql_result, columns, filepath)
            
            if success:
                self.add_message(
                    role="system",
                    text=f"Excelファイルを出力しました。\n保存先: {filepath}"
                )
            else:
                self.add_message(
                    role="system",
                    text=f"Excel出力に失敗しました: {error_msg or '不明なエラー'}"
                )
        except Exception as e:
            logging.error(f"Excel出力エラー: {e}", exc_info=True)
            self.add_message(role="system", text="Excel出力中にエラーが発生しました。")
    
    def _print_list_result(
        self,
        columns: List[str],
        sql_result: List[Dict[str, Any]],
        command: str
    ):
        """
        リスト結果をプリントアウト
        
        Args:
            columns: 列名リスト
            sql_result: SQL実行結果
            command: 元のコマンド
        """
        try:
            # 一時ファイルに保存してから印刷
            import tempfile
            from pathlib import Path
            
            # テキスト形式で一時ファイルに保存
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.txt', delete=False) as f:
                # ヘッダー
                f.write("【リスト表示】\n")
                f.write(f"コマンド: {command}\n")
                f.write("=" * 80 + "\n\n")
                
                # データ
                if sql_result:
                    # ヘッダー行
                    header = " | ".join(columns)
                    f.write(header + "\n")
                    f.write("-" * len(header) + "\n")
                    
                    # データ行
                    for row in sql_result:
                        values = [str(row.get(col, '')) if row.get(col) is not None else '' for col in columns]
                        f.write(" | ".join(values) + "\n")
                else:
                    f.write("該当データがありません\n")
                
                temp_filepath = Path(f.name)
            
            # ファイルを開く（デフォルトのアプリケーションで開く）
            import os
            import platform
            
            if platform.system() == 'Windows':
                os.startfile(str(temp_filepath))
            elif platform.system() == 'Darwin':  # macOS
                os.system(f'open "{temp_filepath}"')
            else:  # Linux
                os.system(f'xdg-open "{temp_filepath}"')
            
            self.add_message(
                role="system",
                text="印刷用ファイルを開きました。ファイルから印刷してください。"
            )
            
        except Exception as e:
            logging.error(f"プリントアウトエラー: {e}", exc_info=True)
            self.add_message(role="system", text="プリントアウト中にエラーが発生しました。")
    
    def _handle_analysis_mode_result(self, user_input: str, ai_response: str, job_id: int):
        """
        分析モードの結果を処理（メインスレッドで実行、Phase D - Step 2/4）
        
        Args:
            user_input: ユーザー入力
            ai_response: AI応答
            job_id: ジョブID
        """
        if job_id != self.current_job_id:
            return
        
        try:
            # SQL文を抽出
            sql = self._extract_sql_from_response(ai_response)
            
            if sql:
                # SQLをログ出力
                logging.info(f"[分析モード] 抽出されたSQL: {sql}")
                
                # テンプレート情報を取得（表示用・出力用）
                template_info = self._extract_template_info_from_response(ai_response)
                template_name = template_info.get('template_name') if template_info else None
                template_params = template_info.get('params', {}) if template_info else {}
                
                # SQLを実行
                sql_result = self._execute_sql_safely(sql)
                
                if sql_result is not None:
                    # 出力：1. SQL、2. 結果、3. 結論、4. ファイル出力情報
                    output_parts = []
                    
                    # 1. 使用したSQL（テンプレート名があれば表示）
                    output_parts.append("【使用したSQL】")
                    if template_name:
                        output_parts.append(f"テンプレート: {template_name}")
                        template_desc = self.sql_templates.TEMPLATE_DESCRIPTIONS.get(template_name, "")
                        if template_desc:
                            output_parts.append(f"説明: {template_desc}")
                        output_parts.append("")
                    output_parts.append("```sql")
                    output_parts.append(sql)
                    output_parts.append("```")
                    output_parts.append("")
                    
                    # 2. SQL結果（Phase E-3.1: 0件でも安全に処理）
                    output_parts.append("【SQL結果】")
                    formatted_result = self._format_sql_result(sql_result)
                    if formatted_result == "該当データがありません":
                        output_parts.append("該当データがありません")
                    else:
                        output_parts.append(formatted_result)
                    output_parts.append("")
                    
                    # 3. 結論（AI応答からSQL部分を除いた部分）
                    conclusion = ai_response
                    # SQL部分とテンプレート宣言部分を除去
                    for pattern in [
                        r'【使用テンプレート】.*?パラメータ:.*?\n',
                        r'```sql\s*.*?\s*```',
                        r'```\s*SELECT.*?\s*```',
                        r'SELECT\s+.*?;'
                    ]:
                        conclusion = re.sub(pattern, '', conclusion, flags=re.DOTALL | re.IGNORECASE)
                    conclusion = conclusion.strip()
                    
                    if conclusion:
                        output_parts.append("【結論】")
                        output_parts.append(conclusion)
                        output_parts.append("")
                    
                    # 4. ファイル出力（Phase D - Step 4）
                    if self.export_csv or self.export_excel:
                        export_info = self._export_analysis_results(
                            sql_result,
                            template_name,
                            template_params
                        )
                        if export_info:
                            output_parts.append("【出力ファイル】")
                            for info in export_info:
                                output_parts.append(info)
                    
                    # 結果を表示
                    self.add_message(role="ai", text="\n".join(output_parts))
                else:
                    # SQL実行エラー（Phase E-3.3: 業務向け文言）
                    if sql_result == []:
                        # 0件の場合
                        self.add_message(role="system", text="該当データがありませんでした")
                    else:
                        # エラーの場合
                        self.add_message(role="system", text="予期しないエラーが発生しました。操作をやり直してください")
            else:
                # SQLが見つからない場合、AI応答をそのまま表示
                self.add_message(role="ai", text=ai_response)
        
        finally:
            if job_id == self.current_job_id:
                self.analysis_running = False
    
    def _export_analysis_results(
        self,
        sql_result: List[Dict[str, Any]],
        template_name: Optional[str],
        template_params: Dict[str, Any]
    ) -> Optional[List[str]]:
        """
        分析結果をCSV/Excelに出力（Phase D - Step 4 / E-3.1/3.2/3.3）
        
        Args:
            sql_result: SQL実行結果（0件でもOK、ヘッダー行のみ出力）
            template_name: テンプレート名
            template_params: テンプレートパラメータ
        
        Returns:
            出力情報のリスト（成功時）、失敗時はNone
        """
        # 0件でも列名を取得できるようにする（Phase E-3.1）
        if sql_result:
            columns = list(sql_result[0].keys())
        else:
            # 0件の場合はテンプレートから列名を推測（簡易対応）
            # 実際にはSQL結果が0件でも列名は取得できるはずだが、念のため
            columns = []
            if template_name:
                # テンプレート名から列名を推測（簡易）
                if "calving" in template_name:
                    columns = ["ym", "lact", "cnt"]
                elif "insemination" in template_name:
                    columns = ["ym", "lact", "cnt"]
                elif "conception_rate" in template_name:
                    columns = ["ym", "lact", "numerator", "denominator", "rate"]
                elif "dim_distribution" in template_name:
                    columns = ["dim_range", "cnt"]
                else:
                    columns = ["result"]
        
        # 0件でも列名が取得できない場合はスキップ（Phase E-3.1）
        if not columns:
            return None
        
        # 保存先ディレクトリ（Phase E-3.2: 自動作成）
        export_dir = self.farm_path / "exports" / "analysis"
        try:
            export_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            logging.error(f"出力ディレクトリ作成エラー: {e}", exc_info=True)
            messagebox.showerror(
                "処理を実行できませんでした",
                "ファイルの書き込みに失敗しました。フォルダの権限を確認してください"
            )
            return None
        
        # パラメータから日付を取得
        start_date = template_params.get('start')
        end_date = template_params.get('end')
        
        export_info = []
        error_messages = []
        
        # CSV出力（Phase E-3.1: 0件でもヘッダー行のみ出力）
        if self.export_csv:
            csv_filename = self.result_exporter.generate_filename(
                template_name, start_date, end_date, "csv"
            )
            csv_filepath = export_dir / csv_filename
            
            success, error_msg = self.result_exporter.export_to_csv(sql_result, columns, csv_filepath)
            if success:
                export_info.append(f"CSV: {csv_filename}")
                export_info.append(f"保存先: {export_dir}")
            else:
                error_messages.append(f"CSV出力: {error_msg or '失敗しました'}")
        
        # Excel出力（Phase E-3.1: 0件でもヘッダー行のみ出力）
        if self.export_excel:
            excel_filename = self.result_exporter.generate_filename(
                template_name, start_date, end_date, "xlsx"
            )
            excel_filepath = export_dir / excel_filename
            
            success, error_msg = self.result_exporter.export_to_excel(sql_result, columns, excel_filepath)
            if success:
                export_info.append(f"Excel: {excel_filename}")
                export_info.append(f"保存先: {export_dir}")
            else:
                error_messages.append(f"Excel出力: {error_msg or '失敗しました'}")
        
        # エラーメッセージがある場合は表示（Phase E-3.3）
        if error_messages:
            messagebox.showerror(
                "処理を実行できませんでした",
                "\n".join(error_messages)
            )
        
        return export_info if export_info else None
    
    def _handle_analysis_mode_error(self, job_id: int):
        """
        分析モードのエラーを処理（Phase D - Step 2 / E-3.3）
        
        Args:
            job_id: ジョブID
        """
        if job_id == self.current_job_id:
            # Phase E-3.3: 業務向け文言
            self.add_message(role="system", text="予期しないエラーが発生しました。操作をやり直してください")
            self.analysis_running = False
    
    def _show_report_list(self):
        """
        定型レポート一覧を表示（Phase E-1）
        """
        reports = self.analysis_reports.list_reports()
        
        if not reports:
            self.add_message(role="system", text="利用可能な定型レポートがありません。")
            return
        
        # 一覧を表示
        list_text = "【定型レポート一覧】\n\n"
        for key, report in reports.items():
            display_name = report.get("display_name", "")
            description = report.get("description", "")
            list_text += f"・{display_name}\n  {description}\n\n"
        
        list_text += "【使用方法】\n"
        list_text += "コマンド欄に以下を入力：\n"
        list_text += "レポート：<表示名>\n"
        list_text += "例：レポート：月別分娩頭数\n"
        list_text += "\n期間指定例：\n"
        list_text += "レポート：月別分娩頭数 2024年\n"
        list_text += "レポート：月別分娩頭数 直近6か月\n"
        list_text += "レポート：月別分娩頭数 3〜8月\n"
        
        # チャット画面に表示
        self.show_chat()
        self.add_message(role="system", text=list_text)
    
    # ========== Phase E-2: UX改善メソッド ==========
    
    def _show_command_placeholder(self):
        """
        コマンド入力欄にプレースホルダーを表示（Phase E-2）
        """
        if not self.command_has_focus and self.command_entry.get() == "":
            self.command_entry.insert(0, self.command_placeholder)
            self.command_entry.config(foreground="gray")
            self.command_placeholder_shown = True
    
    def _hide_command_placeholder(self):
        """
        コマンド入力欄のプレースホルダーを非表示（Phase E-2）
        """
        if self.command_placeholder_shown:
            self.command_entry.delete(0, tk.END)
            self.command_entry.config(foreground="black")
            self.command_placeholder_shown = False
    
    def _on_command_key_press(self, event):
        """
        コマンド入力欄のキー入力イベント（Phase E-2）
        """
        # プレースホルダーを非表示
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
        
        # 入力補助ドロップダウンの表示/非表示
        current_text = self.command_entry.get()
        if current_text.startswith("レポート：") or current_text.startswith("報告："):
            self._show_report_dropdown(current_text)
        else:
            self._hide_report_dropdown()
    
    def _on_command_focus_in(self, event):
        """
        コマンド入力欄にフォーカスが入った時（Phase E-2）
        """
        self.command_has_focus = True
        if self.command_placeholder_shown:
            self._hide_command_placeholder()
    
    def _on_command_focus_out(self, event):
        """
        コマンド入力欄からフォーカスが外れた時（Phase E-2）
        """
        self.command_has_focus = False
        if self.command_entry.get() == "":
            self._show_command_placeholder()
        self._hide_report_dropdown()
    
    def _update_mode_label(self, text: str, color: str):
        """
        モード状態ラベルを更新（Phase E-2）
        
        Args:
            text: 表示テキスト
            color: 色（"gray", "blue", "green"）
        """
        if hasattr(self, 'mode_label'):
            self.mode_label.config(text=text, foreground=color)
    
    def _show_report_dropdown(self, current_text: str):
        """
        定型レポート名のドロップダウンを表示（Phase E-2）
        
        Args:
            current_text: 現在の入力テキスト
        """
        # 既に表示されている場合は更新
        if self.report_dropdown_frame is not None:
            self._update_report_dropdown(current_text)
            return
        
        # ドロップダウンフレームを作成
        command_frame = self.command_entry.master
        self.report_dropdown_frame = ttk.Frame(command_frame)
        self.report_dropdown_frame.pack(side=tk.TOP, fill=tk.X, padx=(60, 0), pady=(5, 0))
        
        # リストボックスを作成
        self.report_dropdown_listbox = tk.Listbox(
            self.report_dropdown_frame,
            height=5,
            font=("", 9)
        )
        self.report_dropdown_listbox.pack(fill=tk.X)
        self.report_dropdown_listbox.bind('<Double-Button-1>', self._on_report_dropdown_select)
        self.report_dropdown_listbox.bind('<Return>', self._on_report_dropdown_select)
        
        # 候補を更新
        self._update_report_dropdown(current_text)
    
    def _update_report_dropdown(self, current_text: str):
        """
        ドロップダウンの候補を更新（Phase E-2）
        
        Args:
            current_text: 現在の入力テキスト
        """
        if self.report_dropdown_listbox is None:
            return
        
        # リストボックスをクリア
        self.report_dropdown_listbox.delete(0, tk.END)
        
        # 定型レポート一覧を取得
        reports = self.analysis_reports.list_reports()
        
        # 入力テキストから検索キーワードを抽出
        prefix = ""
        if current_text.startswith("レポート："):
            prefix = "レポート："
            search_text = current_text[len("レポート："):].strip()
        elif current_text.startswith("報告："):
            prefix = "報告："
            search_text = current_text[len("報告："):].strip()
        else:
            return
        
        # 候補を追加
        for key, report in reports.items():
            display_name = report.get("display_name", "")
            if not search_text or search_text in display_name:
                self.report_dropdown_listbox.insert(tk.END, f"{prefix}{display_name}")
    
    def _hide_report_dropdown(self):
        """
        ドロップダウンを非表示（Phase E-2）
        """
        if self.report_dropdown_frame is not None:
            self.report_dropdown_frame.destroy()
            self.report_dropdown_frame = None
            self.report_dropdown_listbox = None
    
    def _on_report_dropdown_select(self, event):
        """
        ドロップダウンから選択された時（Phase E-2）
        """
        if self.report_dropdown_listbox is None:
            return
        
        selection = self.report_dropdown_listbox.curselection()
        if selection:
            selected_text = self.report_dropdown_listbox.get(selection[0])
            self.command_entry.delete(0, tk.END)
            self.command_entry.insert(0, selected_text)
            self.command_entry.config(foreground="black")
            self.command_placeholder_shown = False
            self._hide_report_dropdown()
            # フォーカスを戻す
            self.command_entry.focus()
    
    def _show_quick_guide_if_first_time(self):
        """
        初回起動時のクイックガイドを表示（Phase E-2）
        """
        # 設定を確認
        if not hasattr(self, 'settings_manager'):
            return
        
        # 既に表示済みかチェック
        guide_shown = self.settings_manager.get("quick_guide_shown", False)
        if guide_shown:
            return
        
        # ガイドを表示
        guide_text = """FALCON2 クイックガイド

【基本的な使い方】

• 数字を入力 → 個体カードを開く
  例：101

• 「レポート：」→ 定型集計
  例：レポート：月別分娩頭数 2024年 CSV

• 「分析：」→ DB限定の分析モード
  例：分析：最近6か月で2産の受胎率が下がった理由

• CSV / Excel を付けると出力されます
  例：レポート：月別分娩頭数 Excel

【ヒント】
• 「定型レポート一覧」ボタンで利用可能なレポートを確認できます
• コマンド入力欄の例示テキストも参考にしてください"""
        
        messagebox.showinfo("FALCON2 クイックガイド", guide_text)
        
        # 表示済みフラグを保存
        self.settings_manager.set("quick_guide_shown", True)
    
    def _execute_query_router(self, query: str):
        """自然文解析は無効化（何もしない）"""
        return
    
    def _run_query_router_in_thread(self, query: str, job_id: int):
        """自然文解析は無効化（何もしない）"""
        return
    
    def _handle_ambiguous_query(self, *args, **kwargs):
        """自然文解析は無効化（何もしない）"""
        return
    
    def _show_item_selection_dialog(self, *args, **kwargs) -> Optional[str]:
        """自然文解析は無効化（ダイアログは使用しない）"""
        return None
    
    def _run_query_router_in_thread_with_item(self, query: str, selected_item_key: str,
                                              ui_period: Optional[Dict[str, Any]], job_id: int):
        """自然文解析は無効化（何もしない）"""
        return
    
    def _handle_query_router_result(self, result_text: Optional[str], job_id: int, 
                                   result_data: Optional[Dict[str, Any]] = None,
                                   period_info: Optional[Dict[str, Any]] = None):
        """自然文解析は無効化（何もしない）"""
        if job_id == self.current_job_id:
            self.analysis_running = False
        return
    
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
        
        # 項目辞書の各項目をチェック
        for item_key, item_def in self.item_dictionary.items():
            display_name = item_def.get('display_name', '')
            if not display_name:
                continue
            
            # display_nameが入力文に含まれているかチェック
            if display_name in query:
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
                # 搾乳中（RC=1: Fresh）の牛のみ対象
                if cow.get('rc') != 1:  # RC_FRESH
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
        初回授精日数を計算
        分娩日から、最初に記録された AI または ET イベント日までの日数
        
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
                # 分娩日がない個体は除外
                return None
            
            try:
                clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
            except:
                return None
            
            # イベント履歴を取得
            events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            # AI または ET イベントを取得（分娩日以降のもの）
            ai_et_events = []
            for event in events:
                event_number = event.get('event_number')
                event_date_str = event.get('event_date')
                if event_number in [self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET] and event_date_str:
                    try:
                        event_date = datetime.strptime(event_date_str, '%Y-%m-%d')
                        if event_date >= clvd_date:
                            ai_et_events.append((event_date, event_date_str))
                    except:
                        continue
            
            if not ai_et_events:
                # 初回AI/ETがない個体は除外
                return None
            
            # 最初のAI/ETイベントを取得（日付順にソート）
            ai_et_events.sort(key=lambda x: x[0])
            first_ai_date = ai_et_events[0][0]
            
            # 分娩日から最初のAI/ETまでの日数
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
                # 搾乳中（RC=1: Fresh）の牛のみ対象
                if cow.get('rc') != 1:  # RC_FRESH
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
    
    def _display_scatter_plot_from_result(self, result: Dict[str, Any]):
        """
        新しいQueryRouterシステムの結果から散布図を表示（クリック可能）
        
        Args:
            result: QueryExecutorV2の結果辞書（kind="scatter", rows, meta）
        """
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
            x_label = self.query_renderer._get_item_display_name(x_item_key)
            y_label = self.query_renderer._get_item_display_name(y_item_key)
            
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
                    # QueryRendererでCSV形式に変換
                    csv_text = self.query_renderer.render(result, query_type="scatter")
                    self.root.clipboard_clear()
                    self.root.clipboard_append(csv_text)
                    self.root.update()
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
            
            save_button = ttk.Button(button_frame, text="画像として保存", command=save_plot_image)
            save_button.pack(side=tk.LEFT, padx=5)
            
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
            
            save_button = ttk.Button(button_frame, text="画像として保存", command=save_plot_image)
            save_button.pack(side=tk.LEFT, padx=5)
            
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
                # 搾乳中（RC=1: Fresh）の牛のみ対象
                if cow.get('rc') != 1:  # RC_FRESH
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
    
    def _ask_chatgpt(self, user_prompt: str):
        """
        ChatGPTに質問を送信（非推奨：ルールエンジンで処理できない場合のみ）
        
        注意：このメソッドは使用されません（ルールエンジンで未対応の場合は「未対応」として返す）
        
        Args:
            user_prompt: ユーザーの質問
        """
        # AIは使用しない（ルールエンジンで処理できない場合は未対応として返す）
        logging.warning(f"[ルールエンジン] _ask_chatgptが呼ばれましたが、AIは使用しません: \"{user_prompt}\"")
        self.add_message(role="system", text="未対応のクエリです。")
    
    def _run_analysis_in_thread(self, user_prompt: str, job_id: int):
        """
        バックグラウンドスレッドで分析処理を実行（非推奨：使用されていません）
        
        注意：このメソッドは使用されません（_run_db_aggregation_in_threadを使用）
        
        Args:
            user_prompt: ユーザーの質問
            job_id: ジョブID
        """
        # このメソッドは使用されていません
        logging.warning(f"[ルールエンジン] _run_analysis_in_threadが呼ばれましたが、使用されていません: job_id={job_id}")
        if job_id == self.current_job_id:
            self.analysis_running = False
    
    def _handle_analysis_result(self, user_prompt: str, calculated_data: Any, job_id: int):
        """
        分析結果を処理（メインスレッドで実行）
        
        Args:
            user_prompt: ユーザーの質問
            calculated_data: 計算済みデータ
            job_id: ジョブID
        """
        try:
            # ジョブIDが古い場合は無視（キャンセルされたジョブ）
            if job_id != self.current_job_id:
                logging.info(f"[分析] 古いジョブの結果を無視: job_id={job_id}, current_job_id={self.current_job_id}")
                # 古いジョブの場合はロック解除は不要（現在のジョブが処理する）
                return
            
            if calculated_data is None:
                # エラーメッセージは表示しない（次の処理に進む）
                matched_items = self._match_items_from_query(user_prompt)
                if not matched_items:
                    logging.info(f"[分析] 該当する分析処理が見つかりませんでした（次の処理に進む）: query=\"{user_prompt}\"")
                return
            
            # イベント抽出系クエリの結果（文字列）を表示
            if self._is_event_extraction_query(user_prompt) and isinstance(calculated_data, str):
                self.add_message(role="ai", text=calculated_data)
                logging.info(f"[分析] イベント抽出結果表示: job_id={job_id}")
                return
            
            # 数値系クエリの場合は数値のみを直接表示
            if self._is_numeric_query(user_prompt):
                if isinstance(calculated_data, str):
                    self.add_message(role="ai", text=calculated_data)
                return
            
            # 散布図クエリの場合はグラフを表示（ChatGPTを経由しない）
            if "散布図" in user_prompt and isinstance(calculated_data, dict):
                if 'type' not in calculated_data or calculated_data.get('type') != 'table_report':
                    try:
                        self._display_scatter_plot(calculated_data)
                    except Exception as e:
                        logging.error(f"[分析] 散布図表示エラー: {e}")
                        import traceback
                        traceback.print_exc()
                        self.add_message(role="system", text="散布図の表示に失敗しました。")
                    return
            
            # 表レポートの場合はReportTableWindowを開く（ChatGPTを経由しない）
            if isinstance(calculated_data, dict) and calculated_data.get('type') == 'table_report':
                try:
                    self._display_table_report(calculated_data)
                except Exception as e:
                    logging.error(f"[分析] 表レポート表示エラー: {e}")
                    import traceback
                    traceback.print_exc()
                    self.add_message(role="system", text="表レポートの表示に失敗しました。")
                return
            
            # その他の表・CSV形式の場合はそのまま表示
            if isinstance(calculated_data, str) and ("表" in user_prompt or "csv" in user_prompt.lower()):
                self.add_message(role="ai", text=calculated_data)
                return
            
            # その他の分析クエリは未対応として返す（AIは使用しない）
            logging.info(f"[分析] 未対応の分析クエリ: job_id={job_id}, query=\"{user_prompt}\"")
            self.add_message(role="system", text="未対応のクエリです。")
        
        finally:
            # 分析ロックを必ず解除（成功・失敗・例外のいずれの場合でも）
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[分析] 分析ロック解除: job_id={job_id}")
    
    def _handle_analysis_error(self, job_id: int, error_message: str):
        """
        分析エラーを処理（メインスレッドで実行）
        
        Args:
            job_id: ジョブID
            error_message: エラーメッセージ
        """
        try:
            # ジョブIDが古い場合は無視
            if job_id != self.current_job_id:
                logging.info(f"[分析] 古いジョブのエラーを無視: job_id={job_id}, current_job_id={self.current_job_id}")
                return
            
            logging.error(f"[分析] 分析処理例外発生: job_id={job_id}, error={error_message}")
            # エラーメッセージは表示しない（内部エラー）
            self.add_message(role="system", text="処理できませんでした。")
        
        finally:
            # 分析ロックを必ず解除（成功・失敗・例外のいずれの場合でも）
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[分析] 分析ロック解除（エラー時）: job_id={job_id}")
    
    def _find_event_by_name(self, search_str: str) -> Optional[int]:
        """
        イベント名（alias または name_jp）でイベント番号を検索
        
        Args:
            search_str: 検索文字列（大文字小文字を区別しない）
        
        Returns:
            見つかったイベント番号、見つからない場合はNone
        """
        if not self.event_dict_path or not self.event_dict_path.exists():
            return None
        
        try:
            # イベント辞書を読み込む
            with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                event_dictionary = json.load(f)
            
            search_lower = search_str.lower().strip()
            
            # イベント辞書を検索（alias または name_jp で前方一致）
            for event_num_str, event_data in event_dictionary.items():
                # deprecated のイベントは除外
                if event_data.get('deprecated', False):
                    continue
                
                # alias で検索
                alias = event_data.get('alias', '')
                if alias and alias.lower().startswith(search_lower):
                    return int(event_num_str)
                
                # name_jp で検索
                name_jp = event_data.get('name_jp', '')
                if name_jp and name_jp.lower().startswith(search_lower):
                    return int(event_num_str)
            
            return None
        except Exception as e:
            logging.error(f"イベント検索エラー: {e}")
            return None
    
    def _open_event_input_for_event(self, event_number: int):
        """
        指定されたイベントに固定されたイベント入力ウィンドウを開く
        
        Args:
            event_number: イベント番号
        """
        if self.event_dict_path is None:
            messagebox.showerror(
                "エラー",
                "event_dictionary.json が見つかりません"
            )
            return
        
        # EventInputWindow を生成して表示（該当イベントに固定）
        event_input_window = EventInputWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 汎用モード
            on_saved=self._on_event_saved,
            farm_path=self.farm_path,
            allowed_event_numbers=[event_number],  # 該当イベントのみ
            default_event_number=event_number  # デフォルトで選択
        )
        event_input_window.show()
    
    def _jump_to_cow_card(self, cow_id: str):
        """
        個体カードウィンドウを開く
        
        Args:
            cow_id: 4桁の牛ID（例: "0980"）
        """
        from ui.cow_card_window import CowCardWindow
        
        # 4桁IDで検索（複数件取得可能）
        cows = self.db.get_cows_by_id(cow_id)
        
        if not cows:
            # 見つからない場合は、正規化されたIDで検索を試みる
            normalized_id = cow_id.lstrip('0')  # 左ゼロを除去
            cow = self.db.get_cow_by_normalized_id(normalized_id)
            if cow:
                cows = [cow]
        
        if not cows:
            # 見つからない場合のUX
            messagebox.showinfo(
                "検索結果",
                f"ID {cow_id} の個体は見つかりませんでした"
            )
            self._add_chat_message("システム", f"ID {cow_id} の個体は見つかりませんでした")
            return
        
        # 1件のみの場合はそのまま表示
        if len(cows) == 1:
            cow = cows[0]
            cow_auto_id = cow.get('auto_id')
            if cow_auto_id:
                # 個体カードウィンドウを開く
                cow_card_window = CowCardWindow(
                    parent=self.root,
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    item_dictionary_path=self.item_dict_path,
                    cow_auto_id=cow_auto_id
                )
                cow_card_window.show()
                self._add_chat_message("システム", f"ID {cow.get('cow_id', cow_id)} カードを開きました。")
            return
        
        # 複数件見つかった場合は選択ダイアログを表示
        selected_cow = self._show_cow_selection_dialog(cows, cow_id)
        if selected_cow:
            cow_auto_id = selected_cow.get('auto_id')
            if cow_auto_id:
                # 個体カードウィンドウを開く
                cow_card_window = CowCardWindow(
                    parent=self.root,
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    item_dictionary_path=self.item_dict_path,
                    cow_auto_id=cow_auto_id
                )
                cow_card_window.show()
                self._add_chat_message("システム", f"ID {selected_cow.get('cow_id', cow_id)} カードを開きました。")
    
    def _show_cow_selection_dialog(self, cows: List[Dict[str, Any]], cow_id: str) -> Optional[Dict[str, Any]]:
        """
        4桁IDが重複している場合の個体選択ダイアログ
        
        Args:
            cows: 検索結果の牛リスト
            cow_id: 検索した4桁ID
        
        Returns:
            選択された牛の情報、キャンセルの場合はNone
        """
        # 選択ダイアログウィンドウを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("個体選択")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # メッセージ
        message = f"拡大4桁ID {cow_id} に該当する個体が複数見つかりました。\n個体識別番号から選択してください。"
        ttk.Label(dialog, text=message, wraplength=450).pack(pady=10)
        
        # リストボックス
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # スクロールバー付きリストボックス
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Courier", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 選択結果を保持
        selected_cow = None
        
        # リストボックスに個体情報を追加
        for cow in cows:
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            bthd = cow.get('bthd', '')
            pen = cow.get('pen', '')
            
            # 表示形式: 個体識別番号 | 品種 | 生年月日 | 群
            display_text = f"{jpn10:12s} | {brd:10s} | {bthd:10s} | {pen:10s}"
            listbox.insert(tk.END, display_text)
        
        # リストボックスの選択イベント
        def on_select(event):
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
        
        listbox.bind('<<ListboxSelect>>', on_select)
        listbox.bind('<Double-Button-1>', lambda e: dialog.destroy())
        
        # ヘッダー行を追加（読みやすさのため）
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill=tk.X, padx=10)
        header_text = f"{'個体識別番号':12s} | {'品種':10s} | {'生年月日':10s} | {'群':10s}"
        ttk.Label(header_frame, text=header_text, font=("Courier", 10, "bold")).pack()
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "個体を選択してください。")
        
        def on_cancel():
            nonlocal selected_cow
            selected_cow = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="選択", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # モーダルダイアログとして表示
        dialog.wait_window()
        
        return selected_cow
    
    def add_message(self, role: Literal["user", "ai", "system"], text: str):
        """
        メッセージを追加（ChatGPT風UI）
        
        Args:
            role: メッセージの役割 ("user", "ai", "system")
            text: メッセージテキスト
        """
        # messages_frame が存在する場合のみ追加
        if not hasattr(self, 'chat_messages_frame') or self.chat_messages_frame is None:
            print(f"[{role}] {text}")
            return
        
        try:
            # メッセージカードを作成
            message_card = self._create_message_card(role, text)
            
            # messages_frameに追加
            message_card.pack(fill=tk.X, pady=8)
            
            # Canvasのスクロール領域を更新
            self.chat_messages_frame.update_idletasks()
            self.chat_canvas.config(scrollregion=self.chat_canvas.bbox("all"))
            
            # 最下部にスクロール
            self.chat_canvas.yview_moveto(1.0)
            
        except Exception as e:
            logging.error(f"メッセージ追加エラー: {e}")
            import traceback
            traceback.print_exc()
        
        print(f"[{role}] {text}")
    
    def _is_table_or_numeric_content(self, text: str) -> bool:
        """
        テキストが表形式または数値が多いかを判定
        
        Args:
            text: メッセージテキスト
        
        Returns:
            表形式または数値が多い場合はTrue
        """
        lines = text.split('\n')
        if len(lines) < 2:
            return False
        
        # CSV形式（カンマ区切り）を検出
        comma_count = sum(line.count(',') for line in lines[:5])  # 最初の5行をチェック
        if comma_count >= 3:
            return True
        
        # 数値が多い行を検出（複数行にわたって数値が多く含まれる場合）
        numeric_line_count = 0
        for line in lines[:10]:  # 最初の10行をチェック
            # 数字、カンマ、スペース、タブが多く含まれる行
            numeric_chars = sum(1 for c in line if c.isdigit() or c in ', \t')
            if len(line) > 0 and numeric_chars / len(line) > 0.4:
                numeric_line_count += 1
        
        if numeric_line_count >= 3:
            return True
        
        return False
    
    def _create_message_card(self, role: Literal["user", "ai", "system"], text: str) -> tk.Frame:
        """
        メッセージカード（Frame + Label）を作成
        
        Args:
            role: メッセージの役割
            text: メッセージテキスト
        
        Returns:
            メッセージカードのFrame
        """
        # メッセージカードのFrame
        card_frame = tk.Frame(self.chat_messages_frame, bg="#FFFFFF")
        
        # フォントを決定（数値・表の場合は等幅フォント）
        if self._is_table_or_numeric_content(text):
            font_family = "Courier New"
        else:
            font_family = "Segoe UI"
        
        # ユーザーメッセージ（右寄せ）
        if role == "user":
            # 右寄せ用のコンテナ
            container = tk.Frame(card_frame, bg="#FFFFFF")
            container.pack(fill=tk.X, anchor=tk.E, padx=10)
            
            # メッセージラベル（右寄せ、背景色#F5F7FA）
            message_label = tk.Label(
                container,
                text=text,
                font=(font_family, 11),
                bg="#F5F7FA",
                fg="#000000",
                wraplength=500,  # 最大幅を設定
                justify=tk.LEFT,
                anchor=tk.W,
                padx=14,
                pady=12,
                relief=tk.FLAT,
                bd=0
            )
            message_label.pack(side=tk.RIGHT)
        
        # AIメッセージまたはシステムメッセージ（左寄せ）
        else:
            # 左寄せ用のコンテナ
            container = tk.Frame(card_frame, bg="#FFFFFF")
            container.pack(fill=tk.X, anchor=tk.W, padx=10)
            
            # 背景色を決定（システムメッセージは薄いグレー）
            bg_color = "#FAFAFA" if role == "system" else "#FFFFFF"
            
            # メッセージラベル（左寄せ、背景色#FFFFFF or #FAFAFA）
            message_label = tk.Label(
                container,
                text=text,
                font=(font_family, 11),
                bg=bg_color,
                fg="#000000",
                wraplength=500,  # 最大幅を設定
                justify=tk.LEFT,
                anchor=tk.W,
                padx=14,
                pady=12,
                relief=tk.FLAT,
                bd=0
            )
            message_label.pack(side=tk.LEFT)
        
        return card_frame
    
    def _add_chat_message(self, sender: str, message: str, color: Optional[str] = None):
        """
        チャットメッセージを追加（後方互換性のため残す）
        
        Args:
            sender: 送信者名 ("ユーザー", "AI回答", "システム" など)
            message: メッセージ
            color: テキストの色（#RRGGBB形式、Noneの場合はデフォルト）
        """
        # senderをroleに変換
        if sender == "ユーザー":
            role = "user"
        elif sender == "AI回答" or sender == "AI":
            role = "ai"
        else:  # "システム" など
            role = "system"
        
        # 新しいadd_messageメソッドを使用
        self.add_message(role=role, text=message)
    
    def _get_event_display_color(self, event_number: int) -> str:
        """
        イベントの表示色を決定（CowCardと同じロジック）
        
        Args:
            event_number: イベント番号
        
        Returns:
            色（#RRGGBB形式）
        """
        # event_dictionary.json を読み込む
        if not hasattr(self, '_event_dictionary') or self._event_dictionary is None:
            self._event_dictionary = {}
            if self.event_dict_path and self.event_dict_path.exists():
                try:
                    with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                        self._event_dictionary = json.load(f)
                except Exception as e:
                    logging.error(f"event_dictionary.json 読み込みエラー: {e}")
        
        # イベント辞書から情報を取得
        event_str = str(event_number)
        event_dict = self._event_dictionary.get(event_str, {})
        
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
    
    def _display_cow_info_in_chat(self, cow_auto_id: int):
        """
        チャット履歴に個体情報とイベント履歴を表示（色付き）
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            # 牛の情報を取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                return
            
            # 個体情報を表示
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            lact = cow.get('lact', '')
            rc = cow.get('rc', '')
            
            self._add_chat_message("システム", f"個体情報: {cow_id} (JPN10: {jpn10}, 品種: {brd}, 産次: {lact}, RC: {rc})")
            
            # イベント履歴を取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            if not events:
                self._add_chat_message("システム", "イベント履歴: なし")
                return
            
            # イベント名を取得するためのヘルパー関数
            def get_event_name(event_number: int) -> str:
                event_str = str(event_number)
                if hasattr(self, '_event_dictionary') and self._event_dictionary:
                    event_dict = self._event_dictionary.get(event_str, {})
                    return event_dict.get('name_jp', f'イベント{event_number}')
                return f'イベント{event_number}'
            
            # イベント履歴を表示（最新から）
            self._add_chat_message("システム", "イベント履歴:")
            for event in events[:10]:  # 最新10件まで
                event_date = event.get('event_date', '')
                event_number = event.get('event_number')
                note = event.get('note', '')
                
                if event_number is None:
                    continue
                
                event_name = get_event_name(event_number)
                color = self._get_event_display_color(event_number)
                
                # イベント情報を色付きで表示
                display_text = f"  {event_date} {event_name}"
                if note:
                    display_text += f" - {note}"
                
                self._add_chat_message("システム", display_text, color=color)
            
            if len(events) > 10:
                self._add_chat_message("システム", f"  ... 他 {len(events) - 10} 件")
                
        except Exception as e:
            logging.error(f"個体情報表示エラー: {e}")
            import traceback
            traceback.print_exc()

