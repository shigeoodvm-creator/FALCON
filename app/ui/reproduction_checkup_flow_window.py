"""
FALCON2 - 繁殖検診フローウインドウ
繁殖検診関連の操作を選択するフローウインドウ
"""

import calendar
import html as html_module
import json
import logging
import tempfile
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine

EXAM_LOG_NOTES_FILENAME = "reproduction_checkup_visit_notes.json"
EXAM_LOG_RATING_LABELS = [
    ("overall", "全体"),
    ("cow_condition", "牛の状態"),
    ("feed", "飼料"),
    ("disease", "疾病"),
    ("milk_volume", "乳量"),
    ("milk_quality", "乳質"),
    ("reproduction", "繁殖"),
    ("calf_rearing", "子牛・育成"),
]


class ReproductionCheckupFlowWindow:
    """繁殖検診フローウインドウ"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 farm_path: Path,
                 event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            farm_path: 農場フォルダのパス
            event_dictionary_path: event_dictionary.json のパス
        """
        self.parent = parent
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        self.event_dict_path = event_dictionary_path
        
        # ウィンドウ作成（データ吸い込みウィンドウと同一デザイン）
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖検診フロー")
        self.window.geometry("520x820")
        self.window.configure(bg="#f5f5f5")
        self.window.minsize(480, 420)
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットを作成（カード型・データ吸い込みと同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        card_bg = "#f5f5f5"
        card_pad = 16
        btn_primary_bg = "#3949ab"
        btn_primary_fg = "#ffffff"
        btn_secondary_bd = "#b0bec5"
        
        # ヘッダー
        header = tk.Frame(self.window, bg=bg, pady=20, padx=24)
        header.pack(fill=tk.X)
        tk.Label(header, text="\u2695", font=(_df, 24), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="繁殖検診フロー", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="操作を選択してください", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)
        
        # 縦スクロール付きコンテンツエリア
        scroll_container = tk.Frame(self.window, bg=bg)
        scroll_container.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(scroll_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas = tk.Canvas(scroll_container, bg=bg, highlightthickness=0, yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=canvas.yview)
        content = tk.Frame(canvas, bg=bg, padx=0, pady=8)
        canvas_window = canvas.create_window(0, 0, window=content, anchor=tk.NW)
        
        def _on_frame_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        def _bind_mousewheel_recursive(widget):
            """ウィジェットとその子孫にマウスホイールをバインド（カード上でもスクロールできるようにする）"""
            widget.bind("<MouseWheel>", _on_mousewheel)
            for child in widget.winfo_children():
                _bind_mousewheel_recursive(child)
        
        content.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<MouseWheel>", _on_mousewheel)
        content.bind("<MouseWheel>", _on_mousewheel)
        
        self._card_icons = []
        
        def add_card(icon_char: Optional[str], title_text: str, desc_text: str, command, icon_path: Optional[Path] = None):
            card = tk.Frame(content, bg=card_bg, padx=card_pad, pady=card_pad,
                            highlightbackground="#e0e7ef", highlightthickness=1)
            card.pack(fill=tk.X, pady=6, padx=24)
            if icon_path is not None and icon_path.exists():
                try:
                    photo = tk.PhotoImage(file=str(icon_path))
                    self._card_icons.append(photo)
                    tk.Label(card, image=photo, bg=card_bg).pack(side=tk.LEFT, padx=(0, 14), pady=4)
                except Exception:
                    if icon_char:
                        tk.Label(card, text=icon_char, font=(_df, 22), bg=card_bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 14), pady=4)
            elif icon_char:
                tk.Label(card, text=icon_char, font=(_df, 22), bg=card_bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 14), pady=4)
            text_frame = tk.Frame(card, bg=card_bg)
            text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            tk.Label(text_frame, text=title_text, font=(_df, 11, "bold"), bg=card_bg, fg="#263238").pack(anchor=tk.W)
            tk.Label(text_frame, text=desc_text, font=(_df, 9), bg=card_bg, fg="#78909c").pack(anchor=tk.W)
            btn = tk.Button(card, text="開く", font=(_df, 10), bg=btn_primary_bg, fg=btn_primary_fg,
                           activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
                           padx=16, pady=8, cursor="hand2", command=command)
            btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        _ai_et_icon_path = Path(__file__).resolve().parent.parent.parent / "assets" / "icons" / "ai_et_icon.png"
        add_card("\U0001f489", "AI/ET入力", "人工授精・胚移植の入力を行います", self._on_ai_et_input, icon_path=_ai_et_icon_path)
        add_card("\U0001f404", "分娩入力", "分娩イベントの入力を行います", self._on_calving_input)
        add_card("\U0001f4e5", "導入", "新規個体の導入登録を行います", self._on_introduction)
        add_card("\U0001f6aa", "退出", "売却・死亡廃用の入力を行います", self._on_exit)
        add_card("\U0001fa7a", "検診入力", "繁殖検診の入力を行います", self._on_checkup_input)
        add_card("\U0001f4ca", "レポート出力", "レポートを出力します", self._on_report_output)
        add_card("\U0001f4cb", "検診表出力", "繁殖検診表を出力します", self._on_checkup_table_output)
        add_card("\U0001f9ea", "必要妊娠シミュレーション", "年間必要妊娠頭数をシミュレーションします", self._on_required_pregnancy_simulation)
        add_card("\U0001f4dd", "検診ログの編集", "検診ログ（日誌）の内容を編集します", self._on_exam_log_editor)

        # カード・ボタン上でもマウスホイールでスクロールできるよう子孫にバインド
        _bind_mousewheel_recursive(content)
        
        content.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # フッター: 閉じる
        footer = tk.Frame(self.window, bg=bg, pady=20)
        footer.pack(fill=tk.X)
        tk.Button(footer, text="閉じる", font=(_df, 10), bg="#fafafa", fg="#546e7a",
                  activebackground="#eceff1", relief=tk.FLAT, padx=24, pady=10,
                  highlightbackground=btn_secondary_bd, highlightthickness=1, cursor="hand2",
                  command=self.window.destroy).pack()
    
    def _on_ai_et_input(self):
        """AI/ET入力ボタンをクリック"""
        from ui.event_input import EventInputWindow
        from modules.rule_engine import RuleEngine
        
        if self.event_dict_path is None:
            messagebox.showerror("エラー", "event_dictionary.json が見つかりません")
            return
        
        # AI/ET入力ウィンドウを開く（AIイベントまたはETイベントのみ）
        # デフォルトでAI(200)を選択（既存があれば再利用）
        event_input_window = EventInputWindow.open_or_focus(
            parent=self.window,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 汎用モード
            on_saved=None,
            farm_path=self.farm_path,
            allowed_event_numbers=[RuleEngine.EVENT_AI, RuleEngine.EVENT_ET],  # AI(200)またはET(201)のみ
            default_event_number=RuleEngine.EVENT_AI  # デフォルトでAI(200)を選択
        )
        event_input_window.show()
    
    def _on_calving_input(self):
        """分娩入力ボタンをクリック"""
        from ui.event_input import EventInputWindow
        from modules.rule_engine import RuleEngine
        
        if self.event_dict_path is None:
            messagebox.showerror("エラー", "event_dictionary.json が見つかりません")
            return
        
        # 分娩入力ウィンドウを開く（分娩イベントのみ、既存があれば再利用）
        event_input_window = EventInputWindow.open_or_focus(
            parent=self.window,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 汎用モード
            on_saved=None,
            farm_path=self.farm_path,
            allowed_event_numbers=[RuleEngine.EVENT_CALV],  # 分娩(202)のみ
            default_event_number=RuleEngine.EVENT_CALV  # デフォルトで分娩(202)を選択
        )
        event_input_window.show()
    
    def _on_introduction(self):
        """導入ボタンをクリック"""
        from ui.introduction_window import IntroductionWindow
        
        # 導入専用ウィンドウを開く（複数個体対応）
        intro_window = IntroductionWindow(
            parent=self.window,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path,
            event_dictionary_path=self.event_dict_path
        )
        intro_window.show()
    
    def _on_exit(self):
        """退出ボタンをクリック"""
        from ui.event_input import EventInputWindow
        from modules.rule_engine import RuleEngine
        
        if self.event_dict_path is None:
            messagebox.showerror("エラー", "event_dictionary.json が見つかりません")
            return
        
        # 退出入力ウィンドウを開く（売却または死廃イベントのみ、既存があれば再利用）
        event_input_window = EventInputWindow.open_or_focus(
            parent=self.window,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 汎用モード
            on_saved=None,
            farm_path=self.farm_path,
            allowed_event_numbers=[RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD]  # 売却(205)または死廃(206)のみ
        )
        event_input_window.show()
    
    def _on_checkup_input(self):
        """検診入力ボタンをクリック"""
        from ui.reproduction_checkup_input_window import ReproductionCheckupInputWindow
        from ui.date_picker_window import DatePickerWindow
        from datetime import datetime
        
        if self.event_dict_path is None:
            messagebox.showerror("エラー", "event_dictionary.json が見つかりません")
            return
        
        # 検診日入力ダイアログ
        dialog = tk.Toplevel(self.window)
        dialog.title("検診日を入力")
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="検診日を入力してください").pack(anchor=tk.W)

        date_frame = ttk.Frame(frame)
        date_frame.pack(fill=tk.X, pady=(8, 0))

        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        date_entry = ttk.Entry(date_frame, textvariable=date_var, width=12)
        date_entry.pack(side=tk.LEFT, padx=(0, 5))

        def on_calendar():
            def on_date_selected(date_str: str):
                date_var.set(date_str)
            picker = DatePickerWindow(parent=dialog, on_date_selected=on_date_selected)
            picker.show()

        ttk.Button(date_frame, text="📅", width=3, command=on_calendar).pack(side=tk.LEFT)

        def on_ok():
            date_str = date_var.get().strip()
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "検診日の形式が正しくありません（YYYY-MM-DD形式）。")
                return

            dialog.destroy()
            input_window = ReproductionCheckupInputWindow(
            parent=self.window,
            db_handler=self.db,
                formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
                farm_path=self.farm_path,
            event_dictionary_path=self.event_dict_path,
                checkup_date=date_str,
            )
            input_window.show()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_exam_log_editor(self):
        """検診ログの編集ボタンをクリック"""
        from ui.reproduction_checkup_exam_log_editor_window import ReproductionCheckupExamLogEditorWindow
        editor = ReproductionCheckupExamLogEditorWindow(
            parent=self.window,
            farm_path=self.farm_path,
            flow_window=self,
        )
        editor.show()
    
    def _on_report_output(self):
        """レポート出力ボタンをクリック → レポート/請求選択ダイアログを表示"""
        from datetime import datetime
        from ui.date_picker_window import DatePickerWindow
        
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        btn_primary_bg = "#3949ab"
        btn_primary_fg = "#ffffff"
        btn_secondary_bg = "#fafafa"
        btn_secondary_fg = "#546e7a"
        btn_secondary_bd = "#b0bec5"
        
        dialog = tk.Toplevel(self.window)
        dialog.title("レポート出力")
        dialog.geometry("440x380")
        dialog.configure(bg=bg)
        dialog.minsize(400, 300)
        
        # ヘッダー（イメージ統一）
        header = tk.Frame(dialog, bg=bg, pady=20, padx=24)
        header.pack(fill=tk.X)
        tk.Label(header, text="\U0001f4ca", font=(_df, 24), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="レポート出力", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="出力する種類を選択し、検診日を設定してください", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)
        
        # コンテンツ: チェックボックス + 検診日
        content = tk.Frame(dialog, bg=bg, padx=24, pady=16)
        content.pack(fill=tk.BOTH, expand=True)
        
        report_var = tk.BooleanVar(value=True)
        billing_var = tk.BooleanVar(value=True)
        log_var = tk.BooleanVar(value=False)
        
        report_cb = tk.Checkbutton(
            content, text="レポート", variable=report_var,
            font=(_df, 11), bg=bg, fg="#263238", activebackground=bg, activeforeground="#263238",
            selectcolor="#e8eaf6", highlightthickness=0
        )
        report_cb.pack(anchor=tk.W, pady=(0, 12))
        billing_cb = tk.Checkbutton(
            content, text="請求", variable=billing_var,
            font=(_df, 11), bg=bg, fg="#263238", activebackground=bg, activeforeground="#263238",
            selectcolor="#e8eaf6", highlightthickness=0
        )
        billing_cb.pack(anchor=tk.W, pady=(0, 12))
        log_cb = tk.Checkbutton(
            content, text="検診ログ", variable=log_var,
            font=(_df, 11), bg=bg, fg="#263238", activebackground=bg, activeforeground="#263238",
            selectcolor="#e8eaf6", highlightthickness=0
        )
        log_cb.pack(anchor=tk.W, pady=(0, 16))
        
        # 検診日
        date_label = tk.Label(content, text="検診日", font=(_df, 11), bg=bg, fg="#263238")
        date_label.pack(anchor=tk.W, pady=(0, 4))
        date_frame = tk.Frame(content, bg=bg)
        date_frame.pack(fill=tk.X)
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        date_entry = tk.Entry(date_frame, textvariable=date_var, font=(_df, 11), width=12)
        date_entry.pack(side=tk.LEFT, padx=(0, 8))
        def _on_calendar():
            def on_date_selected(date_str: str):
                date_var.set(date_str)
            try:
                datetime.strptime(date_var.get().strip(), "%Y-%m-%d")
                initial = date_var.get().strip()
            except ValueError:
                initial = datetime.now().strftime("%Y-%m-%d")
            picker = DatePickerWindow(parent=dialog, initial_date=initial, on_date_selected=on_date_selected)
            picker.show()
        tk.Button(
            date_frame, text="📅", font=(_df, 10), width=3,
            bg=btn_secondary_bg, fg=btn_secondary_fg, relief=tk.FLAT,
            highlightbackground=btn_secondary_bd, highlightthickness=1,
            cursor="hand2", command=_on_calendar
        ).pack(side=tk.LEFT)
        
        # フッター: OK / キャンセル
        footer = tk.Frame(dialog, bg=bg, pady=20)
        footer.pack(fill=tk.X)
        btn_frame = tk.Frame(footer, bg=bg)
        btn_frame.pack()
        
        def _on_ok():
            date_str = date_var.get().strip()
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("エラー", "検診日の形式が正しくありません（YYYY-MM-DD形式）。")
                return
            dialog.destroy()
            self._on_report_output_confirmed(report_var.get(), billing_var.get(), log_var.get(), date_str)
        
        def _on_cancel():
            dialog.destroy()
        
        tk.Button(
            btn_frame, text="OK", font=(_df, 10),
            bg=btn_primary_bg, fg=btn_primary_fg,
            activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
            padx=24, pady=10, cursor="hand2", command=_on_ok
        ).pack(side=tk.LEFT, padx=6)
        tk.Button(
            btn_frame, text="キャンセル", font=(_df, 10),
            bg=btn_secondary_bg, fg=btn_secondary_fg,
            activebackground="#eceff1", relief=tk.FLAT,
            padx=24, pady=10, highlightbackground=btn_secondary_bd, highlightthickness=1,
            cursor="hand2", command=_on_cancel
        ).pack(side=tk.LEFT, padx=6)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_report_output_confirmed(self, report_checked: bool, billing_checked: bool, log_checked: bool, checkup_date: str):
        """レポート出力でOK押下後。請求HTML／レポートHTML／検診ログHTMLを出力する。"""
        if not report_checked and not billing_checked and not log_checked:
            messagebox.showinfo("情報", "レポート・請求・検診ログのいずれかにチェックを入れてください。")
            return
        if billing_checked:
            self._output_billing_html(checkup_date)
        if report_checked:
            self._output_report_html(checkup_date)
        if log_checked:
            self._output_exam_log_html()

    def _output_billing_html(self, checkup_date: str):
        """請求HTMLを生成し、既定ブラウザで表示する。"""
        import tempfile
        import webbrowser
        from modules.reproduction_checkup_billing import build_billing_html
        from settings_manager import SettingsManager

        farm_name = SettingsManager(self.farm_path).get("farm_name", self.farm_path.name)
        html_content = build_billing_html(
            db=self.db,
            farm_path=self.farm_path,
            checkup_date=checkup_date,
            farm_name=farm_name,
        )
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", suffix=".html", delete=False
            ) as f:
                f.write(html_content)
                temp_path = f.name
            webbrowser.open(f"file://{temp_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"HTMLの表示に失敗しました。\n{e}")

    def _output_report_html(self, checkup_date: str):
        """繁殖検診レポート（HTML）を生成し、既定ブラウザで表示する。"""
        import tempfile
        import webbrowser
        from modules.reproduction_checkup_report import build_report_html
        from settings_manager import SettingsManager

        try:
            farm_name = SettingsManager(self.farm_path).get("farm_name", self.farm_path.name)
            html_content = build_report_html(
                db=self.db,
                rule_engine=self.rule_engine,
                farm_path=self.farm_path,
                checkup_date=checkup_date,
                farm_name=farm_name,
            )
        except Exception as e:
            messagebox.showerror("エラー", f"レポートの生成に失敗しました。\n{e}")
            return
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", suffix=".html", delete=False
            ) as f:
                f.write(html_content)
                temp_path = f.name
            webbrowser.open(f"file://{temp_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"レポートHTMLの表示に失敗しました。\n{e}")

    def _get_monthly_stats(self, year: int, month: int) -> Dict[str, Any]:
        """指定月の成績を取得（分娩頭数・授精頭数・妊娠頭数・経産牛受胎率）。

        妊娠頭数はその月の妊娠イベント数ではなく、その月の授精（AI/ET）のうち
        妊娠に至った授精の数として集計する。
        """
        _, last_day = calendar.monthrange(year, month)
        month_start = f"{year}-{month:02d}-01"
        month_end = f"{year}-{month:02d}-{last_day:02d}"
        calving = len(
            self.db.get_events_by_number_and_period(
                self.rule_engine.EVENT_CALV, month_start, month_end, include_deleted=False
            )
        )
        ai_events = self.db.get_events_by_number_and_period(
            self.rule_engine.EVENT_AI, month_start, month_end, include_deleted=False
        )
        et_events = self.db.get_events_by_number_and_period(
            self.rule_engine.EVENT_ET, month_start, month_end, include_deleted=False
        )
        insemination = len(ai_events) + len(et_events)
        preg_plus = (self.rule_engine.EVENT_PDP, self.rule_engine.EVENT_PDP2, self.rule_engine.EVENT_PAGP)
        preg_minus_or_next = (
            self.rule_engine.EVENT_PDN, self.rule_engine.EVENT_PAGN, self.rule_engine.EVENT_ABRT,
            self.rule_engine.EVENT_AI, self.rule_engine.EVENT_ET, self.rule_engine.EVENT_CALV,
        )
        # 妊娠数：その月の授精（AI/ET）のうち妊娠に至った授精数
        pregnancy = 0
        # 経産牛（event_lact >= 1）の授精のうち受胎した割合
        conceived = 0
        total_keisan = 0
        for ev in ai_events + et_events:
            lact = ev.get("event_lact")
            if lact is None:
                lact = 0
            else:
                try:
                    lact = int(lact)
                except (TypeError, ValueError):
                    lact = 0
            cow_auto_id = ev.get("cow_auto_id")
            event_date = ev.get("event_date") or ""
            cow_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            cow_events.sort(key=lambda e: (e.get("event_date", ""), e.get("id", 0)))
            # この授精の後の最初の関連イベントで妊娠／非妊娠を判定
            resulted_in_pregnancy = False
            for e in cow_events:
                d = e.get("event_date") or ""
                if d <= event_date:
                    continue
                num = e.get("event_number")
                if num in preg_plus:
                    resulted_in_pregnancy = True
                    break
                if num in preg_minus_or_next:
                    break
            if resulted_in_pregnancy:
                pregnancy += 1
            # 経産牛のみ受胎率分母に含める
            if lact >= 1:
                total_keisan += 1
                if resulted_in_pregnancy:
                    conceived += 1
        rate = (conceived / total_keisan * 100) if total_keisan else None
        return {
            "calving": calving,
            "insemination": insemination,
            "pregnancy": pregnancy,
            "conception_rate": rate,
            "conception_denom": total_keisan,
        }

    def _output_exam_log_html(self):
        """検診ログ（直近4か月分）をHTMLで生成し、ブラウザで表示する。"""
        path = self.farm_path / EXAM_LOG_NOTES_FILENAME
        if not path.exists():
            messagebox.showinfo("検診ログ", "検診ログのデータがありません。")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("エラー", f"検診ログの読み込みに失敗しました。\n{e}")
            return
        if not isinstance(data, dict):
            messagebox.showinfo("検診ログ", "検診ログのデータがありません。")
            return

        end_date = datetime.now().date()
        start_date = end_date - timedelta(days=4 * 31)
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")

        date_entries: List[tuple] = []
        for date_str, day_data in data.items():
            if not isinstance(day_data, dict):
                continue
            try:
                if start_str <= date_str <= end_str:
                    date_entries.append((date_str, day_data))
            except (TypeError, ValueError):
                continue
        date_entries.sort(key=lambda x: x[0])

        months_seen: Dict[str, bool] = {}
        for date_str, _ in date_entries:
            try:
                y, m = date_str[:4], int(date_str[5:7])
                key = f"{y}-{m:02d}"
                if key not in months_seen:
                    months_seen[key] = True
            except (IndexError, ValueError):
                pass
        month_stats: Dict[str, Dict[str, Any]] = {}
        for key in months_seen.keys():
            try:
                y, m = int(key[:4]), int(key[5:7])
                month_stats[key] = self._get_monthly_stats(y, m)
            except (ValueError, IndexError) as e:
                logging.warning("exam_log month_stats skip %s: %s", key, e)

        from settings_manager import SettingsManager
        farm_name = SettingsManager(self.farm_path).get("farm_name", self.farm_path.name)
        html_content = self._build_exam_log_html(
            date_entries=date_entries,
            farm_name=farm_name,
            month_stats=month_stats,
        )
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", suffix=".html", delete=False
            ) as f:
                f.write(html_content)
                temp_path = f.name
            webbrowser.open(f"file://{temp_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"検診ログHTMLの表示に失敗しました。\n{e}")

    def _build_exam_log_html(
        self, date_entries: List[tuple], farm_name: str, month_stats: Optional[Dict[str, Dict[str, Any]]] = None
    ) -> str:
        """検診ログ用HTMLを組み立てる（1月ログ→1月成績→2月ログ→2月成績…の順で1本のタイムライン）。"""
        month_stats = month_stats or {}
        color_map = {"good": "#2e7d32", "normal": "#f9a825", "caution": "#c62828"}
        # 月ごとにグループ化（キーは YYYY-MM、値は (date_str, day_data) のリスト）
        by_month: Dict[str, List[tuple]] = {}
        for date_str, day_data in date_entries:
            try:
                key = f"{date_str[:4]}-{int(date_str[5:7]):02d}"
                if key not in by_month:
                    by_month[key] = []
                by_month[key].append((date_str, day_data))
            except (IndexError, ValueError):
                pass
        month_keys_ordered = sorted(by_month.keys())

        sb: List[str] = []
        sb.append("<!DOCTYPE html><html><head><meta charset='utf-8'>")
        sb.append("<title>検診ログ</title>")
        sb.append("<style>")
        sb.append("body{font-family:'Meiryo UI',sans-serif;margin:24px;background:#fff;color:#333;")
        sb.append("-webkit-print-color-adjust:exact;print-color-adjust:exact;}")
        sb.append(".header{margin-bottom:28px;}")
        sb.append(".header h1{font-size:1.35em;color:#263238;margin:0 0 8px 0;}")
        sb.append(".header .meta{font-size:0.9em;color:#607d8b;}")
        sb.append(".timeline-wrap{max-width:720px;position:relative;padding-left:56px;}")
        sb.append(".timeline-bar{position:absolute;left:11px;top:12px;bottom:12px;width:6px;background:#c4a574;border-radius:3px;}")
        sb.append(".timeline-item{position:relative;padding-bottom:28px;}")
        sb.append(".timeline-item:last-child{padding-bottom:0;}")
        sb.append(".timeline-dot{position:absolute;left:-52px;top:10px;width:20px;height:20px;border-radius:50%;background:#fff;border:3px solid #c4a574;box-sizing:border-box;z-index:1;}")
        sb.append(".timeline-dot.month-dot{background:#1565c0;border-color:#0d47a1;}")
        sb.append("/* 日誌ログ：白系・余白多めで記録らしさを強調 */")
        sb.append(".timeline-content{background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px 18px;margin-left:4px;box-shadow:0 1px 3px rgba(0,0,0,0.06);}")
        sb.append(".timeline-content .log-badge{display:inline-block;font-size:0.7em;color:#757575;background:#f5f5f5;padding:2px 8px;border-radius:4px;margin-bottom:6px;letter-spacing:0.05em;}")
        sb.append(".timeline-date{font-size:1em;font-weight:bold;color:#424242;margin-bottom:10px;}")
        sb.append(".notes-list{margin:0 0 10px 0;padding-left:20px;color:#37474f;font-size:0.95em;line-height:1.6;}")
        sb.append(".notes-list li{margin-bottom:4px;}")
        sb.append(".ratings{display:flex;flex-wrap:wrap;gap:10px 16px;margin-top:8px;}")
        sb.append(".rating-item{display:inline-flex;align-items:center;gap:5px;font-size:0.85em;color:#5d4e37;}")
        sb.append(".rating-dot{width:12px;height:12px;border-radius:50%;flex-shrink:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}")
        sb.append(".timeline-bar,.timeline-dot,.timeline-content,.summary-block{-webkit-print-color-adjust:exact;print-color-adjust:exact;}")
        sb.append("/* 月別集計：青系・データらしさを強調しログと明確に区別 */")
        sb.append(".summary-block{background:linear-gradient(135deg,#e3f2fd 0%,#e8f4f8 100%);border:1px solid #90caf9;border-left:4px solid #1565c0;border-radius:8px;padding:16px 20px;box-shadow:0 2px 4px rgba(21,101,192,0.12);}")
        sb.append(".summary-block .month-label{font-size:0.75em;color:#0d47a1;font-weight:600;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:6px;}")
        sb.append(".summary-block .month{font-weight:bold;color:#1565c0;margin-bottom:10px;font-size:1.1em;}")
        sb.append(".summary-row{display:flex;flex-wrap:wrap;align-items:center;gap:8px 1.2em;font-size:0.9em;color:#37474f;}")
        sb.append(".summary-row .stat{white-space:nowrap;}")
        sb.append("@media print{.timeline-dot{border:2px solid #5d4e37;} .timeline-dot.month-dot{background:#1565c0;border-color:#0d47a1;} body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}")
        sb.append("</style></head><body>")
        sb.append("<div class='header'>")
        sb.append(f"<h1>検診ログ（直近4か月）</h1>")
        sb.append(f"<p class='meta'>{html_module.escape(farm_name)}　出力: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>")
        sb.append("</div>")
        sb.append("<div class='timeline-wrap'>")
        sb.append("<div class='timeline-bar' aria-hidden='true'></div>")

        for month_key in month_keys_ordered:
            entries = by_month.get(month_key, [])
            # その月のログを連続で出力
            for date_str, day_data in entries:
                ratings = day_data.get("ratings") or {}
                notes = day_data.get("notes") or []
                if not isinstance(notes, list):
                    notes = []
                y, m, d = date_str.split("-")[0], str(int(date_str.split("-")[1])), str(int(date_str.split("-")[2]))
                date_display = f"{y}.{m}.{d}"
                sb.append("<div class='timeline-item'>")
                sb.append("<div class='timeline-dot' aria-hidden='true'></div>")
                sb.append("<div class='timeline-content'>")
                sb.append("<span class='log-badge'>日誌</span>")
                sb.append(f"<div class='timeline-date'>{html_module.escape(date_display)}</div>")
                if notes:
                    sb.append("<ul class='notes-list'>")
                    for line in notes:
                        if line and line.strip():
                            sb.append(f"<li>{html_module.escape(line.strip())}</li>")
                    sb.append("</ul>")
                sb.append("<div class='ratings'>")
                for key, label in EXAM_LOG_RATING_LABELS:
                    val = ratings.get(key)
                    if not val:
                        continue
                    color = color_map.get(val, "#9e9e9e")
                    sb.append(
                        f"<span class='rating-item'>"
                        f"<span class='rating-dot' style='background:{color};'></span>"
                        f"{html_module.escape(label)}</span>"
                    )
                sb.append("</div>")
                sb.append("</div></div>")
            # その月の成績ブロックを直後に出力（横並び: 分娩　授精　妊娠　受胎率）
            stats = month_stats.get(month_key, {})
            calving = stats.get("calving", 0)
            insem = stats.get("insemination", 0)
            pregnancy = stats.get("pregnancy", 0)
            rate = stats.get("conception_rate")
            try:
                month_title = f"{int(month_key[5:7])}月"
            except (ValueError, IndexError):
                month_title = month_key
            sb.append("<div class='timeline-item'>")
            sb.append("<div class='timeline-dot month-dot' aria-hidden='true'></div>")
            sb.append("<div class='timeline-content summary-block'>")
            sb.append("<div class='month-label'>月別集計</div>")
            sb.append(f"<div class='month'>{html_module.escape(month_title)}</div>")
            sb.append("<div class='summary-row'>")
            sb.append(f"<span class='stat'>分娩 {calving} 頭</span>")
            sb.append(f"<span class='stat'>授精 {insem} 頭</span>")
            sb.append(f"<span class='stat'>妊娠 {pregnancy} 頭</span>")
            if rate is not None:
                sb.append(f"<span class='stat'>受胎率（経産） {rate:.0f}%</span>")
            else:
                sb.append("<span class='stat'>受胎率（経産） ―%</span>")
            sb.append("</div>")
            sb.append("</div></div>")
        sb.append("</div></body></html>")
        return "".join(sb)

    def _on_checkup_table_output(self):
        """検診表出力ボタンをクリック"""
        from ui.reproduction_checkup_window import ReproductionCheckupWindow
        
        # 繁殖検診表ウィンドウを開く
        # item_dictionary.json は FALCON側で一括管理（Noneを渡すと自動的にconfig_default/から読み込む）
        repro_checkup_window = ReproductionCheckupWindow(
            parent=self.window,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            farm_path=self.farm_path,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=None,
            rule_engine=self.rule_engine
        )
        repro_checkup_window.show()

    def _on_required_pregnancy_simulation(self):
        """必要妊娠シミュレーション（年間必要妊娠頭数）ボタンをクリック"""
        from ui.reproduction_checkup_report_window import ReproductionCheckupReportWindow

        report_window = ReproductionCheckupReportWindow(
            parent=self.window,
            db=self.db,
            rule_engine=self.rule_engine,
            farm_path=self.farm_path,
        )
        report_window.show()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

