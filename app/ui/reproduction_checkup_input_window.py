"""
FALCON2 - 繁殖検診入力ウインドウ
検診日を指定して抽出し、個体ごとに入力を行う
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List
from datetime import datetime
import json
import logging

VISIT_NOTES_FILENAME = "reproduction_checkup_visit_notes.json"
VISIT_RATING_KEYS = [
    ("overall", "全体"),
    ("cow_condition", "牛の状態"),
    ("feed", "飼料"),
    ("disease", "疾病"),
    ("milk_volume", "乳量"),
    ("milk_quality", "乳質"),
    ("reproduction", "繁殖"),
    ("calf_rearing", "子牛・育成"),
]
VISIT_NOTES_LINES = 8  # 全体＋7部門と1行ずつ対応

from db.db_handler import DBHandler
from modules.activity_log import record_from_event
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.reproduction_checkup_logic import ReproductionCheckupLogic
from modules.reproduction_checkup_billing import REPRO_CHECKUP_EVENT_NUMBERS
from settings_manager import SettingsManager


class ReproductionCheckupInputWindow:
    """繁殖検診入力ウインドウ"""

    EXEC_CODE_EVENT_MAP = {
        "2": 300,  # フレッシュ
        "3": 301,  # 繁殖検査
        "4": RuleEngine.EVENT_PDN,   # 妊娠鑑定マイナス
        "5": RuleEngine.EVENT_PDP,   # 妊娠鑑定プラス
        "6": RuleEngine.EVENT_PDP2,  # 妊娠鑑定プラス（直近以外）
        "7": RuleEngine.EVENT_PAGN,  # PAGマイナス
        "8": RuleEngine.EVENT_PAGP,  # PAGプラス
    }

    INPUT_COLUMNS = [
        "exec_code",
        "treatment",
        "uterine_findings",
        "right_ovary_findings",
        "left_ovary_findings",
        "other",
        "twin",  # 双子チェックボックス
        "female_judgment",  # ♀♂判定の♀
        "male_judgment",   # ♀♂判定の♂
        "bcs",  # BCS（体型スコア）
    ]

    def __init__(
        self,
        parent: tk.Widget,
        db_handler: DBHandler,
        formula_engine: FormulaEngine,
        rule_engine: RuleEngine,
        farm_path: Path,
        event_dictionary_path: Optional[Path],
        checkup_date: str,
    ):
        self.parent = parent
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        self.event_dict_path = event_dictionary_path
        self.checkup_date = checkup_date

        self.settings_manager = SettingsManager(self.farm_path)
        self.checkup_logic = ReproductionCheckupLogic(db_handler, event_dictionary_path)

        # 繁殖処置設定
        self.treatments: Dict[str, Dict[str, Any]] = {}
        self._load_treatment_settings()

        # 抽出結果
        self.results: List[Dict[str, Any]] = []
        # Treeview行データ
        self.row_data: Dict[str, Dict[str, Any]] = {}
        # 表示行（ソート対象）
        self.display_rows: List[Dict[str, Any]] = []
        self.original_order: List[str] = []
        self.sort_state: Dict[str, Optional[str]] = {}
        # 入力ウィジェット
        self.input_widgets: Dict[str, Dict[str, tk.Entry]] = {}
        self.input_checkboxes: Dict[str, tk.BooleanVar] = {}  # 双子・♀判定チェックボックス用
        self.input_entries_order: List[tk.Entry] = []
        # 現在ハイライトされている行ID
        self.current_highlighted_row_id: Optional[str] = None
        # 検診訪問メモ（7部門の評価＋6行箇条書き）
        self.visit_ratings: Dict[str, Optional[str]] = {}  # key -> "good"|"normal"|"caution"|None
        self.visit_notes: List[str] = [""] * VISIT_NOTES_LINES
        self.visit_rating_buttons: Dict[str, Dict[str, tk.Widget]] = {}
        self.visit_note_entries: List[tk.Entry] = []

        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖検診入力")
        self.window.geometry("1280x900")
        self.window.minsize(1000, 750)
        try:
            self.window.state("zoomed")  # 開いたときに最大化（Windows）
        except tk.TclError:
            try:
                self.window.attributes("-zoomed", True)  # Linux等
            except tk.TclError:
                pass
        try:
            self.window.state("zoomed")  # 開いたときに最大化（Windows）
        except tk.TclError:
            try:
                self.window.attributes("-zoomed", True)  # Linux等
            except tk.TclError:
                pass
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        self._create_widgets()
        self._extract_and_display()

        # 最大化できなかった場合のみ中央配置
        try:
            if self.window.state() != "zoomed":
                self.window.update_idletasks()
                x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
                y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
                self.window.geometry(f"+{x}+{y}")
        except tk.TclError:
            self.window.update_idletasks()
            x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
            y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
            self.window.geometry(f"+{x}+{y}")

    def _load_treatment_settings(self):
        """reproduction_treatment_settings.json をロード"""
        settings_file = self.farm_path / "reproduction_treatment_settings.json"
        if not settings_file.exists():
            return

        try:
            with open(settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            treatments_raw = data.get("treatments", {})
            self.treatments = {str(k): v for k, v in treatments_raw.items()}
        except Exception as e:
            logging.error(f"繁殖処置設定読み込みエラー: {e}")
            self.treatments = {}

    def _convert_to_uppercase_on_input(self, event, widget: tk.Widget):
        """入力時にリアルタイムで半角大文字に変換（イベント入力と同様、全角日本語は保持）"""
        if not isinstance(widget, tk.Entry):
            return
        char = event.char
        if not char or len(char) != 1:
            return
        if "a" <= char <= "z":
            cursor_pos = widget.index(tk.INSERT)
            current_text = widget.get()
            new_text = current_text[:cursor_pos] + char.upper() + current_text[cursor_pos:]
            widget.delete(0, tk.END)
            widget.insert(0, new_text)
            widget.icursor(cursor_pos + 1)
            return "break"

    def _convert_to_uppercase(self, widget: tk.Widget):
        """入力されたテキストを半角大文字に変換（全角日本語は保持）"""
        if not isinstance(widget, tk.Entry):
            return
        current_text = widget.get()
        cursor_pos = widget.index(tk.INSERT)
        converted = ""
        for char in current_text:
            if ord("Ａ") <= ord(char) <= ord("Ｚ"):
                converted += chr(ord(char) - ord("Ａ") + ord("A"))
            elif ord("ａ") <= ord(char) <= ord("ｚ"):
                converted += chr(ord(char) - ord("ａ") + ord("A"))
            elif ord("０") <= ord(char) <= ord("９"):
                converted += chr(ord(char) - ord("０") + ord("0"))
            elif "a" <= char <= "z":
                converted += char.upper()
            elif "A" <= char <= "Z" or "0" <= char <= "9":
                converted += char
            else:
                converted += char
        if converted != current_text:
            widget.delete(0, tk.END)
            widget.insert(0, converted)
            try:
                widget.icursor(min(cursor_pos, len(converted)))
            except Exception:
                pass

    def _create_widgets(self):
        # 全体をスクロール可能にするコンテナ
        container = ttk.Frame(self.window)
        container.pack(fill=tk.BOTH, expand=True)
        self.main_canvas = tk.Canvas(container, bg="#f5f5f5", highlightthickness=0)
        main_scrollbar = ttk.Scrollbar(container)
        main_frame = ttk.Frame(self.main_canvas, padding=10)
        self.main_frame_id = self.main_canvas.create_window((0, 0), window=main_frame, anchor=tk.NW)

        def _on_main_configure(event):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

        def _on_canvas_configure(event):
            self.main_canvas.itemconfig(self.main_frame_id, width=event.width)

        main_frame.bind("<Configure>", _on_main_configure)
        self.main_canvas.bind("<Configure>", _on_canvas_configure)

        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.main_canvas.configure(yscrollcommand=main_scrollbar.set)
        main_scrollbar.configure(command=self.main_canvas.yview)

        def _on_main_wheel(event):
            self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_main_wheel_linux(event, direction):
            self.main_canvas.yview_scroll(direction, "units")

        self.main_canvas.bind("<MouseWheel>", _on_main_wheel)
        self.main_canvas.bind("<Button-4>", lambda e: _on_main_wheel_linux(e, -1))
        self.main_canvas.bind("<Button-5>", lambda e: _on_main_wheel_linux(e, 1))
        self.main_canvas.bind("<Enter>", lambda e: (
            self.main_canvas.bind_all("<MouseWheel>", _on_main_wheel),
            self.main_canvas.bind_all("<Button-4>", lambda ev: _on_main_wheel_linux(ev, -1)),
            self.main_canvas.bind_all("<Button-5>", lambda ev: _on_main_wheel_linux(ev, 1)),
        ))
        self.main_canvas.bind("<Leave>", lambda e: (
            self.main_canvas.unbind_all("<MouseWheel>"),
            self.main_canvas.unbind_all("<Button-4>"),
            self.main_canvas.unbind_all("<Button-5>"),
        ))

        title_label = ttk.Label(
            main_frame,
            text=f"繁殖検診入力（検診日: {self.checkup_date}）",
            font=("", 12, "bold"),
        )
        title_label.pack(anchor=tk.W, pady=(0, 8))

        # ID検索欄
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(anchor=tk.W, pady=(0, 6), fill=tk.X)
        
        ttk.Label(search_frame, text="ID検索:", font=("", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = tk.Entry(search_frame, width=15, font=("", 9))
        self.search_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry.bind("<Return>", self._on_search_id)
        self.search_entry.bind("<KeyRelease>", self._on_search_key_release)
        
        ttk.Button(search_frame, text="検索", command=self._on_search_id, width=8).pack(side=tk.LEFT)

        # 実施コードの説明（表の外・表の直上に表示、濃い色で視認性を確保）
        code_hint_frame = ttk.Frame(main_frame)
        code_hint_frame.pack(anchor=tk.W, pady=(4, 6), fill=tk.X)
        code_hint_text = "実施コード: 1=スキップ, 2=フレッシュ, 3=繁殖検査, 4=妊娠マイナス, 5=妊娠プラス, 6=直近以外プラス, 7=PAGマイナス, 8=PAGプラス, 9=その他"
        code_hint_lbl = tk.Label(
            code_hint_frame,
            text=code_hint_text,
            font=("", 9),
            fg="#263238",
            anchor="w",
        )
        code_hint_lbl.pack(anchor=tk.W)

        # テーブル（左：提示／右：入力）
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        left_columns = [
            "cow_id",
            "jpn10",
            "dim",
            "dai",
            "last_checkup_result",
            "checkup_code",
            "cow_id_dup",
        ]
        right_columns = [
            "exec_code",
            "treatment",
            "uterine_findings",
            "right_ovary_findings",
            "left_ovary_findings",
            "other",
        ]

        self.column_labels = {
            "cow_id": "ID",
            "jpn10": "JPN10",
            "dim": "DIM",
            "dai": "授精後",
            "checkup_code": "検診種類",
            "last_checkup_result": "前回検診結果",
            "cow_id_dup": "ID",
            "exec_code": "実施コード",
            "treatment": "処置",  # 農場設定＞繁殖処置設定の処置（WPG, CIDR 等）
            "uterine_findings": "子宮",  # 所見
            "right_ovary_findings": "右",  # 所見
            "left_ovary_findings": "左",  # 所見
            "other": "その他",  # 所見
        }

        # 左テーブル（表示行数を抑えて記録・信号まで1画面に収める）
        style = ttk.Style(self.window)
        style.configure("CheckupLeft.Treeview", rowheight=24)
        # ハイライト用のタグスタイル
        style.configure("CheckupLeft.Treeview", background="white")
        style.map("CheckupLeft.Treeview", background=[("selected", "#4a90e2")])
        self.left_tree = ttk.Treeview(
            table_frame, columns=left_columns, show="headings", height=14, style="CheckupLeft.Treeview"
        )
        # ハイライト用のタグを設定
        self.left_tree.tag_configure("highlighted", background="#b3d9ff")
        left_widths = {
            "cow_id": 40,
            "jpn10": 72,
            "dim": 38,
            "dai": 45,
            "checkup_code": 65,
            "last_checkup_result": 110,
            "cow_id_dup": 40,
        }
        for col in left_columns:
            self.left_tree.heading(
                col,
                text=self.column_labels.get(col, col),
                command=lambda c=col: self._on_column_click(c),
            )
            self.left_tree.column(col, width=left_widths.get(col, 80))

        # 右テーブル（入力欄は常時ボックス表示）
        right_frame = ttk.Frame(table_frame)

        header_frame = ttk.Frame(right_frame, height=24)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        self.right_column_specs = [
            ("exec_code", "実施コード", 76),
            ("treatment", "処置", 84),
            ("uterine_findings", "子宮", 76),
            ("right_ovary_findings", "右", 66),
            ("left_ovary_findings", "左", 66),
            ("other", "その他", 84),
            ("twin", "双子", 48),  # 双子チェックボックス
            ("female_judgment", "♀", 38),  # ♀♂判定の♀
            ("male_judgment", "♂", 38),    # ♀♂判定の♂
            ("bcs", "BCS", 42),
        ]

        self.right_total_width = sum(width for _k, _l, width in self.right_column_specs)
        header_frame.configure(width=self.right_total_width)
        header_frame.pack_propagate(False)

        x_offset = 0
        for col_key, label, width_px in self.right_column_specs:
            lbl = tk.Label(
                header_frame,
                text=label,
                relief="solid",
                bd=1,
                bg="#f2e0a6",
            )
            lbl.place(x=x_offset, y=0, width=width_px, height=24)
            x_offset += width_px

        self.input_canvas = tk.Canvas(
            right_frame, bg="#fff7d6", highlightthickness=0, width=self.right_total_width
        )
        self.input_canvas.configure(yscrollincrement=24)
        self.input_canvas.pack(fill=tk.BOTH, expand=True)

        self.input_inner = ttk.Frame(self.input_canvas)
        self.input_canvas_window = self.input_canvas.create_window(
            (0, 0), window=self.input_inner, anchor="nw"
        )

        self.input_inner.bind(
            "<Configure>",
            lambda e: self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all")),
        )
        self.input_canvas.bind("<Configure>", self._on_input_canvas_configure)

        # スクロールバー（左右共通）
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def yview(*args):
            self.left_tree.yview(*args)
            self.input_canvas.yview(*args)
        
        # yview関数をインスタンス変数として保存（後で使用するため）
        self.yview_func = yview
        
        # 入力フレームにもマウスホイールをバインド
        self.input_inner.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        self.input_inner.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        self.input_inner.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))

        scrollbar.config(command=yview)
        self.left_tree.configure(yscrollcommand=scrollbar.set)
        self.input_canvas.configure(yscrollcommand=scrollbar.set)

        self.left_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH)

        # マウスホイール同期
        self.left_tree.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        self.input_canvas.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        # ウィンドウ全体でマウスホイールを有効化
        self.window.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        main_frame.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        table_frame.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        right_frame.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
        # Linux対応（Button-4: 上スクロール、Button-5: 下スクロール）
        self.left_tree.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        self.left_tree.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))
        self.input_canvas.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        self.input_canvas.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))
        self.window.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        self.window.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))
        main_frame.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        main_frame.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))
        table_frame.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        table_frame.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))
        right_frame.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
        right_frame.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))

        # 検診訪問メモ（表の下）＋保存・閉じるボタン
        self._create_visit_notes_section(main_frame)

    def _create_visit_notes_section(self, parent: tk.Widget):
        """検診訪問メモ：左に部門・信号8行、右に気づいた点（箇条書き）8行を縦揃えで表示。メモは部門対応なしの一般箇条書き。"""
        frame = ttk.LabelFrame(parent, text="記録", padding=8)
        frame.pack(fill=tk.X, pady=(8, 0))

        inner = ttk.Frame(frame)
        inner.pack(fill=tk.BOTH, expand=True)
        # 左列が圧縮されないよう grid で列幅を確保（信号が消えないように）
        inner.grid_columnconfigure(0, minsize=260, weight=0)
        inner.grid_columnconfigure(1, minsize=200, weight=1)
        inner.grid_columnconfigure(2, weight=0)

        rating_colors = [
            ("good", "#c8e6c9", "#2e7d32"),
            ("normal", "#fff9c4", "#f9a825"),
            ("caution", "#ffcdd2", "#c62828"),
        ]
        label_width = 8
        row_pady = 2
        # 左の信号行の高さに合わせる（Canvas 20 + pady 2*2 = 24）
        memo_row_height = 24

        # 左ブロック: 部門・評価（8行）。行高を固定して右のメモ行と揃える。幅を明示して信号が描画されるようにする。
        left_block = ttk.Frame(inner)
        left_block.grid(row=0, column=0, sticky=tk.NW, padx=(0, 12))
        ttk.Label(left_block, text="部門・評価", font=("", 9), width=label_width).pack(anchor=tk.W, pady=(0, 4))
        # 1行の幅（ラベル＋信号3つ＋余白）。pack_propagate(False) 時に幅が0にならないよう指定
        signal_row_width = 220
        for key, label in VISIT_RATING_KEYS:
            self.visit_ratings[key] = None
            row = tk.Frame(left_block, height=memo_row_height, width=signal_row_width)
            row.pack(anchor=tk.W, pady=row_pady)
            row.pack_propagate(False)
            ttk.Label(row, text=f"{label}:", width=label_width, anchor=tk.E).pack(side=tk.LEFT, padx=(0, 4))
            btns = {}
            for val, light_color, dark_color in rating_colors:
                canv = tk.Canvas(row, width=22, height=20, bg="#f5f5f5", highlightthickness=0)
                canv.pack(side=tk.LEFT, padx=1)
                oval = canv.create_oval(2, 2, 20, 18, fill=light_color, outline="#9e9e9e", width=1)
                canv.configure(cursor="hand2")
                canv.bind("<Button-1>", lambda e, k=key, v=val: self._on_visit_rating_click(k, v))
                self._visit_rating_colors = getattr(self, "_visit_rating_colors", {})
                self._visit_rating_colors[(key, val)] = (light_color, dark_color)
                btns[val] = (canv, oval)
            self.visit_rating_buttons[key] = btns

        # 右ブロック: 気づいた点（箇条書き）8行（部門対応なし・一般メモ）。行高を左と揃える。
        notes_block = ttk.Frame(inner)
        notes_block.grid(row=0, column=1, sticky=tk.NSEW)
        ttk.Label(notes_block, text="気づいた点（箇条書き）", font=("", 9)).pack(anchor=tk.W, pady=(0, 4))
        self.visit_note_entries.clear()
        for _ in range(VISIT_NOTES_LINES):
            row_note = tk.Frame(notes_block, height=memo_row_height)
            row_note.pack(fill=tk.X, pady=row_pady)
            row_note.pack_propagate(False)
            entry = tk.Entry(row_note, font=("", 12), width=50)
            entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.visit_note_entries.append(entry)

        btn_frame = ttk.Frame(inner)
        btn_frame.grid(row=0, column=2, sticky=tk.NW, padx=(16, 0))
        ttk.Frame(btn_frame, height=1).pack(anchor=tk.W)
        btn_row = ttk.Frame(btn_frame)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="保存", command=self._on_save, width=10).pack(fill=tk.X, pady=2)
        ttk.Button(btn_row, text="記録を保存", command=self._on_save_record, width=10).pack(fill=tk.X, pady=2)
        ttk.Button(btn_row, text="閉じる", command=self._on_close, width=10).pack(fill=tk.X, pady=2)

        self._load_visit_notes()

    def _on_visit_rating_click(self, category_key: str, value: str):
        """部門の評価ボタンをクリック（1回で選択、同じボタンをもう1回でキャンセル）"""
        current = self.visit_ratings.get(category_key)
        if current == value:
            self.visit_ratings[category_key] = None  # 二回目クリックでキャンセル（コメントなし）
        else:
            self.visit_ratings[category_key] = value
        self._update_visit_rating_button_style(category_key)

    def _update_visit_rating_button_style(self, category_key: str):
        """選択中の評価ボタンをハイライト（薄い→濃い色に変更）"""
        btns = self.visit_rating_buttons.get(category_key, {})
        current = self.visit_ratings.get(category_key)
        colors = getattr(self, "_visit_rating_colors", {})
        for val, (canv, oval) in btns.items():
            light_color, dark_color = colors.get((category_key, val), ("#e0e0e0", "#616161"))
            fill_color = dark_color if val == current else light_color
            outline_color = dark_color if val == current else "#9e9e9e"
            canv.itemconfig(oval, fill=fill_color, outline=outline_color, width=2 if val == current else 1)

    def _load_visit_notes(self):
        """reproduction_checkup_visit_notes.json から当日のメモを読み込み"""
        path = self.farm_path / VISIT_NOTES_FILENAME
        if not path.exists():
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            day_data = data.get(self.checkup_date)
            if not isinstance(day_data, dict):
                return
            ratings = day_data.get("ratings", {})
            for key in self.visit_ratings:
                v = ratings.get(key)
                if v in ("good", "normal", "caution"):
                    self.visit_ratings[key] = v
                    self._update_visit_rating_button_style(key)
            notes = day_data.get("notes", [])
            if isinstance(notes, list):
                for i, entry in enumerate(self.visit_note_entries):
                    entry.delete(0, tk.END)
                    if i < len(notes) and notes[i]:
                        entry.insert(0, str(notes[i]))
        except Exception as e:
            logging.warning(f"検診訪問メモ読み込みエラー: {e}")

    def _on_save_record(self):
        """記録保存ボタン：日付・部門の色・箇条書きを保存（2回目以降は上書き）"""
        self._save_visit_notes(force=True)
        messagebox.showinfo("記録保存", "記録を保存しました。")

    def _save_visit_notes(self, force: bool = False):
        """検診訪問メモを reproduction_checkup_visit_notes.json に保存（force=True のときは空でも上書き）"""
        notes = []
        for entry in self.visit_note_entries:
            notes.append(entry.get().strip() if entry else "")
        ratings = {k: v for k, v in self.visit_ratings.items() if v}
        if not force and not ratings and not any(notes):
            return
        path = self.farm_path / VISIT_NOTES_FILENAME
        try:
            data = {}
            if path.exists():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            data[self.checkup_date] = {
                "notes": notes[:VISIT_NOTES_LINES],
                "ratings": ratings,
                "updated_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            }
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"検診訪問メモ保存エラー: {e}")
            messagebox.showerror("エラー", f"検診訪問メモの保存に失敗しました:\n{e}")

    def _on_close(self):
        """閉じるボタン：検診訪問メモを保存してからウィンドウを閉じる"""
        self._save_visit_notes()
        try:
            if hasattr(self, "main_canvas") and self.main_canvas.winfo_exists():
                self.main_canvas.unbind_all("<MouseWheel>")
                self.main_canvas.unbind_all("<Button-4>")
                self.main_canvas.unbind_all("<Button-5>")
        except tk.TclError:
            pass
        self.window.destroy()

    def _extract_and_display(self):
        settings = self.settings_manager.get("repro_checkup_settings", {})
        if not settings:
            messagebox.showwarning("警告", "繁殖検診設定が未設定です。設定を確認してください。")
            return

        try:
            self.results = self.checkup_logic.extract_cows(self.checkup_date, settings)
        except Exception as e:
            messagebox.showerror("エラー", f"抽出中にエラーが発生しました: {e}")
            return

        self.display_rows = []
        self.original_order = []

        for cow in self.results:
            cow_auto_id = cow.get("auto_id")
            if not cow_auto_id:
                continue

            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            # DIMは検診日時点で計算（本日基準にしない）
            dim_val = self.formula_engine.calculate_dim_at_date(cow_auto_id, self.checkup_date)

            dim_str = ""
            if dim_val is not None:
                dim_str = str(dim_val)

            dai_str = ""
            dai_val = None
            checkup_code = cow.get("checkup_code", "")
            # 原則フレッシュは授精後日数は空欄（分娩を挟んでいる場合も表示しない）
            if checkup_code != ReproductionCheckupLogic.CHECKUP_FRESH:
                try:
                    dai_val = self.formula_engine._calculate_dai_at_event_date(events, self.checkup_date)
                    if dai_val is not None:
                        dai_str = str(dai_val)
                except Exception as e:
                    logging.warning(f"DAI計算エラー (cow_auto_id={cow_auto_id}): {e}")

            # 検診内容（前回検診イベント300～304、307のNOTE）を取得
            last_checkup_result = self.formula_engine._get_last_reproduction_check_note(events) or ""

            row_id = str(cow_auto_id)
            row = {
                "row_id": row_id,
                "cow_id": cow.get("cow_id", ""),
                "jpn10": cow.get("jpn10", ""),
                "dim": dim_val,
                "dim_str": dim_str,
                "dai": dai_val,
                "dai_str": dai_str,
                "checkup_code": checkup_code,
                "last_checkup_result": last_checkup_result,
            }
            self.display_rows.append(row)
            self.original_order.append(row_id)

            if row_id not in self.row_data:
                self.row_data[row_id] = {
                    "cow_auto_id": cow_auto_id,
                    "cow_id": row["cow_id"],
                    "checkup_code": row["checkup_code"],
                    "inputs": {
                        "exec_code": "",
                        "treatment": "",
                        "uterine_findings": "",
                        "right_ovary_findings": "",
                        "left_ovary_findings": "",
                        "other": "",
                        "twin": False,
                        "female_judgment": False,
                        "male_judgment": False,
                        "bcs": "",
                    },
                }

        self._render_rows()

    def _render_rows(self):
        # 左テーブルを更新
        for item in self.left_tree.get_children():
            self.left_tree.delete(item)

        # 入力欄を更新
        for child in self.input_inner.winfo_children():
            child.destroy()

        self.input_widgets.clear()
        self.input_checkboxes.clear()
        self.input_entries_order.clear()

        for row in self.display_rows:
            row_id = row["row_id"]
            values = [
                row["cow_id"],
                row["jpn10"],
                row["dim_str"],
                row["dai_str"],
                row.get("last_checkup_result", ""),
                row["checkup_code"],
                row["cow_id"],
            ]
            self.left_tree.insert("", "end", iid=row_id, values=tuple(values))
            self._build_input_row(row_id)
        
        # 実施コードエントリーの変更を監視して双子・♀♂判定チェックボックスの状態を更新
        if hasattr(self, '_pending_twin_state_updates'):
            for row_id, update_func in self._pending_twin_state_updates.items():
                exec_code_entry = self.input_widgets.get(row_id, {}).get("exec_code")
                if exec_code_entry:
                    def make_update_handler(rid, update_fn):
                        def handler(event=None):
                            update_fn()
                        return handler
                    update_handler = make_update_handler(row_id, update_func)
                    exec_code_entry.bind("<KeyRelease>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_key_release(e, rid), h()))
                    exec_code_entry.bind("<FocusOut>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_focus_out(e, rid), h()))
            self._pending_twin_state_updates.clear()
        if hasattr(self, '_pending_female_judgment_state_updates'):
            for row_id, update_func in self._pending_female_judgment_state_updates.items():
                exec_code_entry = self.input_widgets.get(row_id, {}).get("exec_code")
                if exec_code_entry:
                    def make_update_handler_fj(rid, update_fn):
                        def handler(event=None):
                            update_fn()
                        return handler
                    update_handler = make_update_handler_fj(row_id, update_func)
                    exec_code_entry.bind("<KeyRelease>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_key_release(e, rid), h()))
                    exec_code_entry.bind("<FocusOut>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_focus_out(e, rid), h()))
            self._pending_female_judgment_state_updates.clear()
        if hasattr(self, '_pending_male_judgment_state_updates'):
            for row_id, update_func in self._pending_male_judgment_state_updates.items():
                exec_code_entry = self.input_widgets.get(row_id, {}).get("exec_code")
                if exec_code_entry:
                    def make_update_handler_mj(rid, update_fn):
                        def handler(event=None):
                            update_fn()
                        return handler
                    update_handler = make_update_handler_mj(row_id, update_func)
                    exec_code_entry.bind("<KeyRelease>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_key_release(e, rid), h()))
                    exec_code_entry.bind("<FocusOut>", lambda e, rid=row_id, h=update_handler: (self._on_exec_code_focus_out(e, rid), h()))
            self._pending_male_judgment_state_updates.clear()

        self._sync_input_canvas_size()
        
        # 現在ハイライトされている行があれば再ハイライト
        if self.current_highlighted_row_id:
            self._highlight_row(self.current_highlighted_row_id)
        
        self._focus_first_input()

    def _on_input_canvas_configure(self, event):
        self.input_canvas.itemconfig(self.input_canvas_window, width=self.right_total_width)

    def _sync_input_canvas_size(self):
        self.input_canvas.update_idletasks()
        self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all"))


    def _build_input_row(self, row_id: str):
        row_frame = ttk.Frame(self.input_inner, height=24, width=self.right_total_width)
        row_frame.pack(fill=tk.X)
        row_frame.pack_propagate(False)

        self.input_widgets[row_id] = {}

        x_offset = 0
        for col_key, _label, width_px in self.right_column_specs:
            if col_key == "twin":
                # 双子チェックボックス
                twin_var = tk.BooleanVar()
                checkbox_key = f"{row_id}_twin"
                self.input_checkboxes[checkbox_key] = twin_var
                
                # 現在の値を設定
                current_twin = self.row_data.get(row_id, {}).get("inputs", {}).get("twin", False)
                if isinstance(current_twin, bool):
                    twin_var.set(current_twin)
                elif isinstance(current_twin, str):
                    twin_var.set(current_twin.lower() in ("true", "1", "yes"))
                
                checkbox = tk.Checkbutton(
                    row_frame,
                    variable=twin_var,
                    bg="#fff3c4",
                    relief="solid",
                    bd=1,
                    command=lambda rid=row_id, var=twin_var: self._update_twin_value(rid, var.get())
                )
                checkbox.place(x=x_offset, y=0, width=width_px, height=24)
                x_offset += width_px
                
                # 実施コードに応じて有効/無効を切り替える関数
                def update_twin_state(rid=row_id, ck=checkbox):
                    exec_code = self.row_data.get(rid, {}).get("inputs", {}).get("exec_code", "").strip()
                    # 実施コードが5または6の場合のみ有効
                    enabled = exec_code in ("5", "6")
                    ck.config(state=tk.NORMAL if enabled else tk.DISABLED)
                    if not enabled:
                        twin_var.set(False)
                        self._update_twin_value(rid, False)
                
                # 実施コード変更時に双子チェックボックスの状態を更新
                # 注意: exec_code_entryはまだ作成されていない可能性があるため、
                # _build_input_rowの後で設定する必要がある
                # ここでは、後で設定するためのコールバックを保存
                self._pending_twin_state_updates = getattr(self, '_pending_twin_state_updates', {})
                self._pending_twin_state_updates[row_id] = update_twin_state
                
                # 初期状態を設定
                update_twin_state()
                
                # エンターで次の行の実施コードへ移動（エンターのみで次々個体を入力可能に）
                def on_twin_return(event, rid=row_id):
                    self._focus_next_entry(rid, "twin")
                    return "break"
                checkbox.bind("<Return>", on_twin_return)
                
                self.input_widgets[row_id][col_key] = checkbox
            elif col_key == "female_judgment":
                # ♀判定チェックボックス（妊娠プラス時のみ有効、双子と同様）
                female_judgment_var = tk.BooleanVar()
                checkbox_key = f"{row_id}_female_judgment"
                self.input_checkboxes[checkbox_key] = female_judgment_var
                current_fj = self.row_data.get(row_id, {}).get("inputs", {}).get("female_judgment", False)
                if isinstance(current_fj, bool):
                    female_judgment_var.set(current_fj)
                elif isinstance(current_fj, str):
                    female_judgment_var.set(current_fj.lower() in ("true", "1", "yes"))
                checkbox = tk.Checkbutton(
                    row_frame,
                    variable=female_judgment_var,
                    bg="#fff3c4",
                    relief="solid",
                    bd=1,
                    command=lambda rid=row_id, var=female_judgment_var: self._update_female_judgment_value(rid, var.get())
                )
                checkbox.place(x=x_offset, y=0, width=width_px, height=24)
                x_offset += width_px
                def update_female_judgment_state(rid=row_id, ck=checkbox):
                    exec_code = self.row_data.get(rid, {}).get("inputs", {}).get("exec_code", "").strip()
                    enabled = exec_code in ("5", "6")
                    ck.config(state=tk.NORMAL if enabled else tk.DISABLED)
                    if not enabled:
                        female_judgment_var.set(False)
                        self._update_female_judgment_value(rid, False)
                self._pending_female_judgment_state_updates = getattr(self, '_pending_female_judgment_state_updates', {})
                self._pending_female_judgment_state_updates[row_id] = update_female_judgment_state
                update_female_judgment_state()
                def on_female_judgment_return(event, rid=row_id):
                    self._focus_next_entry(rid, "female_judgment")
                    return "break"
                checkbox.bind("<Return>", on_female_judgment_return)
                self.input_widgets[row_id][col_key] = checkbox
            elif col_key == "male_judgment":
                # ♂判定チェックボックス（妊娠プラス時のみ有効、♀と同様）
                male_judgment_var = tk.BooleanVar()
                checkbox_key = f"{row_id}_male_judgment"
                self.input_checkboxes[checkbox_key] = male_judgment_var
                current_mj = self.row_data.get(row_id, {}).get("inputs", {}).get("male_judgment", False)
                if isinstance(current_mj, bool):
                    male_judgment_var.set(current_mj)
                elif isinstance(current_mj, str):
                    male_judgment_var.set(current_mj.lower() in ("true", "1", "yes"))
                checkbox = tk.Checkbutton(
                    row_frame,
                    variable=male_judgment_var,
                    bg="#fff3c4",
                    relief="solid",
                    bd=1,
                    command=lambda rid=row_id, var=male_judgment_var: self._update_male_judgment_value(rid, var.get())
                )
                checkbox.place(x=x_offset, y=0, width=width_px, height=24)
                x_offset += width_px
                def update_male_judgment_state(rid=row_id, ck=checkbox):
                    exec_code = self.row_data.get(rid, {}).get("inputs", {}).get("exec_code", "").strip()
                    enabled = exec_code in ("5", "6")
                    ck.config(state=tk.NORMAL if enabled else tk.DISABLED)
                    if not enabled:
                        male_judgment_var.set(False)
                        self._update_male_judgment_value(rid, False)
                self._pending_male_judgment_state_updates = getattr(self, '_pending_male_judgment_state_updates', {})
                self._pending_male_judgment_state_updates[row_id] = update_male_judgment_state
                update_male_judgment_state()
                def on_male_judgment_return(event, rid=row_id):
                    self._focus_next_entry(rid, "male_judgment")
                    return "break"
                checkbox.bind("<Return>", on_male_judgment_return)
                self.input_widgets[row_id][col_key] = checkbox
            else:
                # 通常のエントリー
                entry = tk.Entry(
                    row_frame,
                    bg="#fff3c4",
                    relief="solid",
                    bd=1,
                )
                entry.place(x=x_offset, y=0, width=width_px, height=24)
                x_offset += width_px

                current_value = self.row_data.get(row_id, {}).get("inputs", {}).get(col_key, "")
                if current_value:
                    entry.insert(0, current_value)
                    # 処置・子宮・右・左・その他は表示時も大文字に統一
                    if col_key in ("treatment", "uterine_findings", "right_ovary_findings", "left_ovary_findings", "other"):
                        self._convert_to_uppercase(entry)

                def on_focus_in(event, rid=row_id):
                    self._highlight_row(rid)
                
                def on_focus_out(event, rid=row_id, ck=col_key):
                    self._update_input_value(rid, ck, event.widget.get())
                    # 実施コードの場合は双子・♀判定チェックボックスの状態を更新
                    if ck == "exec_code":
                        self._update_twin_checkbox_state(rid)
                        self._update_female_judgment_checkbox_state(rid)
                        self._update_male_judgment_checkbox_state(rid)

                def on_key_release(event, rid=row_id, ck=col_key):
                    # 処置・子宮・右・左・その他はデフォルト大文字（イベント入力と同様）
                    if ck in ("treatment", "uterine_findings", "right_ovary_findings", "left_ovary_findings", "other"):
                        self._convert_to_uppercase(event.widget)
                    if ck == "treatment":
                        normalized = self._normalize_treatment_value(event.widget.get().strip())
                        if normalized != event.widget.get():
                            event.widget.delete(0, tk.END)
                            event.widget.insert(0, normalized)
                    self._update_input_value(rid, ck, event.widget.get())
                    # 実施コードの場合は双子・♀判定チェックボックスの状態を更新
                    if ck == "exec_code":
                        self._update_twin_checkbox_state(rid)
                        self._update_female_judgment_checkbox_state(rid)

                def on_enter(event, rid=row_id, ck=col_key):
                    self._update_input_value(rid, ck, event.widget.get())
                    # 実施コードの場合は双子・♀♂判定チェックボックスの状態を更新
                    if ck == "exec_code":
                        self._update_twin_checkbox_state(rid)
                        self._update_female_judgment_checkbox_state(rid)
                        self._update_male_judgment_checkbox_state(rid)
                    self._focus_next_entry(rid, ck)
                    return "break"
                
                def on_arrow_key(event, rid=row_id, ck=col_key):
                    """矢印キーでセル移動"""
                    direction = event.keysym
                    if direction == "Right":
                        # 右矢印: 次の列（同じ行）
                        self._focus_next_entry(rid, ck)
                    elif direction == "Left":
                        # 左矢印: 前の列（同じ行）
                        self._focus_previous_entry(rid, ck)
                    elif direction == "Down":
                        # 下矢印: 次の行（同じ列）
                        self._focus_entry_below(rid, ck)
                    elif direction == "Up":
                        # 上矢印: 前の行（同じ列）
                        self._focus_entry_above(rid, ck)
                    return "break"

                entry.bind("<FocusIn>", on_focus_in)
                entry.bind("<FocusOut>", on_focus_out)
                entry.bind("<KeyRelease>", on_key_release)
                # 処置・子宮・右・左・その他は入力時リアルタイムで大文字に変換
                if col_key in ("treatment", "uterine_findings", "right_ovary_findings", "left_ovary_findings", "other"):
                    entry.bind("<KeyPress>", lambda e, w=entry: self._convert_to_uppercase_on_input(e, w))
                entry.bind("<Return>", on_enter)
                entry.bind("<Right>", on_arrow_key)
                entry.bind("<Left>", on_arrow_key)
                entry.bind("<Down>", on_arrow_key)
                entry.bind("<Up>", on_arrow_key)
                # マウスホイールでスクロール（エントリーにフォーカスがある場合でも）
                entry.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, self.yview_func))
                entry.bind("<Button-4>", lambda e: self._on_mousewheel_linux(e, self.yview_func, -1))
                entry.bind("<Button-5>", lambda e: self._on_mousewheel_linux(e, self.yview_func, 1))

                self.input_widgets[row_id][col_key] = entry
                self.input_entries_order.append(entry)

    def _update_input_value(self, row_id: str, column_name: str, value: str):
        if row_id in self.row_data:
            self.row_data[row_id]["inputs"][column_name] = value.strip()
    
    def _update_twin_value(self, row_id: str, value: bool):
        """双子チェックボックスの値を更新"""
        if row_id in self.row_data:
            self.row_data[row_id]["inputs"]["twin"] = value
    
    def _update_twin_checkbox_state(self, row_id: str):
        """実施コードに応じて双子チェックボックスの有効/無効を切り替え"""
        exec_code = self.row_data.get(row_id, {}).get("inputs", {}).get("exec_code", "").strip()
        enabled = exec_code in ("5", "6")
        
        checkbox_key = f"{row_id}_twin"
        if checkbox_key in self.input_checkboxes:
            checkbox = self.input_widgets.get(row_id, {}).get("twin")
            if checkbox:
                checkbox.config(state=tk.NORMAL if enabled else tk.DISABLED)
                if not enabled:
                    self.input_checkboxes[checkbox_key].set(False)
                    self._update_twin_value(row_id, False)

    def _update_female_judgment_value(self, row_id: str, value: bool):
        """♀判定チェックボックスの値を更新"""
        if row_id in self.row_data:
            self.row_data[row_id]["inputs"]["female_judgment"] = value

    def _update_female_judgment_checkbox_state(self, row_id: str):
        """実施コードに応じて♀判定チェックボックスの有効/無効を切り替え"""
        exec_code = self.row_data.get(row_id, {}).get("inputs", {}).get("exec_code", "").strip()
        enabled = exec_code in ("5", "6")
        checkbox_key = f"{row_id}_female_judgment"
        if checkbox_key in self.input_checkboxes:
            checkbox = self.input_widgets.get(row_id, {}).get("female_judgment")
            if checkbox:
                checkbox.config(state=tk.NORMAL if enabled else tk.DISABLED)
                if not enabled:
                    self.input_checkboxes[checkbox_key].set(False)
                    self._update_female_judgment_value(row_id, False)

    def _update_male_judgment_value(self, row_id: str, value: bool):
        """♂判定チェックボックスの値を更新"""
        if row_id in self.row_data:
            self.row_data[row_id]["inputs"]["male_judgment"] = value

    def _update_male_judgment_checkbox_state(self, row_id: str):
        """実施コードに応じて♂判定チェックボックスの有効/無効を切り替え"""
        exec_code = self.row_data.get(row_id, {}).get("inputs", {}).get("exec_code", "").strip()
        enabled = exec_code in ("5", "6")
        checkbox_key = f"{row_id}_male_judgment"
        if checkbox_key in self.input_checkboxes:
            checkbox = self.input_widgets.get(row_id, {}).get("male_judgment")
            if checkbox:
                checkbox.config(state=tk.NORMAL if enabled else tk.DISABLED)
                if not enabled:
                    self.input_checkboxes[checkbox_key].set(False)
                    self._update_male_judgment_value(row_id, False)
    
    def _on_exec_code_key_release(self, event, row_id: str):
        """実施コードのキーリリースイベント（既存の処理を保持）"""
        pass
    
    def _on_exec_code_focus_out(self, event, row_id: str):
        """実施コードのフォーカスアウトイベント（既存の処理を保持）"""
        pass

    def _focus_next_entry(self, row_id: str, column_name: str):
        """次のセル（右）にフォーカスを移動。行末の「その他」では次の行の実施コードへ（エンターのみで次々入力可能）"""
        row_ids = [row["row_id"] for row in self.display_rows]
        if row_id not in row_ids:
            return
        row_index = row_ids.index(row_id)
        cols = [c[0] for c in self.right_column_specs]
        try:
            col_index = cols.index(column_name)
        except ValueError:
            return

        next_col_index = col_index + 1
        next_row_index = row_index
        # 行末の「その他」で Enter → 次の行の実施コードへ（双子はスキップして次個体へ）
        if column_name == "other":
            next_col_index = 0
            next_row_index = row_index + 1
        elif next_col_index >= len(cols):
            next_col_index = 0
            next_row_index += 1

        if next_row_index >= len(row_ids):
            return

        next_row_id = row_ids[next_row_index]
        next_col = cols[next_col_index]
        widget = self.input_widgets.get(next_row_id, {}).get(next_col)
        if widget:
            if next_row_index != row_index:
                self._scroll_to_row(next_row_id)
            widget.focus_set()
            if hasattr(widget, "select_range"):
                widget.select_range(0, tk.END)
    
    def _focus_previous_entry(self, row_id: str, column_name: str):
        """前のセル（左）にフォーカスを移動"""
        row_ids = [row["row_id"] for row in self.display_rows]
        if row_id not in row_ids:
            return
        row_index = row_ids.index(row_id)
        cols = [c[0] for c in self.right_column_specs]
        try:
            col_index = cols.index(column_name)
        except ValueError:
            return

        prev_col_index = col_index - 1
        prev_row_index = row_index
        if prev_col_index < 0:
            prev_col_index = len(cols) - 1
            prev_row_index -= 1

        if prev_row_index < 0:
            return

        prev_row_id = row_ids[prev_row_index]
        prev_col = cols[prev_col_index]
        entry = self.input_widgets.get(prev_row_id, {}).get(prev_col)
        if entry:
            if prev_row_index != row_index:
                self._scroll_to_row(prev_row_id)
            entry.focus_set()
            if hasattr(entry, "select_range"):
                entry.select_range(0, tk.END)
    
    def _focus_entry_below(self, row_id: str, column_name: str):
        """下のセル（同じ列）にフォーカスを移動"""
        row_ids = [row["row_id"] for row in self.display_rows]
        if row_id not in row_ids:
            return
        row_index = row_ids.index(row_id)
        cols = [c[0] for c in self.right_column_specs]
        try:
            col_index = cols.index(column_name)
        except ValueError:
            return

        next_row_index = row_index + 1
        if next_row_index >= len(row_ids):
            return

        next_row_id = row_ids[next_row_index]
        entry = self.input_widgets.get(next_row_id, {}).get(column_name)
        if entry:
            self._scroll_to_row(next_row_id)
            entry.focus_set()
            if hasattr(entry, "select_range"):
                entry.select_range(0, tk.END)
    
    def _focus_entry_above(self, row_id: str, column_name: str):
        """上のセル（同じ列）にフォーカスを移動"""
        row_ids = [row["row_id"] for row in self.display_rows]
        if row_id not in row_ids:
            return
        row_index = row_ids.index(row_id)
        cols = [c[0] for c in self.right_column_specs]
        try:
            col_index = cols.index(column_name)
        except ValueError:
            return

        prev_row_index = row_index - 1
        if prev_row_index < 0:
            return

        prev_row_id = row_ids[prev_row_index]
        entry = self.input_widgets.get(prev_row_id, {}).get(column_name)
        if entry:
            self._scroll_to_row(prev_row_id)
            entry.focus_set()
            if hasattr(entry, "select_range"):
                entry.select_range(0, tk.END)

    def _focus_first_input(self):
        if not self.display_rows:
            return
        first_row_id = self.display_rows[0]["row_id"]
        first_col = self.right_column_specs[0][0]
        entry = self.input_widgets.get(first_row_id, {}).get(first_col)
        if entry:
            entry.focus_set()

    def _on_search_id(self, event=None):
        """ID検索機能"""
        search_text = self.search_entry.get().strip()
        if not search_text:
            return
        
        # IDで検索
        found_row_id = None
        for row in self.display_rows:
            if row.get("cow_id") == search_text:
                found_row_id = row["row_id"]
                break
        
        if found_row_id:
            # 該当行の実施コードにフォーカスを移動
            exec_code_entry = self.input_widgets.get(found_row_id, {}).get("exec_code")
            if exec_code_entry:
                exec_code_entry.focus_set()
                exec_code_entry.select_range(0, tk.END)
                # 該当行までスクロール
                self._scroll_to_row(found_row_id)
                # 行をハイライト（フォーカスイベントで自動的にハイライトされるが、念のため）
                self._highlight_row(found_row_id)
        else:
            # 見つからない場合はメッセージを表示して検索欄に戻る
            messagebox.showinfo("検索結果", f"ID '{search_text}' が見つかりませんでした。")
            self.search_entry.focus_set()
            self.search_entry.select_range(0, tk.END)

    def _on_search_key_release(self, event):
        """検索欄のキーリリースイベント（ESCキーでクリア）"""
        if event.keysym == "Escape":
            self.search_entry.delete(0, tk.END)
            self.search_entry.focus_set()

    def _scroll_to_row(self, row_id: str):
        """指定された行までスクロール"""
        row_ids = [row["row_id"] for row in self.display_rows]
        if row_id not in row_ids:
            return
        
        row_index = row_ids.index(row_id)
        # 左テーブルのスクロール
        total_rows = len(self.display_rows)
        if total_rows > 0:
            fraction = row_index / total_rows
            self.left_tree.yview_moveto(fraction)
        # 右入力欄のスクロール
        self.input_canvas.yview_moveto(fraction)

    def _highlight_row(self, row_id: str):
        """指定された行全体をハイライト"""
        # 前の行のハイライトを解除
        if self.current_highlighted_row_id and self.current_highlighted_row_id != row_id:
            self._unhighlight_row(self.current_highlighted_row_id)
        
        # 新しい行をハイライト
        self.current_highlighted_row_id = row_id
        
        # 左テーブルの行をハイライト
        try:
            self.left_tree.set(row_id, "cow_id")  # 行が存在するか確認
            self.left_tree.item(row_id, tags=("highlighted",))
        except tk.TclError:
            pass  # 行が存在しない場合は無視
        
        # 右入力欄の行をハイライト
        row_widgets = self.input_widgets.get(row_id, {})
        highlight_color = "#b3d9ff"  # 薄い青
        for entry in row_widgets.values():
            entry.configure(bg=highlight_color)

    def _unhighlight_row(self, row_id: str):
        """指定された行のハイライトを解除"""
        # 左テーブルの行のハイライトを解除
        try:
            self.left_tree.item(row_id, tags=())
        except tk.TclError:
            pass  # 行が存在しない場合は無視
        
        # 右入力欄の行のハイライトを解除
        row_widgets = self.input_widgets.get(row_id, {})
        normal_color = "#fff3c4"  # 通常の背景色
        for entry in row_widgets.values():
            entry.configure(bg=normal_color)

    def _normalize_treatment_value(self, value: str) -> str:
        if value.isdigit():
            treatment = self.treatments.get(value)
            if treatment:
                name = treatment.get("name", "")
                return name or value
        return value

    def _on_mousewheel(self, event, yview_func):
        """Windows用のマウスホイール処理"""
        delta = -1 if event.delta > 0 else 1
        yview_func("scroll", delta, "units")
        return "break"
    
    def _on_mousewheel_linux(self, event, yview_func, direction):
        """Linux用のマウスホイール処理"""
        yview_func("scroll", direction, "units")
        return "break"

    def _sync_inputs_from_widgets(self):
        """ウィジェットから入力値を同期"""
        for row_id, widgets in self.input_widgets.items():
            for col_key, widget in widgets.items():
                if col_key == "twin":
                    checkbox_key = f"{row_id}_twin"
                    if checkbox_key in self.input_checkboxes:
                        twin_value = self.input_checkboxes[checkbox_key].get()
                        self._update_twin_value(row_id, twin_value)
                elif col_key == "female_judgment":
                    checkbox_key = f"{row_id}_female_judgment"
                    if checkbox_key in self.input_checkboxes:
                        fj_value = self.input_checkboxes[checkbox_key].get()
                        self._update_female_judgment_value(row_id, fj_value)
                elif col_key == "male_judgment":
                    checkbox_key = f"{row_id}_male_judgment"
                    if checkbox_key in self.input_checkboxes:
                        mj_value = self.input_checkboxes[checkbox_key].get()
                        self._update_male_judgment_value(row_id, mj_value)
                else:
                    # 通常のエントリー
                    if isinstance(widget, tk.Entry):
                        self._update_input_value(row_id, col_key, widget.get())

    def _on_column_click(self, column: str):
        if not self.display_rows:
            return

        self._sync_inputs_from_widgets()

        current_state = self.sort_state.get(column)
        if current_state is None:
            new_state = "asc"
        elif current_state == "asc":
            new_state = "desc"
        else:
            new_state = None
        self.sort_state[column] = new_state

        if new_state is None:
            order_map = {row_id: i for i, row_id in enumerate(self.original_order)}
            self.display_rows.sort(key=lambda r: order_map.get(r["row_id"], 999999))
        else:
            reverse = new_state == "desc"

            def sort_key(row):
                if column == "cow_id":
                    try:
                        return int(row.get("cow_id") or 0)
                    except (ValueError, TypeError):
                        return 0
                if column == "cow_id_dup":
                    try:
                        return int(row.get("cow_id") or 0)
                    except (ValueError, TypeError):
                        return 0
                if column == "jpn10":
                    return row.get("jpn10") or ""
                if column == "dim":
                    return row.get("dim") if row.get("dim") is not None else -1
                if column == "dai":
                    return row.get("dai") if row.get("dai") is not None else -1
                if column == "checkup_code":
                    return row.get("checkup_code") or ""
                if column == "last_checkup_result":
                    return row.get("last_checkup_result") or ""
                return ""

            self.display_rows.sort(key=sort_key, reverse=reverse)

        self._render_rows()

    def _build_json_data(self, inputs: Dict[str, Any]) -> Dict[str, Any]:
        json_data: Dict[str, Any] = {}
        for key in ["treatment", "uterine_findings", "right_ovary_findings", "left_ovary_findings", "other"]:
            val = inputs.get(key, "")
            if isinstance(val, str):
                val = val.strip()
            if val:
                json_data[key] = val
        # 双子チェックボックスの値を追加（実施コード5または6の場合）
        twin_value = inputs.get("twin", False)
        if isinstance(twin_value, bool) and twin_value:
            json_data["twin"] = True
        # ♀♂判定チェックボックスの値を追加（実施コード5または6の場合、双子と同様の形で格納）
        female_judgment_value = inputs.get("female_judgment", False)
        if isinstance(female_judgment_value, bool) and female_judgment_value:
            json_data["female_judgment"] = True
        male_judgment_value = inputs.get("male_judgment", False)
        if isinstance(male_judgment_value, bool) and male_judgment_value:
            json_data["male_judgment"] = True
        return json_data

    def _on_save(self):
        self._sync_inputs_from_widgets()
        self._save_visit_notes()
        saved_count = 0
        removed_row_ids = set()
        EVENT_BCS = 101  # イベント辞書のBCS
        for row_id, data in self.row_data.items():
            inputs = data.get("inputs", {})
            cow_auto_id = data.get("cow_auto_id")

            # BCS欄に入力がある場合はイベントBCS（101）として日付・NOTEに保存
            if cow_auto_id:
                bcs_val = inputs.get("bcs", "").strip()
                if bcs_val:
                    try:
                        event_data_bcs = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": EVENT_BCS,
                            "event_date": self.checkup_date,
                            "json_data": None,
                            "note": bcs_val,
                        }
                        event_id = self.db.insert_event(event_data_bcs)
                        self.rule_engine.on_event_added(event_id)
                        record_from_event(
                            self.db,
                            cow_auto_id=cow_auto_id,
                            action="登録",
                            event_number=EVENT_BCS,
                            event_id=event_id,
                            event_date=self.checkup_date,
                            event_dictionary=None,
                        )
                        saved_count += 1
                    except Exception as e:
                        messagebox.showerror("エラー", f"BCSイベントの保存に失敗しました: {e}")
                        return

            exec_code = inputs.get("exec_code", "").strip()
            if not exec_code:
                continue

            if exec_code == "1":
                removed_row_ids.add(row_id)
                continue

            if not cow_auto_id:
                continue

            if exec_code == "9":
                from ui.event_input import EventInputWindow
                event_input_window = EventInputWindow.open_or_focus(
                    parent=self.window,
                    db_handler=self.db,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    cow_auto_id=cow_auto_id,
                    on_saved=None,
                    farm_path=self.farm_path,
                )
                try:
                    event_input_window.date_entry.delete(0, tk.END)
                    event_input_window.date_entry.insert(0, self.checkup_date)
                except Exception:
                    pass
                event_input_window.show()
                continue

            # 実施コード6の場合、受胎日入力が必要
            if exec_code == "6":
                # イベント304（直近以外の妊娠プラス）の受胎日を入力させる
                from ui.event_input import EventInputWindow
                
                # 既存のイベント304を確認
                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                existing_pdp2_event = None
                for event in sorted(events, key=lambda x: (x.get('event_date', ''), x.get('id', 0)), reverse=True):
                    if event.get('event_number') == 304 and event.get('event_date') == self.checkup_date:
                        existing_pdp2_event = event
                        break
                
                # イベント入力ウィンドウを起動
                event_input_window = EventInputWindow.open_or_focus(
                    parent=self.window,
                    db_handler=self.db,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    cow_auto_id=cow_auto_id,
                    on_saved=None,
                    farm_path=self.farm_path,
                    allowed_event_numbers=[304],  # イベント304のみ許可
                    default_event_number=304,  # デフォルトで304を選択
                    edit_event_id=existing_pdp2_event.get('id') if existing_pdp2_event else None,  # 既存イベントがある場合は編集
                )
                try:
                    event_input_window.date_entry.delete(0, tk.END)
                    event_input_window.date_entry.insert(0, self.checkup_date)
                except Exception:
                    pass
                event_input_window.show()
                
                # 受胎日が入力されたか確認
                events_after = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                latest_pdp2_event = None
                for event in sorted(events_after, key=lambda x: (x.get('event_date', ''), x.get('id', 0)), reverse=True):
                    if event.get('event_number') == 304 and event.get('event_date') == self.checkup_date:
                        latest_pdp2_event = event
                        break
                
                if not latest_pdp2_event:
                    messagebox.showwarning("警告", f"個体ID {data.get('cow_id', '')} のイベント304が保存されていません。\n実施コード6を保存するには、イベント304を保存してください。")
                    continue
                
                # 受胎日が入力されているか確認（json_dataにconception_dateがあるか、またはAI/ETイベントが選択されているか）
                json_data_pdp2 = latest_pdp2_event.get('json_data') or {}
                if isinstance(json_data_pdp2, str):
                    try:
                        import json
                        json_data_pdp2 = json.loads(json_data_pdp2)
                    except:
                        json_data_pdp2 = {}
                
                has_conception_date = json_data_pdp2.get('conception_date') is not None
                has_conception_ai_event = json_data_pdp2.get('conception_ai_event_id') is not None
                
                if not has_conception_date and not has_conception_ai_event:
                    messagebox.showwarning("警告", f"個体ID {data.get('cow_id', '')} の受胎日が入力されていません。\n実施コード6を保存するには、受胎日を入力してください。")
                    continue

            event_number = self.EXEC_CODE_EVENT_MAP.get(exec_code)
            if not event_number:
                messagebox.showwarning("警告", f"実施コードが無効です: {exec_code}")
                continue

            json_data = {}
            if event_number in [300, 301, RuleEngine.EVENT_PDN]:
                json_data = self._build_json_data(inputs)
            elif event_number in [RuleEngine.EVENT_PDP, RuleEngine.EVENT_PDP2, RuleEngine.EVENT_PAGP]:
                # 妊娠プラスイベント（303, 304, 307）の場合、双子チェックを追加
                json_data = self._build_json_data(inputs)

            event_data = {
                "cow_auto_id": cow_auto_id,
                "event_number": event_number,
                "event_date": self.checkup_date,
                "json_data": json_data if json_data else None,
                "note": None,
            }

            # 同じ日付で既に繁殖検診イベントがある場合は警告
            if event_number in REPRO_CHECKUP_EVENT_NUMBERS:
                existing_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                date_str = (self.checkup_date or "")[:10]
                same_day_repro = [
                    e for e in existing_events
                    if ((e.get("event_date") or "")[:10] == date_str
                        and e.get("event_number") in REPRO_CHECKUP_EVENT_NUMBERS)
                ]
                if same_day_repro:
                    cow_id_display = data.get("cow_id", "") or str(cow_auto_id)
                    if not messagebox.askyesno(
                        "重複の可能性",
                        f"個体ID {cow_id_display} は同じ日付（{self.checkup_date}）に既に繁殖検診イベントが登録されています。"
                        "重複の可能性があります。\n\n登録しますか？",
                    ):
                        continue

            try:
                event_id = self.db.insert_event(event_data)
                self.rule_engine.on_event_added(event_id)
                record_from_event(
                    self.db,
                    cow_auto_id=cow_auto_id,
                    action="登録",
                    event_number=event_number,
                    event_id=event_id,
                    event_date=self.checkup_date,
                    event_dictionary=None,
                )
                saved_count += 1
                removed_row_ids.add(row_id)
            except Exception as e:
                messagebox.showerror("エラー", f"イベント保存に失敗しました: {e}")
                return

        if removed_row_ids:
            self.display_rows = [row for row in self.display_rows if row["row_id"] not in removed_row_ids]
            self.original_order = [row_id for row_id in self.original_order if row_id not in removed_row_ids]
            for row_id in removed_row_ids:
                self.row_data.pop(row_id, None)
            self._render_rows()

        messagebox.showinfo("完了", f"{saved_count}件のイベントを保存しました。")

    def show(self):
        self.window.wait_window()
