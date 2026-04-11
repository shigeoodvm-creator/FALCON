"""
FALCON2 - 新規農場作成ウィンドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional, Callable

from modules.farm_creator import FarmCreator
from modules.cow_registration_excel_template import prompt_save_bulk_cow_registration_template
from backup_manager import save_backup_settings
import constants  # constants.FARMS_ROOT を動的参照するため module ごと import

# ── カラーパレット ──────────────────────────────
_BG         = "#f5f6fa"
_CARD_OFF   = "#ffffff"
_CARD_ON    = "#edf0f3"
_BORDER_OFF = "#dde1e7"
_BORDER_ON  = "#1e2a3a"
_PRIMARY    = "#1e2a3a"
_PRIMARY_H  = "#2d3e50"
_TEXT       = "#1c2333"
_TEXT_SUB   = "#6b7280"
_ACCENT     = "#edf0f3"
_FONT       = "Meiryo UI"


# ── ユーティリティ ──────────────────────────────

def _bind_tree(widget, event, callback):
    """widget とすべての子孫に同じバインドを設定"""
    widget.bind(event, callback)
    for child in widget.winfo_children():
        _bind_tree(child, event, callback)


def _recolor(widget, bg):
    try:
        widget.config(bg=bg)
    except Exception:
        pass
    for child in widget.winfo_children():
        _recolor(child, bg)


# ── 選択カード ──────────────────────────────────

class _Card(tk.Frame):
    """クリックで選択されるカード型ウィジェット"""

    def __init__(self, parent, value: str, icon: str,
                 title: str, subtitle: str, on_select):
        super().__init__(
            parent,
            bg=_CARD_OFF, cursor="hand2",
            highlightthickness=2, highlightbackground=_BORDER_OFF,
            bd=0
        )
        self.value = value
        self._on_select = on_select
        self._selected = False

        inner = tk.Frame(self, bg=_CARD_OFF, padx=16, pady=12)
        inner.pack(fill=tk.X)

        # アイコン ＋ タイトル
        head = tk.Frame(inner, bg=_CARD_OFF)
        head.pack(fill=tk.X)
        tk.Label(head, text=icon, font=(_FONT, 15), bg=_CARD_OFF,
                 fg=_PRIMARY, width=2, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(head, text=title, font=(_FONT, 10, "bold"), bg=_CARD_OFF,
                 fg=_TEXT, anchor=tk.W).pack(side=tk.LEFT, padx=(6, 0))

        # サブタイトル
        tk.Label(inner, text=subtitle, font=(_FONT, 9), bg=_CARD_OFF,
                 fg=_TEXT_SUB, anchor=tk.W, justify=tk.LEFT).pack(
            fill=tk.X, pady=(3, 0), padx=(26, 0))

        _bind_tree(self, "<Button-1>", lambda _: self._on_select(self.value))
        _bind_tree(self, "<Enter>",   self._hover_on)
        _bind_tree(self, "<Leave>",   self._hover_off)

    def _hover_on(self, _=None):
        if not self._selected:
            self.config(highlightbackground="#9ca3af")

    def _hover_off(self, _=None):
        if not self._selected:
            self.config(highlightbackground=_BORDER_OFF)

    def select(self, yes: bool):
        self._selected = yes
        bg = _CARD_ON if yes else _CARD_OFF
        bd = _BORDER_ON if yes else _BORDER_OFF
        self.config(highlightbackground=bd)
        _recolor(self, bg)


# ── メインウィンドウ ────────────────────────────

class FarmCreateWindow:
    """新規農場作成ウィンドウ"""

    def __init__(self, parent: tk.Tk, on_farm_created: Optional[Callable] = None):
        self.parent = parent
        self.on_farm_created = on_farm_created
        self.created_farm_path: Optional[Path] = None

        self.window = tk.Toplevel(parent)
        self.window.title("新規農場作成")
        self.window.geometry("620x720")
        self.window.minsize(560, 520)
        self.window.resizable(True, True)
        self.window.configure(bg=_BG)

        self.farm_name_var      = tk.StringVar()
        self.template_farm_var  = tk.StringVar()
        self.import_method      = "none"          # "milk_csv" | "bulk_excel" | "none"
        self.milk_csv_path_var  = tk.StringVar()
        self.bulk_excel_path_var = tk.StringVar()
        self.backup_interval_var = tk.StringVar(value="1")
        self.backup_dest_var     = tk.StringVar()

        self._cards: dict[str, _Card] = {}
        self._detail_frames: dict[str, tk.Frame] = {}

        self._build()
        self._load_template_farms()
        self._adjust_initial_geometry()

    # ──────────────────────────────────────────
    #  UI 構築
    # ──────────────────────────────────────────

    def _build(self):
        # ── ヘッダー帯 ──
        header = tk.Frame(self.window, bg=_PRIMARY, height=72)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(
            header,
            text="新規農場を始めましょう",
            font=(_FONT, 14, "bold"), bg=_PRIMARY, fg="white"
        ).place(relx=0.5, rely=0.42, anchor=tk.CENTER)
        tk.Label(
            header,
            text="農場名を入力して、4つのステップで設定してください",
            font=(_FONT, 9), bg=_PRIMARY, fg="#c5cae9"
        ).place(relx=0.5, rely=0.76, anchor=tk.CENTER)

        # ── フッター（先に pack して下端固定） ──
        footer = tk.Frame(self.window, bg="#eaecf0", height=60)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)

        btn_cancel = tk.Button(
            footer, text="キャンセル",
            font=(_FONT, 10), relief=tk.FLAT,
            bg="#eaecf0", fg=_TEXT, activebackground="#dde1e7",
            cursor="hand2", bd=0, padx=20, pady=8,
            command=self._on_cancel
        )
        btn_cancel.pack(side=tk.RIGHT, padx=(6, 20), pady=12)

        btn_create = tk.Button(
            footer, text="  作成する  ",
            font=(_FONT, 10, "bold"), relief=tk.FLAT,
            bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
            cursor="hand2", bd=0, padx=20, pady=8,
            command=self._on_create
        )
        btn_create.pack(side=tk.RIGHT, padx=6, pady=12)

        # ── スクロール可能な本文 ──
        scroll_outer = tk.Frame(self.window, bg=_BG)
        scroll_outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(scroll_outer, bg=_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_outer, orient=tk.VERTICAL, command=canvas.yview)
        body = tk.Frame(canvas, bg=_BG)

        canvas_window = canvas.create_window((0, 0), window=body, anchor=tk.NW)

        def _sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(canvas_window, width=event.width)

        body.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        pad = tk.Frame(body, bg=_BG)
        pad.pack(fill=tk.BOTH, expand=True, padx=28, pady=20)
        body = pad

        # STEP 1: 農場名
        self._section(body, "STEP 1", "農場名")
        name_row = tk.Frame(body, bg=_BG)
        name_row.pack(fill=tk.X, pady=(4, 16))
        self.name_entry = tk.Entry(
            name_row, textvariable=self.farm_name_var,
            font=(_FONT, 11), relief=tk.FLAT,
            highlightthickness=1, highlightbackground=_BORDER_OFF,
            highlightcolor=_BORDER_ON, bg="white", fg=_TEXT
        )
        self.name_entry.pack(fill=tk.X, ipady=7)
        self.name_entry.focus_set()

        # STEP 2: テンプレート農場
        self._section(body, "STEP 2", "テンプレート農場")
        tmpl_outer = tk.Frame(body, bg=_BG)
        tmpl_outer.pack(fill=tk.X, pady=(4, 4))
        tk.Label(tmpl_outer,
                 text="既存農場の設定（ペン・授精師・繁殖処置・薬品など）を引き継ぐ場合に選択します",
                 font=(_FONT, 9), bg=_BG, fg=_TEXT_SUB).pack(anchor=tk.W)
        self.template_combo = ttk.Combobox(
            tmpl_outer, textvariable=self.template_farm_var,
            state="readonly", font=(_FONT, 10)
        )
        self.template_combo.pack(fill=tk.X, pady=(4, 0))

        # STEP 3: 初期データの登録方法
        sep = tk.Frame(body, bg=_BORDER_OFF, height=1)
        sep.pack(fill=tk.X, pady=(16, 12))
        self._section(body, "STEP 3", "初期データの登録方法を選択")

        cards_info = [
            ("milk_csv",   "📋",
             "乳検速報CSVから登録",
             "検定日・産次・乳検成績を一括インポートします（泌乳牛対象）"),
            ("bulk_excel", "📊",
             "個体一括登録 Excel から登録",
             "子牛・育成牛・乾乳牛を含む全頭をテンプレートで一括登録します"),
            ("none",       "✏️",
             "空の農場を作成",
             "後から1頭ずつ、または乳検吸い込みで随時登録します"),
        ]
        for value, icon, title, subtitle in cards_info:
            card = _Card(body, value, icon, title, subtitle, self._select_method)
            card.pack(fill=tk.X, pady=4)
            self._cards[value] = card

            detail = tk.Frame(body, bg=_BG)
            detail.pack(fill=tk.X, pady=(0, 4))
            self._detail_frames[value] = detail
            self._build_detail(detail, value)
            detail.pack_forget()

        # STEP 4: バックアップ
        sep2 = tk.Frame(body, bg=_BORDER_OFF, height=1)
        sep2.pack(fill=tk.X, pady=(16, 12))
        self._section(body, "STEP 4", "自動バックアップ設定")
        tk.Label(
            body,
            text="データはこのPC上に保存されます。消失リスクを下げるため、自動バックアップは事実上必須です。"
                  "下の間隔と保存先を設定してください。",
            font=(_FONT, 9), bg=_BG, fg=_TEXT, justify=tk.LEFT, wraplength=540, anchor=tk.W
        ).pack(fill=tk.X, pady=(4, 8))

        backup_card = tk.Frame(
            body, bg=_ACCENT, padx=14, pady=12,
            highlightthickness=1, highlightbackground="#c5cae9"
        )
        backup_card.pack(fill=tk.X, pady=(0, 6))

        iv_row = tk.Frame(backup_card, bg=_ACCENT)
        iv_row.pack(fill=tk.X)
        tk.Label(iv_row, text="バックアップの間隔（", font=(_FONT, 9), bg=_ACCENT, fg=_TEXT).pack(side=tk.LEFT)
        self.backup_interval_combo = ttk.Combobox(
            iv_row, textvariable=self.backup_interval_var,
            values=["1", "2", "3", "6", "12"],
            width=4, state="readonly", font=(_FONT, 10)
        )
        self.backup_interval_combo.pack(side=tk.LEFT, padx=2)
        tk.Label(iv_row, text="）か月ごと", font=(_FONT, 9), bg=_ACCENT, fg=_TEXT).pack(side=tk.LEFT)

        tk.Label(
            backup_card, text="バックアップ先フォルダ",
            font=(_FONT, 9), bg=_ACCENT, fg=_TEXT
        ).pack(anchor=tk.W, pady=(10, 0))
        dest_row = tk.Frame(backup_card, bg=_ACCENT)
        dest_row.pack(fill=tk.X, pady=(4, 0))
        self.backup_dest_entry = tk.Entry(
            dest_row, textvariable=self.backup_dest_var,
            font=(_FONT, 9), relief=tk.FLAT,
            highlightthickness=1, highlightbackground=_BORDER_OFF,
            highlightcolor=_BORDER_ON, bg="white", fg=_TEXT
        )
        self.backup_dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        tk.Button(
            dest_row, text="参照…", font=(_FONT, 9), relief=tk.FLAT,
            bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
            cursor="hand2", padx=10, pady=5,
            command=self._browse_backup_dest
        ).pack(side=tk.LEFT, padx=(6, 0))

        # デフォルトで "none" を選択状態にする
        self._select_method("none")

        # 本文内のすべての子でホイールスクロール（構築後にバインド）
        _bind_tree(body, "<MouseWheel>", _on_mousewheel)

    def _adjust_initial_geometry(self):
        """初回表示時に必要サイズへ調整し、下部ボタンの欠けを防ぐ。"""
        self.window.update_idletasks()
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()

        req_w = self.window.winfo_reqwidth() + 24
        req_h = self.window.winfo_reqheight() + 24

        width = min(max(req_w, 620), max(620, screen_w - 80))
        height = min(max(req_h, 560), max(520, screen_h - 80))
        self.window.geometry(f"{width}x{height}")

    def _section(self, parent, step_label: str, title: str):
        row = tk.Frame(parent, bg=_BG)
        row.pack(fill=tk.X, pady=(0, 2))
        badge = tk.Label(row, text=step_label, font=(_FONT, 8, "bold"),
                         bg=_PRIMARY, fg="white", padx=6, pady=2)
        badge.pack(side=tk.LEFT)
        tk.Label(row, text=f"  {title}", font=(_FONT, 10, "bold"),
                 bg=_BG, fg=_TEXT).pack(side=tk.LEFT)

    def _build_detail(self, frame: tk.Frame, value: str):
        """カード選択時に表示する追加 UI を構築"""
        if value == "milk_csv":
            inner = tk.Frame(frame, bg=_ACCENT, padx=14, pady=10,
                             highlightthickness=1, highlightbackground="#c5cae9")
            inner.pack(fill=tk.X)
            tk.Label(inner, text="乳検速報CSVファイル", font=(_FONT, 9),
                     bg=_ACCENT, fg=_TEXT).pack(anchor=tk.W)
            row = tk.Frame(inner, bg=_ACCENT)
            row.pack(fill=tk.X, pady=(4, 0))
            self.milk_csv_entry = tk.Entry(
                row, textvariable=self.milk_csv_path_var,
                font=(_FONT, 9), relief=tk.FLAT, state=tk.DISABLED,
                highlightthickness=1, highlightbackground=_BORDER_OFF, bg="white"
            )
            self.milk_csv_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
            tk.Button(row, text="参照…", font=(_FONT, 9), relief=tk.FLAT,
                      bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
                      cursor="hand2", padx=10, pady=5,
                      command=self._browse_milk_csv).pack(side=tk.LEFT, padx=(6, 0))

        elif value == "bulk_excel":
            inner = tk.Frame(frame, bg=_ACCENT, padx=14, pady=10,
                             highlightthickness=1, highlightbackground="#c5cae9")
            inner.pack(fill=tk.X)
            tk.Label(inner, text="個体一括登録 Excel ファイル", font=(_FONT, 9),
                     bg=_ACCENT, fg=_TEXT).pack(anchor=tk.W)
            row = tk.Frame(inner, bg=_ACCENT)
            row.pack(fill=tk.X, pady=(4, 0))
            self.bulk_excel_entry = tk.Entry(
                row, textvariable=self.bulk_excel_path_var,
                font=(_FONT, 9), relief=tk.FLAT, state=tk.DISABLED,
                highlightthickness=1, highlightbackground=_BORDER_OFF, bg="white"
            )
            self.bulk_excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
            tk.Button(row, text="参照…", font=(_FONT, 9), relief=tk.FLAT,
                      bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
                      cursor="hand2", padx=10, pady=5,
                      command=self._browse_bulk_excel).pack(side=tk.LEFT, padx=(6, 0))
            tk.Button(row, text="テンプレートを出力", font=(_FONT, 9), relief=tk.FLAT,
                      bg="#6c757d", fg="white", activebackground="#5a6268",
                      cursor="hand2", padx=10, pady=5,
                      command=self._export_bulk_template).pack(side=tk.LEFT, padx=(6, 0))
            tk.Label(inner,
                     text="※ テンプレートをダウンロード → Excelに記入 → 上の参照から選択",
                     font=(_FONT, 8), bg=_ACCENT, fg=_TEXT_SUB).pack(anchor=tk.W, pady=(6, 0))

    # ──────────────────────────────────────────
    #  カード選択
    # ──────────────────────────────────────────

    def _select_method(self, value: str):
        self.import_method = value
        for v, card in self._cards.items():
            card.select(v == value)
            detail = self._detail_frames[v]
            if v == value and v != "none":
                detail.pack(fill=tk.X, pady=(0, 4))
            else:
                detail.pack_forget()

        # Entry の有効/無効切り替え
        if hasattr(self, "milk_csv_entry"):
            st = tk.NORMAL if value == "milk_csv" else tk.DISABLED
            self.milk_csv_entry.config(state=st)
        if hasattr(self, "bulk_excel_entry"):
            st = tk.NORMAL if value == "bulk_excel" else tk.DISABLED
            self.bulk_excel_entry.config(state=st)

    # ──────────────────────────────────────────
    #  テンプレート農場
    # ──────────────────────────────────────────

    def _load_template_farms(self):
        if not constants.FARMS_ROOT.exists():
            self.template_combo['values'] = []
            return
        farms = [
            item.name for item in constants.FARMS_ROOT.iterdir()
            if item.is_dir() and (item / "farm.db").exists()
        ]
        self.template_combo['values'] = ["（なし）"] + sorted(farms)
        self.template_combo.current(0)

    # ──────────────────────────────────────────
    #  ファイル参照
    # ──────────────────────────────────────────

    def _browse_milk_csv(self):
        fp = filedialog.askopenfilename(
            title="乳検速報CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if fp:
            self.milk_csv_path_var.set(fp)

    def _browse_bulk_excel(self):
        fp = filedialog.askopenfilename(
            title="個体一括登録Excelファイルを選択",
            filetypes=[("Excelファイル", "*.xlsx *.xls"), ("すべてのファイル", "*.*")]
        )
        if fp:
            self.bulk_excel_path_var.set(fp)

    def _browse_backup_dest(self):
        path = filedialog.askdirectory(
            title="バックアップ先フォルダを選択", parent=self.window
        )
        if path:
            self.backup_dest_var.set(path)

    # ──────────────────────────────────────────
    #  テンプレート出力
    # ──────────────────────────────────────────

    def _export_bulk_template(self):
        prompt_save_bulk_cow_registration_template(parent=self.window)

    # ──────────────────────────────────────────
    #  農場作成
    # ──────────────────────────────────────────

    def _on_create(self):
        farm_name = self.farm_name_var.get().strip()
        if not farm_name:
            messagebox.showerror("エラー", "農場名を入力してください")
            self.name_entry.focus_set()
            return

        # ── ライセンス農場数上限チェック ──────────────────────────
        try:
            import builtins
            _lic = getattr(builtins, "_falcon_license", None)
            if _lic is not None:
                farm_limit = getattr(_lic, "farm_limit", 999)
                existing_farms = [
                    p for p in constants.FARMS_ROOT.iterdir()
                    if p.is_dir() and (p / "farm.db").exists()
                ] if constants.FARMS_ROOT.exists() else []
                if len(existing_farms) >= farm_limit:
                    messagebox.showerror(
                        "農場数上限",
                        f"現在のライセンスでは最大 {farm_limit} 農場まで登録できます。\n"
                        "ライセンスのアップグレードについてはサポートにお問い合わせください。"
                    )
                    return
        except Exception:
            pass

        farm_path = constants.FARMS_ROOT / farm_name

        if farm_path.exists() and (farm_path / "farm.db").exists():
            if not messagebox.askyesno("確認",
                    f"農場「{farm_name}」は既に存在します。\n上書きしますか？"):
                return

        method     = self.import_method
        csv_path   = None
        excel_path = None

        if method == "milk_csv":
            p = self.milk_csv_path_var.get().strip()
            if not p:
                messagebox.showerror("エラー", "乳検速報CSVファイルを選択してください")
                return
            csv_path = Path(p)
            if not csv_path.exists():
                messagebox.showerror("エラー", f"CSVファイルが見つかりません:\n{csv_path}")
                return

        elif method == "bulk_excel":
            p = self.bulk_excel_path_var.get().strip()
            if not p:
                messagebox.showerror("エラー", "個体一括登録Excelファイルを選択してください")
                return
            excel_path = Path(p)
            if not excel_path.exists():
                messagebox.showerror("エラー", f"Excelファイルが見つかりません:\n{excel_path}")
                return

        template_farm_path = None
        template_name = self.template_farm_var.get()
        if template_name and template_name != "（なし）":
            template_farm_path = constants.FARMS_ROOT / template_name
            if not (template_farm_path / "farm.db").exists():
                messagebox.showerror("エラー",
                    f"テンプレート農場が見つかりません:\n{template_farm_path}")
                return

        try:
            interval = int(self.backup_interval_var.get())
        except ValueError:
            messagebox.showerror("エラー", "バックアップの間隔は 1, 2, 3, 6, 12 のいずれかを選択してください")
            return
        if interval not in (1, 2, 3, 6, 12):
            messagebox.showerror("エラー", "バックアップの間隔は 1, 2, 3, 6, 12 のいずれかを選択してください")
            return
        backup_dest = self.backup_dest_var.get().strip()

        try:
            creator = FarmCreator(farm_path)
            creator.create_farm(
                farm_name=farm_name,
                csv_path=csv_path,
                excel_path=excel_path,
                template_farm_path=template_farm_path
            )
            save_backup_settings(farm_path, interval, backup_dest)
            self.created_farm_path = farm_path
            messagebox.showinfo("完了", f"農場「{farm_name}」を作成しました")
            self.window.destroy()
            if self.on_farm_created:
                self.on_farm_created(farm_path)

        except Exception as e:
            messagebox.showerror("エラー", f"農場の作成に失敗しました:\n{e}")
            import traceback
            traceback.print_exc()

    def _on_cancel(self):
        self.window.destroy()

    def show(self):
        self.window.wait_window()
        return self.created_farm_path
