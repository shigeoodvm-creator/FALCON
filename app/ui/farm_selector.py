"""
FALCON2 - 農場選択ウィンドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from pathlib import Path
from typing import Optional, Callable

import constants  # FARMS_ROOT を動的参照するため module ごと import
from ui.farm_create_window import FarmCreateWindow

# デザイン定数
_BG           = "#f5f5f5"
_CARD_BG      = "#ffffff"
_CARD_BORDER  = "#e0e0e0"
_ACCENT       = "#1e2a3a"
_ACCENT_HOVER = "#2d3e50"
_TEXT_PRIMARY = "#263238"
_TEXT_SECONDARY = "#607d8b"
_SELECT_BG    = "#e8eaf6"
_FONT         = "Meiryo UI"


class FarmSelectorWindow:
    """農場選択ウィンドウ"""

    def __init__(self, parent: tk.Tk, on_farm_selected: Optional[Callable] = None):
        self.parent = parent
        self.on_farm_selected = on_farm_selected
        self.selected_farm_path: Optional[Path] = None

        self.window = tk.Toplevel(parent)
        self.window.title("農場選択")
        self.window.geometry("480x480")
        self.window.minsize(420, 420)
        self.window.configure(bg=_BG)

        self._ensure_farms_root_configured()
        self._create_widgets()
        self._load_farms()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth()  // 2) - (self.window.winfo_width()  // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        self.window.lift()
        self.window.focus_force()

    # ------------------------------------------------------------------ #
    #  初回セットアップ（FARMSルート）
    # ------------------------------------------------------------------ #

    def _is_farms_root_configured(self) -> bool:
        """app_config.json に farms_root が設定済みかを判定する。"""
        app_config_path = constants.APP_CONFIG_DIR / "app_config.json"
        if not app_config_path.exists():
            return False
        try:
            with open(app_config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            value = cfg.get("farms_root")
            return bool(str(value).strip()) if value is not None else False
        except Exception:
            return False

    def _ensure_farms_root_configured(self):
        """
        FARMSルートが未設定の場合に初回セットアップを行う。
        設定済みの場合は何もしない。
        """
        if self._is_farms_root_configured():
            return

        default_path = constants.FARMS_ROOT
        choice = messagebox.askyesnocancel(
            "初回設定",
            "農場データの保存先（FARMSフォルダ）が未設定です。\n\n"
            f"既定の保存先を使用しますか？\n{default_path}\n\n"
            "「いいえ」を選ぶと保存先を手動で選択できます。"
        )

        if choice is None:
            self.window.destroy()
            self.parent.quit()
            return

        selected_path: Optional[Path] = None
        if choice is True:
            selected_path = default_path
        else:
            selected = filedialog.askdirectory(
                title="農場データの保存先フォルダを選択",
                initialdir=str(Path.home()),
                mustexist=False
            )
            if not selected:
                messagebox.showwarning("中止", "保存先が選択されなかったため終了します。")
                self.window.destroy()
                self.parent.quit()
                return
            selected_path = Path(selected)

        try:
            selected_path.mkdir(parents=True, exist_ok=True)
            constants.set_farms_root(selected_path)
            messagebox.showinfo(
                "設定完了",
                "農場データの保存先を設定しました。\n\n"
                f"{selected_path}"
            )
        except Exception as e:
            messagebox.showerror(
                "エラー",
                f"保存先フォルダの設定に失敗しました。\n\n{selected_path}\n\n{e}"
            )
            self.window.destroy()
            self.parent.quit()

    # ------------------------------------------------------------------ #
    #  UI 構築
    # ------------------------------------------------------------------ #

    def _create_widgets(self):
        main = tk.Frame(self.window, bg=_BG, padx=28, pady=24)
        main.pack(fill=tk.BOTH, expand=True)
        main.grid_columnconfigure(0, weight=1)

        # ── ヘッダー ──
        header = tk.Frame(main, bg=_BG)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 12))

        tk.Label(header, text="農場選択",
                 font=(_FONT, 17, "bold"), bg=_BG, fg=_TEXT_PRIMARY).pack(anchor=tk.W)
        tk.Label(header, text="農場を選択してください",
                 font=(_FONT, 10), bg=_BG, fg=_TEXT_SECONDARY).pack(anchor=tk.W)

        # ── FARMSフォルダ表示行 ──
        folder_row = tk.Frame(main, bg=_BG)
        folder_row.grid(row=1, column=0, sticky=tk.EW, pady=(0, 4))

        tk.Label(folder_row, text="FARMSフォルダ:", font=(_FONT, 9),
                 bg=_BG, fg=_TEXT_SECONDARY).pack(side=tk.LEFT)

        self.folder_label = tk.Label(
            folder_row, text=str(constants.FARMS_ROOT),
            font=(_FONT, 9), bg=_BG, fg=_TEXT_PRIMARY, anchor=tk.W
        )
        self.folder_label.pack(side=tk.LEFT, padx=(4, 8))

        change_btn = tk.Label(
            folder_row, text="変更",
            font=(_FONT, 9, "underline"), bg=_BG,
            fg=_ACCENT, cursor="hand2"
        )
        change_btn.pack(side=tk.LEFT)
        change_btn.bind("<Button-1>", lambda _: self._change_folder())
        change_btn.bind("<Enter>", lambda _: change_btn.config(fg=_ACCENT_HOVER))
        change_btn.bind("<Leave>", lambda _: change_btn.config(fg=_ACCENT))

        # ── 農場リスト ──
        list_card = tk.Frame(main, bg=_CARD_BORDER, padx=1, pady=1)
        list_card.grid(row=2, column=0, sticky=tk.NSEW, pady=(8, 16))
        list_card.grid_columnconfigure(0, weight=1)
        list_card.grid_rowconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        list_inner = tk.Frame(list_card, bg=_CARD_BG)
        list_inner.grid(row=0, column=0, sticky=tk.NSEW)
        list_inner.grid_columnconfigure(0, weight=1)
        list_inner.grid_rowconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(list_inner)
        scrollbar.grid(row=0, column=1, sticky=tk.NS, pady=2)

        self.farm_listbox = tk.Listbox(
            list_inner,
            yscrollcommand=scrollbar.set,
            font=(_FONT, 11),
            selectbackground=_SELECT_BG,
            selectforeground=_TEXT_PRIMARY,
            highlightthickness=0,
            bd=0, relief=tk.FLAT,
            bg=_CARD_BG, fg=_TEXT_PRIMARY,
            activestyle=tk.NONE
        )
        self.farm_listbox.grid(row=0, column=0, sticky=tk.NSEW, padx=(8, 2), pady=6)
        scrollbar.config(command=self.farm_listbox.yview)
        self.farm_listbox.bind("<Double-Button-1>", lambda _: self._on_open_farm())

        # ── フッター ──
        footer = tk.Frame(main, bg=_BG)
        footer.grid(row=3, column=0, sticky=tk.EW, pady=(4, 0))
        footer.grid_columnconfigure(0, weight=1)

        tk.Frame(footer, height=1, bg=_CARD_BORDER).grid(
            row=0, column=0, sticky=tk.EW, pady=(0, 14))

        btn_frame = tk.Frame(footer, bg=_BG)
        btn_frame.grid(row=1, column=0, sticky=tk.E)

        ttk.Button(btn_frame, text="キャンセル",
                   command=self._on_cancel, width=11).pack(side=tk.RIGHT, padx=(8, 0))

        ttk.Button(btn_frame, text="新規農場作成",
                   command=self._on_create_farm, width=12).pack(side=tk.RIGHT, padx=(8, 0))

        open_btn = tk.Button(
            btn_frame, text="農場を開く",
            font=(_FONT, 10), bg=_ACCENT, fg="white",
            activebackground=_ACCENT_HOVER, activeforeground="white",
            relief=tk.FLAT, bd=0, padx=16, pady=8, cursor="hand2",
            command=self._on_open_farm
        )
        open_btn.pack(side=tk.RIGHT)
        open_btn.bind("<Enter>", lambda _: open_btn.config(bg=_ACCENT_HOVER))
        open_btn.bind("<Leave>", lambda _: open_btn.config(bg=_ACCENT))

    # ------------------------------------------------------------------ #
    #  FARMSフォルダ変更
    # ------------------------------------------------------------------ #

    def _change_folder(self):
        """FARMSフォルダを変更する（説明ダイアログのあとフォルダ選択）"""
        explain = (
            "FARMSフォルダの保存場所を変えるときの説明です。\n\n"
            "【フォルダ構成】\n"
            "「FARMSフォルダ」の直下に、農場ごとのフォルダを置きます。\n\n"
            "  C:\\FARMS           ← FARMSフォルダ（ここを指定）\n"
            "     ├─ 牧場A         ← 農場ごとのフォルダ\n"
            "     ├─ 牧場B\n"
            "     └─ DEMO\n\n"
            "農場名のフォルダではなく、FARMSフォルダ自体\n"
            "（例: C:\\FARMS）を選んでください。\n\n"
            "「OK」でフォルダ選択、「キャンセル」で中止します。"
        )
        if not messagebox.askokcancel(
            "FARMSフォルダの変更", explain, parent=self.window
        ):
            return

        new_dir = filedialog.askdirectory(
            title="FARMSフォルダを選択（例: C:\\FARMS）",
            initialdir=str(constants.FARMS_ROOT),
            mustexist=False
        )
        if not new_dir:
            return

        new_path = Path(new_dir)

        # 存在しない場合は作成するか確認
        if not new_path.exists():
            if not messagebox.askyesno(
                "確認",
                f"フォルダが存在しません。作成しますか？\n\n{new_path}"
            ):
                return
            new_path.mkdir(parents=True, exist_ok=True)

        # constants.FARMS_ROOT を更新して設定ファイルに保存
        constants.set_farms_root(new_path)

        # 表示を更新
        self.folder_label.config(text=str(constants.FARMS_ROOT))
        self._load_farms()

        messagebox.showinfo(
            "変更しました",
            f"FARMSフォルダを変更しました。\n\n{new_path}\n\n"
            "この設定は次回起動時にも引き継がれます。"
        )

    # ------------------------------------------------------------------ #
    #  農場リスト
    # ------------------------------------------------------------------ #

    def _load_farms(self):
        farms_root = constants.FARMS_ROOT

        if not farms_root.exists():
            farms_root.mkdir(parents=True, exist_ok=True)

        farms = sorted(
            item.name for item in farms_root.iterdir()
            if item.is_dir() and (item / "farm.db").exists()
        ) if farms_root.exists() else []

        self.farm_listbox.delete(0, tk.END)
        if farms:
            for name in farms:
                self.farm_listbox.insert(tk.END, name)
        else:
            self.farm_listbox.insert(tk.END, "（農場が見つかりません）")

    # ------------------------------------------------------------------ #
    #  アクション
    # ------------------------------------------------------------------ #

    def _on_open_farm(self):
        selection = self.farm_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "農場を選択してください")
            return

        farm_name = self.farm_listbox.get(selection[0])
        if farm_name == "（農場が見つかりません）":
            messagebox.showinfo("情報",
                "「新規農場作成」から農場を作成するか、\n"
                "「変更」で既存のFARMSフォルダを指定してください。")
            return

        farm_path = constants.FARMS_ROOT / farm_name
        if not (farm_path / "farm.db").exists():
            messagebox.showerror("エラー", f"農場データが見つかりません:\n{farm_path}")
            return

        self.selected_farm_path = farm_path
        self.window.destroy()
        if self.on_farm_selected:
            self.on_farm_selected(farm_path)

    def _on_create_farm(self):
        def on_farm_created(farm_path: Path):
            self._load_farms()
            farm_name = farm_path.name
            for i in range(self.farm_listbox.size()):
                if self.farm_listbox.get(i) == farm_name:
                    self.farm_listbox.selection_clear(0, tk.END)
                    self.farm_listbox.selection_set(i)
                    self.farm_listbox.see(i)
                    break

        create_window = FarmCreateWindow(self.window, on_farm_created=on_farm_created)
        created_path = create_window.show()

        if created_path:
            self.selected_farm_path = created_path
            self.window.destroy()
            if self.on_farm_selected:
                self.on_farm_selected(created_path)

    def _on_cancel(self):
        self.window.destroy()
        self.parent.quit()

    def show(self):
        self.window.wait_window()
        return self.selected_farm_path
