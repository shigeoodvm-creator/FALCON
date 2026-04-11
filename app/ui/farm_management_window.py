"""
FALCON2 - 農場管理ウィンドウ
農場の切り替えと新規作成を管理
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Callable, cast

from ui.farm_selector import FarmSelectorWindow
from ui.farm_create_window import FarmCreateWindow
from settings_manager import SettingsManager

# デザイン定数（他ウィンドウと統一）
_BG = "#f5f5f5"
_CARD_BG = "#ffffff"
_CARD_BORDER = "#e0e0e0"
_ACCENT = "#1e2a3a"
_ACCENT_HOVER = "#2d3e50"
_TEXT_PRIMARY = "#263238"
_TEXT_SECONDARY = "#607d8b"
_FONT = "Meiryo UI"


class FarmManagementWindow:
    """農場管理ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, current_farm_path: Optional[Path] = None,
                 on_farm_changed: Optional[Callable[[Path], None]] = None,
                 on_farm_name_changed: Optional[Callable[[str], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            current_farm_path: 現在開いている農場のパス
            on_farm_changed: 農場が変更された時のコールバック関数（farm_path を引数に取る）
        """
        self.parent = parent
        self.current_farm_path = current_farm_path
        self.on_farm_changed = on_farm_changed
        self.on_farm_name_changed = on_farm_name_changed
        self._display_name_var = tk.StringVar()
        
        self.window = tk.Toplevel(parent)
        self.window.title("農場管理")
        self.window.minsize(440, 420)
        self.window.geometry("520x480")
        self.window.configure(bg=_BG)
        
        self._create_widgets()
        self._adjust_initial_geometry()
        self._center_on_screen()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main = tk.Frame(self.window, bg=_BG, padx=28, pady=24)
        main.pack(fill=tk.BOTH, expand=True)
        main.grid_columnconfigure(0, weight=1)
        # 内容の高さに合わせる（余白だけ広がって閉じるが画面外に押し出されるのを防ぐ）
        main.grid_rowconfigure(2, weight=0)

        # ヘッダー
        header = tk.Frame(main, bg=_BG)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 14))
        tk.Label(
            header,
            text="農場管理",
            font=(_FONT, 18, "bold"),
            bg=_BG,
            fg=_TEXT_PRIMARY
        ).pack(anchor=tk.W)
        tk.Label(
            header,
            text="農場の切り替えや新規作成を行います",
            font=(_FONT, 10),
            bg=_BG,
            fg=_TEXT_SECONDARY
        ).pack(anchor=tk.W, pady=(2, 0))

        # 現在の農場カード
        farm_card = tk.Frame(main, bg=_CARD_BORDER, padx=1, pady=1)
        farm_card.grid(row=1, column=0, sticky=tk.EW, pady=(0, 16))
        farm_inner = tk.Frame(farm_card, bg=_CARD_BG, padx=14, pady=12)
        farm_inner.pack(fill=tk.X)

        farm_name = self.current_farm_path.name if self.current_farm_path else "なし"
        if self.current_farm_path:
            try:
                sm = SettingsManager(self.current_farm_path)
                farm_name = sm.get("farm_name", self.current_farm_path.name)
            except Exception:
                farm_name = self.current_farm_path.name
        self._display_name_var.set(farm_name if self.current_farm_path else "")
        farm_fg = _TEXT_PRIMARY if self.current_farm_path else _TEXT_SECONDARY

        tk.Label(
            farm_inner,
            text="現在の農場",
            font=(_FONT, 9),
            bg=_CARD_BG,
            fg=_TEXT_SECONDARY
        ).pack(anchor=tk.W)
        tk.Label(
            farm_inner,
            text=farm_name,
            font=(_FONT, 13, "bold"),
            bg=_CARD_BG,
            fg=farm_fg
        ).pack(anchor=tk.W, pady=(2, 0))

        tk.Frame(farm_inner, height=1, bg=_CARD_BORDER).pack(fill=tk.X, pady=(10, 10))
        tk.Label(
            farm_inner,
            text="表示名（アプリ表示用）",
            font=(_FONT, 9),
            bg=_CARD_BG,
            fg=_TEXT_SECONDARY
        ).pack(anchor=tk.W)

        name_row = tk.Frame(farm_inner, bg=_CARD_BG)
        name_row.pack(fill=tk.X, pady=(6, 0))
        self._display_name_entry = ttk.Entry(
            name_row,
            textvariable=self._display_name_var,
            font=(_FONT, 10)
        )
        self._display_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._display_name_entry.bind("<Return>", lambda _e: self._on_save_farm_display_name())

        self._save_name_btn = tk.Button(
            name_row,
            text="保存",
            font=(_FONT, 9),
            bg="#eceff1",
            fg="#37474f",
            activebackground="#cfd8dc",
            relief=tk.FLAT,
            padx=10,
            pady=6,
            cursor="hand2",
            highlightbackground="#b0bec5",
            highlightthickness=1,
            command=self._on_save_farm_display_name
        )
        self._save_name_btn.pack(side=tk.LEFT, padx=(8, 0))

        # 操作ボタンカード
        action_card = tk.Frame(main, bg=_CARD_BORDER, padx=1, pady=1)
        action_card.grid(row=2, column=0, sticky=tk.NSEW)
        action_inner = tk.Frame(action_card, bg=_CARD_BG, padx=14, pady=14)
        action_inner.pack(fill=tk.X)

        switch_farm_btn = tk.Button(
            action_inner,
            text="農場切り替え",
            font=(_FONT, 11, "bold"),
            bg=_ACCENT,
            fg="white",
            activebackground=_ACCENT_HOVER,
            activeforeground="white",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=9,
            cursor="hand2",
            command=self._on_switch_farm
        )
        switch_farm_btn.pack(fill=tk.X, pady=(0, 10))
        switch_farm_btn.bind("<Enter>", lambda _: switch_farm_btn.config(bg=_ACCENT_HOVER))
        switch_farm_btn.bind("<Leave>", lambda _: switch_farm_btn.config(bg=_ACCENT))

        create_farm_btn = tk.Button(
            action_inner,
            text="農場新規作成",
            font=(_FONT, 10),
            bg="#eaecf0",
            fg=_TEXT_PRIMARY,
            activebackground="#dde1e7",
            activeforeground=_TEXT_PRIMARY,
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=9,
            cursor="hand2",
            command=self._on_create_farm
        )
        create_farm_btn.pack(fill=tk.X, pady=(0, 8))

        tk.Frame(action_inner, height=1, bg=_CARD_BORDER).pack(fill=tk.X, pady=(4, 8))

        close_btn = ttk.Button(
            action_inner,
            text="閉じる",
            command=self._on_close
        )
        close_btn.pack(anchor=tk.E)
    
    def _adjust_initial_geometry(self):
        """描画後の必要サイズに合わせ、閉じるボタンが欠けないようにする。"""
        self.window.update_idletasks()
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        req_w = self.window.winfo_reqwidth() + 32
        req_h = self.window.winfo_reqheight() + 32
        width = min(max(req_w, 520), max(520, screen_w - 80))
        height = min(max(req_h, 460), max(420, screen_h - 80))
        self.window.geometry(f"{width}x{height}")
    
    def _center_on_screen(self):
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _on_switch_farm(self):
        """農場切り替えボタンをクリック"""
        def on_farm_selected(farm_path: Path):
            """農場が選択された時のコールバック"""
            if self.on_farm_changed:
                self.on_farm_changed(farm_path)
            # ウィンドウを閉じる（切り替え処理はコールバック内で行われる）
            self.window.destroy()
        
        # 農場選択ウィンドウを表示
        farm_selector = FarmSelectorWindow(
            parent=cast(tk.Tk, self.window),
            on_farm_selected=on_farm_selected
        )
        farm_selector.show()
    
    def _on_create_farm(self):
        """農場新規作成ボタンをクリック"""
        def on_farm_created(farm_path: Path):
            """農場が作成された時のコールバック"""
            if self.on_farm_changed:
                self.on_farm_changed(farm_path)
            # ウィンドウを閉じる（切り替え処理はコールバック内で行われる）
            self.window.destroy()
        
        # 新規農場作成ウィンドウを表示
        create_window = FarmCreateWindow(
            parent=cast(tk.Tk, self.window),
            on_farm_created=on_farm_created
        )
        created_path = create_window.show()
        
        # 作成された場合は自動的に切り替える（コールバックはFarmCreateWindow内で呼ばれる）
        if created_path and self.on_farm_changed:
            self.on_farm_changed(created_path)
            self.window.destroy()
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()

    def _on_save_farm_display_name(self):
        """農場表示名を farm_settings.json に保存"""
        if not self.current_farm_path:
            messagebox.showwarning("警告", "現在の農場が選択されていません。")
            return

        new_name = (self._display_name_var.get() or "").strip()
        if not new_name:
            messagebox.showwarning("警告", "農場名を入力してください。")
            return
        if len(new_name) > 50:
            messagebox.showwarning("警告", "農場名は50文字以内で入力してください。")
            return
        if any(ch in new_name for ch in ("\r", "\n", "\t")):
            messagebox.showwarning("警告", "改行やタブを含む名前は使用できません。")
            return

        try:
            sm = SettingsManager(self.current_farm_path)
            old_name = sm.get("farm_name", self.current_farm_path.name)
            sm.set("farm_name", new_name)
            if hasattr(self, "_display_name_entry") and self._display_name_entry.winfo_exists():
                self._display_name_entry.selection_clear()
            if self.on_farm_name_changed:
                self.on_farm_name_changed(new_name)
            if old_name != new_name:
                messagebox.showinfo("完了", f"農場表示名を更新しました。\n{old_name} → {new_name}")
        except Exception as e:
            messagebox.showerror("エラー", f"農場名の保存に失敗しました。\n{e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
