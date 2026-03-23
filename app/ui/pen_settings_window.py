"""
PEN設定ウインドウ
農場ごとの PENコード / PEN名 を管理する
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict

from settings_manager import SettingsManager


class PenSettingsWindow:
    """PEN設定ウインドウ"""

    def __init__(self, parent: tk.Tk, farm_path: Path):
        """
        初期化

        Args:
            parent: 親ウィンドウ
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_manager = SettingsManager(self.farm_path)

        # 設定をロード
        self.pen_settings: Dict[str, str] = self.settings_manager.load_pen_settings()

        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("PEN設定")
        self.window.geometry("520x500")

        # UI作成
        self._create_widgets()

        # 中央配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview
        tree_frame = ttk.LabelFrame(main_frame, text="PEN一覧", padding=10)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("code", "name"),
            show="headings",
            yscrollcommand=tree_scroll.set,
            height=12
        )
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)

        self.tree.heading("code", text="PENコード")
        self.tree.heading("name", text="PEN名")
        self.tree.column("code", width=120, anchor=tk.CENTER)
        self.tree.column("name", width=280)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # 入力欄
        form_frame = ttk.LabelFrame(main_frame, text="編集", padding=10)
        form_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(form_frame, text="PENコード:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        vcmd = (self.window.register(self._validate_code), "%P")
        self.code_entry = ttk.Entry(form_frame, width=20, validate="key", validatecommand=vcmd)
        self.code_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form_frame, text="PEN名:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.name_entry = ttk.Entry(form_frame, width=30)
        self.name_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        # ボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=15)

        ttk.Button(
            button_frame,
            text="追加 / 更新",
            command=self._on_add_update
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="削除",
            command=self._on_delete
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close
        ).pack(side=tk.RIGHT, padx=5)

        # データ表示
        self._refresh_tree()

    @staticmethod
    def _validate_code(new_value: str) -> bool:
        """コード欄を数字のみ許可"""
        return new_value.isdigit() or new_value == ""

    def _refresh_tree(self):
        """Treeviewをリフレッシュ"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        def sort_key(item):
            code = item[0]
            try:
                return int(code)
            except ValueError:
                return float("inf")

        for code, name in sorted(self.pen_settings.items(), key=sort_key):
            self.tree.insert("", tk.END, values=(code, name))

    def _on_tree_select(self, _event=None):
        """選択行を入力欄に反映"""
        selection = self.tree.selection()
        if not selection:
            return
        item = self.tree.item(selection[0])
        code, name = item["values"]
        self.code_entry.delete(0, tk.END)
        self.code_entry.insert(0, code)
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, name)

    def _on_add_update(self):
        """追加 / 更新ボタン"""
        code = self.code_entry.get().strip()
        name = self.name_entry.get().strip()

        if not code:
            messagebox.showwarning("警告", "PENコードを入力してください")
            return
        if not name:
            messagebox.showwarning("警告", "PEN名を入力してください")
            return
        if not code.isdigit():
            messagebox.showwarning("警告", "PENコードは数字のみ入力できます")
            return

        # 更新または追加
        self.pen_settings[code] = name
        self.settings_manager.save_pen_settings(self.pen_settings)
        self._refresh_tree()
        messagebox.showinfo("完了", f"PEN設定を保存しました（{code}: {name}）")

    def _on_delete(self):
        """削除ボタン"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "削除するPENを選択してください")
            return

        item = self.tree.item(selection[0])
        code, name = item["values"]

        if messagebox.askyesno("確認", f"PEN「{code}: {name}」を削除しますか？"):
            if code in self.pen_settings:
                del self.pen_settings[code]
                self.settings_manager.save_pen_settings(self.pen_settings)
                self._refresh_tree()

    def _on_close(self):
        """閉じるボタン"""
        self.window.destroy()

    def show(self):
        """モーダル表示"""
        self.window.wait_window()






































