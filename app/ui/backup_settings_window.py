"""
FALCON2 - バックアップ設定ウィンドウ
バックアップの間隔と保存先を設定
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional

from backup_manager import load_backup_settings, save_backup_settings


class BackupSettingsWindow:
    """バックアップ設定ウィンドウ"""

    def __init__(self, parent: tk.Tk, farm_path: Path):
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.window = tk.Toplevel(parent)
        self.window.title("バックアップ設定")
        self.window.minsize(420, 260)
        self.window.geometry("480x280")

        self._create_widgets()
        self._load_current_settings()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main = ttk.Frame(self.window, padding=16)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="バックアップ設定", font=("", 12, "bold")).pack(pady=(0, 12))

        # バックアップの間隔（○か月ごと）
        interval_frame = ttk.Frame(main)
        interval_frame.pack(fill=tk.X, pady=6)
        ttk.Label(interval_frame, text="バックアップの間隔（").pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="1")
        interval_combo = ttk.Combobox(
            interval_frame,
            textvariable=self.interval_var,
            values=["1", "2", "3", "6", "12"],
            width=4,
            state="readonly",
        )
        interval_combo.pack(side=tk.LEFT, padx=2)
        ttk.Label(interval_frame, text="）か月ごと").pack(side=tk.LEFT)

        # バックアップ先
        dest_frame = ttk.Frame(main)
        dest_frame.pack(fill=tk.X, pady=6)
        ttk.Label(dest_frame, text="バックアップ先：").pack(anchor=tk.W)
        row = ttk.Frame(dest_frame)
        row.pack(fill=tk.X, pady=4)
        self.dest_var = tk.StringVar()
        dest_entry = ttk.Entry(row, textvariable=self.dest_var, width=40)
        dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(row, text="参照…", command=self._browse_destination).pack(side=tk.LEFT)

        # 説明
        note = (
            "その月にアプリを初めて開いたときに、農場フォルダを指定先に日付付きでコピーします。\n"
            "例：20250205デモファーム。データベース破損時は、このフォルダをFARMS内の農場フォルダと置き換えて復旧できます。"
        )
        ttk.Label(main, text=note, font=("", 8), foreground="gray").pack(anchor=tk.W, pady=8)

        # ボタン
        btn_frame = ttk.Frame(main)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="保存", command=self._on_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="キャンセル", command=self.window.destroy).pack(side=tk.LEFT, padx=4)

    def _load_current_settings(self):
        s = load_backup_settings(self.farm_path)
        self.interval_var.set(str(s.get("interval_months", 1)))
        self.dest_var.set(s.get("destination_path") or "")

    def _browse_destination(self):
        path = filedialog.askdirectory(title="バックアップ先フォルダを選択", parent=self.window)
        if path:
            self.dest_var.set(path)

    def _on_save(self):
        try:
            interval = int(self.interval_var.get())
        except ValueError:
            messagebox.showerror("エラー", "間隔には 1, 2, 3, 6, 12 のいずれかを選択してください。", parent=self.window)
            return
        if interval not in (1, 2, 3, 6, 12):
            messagebox.showerror("エラー", "間隔には 1, 2, 3, 6, 12 のいずれかを選択してください。", parent=self.window)
            return
        dest = self.dest_var.get().strip()
        if not dest:
            messagebox.showwarning("確認", "バックアップ先を指定すると、その月の初回起動時に自動でバックアップが実行されます。\n空のまま保存すると自動バックアップは行われません。", parent=self.window)
        try:
            save_backup_settings(self.farm_path, interval, dest)
            messagebox.showinfo("保存完了", "バックアップ設定を保存しました。", parent=self.window)
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}", parent=self.window)

    def show(self):
        self.window.wait_window()
