"""
FALCON2 - バックアップ設定ウィンドウ
バックアップの間隔と保存先を設定
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional

from backup_manager import load_backup_settings, save_backup_settings

# farm_create_window と同系のカラーパレット（画面の統一感のため）
_BG         = "#f5f6fa"
_BORDER_OFF = "#dde1e7"
_BORDER_ON  = "#1e2a3a"
_PRIMARY    = "#1e2a3a"
_PRIMARY_H  = "#2d3e50"
_TEXT       = "#1c2333"
_TEXT_SUB   = "#6b7280"
_ACCENT     = "#edf0f3"
_FONT       = "Meiryo UI"


class BackupSettingsWindow:
    """バックアップ設定ウィンドウ"""

    def __init__(self, parent: tk.Tk, farm_path: Path):
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.window = tk.Toplevel(parent)
        self.window.title("バックアップ設定")
        self.window.minsize(480, 420)
        self.window.geometry("560x480")
        self.window.configure(bg=_BG)

        self._create_widgets()
        self._load_current_settings()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        header = tk.Frame(self.window, bg=_PRIMARY, height=72)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(
            header, text="バックアップ設定",
            font=(_FONT, 14, "bold"), bg=_PRIMARY, fg="white"
        ).place(relx=0.5, rely=0.42, anchor=tk.CENTER)
        tk.Label(
            header,
            text="農場フォルダの自動コピー先と間隔を指定します",
            font=(_FONT, 9), bg=_PRIMARY, fg="#c5cae9"
        ).place(relx=0.5, rely=0.76, anchor=tk.CENTER)

        body = tk.Frame(self.window, bg=_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

        tk.Label(
            body,
            text="農場データはこのPC上のフォルダに保存されています。別のドライブやクラウドへ定期的にコピーしておくと、"
                  "故障・ウイルス・誤操作によるデータ消失に備えられます。",
            font=(_FONT, 9), bg=_BG, fg=_TEXT, justify=tk.LEFT, wraplength=500, anchor=tk.W
        ).pack(fill=tk.X, pady=(0, 10))

        card = tk.Frame(
            body, bg=_ACCENT, padx=14, pady=12,
            highlightthickness=1, highlightbackground="#c5cae9"
        )
        card.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            card, text="自動バックアップのしくみ",
            font=(_FONT, 9, "bold"), bg=_ACCENT, fg=_TEXT
        ).pack(anchor=tk.W)
        note = (
            "・指定した「か月ごと」の間隔で、その月に初めてアプリを開いたときに、"
            "農場フォルダ全体をバックアップ先へ日付＋農場名のフォルダ名でコピーします。\n"
            "・バックアップ先を空のまま保存すると、自動バックアップは行われません。\n"
            "・保存先の例：外付けHDD、USB、OneDrive / Google ドライブの同期フォルダなど。\n"
            "・例：フォルダ名 20250205デモ農場。データベース破損時は、このフォルダの内容で "
            "農場フォルダを差し替えることで復旧できる場合があります。"
        )
        tk.Label(
            card, text=note,
            font=(_FONT, 8), bg=_ACCENT, fg=_TEXT_SUB, justify=tk.LEFT,
            wraplength=500, anchor=tk.W
        ).pack(fill=tk.X, pady=(6, 12))

        iv_row = tk.Frame(card, bg=_ACCENT)
        iv_row.pack(fill=tk.X)
        tk.Label(iv_row, text="バックアップの間隔（", font=(_FONT, 9), bg=_ACCENT, fg=_TEXT).pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="1")
        self.interval_combo = ttk.Combobox(
            iv_row,
            textvariable=self.interval_var,
            values=["1", "2", "3", "6", "12"],
            width=4,
            state="readonly",
            font=(_FONT, 10),
        )
        self.interval_combo.pack(side=tk.LEFT, padx=2)
        tk.Label(iv_row, text="）か月ごと", font=(_FONT, 9), bg=_ACCENT, fg=_TEXT).pack(side=tk.LEFT)

        tk.Label(
            card, text="バックアップ先フォルダ",
            font=(_FONT, 9), bg=_ACCENT, fg=_TEXT
        ).pack(anchor=tk.W, pady=(12, 0))
        row = tk.Frame(card, bg=_ACCENT)
        row.pack(fill=tk.X, pady=(4, 0))
        self.dest_var = tk.StringVar()
        self.dest_entry = tk.Entry(
            row, textvariable=self.dest_var,
            font=(_FONT, 9), relief=tk.FLAT,
            highlightthickness=1, highlightbackground=_BORDER_OFF,
            highlightcolor=_BORDER_ON, bg="white", fg=_TEXT
        )
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        tk.Button(
            row, text="参照…", font=(_FONT, 9), relief=tk.FLAT,
            bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
            cursor="hand2", padx=10, pady=6,
            command=self._browse_destination
        ).pack(side=tk.LEFT, padx=(8, 0))

        footer = tk.Frame(self.window, bg="#eaecf0", height=60)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)

        tk.Button(
            footer, text="キャンセル",
            font=(_FONT, 10), relief=tk.FLAT,
            bg="#eaecf0", fg=_TEXT, activebackground="#dde1e7",
            cursor="hand2", bd=0, padx=20, pady=8,
            command=self.window.destroy
        ).pack(side=tk.RIGHT, padx=(6, 20), pady=12)

        tk.Button(
            footer, text="  保存  ",
            font=(_FONT, 10, "bold"), relief=tk.FLAT,
            bg=_PRIMARY, fg="white", activebackground=_PRIMARY_H,
            cursor="hand2", bd=0, padx=20, pady=8,
            command=self._on_save
        ).pack(side=tk.RIGHT, padx=6, pady=12)

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
            messagebox.showwarning(
                "確認",
                "バックアップ先を空のまま保存すると、自動バックアップは行われません。\n"
                "あとから同じ画面で保存先を指定できます。",
                parent=self.window
            )
        try:
            save_backup_settings(self.farm_path, interval, dest)
            messagebox.showinfo("保存完了", "バックアップ設定を保存しました。", parent=self.window)
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}", parent=self.window)

    def show(self):
        self.window.wait_window()
