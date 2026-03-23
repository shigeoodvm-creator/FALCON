"""
FALCON - コマンド辞書ウィンドウ
command_dictionary.json を編集する（登録名 → 展開コマンド）
"""

import json
import logging
from pathlib import Path
from typing import Dict, Optional, Callable

import tkinter as tk
from tkinter import ttk, messagebox

logger = logging.getLogger(__name__)


class CommandDictionaryWindow:
    """コマンド辞書一覧・編集ウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        command_dictionary_path: Optional[Path] = None,
        on_updated: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.command_dict_path = command_dictionary_path
        self.on_updated = on_updated
        self.command_dictionary: Dict[str, str] = {}
        self.selected_name: Optional[str] = None

        self.window = tk.Toplevel(parent)
        self.window.title("コマンド辞書")
        self.window.geometry("820x480")
        self.window.configure(bg="#f5f5f5")

        self._load_dictionary()
        self._create_widgets()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _load_dictionary(self):
        """command_dictionary.json を読み込む"""
        if self.command_dict_path and self.command_dict_path.exists():
            try:
                with open(self.command_dict_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.command_dictionary = {str(k): str(v) for k, v in (data or {}).items()}
            except Exception as e:
                logger.error(f"command_dictionary.json 読み込みエラー: {e}")
                self.command_dictionary = {}
        else:
            self.command_dictionary = {}

    def _save_dictionary(self) -> bool:
        """command_dictionary.json に保存"""
        if not self.command_dict_path:
            messagebox.showerror("エラー", "コマンド辞書のパスが設定されていません")
            return False
        try:
            self.command_dict_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.command_dict_path, "w", encoding="utf-8") as f:
                json.dump(self.command_dictionary, f, ensure_ascii=False, indent=2)
            if self.on_updated:
                self.on_updated()
            return True
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")
            return False

    def _create_widgets(self):
        _df = "Meiryo UI"
        bg = "#f5f5f5"

        main = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        main.pack(fill=tk.BOTH, expand=True)
        main.grid_rowconfigure(2, weight=1)
        main.grid_columnconfigure(0, weight=1)

        header = tk.Frame(main, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\u2699\ufe0f", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="コマンド辞書", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(
            title_frame,
            text="登録名をコマンド欄に入力すると、登録されたコマンドに展開されて実行されます",
            font=(_df, 10),
            bg=bg,
            fg="#607d8b",
        ).pack(anchor=tk.W)

        table_frame = ttk.Frame(main)
        table_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        columns = ("name", "command")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        self.tree.heading("name", text="登録名")
        self.tree.heading("command", text="展開コマンド")
        self.tree.column("name", width=180, anchor=tk.W)
        self.tree.column("command", width=580, anchor=tk.W)

        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        self._populate_tree()
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-Button-1>", lambda e: self._on_edit())

        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=4, column=0, pady=(0, 10))
        ttk.Button(btn_frame, text="使い方", command=self._show_usage, width=10).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="追加", command=self._on_add, width=10).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="編集", command=self._on_edit, width=10).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="削除", command=self._on_delete, width=10).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="閉じる", command=self.window.destroy, width=10).pack(side=tk.LEFT, padx=3)

        path_str = str(self.command_dict_path) if self.command_dict_path else "(未指定)"
        ttk.Label(main, text=f"読み込み元: {path_str}", font=("", 8), foreground="gray").grid(
            row=5, column=0, sticky=tk.W, pady=2
        )

    def _populate_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for name, command in self.command_dictionary.items():
            self.tree.insert("", tk.END, values=(name, command))

    def _on_select(self, event=None):
        sel = self.tree.selection()
        self.selected_name = None
        if sel:
            item = self.tree.item(sel[0])
            vals = item.get("values", [])
            if vals:
                self.selected_name = str(vals[0])

    def _show_usage(self):
        """使い方（リストコマンドの書式など）をダイアログで表示"""
        _usage_text = """【コマンド辞書の使い方】

登録名をコマンド欄に入力して実行すると、ここに登録した「展開コマンド」に置き換わって実行されます。
自作の登録名（例：分娩予定表、受胎牛一覧）でよく使うコマンドを呼び出せます。

■ リストコマンドの書式（自作するときの参考）

  リスト：項目1 項目2 項目3 … ：条件 ：並び順

  • 項目 … 表示したい列をスペース区切りで指定（ID, 受胎日, 分娩予定日 など）
  • 条件 … 省略可。例：RC=5,6（受胎のみ）、RC=1,2,3,4（空胎のみ）、DIM>150
  • 並び順 … 省略可。指定する場合は「昇順 項目名」または「降順 項目名」

■ 並び順の指定（コロン「：」で区切って3番目に書く）

  • 昇順 項目名 … その項目の小さい順（古い日付順など）
  • 降順 項目名 … その項目の大きい順（新しい日付順など）

  例）分娩予定日で新しい順にしたい場合：
      リスト：ID 受胎日 分娩予定日 RC：RC=5,6：降順 分娩予定日

■ 注意
  • コロン「：」は半角・全角どちらでも可です。
  • 条件だけ付ける場合：リスト：ID 分娩予定日 RC：RC=5,6
  • 並び順だけ付ける場合：リスト：ID 分娩予定日：降順 分娩予定日
"""
        dlg = tk.Toplevel(self.window)
        dlg.title("コマンド辞書の使い方")
        dlg.geometry("560x420")
        dlg.configure(bg="#f5f5f5")
        f = tk.Frame(dlg, padx=12, pady=12, bg="#f5f5f5")
        f.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(f, wrap=tk.WORD, font=("Meiryo UI", 10), padx=8, pady=8, height=24, width=62)
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(tk.END, _usage_text.strip())
        txt.config(state=tk.DISABLED)
        ttk.Button(f, text="閉じる", command=dlg.destroy, width=10).pack(pady=(8, 0))

    def _on_add(self):
        self._show_edit_dialog(is_new=True)

    def _on_edit(self):
        if not self.selected_name:
            messagebox.showwarning("警告", "編集する行を選択してください")
            return
        self._show_edit_dialog(is_new=False, current_name=self.selected_name)

    def _show_edit_dialog(self, is_new: bool, current_name: str = ""):
        dlg = tk.Toplevel(self.window)
        dlg.title("追加" if is_new else "編集")
        dlg.geometry("520x200")

        f = ttk.Frame(dlg, padding=10)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="登録名:").grid(row=0, column=0, sticky=tk.W, pady=4)
        name_var = tk.StringVar(value=current_name)
        name_entry = ttk.Entry(f, textvariable=name_var, width=50)
        name_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=4)
        if not is_new:
            name_entry.config(state="readonly")  # 編集時は名前変更不可（キーなので）

        ttk.Label(f, text="展開コマンド:").grid(row=1, column=0, sticky=tk.NW, pady=4)
        cmd_var = tk.StringVar(value=self.command_dictionary.get(current_name, ""))
        cmd_entry = ttk.Entry(f, textvariable=cmd_var, width=60)
        cmd_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=4)

        f.grid_columnconfigure(1, weight=1)

        def save():
            name = name_var.get().strip()
            cmd = cmd_var.get().strip()
            if not name:
                messagebox.showwarning("警告", "登録名を入力してください", parent=dlg)
                return
            if not cmd:
                messagebox.showwarning("警告", "展開コマンドを入力してください", parent=dlg)
                return
            if is_new and name in self.command_dictionary:
                messagebox.showwarning("警告", f"登録名「{name}」は既に存在します", parent=dlg)
                return
            if is_new:
                self.command_dictionary[name] = cmd
            else:
                if current_name != name:
                    return
                self.command_dictionary[current_name] = cmd
            if self._save_dictionary():
                self._load_dictionary()
                self._populate_tree()
                dlg.destroy()
                messagebox.showinfo("完了", "保存しました", parent=self.window)

        btn_f = ttk.Frame(dlg)
        btn_f.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_f, text="保存", command=save, width=10).pack(side=tk.RIGHT, padx=3)
        ttk.Button(btn_f, text="キャンセル", command=dlg.destroy, width=10).pack(side=tk.RIGHT)

    def _on_delete(self):
        if not self.selected_name:
            messagebox.showwarning("警告", "削除する行を選択してください")
            return
        if not messagebox.askyesno("確認", f"「{self.selected_name}」を削除しますか？", parent=self.window):
            return
        if self.selected_name in self.command_dictionary:
            del self.command_dictionary[self.selected_name]
            if self._save_dictionary():
                self._load_dictionary()
                self._populate_tree()
                self.selected_name = None
                messagebox.showinfo("完了", "削除しました", parent=self.window)

    def show(self):
        self.window.focus_set()
        self.window.lift()
