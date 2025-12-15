"""
FALCON2 - 新規項目追加ウィンドウ
"""

import json
import logging
import re
from pathlib import Path
from typing import Optional, Callable

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

logger = logging.getLogger(__name__)


class ItemNewWindow:
    """新規項目追加ウィンドウ"""

    KEY_PATTERN = re.compile(r"^[A-Z0-9_]+$")

    def __init__(
        self,
        parent: tk.Tk,
        item_dictionary_path: Path,
        on_saved: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.item_dictionary_path = item_dictionary_path
        self.on_saved = on_saved

        self.window = tk.Toplevel(parent)
        self.window.title("新規項目を追加")
        self.window.geometry("640x560")
        self.window.transient(parent)
        self.window.grab_set()

        self._create_widgets()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        form.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(form, text="項目キー（英大文字・アンダースコア）:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.key_entry = ttk.Entry(form, width=32)
        self.key_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="表示名:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.label_entry = ttk.Entry(form, width=32)
        self.label_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(form, text="origin:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.origin_combo = ttk.Combobox(
            form,
            values=["custom", "calc"],
            width=29,
            state="readonly",
        )
        self.origin_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.origin_combo.set("custom")
        self.origin_combo.bind("<<ComboboxSelected>>", lambda e: self._toggle_formula())

        ttk.Label(form, text="データ型:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.datatype_combo = ttk.Combobox(
            form,
            values=["str", "int", "float", "bool", "date", "datetime"],
            width=29,
            state="readonly",
        )
        self.datatype_combo.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        self.datatype_combo.set("str")

        self.editable_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text="編集可能", variable=self.editable_var).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5
        )

        desc_frame = ttk.LabelFrame(main_frame, text="説明", padding=10)
        desc_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.description_text = scrolledtext.ScrolledText(desc_frame, height=6)
        self.description_text.pack(fill=tk.BOTH, expand=True)

        self.formula_frame = ttk.LabelFrame(main_frame, text="計算式（calc の場合は任意）", padding=10)
        self.formula_text = scrolledtext.ScrolledText(self.formula_frame, height=6)
        self.formula_text.pack(fill=tk.BOTH, expand=True)
        self._toggle_formula()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=8)
        ttk.Button(btn_frame, text="保存", command=self._on_save, width=12).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(btn_frame, text="キャンセル", command=self.window.destroy, width=12).pack(
            side=tk.LEFT, padx=5
        )

    def _toggle_formula(self):
        origin = self.origin_combo.get()
        if origin == "calc":
            self.formula_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        else:
            self.formula_frame.forget()

    def _on_save(self):
        key = self.key_entry.get().strip()
        label = self.label_entry.get().strip()
        origin = self.origin_combo.get() or "custom"
        datatype = self.datatype_combo.get()
        editable = self.editable_var.get()
        description = self.description_text.get(1.0, tk.END).strip()
        formula = self.formula_text.get(1.0, tk.END).strip() if origin == "calc" else ""

        if not key or not self.KEY_PATTERN.match(key):
            messagebox.showerror("エラー", "項目キーは英大文字・数字・アンダースコアのみ使用できます。")
            return
        if not label:
            messagebox.showerror("エラー", "表示名を入力してください")
            return

        try:
            with open(self.item_dictionary_path, "r", encoding="utf-8") as f:
                item_dict = json.load(f)
        except FileNotFoundError:
            item_dict = {}
        except Exception as e:
            messagebox.showerror("エラー", f"item_dictionary.json の読み込みに失敗しました: {e}")
            return

        if key in item_dict:
            messagebox.showerror("エラー", f"{key} は既に存在します")
            return

        new_entry = {
            "origin": origin,
            "type": "calc" if origin == "calc" else "custom",
            "label": label,
            "display_name": label,
            "datatype": datatype,
            "editable": editable,
            "visible": True,
        }
        if description:
            new_entry["description"] = description
        if origin == "calc" and formula:
            new_entry["formula"] = formula

        item_dict[key] = new_entry

        try:
            with open(self.item_dictionary_path, "w", encoding="utf-8") as f:
                json.dump(item_dict, f, ensure_ascii=False, indent=2)
            logger.info(f"item_dictionary added: {key}")
            messagebox.showinfo("完了", "項目を追加しました")
            if self.on_saved:
                self.on_saved()
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    def show(self):
        self.window.wait_window()


