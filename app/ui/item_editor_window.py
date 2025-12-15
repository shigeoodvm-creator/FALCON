"""
FALCON2 - 項目編集ウィンドウ
item_dictionary.json の単一項目を編集する
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, Callable

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

logger = logging.getLogger(__name__)


class ItemEditorWindow:
    """項目編集ウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        item_key: str,
        item_data: Dict[str, Any],
        item_dictionary_path: Path,
        on_saved: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.item_key = item_key
        self.original_data = item_data.copy()
        self.item_dictionary_path = item_dictionary_path
        self.on_saved = on_saved

        self.window = tk.Toplevel(parent)
        self.window.title(f"項目を編集 - {item_key}")
        self.window.geometry("640x520")
        self.window.transient(parent)
        self.window.grab_set()

        self.origin = item_data.get("origin") or item_data.get("type") or "custom"
        self._create_widgets()
        self._populate_fields()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(info_frame, text="項目キー:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(info_frame, text=self.item_key, font=("", 10, "bold")).grid(
            row=0, column=1, sticky=tk.W, padx=5, pady=5
        )

        ttk.Label(info_frame, text="表示名:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.label_entry = ttk.Entry(info_frame, width=30)
        self.label_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(info_frame, text="データ型:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.datatype_combo = ttk.Combobox(
            info_frame,
            values=["str", "int", "float", "bool", "date", "datetime"],
            width=27,
            state="readonly",
        )
        self.datatype_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        self.editable_var = tk.BooleanVar(value=True)
        self.editable_check = ttk.Checkbutton(
            info_frame, text="編集可能", variable=self.editable_var
        )
        self.editable_check.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(info_frame, text="origin:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(info_frame, text=self.origin).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)

        # 計算式/ソース（calc/event項目のみ）
        self.formula_frame = ttk.LabelFrame(main_frame, text="計算式 / ソース", padding=10)
        self.formula_text = tk.Text(self.formula_frame, height=4, width=60)

        desc_frame = ttk.LabelFrame(main_frame, text="説明", padding=10)
        desc_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.description_text = scrolledtext.ScrolledText(desc_frame, height=6)
        self.description_text.pack(fill=tk.BOTH, expand=True)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="保存", command=self._on_save, width=12).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(button_frame, text="キャンセル", command=self.window.destroy, width=12).pack(
            side=tk.LEFT, padx=5
        )

    def _populate_fields(self):
        label = (
            self.original_data.get("label")
            or self.original_data.get("display_name")
            or self.original_data.get("name_jp")
            or self.item_key
        )
        self.label_entry.insert(0, label)

        datatype = (
            self.original_data.get("datatype")
            or self.original_data.get("data_type")
            or ""
        )
        if datatype:
            self.datatype_combo.set(datatype)
        else:
            self.datatype_combo.set("str")

        self.editable_var.set(self.original_data.get("editable", True))

        desc = self.original_data.get("description") or ""
        self.description_text.insert(1.0, desc)

        # 計算式/ソースの扱い
        if self.origin in ("calc", "event"):
            self.formula_frame.pack(fill=tk.X, pady=(0, 10))
            self.formula_text.pack(fill=tk.X)
            formula_value = self.original_data.get("formula") or self.original_data.get("source") or ""
            self.formula_text.insert(1.0, formula_value)
            if self.origin == "event":
                self.formula_text.config(state=tk.DISABLED)
        else:
            # custom/core では非表示
            self.formula_frame.forget()

        # core/event/calc は編集可能範囲を制限（editableだけ変更可とする）
        if self.origin in ("core", "event", "calc"):
            self.label_entry.config(state="normal")
            # label変更は許可する（表示名だけ微調整可能）
            # データ型は選択可能にする（calc項目などで必要）
            self.datatype_combo.config(state="readonly")

    def _on_save(self):
        label = self.label_entry.get().strip()
        datatype = self.datatype_combo.get()
        editable = self.editable_var.get()
        description = self.description_text.get(1.0, tk.END).strip()

        if not label:
            messagebox.showerror("エラー", "表示名を入力してください")
            return

        if not datatype:
            messagebox.showerror("エラー", "データ型を選択してください")
            return

        try:
            with open(self.item_dictionary_path, "r", encoding="utf-8") as f:
                item_dict = json.load(f)
        except Exception as e:
            messagebox.showerror("エラー", f"item_dictionary.json の読み込みに失敗しました: {e}")
            return

        # 既存データをマージ
        current = item_dict.get(self.item_key, {}).copy()
        current["label"] = label
        current["display_name"] = label
        current["datatype"] = datatype
        current["data_type"] = datatype  # data_type も保存（互換性のため）
        current["editable"] = editable
        if description:
            current["description"] = description
        elif "description" in current:
            del current["description"]

        current["origin"] = current.get("origin") or current.get("type") or self.origin
        current["type"] = current.get("type") or current["origin"]

        if self.origin == "calc":
            formula = self.formula_text.get(1.0, tk.END).strip() if self.formula_text.winfo_exists() else ""
            if formula:
                current["formula"] = formula
            elif "formula" in current:
                del current["formula"]
        elif self.origin == "event":
            # 参照のみ。編集不可のまま保持
            pass

        item_dict[self.item_key] = current

        try:
            with open(self.item_dictionary_path, "w", encoding="utf-8") as f:
                json.dump(item_dict, f, ensure_ascii=False, indent=2)
            logger.info(f"item_dictionary updated: {self.item_key}")
            messagebox.showinfo("完了", "項目を保存しました")
            if self.on_saved:
                self.on_saved()
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    def show(self):
        self.window.wait_window()


