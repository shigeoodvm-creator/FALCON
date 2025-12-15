"""
FALCON2 - 項目辞書一覧ウィンドウ
item_dictionary.json を読み書きする
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, Callable

import tkinter as tk
from tkinter import ttk, messagebox

logger = logging.getLogger(__name__)


def _get_origin(data: Dict[str, Any]) -> str:
    """origin/type のフォールバック取得"""
    return data.get("origin") or data.get("type") or ""


def _get_label(data: Dict[str, Any], key: str) -> str:
    """表示名のフォールバック取得"""
    return (
        data.get("label")
        or data.get("display_name")
        or data.get("name_jp")
        or key
    )


def _get_datatype(data: Dict[str, Any]) -> str:
    """データ型のフォールバック取得"""
    return data.get("datatype") or data.get("data_type") or "-"


def _is_editable(data: Dict[str, Any]) -> bool:
    """編集可否を取得（デフォルト True）"""
    return data.get("editable", True)


class ItemDictionaryWindow:
    """項目辞書一覧ウィンドウ"""

    ORIGIN_COLORS = {
        "core": "#1e90ff",
        "calc": "#cc6600",
        "event": "#2e8b57",
        "custom": "#555555",
    }

    def __init__(
        self,
        parent: tk.Tk,
        item_dictionary_path: Optional[Path] = None,
        on_item_updated: Optional[Callable[[], None]] = None,
        formula_engine: Optional[Any] = None,
    ):
        """
        初期化

        Args:
            parent: 親ウィンドウ
            item_dictionary_path: item_dictionary.json のパス
            on_item_updated: 保存後に呼び出されるコールバック（FormulaEngine再読込等）
            formula_engine: FormulaEngine インスタンス（任意）
        """
        self.parent = parent
        self.item_dict_path = item_dictionary_path
        self.on_item_updated = on_item_updated
        self.formula_engine = formula_engine

        self.item_dictionary: Dict[str, Any] = {}
        self.selected_item_key: Optional[str] = None

        self.window = tk.Toplevel(parent)
        self.window.title("項目辞書")
        self.window.geometry("900x640")
        self.window.transient(parent)

        self._load_item_dictionary()
        self._create_widgets()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _load_item_dictionary(self):
        """item_dictionary.json を読み込む"""
        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, "r", encoding="utf-8") as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                logger.error(f"item_dictionary.json 読み込みエラー: {e}")
                self.item_dictionary = {}
        else:
            messagebox.showwarning(
                "警告", f"item_dictionary.json が見つかりません: {self.item_dict_path}"
            )
            self.item_dictionary = {}

    def _create_widgets(self):
        """ウィジェットを作成"""
        title_label = ttk.Label(self.window, text="項目辞書一覧", font=("", 14, "bold"))
        title_label.pack(pady=10)

        desc_label = ttk.Label(
            self.window,
            text="item_dictionary.json の内容を表示・編集します",
            font=("", 9),
        )
        desc_label.pack(pady=(0, 10))

        table_frame = ttk.Frame(self.window)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("key", "label", "origin", "datatype", "editable")
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", height=20
        )

        self.tree.heading("key", text="項目キー")
        self.tree.heading("label", text="表示名")
        self.tree.heading("origin", text="origin")
        self.tree.heading("datatype", text="データ型")
        self.tree.heading("editable", text="編集可否")

        self.tree.column("key", width=180, anchor=tk.W)
        self.tree.column("label", width=200, anchor=tk.W)
        self.tree.column("origin", width=80, anchor=tk.CENTER)
        self.tree.column("datatype", width=100, anchor=tk.CENTER)
        self.tree.column("editable", width=80, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self._populate_tree()

        detail_frame = ttk.LabelFrame(self.window, text="詳細情報", padding=10)
        detail_frame.pack(fill=tk.X, padx=10, pady=5)

        self.detail_text = tk.Text(detail_frame, height=7, wrap=tk.WORD, state=tk.DISABLED)
        self.detail_text.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(label="項目を編集", command=self._on_edit_selected)
        self.context_menu.add_command(label="計算式を確認", command=self._on_formula_selected)
        self.context_menu.add_command(label="項目を削除", command=self._on_delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="新規項目を追加", command=self._on_add_new)

        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-Button-1>", lambda e: self._on_edit_selected())

        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=10)

        add_btn = ttk.Button(
            button_frame, text="新規項目を追加", command=self._on_add_new, width=18
        )
        add_btn.pack(side=tk.LEFT, padx=5)

        close_button = ttk.Button(
            button_frame, text="閉じる", command=self.window.destroy, width=12
        )
        close_button.pack(side=tk.LEFT, padx=5)

        debug_frame = ttk.LabelFrame(self.window, text="デバッグ情報", padding=5)
        debug_frame.pack(fill=tk.X, padx=10, pady=5)

        dict_path_str = str(self.item_dict_path) if self.item_dict_path else "(未指定)"
        ttk.Label(
            debug_frame,
            text=f"読み込み元: {dict_path_str}",
            font=("", 8),
            foreground="gray",
        ).pack(anchor=tk.W, padx=5, pady=2)

        count_label = ttk.Label(
            debug_frame,
            text=f"項目数: {len(self.item_dictionary)}",
            font=("", 8),
            foreground="gray",
        )
        count_label.pack(anchor=tk.W, padx=5, pady=2)

    def _populate_tree(self):
        """ツリービューにデータを挿入"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        sorted_items = sorted(self.item_dictionary.items(), key=lambda x: x[0])
        for key, data in sorted_items:
            origin = _get_origin(data)
            label = _get_label(data, key)
            datatype = _get_datatype(data)
            editable = "可" if _is_editable(data) else "不可"
            tags = (origin,)
            self.tree.insert("", tk.END, values=(key, label, origin, datatype, editable), tags=tags)
            self._ensure_origin_tag(origin)

    def _ensure_origin_tag(self, origin: str):
        """originごとの色タグ設定"""
        if not origin:
            return
        color = self.ORIGIN_COLORS.get(origin)
        if color:
            try:
                self.tree.tag_configure(origin, foreground=color)
            except tk.TclError:
                pass

    def _on_select(self, event=None):
        selection = self.tree.selection()
        if not selection:
            self.selected_item_key = None
            return
        item = self.tree.item(selection[0])
        values = item.get("values", [])
        if not values:
            return
        key = str(values[0])
        self.selected_item_key = key
        data = self.item_dictionary.get(key)
        if data:
            self._show_detail(key, data)

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
            item = self.tree.item(row_id)
            values = item.get("values", [])
            self.selected_item_key = str(values[0]) if values else None
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _on_edit_selected(self):
        key, data = self._get_selected_item()
        if not key or not data:
            messagebox.showwarning("警告", "項目を選択してください")
            return
        if not self.item_dict_path:
            messagebox.showerror("エラー", "item_dictionary.json のパスが設定されていません")
            return
        from ui.item_editor_window import ItemEditorWindow

        def on_saved():
            self._reload_dictionary()
        ItemEditorWindow(
            parent=self.window,
            item_key=key,
            item_data=data,
            item_dictionary_path=self.item_dict_path,
            on_saved=on_saved,
        ).show()

    def _on_formula_selected(self):
        key, data = self._get_selected_item()
        if not key or not data:
            messagebox.showwarning("警告", "項目を選択してください")
            return
        if not self.item_dict_path:
            messagebox.showerror("エラー", "item_dictionary.json のパスが設定されていません")
            return
        from ui.item_formula_window import ItemFormulaWindow

        def on_saved():
            self._reload_dictionary()
        ItemFormulaWindow(
            parent=self.window,
            item_key=key,
            item_data=data,
            item_dictionary_path=self.item_dict_path,
            on_saved=on_saved,
        ).show()

    def _on_delete_selected(self):
        key, data = self._get_selected_item()
        if not key or not data:
            messagebox.showwarning("警告", "項目を選択してください")
            return
        origin = _get_origin(data)
        if origin in ("core", "calc", "event"):
            messagebox.showwarning("警告", "core/event/calc 項目は削除できません")
            return
        result = messagebox.askyesno(
            "確認", f"項目「{_get_label(data, key)} ({key})」を削除しますか？"
        )
        if not result:
            return
        try:
            with open(self.item_dict_path, "r", encoding="utf-8") as f:
                item_dict = json.load(f)
            if key in item_dict:
                del item_dict[key]
                with open(self.item_dict_path, "w", encoding="utf-8") as f:
                    json.dump(item_dict, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("完了", "項目を削除しました")
                self._reload_dictionary()
            else:
                messagebox.showerror("エラー", f"{key} が見つかりません")
        except Exception as e:
            messagebox.showerror("エラー", f"削除に失敗しました: {e}")

    def _on_add_new(self):
        if not self.item_dict_path:
            messagebox.showerror("エラー", "item_dictionary.json のパスが設定されていません")
            return
        from ui.item_new_window import ItemNewWindow

        def on_saved():
            self._reload_dictionary()
        ItemNewWindow(
            parent=self.window,
            item_dictionary_path=self.item_dict_path,
            on_saved=on_saved,
        ).show()

    def _get_selected_item(self):
        selection = self.tree.selection()
        if not selection:
            return None, None
        item = self.tree.item(selection[0])
        values = item.get("values", [])
        if not values:
            return None, None
        key = str(values[0])
        data = self.item_dictionary.get(key)
        return key, data

    def _show_detail(self, key: str, data: Dict[str, Any]):
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete(1.0, tk.END)
        lines = [
            f"項目キー: {key}",
            f"表示名: {_get_label(data, key)}",
            f"origin: {_get_origin(data)}",
            f"データ型: {_get_datatype(data)}",
            f"編集可否: {'可' if _is_editable(data) else '不可'}",
        ]
        if data.get("description"):
            lines.append(f"説明: {data.get('description')}")
        if data.get("formula"):
            lines.append(f"計算式: {data.get('formula')}")
        if data.get("source"):
            lines.append(f"source: {data.get('source')}")
        self.detail_text.insert(1.0, "\n".join(lines))
        self.detail_text.config(state=tk.DISABLED)

    def _reload_dictionary(self):
        """辞書を再読込して表示更新"""
        prev_selected = self.selected_item_key
        self._load_item_dictionary()
        self._populate_tree()
        # 再選択
        if prev_selected:
            for iid in self.tree.get_children():
                item = self.tree.item(iid)
                values = item.get("values", [])
                if values and str(values[0]) == prev_selected:
                    self.tree.selection_set(iid)
                    self.tree.focus(iid)
                    break
        if self.on_item_updated:
            try:
                self.on_item_updated()
            except Exception as e:
                logger.error(f"on_item_updated callback error: {e}")
        # FormulaEngine 内部の辞書キャッシュもリロード
        if self.formula_engine:
            reload_method = getattr(self.formula_engine, "reload_item_dictionary", None)
            if callable(reload_method):
                try:
                    reload_method()
                except Exception as e:
                    logger.error(f"FormulaEngine reload failed: {e}")

    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()
        self.window.grab_set()


