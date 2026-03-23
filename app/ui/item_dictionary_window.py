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


def _get_category(data: Dict[str, Any]) -> str:
    """カテゴリーを取得"""
    return data.get("category") or "-"


class ItemDictionaryWindow:
    """項目辞書一覧ウィンドウ"""

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
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *args: self._on_search_changed())
        
        # ソート状態管理（列名 -> ソート状態: None=元の順序, 'asc'=昇順, 'desc'=降順）
        self.sort_states: Dict[str, Optional[str]] = {
            "key": None,
            "label": None,
            "category": None,
            "origin": None,
            "datatype": None
        }
        # 元の順序を保持（項目キーの順序）
        self.original_order: list = []

        self.window = tk.Toplevel(parent)
        self.window.title("項目辞書")
        self.window.geometry("900x640")
        self.window.configure(bg="#f5f5f5")

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
                # 元の順序を保持（項目キーの順序）
                sorted_items = sorted(self.item_dictionary.items(), key=lambda x: x[0])
                self.original_order = [key for key, _ in sorted_items]
                logger.info(f"[ItemDictionaryWindow] 読み込み成功: {self.item_dict_path}")
                logger.info(f"[ItemDictionaryWindow] 読み込んだ項目数: {len(self.item_dictionary)}")
                logger.info(f"[ItemDictionaryWindow] 項目キー: {list(self.item_dictionary.keys())}")
                
                # TWIN項目が存在しない場合は追加
                if "TWIN" not in self.item_dictionary:
                    logger.info("[ItemDictionaryWindow] TWIN項目が見つからないため追加します")
                    self.item_dictionary["TWIN"] = {
                        "type": "calc",
                        "origin": "calc",
                        "data_type": "str",
                        "formula": "twin_flag(events, cow)",
                        "display_name": "双子フラグ",
                        "description": "略称: TWIN（妊娠イベントで双子とされた個体で「双子」と表示。妊娠していない個体は未受胎と表示、妊娠しているかつ双子とされていない場合は空欄）",
                        "category": "REPRODUCTION"
                    }
                    # ファイルに保存（フォルダが存在しない場合は作成）
                    try:
                        if self.item_dict_path:
                            self.item_dict_path.parent.mkdir(parents=True, exist_ok=True)
                        with open(self.item_dict_path, "w", encoding="utf-8") as f:
                            json.dump(self.item_dictionary, f, ensure_ascii=False, indent=2)
                        logger.info("[ItemDictionaryWindow] TWIN項目を追加して保存しました")
                    except Exception as e:
                        logger.error(f"[ItemDictionaryWindow] TWIN項目の保存に失敗: {e}")
            except Exception as e:
                logger.error(f"item_dictionary.json 読み込みエラー: {e}", exc_info=True)
                self.item_dictionary = {}
                self.original_order = []
        else:
            logger.warning(f"[ItemDictionaryWindow] item_dictionary.json が見つかりません: {self.item_dict_path}")
            messagebox.showwarning(
                "警告", f"item_dictionary.json が見つかりません: {self.item_dict_path}"
            )
            self.item_dictionary = {}
            self.original_order = []

    def _create_widgets(self):
        """ウィジェットを作成（他ウィンドウと同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"

        self.main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.main_container.grid_rowconfigure(2, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # ヘッダー（他ウィンドウと同一イメージ）
        header = tk.Frame(self.main_container, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\u2695", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="項目辞書一覧", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="item_dictionary.json の内容を表示します", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)

        # 検索欄
        search_frame = ttk.Frame(self.main_container)
        search_frame.grid(row=1, column=0, sticky=tk.EW, pady=(0, 10))
        search_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="検索:", font=(_df, 9)).grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(search_frame, text="クリア", command=self._clear_search, width=8).grid(row=0, column=2, padx=(5, 0))

        table_frame = ttk.Frame(self.main_container)
        table_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        columns = ("key", "label", "category", "origin", "datatype")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)

        self.tree.heading("key", text="項目キー", command=lambda: self._on_column_click("key"))
        self.tree.heading("label", text="表示名", command=lambda: self._on_column_click("label"))
        self.tree.heading("category", text="カテゴリー", command=lambda: self._on_column_click("category"))
        self.tree.heading("origin", text="origin", command=lambda: self._on_column_click("origin"))
        self.tree.heading("datatype", text="データ型", command=lambda: self._on_column_click("datatype"))

        self.tree.column("key", width=150, anchor=tk.W)
        self.tree.column("label", width=240, anchor=tk.W)
        self.tree.column("category", width=120, anchor=tk.W)
        self.tree.column("origin", width=80, anchor=tk.CENTER)
        self.tree.column("datatype", width=100, anchor=tk.CENTER)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self._populate_tree()

        detail_frame = ttk.LabelFrame(self.main_container, text="詳細情報", padding=10)
        detail_frame.grid(row=3, column=0, sticky=tk.EW, pady=(0, 10))
        detail_frame.grid_columnconfigure(0, weight=1)

        self.detail_text = tk.Text(detail_frame, height=7, wrap=tk.WORD, state=tk.DISABLED)
        self.detail_text.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(label="項目を編集", command=self._on_edit_selected)
        self.context_menu.add_command(label="計算式を確認", command=self._on_formula_selected)
        self.context_menu.add_command(label="項目を削除", command=self._on_delete_selected)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-Button-1>", lambda e: self._on_edit_selected())

        button_frame = ttk.Frame(self.main_container)
        button_frame.grid(row=4, column=0, pady=(0, 10))
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy, width=12).pack(side=tk.LEFT, padx=5)

        debug_frame = ttk.LabelFrame(self.main_container, text="デバッグ情報", padding=5)
        debug_frame.grid(row=5, column=0, sticky=tk.EW, pady=(0, 5))
        debug_frame.grid_columnconfigure(0, weight=1)
        dict_path_str = str(self.item_dict_path) if self.item_dict_path else "(未指定)"
        ttk.Label(debug_frame, text=f"読み込み元: {dict_path_str}", font=("", 8), foreground="gray").pack(anchor=tk.W, padx=5, pady=2)
        self._debug_count_label = ttk.Label(debug_frame, text=f"項目数: {len(self.item_dictionary)}", font=("", 8), foreground="gray")
        self._debug_count_label.pack(anchor=tk.W, padx=5, pady=2)
    
    def _on_column_click(self, column: str):
        """列ヘッダーがクリックされたときの処理（3段階ソート：昇順→降順→元の順序）"""
        current_state = self.sort_states.get(column, None)
        
        # 他の列のソート状態をリセット
        for col in self.sort_states:
            if col != column:
                self.sort_states[col] = None
        
        # ソート状態を切り替え：None → 'asc' → 'desc' → None
        if current_state is None:
            # 1クリック目：昇順
            self.sort_states[column] = 'asc'
            self._sort_tree(column, 'asc')
        elif current_state == 'asc':
            # 2クリック目：降順
            self.sort_states[column] = 'desc'
            self._sort_tree(column, 'desc')
        else:
            # 3クリック目：元の順序に戻す
            self.sort_states[column] = None
            self._sort_tree(column, None)
    
    def _sort_tree(self, column: str, sort_order: Optional[str]):
        """Treeviewをソート（_populate_treeを呼び出して再構築）"""
        # ソート状態を更新した後、_populate_treeを呼び出して再構築
        self._populate_tree()
    
    def _update_column_headers(self):
        """列ヘッダーにソートインジケーターを表示"""
        headers = {
            "key": "項目キー",
            "label": "表示名",
            "category": "カテゴリー",
            "origin": "origin",
            "datatype": "データ型"
        }
        
        for col, base_text in headers.items():
            state = self.sort_states.get(col, None)
            if state == 'asc':
                text = f"{base_text} ▲"
            elif state == 'desc':
                text = f"{base_text} ▼"
            else:
                text = base_text
            
            self.tree.heading(col, text=text, command=lambda c=col: self._on_column_click(c))

        if hasattr(self, "_debug_count_label") and self._debug_count_label.winfo_exists():
            self._debug_count_label.config(text=f"項目数: {len(self.item_dictionary)}")

    def _populate_tree(self):
        """ツリービューにデータを挿入（検索フィルタリングとソートを適用）"""
        # 既存の項目をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 検索フィルタリング
        search_text = self.search_var.get().strip().lower()
        
        # 全項目を取得
        all_items = list(self.item_dictionary.items())
        
        # 検索フィルタリング
        filtered_items = []
        # 出生年・出生年月：alias が辞書に無くても検索でヒットするようフォールバック
        _VIRTUAL_ALIAS = {"BTHYR": "出生年", "BTHYM": "出生年月"}
        for key, data in all_items:
            origin = _get_origin(data)
            label = _get_label(data, key)
            category = _get_category(data)
            alias = (data.get("alias") or "").strip().lower()
            virtual = (_VIRTUAL_ALIAS.get(key) or "").strip().lower()
            search_lower = search_text.lower() if search_text else ""
            match_alias = alias and search_lower and search_lower in alias
            match_virtual = virtual and search_lower and search_lower in virtual
            
            # 検索フィルタリング（項目キー・表示名・カテゴリー・別名(alias)・仮の別名）
            if search_text:
                if (search_text not in key.lower() and
                    search_text not in label.lower() and
                    search_text not in category.lower() and
                    not match_alias and
                    not match_virtual):
                    continue
            
            filtered_items.append((key, data))
        
        # ソート状態を確認
        active_sort_column = None
        active_sort_order = None
        for col, state in self.sort_states.items():
            if state is not None:
                active_sort_column = col
                active_sort_order = state
                break
        
        # ソート適用
        if active_sort_column and active_sort_order:
            # ソート用のキー関数を定義
            column_index_map = {
                "key": 0,
                "label": 1,
                "category": 2,
                "origin": 3,
                "datatype": 4
            }
            col_idx = column_index_map.get(active_sort_column, 0)
            
            def get_sort_key(item):
                key, data = item
                origin = _get_origin(data)
                label = _get_label(data, key)
                category = _get_category(data)
                datatype = _get_datatype(data)
                
                values = (key, label, category, origin, datatype)
                value = values[col_idx] if col_idx < len(values) else ""
                
                # 数値として比較できる場合は数値に変換
                try:
                    return (0, value) if isinstance(value, (int, float)) else (1, str(value).lower())
                except (ValueError, TypeError):
                    return (1, str(value).lower())
            
            reverse = (active_sort_order == 'desc')
            filtered_items = sorted(filtered_items, key=get_sort_key, reverse=reverse)
        elif active_sort_column is None:
            # 元の順序に戻す（項目キー順）
            filtered_items = sorted(filtered_items, key=lambda x: self.original_order.index(x[0]) if x[0] in self.original_order else len(self.original_order))
        
        # Treeviewに表示
        filtered_count = 0
        for key, data in filtered_items:
            origin = _get_origin(data)
            label = _get_label(data, key)
            category = _get_category(data)
            datatype = _get_datatype(data)
            
            self.tree.insert("", tk.END, values=(key, label, category, origin, datatype))
            filtered_count += 1
        
        # ヘッダーにソートインジケーターを表示
        self._update_column_headers()
        
        logger.info(f"[ItemDictionaryWindow] 表示した項目数: {filtered_count}")

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
            # normalization辞書を再生成
            self._regenerate_normalization_dict()
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
            # normalization辞書を再生成
            self._regenerate_normalization_dict()
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
                # フォルダが存在しない場合は作成
                if self.item_dict_path:
                    self.item_dict_path.parent.mkdir(parents=True, exist_ok=True)
                with open(self.item_dict_path, "w", encoding="utf-8") as f:
                    json.dump(item_dict, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("完了", "項目を削除しました")
                self._reload_dictionary()
                # normalization辞書を再生成
                self._regenerate_normalization_dict()
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
            # normalization辞書を再生成
            self._regenerate_normalization_dict()
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
            f"カテゴリー: {_get_category(data)}",
            f"origin: {_get_origin(data)}",
            f"データ型: {_get_datatype(data)}",
        ]
        if data.get("description"):
            lines.append(f"説明: {data.get('description')}")
        if data.get("formula"):
            lines.append(f"計算式: {data.get('formula')}")
        if data.get("source"):
            lines.append(f"source: {data.get('source')}")
        self.detail_text.insert(1.0, "\n".join(lines))
        self.detail_text.config(state=tk.DISABLED)

    def _on_search_changed(self):
        """検索欄の変更時にフィルタリングを実行"""
        self._populate_tree()
        # 選択をクリア
        self.tree.selection_remove(self.tree.selection())

    def _clear_search(self):
        """検索欄をクリア"""
        self.search_var.set("")

    def _regenerate_normalization_dict(self):
        """normalization辞書を再生成"""
        try:
            from modules.normalization_generator import generate_normalization_dict
            generate_normalization_dict()
            logger.info("[ItemDictionaryWindow] normalization辞書を再生成しました")
        except Exception as e:
            logger.error(f"[ItemDictionaryWindow] normalization辞書の再生成に失敗: {e}")
    
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


