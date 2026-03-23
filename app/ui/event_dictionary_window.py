"""
FALCON2 - イベント辞書一覧ウィンドウ
登録されているイベントの一覧を表示
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Dict, Any, Tuple
import json
import logging

logger = logging.getLogger(__name__)


class EventDictionaryWindow:
    """イベント辞書一覧ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            event_dictionary_path: event_dictionary.json のパス
        """
        self.parent = parent
        self.event_dict_path = event_dictionary_path
        self.event_dictionary: Dict[str, Any] = {}
        self.selected_item_id: Optional[str] = None  # 右クリックで選択されたアイテムID
        
        # 検索用
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *args: self._on_search_changed())
        
        # ソート状態管理（列名 -> ソート状態: None=元の順序, 'asc'=昇順, 'desc'=降順）
        self.sort_states: Dict[str, Optional[str]] = {
            "event_number": None,
            "alias": None,
            "name_jp": None,
            "category": None
        }
        # 元の順序を保持（イベント番号の順序）
        self.original_order: list = []
        
        # ウィンドウを作成
        self.window = tk.Toplevel(parent)
        self.window.title("イベント辞書")
        self.window.geometry("840x640")
        self.window.configure(bg="#f5f5f5")
        
        self._load_event_dictionary()
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _regenerate_normalization_dict(self):
        """normalization辞書を再生成"""
        try:
            from modules.normalization_generator import generate_normalization_dict
            generate_normalization_dict()
            logger.info("[EventDictionaryWindow] normalization辞書を再生成しました")
        except Exception as e:
            logger.error(f"[EventDictionaryWindow] normalization辞書の再生成に失敗: {e}")
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
                # 元の順序を保持（イベント番号の順序）
                sorted_events = sorted(
                    self.event_dictionary.items(),
                    key=lambda x: int(x[0]) if x[0].isdigit() else 9999
                )
                self.original_order = [key for key, _ in sorted_events]
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
                self.original_order = []
        else:
            # 農場フォルダ側の辞書が存在しない場合は空辞書
            # （同期処理で作成されるはずだが、念のため）
            print(f"警告: event_dictionary.json が見つかりません: {self.event_dict_path}")
            self.event_dictionary = {}
            self.original_order = []
    
    def _create_widgets(self):
        """ウィジェットを作成（項目辞書一覧と同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"

        # このウィンドウ用の ttk スタイル（Treeview フォント統一）
        try:
            style = ttk.Style(self.window)
            style.configure("Treeview", font=(_df, 10))
            style.configure("Treeview.Heading", font=(_df, 10, "bold"))
        except tk.TclError:
            pass

        self.main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.main_container.grid_rowconfigure(2, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # ヘッダー（アイコン・タイトル・サブタイトル）
        header = tk.Frame(self.main_container, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\U0001F4CB", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="イベント辞書一覧", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="登録されているイベントの一覧です", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)

        # 検索欄
        search_frame = ttk.Frame(self.main_container)
        search_frame.grid(row=1, column=0, sticky=tk.EW, pady=(0, 10))
        search_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(search_frame, text="検索:", font=(_df, 9)).grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(search_frame, text="クリア", command=self._clear_search, width=8).grid(row=0, column=2, padx=(5, 0))

        # テーブルフレーム
        table_frame = ttk.Frame(self.main_container)
        table_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # ツリービュー
        self.tree = ttk.Treeview(table_frame, columns=("event_number", "alias", "name_jp", "category"),
                                 show="headings", height=20)

        self.tree.heading("event_number", text="イベント番号", command=lambda: self._on_column_click("event_number"))
        self.tree.heading("alias", text="エイリアス", command=lambda: self._on_column_click("alias"))
        self.tree.heading("name_jp", text="日本語名", command=lambda: self._on_column_click("name_jp"))
        self.tree.heading("category", text="カテゴリ", command=lambda: self._on_column_click("category"))

        self.tree.column("event_number", width=120, anchor=tk.CENTER)
        self.tree.column("alias", width=150, anchor=tk.W)
        self.tree.column("name_jp", width=250, anchor=tk.W)
        self.tree.column("category", width=150, anchor=tk.W)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self._populate_tree()

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(label="編集", command=self._on_edit_selected)
        self.context_menu.add_command(label="削除", command=self._on_delete_selected)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-Button-1>", lambda e: self._on_edit_selected())

        # ボタンフレーム
        button_frame = ttk.Frame(self.main_container)
        button_frame.grid(row=3, column=0, pady=(0, 10))
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy, width=12).pack(side=tk.LEFT, padx=5)

        # デバッグ情報フレーム
        debug_frame = ttk.LabelFrame(self.main_container, text="デバッグ情報", padding=5)
        debug_frame.grid(row=4, column=0, sticky=tk.EW, pady=(0, 5))
        debug_frame.grid_columnconfigure(0, weight=1)
        dict_path_str = str(self.event_dict_path) if self.event_dict_path else "(未指定)"
        ttk.Label(debug_frame, text=f"読み込み元: {dict_path_str}", font=("", 8), foreground="gray").pack(anchor=tk.W, padx=5, pady=2)
        visible_count = sum(1 for d in self.event_dictionary.values() if not d.get("deprecated", False))
        total_count = len(self.event_dictionary)
        self._debug_count_label = ttk.Label(
            debug_frame,
            text=f"表示中イベント数: {visible_count} / 総イベント数: {total_count}",
            font=("", 8),
            foreground="gray"
        )
        self._debug_count_label.pack(anchor=tk.W, padx=5, pady=2)

    def _clear_search(self):
        """検索欄をクリア"""
        self.search_var.set("")

    def _on_search_changed(self):
        """検索テキスト変更時に一覧を再描画"""
        self._populate_tree()
    
    def _populate_tree(self):
        """ツリービューにデータを挿入（検索フィルタ・ソートを適用）"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        search_text = (self.search_var.get() or "").strip().lower()

        # 全イベントを取得（deprecatedを除く）
        all_events = []
        for event_number, event_data in self.event_dictionary.items():
            if event_data.get("deprecated", False):
                continue
            all_events.append((event_number, event_data))

        # 検索フィルタ
        if search_text:
            filtered = []
            for event_number, event_data in all_events:
                alias = (event_data.get("alias") or "").lower()
                name_jp = (event_data.get("name_jp") or "").lower()
                category = (event_data.get("category") or "").lower()
                if (search_text in event_number or
                    search_text in alias or
                    search_text in name_jp or
                    search_text in category):
                    filtered.append((event_number, event_data))
            all_events = filtered
        
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
                "event_number": 0,
                "alias": 1,
                "name_jp": 2,
                "category": 3
            }
            col_idx = column_index_map.get(active_sort_column, 0)
            
            def get_sort_key(item):
                event_number, event_data = item
                alias = event_data.get("alias", "")
                name_jp = event_data.get("name_jp", "")
                category = event_data.get("category", "")
                
                values = (event_number, alias, name_jp, category)
                value = values[col_idx] if col_idx < len(values) else ""
                
                # イベント番号の場合は数値として比較
                if active_sort_column == "event_number":
                    try:
                        return (0, int(value)) if value.isdigit() else (1, value)
                    except (ValueError, TypeError):
                        return (1, str(value))
                # その他の場合は文字列として比較
                return (1, str(value).lower())
            
            reverse = (active_sort_order == 'desc')
            all_events = sorted(all_events, key=get_sort_key, reverse=reverse)
        elif active_sort_column is None:
            # 元の順序に戻す（イベント番号順）
            all_events = sorted(all_events, key=lambda x: self.original_order.index(x[0]) if x[0] in self.original_order else len(self.original_order))
        
        # Treeviewに表示
        for event_number, event_data in all_events:
            alias = event_data.get("alias", "")
            name_jp = event_data.get("name_jp", "")
            category = event_data.get("category", "")
            self.tree.insert(
                "",
                tk.END,
                values=(event_number, alias, name_jp, category),
                tags=(event_number,)
            )

        self._update_column_headers()
        # デバッグ表示中の件数を更新
        if hasattr(self, "_debug_count_label") and self._debug_count_label.winfo_exists():
            total_count = len(self.event_dictionary)
            self._debug_count_label.config(
                text=f"表示中イベント数: {len(all_events)} / 総イベント数: {total_count}"
            )
    
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
            self._populate_tree()
        elif current_state == 'asc':
            # 2クリック目：降順
            self.sort_states[column] = 'desc'
            self._populate_tree()
        else:
            # 3クリック目：元の順序に戻す
            self.sort_states[column] = None
            self._populate_tree()
    
    def _update_column_headers(self):
        """列ヘッダーにソートインジケーターを表示"""
        headers = {
            "event_number": "イベント番号",
            "alias": "エイリアス",
            "name_jp": "日本語名",
            "category": "カテゴリ"
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
    
    def _on_select(self, event):
        """イベント選択時の処理（詳細情報は表示しない）"""
        selection = self.tree.selection()
        if not selection:
            self.selected_item_id = None
            return
        self.selected_item_id = selection[0]
    
    def _on_right_click(self, event):
        """右クリック時の処理（コンテキストメニューを表示）"""
        # event.y を使って Treeview.identify_row(event.y) を取得
        row_id = self.tree.identify_row(event.y)
        if row_id:
            # 該当行があれば、その行を selection_set する
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
            # 選択されたアイテムIDを保存
            self.selected_item_id = row_id
            # コンテキストメニューを表示
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
        else:
            # アイテム以外をクリックした場合は選択をクリア
            self.selected_item_id = None
    
    def _get_selected_event(self) -> Tuple[Optional[str], Optional[Dict[str, Any]]]:
        """選択されたイベントを取得"""
        # tree.selection() が空でないことを確認
        selection = self.tree.selection()
        if not selection:
            return None, None
        
        # 選択された行の values から event_number を取得
        # values の想定: (values[0] = event_number)
        item = self.tree.item(selection[0])
        values = item.get('values', [])
        
        if not values or len(values) == 0:
            return None, None
        
        # values[0] から event_number を取得
        event_number = str(values[0])
        
        # event_dictionary[event_number] を取得する
        if event_number in self.event_dictionary:
            event_data = self.event_dictionary[event_number]
            return event_number, event_data
        
        return None, None
    
    def _on_edit_selected(self):
        """選択されたイベントを編集"""
        event_number, event_data = self._get_selected_event()
        if not event_number or not event_data:
            messagebox.showwarning("警告", "イベントを選択してください")
            return
        
        # deprecated=true のイベントは編集不可
        if event_data.get('deprecated', False):
            messagebox.showwarning(
                "警告",
                f"イベント番号 {event_number} は非推奨（deprecated）のため編集できません。"
            )
            return
        
        # 編集後、選択状態をクリア
        self.selected_item_id = None
        self._open_edit_window(event_number, event_data)
    
    def _on_delete_selected(self):
        """選択されたイベントを削除"""
        event_number, event_data = self._get_selected_event()
        if not event_number or not event_data:
            messagebox.showwarning("警告", "イベントを選択してください")
            return
        
        # 確認ダイアログ
        name_jp = event_data.get('name_jp', event_number)
        result = messagebox.askyesno(
            "確認",
            f"イベント「{name_jp} (番号: {event_number})」を削除しますか？\n\n"
            "削除すると、このイベントは非推奨（deprecated）としてマークされます。"
        )
        
        if result:
            # 削除後、選択状態をクリア
            self.selected_item_id = None
            self._delete_event(event_number)
    
    def _open_edit_window(self, event_number: str, event_data: Dict[str, Any]):
        """編集ウィンドウを開く"""
        if not self.event_dict_path:
            messagebox.showerror("エラー", "event_dictionary.json のパスが設定されていません")
            return
        
        from ui.event_dictionary_edit_window import EventDictionaryEditWindow
        
        def on_saved():
            """保存後のコールバック（辞書を再読み込みして表示を更新）"""
            self._load_event_dictionary()
            # ツリービューをクリアして再構築
            for item in self.tree.get_children():
                self.tree.delete(item)
            self._populate_tree()
            # normalization辞書を再生成
            self._regenerate_normalization_dict()
        
        edit_window = EventDictionaryEditWindow(
            parent=self.window,
            event_number=event_number,
            event_data=event_data,
            event_dictionary_path=self.event_dict_path,
            on_saved=on_saved
        )
        edit_window.show()
    
    def _delete_event(self, event_number: str):
        """イベントを削除（deprecatedフラグをtrueに設定）"""
        if not self.event_dict_path:
            messagebox.showerror("エラー", "event_dictionary.json のパスが設定されていません")
            return
        
        try:
            # event_dictionary.json を読み込む
            with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                event_dictionary = json.load(f)
            
            # deprecatedフラグを設定
            if event_number in event_dictionary:
                event_dictionary[event_number]['deprecated'] = True
                
                # 保存（フォルダが存在しない場合は作成）
                if self.event_dict_path:
                    self.event_dict_path.parent.mkdir(parents=True, exist_ok=True)
                with open(self.event_dict_path, 'w', encoding='utf-8') as f:
                    json.dump(event_dictionary, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("完了", f"イベント番号 {event_number} を削除しました（非推奨としてマーク）")
                
                # 表示を更新
                self._load_event_dictionary()
                # ツリービューをクリアして再構築
                for item in self.tree.get_children():
                    self.tree.delete(item)
                # normalization辞書を再生成
                self._regenerate_normalization_dict()
                self._populate_tree()
            else:
                messagebox.showerror("エラー", f"イベント番号 {event_number} が見つかりません")
                
        except Exception as e:
            messagebox.showerror("エラー", f"削除に失敗しました: {e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()

