"""
FALCON2 - 項目一覧ウィンドウ
コマンドライン入力欄から項目を選択するためのウィンドウ
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any, Optional, Callable, List, Tuple
from collections import defaultdict


def _get_label(data: Dict[str, Any], key: str) -> str:
    """表示名のフォールバック取得"""
    return (
        data.get("label")
        or data.get("display_name")
        or data.get("name_jp")
        or key
    )


def _get_category(data: Dict[str, Any]) -> str:
    """カテゴリーを取得"""
    return data.get("category") or "その他"


class ItemListWindow:
    """項目一覧ウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        item_dictionary: Dict[str, Any],
        command_history: Dict[str, int],
        on_item_selected: Optional[Callable[[str], None]] = None,
        event_dictionary: Optional[Dict[str, Any]] = None,
    ):
        """
        初期化

        Args:
            parent: 親ウィンドウ
            item_dictionary: 項目辞書
            command_history: コマンド入力履歴（{項目名: 入力回数}）
            on_item_selected: 項目選択時のコールバック関数（項目キーを引数に取る）
            event_dictionary: イベント辞書（オプション）
        """
        self.parent = parent
        self.item_dictionary = item_dictionary
        self.event_dictionary = event_dictionary or {}
        self.command_history = command_history
        self.on_item_selected = on_item_selected

        self.window = tk.Toplevel(parent)
        self.window.title("項目一覧")
        self.window.geometry("750x550")
        
        # ウィンドウの背景色を設定
        self.window.configure(bg="#f5f5f5")
        
        # 検索用の変数
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *args: self._on_search_changed())

        self._create_widgets()
        self._populate_items()

        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """ウィジェットを作成"""
        # スタイル設定
        style = ttk.Style()
        
        # タイトルフレーム
        title_frame = ttk.Frame(self.window)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 10))
        
        title_label = tk.Label(
            title_frame,
            text="項目一覧",
            font=("", 14, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(
            title_frame,
            text="ダブルクリックでコマンドラインに転記",
            font=("", 9),
            foreground="#7f8c8d"
        )
        subtitle_label.pack(anchor=tk.W, pady=(2, 0))
        
        # 検索欄
        search_frame = ttk.Frame(self.window)
        search_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        ttk.Label(search_frame, text="検索:", font=("", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        clear_btn = ttk.Button(search_frame, text="クリア", command=self._clear_search, width=8)
        clear_btn.pack(side=tk.LEFT, padx=(5, 0))

        # メインフレーム（スクロール可能なリスト）
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        # Treeview（2列表示用）
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # スクロールバー
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview（2列表示）
        self.item_treeview = ttk.Treeview(
            tree_frame,
            columns=("japanese",),
            show="tree headings",
            yscrollcommand=scrollbar.set,
            selectmode=tk.BROWSE
        )
        self.item_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.item_treeview.yview)

        # カラムの設定
        self.item_treeview.heading("#0", text="項目名", anchor=tk.W)
        self.item_treeview.heading("japanese", text="日本語名", anchor=tk.W)
        self.item_treeview.column("#0", width=250, anchor=tk.W, minwidth=200)
        self.item_treeview.column("japanese", width=400, anchor=tk.W, minwidth=300)
        
        # モダンなスタイル設定
        style.configure("Treeview",
                       font=("", 10),
                       rowheight=25,
                       background="#fafafa",
                       foreground="#2c3e50",
                       fieldbackground="#fafafa")
        style.configure("Treeview.Heading",
                       font=("", 10, "bold"),
                       background="#ecf0f1",
                       foreground="#2c3e50",
                       relief=tk.FLAT)
        style.map("Treeview.Heading",
                 background=[("active", "#d5dbdb")])
        style.map("Treeview",
                 background=[("selected", "#3498db")],
                 foreground=[("selected", "#ffffff")])

        # カテゴリ見出し用のタグスタイル（フォント設定はTreeviewタグでは使用不可の可能性があるため削除）
        self.item_treeview.tag_configure("category", 
                                         background="#34495e",
                                         foreground="#ffffff")

        # ダブルクリックで選択
        self.item_treeview.bind("<Double-Button-1>", self._on_item_double_click)

        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=(0, 15))

        close_button = ttk.Button(
            button_frame,
            text="閉じる",
            command=self.window.destroy,
            width=12
        )
        close_button.pack()

    def _populate_items(self):
        """項目をリストに追加（2列表示、検索フィルタリング対応）"""
        try:
            # 既存の項目をクリア
            for item in self.item_treeview.get_children():
                self.item_treeview.delete(item)
            
            # 項目キーとTreeviewのアイテムIDの対応を保存
            self.item_key_map: Dict[str, str] = {}  # {item_id: item_key}

            # 項目のみを表示（イベントは除外）
            all_items = self.item_dictionary if self.item_dictionary else {}
            
            # デバッグ用：項目辞書が空でないか確認
            if not all_items:
                # 項目辞書が空の場合はメッセージを表示
                self.item_treeview.insert("", tk.END, text="(項目が見つかりません)", values=("",))
                return

            # 検索フィルタリング
            search_text = self.search_var.get().strip().lower()

            # 頻度の高い項目を取得（上位7個）
            frequent_items = []
            if self.command_history:
                sorted_by_frequency = sorted(
                    self.command_history.items(),
                    key=lambda x: x[1],
                    reverse=True
                )[:7]  # 上位7個
                frequent_items = [item[0] for item in sorted_by_frequency if item[0] in all_items]

            # 頻度の高い項目を表示
            if frequent_items:
                # 検索フィルタリングを適用
                filtered_frequent_items = []
                for item_key in frequent_items:
                    item_data = all_items.get(item_key)
                    if item_data:
                        label = _get_label(item_data, item_key)
                        # 検索フィルタリング
                        if search_text:
                            if (search_text not in item_key.lower() and 
                                search_text not in label.lower()):
                                continue
                        filtered_frequent_items.append((item_key, item_data))
                
                if filtered_frequent_items:
                    frequent_parent = self.item_treeview.insert("", tk.END, text="よく使う項目", values=("",), tags=("category",))
                    for item_key, item_data in filtered_frequent_items:
                        label = _get_label(item_data, item_key)
                        item_id = self.item_treeview.insert(
                            frequent_parent,
                            tk.END,
                            text=item_key,
                            values=(label,),
                            tags=("frequent",)
                        )
                        # 項目キーを保存
                        self.item_key_map[item_id] = item_key
                    # タグのスタイル設定（よく使う項目の背景色をより洗練された色に）
                    self.item_treeview.tag_configure("frequent", background="#ebf5fb")

            # カテゴリ別に項目を整理
            items_by_category = defaultdict(list)
            for item_key, item_data in all_items.items():
                # 頻度の高い項目はスキップ
                if item_key in frequent_items:
                    continue
                
                # 検索フィルタリング
                if search_text:
                    label = _get_label(item_data, item_key)
                    if (search_text not in item_key.lower() and 
                        search_text not in label.lower()):
                        continue
                
                category = _get_category(item_data)
                items_by_category[category].append((item_key, item_data))

            # カテゴリごとに表示
            if items_by_category:
                for category in sorted(items_by_category.keys()):
                    category_parent = self.item_treeview.insert("", tk.END, text=category, values=("",), tags=("category",))
                    # カテゴリ内の項目をソート（キー順）
                    sorted_items = sorted(items_by_category[category], key=lambda x: x[0])
                    for item_key, item_data in sorted_items:
                        label = _get_label(item_data, item_key)
                        item_id = self.item_treeview.insert(
                            category_parent,
                            tk.END,
                            text=item_key,
                            values=(label,)
                        )
                        # 項目キーを保存
                        self.item_key_map[item_id] = item_key
            elif not frequent_items:
                # 項目辞書はあるが、カテゴリがない場合（通常はないはず）
                if not search_text:
                    self.item_treeview.insert("", tk.END, text="(項目が見つかりません)", values=("",))
                else:
                    self.item_treeview.insert("", tk.END, text="(検索結果が見つかりません)", values=("",))
        except Exception as e:
            # エラーが発生した場合はメッセージを表示
            import logging
            logging.error(f"項目一覧の表示でエラーが発生しました: {e}", exc_info=True)
            self.item_treeview.insert("", tk.END, text=f"(エラー: {str(e)})", values=("",))
    
    def _on_search_changed(self):
        """検索欄の変更時にフィルタリングを実行"""
        self._populate_items()
        # 選択をクリア
        self.item_treeview.selection_remove(self.item_treeview.selection())
    
    def _clear_search(self):
        """検索欄をクリアしてフィルタリングをリセット"""
        self.search_var.set("")
        self.search_entry.delete(0, tk.END)
        # フィルタリングを実行（空文字列で全項目を表示）
        self._populate_items()
        # 検索欄にフォーカスを戻す
        self.search_entry.focus_set()

    def _on_item_double_click(self, event):
        """項目がダブルクリックされた時の処理"""
        selection = self.item_treeview.selection()
        if not selection:
            return

        item_id = selection[0]
        # カテゴリ行の場合は何もしない（タグで判定）
        tags = self.item_treeview.item(item_id, "tags")
        if "category" in tags:
            return

        # 項目キーを取得
        item_key = self.item_key_map.get(item_id)
        if item_key and self.on_item_selected:
            self.on_item_selected(item_key)
            # ウィンドウを閉じずに、連続選択できるようにする

