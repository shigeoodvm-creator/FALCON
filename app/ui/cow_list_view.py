"""
FALCON2 - 個体一覧ビュー
個体一覧を表示し、ダブルクリックで個体カードを表示
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable
from db.db_handler import DBHandler
from ui.cow_edit_window import CowEditWindow


class CowListView:
    """個体一覧ビュー"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 on_cow_selected: Optional[Callable[[int], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            on_cow_selected: 個体選択時のコールバック（cow_auto_id を引数に取る）
        """
        self.db = db_handler
        self.on_cow_selected = on_cow_selected
        
        # UI作成
        self.frame = ttk.Frame(parent)
        self._create_widgets()
        
        # 個体リストを読み込み
        self._load_cows()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.frame,
            text="個体一覧",
            font=("", 14, "bold")
        )
        title_label.pack(pady=10)
        
        # 検索欄
        search_frame = ttk.Frame(self.frame)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_changed)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # 個体リスト（Treeview）
        list_frame = ttk.Frame(self.frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ('cow_id', 'jpn10', 'brd', 'lact', 'rc', 'pen')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)
        
        # カラム設定
        self.tree.heading('cow_id', text='拡大4桁ID')
        self.tree.heading('jpn10', text='個体識別番号')
        self.tree.heading('brd', text='品種')
        self.tree.heading('lact', text='産次')
        self.tree.heading('rc', text='繁殖コード')
        self.tree.heading('pen', text='群')
        
        self.tree.column('cow_id', width=100)
        self.tree.column('jpn10', width=120)
        self.tree.column('brd', width=100)
        self.tree.column('lact', width=60)
        self.tree.column('rc', width=100)
        self.tree.column('pen', width=100)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 配置
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ダブルクリックイベント
        self.tree.bind('<Double-Button-1>', self._on_cow_double_click)
        
        # 右クリックメニュー
        self.tree.bind('<Button-3>', self._on_right_click)  # Windows/Linux
        self.tree.bind('<Button-2>', self._on_right_click)  # Mac
        
        # 右クリックメニューを作成
        self.context_menu = tk.Menu(self.frame, tearoff=0)
        self.context_menu.add_command(label="編集", command=self._on_edit_cow)
        self.context_menu.add_command(label="削除", command=self._on_delete_cow)
        
        # 選択中の個体IDを保持
        self.selected_cow_auto_id: Optional[int] = None
    
    def _load_cows(self, search_text: str = ""):
        """
        個体リストを読み込み
        
        Args:
            search_text: 検索テキスト（拡大4桁IDまたは個体識別番号で検索）
        """
        # 既存のアイテムをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 全個体を取得
        cows = self.db.get_all_cows()
        
        # 検索フィルタリング
        if search_text:
            search_text = search_text.strip().lower()
            cows = [
                cow for cow in cows
                if search_text in str(cow.get('cow_id', '')).lower()
                or search_text in str(cow.get('jpn10', '')).lower()
            ]
        
        # Treeviewに追加
        for cow in cows:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            lact = cow.get('lact', '')
            rc = cow.get('rc', '')
            pen = cow.get('pen', '')
            
            # 繁殖コードを表示名に変換
            rc_names = {
                1: "Fresh",
                2: "Bred",
                3: "Pregnant",
                4: "Dry",
                5: "Open",
                6: "Stopped"
            }
            rc_text = rc_names.get(rc, '') if rc is not None else ''
            
            self.tree.insert('', 'end', values=(
                cow_id, jpn10, brd, lact, rc_text, pen
            ), tags=(cow['auto_id'],))
    
    def _on_search_changed(self, *args):
        """検索テキスト変更時の処理"""
        search_text = self.search_var.get()
        self._load_cows(search_text)
    
    def _on_cow_double_click(self, event):
        """個体のダブルクリック処理"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            tags = item.get('tags', [])
            if tags and self.on_cow_selected:
                cow_auto_id = int(tags[0])
                self.on_cow_selected(cow_auto_id)
    
    def _on_right_click(self, event):
        """右クリック処理"""
        # クリック位置のアイテムを選択
        item = self.tree.identify_row(event.y)
        if item:
            # アイテムを選択状態にする
            self.tree.selection_set(item)
            
            # 個体IDを取得
            item_data = self.tree.item(item)
            tags = item_data.get('tags', [])
            if tags:
                self.selected_cow_auto_id = int(tags[0])
                
                # メニューを表示
                try:
                    self.context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    self.context_menu.grab_release()
    
    def _on_edit_cow(self):
        """個体編集"""
        if self.selected_cow_auto_id is None:
            return
        
        # 編集ウィンドウを表示
        edit_window = CowEditWindow(
            self.frame,
            self.db,
            self.selected_cow_auto_id,
            on_saved=self.refresh
        )
        edit_window.show()
    
    def _on_delete_cow(self):
        """個体削除"""
        if self.selected_cow_auto_id is None:
            return
        
        # 個体情報を取得
        cow = self.db.get_cow_by_auto_id(self.selected_cow_auto_id)
        if not cow:
            messagebox.showerror("エラー", "個体が見つかりません")
            return
        
        cow_id = cow.get('cow_id', '')
        jpn10 = cow.get('jpn10', '')
        
        # 確認ダイアログ
        result = messagebox.askyesno(
            "確認",
            f"個体を削除しますか？\n\n"
            f"拡大4桁ID: {cow_id}\n"
            f"個体識別番号: {jpn10}\n\n"
            f"※ この操作は取り消せません。\n"
            f"※ 関連するイベントも削除されます。"
        )
        
        if result:
            try:
                # 個体を削除
                self.db.delete_cow(self.selected_cow_auto_id)
                messagebox.showinfo("完了", "個体を削除しました")
                
                # 一覧を更新
                self.refresh()
                
                # 選択をクリア
                self.selected_cow_auto_id = None
            except Exception as e:
                messagebox.showerror("エラー", f"削除に失敗しました:\n{e}")
    
    def get_cow_by_id(self, cow_id: str) -> Optional[int]:
        """
        拡大4桁IDから個体のauto_idを取得
        
        Args:
            cow_id: 拡大4桁ID
        
        Returns:
            cow_auto_id、見つからない場合はNone
        """
        cow = self.db.get_cow_by_id(cow_id)
        if cow:
            return cow.get('auto_id')
        return None
    
    def refresh(self):
        """一覧を更新"""
        search_text = self.search_var.get()
        self._load_cows(search_text)
    
    def get_widget(self) -> ttk.Frame:
        """ウィジェットを取得"""
        return self.frame

