"""
FALCON2 - 個体一覧ウィンドウ
個体一覧を表示し、ダブルクリックで個体カードを開く
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable
from pathlib import Path

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from ui.cow_card import CowCard


class CowListWindow:
    """個体一覧ウィンドウ（Toplevel）"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler,
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 event_dictionary_path: Path,
                 item_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（MainWindow）
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス（オプション）
        """
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.event_dict_path = event_dictionary_path
        self.item_dict_path = item_dictionary_path
        
        # 開いている個体カードウィンドウを管理（cow_auto_id -> window）
        self.open_cow_windows: Dict[int, tk.Toplevel] = {}
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("個体一覧")
        self.window.geometry("1000x600")
        self.window.transient(parent)
        
        self._create_widgets()
        self._load_cows()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        title_label = ttk.Label(
            main_frame,
            text="個体一覧",
            font=("", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 検索欄（将来の拡張用）
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_changed)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # 個体リスト（Treeview）
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('cow_id', 'jpn10', 'brd', 'lact', 'clvd', 'rc', 'pen')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)
        
        # カラム設定
        self.tree.heading('cow_id', text='管理番号')
        self.tree.heading('jpn10', text='個体識別番号')
        self.tree.heading('brd', text='品種')
        self.tree.heading('lact', text='産次')
        self.tree.heading('clvd', text='最終分娩日')
        self.tree.heading('rc', text='繁殖区分')
        self.tree.heading('pen', text='群')
        
        self.tree.column('cow_id', width=100)
        self.tree.column('jpn10', width=120)
        self.tree.column('brd', width=100)
        self.tree.column('lact', width=60)
        self.tree.column('clvd', width=100)
        self.tree.column('rc', width=100)
        self.tree.column('pen', width=100)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 配置
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ダブルクリックイベント
        self.tree.bind('<Double-1>', self._on_cow_double_click)
        
        # 右クリックメニュー
        self.tree.bind('<Button-3>', self._on_right_click)  # Windows/Linux
        self.tree.bind('<Button-2>', self._on_right_click)  # Mac
        
        # 右クリックメニューを作成
        self.context_menu = tk.Menu(self.window, tearoff=0)
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
        
        # 繁殖コード表示名
        rc_names = {
            1: "Fresh（分娩後）",
            2: "Bred（授精後）",
            3: "Pregnant（妊娠中）",
            4: "Dry（乾乳中）",
            5: "Open（空胎）",
            6: "Stopped（繁殖停止）"
        }
        
        # Treeviewに追加
        for cow in cows:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            lact = cow.get('lact', '')
            clvd = cow.get('clvd', '') or ''
            rc = cow.get('rc', '')
            pen = cow.get('pen', '')
            
            # 繁殖コードを表示名に変換
            rc_text = rc_names.get(rc, '') if rc is not None else ''
            
            self.tree.insert('', 'end', values=(
                cow_id, jpn10, brd, lact, clvd, rc_text, pen
            ), tags=(cow['auto_id'],))
    
    def _on_search_changed(self, *args):
        """検索テキスト変更時の処理"""
        search_text = self.search_var.get()
        self._load_cows(search_text)
    
    def _on_cow_double_click(self, event):
        """個体のダブルクリック処理（個体カードを開く）"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = self.tree.item(selection[0])
        tags = item.get('tags', [])
        if tags:
            cow_auto_id = int(tags[0])
            self._open_cow_card(cow_auto_id)
    
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
        """個体編集（個体カードを開く）"""
        if self.selected_cow_auto_id is None:
            return
        
        self._open_cow_card(self.selected_cow_auto_id)
    
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
            f"管理番号: {cow_id}\n"
            f"個体識別番号: {jpn10}\n\n"
            f"※ この操作は取り消せません。\n"
            f"※ 関連するイベントも削除されます。"
        )
        
        if result:
            try:
                # 個体を削除
                self.db.delete_cow(self.selected_cow_auto_id)
                messagebox.showinfo("完了", "個体を削除しました")
                
                # 開いている個体カードウィンドウがあれば閉じる
                if self.selected_cow_auto_id in self.open_cow_windows:
                    window = self.open_cow_windows[self.selected_cow_auto_id]
                    window.destroy()
                    del self.open_cow_windows[self.selected_cow_auto_id]
                
                # 一覧を更新
                self._load_cows(self.search_var.get())
                
                # 選択をクリア
                self.selected_cow_auto_id = None
            except Exception as e:
                messagebox.showerror("エラー", f"削除に失敗しました:\n{e}")
    
    def _open_cow_card(self, cow_auto_id: int):
        """
        個体カードウィンドウを開く
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 既に同じ個体カードが開いている場合は前面に出す
        if cow_auto_id in self.open_cow_windows:
            window = self.open_cow_windows[cow_auto_id]
            window.lift()
            window.focus()
            return
        
        # 新しい個体カードウィンドウを作成
        card_window = tk.Toplevel(self.window)
        card_window.title("個体カード")
        card_window.geometry("800x700")
        card_window.transient(self.window)
        
        # 個体カードを埋め込む
        cow_card = CowCard(
            parent=card_window,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=self.item_dict_path
        )
        
        # 個体を読み込む
        cow_card.load_cow(cow_auto_id)
        
        # ウィジェットを配置
        cow_card.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ウィンドウを管理
        self.open_cow_windows[cow_auto_id] = card_window
        
        # ウィンドウが閉じられたときに管理から削除
        def on_window_close():
            if cow_auto_id in self.open_cow_windows:
                del self.open_cow_windows[cow_auto_id]
            card_window.destroy()
        
        card_window.protocol("WM_DELETE_WINDOW", on_window_close)
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()




