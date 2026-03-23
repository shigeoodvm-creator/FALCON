"""
FALCON2 - 項目一括編集ウィンドウ
指定された項目を複数の個体で連続編集するためのウィンドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable, Dict, Any, List
from datetime import datetime

from db.db_handler import DBHandler
from modules.batch_update import BatchUpdate


class BatchItemEditWindow:
    """項目一括編集ウィンドウ"""
    
    # 編集可能なcore項目とDBカラムのマッピング
    CORE_ITEM_TO_DB_COLUMN = {
        'COW_ID': 'cow_id',
        'JPN10': 'jpn10',
        'BRD': 'brd',
        'BTHD': 'bthd',
        'ENTR': 'entr',
        'LACT': 'lact',
        'CLVD': 'clvd',
        'PEN': 'pen',
        'RC': 'rc',
    }
    
    # 日付形式の項目
    DATE_ITEMS = ['BTHD', 'ENTR', 'CLVD']
    
    # 数値形式の項目
    INT_ITEMS = ['LACT', 'RC']
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 item_name: str, item_info: Dict[str, Any],
                 on_closed: Optional[Callable[[], None]] = None,
                 formula_engine: Optional[Any] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            item_name: 編集する項目名（例: 'PEN'）
            item_info: 項目の情報（item_dictionaryから）
            on_closed: ウィンドウが閉じられた時のコールバック
            formula_engine: 条件絞り込み用（指定時のみ「条件で対象を絞る」を表示）
        """
        self.parent = parent
        self.db = db_handler
        self.item_name = item_name.upper()
        self.item_info = item_info
        self.on_closed = on_closed
        self.formula_engine = formula_engine
        
        # DBカラム名を取得（core項目はcowテーブル、customはitem_valueテーブル）
        self.db_column = self.CORE_ITEM_TO_DB_COLUMN.get(self.item_name)
        origin = (item_info.get('origin') or item_info.get('type', '')).lower()
        self.is_custom = origin == 'custom' or (not self.db_column and origin != 'core')
        if not self.db_column and not self.is_custom:
            raise ValueError(f"編集不可の項目です: {item_name}")
        
        # 項目の表示名
        self.display_name = item_info.get('display_name', item_name)
        
        # 現在編集中の個体
        self.current_cow: Optional[Dict[str, Any]] = None
        
        # 編集履歴（キャンセル時のロールバック用ではなく、記録用）
        self.edit_history = []
        
        # 条件で絞り込んだ対象リスト（「対象を表示」で設定）
        self.condition_matching_cows: List[Dict[str, Any]] = []
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title(f"項目一括編集 - {self.display_name}")
        self.window.geometry("520x620")
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self._create_widgets()
        
        # 入力欄にフォーカス
        self.cow_id_entry.focus_set()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.window,
            text=f"項目一括編集: {self.display_name} ({self.item_name})",
            font=("", 12, "bold")
        )
        title_label.pack(pady=10)
        
        # 説明
        desc_text = "個体IDを入力して項目を編集するか、条件で対象を絞って一括設定できます。"
        desc_label = ttk.Label(
            self.window,
            text=desc_text,
            foreground="gray"
        )
        desc_label.pack(pady=5)
        
        # メインフレーム
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 条件で対象を絞る（formula_engine が渡されている場合のみ表示）
        if self.formula_engine:
            self._create_condition_filter_section(main_frame)
        
        # 個体ID入力行
        id_row = ttk.Frame(main_frame)
        id_row.pack(fill=tk.X, pady=5)
        ttk.Label(id_row, text="個体ID:", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.cow_id_var = tk.StringVar()
        self.cow_id_entry = ttk.Entry(id_row, textvariable=self.cow_id_var, width=20)
        self.cow_id_entry.pack(side=tk.LEFT, padx=5)
        self.cow_id_entry.bind('<Return>', self._on_cow_id_enter)
        self.cow_id_entry.bind('<KeyRelease>', self._on_cow_id_changed)
        
        # 検索ボタン
        self.search_btn = ttk.Button(id_row, text="検索", command=self._search_cow)
        self.search_btn.pack(side=tk.LEFT, padx=5)
        
        # 候補リスト用のフレーム
        self.cow_candidate_frame = ttk.Frame(main_frame)
        self.cow_candidate_frame.pack(fill=tk.X, pady=5)
        
        # 候補リスト用のTreeview
        candidate_columns = ('cow_id', 'jpn10')
        self.cow_candidate_tree = ttk.Treeview(
            self.cow_candidate_frame,
            columns=candidate_columns,
            show='headings',
            height=5
        )
        self.cow_candidate_tree.heading('cow_id', text='牛ID')
        self.cow_candidate_tree.heading('jpn10', text='個体識別番号')
        self.cow_candidate_tree.column('cow_id', width=100)
        self.cow_candidate_tree.column('jpn10', width=150)
        
        # スクロールバー
        candidate_scrollbar = ttk.Scrollbar(
            self.cow_candidate_frame,
            orient=tk.VERTICAL,
            command=self.cow_candidate_tree.yview
        )
        self.cow_candidate_tree.configure(yscrollcommand=candidate_scrollbar.set)
        
        self.cow_candidate_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        candidate_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ダブルクリックで選択
        self.cow_candidate_tree.bind('<Double-Button-1>', self._on_cow_candidate_selected)
        
        # 初期状態では候補リストを非表示
        self.cow_candidate_frame.pack_forget()
        
        # 個体情報表示（検索後）
        self.cow_info_frame = ttk.LabelFrame(main_frame, text="個体情報")
        self.cow_info_frame.pack(fill=tk.X, pady=10)
        
        self.cow_info_label = ttk.Label(
            self.cow_info_frame, 
            text="個体IDを入力してください",
            foreground="gray"
        )
        self.cow_info_label.pack(padx=10, pady=5)
        
        # 項目編集行
        edit_row = ttk.Frame(main_frame)
        edit_row.pack(fill=tk.X, pady=5)
        ttk.Label(edit_row, text=f"{self.display_name}:", width=12, anchor=tk.W).pack(side=tk.LEFT)
        
        # 編集入力欄
        self.edit_var = tk.StringVar()
        self.edit_entry = ttk.Entry(edit_row, textvariable=self.edit_var, width=30, state="disabled")
        self.edit_entry.pack(side=tk.LEFT, padx=5)
        self.edit_entry.bind('<Return>', self._on_edit_enter)
        
        # 現在の値表示
        self.current_value_label = ttk.Label(edit_row, text="", foreground="blue")
        self.current_value_label.pack(side=tk.LEFT, padx=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=15)
        
        # 保存して次へボタン
        self.save_next_btn = ttk.Button(
            button_frame,
            text="保存して次へ",
            command=self._save_and_next,
            width=15,
            state="disabled"
        )
        self.save_next_btn.pack(side=tk.LEFT, padx=5)
        
        # 終了ボタン
        self.close_btn = ttk.Button(
            button_frame,
            text="終了",
            command=self._on_close,
            width=10
        )
        self.close_btn.pack(side=tk.LEFT, padx=5)
        
        # 編集件数表示
        self.count_label = ttk.Label(
            self.window,
            text="編集件数: 0",
            foreground="gray"
        )
        self.count_label.pack(pady=5)
    
    def _create_condition_filter_section(self, parent: tk.Widget):
        """「条件で対象を絞る」セクションのウィジェットを作成"""
        frame = ttk.LabelFrame(parent, text="条件で対象を絞る")
        frame.pack(fill=tk.X, pady=10)
        
        # 条件入力行
        cond_row = ttk.Frame(frame)
        cond_row.pack(fill=tk.X, pady=5)
        ttk.Label(cond_row, text="条件:", width=8, anchor=tk.W).pack(side=tk.LEFT)
        self.condition_var = tk.StringVar()
        self.condition_entry = ttk.Entry(cond_row, textvariable=self.condition_var, width=25)
        self.condition_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(cond_row, text="対象を表示", command=self._on_show_condition_targets).pack(side=tk.LEFT, padx=5)
        
        # 条件結果ツリー（牛ID, 個体識別番号, 現在値）
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.X, pady=5)
        cond_columns = ('cow_id', 'jpn10', 'current_val')
        self.condition_result_tree = ttk.Treeview(
            tree_frame,
            columns=cond_columns,
            show='headings',
            height=4
        )
        self.condition_result_tree.heading('cow_id', text='牛ID')
        self.condition_result_tree.heading('jpn10', text='個体識別番号')
        self.condition_result_tree.heading('current_val', text='現在の値')
        self.condition_result_tree.column('cow_id', width=80)
        self.condition_result_tree.column('jpn10', width=120)
        self.condition_result_tree.column('current_val', width=80)
        cond_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.condition_result_tree.yview)
        self.condition_result_tree.configure(yscrollcommand=cond_scroll.set)
        self.condition_result_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        cond_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 一括でこの値に設定
        batch_row = ttk.Frame(frame)
        batch_row.pack(fill=tk.X, pady=5)
        ttk.Label(batch_row, text="一括設定値:", width=10, anchor=tk.W).pack(side=tk.LEFT)
        self.batch_value_var = tk.StringVar()
        self.batch_value_entry = ttk.Entry(batch_row, textvariable=self.batch_value_var, width=15)
        self.batch_value_entry.pack(side=tk.LEFT, padx=5)
        self.batch_apply_btn = ttk.Button(
            batch_row,
            text="全対象に適用",
            command=self._on_apply_batch_to_condition_targets,
            state="disabled"
        )
        self.batch_apply_btn.pack(side=tk.LEFT, padx=5)
    
    def _on_show_condition_targets(self):
        """「対象を表示」クリック: 条件に合う個体を取得してツリーに表示"""
        condition = self.condition_var.get().strip()
        if not condition:
            messagebox.showwarning("警告", "条件を入力してください。例: LACT=0")
            return
        try:
            batch_update = BatchUpdate(self.db, self.formula_engine)
            self.condition_matching_cows = batch_update.get_matching_cows(condition)
        except Exception as e:
            messagebox.showerror("エラー", f"条件の評価に失敗しました:\n{e}")
            return
        for item in self.condition_result_tree.get_children():
            self.condition_result_tree.delete(item)
        if not self.condition_matching_cows:
            messagebox.showinfo("結果", f"条件「{condition}」に該当する個体はありません。")
            self.batch_apply_btn.config(state="disabled")
            return
        for cow in self.condition_matching_cows:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            if self.db_column:
                cur = cow.get(self.db_column, '')
            else:
                cur = self.db.get_item_value(cow.get('auto_id'), self.item_name) or ''
            if cur is None:
                cur = ''
            self.condition_result_tree.insert('', 'end', values=(cow_id, jpn10, str(cur)))
        self.batch_apply_btn.config(state="normal")
        messagebox.showinfo("結果", f"条件「{condition}」に該当する {len(self.condition_matching_cows)} 頭を表示しました。")
    
    def _on_apply_batch_to_condition_targets(self):
        """「全対象に適用」クリック: 表示中の全個体に一括設定値を適用"""
        if not self.condition_matching_cows:
            messagebox.showwarning("警告", "先に「対象を表示」で対象を表示してください。")
            return
        value_str = self.batch_value_var.get().strip()
        if not self._validate_value(value_str):
            return
        converted = self._convert_value(value_str)
        updated = 0
        for cow in self.condition_matching_cows:
            auto_id = cow.get('auto_id')
            if not auto_id:
                continue
            try:
                if self.db_column:
                    old_val = cow.get(self.db_column)
                    self.db.update_cow(auto_id, {self.db_column: converted})
                else:
                    old_val = self.db.get_item_value(auto_id, self.item_name)
                    self.db.set_item_value(auto_id, self.item_name, converted)
                updated += 1
                self.edit_history.append({
                    'cow_id': cow.get('cow_id'),
                    'old_value': old_val,
                    'new_value': converted
                })
            except Exception as e:
                messagebox.showerror("エラー", f"更新に失敗しました:\n{e}")
                return
        self.count_label.config(text=f"編集件数: {len(self.edit_history)}")
        messagebox.showinfo("完了", f"{updated} 頭の {self.display_name} を「{value_str}」に更新しました。")
        self.batch_apply_btn.config(state="disabled")
        self.condition_matching_cows = []
        for item in self.condition_result_tree.get_children():
            self.condition_result_tree.delete(item)
        self.batch_value_var.set("")
    
    def _on_cow_id_changed(self, event=None):
        """牛ID入力変更時の処理（リアルタイムで候補を絞り込む）"""
        # 特殊キー（Enter, Tab, Shift等）の場合は処理しない
        if event and event.keysym in ('Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Up', 'Down'):
            return
        
        cow_id_str = self.cow_id_var.get().strip()
        
        if not cow_id_str:
            # 空欄の場合は候補リストを非表示
            self.cow_candidate_frame.pack_forget()
            return
        
        # 数字のみの場合は前方一致で検索
        if cow_id_str.isdigit():
            candidates = self.db.search_cows_by_id_prefix(cow_id_str, limit=50)
            self._update_cow_candidates(candidates)
        else:
            # 数字以外が含まれている場合は候補リストを非表示
            self.cow_candidate_frame.pack_forget()
    
    def _update_cow_candidates(self, candidates: List[Dict[str, Any]]):
        """
        牛の候補リストを更新
        
        Args:
            candidates: 候補牛のリスト
        """
        # 既存のアイテムをクリア
        for item in self.cow_candidate_tree.get_children():
            self.cow_candidate_tree.delete(item)
        
        if not candidates:
            # 候補がない場合は非表示
            self.cow_candidate_frame.pack_forget()
            return
        
        # 候補を表示
        for cow in candidates:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            
            # auto_idをtagsに保存
            auto_id = cow.get('auto_id')
            self.cow_candidate_tree.insert(
                '',
                'end',
                values=(cow_id, jpn10),
                tags=(str(auto_id),)
            )
        
        # 候補リストを表示（個体情報フレームの前に配置）
        self.cow_candidate_frame.pack(fill=tk.X, pady=5, before=self.cow_info_frame)
    
    def _on_cow_candidate_selected(self, event=None):
        """候補リストでダブルクリックされた時の処理"""
        selected_items = self.cow_candidate_tree.selection()
        if not selected_items:
            # ダブルクリックされた行を取得
            item = self.cow_candidate_tree.identify_row(event.y) if event else None
            if item:
                selected_items = [item]
            else:
                return
        
        item = selected_items[0]
        item_data = self.cow_candidate_tree.item(item)
        tags = item_data.get('tags', [])
        values = item_data.get('values', [])
        
        if tags and values:
            auto_id = int(tags[0])
            cow_id = values[0]
            
            # 牛を選択
            self.cow_id_var.set(cow_id)
            cow = self.db.get_cow_by_auto_id(auto_id)
            if cow:
                self._select_cow(cow)
            
            # 候補リストを非表示
            self.cow_candidate_frame.pack_forget()
    
    def _on_cow_id_enter(self, event=None):
        """個体ID入力欄でEnterキーが押された時"""
        # 候補リストが表示されている場合は、候補が1件だけなら自動選択
        if self.cow_candidate_frame.winfo_viewable():
            candidates = self.cow_candidate_tree.get_children()
            if len(candidates) == 1:
                # 候補が1件だけの場合は自動選択
                self.cow_candidate_tree.selection_set(candidates[0])
                self._on_cow_candidate_selected(None)
                return
            elif len(candidates) > 1:
                # 候補が複数ある場合は候補リストから選択を促す
                return
        
        # 候補リストが表示されていない場合は通常の検索処理
        self._search_cow()
    
    def _search_cow(self):
        """個体を検索"""
        cow_id = self.cow_id_var.get().strip()
        if not cow_id:
            messagebox.showwarning("警告", "個体IDを入力してください")
            return
        
        # 4桁にゼロパディング
        cow_id_padded = cow_id.zfill(4)
        
        # 個体を検索
        cow = self.db.get_cow_by_id(cow_id_padded)
        if not cow:
            # 個体識別番号（10桁）でも検索
            cow = self.db.get_cow_by_jpn10(cow_id)
        
        if not cow:
            messagebox.showwarning("警告", f"個体が見つかりません: {cow_id}")
            self.cow_id_entry.select_range(0, tk.END)
            self.cow_id_entry.focus_set()
            return
        
        self._select_cow(cow)
    
    def _select_cow(self, cow: Dict[str, Any]):
        """個体を選択して表示"""
        self.current_cow = cow
        
        # 個体情報を表示
        cow_id_display = cow.get('cow_id', '')
        jpn10 = cow.get('jpn10', '')
        brd = cow.get('brd', '')
        lact = cow.get('lact', '')
        self.cow_info_label.config(
            text=f"ID: {cow_id_display}  JPN10: {jpn10}  品種: {brd}  産次: {lact}",
            foreground="black"
        )
        
        # 現在の値を取得して表示（coreはcowテーブル、customはitem_valueテーブル）
        if self.db_column:
            current_value = cow.get(self.db_column, '')
        else:
            current_value = self.db.get_item_value(cow.get('auto_id'), self.item_name) or ''
        if current_value is None:
            current_value = ''
        self.current_value_label.config(text=f"(現在: {current_value})")
        
        # 編集入力欄を有効化して値をセット
        self.edit_entry.config(state="normal")
        self.edit_var.set(str(current_value))
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus_set()
        
        # 保存ボタンを有効化
        self.save_next_btn.config(state="normal")
    
    def _on_edit_enter(self, event=None):
        """編集入力欄でEnterキーが押された時"""
        self._save_and_next()
    
    def _save_and_next(self):
        """保存して次の個体へ"""
        if not self.current_cow:
            return
        
        new_value = self.edit_var.get().strip()
        
        # バリデーション
        if not self._validate_value(new_value):
            return
        
        # 型変換
        converted_value = self._convert_value(new_value)
        
        try:
            auto_id = self.current_cow.get('auto_id')
            if self.db_column:
                # core項目: cowテーブルを更新
                cow_data = {self.db_column: converted_value}
                self.db.update_cow(auto_id, cow_data)
                old_value = self.current_cow.get(self.db_column)
            else:
                # custom項目: item_valueテーブルを更新
                old_value = self.db.get_item_value(auto_id, self.item_name)
                self.db.set_item_value(auto_id, self.item_name, converted_value)
            
            # 編集履歴に追加
            self.edit_history.append({
                'cow_id': self.current_cow.get('cow_id'),
                'old_value': old_value,
                'new_value': converted_value
            })
            
            # 件数表示を更新
            self.count_label.config(text=f"編集件数: {len(self.edit_history)}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"更新に失敗しました:\n{e}")
            return
        
        # 次の入力に備えてリセット
        self._reset_for_next()
    
    def _validate_value(self, value: str) -> bool:
        """入力値のバリデーション"""
        # 日付項目のチェック
        if self.item_name in self.DATE_ITEMS:
            if value and not self._is_valid_date(value):
                messagebox.showerror("エラー", f"{self.display_name}の形式が正しくありません (YYYY-MM-DD)")
                return False
        
        # 数値項目のチェック（core + customのdata_type）
        is_int_item = self.item_name in self.INT_ITEMS or (
            self.is_custom and (self.item_info.get('data_type') or '').lower() == 'int'
        )
        if is_int_item:
            if value:
                try:
                    int(value)
                except ValueError:
                    messagebox.showerror("エラー", f"{self.display_name}は数値である必要があります")
                    return False
        
        return True
    
    def _convert_value(self, value: str) -> Any:
        """値を適切な型に変換"""
        if not value:
            return None
        
        if self.item_name in self.INT_ITEMS:
            return int(value)
        if self.is_custom and (self.item_info.get('data_type') or '').lower() == 'int':
            return int(value)
        
        return value
    
    def _is_valid_date(self, date_str: str) -> bool:
        """日付形式の簡易チェック"""
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            return True
        except ValueError:
            return False
    
    def _reset_for_next(self):
        """次の入力のためにリセット"""
        self.current_cow = None
        self.cow_id_var.set("")
        self.edit_var.set("")
        self.edit_entry.config(state="disabled")
        self.save_next_btn.config(state="disabled")
        self.cow_info_label.config(text="個体IDを入力してください", foreground="gray")
        self.current_value_label.config(text="")
        
        # 候補リストを非表示
        self.cow_candidate_frame.pack_forget()
        
        # 個体ID入力欄にフォーカス
        self.cow_id_entry.focus_set()
    
    def _on_close(self):
        """終了処理"""
        if self.edit_history:
            # 編集した件数を表示
            messagebox.showinfo(
                "完了",
                f"{len(self.edit_history)}件の{self.display_name}を編集しました"
            )
        
        self.window.destroy()
        
        if self.on_closed:
            self.on_closed()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
