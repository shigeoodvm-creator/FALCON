"""
FALCON2 - 個体一覧ウィンドウ
個体一覧を表示し、ダブルクリックで個体カードを開く
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable
from pathlib import Path
from datetime import date, datetime

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from ui.cow_card import CowCard
from ui.cow_card_window import CowCardWindow


class CowListWindow:
    """個体一覧ウィンドウ（Toplevel）"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler,
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 event_dictionary_path: Path,
                 item_dictionary_path: Optional[Path] = None,
                 on_close: Optional[Callable[[], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（MainWindow）
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス（オプション）
            on_close: ウィンドウ閉鎖時に呼ぶコールバック（メインが参照を外す用）
        """
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.event_dict_path = event_dictionary_path
        self.item_dict_path = item_dictionary_path
        self.on_close = on_close
        
        # 開いている個体カードウィンドウを管理（cow_auto_id -> window）
        self.open_cow_windows: Dict[int, CowCardWindow] = {}
        
        # ウィンドウ作成（繁殖検診フロー・個体カードと同一デザイン）
        self.window = tk.Toplevel(parent)
        self.window.title("個体一覧")
        self.window.minsize(800, 500)
        self.window.geometry("1000x600")
        self.window.configure(bg="#f5f5f5")
        self.window.update_idletasks()  # レイアウト確定を促す
        
        def _on_delete():
            if self.on_close:
                self.on_close()
            self.window.destroy()
        self.window.protocol("WM_DELETE_WINDOW", _on_delete)
        
        self._create_widgets()
        self._load_cows()
        self.window.update_idletasks()  # 子ウィジェット配置後に再計算
    
    def _create_widgets(self):
        """ウィジェットを作成（繁殖検診フロー・個体カードと同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        
        main_frame = tk.Frame(self.window, bg=bg, padx=24, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(3, weight=1)   # 一覧行が残りスペースを取得
        main_frame.grid_columnconfigure(0, weight=1)
        
        # ========== ヘッダー（繁殖検診フローと同じイメージ） ==========
        header = tk.Frame(main_frame, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\U0001f4dc", font=(_df, 24), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="個体一覧", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="一覧から個体を選択し、ダブルクリックで個体カードを開きます", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)
        
        # ========== 検索・フィルタ（1段目） ==========
        search_row = tk.Frame(main_frame, bg=bg)
        search_row.grid(row=1, column=0, sticky=tk.W, pady=(10, 4))
        
        tk.Label(search_row, text="検索:", font=(_df, 10), bg=bg, fg="#263238").pack(side=tk.LEFT, padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_changed)
        search_entry = tk.Entry(search_row, textvariable=self.search_var, width=28, font=(_df, 10), relief=tk.SOLID, borderwidth=1)
        search_entry.pack(side=tk.LEFT, padx=4)
        
        self.show_existing_only_var = tk.BooleanVar(value=False)
        show_existing_cb = tk.Checkbutton(
            search_row, text="現存牛のみ", variable=self.show_existing_only_var,
            font=(_df, 10), bg=bg, fg="#263238", selectcolor=bg, activebackground=bg, activeforeground="#263238",
            command=self._on_filter_changed
        )
        show_existing_cb.pack(side=tk.LEFT, padx=16)
        
        # ========== 牛群内訳（2段目） ==========
        stats_row = tk.Frame(main_frame, bg=bg)
        stats_row.grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        self.stats_label = tk.Label(stats_row, text="", font=(_df, 9), bg=bg, fg="#607d8b")
        self.stats_label.pack(side=tk.LEFT)
        
        # ========== 個体リスト（Treeview） ==========
        list_frame = tk.Frame(main_frame, bg=bg)
        list_frame.grid(row=3, column=0, sticky=tk.NSEW)
        
        # 個体一覧専用の Treeview スタイル（他画面で Treeview が上書きされても一覧のフォントを維持）
        try:
            style = ttk.Style()
            style.configure("CowList.Treeview", font=(_df, 10))
            style.configure("CowList.Treeview.Heading", font=(_df, 10, "bold"))
        except tk.TclError:
            pass
        
        columns = ('cow_id', 'jpn10', 'lact', 'clvd', 'dim', 'rc', 'pen')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20, style="CowList.Treeview")
        
        # カラム設定（ID, JPN10, 産次, 最終分娩日, DIM, 繁殖区分, 群）
        self.tree.heading('cow_id', text='管理番号')
        self.tree.heading('jpn10', text='個体識別番号')
        self.tree.heading('lact', text='産次')
        self.tree.heading('clvd', text='最終分娩日')
        self.tree.heading('dim', text='分娩後日数DIM')
        self.tree.heading('rc', text='繁殖区分')
        self.tree.heading('pen', text='群')
        
        self.tree.column('cow_id', width=100)
        self.tree.column('jpn10', width=120)
        self.tree.column('lact', width=60)
        self.tree.column('clvd', width=100)
        self.tree.column('dim', width=100)
        self.tree.column('rc', width=100)
        self.tree.column('pen', width=100)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 配置
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # タグのスタイルを設定（グレーの網掛け用）
        self.tree.tag_configure('disposed', background='#E0E0E0')
        
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
        
        # ソート状態を管理（列名 -> ソート状態: None, 'asc', 'desc'）
        self.sort_states: Dict[str, Optional[str]] = {col: None for col in columns}
        
        # 列ヘッダークリックイベントをバインド
        for col in columns:
            self.tree.heading(col, command=lambda c=col: self._on_column_click(c))
    
    def _calculate_statistics(self, cows: list) -> Dict[str, Any]:
        """
        統計情報を計算
        
        Args:
            cows: 牛のリスト
            
        Returns:
            統計情報の辞書
        """
        total_count = len(cows)
        
        # 経産牛（産次が1以上）
        lactated_cows = [cow for cow in cows if cow.get('lact') is not None and cow.get('lact', 0) >= 1]
        lactated_count = len(lactated_cows)
        
        # 妊娠牛（繁殖コードが5（Pregnant）または6（Dry））のうち経産牛
        pregnant_dry_cows = [
            cow for cow in lactated_cows
            if cow.get('rc') in [RuleEngine.RC_PREGNANT, RuleEngine.RC_DRY]
        ]
        pregnant_dry_count = len(pregnant_dry_cows)
        
        # 妊娠牛割合（経産牛に対する割合）
        pregnant_dry_ratio = (pregnant_dry_count / lactated_count * 100) if lactated_count > 0 else 0.0
        
        # 産次別の統計
        first_lactation = [cow for cow in lactated_cows if cow.get('lact') == 1]
        second_lactation = [cow for cow in lactated_cows if cow.get('lact') == 2]
        third_plus_lactation = [cow for cow in lactated_cows if cow.get('lact') is not None and cow.get('lact', 0) >= 3]
        
        first_count = len(first_lactation)
        second_count = len(second_lactation)
        third_plus_count = len(third_plus_lactation)
        
        first_ratio = (first_count / lactated_count * 100) if lactated_count > 0 else 0.0
        second_ratio = (second_count / lactated_count * 100) if lactated_count > 0 else 0.0
        third_plus_ratio = (third_plus_count / lactated_count * 100) if lactated_count > 0 else 0.0
        
        return {
            'total_count': total_count,
            'lactated_count': lactated_count,
            'pregnant_dry_count': pregnant_dry_count,
            'pregnant_dry_ratio': pregnant_dry_ratio,
            'first_count': first_count,
            'first_ratio': first_ratio,
            'second_count': second_count,
            'second_ratio': second_ratio,
            'third_plus_count': third_plus_count,
            'third_plus_ratio': third_plus_ratio
        }
    
    def _update_statistics_display(self, stats: Dict[str, Any]):
        """
        統計情報の表示を更新
        
        Args:
            stats: 統計情報の辞書
        """
        # 基本統計
        stats_text = f"総頭数: {stats['total_count']}  |  "
        stats_text += f"経産牛: {stats['lactated_count']}  |  "
        stats_text += f"経産妊娠牛割合: {stats['pregnant_dry_ratio']:.1f}% ({stats['pregnant_dry_count']})"
        
        # 産次別統計（余裕があれば）
        stats_text += f"  |  初産: {stats['first_ratio']:.1f}% ({stats['first_count']})"
        stats_text += f"  |  ２産: {stats['second_ratio']:.1f}% ({stats['second_count']})"
        stats_text += f"  |  ３産以上: {stats['third_plus_ratio']:.1f}% ({stats['third_plus_count']})"
        
        self.stats_label.config(text=stats_text)
    
    def _is_cow_disposed(self, cow_auto_id: int) -> bool:
        """
        個体が売却または死亡廃用されているかをチェック
        
        Args:
            cow_auto_id: 牛の auto_id
            
        Returns:
            売却または死亡廃用されている場合True
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for event in events:
            event_number = event.get('event_number')
            if event_number in [RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD]:
                return True
        return False
    
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
        all_cows = self.db.get_all_cows()
        
        # 現存牛のみを取得（統計情報計算用）
        existing_cows = [
            cow for cow in all_cows
            if not self._is_cow_disposed(cow['auto_id'])
        ]
        
        # 統計情報を計算（常に現存牛のみで計算）
        stats = self._calculate_statistics(existing_cows)
        self._update_statistics_display(stats)
        
        # 検索フィルタリング
        if search_text:
            search_text = search_text.strip().lower()
            cows = [
                cow for cow in all_cows
                if search_text in str(cow.get('cow_id', '')).lower()
                or search_text in str(cow.get('jpn10', '')).lower()
            ]
        else:
            cows = all_cows
        
        # 「現存牛のみ」フィルタリング
        if self.show_existing_only_var.get():
            cows = [
                cow for cow in cows
                if not self._is_cow_disposed(cow['auto_id'])
            ]
        
        # ソートを適用
        cows = self._apply_sort(cows)
        
        # 繁殖コード表示：コード番号：日本語のみ
        rc_names = {
            RuleEngine.RC_STOPPED: "1：繁殖停止",
            RuleEngine.RC_FRESH: "2：分娩後",
            RuleEngine.RC_BRED: "3：授精後",
            RuleEngine.RC_OPEN: "4：空胎",
            RuleEngine.RC_PREGNANT: "5：妊娠中",
            RuleEngine.RC_DRY: "6：乾乳中"
        }
        
        today = date.today()
        
        # Treeviewに追加
        for cow in cows:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            lact = cow.get('lact', '')
            clvd = cow.get('clvd', '') or ''
            rc = cow.get('rc', '')
            pen = cow.get('pen', '')
            
            # 分娩後日数DIMを計算（最終分娩日から今日までの日数）
            dim_display = ''
            if clvd:
                try:
                    clvd_date = datetime.strptime(clvd[:10], '%Y-%m-%d').date()
                    dim_days = (today - clvd_date).days
                    dim_display = str(dim_days) if dim_days >= 0 else ''
                except (ValueError, TypeError):
                    pass
            
            # 繁殖コードを表示名に変換
            rc_text = rc_names.get(rc, f'{rc}：' if rc is not None else '') if rc is not None else ''
            
            # 売却・死亡廃用されているかチェック
            is_disposed = self._is_cow_disposed(cow['auto_id'])
            
            # タグを設定（グレーの網掛け用）
            # auto_idを文字列としてタグに含める（後で個体を特定するため）
            tags = (str(cow['auto_id']),)
            if is_disposed:
                tags = (str(cow['auto_id']), 'disposed')
            
            item = self.tree.insert('', 'end', values=(
                cow_id, jpn10, lact, clvd, dim_display, rc_text, pen
            ), tags=tags)
    
    def _on_filter_changed(self):
        """フィルター変更時の処理"""
        self._load_cows(self.search_var.get())
    
    def _on_column_click(self, column: str):
        """
        列ヘッダークリック時の処理
        
        Args:
            column: クリックされた列名
        """
        # 現在のソート状態を取得
        current_state = self.sort_states.get(column, None)
        
        # ソート状態を切り替え: None -> 'asc' -> 'desc' -> None
        if current_state is None:
            # 他の列のソートをクリア
            for col in self.sort_states:
                self.sort_states[col] = None
                # ヘッダーの表示をリセット
                self._update_column_heading(col)
            
            # 昇順に設定
            self.sort_states[column] = 'asc'
        elif current_state == 'asc':
            # 降順に設定
            self.sort_states[column] = 'desc'
        else:  # 'desc'
            # 元に戻す（ソートなし）
            self.sort_states[column] = None
        
        # ヘッダーの表示を更新
        self._update_column_heading(column)
        
        # データを再読み込み（ソート適用）
        self._load_cows(self.search_var.get())
    
    def _update_column_heading(self, column: str):
        """
        列ヘッダーの表示を更新（ソート状態を示す矢印を追加）
        
        Args:
            column: 列名
        """
        # 列名のマッピング
        column_names = {
            'cow_id': '管理番号',
            'jpn10': '個体識別番号',
            'lact': '産次',
            'clvd': '最終分娩日',
            'dim': '分娩後日数DIM',
            'rc': '繁殖区分',
            'pen': '群'
        }
        
        base_name = column_names.get(column, column)
        sort_state = self.sort_states.get(column, None)
        
        if sort_state == 'asc':
            self.tree.heading(column, text=f"{base_name} ▲")
        elif sort_state == 'desc':
            self.tree.heading(column, text=f"{base_name} ▼")
        else:
            self.tree.heading(column, text=base_name)
    
    def _apply_sort(self, cows: list) -> list:
        """
        ソートを適用
        
        Args:
            cows: 牛のリスト
            
        Returns:
            ソートされた牛のリスト
        """
        # ソート対象の列を探す
        sort_column = None
        sort_order = None
        
        for col, state in self.sort_states.items():
            if state is not None:
                sort_column = col
                sort_order = state
                break
        
        # ソート対象がない場合はそのまま返す
        if sort_column is None:
            return cows
        
        # ソート用のキー関数を定義
        def get_sort_key(cow):
            value = cow.get(sort_column, '')
            
            # データ型に応じたソート
            if sort_column == 'lact':
                # 産次は数値としてソート
                try:
                    return int(value) if value is not None and value != '' else 0
                except (ValueError, TypeError):
                    return 0
            elif sort_column == 'clvd':
                # 日付は文字列としてソート（YYYY-MM-DD形式）
                return str(value) if value else ''
            elif sort_column == 'dim':
                # 分娩後日数DIMは数値としてソート（表示用に計算）
                clvd = cow.get('clvd', '') or ''
                if not clvd:
                    return -1  # 分娩日なしは最後に
                try:
                    clvd_date = datetime.strptime(clvd[:10], '%Y-%m-%d').date()
                    dim_days = (date.today() - clvd_date).days
                    return dim_days if dim_days >= 0 else -1
                except (ValueError, TypeError):
                    return -1
            elif sort_column == 'rc':
                # 繁殖区分は数値としてソート（表示名ではなく元の値）
                return int(value) if value is not None and value != '' else 0
            else:
                # その他は文字列としてソート
                return str(value) if value is not None else ''
        
        # ソートを実行
        sorted_cows = sorted(cows, key=get_sort_key, reverse=(sort_order == 'desc'))
        
        return sorted_cows
    
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
            # タグからauto_idを取得（最初のタグがauto_idの文字列）
            # 'disposed'タグが含まれている可能性があるので、数値に変換できるものを探す
            for tag in tags:
                try:
                    cow_auto_id = int(tag)
                    self._open_cow_card(cow_auto_id)
                    break
                except (ValueError, TypeError):
                    continue
    
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
                # タグからauto_idを取得（最初のタグがauto_idの文字列）
                # 'disposed'タグが含まれている可能性があるので、数値に変換できるものを探す
                for tag in tags:
                    try:
                        self.selected_cow_auto_id = int(tag)
                        break
                    except (ValueError, TypeError):
                        continue
                
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
                    window.window.destroy()
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
            window.window.lift()
            window.window.focus_set()
            return
        
        # 新しい個体カードウィンドウを作成（CowCardWindowを使用）
        # parentはルートウィンドウを取得（画面の高さを正しく取得するため）
        root_window = self.window.winfo_toplevel()
        cow_card_window = CowCardWindow(
            parent=root_window,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=self.item_dict_path,
            cow_auto_id=cow_auto_id
        )
        
        # ウィンドウを管理
        self.open_cow_windows[cow_auto_id] = cow_card_window
        
        # ウィンドウが閉じられたときに管理から削除
        def on_window_close():
            if cow_auto_id in self.open_cow_windows:
                del self.open_cow_windows[cow_auto_id]
            cow_card_window.window.destroy()
        
        cow_card_window.window.protocol("WM_DELETE_WINDOW", on_window_close)
        
        # ウィンドウを表示
        cow_card_window.show()
    
    def show(self):
        """ウィンドウを前面に表示（ブロックせずに返す）"""
        self.window.lift()
        self.window.focus_set()














