"""
FALCON2 - MainWindow（メインウィンドウ）
アプリ全体のフレーム（左：サイドメニュー、右：メイン表示領域）
設計書 第11章・第13章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from typing import Optional, Literal, List, Dict, Any
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from ui.cow_card import CowCard
from ui.event_input import EventInputWindow


class MainWindow:
    """メインウィンドウ（サイドメニュー + メイン表示領域）"""
    
    def __init__(self, root: tk.Tk, db_handler: DBHandler, 
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 farm_path: Path):
        """
        初期化
        
        Args:
            root: Tkinter ルートウィンドウ
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            farm_path: 農場フォルダのパス
        """
        self.root = root
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.farm_path = farm_path
        
        # event_dictionary.json のパス（農場フォルダ側を優先）
        # 同期処理で作成されるため、農場フォルダ側のみを参照
        self.event_dict_path = farm_path / "event_dictionary.json"
        if not self.event_dict_path.exists():
            # 同期処理で作成されるはずだが、念のため警告
            print(f"警告: event_dictionary.json が見つかりません: {self.event_dict_path}")
            self.event_dict_path = None
        
        # item_dictionary.json のパス（農場フォルダ側を優先）
        # 同期処理で作成されるため、農場フォルダ側のみを参照
        self.item_dict_path = farm_path / "item_dictionary.json"
        if not self.item_dict_path.exists():
            # 同期処理で作成されるはずだが、念のため警告
            print(f"警告: item_dictionary.json が見つかりません: {self.item_dict_path}")
            self.item_dict_path = None
        
        # 現在選択中の牛
        self.current_cow_auto_id: Optional[int] = None
        self.current_cow_card: Optional[CowCard] = None
        
        # 現在表示中のView
        self.current_view: Optional[tk.Widget] = None
        self.current_view_type: Optional[Literal['chat', 'cow_card', 'list', 'report']] = None
        
        # ビュー管理（再利用のため保持）
        self.views: Dict[str, tk.Widget] = {}
        
        # ウィンドウサイズを設定
        self.root.title("FALCON2")
        self.root.geometry("1200x800")
        
        # UI作成
        self._create_widgets()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインPanedWindow（左右分割）
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左カラム：サイドメニュー ==========
        menu_frame = ttk.Frame(main_paned, width=200)
        main_paned.add(menu_frame, weight=0)
        
        # メニューボタン
        menu_buttons = [
            ("個体カード", self._on_cow_card),
            ("繁殖検診", self._on_reproduction_checkup),
            ("イベント入力", self._on_event_input),
            ("データ出力", self._on_data_output),
            ("辞書設定", self._on_dictionary_settings),
            ("農場管理", self._on_farm_management),
        ]
        
        for text, command in menu_buttons:
            btn = ttk.Button(menu_frame, text=text, command=command, width=20)
            btn.pack(pady=5, padx=10)
        
        # ========== 右カラム：メイン表示領域 ==========
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        # 右上：コマンド入力欄
        command_frame = ttk.Frame(right_frame)
        command_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(command_frame, text="コマンド:").pack(side=tk.LEFT, padx=5)
        self.command_entry = ttk.Entry(command_frame, width=50)
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.command_entry.bind('<Return>', self._on_command_enter)
        
        command_btn = ttk.Button(command_frame, text="実行", command=self._on_command_execute)
        command_btn.pack(side=tk.LEFT, padx=5)
        
        # コマンド入力欄の直下：最終日付表示
        latest_dates_frame = ttk.Frame(right_frame)
        latest_dates_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        self.latest_dates_label = ttk.Label(
            latest_dates_frame,
            text="最終分娩：—　最終AI：—　最終乳検：—　最終イベント：—",
            foreground="gray"
        )
        self.latest_dates_label.pack(side=tk.LEFT, padx=5)
        
        # 右側：メイン表示領域（ChatGPT / View切替）
        self.main_content_frame = ttk.Frame(right_frame)
        self.main_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 初期表示：ChatGPT画面
        self._show_chat_view()
        
        # 群全体の最終日付を計算して表示
        self._calculate_and_update_farm_latest_dates()
    
    def show_view(self, view_widget: tk.Widget, view_type: Literal['chat', 'cow_card', 'list', 'report']):
        """
        Viewを表示（切替機構）
        
        Args:
            view_widget: 表示するウィジェット
            view_type: Viewの種類
        """
        # デバッグログ：親フレームの存在確認
        try:
            parent_exists = self.main_content_frame.winfo_exists()
            logging.debug(f"show_view: parent exists={parent_exists}, view_type={view_type}")
        except tk.TclError as e:
            logging.error(f"show_view: parent check failed: {e}")
            return
        
        # 既存のビューを pack_forget() で非表示にする（destroy しない）
        for v in self.views.values():
            try:
                if v.winfo_exists():
                    v.pack_forget()
            except tk.TclError:
                # 既に破棄されている場合は無視
                pass
        
        # 新しいViewを配置
        try:
            view_widget.pack(fill=tk.BOTH, expand=True)
        except tk.TclError as e:
            logging.error(f"show_view: pack failed: {e}, view_type={view_type}")
            return
        
        # ビューを辞書に保存（再利用のため）
        self.views[view_type] = view_widget
        
        # 現在のViewを記録
        self.current_view = view_widget
        self.current_view_type = view_type
    
    def _show_chat_view(self):
        """ChatGPT画面を表示"""
        # 既存のchat_viewがあれば再利用、なければ新規作成
        if 'chat' in self.views:
            chat_frame = self.views['chat']
        else:
            # 新しいchat_frameを作成（parent は main_content_frame）
            chat_frame = ttk.Frame(self.main_content_frame)
            
            # チャット履歴表示エリア
            chat_history = scrolledtext.ScrolledText(
                chat_frame,
                wrap=tk.WORD,
                width=80,
                height=30,
                state=tk.DISABLED
            )
            chat_history.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # チャット履歴を保存（後でAI機能で使用）
            self.chat_history = chat_history
        
        # View切替（show_view 内で既存のViewを pack_forget）
        self.show_view(chat_frame, 'chat')
        
        # 初期メッセージ（初回のみ表示）
        if not hasattr(self, '_chat_initialized'):
            self._add_chat_message("システム", "FALCON2 にようこそ。コマンドを入力するか、メニューから操作を選択してください。")
            self._chat_initialized = True
    
    def _on_cow_card(self):
        """個体カードメニューをクリック（個体一覧ウィンドウを開く）"""
        from ui.cow_list_window import CowListWindow
        
        # 個体一覧ウィンドウを開く
        list_window = CowListWindow(
            parent=self.root,
            db_handler=self.db,
            formula_engine=self.formula_engine,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            item_dictionary_path=self.item_dict_path
        )
        list_window.show()
    
    def _show_cow_card_view(self, cow_auto_id: int):
        """
        個体カードViewを表示
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 既存のCowCardがあれば再利用、なければ新規作成
        if 'cow_card' in self.views and self.current_cow_card:
            cow_card = self.current_cow_card
            cow_card_widget = cow_card.get_widget()
        else:
            # CowCard を作成（parent は必ず main_content_frame）
            cow_card = CowCard(
                parent=self.main_content_frame,
                db_handler=self.db,
                formula_engine=self.formula_engine,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path
            )
            cow_card_widget = cow_card.get_widget()
            # イベント保存時のコールバックを設定
            cow_card.set_on_event_saved(self._on_event_saved)
            # 現在の牛を記録
            self.current_cow_card = cow_card
        
        # 牛の情報を読み込んで表示（既存のCowCardでも再読み込み）
        cow_card.load_cow(cow_auto_id)
        
        # View切替（show_view 内で既存のViewを pack_forget）
        self.show_view(cow_card_widget, 'cow_card')
        
        # 現在の牛を記録
        self.current_cow_auto_id = cow_auto_id
        
        # 群全体の最終日付を再計算して表示を更新（個体カード表示時も更新）
        self._calculate_and_update_farm_latest_dates()
    
    def _on_reproduction_checkup(self):
        """繁殖検診メニューをクリック"""
        print("繁殖検診メニューをクリック")
        # TODO: 繁殖検診画面を実装
    
    def _on_event_input(self):
        """
        イベント入力メニューをクリック
        
        【重要】左メニューからのイベント入力は常に汎用モード（ID入力から開始）
        メイン画面の状態（個体カード表示有無）に依存しない
        """
        # event_dictionary_path が None の場合はエラー
        if self.event_dict_path is None:
            messagebox.showerror(
                "エラー",
                "event_dictionary.json が見つかりません"
            )
            return
        
        # EventInputWindow を生成して表示
        # 左メニューからの起動は常に cow_auto_id=None で汎用モード
        # 必ずID入力欄を表示状態で起動する
        event_input_window = EventInputWindow(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 常にNone（汎用モード、ID入力から開始）
            on_saved=self._on_event_saved,
            farm_path=self.farm_path
        )
        event_input_window.show()
    
    def _on_data_output(self):
        """データ出力メニューをクリック"""
        print("データ出力メニューをクリック")
        # TODO: データ出力画面を実装
    
    def _on_dictionary_settings(self):
        """辞書設定メニューをクリック"""
        # 辞書設定選択ダイアログを表示
        dialog = tk.Toplevel(self.root)
        dialog.title("辞書・設定")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # タイトル
        title_label = ttk.Label(
            dialog,
            text="辞書・設定",
            font=("", 12, "bold")
        )
        title_label.pack(pady=20)
        
        # 説明
        desc_label = ttk.Label(
            dialog,
            text="操作を選択してください",
            font=("", 9)
        )
        desc_label.pack(pady=(0, 20))
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        # イベント辞書ボタン
        event_dict_btn = ttk.Button(
            button_frame,
            text="イベント辞書",
            command=lambda: self._on_event_dictionary(dialog),
            width=20
        )
        event_dict_btn.pack(pady=5, padx=20)
        
        # 項目辞書ボタン
        item_dict_btn = ttk.Button(
            button_frame,
            text="項目辞書",
            command=lambda: self._on_item_dictionary(dialog),
            width=20
        )
        item_dict_btn.pack(pady=5, padx=20)
        
        # 農場設定ボタン
        farm_settings_btn = ttk.Button(
            button_frame,
            text="農場設定",
            command=lambda: self._on_farm_settings(dialog),
            width=20
        )
        farm_settings_btn.pack(pady=5, padx=20)
        
        # キャンセルボタン
        cancel_btn = ttk.Button(
            button_frame,
            text="キャンセル",
            command=dialog.destroy,
            width=20
        )
        cancel_btn.pack(pady=10, padx=20)
        
        # ウィンドウを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _on_event_dictionary(self, parent_dialog):
        """イベント辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        from ui.event_dictionary_window import EventDictionaryWindow
        
        event_dict_window = EventDictionaryWindow(
            parent=self.root,
            event_dictionary_path=self.event_dict_path
        )
        event_dict_window.show()
    
    def _on_item_dictionary(self, parent_dialog):
        """項目辞書ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        if not self.item_dict_path:
            messagebox.showerror("エラー", "item_dictionary.json のパスが設定されていません")
            return
        
        from ui.item_dictionary_window import ItemDictionaryWindow
        
        item_window = ItemDictionaryWindow(
            parent=self.root,
            item_dictionary_path=self.item_dict_path,
            on_item_updated=self._on_item_dictionary_changed,
            formula_engine=self.formula_engine,
        )
        item_window.show()

    def _on_item_dictionary_changed(self):
        """項目辞書更新後のハンドラ（FormulaEngineを再読込し、表示を更新）"""
        try:
            if self.formula_engine:
                self.formula_engine.reload_item_dictionary()
        except Exception as e:
            logging.error(f"item_dictionary reload failed: {e}")
        
        # CowCardを再描画
        if self.current_view_type == 'cow_card' and self.current_cow_card and self.current_cow_auto_id:
            try:
                self.current_cow_card.load_cow(self.current_cow_auto_id)
                self._add_chat_message("システム", "項目辞書の更新を反映しました。")
            except Exception as e:
                logging.error(f"CowCard refresh failed after item_dictionary update: {e}")
    
    def _on_farm_settings(self, parent_dialog):
        """農場設定ボタンをクリック"""
        parent_dialog.destroy()  # 選択ダイアログを閉じる
        
        from ui.farm_settings_window import FarmSettingsWindow
        
        farm_settings_window = FarmSettingsWindow(
            parent=self.root,
            farm_path=self.farm_path
        )
        farm_settings_window.show()
    
    def _on_farm_management(self):
        """農場管理メニューをクリック"""
        print("農場管理メニューをクリック")
        # TODO: 農場管理画面を実装
    
    def _on_event_saved(self, cow_auto_id: int):
        """
        イベント保存後に現在のViewを更新するコールバック
        
        Args:
            cow_auto_id: イベントが保存された牛の auto_id
        """
        # 現在CowCardが表示されている場合は更新
        if self.current_view_type == 'cow_card' and self.current_cow_card:
            if self.current_cow_auto_id == cow_auto_id:
                self.current_cow_card.load_cow(cow_auto_id)
                self._add_chat_message("システム", f"個体カードを更新しました: {cow_auto_id}")
        else:
            # 他のViewが表示されている場合はチャットにメッセージを追加
            self._add_chat_message("システム", f"イベントを保存しました: 個体ID {cow_auto_id}")
        
        # 群全体の最終日付を再計算して表示を更新
        self._calculate_and_update_farm_latest_dates()
    
    def _calculate_and_update_farm_latest_dates(self):
        """
        群全体（農場全体）の最終日付を計算して表示を更新
        """
        try:
            # 農場全体のイベントを取得
            events = self.db.get_all_events(include_deleted=False)
            
            # 最終分娩日を計算（EVENT_CALV = 202、baselineも含む）
            latest_calving = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_CALV])
            
            # 最終AI日を計算（AI/ET系イベント: 200, 201）
            latest_ai = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET])
            
            # 最終乳検日を計算（EVENT_MILK_TEST = 601）
            latest_milk = CowCard.get_latest_event_date(events, [RuleEngine.EVENT_MILK_TEST])
            
            # 最終イベント日を計算（全イベント対象）
            latest_any = CowCard.get_latest_any_event_date(events)
            
            # 日付を表示用文字列に変換（Noneの場合は"—"）
            calving_str = latest_calving if latest_calving else "—"
            ai_str = latest_ai if latest_ai else "—"
            milk_str = latest_milk if latest_milk else "—"
            any_str = latest_any if latest_any else "—"
            
            # 表示文字列を作成
            display_text = f"最終分娩：{calving_str}　最終AI：{ai_str}　最終乳検：{milk_str}　最終イベント：{any_str}"
            
            # Labelを更新
            self.latest_dates_label.config(text=display_text)
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _calculate_and_update_farm_latest_dates で例外が発生しました: {e}")
            traceback.print_exc()
            # エラー時は"—"を表示
            display_text = "最終分娩：—　最終AI：—　最終乳検：—　最終イベント：—"
            self.latest_dates_label.config(text=display_text)
    
    def _on_command_enter(self, event):
        """コマンド入力欄でEnterキー押下時"""
        self._on_command_execute()
    
    def _on_command_execute(self):
        """コマンド実行"""
        raw_input = self.command_entry.get().strip()
        if not raw_input:
            return
        
        # デバッグ補助
        print("COMMAND:", raw_input)
        
        # コマンドをチャット履歴に追加
        self._add_chat_message("ユーザー", raw_input)
        
        # コマンドをクリア
        self.command_entry.delete(0, tk.END)
        
        # 数字のみの場合は個体ID検索として扱う
        if raw_input.isdigit():
            # 4桁にゼロパディング（例: 980 → 0980）
            padded_id = raw_input.zfill(4)
            
            print("PADDED ID:", padded_id)
            
            # 個体カードへジャンプ
            self._jump_to_cow_card(padded_id)
            return
        
        # 空白や演算子が含まれる場合は従来のコマンド解析ロジックに渡す
        if ' ' in raw_input or any(op in raw_input for op in ['>', '<', '=', '!']):
            # TODO: コマンド解析と実行（AI機能など）
            self._add_chat_message("システム", f"コマンド '{raw_input}' を実行しました（実装予定）")
            return
        
        # その他の場合は従来のコマンド解析
        # TODO: コマンド解析と実行（AI機能など）
        self._add_chat_message("システム", f"コマンド '{raw_input}' を実行しました（実装予定）")
    
    def _jump_to_cow_card(self, cow_id: str):
        """
        個体カードへジャンプ
        
        Args:
            cow_id: 4桁の牛ID（例: "0980"）
        """
        # 4桁IDで検索（複数件取得可能）
        cows = self.db.get_cows_by_id(cow_id)
        
        if not cows:
            # 見つからない場合のUX
            messagebox.showinfo(
                "検索結果",
                f"ID {cow_id} の個体は見つかりませんでした"
            )
            self._add_chat_message("システム", f"ID {cow_id} の個体は見つかりませんでした")
            return
        
        # 1件のみの場合はそのまま表示
        if len(cows) == 1:
            cow = cows[0]
            cow_auto_id = cow.get('auto_id')
            if cow_auto_id:
                self._show_cow_card_view(cow_auto_id)
                self._add_chat_message("システム", f"個体カードを表示しました: {cow.get('cow_id', cow_id)}")
                # チャット履歴に個体情報とイベント履歴を表示（色付き）
                self._display_cow_info_in_chat(cow_auto_id)
            return
        
        # 複数件見つかった場合は選択ダイアログを表示
        selected_cow = self._show_cow_selection_dialog(cows, cow_id)
        if selected_cow:
            cow_auto_id = selected_cow.get('auto_id')
            if cow_auto_id:
                self._show_cow_card_view(cow_auto_id)
                self._add_chat_message("システム", f"個体カードを表示しました: {selected_cow.get('cow_id', cow_id)} (個体識別番号: {selected_cow.get('jpn10', '')})")
                # チャット履歴に個体情報とイベント履歴を表示（色付き）
                self._display_cow_info_in_chat(cow_auto_id)
    
    def _show_cow_selection_dialog(self, cows: List[Dict[str, Any]], cow_id: str) -> Optional[Dict[str, Any]]:
        """
        4桁IDが重複している場合の個体選択ダイアログ
        
        Args:
            cows: 検索結果の牛リスト
            cow_id: 検索した4桁ID
        
        Returns:
            選択された牛の情報、キャンセルの場合はNone
        """
        # 選択ダイアログウィンドウを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("個体選択")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # メッセージ
        message = f"拡大4桁ID {cow_id} に該当する個体が複数見つかりました。\n個体識別番号から選択してください。"
        ttk.Label(dialog, text=message, wraplength=450).pack(pady=10)
        
        # リストボックス
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # スクロールバー付きリストボックス
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Courier", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 選択結果を保持
        selected_cow = None
        
        # リストボックスに個体情報を追加
        for cow in cows:
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            bthd = cow.get('bthd', '')
            pen = cow.get('pen', '')
            
            # 表示形式: 個体識別番号 | 品種 | 生年月日 | 群
            display_text = f"{jpn10:12s} | {brd:10s} | {bthd:10s} | {pen:10s}"
            listbox.insert(tk.END, display_text)
        
        # リストボックスの選択イベント
        def on_select(event):
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
        
        listbox.bind('<<ListboxSelect>>', on_select)
        listbox.bind('<Double-Button-1>', lambda e: dialog.destroy())
        
        # ヘッダー行を追加（読みやすさのため）
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill=tk.X, padx=10)
        header_text = f"{'個体識別番号':12s} | {'品種':10s} | {'生年月日':10s} | {'群':10s}"
        ttk.Label(header_frame, text=header_text, font=("Courier", 10, "bold")).pack()
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "個体を選択してください。")
        
        def on_cancel():
            nonlocal selected_cow
            selected_cow = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="選択", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # モーダルダイアログとして表示
        dialog.wait_window()
        
        return selected_cow
    
    def _add_chat_message(self, sender: str, message: str, color: Optional[str] = None):
        """
        チャットメッセージを追加
        
        Args:
            sender: 送信者名
            message: メッセージ
            color: テキストの色（#RRGGBB形式、Noneの場合はデフォルト）
        """
        # chat_history が存在し、有効な場合のみ追加
        if hasattr(self, 'chat_history') and self.chat_history is not None:
            try:
                self.chat_history.config(state=tk.NORMAL)
                
                # 色が指定されている場合はタグを使用
                if color:
                    tag_name = f"color_{color}"
                    # タグが存在しない場合は作成
                    if tag_name not in self.chat_history.tag_names():
                        self.chat_history.tag_configure(tag_name, foreground=color)
                    
                    # メッセージを挿入してタグを適用
                    start_pos = self.chat_history.index(tk.END)
                    self.chat_history.insert(tk.END, f"[{sender}] {message}\n")
                    end_pos = self.chat_history.index(tk.END + "-1c")
                    self.chat_history.tag_add(tag_name, start_pos, end_pos)
                else:
                    # 色が指定されていない場合は通常通り
                    self.chat_history.insert(tk.END, f"[{sender}] {message}\n")
                
                self.chat_history.see(tk.END)
                self.chat_history.config(state=tk.DISABLED)
            except tk.TclError:
                # chat_history が破壊されている場合は無視
                pass
        print(f"[{sender}] {message}")
    
    def _get_event_display_color(self, event_number: int) -> str:
        """
        イベントの表示色を決定（CowCardと同じロジック）
        
        Args:
            event_number: イベント番号
        
        Returns:
            色（#RRGGBB形式）
        """
        # event_dictionary.json を読み込む
        if not hasattr(self, '_event_dictionary') or self._event_dictionary is None:
            self._event_dictionary = {}
            if self.event_dict_path and self.event_dict_path.exists():
                try:
                    with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                        self._event_dictionary = json.load(f)
                except Exception as e:
                    logging.error(f"event_dictionary.json 読み込みエラー: {e}")
        
        # イベント辞書から情報を取得
        event_str = str(event_number)
        event_dict = self._event_dictionary.get(event_str, {})
        
        # 1. display_color が指定されていればそれを使用（最優先）
        display_color = event_dict.get('display_color')
        if display_color:
            return display_color
        
        # 2. category / outcome に応じてデフォルト色を使用
        category = event_dict.get('category', '')
        
        if category == "CALVING":
            return "#0066cc"  # 青
        elif category == "PREGNANCY":
            outcome = event_dict.get('outcome', '')
            if outcome == "NEGATIVE":
                return "#cc0000"  # 赤
            else:
                return "#008000"  # 緑（POSITIVE または未指定）
        elif category == "BREEDING":
            return "#000000"  # 黒
        
        # 3. それ以外は黒
        return "#000000"
    
    def _display_cow_info_in_chat(self, cow_auto_id: int):
        """
        チャット履歴に個体情報とイベント履歴を表示（色付き）
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            # 牛の情報を取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                return
            
            # 個体情報を表示
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            lact = cow.get('lact', '')
            rc = cow.get('rc', '')
            
            self._add_chat_message("システム", f"個体情報: {cow_id} (JPN10: {jpn10}, 品種: {brd}, 産次: {lact}, RC: {rc})")
            
            # イベント履歴を取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            if not events:
                self._add_chat_message("システム", "イベント履歴: なし")
                return
            
            # イベント名を取得するためのヘルパー関数
            def get_event_name(event_number: int) -> str:
                event_str = str(event_number)
                if hasattr(self, '_event_dictionary') and self._event_dictionary:
                    event_dict = self._event_dictionary.get(event_str, {})
                    return event_dict.get('name_jp', f'イベント{event_number}')
                return f'イベント{event_number}'
            
            # イベント履歴を表示（最新から）
            self._add_chat_message("システム", "イベント履歴:")
            for event in events[:10]:  # 最新10件まで
                event_date = event.get('event_date', '')
                event_number = event.get('event_number')
                note = event.get('note', '')
                
                if event_number is None:
                    continue
                
                event_name = get_event_name(event_number)
                color = self._get_event_display_color(event_number)
                
                # イベント情報を色付きで表示
                display_text = f"  {event_date} {event_name}"
                if note:
                    display_text += f" - {note}"
                
                self._add_chat_message("システム", display_text, color=color)
            
            if len(events) > 10:
                self._add_chat_message("システム", f"  ... 他 {len(events) - 10} 件")
                
        except Exception as e:
            logging.error(f"個体情報表示エラー: {e}")
            import traceback
            traceback.print_exc()

