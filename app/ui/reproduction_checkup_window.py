"""
FALCON2 - 繁殖検診ウインドウ
繁殖検診表を表示するウインドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
import json
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.reproduction_checkup_logic import ReproductionCheckupLogic
from settings_manager import SettingsManager


class ReproductionCheckupWindow:
    """繁殖検診ウインドウ"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 formula_engine: FormulaEngine,
                 farm_path: Path,
                 event_dictionary_path: Optional[Path] = None,
                 item_dictionary_path: Optional[Path] = None,
                 rule_engine: Optional[RuleEngine] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            farm_path: 農場フォルダのパス
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス
        """
        self.parent = parent
        self.db = db_handler
        self.formula_engine = formula_engine
        self.farm_path = Path(farm_path)
        self.event_dict_path = event_dictionary_path
        # item_dictionary.json のパス（FALCON側で一括管理）
        if item_dictionary_path:
            self.item_dict_path = item_dictionary_path
        else:
            app_root = Path(__file__).parent.parent.parent
            self.item_dict_path = app_root / "config_default" / "item_dictionary.json"
        self.settings_manager = SettingsManager(self.farm_path)
        self.checkup_logic = ReproductionCheckupLogic(db_handler, event_dictionary_path)
        # RuleEngine（個体カード表示用）
        self.rule_engine = rule_engine or RuleEngine(db_handler)
        
        # 項目辞書を読み込む
        self.item_dictionary: Dict[str, Any] = {}
        self._load_item_dictionary()
        
        # ウィンドウ作成（他ウィンドウと同一デザイン）
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖検診表")
        self.window.geometry("1200x700")
        self.window.configure(bg="#f5f5f5")
        
        # 抽出結果
        self.results: List[Dict[str, Any]] = []
        # 抽出時の元の順序を保持
        self.original_results_order: List[Dict[str, Any]] = []
        
        # 並び替え状態管理
        self.sort_state = {}  # {column: 'asc'|'desc'|None}
        
        # チェック状態管理（cow_auto_id -> checked）
        self.checked_cows: Dict[int, bool] = {}  # デフォルトはTrue（チェック済み）
        
        # 追加項目（動的に追加される）
        self.additional_items: Dict[str, str] = {}  # {column_id: item_code}
        
        # 列幅の設定（カラム名: 文字数）
        self.column_widths_settings: Dict[str, int] = {}
        # 保存された列幅設定を読み込む
        self._load_column_widths_settings()
        # 追加項目のインデックスマッピング（プルダウンとの対応）
        self.additional_item_index_map: Dict[int, str] = {}  # {index: column_id}
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _load_item_dictionary(self):
        """item_dictionary.jsonを読み込む"""
        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                print(f"項目辞書読み込みエラー: {e}")
                self.item_dictionary = {}
        else:
            # デフォルトパスを試す
            default_path = Path("docs/item_dictionary.json")
            if default_path.exists():
                try:
                    with open(default_path, 'r', encoding='utf-8') as f:
                        self.item_dictionary = json.load(f)
                except Exception as e:
                    print(f"項目辞書読み込みエラー: {e}")
                    self.item_dictionary = {}
    
    def _load_column_widths_settings(self):
        """保存された列幅設定を読み込む"""
        self.column_widths_settings = self.settings_manager.get('repro_checkup_column_widths', {})
    
    def _save_column_widths_settings(self):
        """列幅設定を保存"""
        self.settings_manager.set('repro_checkup_column_widths', self.column_widths_settings)
    
    def _create_widgets(self):
        """ウィジェットを作成（個体一覧・個体カードなどと同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        
        main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        main_container.pack(fill=tk.BOTH, expand=True)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # ヘッダー（他ウィンドウと同一イメージ）
        header = tk.Frame(main_container, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\u2695", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="繁殖検診表", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="検診予定日を選び、抽出して印刷プレビューできます", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)
        
        # コンテンツ（コントロール＋表）
        content_frame = tk.Frame(main_container, bg=bg)
        content_frame.grid(row=1, column=0, sticky=tk.NSEW)
        content_frame.grid_columnconfigure(0, weight=1)
        
        main_frame = ttk.Frame(content_frame)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 検診予定日入力フレーム（左詰め）
        date_frame = ttk.Frame(main_frame)
        date_frame.pack(anchor=tk.W, pady=(0, 8))
        
        ttk.Label(date_frame, text="検診予定日:").pack(side=tk.LEFT, padx=(0, 3))
        
        # デフォルトは今日
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_var = tk.StringVar(value=today)
        date_entry = ttk.Entry(date_frame, textvariable=self.date_var, width=12)
        date_entry.pack(side=tk.LEFT, padx=(0, 3))
        
        # カレンダーボタン
        calendar_btn = ttk.Button(
            date_frame,
            text="📅",
            command=self._on_calendar_click,
            width=3
        )
        calendar_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        extract_btn = ttk.Button(
            date_frame,
            text="抽出",
            command=self._on_extract,
            width=8
        )
        extract_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 設定ボタン
        settings_btn = ttk.Button(
            date_frame,
            text="繁殖検診設定",
            command=self._on_settings,
            width=12
        )
        settings_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 印刷プレビューボタン
        print_preview_btn = ttk.Button(
            date_frame,
            text="印刷プレビュー",
            command=self._on_print_preview,
            width=12
        )
        print_preview_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 追加項目選択フレーム（左詰め）
        additional_frame = ttk.Frame(main_frame)
        additional_frame.pack(anchor=tk.W, pady=(0, 8))
        
        ttk.Label(additional_frame, text="追加項目:").pack(side=tk.LEFT, padx=(0, 3))
        
        # 追加項目選択（3つ）
        self.additional_item_vars = []
        self.additional_item_combos = []
        self.additional_item_buttons = []
        for i in range(3):
            # 項目フレーム（プルダウンとボタンをまとめる）
            item_frame = ttk.Frame(additional_frame)
            item_frame.pack(side=tk.LEFT, padx=(0, 5))
            
            var = tk.StringVar(value='')
            self.additional_item_vars.append(var)
            combo = ttk.Combobox(
                item_frame,
                textvariable=var,
                width=18,
                state='readonly'
            )
            combo.pack(side=tk.LEFT, padx=(0, 2))
            # 項目リストを設定（後で更新される）
            combo['values'] = []
            combo.bind('<<ComboboxSelected>>', lambda e, idx=i: self._on_additional_item_changed(idx))
            combo.bind('<Button-3>', lambda e, idx=i: self._on_combo_right_click(e, idx))
            self.additional_item_combos.append(combo)
            
            # カテゴリー別選択ボタン
            select_btn = ttk.Button(
                item_frame,
                text="📋",
                width=3,
                command=lambda idx=i: self._show_category_item_dialog(idx)
            )
            select_btn.pack(side=tk.LEFT)
            self.additional_item_buttons.append(select_btn)
        
        # プルダウンのリストを初期化
        self._update_all_combo_lists()
        
        # 統計表示フレーム（左詰め）
        stats_frame = tk.Frame(main_frame, bg=bg)
        stats_frame.pack(anchor=tk.W, pady=(0, 8))
        
        self.stats_label = tk.Label(
            stats_frame,
            text="抽出頭数：0",
            font=(_df, 9),
            bg=bg,
            fg="#607d8b"
        )
        self.stats_label.pack(side=tk.LEFT, padx=0)
        
        # 結果表示フレーム（表を残りスペースで表示）
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 0))
        
        # カラム定義（順序変更可能）
        self.column_order = [
            'checked',  # チェックボックスカラム（左端）
            'cow_id', 'jpn10', 'lact', 'dim_or_age', 'last_checkup_date',
            'last_checkup_result', 'last_ai_date', 'dai', 'checkup_code'
        ]
        # 追加項目カラム（動的に追加）
        self.additional_columns = []  # 追加された項目コードのリスト
        
        # カラム名と表示名のマッピング
        self.column_labels = {
            'checked': '✓',  # チェックボックスカラム
            'cow_id': '個体ID',
            'jpn10': '個体識別番号',
            'lact': '産次',
            'dim_or_age': '分娩後日数/月齢',
            'last_checkup_date': '前回検診日',
            'last_checkup_result': '前回検診結果',
            'last_ai_date': '最終AI日/ET日',
            'dai': '授精後日数',
            'checkup_code': '検診コード'
        }
        # 日付は必ず見える幅を確保。前回検診結果は控えめにして表全体がウィンドウ内に収まるように
        self._default_column_widths = {
            'last_checkup_date': 120,   # 日付 YYYY-MM-DD が切れない幅
            'last_checkup_result': 200, # 所見欄（ウインドウ内に全列収まるよう調整）
            'last_ai_date': 120,        # 日付 YYYY-MM-DD が切れない幅
            'jpn10': 110,
            'dai': 72,
            'checkup_code': 72,
            'cow_id': 72,
            'lact': 44,
            'dim_or_age': 100,
        }
        
        # Treeview用スタイル（Meiryo UI）
        try:
            style = ttk.Style()
            style.configure("ReproCheckup.Treeview", font=(_df, 10))
            style.configure("ReproCheckup.Treeview.Heading", font=(_df, 10, "bold"))
        except tk.TclError:
            pass
        
        # Treeviewを作成（高さを調整、コンパクトに）
        all_columns = self.column_order + self.additional_columns
        self.tree = ttk.Treeview(result_frame, columns=all_columns, show='headings', height=20, style="ReproCheckup.Treeview")
        
        # カラムヘッダーを設定
        for col in all_columns:
            if col in self.column_labels:
                label = self.column_labels[col]
            elif col in self.additional_items:
                item_code = self.additional_items[col]
                item_data = self.item_dictionary.get(item_code, {})
                label = item_data.get('display_name', item_code)
            else:
                label = col
            # チェックボックスカラムはクリックでソートしない
            if col == 'checked':
                self.tree.heading(col, text=label)
            else:
                self.tree.heading(col, text=label, command=lambda c=col: self._on_column_click(c))
            # 初期幅（保存設定優先、次に日付・長文用デフォルト、それ以外100）
            saved_chars = self.column_widths_settings.get(col)
            if saved_chars:
                pixel_width = max(saved_chars * 10 + 20, self._default_column_widths.get(col, 100))
                self.tree.column(col, width=pixel_width)
            else:
                if col == 'checked':
                    self.tree.column(col, width=40)
                else:
                    default_w = self._default_column_widths.get(col, 100)
                    self.tree.column(col, width=default_w)
        
        # 縦・横スクロールバー（全列がウィンドウ内に収まらない場合は横スクロールで表示）
        self.scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.hscroll = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.scrollbar.set, xscrollcommand=self.hscroll.set)
        
        # 配置（上段: ツリー＋縦スクロール、下段: 横スクロール）
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # ダブルクリックで個体カードを開く
        self.tree.bind('<Double-Button-1>', self._on_cow_double_click)
        
        # チェックボックスのクリックイベント（シングルクリック）
        self.tree.bind('<Button-1>', self._on_tree_click)
        
        # 右クリックメニュー
        self.tree.bind('<Button-3>', self._on_tree_right_click)
        
        # result_frameを保存（後で使用）
        self.result_frame = result_frame
    
    def _on_extract(self):
        """抽出ボタンをクリック"""
        # 検診予定日を取得
        checkup_date_str = self.date_var.get().strip()
        if not checkup_date_str:
            messagebox.showerror("エラー", "検診予定日を入力してください。")
            return
        
        # 日付形式を検証
        try:
            datetime.strptime(checkup_date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("エラー", "検診予定日の形式が正しくありません（YYYY-MM-DD形式）。")
            return
        
        # 設定を読み込む
        settings = self.settings_manager.get('repro_checkup_settings', {})
        if not settings:
            messagebox.showwarning("警告", "繁殖検診設定が未設定です。設定を確認してください。")
            return
        
        # 抽出実行
        try:
            self.results = self.checkup_logic.extract_cows(checkup_date_str, settings)
            # 元の順序を保存（ディープコピー）
            self.original_results_order = [cow.copy() for cow in self.results]
            # 並び替え状態をリセット
            self.sort_state = {}
            # チェック状態をリセット（全員チェック済み）
            self.checked_cows = {}
            for cow in self.results:
                cow_auto_id = cow.get('auto_id')
                if cow_auto_id:
                    self.checked_cows[cow_auto_id] = True
        except Exception as e:
            messagebox.showerror("エラー", f"抽出中にエラーが発生しました: {e}")
            import traceback
            traceback.print_exc()
            return
        
        # 結果を表示
        self._display_results()
        self._update_stats()
    
    def _update_stats(self):
        """統計情報を更新"""
        if not self.results:
            self.stats_label.config(text="抽出頭数：0")
            return
        
        # 検診コード別の集計
        stats = {}
        for cow in self.results:
            code = cow.get('checkup_code', '')
            if code:
                stats[code] = stats.get(code, 0) + 1
        
        # 統計文字列を作成
        total = len(self.results)
        stats_text = f"抽出頭数：{total}"
        
        # 検診コードの順序（定義順）
        code_order = [
            "ﾌﾚｯｼｭ", "繁殖1", "繁殖2", "妊鑑", "再妊", "妊鑑2",
            "分娩？", "チェック", "育繁1", "育繁2", "育妊", "育再妊"
        ]
        
        for code in code_order:
            if code in stats:
                stats_text += f"　{code}：{stats[code]}"
        
        # その他の検診コード
        for code, count in sorted(stats.items()):
            if code not in code_order:
                stats_text += f"　{code}：{count}"
        
        self.stats_label.config(text=stats_text)
    
    def _format_value(self, value):
        """値をフォーマット（Noneや空文字列は「ー」に変換）"""
        if value is None or value == '':
            return 'ー'
        return str(value)
    
    def _auto_resize_columns(self):
        """カラム幅を内容に応じて自動調整（保存された設定がない場合のみ）"""
        all_columns = self.column_order + self.additional_columns
        for col in all_columns:
            # チェックボックスカラムは固定幅なのでスキップ
            if col == 'checked':
                continue
            # 保存された設定があればそれを使用
            if col in self.column_widths_settings:
                continue
            
            # ヘッダーテキストの幅を計算
            header_text = self.tree.heading(col, 'text')
            # 日本語文字を考慮（全角文字は約2倍の幅）
            header_width = sum(14 if ord(c) > 127 else 7 for c in header_text) + 20
            
            # データの最大幅を計算
            max_width = header_width
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                col_index = all_columns.index(col)
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    # 日本語文字を考慮した幅計算（全角文字は約14ピクセル、半角は約7ピクセル）
                    text_width = sum(14 if ord(c) > 127 else 7 for c in cell_text) + 20
                    max_width = max(max_width, text_width)
            
            # 日付は最小幅を確保。前回検診結果は最大幅を抑えて表がウィンドウに収まるように
            min_w = getattr(self, '_default_column_widths', {}).get(col, 60)
            max_w = 450
            if col == 'last_checkup_result':
                max_w = 240  # 所見欄は広げすぎない
            self.tree.column(col, width=max(min_w, min(max_width, max_w)))
    
    def _refresh_tree_columns(self):
        """Treeviewのカラムを更新（追加項目が変更された場合）"""
        # 現在のカラムリスト
        current_columns = list(self.tree['columns'])
        new_columns = self.column_order + self.additional_columns
        
        # カラムが変更された場合は再作成が必要
        if current_columns != new_columns:
            # Treeviewを再作成
            old_tree = self.tree
            old_scrollbar = self.scrollbar
            
            # 新しいTreeviewを作成（スタイル・デフォルト幅を維持）
            self.tree = ttk.Treeview(self.result_frame, columns=new_columns, show='headings', height=20, style="ReproCheckup.Treeview")
            
            # カラムヘッダーを設定
            for col in new_columns:
                if col in self.column_labels:
                    label = self.column_labels[col]
                elif col in self.additional_items:
                    item_code = self.additional_items[col]
                    item_data = self.item_dictionary.get(item_code, {})
                    label = item_data.get('display_name', item_code)
                else:
                    label = col
                if col == 'checked':
                    self.tree.heading(col, text=label)
                else:
                    self.tree.heading(col, text=label, command=lambda c=col: self._on_column_click(c))
                saved_chars = self.column_widths_settings.get(col)
                if saved_chars:
                    pixel_width = max(saved_chars * 10 + 20, getattr(self, '_default_column_widths', {}).get(col, 100))
                    self.tree.column(col, width=pixel_width)
                else:
                    if col == 'checked':
                        self.tree.column(col, width=40)
                    else:
                        default_w = getattr(self, '_default_column_widths', {}).get(col, 100)
                        self.tree.column(col, width=default_w)
            
            # スクロールバーを再設定（縦・横）
            self.scrollbar = ttk.Scrollbar(self.result_frame, orient=tk.VERTICAL, command=self.tree.yview)
            self.tree.configure(yscrollcommand=self.scrollbar.set, xscrollcommand=self.hscroll.set)
            
            # イベントバインディング
            self.tree.bind('<Double-Button-1>', self._on_cow_double_click)
            self.tree.bind('<Button-1>', self._on_tree_click)
            self.tree.bind('<Button-3>', self._on_tree_right_click)
            
            # 古いTreeviewを削除して新しいものを配置
            old_tree.destroy()
            old_scrollbar.destroy()
            self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        else:
            # カラムが同じ場合はヘッダーのみ更新
            for col in self.additional_columns:
                if col in self.additional_items:
                    item_code = self.additional_items[col]
                    item_data = self.item_dictionary.get(item_code, {})
                    label = item_data.get('display_name', item_code)
                    self.tree.heading(col, text=label)
    
    def _display_results(self):
        """結果を表示"""
        # カラムを更新（追加項目がある場合）
        self._refresh_tree_columns()
        
        
        # 既存のアイテムをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 各牛の詳細情報を計算して表示
        checkup_date_str = self.date_var.get().strip()
        for cow in self.results:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # イベント履歴を取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            # 当該産次のイベントのみ使用（産次更新後は前回検診・最終AI・授精後日数をリセット）
            events_this_lact = self.checkup_logic.filter_events_current_lactation(events, cow)
            
            # 計算項目を取得（DIMは検診予定日起算、AGE等は従来通り）
            calculated = self.formula_engine.calculate(cow_auto_id)
            
            # 各項目を取得
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            lact = cow.get('lact') or 0
            
            # 分娩後日数/月齢（経産牛は検診予定日時点のDIM）
            dim_or_age = ''
            if lact >= 1:
                dim = self.formula_engine.calculate_dim_at_date(cow_auto_id, checkup_date_str) if checkup_date_str else calculated.get('DIM')
                if dim is not None:
                    dim_or_age = str(dim)
            else:
                age = calculated.get('AGE')
                if age is not None:
                    dim_or_age = f"{age:.1f}"
            
            # 前回検診日・結果（当該産次のイベントのみ）
            try:
                last_checkup_date = self.formula_engine._get_last_reproduction_check_date(events_this_lact) or ''
                last_checkup_result = self.formula_engine._get_last_reproduction_check_note(events_this_lact) or ''
            except Exception:
                last_checkup_date = ''
                last_checkup_result = ''
            
            # 最終AI日/ET日（当該産次のイベントのみ）
            last_ai_date = self.checkup_logic._get_last_ai_date(events_this_lact, checkup_date_str) or ''
            
            # 授精後日数（当該産次のイベントのみ、検診予定日を基準として計算）
            dai = None
            if checkup_date_str:
                try:
                    dai = self.checkup_logic._calculate_dai(events_this_lact, checkup_date_str)
                except Exception as e:
                    logging.warning(f"DAI計算エラー (cow_auto_id={cow_auto_id}): {e}")
            dai_str = str(dai) if dai is not None else ''
            
            # 検診コード（抽出時に当該産次で判定済み）
            checkup_code = cow.get('checkup_code', '')
            
            # チェック状態を取得（デフォルトはTrue）
            is_checked = self.checked_cows.get(cow_auto_id, True)
            checked_display = '☑' if is_checked else '☐'
            
            # 基本項目の値を取得（チェックボックスを最初に追加）
            values = [
                checked_display,  # チェックボックス
                self._format_value(cow_id),
                self._format_value(jpn10),
                str(lact) if lact is not None else 'ー',
                self._format_value(dim_or_age),
                self._format_value(last_checkup_date),
                self._format_value(last_checkup_result),
                self._format_value(last_ai_date),
                self._format_value(dai_str),
                self._format_value(checkup_code)
            ]
            
            # 追加項目の値を取得
            for col_id in self.additional_columns:
                item_code = self.additional_items.get(col_id, '')
                if item_code and item_code in calculated:
                    value = calculated.get(item_code)
                    values.append(self._format_value(value))
                else:
                    values.append('ー')
            
            # Treeviewにアイテムを追加（cow_auto_idをtagsに保存）
            item_id = self.tree.insert('', 'end', values=tuple(values), tags=(str(cow_auto_id),))
        
        # カラム幅を自動調整
        self._auto_resize_columns()
    
    def _on_column_click(self, column: str):
        """カラムヘッダーをクリック"""
        if not self.results:
            return
        
        # チェックボックスカラムはソートしない
        if column == 'checked':
            return
        
        # カラム名をソートキーに変換
        column_to_key = {
            'cow_id': 'cow_id',
            'jpn10': 'jpn10',
            'lact': 'lact',
            'dim_or_age': 'dim_or_age',
            'last_checkup_date': 'last_checkup_date',
            'last_checkup_result': 'last_checkup_result',
            'last_ai_date': 'last_ai_date',
            'dai': 'dai',
            'checkup_code': 'checkup_code'
        }
        
        sort_key = column_to_key.get(column)
        # 追加項目の場合は、項目コードをソートキーとして使用
        if not sort_key and column in self.additional_items:
            sort_key = self.additional_items[column]
        
        if not sort_key:
            return
        
        # 追加項目の場合、ソートキーは項目コードとして扱う
        is_additional_item = (column in self.additional_items)
        
        # 現在の状態を取得
        current_state = self.sort_state.get(column, None)
        
        # 状態を切り替え：None -> 'asc' -> 'desc' -> None
        if current_state is None:
            new_state = 'asc'
        elif current_state == 'asc':
            new_state = 'desc'
        else:  # 'desc'
            new_state = None
        
        self.sort_state[column] = new_state
        
        # 元に戻す場合は元の順序に戻す（抽出時の順序）
        if new_state is None:
            # 抽出時の順序に戻す
            if self.original_results_order:
                # 元の順序を保持するために、auto_idでマッピング
                order_map = {cow.get('auto_id'): i for i, cow in enumerate(self.original_results_order)}
                self.results.sort(key=lambda c: order_map.get(c.get('auto_id'), 999999))
        else:
            # 並び替え実行
            reverse = (new_state == 'desc')
            
            def sort_key_func(cow):
                cow_auto_id = cow.get('auto_id')
                if not cow_auto_id:
                    return ''
                
                # 計算が必要な項目は計算して取得
                if sort_key == 'dim_or_age':
                    lact = cow.get('lact') or 0
                    if lact >= 1:
                        checkup_date_str = self.date_var.get().strip()
                        dim = self.formula_engine.calculate_dim_at_date(cow_auto_id, checkup_date_str) if checkup_date_str else None
                        if dim is None:
                            calculated = self.formula_engine.calculate(cow_auto_id)
                            dim = calculated.get('DIM')
                        return dim if dim is not None else 0
                    else:
                        calculated = self.formula_engine.calculate(cow_auto_id)
                        age = calculated.get('AGE')
                        return age if age is not None else 0.0
                elif sort_key == 'dai':
                    # 検診予定日を基準としてDAIを計算
                    checkup_date_str = self.date_var.get().strip()
                    if checkup_date_str:
                        try:
                            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                            dai = self.formula_engine._calculate_dai_at_event_date(events, checkup_date_str)
                            return dai if dai is not None else 0
                        except Exception:
                            return 0
                    else:
                        # 検診予定日が設定されていない場合は0を返す
                        return 0
                elif sort_key == 'last_checkup_date':
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    date_str = calculated.get('LREPRO', '')
                    return date_str if date_str else ''
                elif sort_key == 'last_checkup_result':
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    result = calculated.get('REPCT', '')
                    return result if result else ''
                elif sort_key == 'last_ai_date':
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    date_str = calculated.get('LASTAI', '')
                    return date_str if date_str else ''
                elif is_additional_item or sort_key in self.item_dictionary:
                    # 追加項目の場合
                    calculated = self.formula_engine.calculate(cow_auto_id)
                    value = calculated.get(sort_key, '')
                    # データ型に応じて適切な型で返す
                    item_data = self.item_dictionary.get(sort_key, {})
                    data_type = item_data.get('data_type', 'str')
                    if data_type in ('int', 'integer'):
                        try:
                            return int(value) if value else 0
                        except (ValueError, TypeError):
                            return 0
                    elif data_type in ('float', 'decimal'):
                        try:
                            return float(value) if value else 0.0
                        except (ValueError, TypeError):
                            return 0.0
                    return str(value) if value else ''
                else:
                    value = cow.get(sort_key, '')
                    # 数値として比較可能な場合は数値として比較
                    if sort_key in ('lact', 'cow_id'):
                        try:
                            return int(value) if value else 0
                        except (ValueError, TypeError):
                            return 0
                    return value
            
            self.results.sort(key=sort_key_func, reverse=reverse)
        
        # 再表示
        self._display_results()
    
    def _on_calendar_click(self):
        """カレンダーボタンをクリック"""
        self._show_calendar()
    
    def _show_calendar(self):
        """カレンダーウィンドウを表示（統一コンポーネントを使用）"""
        from ui.date_picker_window import DatePickerWindow
        
        # 現在の日付を初期値として使用
        initial_date = None
        try:
            datetime.strptime(self.date_var.get(), '%Y-%m-%d')
            initial_date = self.date_var.get()
        except:
            pass
    
        def on_date_selected(date_str: str):
            """日付選択時のコールバック"""
            self.date_var.set(date_str)
        
        date_picker = DatePickerWindow(
            parent=self.window,
            initial_date=initial_date,
            on_date_selected=on_date_selected
        )
        date_picker.show()
    
    def _on_settings(self):
        """設定ボタンをクリック"""
        from ui.reproduction_checkup_settings_window import ReproductionCheckupSettingsWindow
        
        settings_window = ReproductionCheckupSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        settings_window.show()
    
    def _on_tree_click(self, event):
        """Treeviewのクリック処理（チェックボックスの切り替え）"""
        # クリックされた位置を取得
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        
        # クリックされたカラムを取得
        column = self.tree.identify_column(event.x)
        if not column:
            return
        
        # カラム番号（#1, #2など）からカラム名を取得
        col_index = int(column.replace('#', '')) - 1
        all_columns = self.column_order + self.additional_columns
        if col_index < 0 or col_index >= len(all_columns):
            return
        
        col_name = all_columns[col_index]
        
        # チェックボックスカラムの場合のみ処理
        if col_name != 'checked':
            return
        
        # クリックされた行を取得
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        # タグからcow_auto_idを取得
        tags = self.tree.item(item, 'tags')
        if not tags:
            return
        
        try:
            cow_auto_id = int(tags[0])
        except (ValueError, TypeError):
            return
        
        # チェック状態を切り替え
        current_checked = self.checked_cows.get(cow_auto_id, True)
        self.checked_cows[cow_auto_id] = not current_checked
        
        # 表示を更新
        values = list(self.tree.item(item, 'values'))
        if values:
            values[0] = '☑' if not current_checked else '☐'
            self.tree.item(item, values=tuple(values))
    
    def _on_cow_double_click(self, event):
        """個体のダブルクリック処理"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = self.tree.item(selection[0])
        values = item['values']
        if not values:
            return
        
        # チェックボックスカラムが追加されたため、インデックスを調整
        cow_id = values[1]  # 0番目はチェックボックス、1番目が個体ID
        
        # 個体を検索
        cows = self.db.get_cows_by_id(cow_id)
        if not cows:
            return
        
        # 個体カードウィンドウを開く
        from ui.cow_card_window import CowCardWindow
        
        cow = cows[0]
        cow_auto_id = cow.get('auto_id')
        if cow_auto_id:
            cow_card_window = CowCardWindow(
                parent=self.window,
                db_handler=self.db,
                formula_engine=self.formula_engine,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path,
                item_dictionary_path=None,
                cow_auto_id=cow_auto_id
            )
            cow_card_window.show()
    
    def _get_available_items(self) -> List[tuple]:
        """利用可能な項目のリストを取得（既に追加されている項目を除く）"""
        items = []
        added_codes = set(self.additional_items.values())
        for item_code, item_data in self.item_dictionary.items():
            if item_code not in added_codes:
                display_name = item_data.get('display_name', item_code)
                items.append((item_code, f"{item_code} - {display_name}"))
        return sorted(items, key=lambda x: x[1])
    
    def _get_available_items_for_combo(self) -> List[str]:
        """プルダウン用の項目リストを取得（空文字列を含む）"""
        items = ['']  # 空文字列（選択なし）
        added_codes = set(self.additional_items.values())
        for item_code, item_data in self.item_dictionary.items():
            if item_code not in added_codes:
                display_name = item_data.get('display_name', item_code)
                items.append(f"{item_code} - {display_name}")
        return sorted(items)
    
    def _on_tree_right_click(self, event):
        """Treeviewの右クリック処理"""
        # クリック位置からカラムを判定
        region = self.tree.identify_region(event.x, event.y)
        if region == 'heading':
            # カラムヘッダーが右クリックされた
            column = self.tree.identify_column(event.x)
            if column:
                # カラム番号（#1, #2など）からカラム名を取得
                col_index = int(column.replace('#', '')) - 1
                all_columns = self.column_order + self.additional_columns
                if 0 <= col_index < len(all_columns):
                    col_name = all_columns[col_index]
                    self._show_column_menu(event.x_root, event.y_root, col_name)
        elif region == 'cell':
            # セルが右クリックされた（項目追加用）
            self._show_add_item_menu(event.x_root, event.y_root)
    
    def _show_column_menu(self, x, y, column_name):
        """カラムヘッダーの右クリックメニューを表示"""
        menu = tk.Menu(self.window, tearoff=0)
        
        # チェックボックスカラムの場合はメニューを表示しない
        if column_name == 'checked':
            return
        
        # 列幅を設定
        menu.add_command(
            label="列幅を設定",
            command=lambda: self._set_column_width(column_name)
        )
        
        # 列幅を自動調整に戻す
        if column_name in self.column_widths_settings:
            menu.add_command(
                label="列幅を自動調整に戻す",
                command=lambda: self._reset_column_width(column_name)
            )
        
        menu.add_separator()
        
        # 左に移動
        if column_name in self.column_order:
            current_index = self.column_order.index(column_name)
            if current_index > 0:
                menu.add_command(
                    label="左に移動",
                    command=lambda: self._move_column_left(column_name)
                )
        
        # 右に移動
        if column_name in self.column_order:
            current_index = self.column_order.index(column_name)
            if current_index < len(self.column_order) - 1:
                menu.add_command(
                    label="右に移動",
                    command=lambda: self._move_column_right(column_name)
                )
        
        # 追加項目の場合は削除オプション
        if column_name in self.additional_columns:
            menu.add_separator()
            menu.add_command(
                label="項目を削除",
                command=lambda: self._remove_additional_column(column_name)
            )
        
        menu.post(x, y)
    
    def _show_add_item_menu(self, x, y):
        """項目追加メニューを表示"""
        if len(self.additional_columns) >= 3:
            messagebox.showinfo("情報", "追加項目は最大3つまでです。")
            return
        
        menu = tk.Menu(self.window, tearoff=0)
        # 空いているプルダウンのインデックスを探す
        empty_index = None
        for i in range(3):
            if i not in self.additional_item_index_map:
                empty_index = i
                break
        
        if empty_index is not None:
            menu.add_command(
                label=f"項目を追加（追加項目{empty_index + 1}）",
                command=lambda idx=empty_index: self._show_add_item_dialog_for_index(idx)
            )
        else:
            # すべて埋まっている場合は最後のものを置き換え
            menu.add_command(
                label="項目を追加（追加項目3を置き換え）",
                command=lambda idx=2: self._show_add_item_dialog_for_index(idx)
            )
        menu.post(x, y)
    
    def _add_additional_column(self, item_code: str, index: Optional[int] = None):
        """追加項目カラムを追加
        
        Args:
            item_code: 項目コード
            index: プルダウンのインデックス（0-2）。Noneの場合は自動割り当て
        """
        # 既に追加されている場合は何もしない
        if item_code in self.additional_items.values():
            return
        
        # インデックスが指定されていない場合は、空いているインデックスを探す
        if index is None:
            for i in range(3):
                if i not in self.additional_item_index_map:
                    index = i
                    break
            if index is None:
                # すべて埋まっている場合は最後のものを置き換え
                index = 2
        
        # 既にそのインデックスに項目がある場合は削除
        if index in self.additional_item_index_map:
            old_col_id = self.additional_item_index_map[index]
            if old_col_id in self.additional_columns:
                self.additional_columns.remove(old_col_id)
            if old_col_id in self.additional_items:
                del self.additional_items[old_col_id]
        
        # 新しいカラムIDを生成
        col_id = f"additional_{index + 1}"
        self.additional_columns.append(col_id)
        self.additional_items[col_id] = item_code
        self.additional_item_index_map[index] = col_id
        
        # プルダウンの値を更新
        display_name = self.item_dictionary.get(item_code, {}).get('display_name', item_code)
        self.additional_item_vars[index].set(f"{item_code} - {display_name}")
        
        # すべてのプルダウンのリストを更新
        self._update_all_combo_lists()
        
        # 表示を更新
        if self.results:
            self._display_results()
    
    def _update_all_combo_lists(self):
        """すべてのプルダウンの項目リストを更新"""
        items_list = self._get_available_items_for_combo()
        for combo in self.additional_item_combos:
            combo['values'] = items_list
    
    def _on_additional_item_changed(self, index: int):
        """追加項目が変更された（プルダウン選択時）"""
        var = self.additional_item_vars[index]
        selected = var.get()
        
        if not selected or selected == '':
            # 選択がクリアされた場合は項目を削除
            if index in self.additional_item_index_map:
                col_id = self.additional_item_index_map[index]
                self._remove_additional_column(col_id, index)
        else:
            # "CODE - 表示名" から CODE を抽出
            item_code = selected.split(' - ')[0]
            self._add_additional_column(item_code, index)
    
    def _on_combo_right_click(self, event, index: int):
        """プルダウンの右クリック処理"""
        menu = tk.Menu(self.window, tearoff=0)
        menu.add_command(
            label="項目を選択",
            command=lambda: self._show_add_item_dialog_for_index(index)
        )
        menu.post(event.x_root, event.y_root)
    
    def _show_add_item_dialog_for_index(self, index: int):
        """特定のインデックス用の項目追加ダイアログを表示（簡易版）"""
        self._show_category_item_dialog(index)
    
    def _show_category_item_dialog(self, index: int):
        """カテゴリー別項目選択ダイアログを表示"""
        # 利用可能な項目を取得（カテゴリー情報付き）
        selectable_items = self._get_available_items_with_category()
        if not selectable_items:
            messagebox.showinfo("情報", "追加可能な項目がありません。")
            return
        
        # ダイアログウィンドウ
        dialog = tk.Toplevel(self.window)
        dialog.title("項目を選択（カテゴリー別）")
        dialog.geometry("550x600")
        
        # カテゴリー名のマッピング
        category_names = {
            "CORE": "基本情報",
            "REPRODUCTION": "繁殖",
            "DHI": "乳検",
            "GENOMIC": "ゲノム",
            "HEALTH": "疾病",
            "OTHERS": "その他",
            "USER": "ユーザー"
        }
        
        # メインフレーム
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 検索フレーム
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT, padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Treeview（カテゴリー別表示）
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        tree = ttk.Treeview(tree_frame, columns=('code',), show='tree headings', height=20, yscrollcommand=scrollbar.set)
        scrollbar.config(command=tree.yview)
        tree.heading('#0', text='項目名')
        tree.heading('code', text='ｺｰﾄﾞ')
        tree.column('#0', width=350)
        tree.column('code', width=150)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # カテゴリー別に項目を整理
        category_items: Dict[str, List[tuple]] = {}
        for code, name, category in selectable_items:
            cat = category or "OTHERS"
            if cat not in category_items:
                category_items[cat] = []
            category_items[cat].append((code, name))
        
        # カテゴリーごとにソート
        for cat in category_items:
            category_items[cat].sort(key=lambda x: x[1])
        
        # カテゴリー順序（表示順）
        category_order = ["CORE", "REPRODUCTION", "DHI", "GENOMIC", "HEALTH", "USER", "OTHERS"]
        
        # Treeviewに項目を追加
        category_nodes: Dict[str, str] = {}
        all_items: List[tuple] = []  # (code, name, category, node_id)
        
        for cat in category_order:
            if cat in category_items and category_items[cat]:
                cat_name = category_names.get(cat, cat)
                cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                category_nodes[cat] = cat_node
                
                for code, name in category_items[cat]:
                    item_node = tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                    all_items.append((code, name, cat, item_node))
        
        # カテゴリータグのスタイル設定
        tree.tag_configure('category', font=('', 9, 'bold'), background='#E0E0E0')
        tree.tag_configure('item', font=('', 9))
        
        # 初期状態でカテゴリーを展開
        for cat_node in category_nodes.values():
            tree.item(cat_node, open=True)
        
        selected_code: Optional[str] = None
        
        def filter_tree(*args):
            """検索文字列でフィルタリング"""
            nonlocal all_items
            search_text = search_var.get().lower()
            
            # すべての項目を削除
            for item in tree.get_children():
                tree.delete(item)
            
            category_nodes.clear()
            all_items = []
            
            # フィルタリング後の項目
            filtered_category_items: Dict[str, List[tuple]] = {}
            
            if search_text:
                # 検索文字列がある場合
                for cat in category_order:
                    if cat in category_items:
                        filtered = [
                            (code, name) for code, name in category_items[cat]
                            if search_text in code.lower() or search_text in name.lower()
                        ]
                        if filtered:
                            filtered_category_items[cat] = filtered
            else:
                # 検索文字列がない場合はすべて表示
                filtered_category_items = category_items
            
            # Treeviewに項目を再追加
            for cat in category_order:
                if cat in filtered_category_items and filtered_category_items[cat]:
                    cat_name = category_names.get(cat, cat)
                    cat_node = tree.insert('', 'end', text=cat_name, values=('',), tags=('category',))
                    category_nodes[cat] = cat_node
                    
                    for code, name in filtered_category_items[cat]:
                        item_node = tree.insert(cat_node, 'end', text=name, values=(code,), tags=('item',))
                        all_items.append((code, name, cat, item_node))
            
            # 展開
            for cat_node in category_nodes.values():
                tree.item(cat_node, open=True)
        
        # 検索エントリの変更を監視
        search_var.trace_add('write', filter_tree)
        
        # ダブルクリックで選択
        def on_double_click(event):
            item = tree.selection()[0] if tree.selection() else None
            if item:
                tags = tree.item(item, 'tags')
                if 'item' in tags:
                    values = tree.item(item, 'values')
                    if values:
                        selected_code = values[0]
                        self._add_additional_column(selected_code, index)
                        dialog.destroy()
        
        tree.bind('<Double-Button-1>', on_double_click)
        
        # ボタンフレーム
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_ok():
            selection = tree.selection()
            if selection:
                item = selection[0]
                tags = tree.item(item, 'tags')
                if 'item' in tags:
                    values = tree.item(item, 'values')
                    if values:
                        selected_code = values[0]
                        self._add_additional_column(selected_code, index)
                        dialog.destroy()
                else:
                    messagebox.showinfo("情報", "項目を選択してください。")
            else:
                messagebox.showinfo("情報", "項目を選択してください。")
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="選択", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # 検索エントリにフォーカス
        search_entry.focus()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
    
    def _get_available_items_with_category(self) -> List[tuple]:
        """利用可能な項目のリストを取得（カテゴリー情報付き）"""
        items = []
        added_codes = set(self.additional_items.values())
        for item_code, item_data in self.item_dictionary.items():
            if item_code not in added_codes:
                display_name = item_data.get('display_name', item_code)
                category = (item_data.get('category') or 'OTHERS').upper()
                items.append((item_code, display_name, category))
        return sorted(items, key=lambda x: (x[2], x[1]))  # カテゴリー、名前の順でソート
    
    def _remove_additional_column(self, column_name: str, index: Optional[int] = None):
        """追加項目カラムを削除
        
        Args:
            column_name: カラムID
            index: プルダウンのインデックス（指定された場合のみプルダウンを更新）
        """
        if column_name in self.additional_columns:
            self.additional_columns.remove(column_name)
            if column_name in self.additional_items:
                del self.additional_items[column_name]
            
            # インデックスマッピングから削除
            if index is not None:
                if index in self.additional_item_index_map:
                    del self.additional_item_index_map[index]
                # プルダウンの値をクリア
                self.additional_item_vars[index].set('')
            else:
                # インデックスが見つからない場合は検索
                for idx, col_id in list(self.additional_item_index_map.items()):
                    if col_id == column_name:
                        del self.additional_item_index_map[idx]
                        self.additional_item_vars[idx].set('')
                        break
            
            # すべてのプルダウンのリストを更新
            self._update_all_combo_lists()
            
            # 表示を更新
            if self.results:
                self._display_results()
    
    def _move_column_left(self, column_name: str):
        """カラムを左に移動"""
        if column_name in self.column_order:
            index = self.column_order.index(column_name)
            if index > 0:
                self.column_order[index], self.column_order[index - 1] = \
                    self.column_order[index - 1], self.column_order[index]
                if self.results:
                    self._display_results()
    
    def _move_column_right(self, column_name: str):
        """カラムを右に移動"""
        if column_name in self.column_order:
            index = self.column_order.index(column_name)
            if index < len(self.column_order) - 1:
                self.column_order[index], self.column_order[index + 1] = \
                    self.column_order[index + 1], self.column_order[index]
                if self.results:
                    self._display_results()
    
    def _set_column_width(self, column_name: str):
        """列幅を設定（文字数ベース）"""
        # 現在の設定を取得（文字数、なければピクセルから推測）
        current_chars = self.column_widths_settings.get(column_name)
        if current_chars is None:
            # ピクセル幅から文字数を推測（約10px/文字）
            current_pixel_width = self.tree.column(column_name, 'width')
            current_chars = max(1, (current_pixel_width - 20) // 10)
        
        # ダイアログウィンドウ
        dialog = tk.Toplevel(self.window)
        dialog.title("列幅を設定")
        dialog.resizable(False, False)  # リサイズ不可
        
        # カラム名を取得
        column_label = self.tree.heading(column_name, 'text')
        
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            frame,
            text=f"「{column_label}」の列幅を設定してください"
        ).pack(pady=10)
        
        ttk.Label(
            frame,
            text="文字数で指定します（例: 3 → 3文字分）",
            font=("", 8),
            foreground="gray"
        ).pack(pady=2)
        
        width_var = tk.IntVar(value=current_chars)
        
        width_frame = ttk.Frame(frame)
        width_frame.pack(pady=10)
        
        ttk.Label(width_frame, text="幅 (文字数):").pack(side=tk.LEFT, padx=5)
        
        width_spinbox = ttk.Spinbox(
            width_frame,
            from_=1,
            to=50,
            textvariable=width_var,
            width=10
        )
        width_spinbox.pack(side=tk.LEFT, padx=5)
        
        def on_ok():
            new_chars = width_var.get()
            # 文字数をピクセルに変換して設定
            pixel_width = new_chars * 10 + 20
            self.tree.column(column_name, width=pixel_width)
            # 設定を保存（文字数で保存）
            self.column_widths_settings[column_name] = new_chars
            self._save_column_widths_settings()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text="設定", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # ウィンドウサイズを自動調整してから固定
        dialog.update_idletasks()
        dialog.geometry(f"{dialog.winfo_reqwidth()}x{dialog.winfo_reqheight()}")
        
        # ウィンドウを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _reset_column_width(self, column_name: str):
        """列幅を自動調整に戻す"""
        # 設定から削除
        if column_name in self.column_widths_settings:
            del self.column_widths_settings[column_name]
            self._save_column_widths_settings()
        
        # 自動調整を実行
        if self.results:
            # 一時的に設定を無視して自動調整
            header_text = self.tree.heading(column_name, 'text')
            header_width = sum(14 if ord(c) > 127 else 7 for c in header_text) + 20
            
            max_width = header_width
            all_columns = self.column_order + self.additional_columns
            col_index = all_columns.index(column_name)
            
            for item in self.tree.get_children():
                values = self.tree.item(item, 'values')
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    text_width = sum(14 if ord(c) > 127 else 7 for c in cell_text) + 20
                    max_width = max(max_width, text_width)
            cap = 240 if column_name == 'last_checkup_result' else 400
            self.tree.column(column_name, width=max(60, min(max_width, cap)))
    
    def _on_print_preview(self):
        """印刷プレビューボタンをクリック"""
        if not self.results:
            messagebox.showwarning("警告", "抽出結果がありません。")
            return
        
        # チェックが入っている個体のみをフィルタリング
        filtered_results = []
        for cow in self.results:
            cow_auto_id = cow.get('auto_id')
            if cow_auto_id and self.checked_cows.get(cow_auto_id, True):
                filtered_results.append(cow)
        
        if not filtered_results:
            messagebox.showwarning("警告", "印刷対象の個体がありません。チェックボックスを確認してください。")
            return
        
        # HTML方式でブラウザに表示
        from ui.reproduction_checkup_print_preview import open_html_print_preview
        
        # 農場名を取得
        from settings_manager import SettingsManager
        settings_manager = SettingsManager(self.farm_path)
        farm_name = settings_manager.get('farm_name', self.farm_path.name)
        
        # チェックボックスカラムを除外したcolumn_orderを渡す
        print_column_order = [col for col in self.column_order if col != 'checked']
        
        open_html_print_preview(
            results=filtered_results,
            checkup_date=self.date_var.get(),
            formula_engine=self.formula_engine,
            item_dictionary=self.item_dictionary,
            additional_items=self.additional_items,
            column_order=print_column_order,
            column_widths_settings=self.column_widths_settings,
            farm_name=farm_name,
            checkup_logic=self.checkup_logic
        )
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

