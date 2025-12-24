"""
FALCON2 - 分析UI（改修版）
分析種別選択、複数行コマンド入力、右パネル（参照・挿入専用）を実装
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any, Optional, List, Callable
from pathlib import Path
import json
import logging

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    logging.warning("matplotlib がインストールされていません。グラフ機能は使用できません。")

from modules.execution_plan import AnalysisType, GraphType
from modules.query_router_v2 import QueryRouterV2
from modules.executor_v2 import ExecutorV2

logger = logging.getLogger(__name__)


class AnalysisUI:
    """分析UIフレーム"""
    
    # 分析種別のガイド
    ANALYSIS_GUIDES = {
        "list": [
            "（表示したい項目：例 ID、分娩後日数、産次）",
            "（絞り込み：例 2産以上、DIM<150、妊娠牛）"
        ],
        "agg": [
            "（見たい項目：例 授精回数、初回授精日数）",
            "（分けて見る基準：例 産次、分娩月）",
            "（絞り込み：例 DIM>150）"
        ],
        "eventcount": [
            "（集計したいイベント：例 授精、分娩）",
            "（分けて見る基準：例 月、産次）",
            "（絞り込み）"
        ],
        "graph": [
            "（縦軸にする項目：例 初回授精日数）",
            "（横軸にする項目：例 分娩月）",
            "（グラフの種類：例 折れ線、棒、プロット、箱ひげ、生存曲線）",
            "（色・線を分ける基準：例 産次、DIM21）",
            "（絞り込み：例 妊娠牛、産次>1）"
        ],
        "repro": [
            "（見たい繁殖指標：例 受胎率、発情発見率、妊娠率）",
            "（分けて見る基準：例 産次、月、DIM21）",
            "（絞り込み：例 初回授精）"
        ]
    }
    
    def __init__(self, parent: tk.Widget, farm_path: Path,
                 query_router: QueryRouterV2, executor: ExecutorV2,
                 on_execute: Optional[Callable] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            farm_path: 農場フォルダパス
            query_router: QueryRouterV2 インスタンス
            executor: ExecutorV2 インスタンス
            on_execute: 実行時のコールバック
        """
        self.parent = parent
        self.farm_path = farm_path
        self.query_router = query_router
        self.executor = executor
        self.on_execute = on_execute
        
        # 辞書を読み込む
        self.item_dictionary: Dict[str, Any] = {}
        self.event_dictionary: Dict[str, Any] = {}
        self._load_dictionaries()
        
        # UI作成
        self.frame = ttk.Frame(parent)
        self._create_widgets()
    
    def _load_dictionaries(self):
        """辞書を読み込む"""
        # item_dictionary.json
        item_dict_path = self.farm_path / "item_dictionary.json"
        if item_dict_path.exists():
            try:
                with open(item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                logger.error(f"item_dictionary.json読み込みエラー: {e}")
        
        # event_dictionary.json
        event_dict_path = self.farm_path / "event_dictionary.json"
        if event_dict_path.exists():
            try:
                with open(event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                logger.error(f"event_dictionary.json読み込みエラー: {e}")
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインPanedWindow（左右分割：分析UI vs 結果表示）
        main_paned = ttk.PanedWindow(self.frame, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左側：分析UI（仮説入力）エリア ==========
        analysis_frame = ttk.Frame(main_paned)
        main_paned.add(analysis_frame, weight=1)
        
        # 分析UIの内部PanedWindow（上下分割：入力エリア vs 参照パネル）
        analysis_paned = ttk.PanedWindow(analysis_frame, orient=tk.VERTICAL)
        analysis_paned.pack(fill=tk.BOTH, expand=True)
        
        # 上：入力エリア
        input_frame = ttk.Frame(analysis_paned)
        analysis_paned.add(input_frame, weight=2)
        self._create_input_area(input_frame)
        
        # 下：参照・挿入専用パネル
        reference_frame = ttk.Frame(analysis_paned)
        analysis_paned.add(reference_frame, weight=1)
        self._create_reference_panel(reference_frame)
        
        # ========== 右側：結果表示エリア（完全に分離） ==========
        result_frame = ttk.Frame(main_paned)
        main_paned.add(result_frame, weight=1)
        self._create_result_view(result_frame)
    
    def _create_input_area(self, parent: tk.Widget):
        """入力エリアを作成"""
        # 分析種別選択
        analysis_type_frame = ttk.LabelFrame(parent, text="分析種別", padding=5)
        analysis_type_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.analysis_type_var = tk.StringVar(value="list")
        analysis_types = [
            ("リスト", "list"),
            ("集計", "agg"),
            ("イベント集計", "eventcount"),
            ("グラフ", "graph"),
            ("繁殖分析", "repro")
        ]
        
        for text, value in analysis_types:
            rb = ttk.Radiobutton(
                analysis_type_frame,
                text=text,
                variable=self.analysis_type_var,
                value=value,
                command=self._on_analysis_type_changed
            )
            rb.pack(side=tk.LEFT, padx=5)
        
        # 期間指定
        period_frame = ttk.LabelFrame(parent, text="期間", padding=5)
        period_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(period_frame, text="開始:").pack(side=tk.LEFT, padx=5)
        self.period_start_entry = ttk.Entry(period_frame, width=12)
        self.period_start_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(period_frame, text="終了:").pack(side=tk.LEFT, padx=5)
        self.period_end_entry = ttk.Entry(period_frame, width=12)
        self.period_end_entry.pack(side=tk.LEFT, padx=5)
        
        # 複数行コマンド入力欄
        command_frame = ttk.LabelFrame(parent, text="コマンド", padding=5)
        command_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Textウィジェット + Scrollbar（オーバーレイ用にFrameで囲む）
        text_container = ttk.Frame(command_frame)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        text_frame = ttk.Frame(text_container)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.command_text = tk.Text(text_frame, height=10, wrap=tk.WORD)
        self.command_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.command_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.command_text.config(yscrollcommand=lambda *args: (scrollbar.set(*args), self._update_guides()))
        
        # ガイド表示用Canvas（オーバーレイ）
        # Textウィジェットの背景色を取得
        text_bg = self.command_text.cget("bg")
        self.guide_canvas = tk.Canvas(
            text_container,
            highlightthickness=0,
            bg=text_bg,
            cursor="text",  # Textウィジェットと同じカーソル
            takefocus=False  # フォーカスを受け取らない
        )
        # CanvasをTextウィジェットの上に配置（placeで重ねる）
        self.guide_canvas.place(in_=text_frame, x=0, y=0, relwidth=1, relheight=1)
        
        # CanvasのクリックイベントをTextウィジェットに転送（透過）
        self.guide_canvas.bind('<Button-1>', self._on_canvas_click)
        self.guide_canvas.bind('<Key>', lambda e: self.command_text.focus_set())
        self.guide_canvas.bind('<Motion>', lambda e: self.command_text.focus_set())
        
        # ガイドラベルの管理
        self.guide_labels: Dict[int, int] = {}  # {line_num: canvas_text_id}
        
        # Textウィジェットのイベントバインド
        self.command_text.bind('<KeyRelease>', self._on_text_changed)
        self.command_text.bind('<Button-1>', self._on_text_changed)
        self.command_text.bind('<FocusIn>', self._on_text_changed)
        self.command_text.bind('<FocusOut>', self._on_text_changed)
        self.command_text.bind('<Configure>', self._on_text_configure)
        
        # 初期ガイド表示
        self._update_guides()
        
        # 実行ボタン
        execute_btn = ttk.Button(command_frame, text="実行", command=self._on_execute)
        execute_btn.pack(pady=5)
    
    def _create_reference_panel(self, parent: tk.Widget):
        """参照・挿入専用パネルを作成"""
        # ラベルフレーム
        ref_label_frame = ttk.LabelFrame(parent, text="参照・挿入", padding=5)
        ref_label_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 検索ボックス
        search_frame = ttk.Frame(ref_label_frame)
        search_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(search_frame, text="検索:").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.search_entry.bind('<KeyRelease>', self._on_search)
        
        # Notebook（タブ）
        notebook = ttk.Notebook(ref_label_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 項目タブ
        items_frame = ttk.Frame(notebook)
        notebook.add(items_frame, text="項目")
        self._create_items_tab(items_frame)
        
        # イベントタブ
        events_frame = ttk.Frame(notebook)
        notebook.add(events_frame, text="イベント")
        self._create_events_tab(events_frame)
        
        # 区分タブ
        categories_frame = ttk.Frame(notebook)
        notebook.add(categories_frame, text="区分")
        self._create_categories_tab(categories_frame)
        
        # DIM区分タブ
        dim_bins_frame = ttk.Frame(notebook)
        notebook.add(dim_bins_frame, text="DIM区分")
        self._create_dim_bins_tab(dim_bins_frame)
        
        # グラフ種類タブ（グラフモードのみ表示）
        graph_types_frame = ttk.Frame(notebook)
        notebook.add(graph_types_frame, text="グラフ種類")
        self._create_graph_types_tab(graph_types_frame)
    
    def _create_items_tab(self, parent: tk.Widget):
        """項目タブを作成"""
        # Treeview + Scrollbar
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(tree_frame, columns=("key",), show="tree headings", height=15)
        tree.heading("#0", text="項目")
        tree.heading("key", text="キー")
        tree.column("#0", width=200)
        tree.column("key", width=100)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.config(yscrollcommand=scrollbar.set)
        
        # カテゴリ別に表示
        categories: Dict[str, List[str]] = {}
        for item_key, item_def in self.item_dictionary.items():
            if isinstance(item_def, dict):
                category = item_def.get("category", "OTHER")
                if category not in categories:
                    categories[category] = []
                categories[category].append(item_key)
        
        for category, item_keys in sorted(categories.items()):
            category_node = tree.insert("", tk.END, text=category, open=False)
            for item_key in sorted(item_keys):
                item_def = self.item_dictionary[item_key]
                display_name = item_def.get("display_name", item_key)
                tree.insert(category_node, tk.END, text=f"{display_name}（{item_key}）", values=(item_key,))
        
        tree.bind('<Double-1>', lambda e: self._on_item_double_click(tree))
        tree.bind('<Button-1>', lambda e: self._on_item_click(tree))
        
        self.items_tree = tree
    
    def _create_events_tab(self, parent: tk.Widget):
        """イベントタブを作成"""
        # Treeview + Scrollbar
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(tree_frame, columns=("number",), show="tree headings", height=15)
        tree.heading("#0", text="イベント")
        tree.heading("number", text="番号")
        tree.column("#0", width=200)
        tree.column("number", width=100)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.config(yscrollcommand=scrollbar.set)
        
        # カテゴリ別に表示
        categories: Dict[str, List[str]] = {}
        for event_number_str, event_data in self.event_dictionary.items():
            if isinstance(event_data, dict):
                category = event_data.get("category", "OTHER")
                if category not in categories:
                    categories[category] = []
                categories[category].append(event_number_str)
        
        for category, event_numbers in sorted(categories.items()):
            category_node = tree.insert("", tk.END, text=category, open=False)
            for event_number_str in sorted(event_numbers, key=int):
                event_data = self.event_dictionary[event_number_str]
                name_jp = event_data.get("name_jp", "")
                alias = event_data.get("alias", "")
                tree.insert(category_node, tk.END, text=f"{name_jp}（{alias} / {event_number_str}）", values=(event_number_str,))
        
        tree.bind('<Double-1>', lambda e: self._on_event_double_click(tree))
        tree.bind('<Button-1>', lambda e: self._on_event_click(tree))
        
        self.events_tree = tree
    
    def _create_categories_tab(self, parent: tk.Widget):
        """区分タブを作成（固定）"""
        # 固定区分を表示
        listbox = tk.Listbox(parent, height=15)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        # 固定区分（後で実装）
        listbox.insert(tk.END, "区分1")
        listbox.insert(tk.END, "区分2")
        
        self.categories_listbox = listbox
    
    def _create_dim_bins_tab(self, parent: tk.Widget):
        """DIM区分タブを作成（固定）"""
        listbox = tk.Listbox(parent, height=15)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        dim_bins = ["DIM7", "DIM14", "DIM21", "DIM30", "DIM50"]
        for dim_bin in dim_bins:
            listbox.insert(tk.END, dim_bin)
        
        listbox.bind('<Double-1>', lambda e: self._on_dim_bin_double_click(listbox))
        listbox.bind('<Button-1>', lambda e: self._on_dim_bin_click(listbox))
        
        self.dim_bins_listbox = listbox
    
    def _create_graph_types_tab(self, parent: tk.Widget):
        """グラフ種類タブを作成（固定）"""
        listbox = tk.Listbox(parent, height=15)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        graph_types = ["折れ線", "棒", "プロット", "箱ひげ", "生存曲線"]
        for graph_type in graph_types:
            listbox.insert(tk.END, graph_type)
        
        listbox.bind('<Double-1>', lambda e: self._on_graph_type_double_click(listbox))
        listbox.bind('<Button-1>', lambda e: self._on_graph_type_click(listbox))
        
        self.graph_types_listbox = listbox
    
    def _on_analysis_type_changed(self):
        """分析種別変更時の処理"""
        self._update_guides()
    
    
    def _on_text_changed(self, event=None):
        """Textウィジェットの内容変更時にガイドを更新"""
        self._update_guides()
    
    def _on_text_configure(self, event=None):
        """Textウィジェットのサイズ変更時にガイドを更新"""
        self._update_guides()
    
    def _update_guides(self):
        """ガイドを更新（行単位で表示/非表示を切り替え）"""
        if not hasattr(self, 'guide_canvas') or not hasattr(self, 'command_text'):
            return
        
        # 既存のガイドを削除
        for canvas_text_id in self.guide_labels.values():
            self.guide_canvas.delete(canvas_text_id)
        self.guide_labels.clear()
        
        # 現在の分析種別を取得
        analysis_type = self.analysis_type_var.get()
        guides = self.ANALYSIS_GUIDES.get(analysis_type, [])
        
        if not guides:
            return
        
        # Textウィジェットの内容を取得
        content = self.command_text.get("1.0", tk.END)
        lines = content.split('\n')
        
        # 各行のガイドを表示/非表示
        for i, guide in enumerate(guides):
            line_num = i + 1
            
            # 該当行の内容を取得
            if line_num <= len(lines):
                line_content = lines[line_num - 1].strip()
            else:
                line_content = ""
            
            # 行が空の場合のみガイドを表示
            if not line_content:
                self._show_guide_for_line(line_num, guide)
            else:
                # 行に入力がある場合はガイドを非表示（既に削除済み）
                pass
    
    def _show_guide_for_line(self, line_num: int, guide_text: str):
        """指定行にガイドを表示"""
        try:
            # Textウィジェットの行の位置を取得
            bbox = self.command_text.bbox(f"{line_num}.0")
            if not bbox:
                return
            
            x, y, width, height = bbox
            
            # パディングを追加（Textウィジェットの内部パディング）
            padding_x = 2
            padding_y = 2
            
            # Canvas上にテキストを描画
            canvas_text_id = self.guide_canvas.create_text(
                x + padding_x,
                y + padding_y,
                text=guide_text,
                anchor=tk.NW,
                fill="gray",
                font=("", 9),
                state=tk.NORMAL
            )
            
            # ガイドIDを保存
            self.guide_labels[line_num] = canvas_text_id
            
        except Exception as e:
            logger.debug(f"ガイド表示エラー (line {line_num}): {e}")
    
    def _on_canvas_click(self, event):
        """Canvasクリック時にTextウィジェットにフォーカスを移し、クリック位置にカーソルを移動"""
        # Textウィジェットにフォーカスを移す
        self.command_text.focus_set()
        
        # Canvasのクリック位置をTextウィジェットの座標系に変換
        # CanvasとTextウィジェットは同じ親フレーム内にあるため、相対位置を計算
        canvas_x = self.guide_canvas.winfo_x()
        canvas_y = self.guide_canvas.winfo_y()
        text_x = self.command_text.winfo_x()
        text_y = self.command_text.winfo_y()
        
        # クリック位置をTextウィジェットの座標に変換
        relative_x = event.x + (canvas_x - text_x)
        relative_y = event.y + (canvas_y - text_y)
        
        # Textウィジェットの該当位置にカーソルを移動
        try:
            index = self.command_text.index(f"@{relative_x},{relative_y}")
            self.command_text.mark_set(tk.INSERT, index)
            self.command_text.see(tk.INSERT)
        except Exception as e:
            logger.debug(f"Canvasクリック位置変換エラー: {e}")
    
    def _on_search(self, event):
        """検索処理"""
        search_text = self.search_entry.get().strip().lower()
        
        if not search_text:
            # 検索テキストが空の場合はすべて表示
            return
        
        # 項目タブを検索（簡易実装：選択してスクロール）
        if hasattr(self, 'items_tree'):
            self._search_in_treeview(self.items_tree, search_text)
        
        # イベントタブを検索
        if hasattr(self, 'events_tree'):
            self._search_in_treeview(self.events_tree, search_text)
    
    def _search_in_treeview(self, tree: ttk.Treeview, search_text: str):
        """Treeview内を検索して該当アイテムを選択・スクロール"""
        # すべてのアイテムを走査
        for item in tree.get_children():
            found = self._search_treeview_item(tree, item, search_text)
            if found:
                # 見つかったアイテムを選択してスクロール
                tree.selection_set(found)
                tree.see(found)
                return
    
    def _search_treeview_item(self, tree: ttk.Treeview, item: str, search_text: str) -> Optional[str]:
        """Treeviewアイテムを検索"""
        item_text = tree.item(item, "text").lower()
        if search_text in item_text:
            return item
        
        # 子アイテムも検索
        for child in tree.get_children(item):
            found = self._search_treeview_item(tree, child, search_text)
            if found:
                return found
        
        return None
    
    def _on_item_click(self, tree):
        """項目クリック時の処理：カーソル位置に挿入"""
        selection = tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = tree.item(item, "values")
        if not values:
            return
        
        item_key = values[0]
        
        # 表示名を取得
        display_name = self._get_item_display_name(item_key)
        insert_text = display_name
        
        # カーソル位置に挿入
        self.command_text.insert(tk.INSERT, insert_text)
        self.command_text.focus_set()
    
    def _on_item_double_click(self, tree):
        """項目ダブルクリック時の処理：分析種別とカーソル行に応じて適切な行へ挿入/置換"""
        selection = tree.selection()
        if not selection:
            return
        
        item = tree.selection()[0]
        values = tree.item(item, "values")
        if not values:
            return
        
        item_key = values[0]
        display_name = self._get_item_display_name(item_key)
        
        # 現在のカーソル行番号を取得
        cursor_index = self.command_text.index(tk.INSERT)
        current_line = int(cursor_index.split('.')[0])
        
        # 分析種別を取得
        analysis_type = self.analysis_type_var.get()
        
        # 分析種別と行番号に応じて挿入/置換
        target_line = self._determine_target_line(analysis_type, current_line, "item")
        
        if target_line:
            self._insert_or_replace_line(target_line, display_name, append=True)
    
    def _on_event_click(self, tree):
        """イベントクリック時の処理：カーソル位置に挿入"""
        selection = tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = tree.item(item, "values")
        if not values:
            return
        
        event_number_str = values[0]
        event_data = self.event_dictionary.get(event_number_str, {})
        name_jp = event_data.get("name_jp", "")
        
        # カーソル位置に挿入
        self.command_text.insert(tk.INSERT, name_jp)
        self.command_text.focus_set()
    
    def _on_event_double_click(self, tree):
        """イベントダブルクリック時の処理：分析種別とカーソル行に応じて適切な行へ挿入/置換"""
        selection = tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = tree.item(item, "values")
        if not values:
            return
        
        event_number_str = values[0]
        event_data = self.event_dictionary.get(event_number_str, {})
        name_jp = event_data.get("name_jp", "")
        
        # 現在のカーソル行番号を取得
        cursor_index = self.command_text.index(tk.INSERT)
        current_line = int(cursor_index.split('.')[0])
        
        # 分析種別を取得
        analysis_type = self.analysis_type_var.get()
        
        # 分析種別と行番号に応じて挿入/置換
        target_line = self._determine_target_line(analysis_type, current_line, "event")
        
        if target_line:
            self._insert_or_replace_line(target_line, name_jp, append=True)
    
    def _on_dim_bin_click(self, listbox):
        """DIM区分クリック時の処理：カーソル位置に挿入"""
        selection = listbox.curselection()
        if not selection:
            return
        
        dim_bin = listbox.get(selection[0])
        
        # カーソル位置に挿入
        self.command_text.insert(tk.INSERT, dim_bin)
        self.command_text.focus_set()
    
    def _on_dim_bin_double_click(self, listbox):
        """DIM区分ダブルクリック時の処理：GROUP BY行（2行目）に挿入/置換"""
        selection = listbox.curselection()
        if not selection:
            return
        
        dim_bin = listbox.get(selection[0])
        
        # 現在のカーソル行番号を取得
        cursor_index = self.command_text.index(tk.INSERT)
        current_line = int(cursor_index.split('.')[0])
        
        # 分析種別を取得
        analysis_type = self.analysis_type_var.get()
        
        # GROUP BY行（2行目）に挿入/置換
        target_line = self._determine_target_line(analysis_type, current_line, "dim_bin")
        
        if target_line:
            self._insert_or_replace_line(target_line, dim_bin, append=False)
    
    def _on_graph_type_click(self, listbox):
        """グラフ種類クリック時の処理：カーソル位置に挿入"""
        selection = listbox.curselection()
        if not selection:
            return
        
        graph_type = listbox.get(selection[0])
        
        # カーソル位置に挿入
        self.command_text.insert(tk.INSERT, graph_type)
        self.command_text.focus_set()
    
    def _on_graph_type_double_click(self, listbox):
        """グラフ種類ダブルクリック時の処理：グラフ種類行（3行目）に挿入/置換"""
        selection = listbox.curselection()
        if not selection:
            return
        
        graph_type = listbox.get(selection[0])
        
        # グラフモードの場合のみ
        if self.analysis_type_var.get() != "graph":
            # カーソル位置に挿入
            self.command_text.insert(tk.INSERT, graph_type)
            self.command_text.focus_set()
            return
        
        # グラフ種類行（3行目）に挿入/置換
        self._insert_or_replace_line(3, graph_type, append=False)
    
    def _determine_target_line(self, analysis_type: str, current_line: int, item_type: str) -> Optional[int]:
        """
        分析種別とカーソル行に応じて挿入/置換先の行番号を決定
        
        Args:
            analysis_type: 分析種別（"list", "agg", "eventcount", "graph", "repro"）
            current_line: 現在のカーソル行番号
            item_type: アイテムタイプ（"item", "event", "dim_bin"）
        
        Returns:
            ターゲット行番号（Noneの場合はカーソル位置に挿入）
        """
        if analysis_type == "list":
            # リスト: 1行目=項目、2行目以降=条件
            if item_type == "item":
                if current_line == 1:
                    return 1  # 1行目に置換
                else:
                    return None  # カーソル位置に挿入
            else:
                return None  # カーソル位置に挿入
        
        elif analysis_type == "agg":
            # 集計: 1行目=集計項目、2行目=GROUP BY基準、3行目以降=条件
            if item_type == "item":
                if current_line == 1:
                    return 1  # 1行目に置換
                else:
                    return None  # カーソル位置に挿入
            elif item_type == "dim_bin":
                return 2  # 2行目（GROUP BY行）に置換
            else:
                return None
        
        elif analysis_type == "eventcount":
            # イベント集計: 1行目=イベント、2行目=GROUP BY基準、3行目以降=条件
            if item_type == "event":
                if current_line == 1:
                    return 1  # 1行目に置換
                else:
                    return None  # カーソル位置に挿入
            elif item_type == "dim_bin":
                return 2  # 2行目（GROUP BY行）に置換
            else:
                return None
        
        elif analysis_type == "graph":
            # グラフ: 1行目=Y軸、2行目=X軸、3行目=グラフ種類、4行目=分ける基準、5行目以降=条件
            if item_type == "item":
                if current_line == 1:
                    return 1  # 1行目（Y軸）に置換
                elif current_line == 2:
                    return 2  # 2行目（X軸）に置換
                elif current_line == 4:
                    return 4  # 4行目（分ける基準）に置換
                else:
                    return None  # カーソル位置に挿入
            elif item_type == "dim_bin":
                return 4  # 4行目（分ける基準）に置換
            else:
                return None
        
        elif analysis_type == "repro":
            # 繁殖分析: 1行目=繁殖指標、2行目=GROUP BY基準、3行目以降=条件
            if item_type == "item":
                if current_line == 1:
                    return 1  # 1行目に置換
                else:
                    return None  # カーソル位置に挿入
            elif item_type == "dim_bin":
                return 2  # 2行目（GROUP BY行）に置換
            else:
                return None
        
        return None
    
    def _insert_or_replace_line(self, line_num: int, text: str, append: bool = False):
        """
        指定行に挿入または置換
        
        Args:
            line_num: 行番号（1始まり）
            text: 挿入/置換するテキスト
            append: Trueの場合は既存テキストに追加、Falseの場合は置換
        """
        # 行の開始位置と終了位置を取得
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        
        # 既存の行内容を取得
        existing_text = self.command_text.get(line_start, line_end).strip()
        
        if append and existing_text:
            # 既存テキストに追加（カンマ区切り）
            new_text = f"{existing_text}, {text}"
        else:
            # 置換
            new_text = text
        
        # 行を置換
        self.command_text.delete(line_start, line_end)
        self.command_text.insert(line_start, new_text)
        
        # カーソルを該当行の末尾に移動
        self.command_text.mark_set(tk.INSERT, f"{line_num}.end")
        self.command_text.focus_set()
    
    def _create_result_view(self, parent: tk.Widget):
        """結果表示ビューを作成（分析UIから完全に分離）"""
        # 結果表示フレーム（視覚的に明確に分離）
        result_label_frame = ttk.LabelFrame(parent, text="結果", padding=5)
        result_label_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ビュー切替タブ（表/グラフ/イベント一覧を1つの領域で切り替え）
        self.result_notebook = ttk.Notebook(result_label_frame)
        self.result_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 表ビュー
        table_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(table_frame, text="表")
        self._create_table_view(table_frame)
        
        # グラフビュー
        graph_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(graph_frame, text="グラフ")
        self._create_graph_view(graph_frame)
        
        # イベント一覧ビュー
        event_list_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(event_list_frame, text="イベント一覧")
        self._create_event_list_view(event_list_frame)
        
        # エラー/警告表示（結果表示エリアの下部）
        error_frame = ttk.Frame(result_label_frame)
        error_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.error_text = tk.Text(error_frame, height=3, wrap=tk.WORD, foreground="red")
        self.error_text.pack(fill=tk.X)
        
        # 初期状態：結果なし
        self._clear_result()
    
    def _create_event_list_view(self, parent: tk.Widget):
        """イベント一覧ビューを作成"""
        # Treeview + Scrollbar
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview（イベント一覧用）
        self.event_list_tree = ttk.Treeview(
            tree_frame,
            columns=("event_date", "event_name", "cow_id", "note"),
            show="headings",
            height=20
        )
        self.event_list_tree.heading("#0", text="ID")
        self.event_list_tree.heading("event_date", text="日付")
        self.event_list_tree.heading("event_name", text="イベント名")
        self.event_list_tree.heading("cow_id", text="個体ID")
        self.event_list_tree.heading("note", text="備考")
        
        self.event_list_tree.column("#0", width=50)
        self.event_list_tree.column("event_date", width=100)
        self.event_list_tree.column("event_name", width=150)
        self.event_list_tree.column("cow_id", width=100)
        self.event_list_tree.column("note", width=200)
        
        self.event_list_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar_v = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.event_list_tree.yview)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        self.event_list_tree.config(yscrollcommand=scrollbar_v.set)
        
        scrollbar_h = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.event_list_tree.xview)
        scrollbar_h.pack(fill=tk.X)
        self.event_list_tree.config(xscrollcommand=scrollbar_h.set)
    
    def _create_table_view(self, parent: tk.Widget):
        """表ビューを作成"""
        # Treeview + Scrollbar
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview（列は動的に設定）
        self.result_tree = ttk.Treeview(tree_frame, show="headings")
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar_v = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.config(yscrollcommand=scrollbar_v.set)
        
        scrollbar_h = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        scrollbar_h.pack(fill=tk.X)
        self.result_tree.config(xscrollcommand=scrollbar_h.set)
    
    def _create_graph_view(self, parent: tk.Widget):
        """グラフビューを作成"""
        if not MATPLOTLIB_AVAILABLE:
            label = ttk.Label(parent, text="matplotlib がインストールされていません。", foreground="red")
            label.pack(expand=True)
            self.graph_canvas = None
            return
        
        # matplotlib Figure
        self.graph_figure = Figure(figsize=(8, 6), dpi=100)
        self.graph_canvas = FigureCanvasTkAgg(self.graph_figure, parent)
        self.graph_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _on_execute(self):
        """実行処理"""
        try:
            # 分析種別を取得
            analysis_type = self.analysis_type_var.get()
            
            # 期間を取得
            period_start = self.period_start_entry.get().strip() or None
            period_end = self.period_end_entry.get().strip() or None
            
            # コマンドテキストを取得
            multiline_text = self.command_text.get("1.0", tk.END).strip()
            
            # QueryRouterで解析
            try:
                plan = self.query_router.parse(analysis_type, period_start, period_end, multiline_text)
            except Exception as e:
                import logging
                logging.error(f"QueryRouter解析エラー: {e}", exc_info=True)
                self.error_text.delete("1.0", tk.END)
                self.error_text.insert("1.0", f"コマンドの解析に失敗しました: {str(e)}")
                self._clear_result()
                return
            
            # Executorで実行
            try:
                result = self.executor.execute(plan)
            except Exception as e:
                import logging
                logging.error(f"Executor実行エラー: {e}", exc_info=True)
                self.error_text.delete("1.0", tk.END)
                self.error_text.insert("1.0", f"分析の実行に失敗しました: {str(e)}")
                self._clear_result()
                return
            
            # 結果を表示
            try:
                self._display_result(result)
            except Exception as e:
                import logging
                logging.error(f"結果表示エラー: {e}", exc_info=True)
                self.error_text.delete("1.0", tk.END)
                self.error_text.insert("1.0", f"結果の表示に失敗しました: {str(e)}")
                self._clear_result()
                return
            
            # コールバックを呼び出し
            if self.on_execute:
                try:
                    self.on_execute(result)
                except Exception as e:
                    import logging
                    logging.error(f"コールバック実行エラー: {e}", exc_info=True)
                    # コールバックのエラーは結果表示に影響しない
        
        except Exception as e:
            import logging
            logging.error(f"実行処理エラー: {e}", exc_info=True)
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert("1.0", f"予期しないエラーが発生しました: {str(e)}")
            self._clear_result()
    
    def _display_result(self, result: Dict[str, Any]):
        """結果を表示"""
        try:
            # エラー/警告をクリア
            self.error_text.delete("1.0", tk.END)
            
            # エラーがある場合（エラーは赤で表示）
            if result.get("errors"):
                error_text = "\n".join(result["errors"])
                self.error_text.insert("1.0", f"エラー:\n{error_text}")
                self.error_text.tag_add("error", "1.0", "end")
                self.error_text.tag_config("error", foreground="red")
                self._clear_result()
                return
            
            # 警告がある場合（警告は黄色で表示）
            if result.get("warnings"):
                warning_text = "\n".join(result["warnings"])
                self.error_text.insert("1.0", f"警告:\n{warning_text}")
                # 警告は黄色で表示（エラーは赤のまま）
                self.error_text.tag_add("warning", "1.0", "end")
                self.error_text.tag_config("warning", foreground="orange")
            
            # データがない場合
            if not result.get("success") or not result.get("data"):
                self.error_text.insert("1.0", "データがありません。条件を変更して再度実行してください。")
                self._clear_result()
                return
            
            data = result.get("data")
            if not data:
                self.error_text.insert("1.0", "データがありません")
                self._clear_result()
                return
            
            data_type = data.get("type")
            
            # データタイプに応じて表示
            if data_type == "table":
                self._display_table(data)
            elif data_type == "graph":
                self._display_graph(data)
            elif data_type == "event_list":
                self._display_event_list(data)
            else:
                self.error_text.insert("1.0", f"未対応のデータタイプ: {data_type}")
        
        except Exception as e:
            import logging
            logging.error(f"結果表示エラー: {e}", exc_info=True)
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert("1.0", f"結果の表示中にエラーが発生しました: {str(e)}")
            self._clear_result()
    
    def _display_table(self, data: Dict[str, Any]):
        """表を表示"""
        try:
            # 既存の列とデータをクリア
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            
            columns = data.get("columns", [])
            rows = data.get("rows", [])
            
            if not columns:
                return
            
            # 列名を表示名に変換
            column_display_names = []
            for col in columns:
                display_name = self._get_item_display_name(col)
                column_display_names.append(display_name)
            
            # 列を設定
            self.result_tree["columns"] = columns
            for col, display_name in zip(columns, column_display_names):
                self.result_tree.heading(col, text=display_name)
                self.result_tree.column(col, width=120)
            
            # データを挿入
            for row in rows:
                if isinstance(row, dict):
                    values = [str(row.get(col, "")) for col in columns]
                elif isinstance(row, (list, tuple)):
                    values = [str(v) for v in row[:len(columns)]]
                else:
                    continue
                self.result_tree.insert("", tk.END, values=values)
            
            # 表タブを選択
            self.result_notebook.select(0)
        
        except Exception as e:
            import logging
            logging.error(f"表表示エラー: {e}", exc_info=True)
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert("1.0", f"表の表示中にエラーが発生しました: {str(e)}")
    
    def _get_item_display_name(self, item_key: str) -> str:
        """項目キーから表示名を取得"""
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            if isinstance(item_def, dict):
                return item_def.get("display_name", item_key)
        return item_key
    
    def _display_graph(self, data: Dict[str, Any]):
        """グラフを表示"""
        if not MATPLOTLIB_AVAILABLE or not self.graph_canvas:
            self.error_text.insert("1.0", "グラフ機能は使用できません（matplotlib未インストール）")
            return
        
        # グラフをクリア
        self.graph_figure.clear()
        
        graph_type = data.get("graph_type", "line")
        x_data = data.get("x_data", [])
        y_data = data.get("y_data", [])
        
        if not x_data or not y_data:
            self.error_text.insert("1.0", "グラフデータがありません")
            return
        
        ax = self.graph_figure.add_subplot(111)
        
        if graph_type == "line":
            ax.plot(x_data, y_data, marker='o')
        elif graph_type == "bar":
            ax.bar(x_data, y_data)
        elif graph_type == "scatter":
            ax.scatter(x_data, y_data)
        elif graph_type == "box":
            ax.boxplot(y_data)
        else:
            ax.plot(x_data, y_data)
        
        ax.set_xlabel(data.get("x_label", "X"))
        ax.set_ylabel(data.get("y_label", "Y"))
        ax.set_title(data.get("title", "グラフ"))
        
        self.graph_figure.tight_layout()
        self.graph_canvas.draw()
        
        # グラフタブを選択
        self.result_notebook.select(1)
    
    def _display_event_list(self, data: Dict[str, Any]):
        """イベント一覧を表示"""
        # 既存のデータをクリア
        if hasattr(self, 'event_list_tree'):
            for item in self.event_list_tree.get_children():
                self.event_list_tree.delete(item)
        
        events = data.get("events", [])
        
        if not events:
            return
        
        # イベントを挿入
        for event in events:
            event_id = str(event.get("id", ""))
            event_date = event.get("event_date", "")
            event_name = event.get("event_name", "")
            cow_id = event.get("cow_id", "")
            note = event.get("note", "")
            
            self.event_list_tree.insert(
                "",
                tk.END,
                text=event_id,
                values=(event_date, event_name, cow_id, note)
            )
        
        # イベント一覧タブを選択
        self.result_notebook.select(2)
    
    def _clear_result(self):
        """結果をクリア"""
        # 表をクリア
        if hasattr(self, 'result_tree'):
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
        
        # イベント一覧をクリア
        if hasattr(self, 'event_list_tree'):
            for item in self.event_list_tree.get_children():
                self.event_list_tree.delete(item)
        
        # グラフをクリア
        if MATPLOTLIB_AVAILABLE and hasattr(self, 'graph_canvas') and self.graph_canvas:
            self.graph_figure.clear()
            self.graph_canvas.draw()

