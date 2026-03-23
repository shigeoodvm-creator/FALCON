"""
FALCON2 - 繁殖検診表 印刷プレビュー
A4縦サイズで印刷プレビューを表示
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List
from datetime import datetime
import json
import tempfile
import webbrowser
import html

from modules.formula_engine import FormulaEngine


class ReproductionCheckupPrintPreview:
    """繁殖検診表 印刷プレビューウィンドウ"""
    
    # A4縦サイズ（mm）
    A4_WIDTH_MM = 210
    A4_HEIGHT_MM = 297
    # 1mm = 約3.78ピクセル（96 DPI）
    MM_TO_PIXEL = 3.78
    
    def __init__(self, parent: tk.Widget,
                 results: List[Dict[str, Any]],
                 checkup_date: str,
                 formula_engine: FormulaEngine,
                 item_dictionary: Dict[str, Any],
                 additional_items: List[str],
                 column_order: Optional[List[str]] = None,
                 column_widths_settings: Optional[Dict[str, int]] = None,
                 farm_name: Optional[str] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            results: 抽出結果
            checkup_date: 検診予定日
            formula_engine: FormulaEngine インスタンス
            item_dictionary: 項目辞書
            additional_items: 追加項目コードリスト
            column_order: カラムの順序
            column_widths_settings: 列幅設定
        """
        self.parent = parent
        self.results = results
        self.checkup_date = checkup_date
        self.formula_engine = formula_engine
        self.item_dictionary = item_dictionary
        self.additional_items = additional_items
        self.column_order = column_order or [
            'cow_id', 'jpn10', 'lact', 'dim_or_age', 'last_checkup_date',
            'last_checkup_result', 'last_ai_date', 'dai', 'checkup_code'
        ]
        # 列幅設定（文字数ベース）
        self.column_widths_settings = column_widths_settings or {}
        # 日付を短縮表示するか（年月日→月日）
        self.short_date_format = True
        # 農場名
        self.farm_name = farm_name or '農場'
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖検診表 - 印刷プレビュー")
        
        # A4縦サイズに合わせたウィンドウサイズ（少し大きめに表示）
        width_px = int(self.A4_WIDTH_MM * self.MM_TO_PIXEL * 1.2)  # 120%で表示
        height_px = int(self.A4_HEIGHT_MM * self.MM_TO_PIXEL * 1.2)  # 120%で表示
        self.window.geometry(f"{width_px}x{height_px}")
        
        # 経産牛と育成牛を分ける
        parous_cows = [cow for cow in results if (cow.get('lact') or 0) >= 1]
        heifer_cows = [cow for cow in results if (cow.get('lact') or 0) == 0]
        
        # UI作成
        self._create_widgets(parous_cows, heifer_cows)
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self, parous_cows: List[Dict[str, Any]], 
                       heifer_cows: List[Dict[str, Any]]):
        """ウィジェットを作成"""
        # メインフレーム（スクロール可能）
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # キャンバスとスクロールバー
        canvas = tk.Canvas(main_frame, bg='white')
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 経産牛ページ
        if parous_cows:
            parous_frame = self._create_page(scrollable_frame, "経産牛", parous_cows)
            parous_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        # 育成牛ページ
        if heifer_cows:
            heifer_frame = self._create_page(scrollable_frame, "育成牛", heifer_cows)
            heifer_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # マウスホイールでスクロール
        def _on_mousewheel(event):
            # ウィンドウがまだ存在するかチェック
            try:
                if not self.window.winfo_exists():
                    return
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except (tk.TclError, AttributeError):
                # ウィンドウが既に破棄されている場合は何もしない
                pass
        
        # マウスホイールイベントをバインド
        self.mousewheel_handler = canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # ウィンドウが閉じられるときにバインディングを解除
        def _on_closing():
            try:
                canvas.unbind_all("<MouseWheel>")
            except:
                pass
            self.window.destroy()
        
        self.window.protocol("WM_DELETE_WINDOW", _on_closing)
        
        # 印刷ボタン
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        print_btn = ttk.Button(
            button_frame,
            text="印刷",
            command=self._on_print,
            width=15
        )
        print_btn.pack(side=tk.LEFT, padx=5)
        
        close_btn = ttk.Button(
            button_frame,
            text="閉じる",
            command=_on_closing,
            width=15
        )
        close_btn.pack(side=tk.LEFT, padx=5)
    
    def _create_page(self, parent: tk.Widget, title: str, 
                    cows: List[Dict[str, Any]]) -> ttk.Frame:
        """
        ページを作成
        
        Args:
            parent: 親ウィジェット
            title: タイトル（"経産牛" または "育成牛"）
            cows: 牛のリスト
        """
        """ページを作成"""
        page_frame = ttk.Frame(parent, relief=tk.SOLID, borderwidth=1)
        
        # ヘッダー部分（左揃え）
        header_frame = ttk.Frame(page_frame)
        header_frame.pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        # 繁殖検診表（左揃え）
        title_label = ttk.Label(
            header_frame,
            text="繁殖検診表",
            font=("", 14, "bold")
        )
        title_label.pack(anchor=tk.W, pady=(0, 2))
        
        # 区切り線1
        separator1 = ttk.Label(
            header_frame,
            text="＿" * 20,  # 全角アンダースコアで区切り線
            font=("", 10)
        )
        separator1.pack(anchor=tk.W, pady=(0, 5))
        
        # 農場名（左揃え）
        farm_label = ttk.Label(
            header_frame,
            text=self.farm_name,
            font=("", 10)
        )
        farm_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 検診日、経産牛/育成牛、出力日（同じ行に配置）
        date_cow_frame = ttk.Frame(header_frame)
        date_cow_frame.pack(anchor=tk.W, pady=(0, 5))
        
        # 検診日
        date_label = ttk.Label(
            date_cow_frame,
            text=f"検診日　{self.checkup_date}",
            font=("", 10)
        )
        date_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # 経産牛/育成牛
        cow_type_label = ttk.Label(
            date_cow_frame,
            text=title,
            font=("", 10)
        )
        cow_type_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # 出力日（本日の日付）
        output_date = datetime.now().strftime("%Y-%m-%d")
        output_label = ttk.Label(
            date_cow_frame,
            text=f"出力日：{output_date}",
            font=("", 10)
        )
        output_label.pack(side=tk.LEFT)
        
        # 統計情報を計算して表示
        stats_text = self._calculate_statistics(cows)
        if stats_text:
            stats_label = ttk.Label(
                header_frame,
                text=stats_text,
                font=("", 9)
            )
            stats_label.pack(anchor=tk.W, pady=(0, 5))
        
        # 区切り線2
        separator2 = ttk.Label(
            header_frame,
            text="＿" * 25,  # 全角アンダースコアで区切り線（少し長め）
            font=("", 10)
        )
        separator2.pack(anchor=tk.W, pady=(0, 5))
        
        # 表を作成
        table_frame = ttk.Frame(page_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # カラム名と表示名のマッピング（経産牛/育成牛で異なる）
        is_parous = (title == "経産牛")
        
        # 基本のカラム名マッピング（短縮版）
        column_labels = {
            'cow_id': 'ID',
            'jpn10': 'JPN10',
            'lact': '産次',
            'dim_or_age': 'DIM' if is_parous else '月齢',  # 経産牛はDIM、育成牛は月齢
            'last_checkup_date': '検診',
            'last_checkup_result': '前回検診結果',
            'last_ai_date': 'AI日',
            'dai': 'AI後',
            'checkup_code': 'ｺｰﾄﾞ'
        }
        
        # 追加項目のカラムIDリスト
        additional_column_ids = []
        if isinstance(self.additional_items, dict):
            # Dict形式の場合（新しい形式）
            additional_column_ids = sorted(self.additional_items.keys())
        else:
            # List形式の場合（旧形式）
            additional_column_ids = [f"additional_{i+1}" for i in range(len(self.additional_items))]
        
        # カラムIDリスト（列幅設定用）
        all_column_ids = self.column_order + additional_column_ids + ['result']  # resultは検診結果欄
        
        # ヘッダー行（column_orderに従って順序を決定）
        headers = []
        
        # 基本カラムのヘッダー
        for col in self.column_order:
            if col in column_labels:
                headers.append(column_labels[col])
        
        # 追加項目のヘッダー
        if isinstance(self.additional_items, dict):
            for col_id in additional_column_ids:
                item_code = self.additional_items.get(col_id, '')
                if item_code and item_code in self.item_dictionary:
                    item_data = self.item_dictionary[item_code]
                    display_name = item_data.get('display_name', item_code)
                    headers.append(display_name)
        else:
            for i, item_code in enumerate(self.additional_items):
                if item_code and item_code in self.item_dictionary:
                    item_data = self.item_dictionary[item_code]
                    display_name = item_data.get('display_name', item_code)
                    headers.append(display_name)
        
        # 検診結果記入欄を右端に追加
        headers.append('')  # 検診結果記入欄（右端）
        
        # ヘッダーを表示
        header_frame = tk.Frame(table_frame, bg='white')
        header_frame.pack(fill=tk.X)
        
        # カラム幅を計算（A4幅に合わせて調整、文字数ベース）
        num_columns = len(headers)
        available_width = int(self.A4_WIDTH_MM * self.MM_TO_PIXEL - 40)  # 余白を考慮
        
        # 文字幅の定義（全角1文字約10px、半角1文字約7px、平均約9pxで計算）
        char_width_px = 9
        
        # 検診結果欄の幅（10文字分、右端）
        result_column_chars = 10
        result_column_width = result_column_chars * char_width_px
        
        # 列幅設定から幅を取得（文字数ベース）
        column_widths_list = []
        total_set_chars = 0
        
        for col_id in all_column_ids:
            if col_id == 'result':
                # 検診結果欄は固定（10文字分）
                column_widths_list.append(result_column_width)
                total_set_chars += result_column_chars
            elif col_id in self.column_widths_settings:
                # 保存された設定（文字数）を使用
                saved_chars = self.column_widths_settings[col_id]
                saved_width = saved_chars * char_width_px
                column_widths_list.append(saved_width)
                total_set_chars += saved_chars
            elif col_id == 'last_checkup_result':
                # 所見記入欄（前回検診結果）はデフォルトで長めに
                note_chars = 42
                column_widths_list.append(note_chars * char_width_px)
                total_set_chars += note_chars
            else:
                # 設定がない場合は0（後で均等分配）
                column_widths_list.append(0)
        
        # 設定されていないカラムの幅を計算
        num_unset = sum(1 for w in column_widths_list if w == 0)
        if num_unset > 0:
            # 残りの幅を文字数ベースで分配
            remaining_width = available_width - (total_set_chars * char_width_px)
            remaining_chars = remaining_width // char_width_px
            unset_chars = remaining_chars // num_unset if num_unset > 0 else 0
            unset_width = unset_chars * char_width_px if unset_chars > 0 else remaining_width // num_unset
            
            # 0の部分を均等幅で埋める
            for i, width in enumerate(column_widths_list):
                if width == 0:
                    column_widths_list[i] = unset_width
        
        self.column_widths = column_widths_list
        
        for i, header in enumerate(headers):
            if i == len(headers) - 1:
                # 検診結果欄（下線のみ、右端）
                label = tk.Label(
                    header_frame,
                    text='',
                    font=("", 7, "bold"),
                    anchor='center',
                    bg='white'
                )
                label.grid(row=0, column=i, sticky='nsew', padx=0, pady=0)
                # 下線のみを描画するためのFrameを作成（10文字分以上の幅）
                underline_frame = tk.Frame(header_frame, height=2, bg='black')
                underline_frame.grid(row=1, column=i, sticky='ew', padx=2)
                header_frame.rowconfigure(1, minsize=2)
            else:
                # 項目名が2行になる場合の処理
                # 文字数から必要行数を計算（全角1文字約2文字分、折り返しを考慮）
                header_text = str(header)
                char_count = len(header_text) + sum(1 for c in header_text if ord(c) > 127)  # 全角文字は2文字分
                max_chars_per_line = max(6, (self.column_widths[i] - 4) // 9)  # 1文字約9px
                
                if char_count > max_chars_per_line:
                    # 2行に分ける
                    mid = len(header_text) // 2
                    # スペースや区切り文字の位置で分割
                    split_pos = mid
                    for j in range(mid - 2, min(mid + 3, len(header_text))):
                        if header_text[j] in [' ', '・', '、', '/']:
                            split_pos = j + 1
                            break
                    header_line1 = header_text[:split_pos].strip()
                    header_line2 = header_text[split_pos:].strip()
                    
                    # 2行ラベルを作成
                    label_frame = tk.Frame(header_frame, bg='white')
                    label_frame.grid(row=0, column=i, sticky='nsew', padx=0, pady=2)
                    
                    label1 = tk.Label(
                        label_frame,
                        text=header_line1,
                        font=("", 7, "bold"),
                        anchor='center',
                        bg='white'
                    )
                    label1.pack(fill=tk.X)
                    
                    label2 = tk.Label(
                        label_frame,
                        text=header_line2,
                        font=("", 7, "bold"),
                        anchor='center',
                        bg='white'
                    )
                    label2.pack(fill=tk.X)
                else:
                    # 1行で表示
                    label = tk.Label(
                        header_frame,
                        text=header,
                        font=("", 7, "bold"),
                        anchor='center',
                        bg='white'
                    )
                    label.grid(row=0, column=i, sticky='nsew', padx=0, pady=2)
            header_frame.columnconfigure(i, weight=1, minsize=self.column_widths[i])
        
        # データ行
        data_frame = tk.Frame(table_frame, bg='white')
        data_frame.pack(fill=tk.BOTH, expand=True)
        
        # 各行の高さを計算（A4高さに合わせて調整）
        max_rows_per_page = int((self.A4_HEIGHT_MM * self.MM_TO_PIXEL - 100) / 25)  # ヘッダーと余白を考慮
        
        for row_idx, cow in enumerate(cows[:max_rows_per_page]):  # 1ページ分のみ表示
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # 計算項目を取得（DIMは検診日時点で表示）
            calculated = self.formula_engine.calculate(cow_auto_id)
            
            # 各項目を取得
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            lact = cow.get('lact') or 0
            
            # 分娩後日数/月齢（経産牛は検診日時点のDIM）
            dim_or_age = ''
            if lact >= 1:
                dim = self.formula_engine.calculate_dim_at_date(cow_auto_id, self.checkup_date) if self.checkup_date else calculated.get('DIM')
                if dim is not None:
                    dim_or_age = str(dim)
            else:
                age = calculated.get('AGE')
                if age is not None:
                    dim_or_age = f"{age:.1f}"
            
            # 前回検診日・結果・最終AI日・授精後日数は当該産次のイベントのみで算出
            events = self.formula_engine.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            if lact >= 1 and cow.get('clvd'):
                clvd = cow.get('clvd')
                events_this_lact = [e for e in events if e.get('event_date') and e.get('event_date') >= clvd]
                repro_display = self.formula_engine.get_reproduction_checkup_display(
                    events_this_lact, self.checkup_date or ''
                )
                last_checkup_date = repro_display.get('LREPRO', '')
                last_checkup_result = repro_display.get('REPCT', '')
                last_ai_date = repro_display.get('LASTAI', '')
                dai = repro_display.get('DAI')
            else:
                last_checkup_date = calculated.get('LREPRO', '')
                last_checkup_result = calculated.get('REPCT', '')
                last_ai_date = calculated.get('LASTAI', '')
                dai = None
                if self.checkup_date:
                    try:
                        dai = self.formula_engine._calculate_dai_at_event_date(events, self.checkup_date)
                    except Exception as e:
                        import logging
                        logging.warning(f"DAI計算エラー (cow_auto_id={cow_auto_id}): {e}")
            dai_str = str(dai) if dai is not None else ''
            
            # 検診コード
            checkup_code = cow.get('checkup_code', '')
            
            # 値をフォーマット（column_orderに従って順序を決定）
            values = []
            
            # 基本カラムの値を取得
            for col in self.column_order:
                if col == 'cow_id':
                    values.append(self._format_value(cow_id))
                elif col == 'jpn10':
                    values.append(self._format_value(jpn10))
                elif col == 'lact':
                    values.append(str(lact) if lact is not None else 'ー')
                elif col == 'dim_or_age':
                    values.append(self._format_value(dim_or_age))
                elif col == 'last_checkup_date':
                    values.append(self._format_value(last_checkup_date, is_date=True))
                elif col == 'last_checkup_result':
                    values.append(self._format_value(last_checkup_result))
                elif col == 'last_ai_date':
                    values.append(self._format_value(last_ai_date, is_date=True))
                elif col == 'dai':
                    values.append(self._format_value(dai_str))
                elif col == 'checkup_code':
                    values.append(self._format_value(checkup_code))
            
            # 追加項目の値を取得（column_orderに従って）
            if isinstance(self.additional_items, dict):
                for col_id in additional_column_ids:
                    item_code = self.additional_items.get(col_id, '')
                    if item_code and item_code in calculated:
                        value = calculated.get(item_code)
                        values.append(self._format_value(value))
                    else:
                        values.append('ー')
            else:
                for item_code in self.additional_items:
                    if item_code and item_code in calculated:
                        value = calculated.get(item_code)
                        values.append(self._format_value(value))
                    else:
                        values.append('ー')
            
            # 検診結果記入欄を右端に追加
            values.append('')  # 検診結果記入欄（空白、右端）
            
            # 行を表示
            for col_idx, value in enumerate(values):
                if col_idx == len(values) - 1:
                    # 検診結果欄（下線のみ、右端）
                    cell_frame = tk.Frame(data_frame, bg='white')
                    cell_frame.grid(row=row_idx, column=col_idx, sticky='nsew', padx=0, pady=0)
                    # 下線を描画（10文字分以上の幅を確保）
                    underline = tk.Frame(cell_frame, height=1, bg='black')
                    underline.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=2)
                    data_frame.columnconfigure(col_idx, weight=1, minsize=self.column_widths[col_idx])
                else:
                    # 値が長い場合は切り詰め
                    display_text = str(value)
                    if len(display_text) > 12:
                        display_text = display_text[:12] + '...'
                    
                    label = tk.Label(
                        data_frame,
                        text=display_text,
                        font=("", 6),
                        anchor='center',
                        bg='white'
                    )
                    label.grid(row=row_idx, column=col_idx, sticky='nsew', padx=0, pady=0)
                    data_frame.columnconfigure(col_idx, weight=1, minsize=self.column_widths[col_idx])
        
        # ページ番号表示（複数ページの場合）
        if len(cows) > max_rows_per_page:
            page_info = ttk.Label(
                page_frame,
                text=f"（1ページ目 / 全{len(cows)}頭中{max_rows_per_page}頭表示）",
                font=("", 8)
            )
            page_info.pack(pady=2)
        
        return page_frame
    
    def _calculate_statistics(self, cows: List[Dict[str, Any]]) -> str:
        """
        統計情報を計算
        
        Args:
            cows: 牛のリスト
        
        Returns:
            統計情報の文字列
        """
        if not cows:
            return ""
        
        # 検診コード別の集計
        stats = {}
        for cow in cows:
            code = cow.get('checkup_code', '')
            if code:
                stats[code] = stats.get(code, 0) + 1
        
        # 統計文字列を作成
        total = len(cows)
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
        
        return stats_text
    
    def _format_value(self, value, is_date=False):
        """値をフォーマット（Noneや空文字列は「ー」に変換）
        
        Args:
            value: フォーマットする値
            is_date: 日付の場合True（年月日→月日に変換）
        """
        if value is None or value == '':
            return 'ー'
        
        value_str = str(value).strip()
        
        # 日付の短縮表示（年月日→月日）
        if is_date and self.short_date_format:
            import re
            # YYYY-MM-DD または YYYY/MM/DD 形式を MM-DD または MM/DD に変換
            date_pattern = r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})'
            match = re.match(date_pattern, value_str)
            if match:
                month = match.group(2)
                day = match.group(3)
                separator = '-' if '-' in value_str else '/'
                return f"{month}{separator}{day}"
            # YYYY年MM月DD日 形式の場合
            date_pattern_jp = r'^(\d{4})年(\d{1,2})月(\d{1,2})日'
            match_jp = re.match(date_pattern_jp, value_str)
            if match_jp:
                month = match_jp.group(2)
                day = match_jp.group(3)
                return f"{month}/{day}"
            # その他の形式（既に短縮されているなど）はそのまま返す
        
        return value_str


def _format_value(value, is_date=False, short_date_format=True) -> str:
    """値をフォーマット（Noneや空文字列は「ー」に変換）"""
    if value is None or value == '':
        return 'ー'
    value_str = str(value).strip()
    if is_date and short_date_format:
        import re
        match = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$', value_str)
        if match:
            month = match.group(2).zfill(2)
            day = match.group(3).zfill(2)
            return f"{month}-{day}"
    return value_str


def _calculate_statistics(cows: List[Dict[str, Any]]) -> str:
    """統計情報を計算"""
    if not cows:
        return ""
    stats = {}
    for cow in cows:
        code = cow.get('checkup_code', '')
        if code:
            stats[code] = stats.get(code, 0) + 1
    total = len(cows)
    stats_text = f"抽出頭数：{total}"
    code_order = [
        "ﾌﾚｯｼｭ", "繁殖1", "繁殖2", "妊鑑", "再妊", "妊鑑2",
        "分娩？", "チェック", "育繁1", "育繁2", "育妊", "育再妊"
    ]
    for code in code_order:
        if code in stats:
            stats_text += f"　{code}：{stats[code]}"
    for code, count in sorted(stats.items()):
        if code not in code_order:
            stats_text += f"　{code}：{count}"
    return stats_text


def open_html_print_preview(
    results: List[Dict[str, Any]],
    checkup_date: str,
    formula_engine: FormulaEngine,
    item_dictionary: Dict[str, Any],
    additional_items: List[str],
    column_order: Optional[List[str]] = None,
    column_widths_settings: Optional[Dict[str, int]] = None,
    farm_name: Optional[str] = None,
    checkup_logic: Optional[Any] = None
):
    """繁殖検診表をHTML方式で表示（ブラウザ印刷）
    
    checkup_logic: 産次でイベントをフィルタするために使用。
                  指定時は前回検診日・前回検診結果・最終AI/ET日・授精後日数を
                  当該産次のイベントのみから算出する。
    """
    column_order = column_order or [
        'cow_id', 'jpn10', 'lact', 'dim_or_age', 'last_checkup_date',
        'last_checkup_result', 'last_ai_date', 'dai', 'checkup_code'
    ]
    farm_name = farm_name or '農場'
    short_date_format = True
    
    parous_cows = [cow for cow in results if (cow.get('lact') or 0) >= 1]
    heifer_cows = [cow for cow in results if (cow.get('lact') or 0) == 0]
    
    def build_headers(is_parous: bool):
        column_labels = {
            'cow_id': 'ID',
            'jpn10': 'JPN10',
            'lact': '産次',
            'dim_or_age': 'DIM' if is_parous else '月齢',
            'last_checkup_date': '検診',
            'last_checkup_result': '前回検診結果',
            'last_ai_date': 'AI日',
            'dai': 'AI後',
            'checkup_code': 'ｺｰﾄﾞ'
        }
        additional_column_ids = []
        if isinstance(additional_items, dict):
            additional_column_ids = sorted(additional_items.keys())
        else:
            additional_column_ids = [f"additional_{i+1}" for i in range(len(additional_items))]
        headers = []
        for col in column_order:
            if col in column_labels:
                headers.append(column_labels[col])
        if isinstance(additional_items, dict):
            for col_id in additional_column_ids:
                item_code = additional_items.get(col_id, '')
                if item_code and item_code in item_dictionary:
                    display_name = item_dictionary[item_code].get('display_name', item_code)
                    headers.append(display_name)
        else:
            for item_code in additional_items:
                if item_code and item_code in item_dictionary:
                    display_name = item_dictionary[item_code].get('display_name', item_code)
                    headers.append(display_name)
        headers.append('')
        return headers, additional_column_ids
    
    def build_rows(cows: List[Dict[str, Any]], additional_column_ids: List[str]):
        rows = []
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            calculated = formula_engine.calculate(cow_auto_id)
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            lact = cow.get('lact') or 0
            dim_or_age = ''
            if lact >= 1:
                dim = formula_engine.calculate_dim_at_date(cow_auto_id, checkup_date) if checkup_date else calculated.get('DIM')
                if dim is not None:
                    dim_or_age = str(dim)
            else:
                age = calculated.get('AGE')
                if age is not None:
                    dim_or_age = f"{age:.1f}"
            # 前回検診日・前回検診結果・最終AI/ET日・授精後日数は当該産次のイベントのみで算出
            # （産次更新＝分娩後は前産次の検診・AI等をリセット）
            if checkup_logic:
                events = formula_engine.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                events_this_lact = checkup_logic.filter_events_current_lactation(events, cow)
                last_checkup_date = formula_engine._get_last_reproduction_check_date(events_this_lact) or ''
                last_checkup_result = formula_engine._get_last_reproduction_check_note(events_this_lact) or ''
                last_ai_date = (checkup_logic._get_last_ai_date(events_this_lact, checkup_date) or '') if checkup_date else ''
                dai = checkup_logic._calculate_dai(events_this_lact, checkup_date) if checkup_date else None
            else:
                last_checkup_date = calculated.get('LREPRO', '')
                last_checkup_result = calculated.get('REPCT', '')
                last_ai_date = calculated.get('LASTAI', '')
                dai = None
                if checkup_date:
                    try:
                        events = formula_engine.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                        dai = formula_engine._calculate_dai_at_event_date(events, checkup_date)
                    except Exception:
                        dai = None
            dai_str = str(dai) if dai is not None else ''
            checkup_code = cow.get('checkup_code', '')
            
            values = []
            for col in column_order:
                if col == 'cow_id':
                    values.append(_format_value(cow_id, short_date_format=short_date_format))
                elif col == 'jpn10':
                    values.append(_format_value(jpn10, short_date_format=short_date_format))
                elif col == 'lact':
                    values.append(str(lact) if lact is not None else 'ー')
                elif col == 'dim_or_age':
                    values.append(_format_value(dim_or_age, short_date_format=short_date_format))
                elif col == 'last_checkup_date':
                    values.append(_format_value(last_checkup_date, is_date=True, short_date_format=short_date_format))
                elif col == 'last_checkup_result':
                    values.append(_format_value(last_checkup_result, short_date_format=short_date_format))
                elif col == 'last_ai_date':
                    values.append(_format_value(last_ai_date, is_date=True, short_date_format=short_date_format))
                elif col == 'dai':
                    values.append(_format_value(dai_str, short_date_format=short_date_format))
                elif col == 'checkup_code':
                    values.append(_format_value(checkup_code, short_date_format=short_date_format))
            if isinstance(additional_items, dict):
                for col_id in additional_column_ids:
                    item_code = additional_items.get(col_id, '')
                    if item_code and item_code in calculated:
                        values.append(_format_value(calculated.get(item_code), short_date_format=short_date_format))
                    else:
                        values.append('ー')
            else:
                for item_code in additional_items:
                    if item_code and item_code in calculated:
                        values.append(_format_value(calculated.get(item_code), short_date_format=short_date_format))
                    else:
                        values.append('ー')
            values.append('')
            rows.append(values)
        return rows
    
    output_date = datetime.now().strftime("%Y-%m-%d")
    html_content = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>繁殖検診表 - 印刷</title>
    <style>
        body { font-family: "MS Gothic","MS PGothic","Meiryo",sans-serif; margin: 20px; font-size: 11px; }
        h1 { font-size: 16px; margin-bottom: 6px; }
        .section { margin-bottom: 24px; page-break-after: always; }
        .header { margin-bottom: 8px; }
        .stats { margin-top: 4px; font-size: 10px; }
        table { border-collapse: collapse; width: 100%; margin-top: 6px; }
        th, td { border: 1px solid #333; padding: 4px 6px; text-align: left; }
        th { background: #e0e0e0; }
        tbody tr:nth-child(even) { background: #f3f3f3; }
        .result-cell { border-bottom: 1px solid #000; min-width: 9ch; }
        @media print {
            @page {
                size: portrait;
                margin: 0.5cm;
            }
            .section { page-break-after: always; }
            tbody tr:nth-child(even) { background: #f3f3f3 !important; }
        }
        body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    </style>
</head>
<body>
"""
    
    def render_section(title: str, cows: List[Dict[str, Any]]):
        if not cows:
            return ""
        headers, additional_column_ids = build_headers(title == "経産牛")
        rows = build_rows(cows, additional_column_ids)
        stats_text = _calculate_statistics(cows)
        section_html = f"""
<div class="section">
  <div class="header">
    <h1>繁殖検診表</h1>
    <div>{html.escape(str(farm_name))}</div>
    <div>検診日　{html.escape(str(checkup_date))}　{html.escape(title)}　出力日：{output_date}</div>
    <div class="stats">{html.escape(stats_text)}</div>
  </div>
  <table>
    <thead><tr>
"""
        for h in headers:
            section_html += f"<th>{html.escape(str(h))}</th>"
        section_html += "</tr></thead><tbody>"
        for row in rows:
            section_html += "<tr>"
            for idx, val in enumerate(row):
                if idx == len(row) - 1:
                    section_html += "<td class='result-cell'>&nbsp;</td>"
                else:
                    section_html += f"<td>{html.escape(str(val))}</td>"
            section_html += "</tr>"
        section_html += "</tbody></table></div>"
        return section_html
    
    html_content += render_section("経産牛", parous_cows)
    html_content += render_section("育成牛", heifer_cows)
    html_content += "</body></html>"
    
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.html', delete=False) as f:
        f.write(html_content)
        temp_file_path = f.name
    # 直接ブラウザで開き、追加の案内ダイアログは表示しない
    webbrowser.open(f'file://{temp_file_path}')
    
    def _on_print(self):
        """印刷ボタンをクリック"""
        messagebox.showinfo("情報", "印刷機能は今後実装予定です。")
        # TODO: 実際の印刷機能を実装
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

