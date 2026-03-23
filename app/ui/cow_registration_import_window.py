"""
FALCON2 - 新規個体登録取り込みウィンドウ
CSVファイルから新規個体を登録する
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple
import pandas as pd
import logging
import re

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)


class CowRegistrationImportWindow:
    """新規個体登録取り込みウィンドウ"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            db_handler: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        
        self.window = tk.Toplevel(parent)
        self.window.title("新規一括導入")
        self.window.geometry("1200x700")
        
        self.csv_path: Optional[Path] = None
        self.data_rows: List[Dict[str, Any]] = []
        
        self._create_widgets()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # CSVファイル選択
        file_frame = ttk.LabelFrame(main_frame, text="CSVファイル選択", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=80, state='readonly').pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(file_frame, text="参照...", command=self._select_csv_file).pack(side=tk.LEFT)
        
        # プレビューエリア
        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー（編集可能）", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # ツリービュー（プレビュー用、編集可能）
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 列定義
        columns = ('cow_id', 'jpn10', 'brd', 'pen', 'birth_date', 'lact', 'clvd', 'last_ai_date', 'ai_count', 'dam', 'pregnant_status')
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        self.tree.heading('cow_id', text='ID（4桁）')
        self.tree.heading('jpn10', text='JPN10（個体識別番号）')
        self.tree.heading('brd', text='品種')
        self.tree.heading('pen', text='群')
        self.tree.heading('birth_date', text='生年月日')
        self.tree.heading('lact', text='産次')
        self.tree.heading('clvd', text='分娩月日')
        self.tree.heading('last_ai_date', text='最終授精月日')
        self.tree.heading('ai_count', text='授精回数')
        self.tree.heading('dam', text='母牛（母牛識別番号）')
        self.tree.heading('pregnant_status', text='妊娠状態')
        
        self.tree.column('cow_id', width=80)
        self.tree.column('jpn10', width=120)
        self.tree.column('brd', width=100)
        self.tree.column('pen', width=60)
        self.tree.column('birth_date', width=100)
        self.tree.column('lact', width=60)
        self.tree.column('clvd', width=100)
        self.tree.column('last_ai_date', width=120)
        self.tree.column('ai_count', width=80)
        self.tree.column('dam', width=120)
        self.tree.column('pregnant_status', width=120)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # 編集可能にするためのイベントバインディング
        self.tree.bind('<Double-Button-1>', self._on_cell_double_click)
        
        # 情報表示
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.info_label = ttk.Label(info_frame, text="CSVファイルを選択してください")
        self.info_label.pack(side=tk.LEFT)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="取り込み実行", command=self._execute_import).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)
    
    def _select_csv_file(self):
        """CSVファイルを選択"""
        file_path = filedialog.askopenfilename(
            title="新規一括導入CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_path = Path(file_path)
            self.file_path_var.set(str(self.csv_path))
            self._parse_csv()
    
    def _normalize_date(self, date_str: str) -> Optional[str]:
        """日付文字列をYYYY-MM-DD形式に正規化"""
        if not date_str or pd.isna(date_str):
            return None
        
        date_str = str(date_str).strip()
        if not date_str:
            return None
        
        # Excelの「########」表示をスキップ
        if '#' in date_str or date_str == '########':
            return None
        
        # 様々な日付形式を試す
        date_formats = [
            '%Y-%m-%d',
            '%Y/%m/%d',
            '%Y/%m/%d',  # 重複だが念のため
            '%Y%m%d',
            '%m/%d/%Y',
            '%d/%m/%Y',
        ]
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                return date_obj.strftime('%Y-%m-%d')
            except:
                continue
        
        # pandasの日付解析を試す
        try:
            date_obj = pd.to_datetime(date_str, errors='coerce')
            if pd.notna(date_obj):
                return date_obj.strftime('%Y-%m-%d')
        except:
            pass
        
        return None
    
    def _find_column_index(self, headers: List[str], keywords: List[str]) -> Optional[int]:
        """ヘッダー行からキーワードを含む列のインデックスを検索"""
        # まず完全一致を試す（空白を除去して比較）
        for idx, header in enumerate(headers):
            if pd.notna(header):
                header_str = str(header).strip()
                header_clean = header_str.replace(' ', '').replace('　', '')  # 全角・半角スペースを除去
                for keyword in keywords:
                    keyword_clean = keyword.replace(' ', '').replace('　', '')
                    if header_str == keyword or header_str.upper() == keyword.upper() or header_clean == keyword_clean or header_clean.upper() == keyword_clean.upper():
                        logger.info(f"列検出（完全一致）: '{header_str}' -> キーワード '{keyword}'")
                        return idx
        
        # 完全一致がない場合は部分一致を試す
        for idx, header in enumerate(headers):
            if pd.notna(header):
                header_str = str(header).strip()
                header_upper = header_str.upper()
                for keyword in keywords:
                    keyword_upper = keyword.upper()
                    if keyword_upper in header_upper or header_upper in keyword_upper:
                        logger.info(f"列検出（部分一致）: '{header_str}' -> キーワード '{keyword}'")
                        return idx
        return None
    
    def _parse_csv(self):
        """CSVファイルをパースして自動判別"""
        if not self.csv_path or not self.csv_path.exists():
            messagebox.showerror("エラー", "CSVファイルが見つかりません")
            return
        
        try:
            # CSVを読み込む（複数のエンコーディングを試す）
            df = None
            for encoding in ['shift_jis', 'utf-8', 'cp932']:
                try:
                    df = pd.read_csv(self.csv_path, encoding=encoding, dtype=str)
                    break
                except:
                    continue
            
            if df is None:
                messagebox.showerror("エラー", "CSVファイルの読み込みに失敗しました")
                return
            
            # ヘッダー行を取得（最初の行をヘッダーとして使用）
            headers = df.columns.tolist()
            
            # 列位置を自動判別
            col_jpn10 = self._find_column_index(headers, ['JPN10', '個体識別番号', '個体識別', '識別番号', 'JPN', '個体識別番'])
            col_brd = self._find_column_index(headers, ['品種', 'BRD', '品種名'])
            col_pen = self._find_column_index(headers, ['群', 'PEN', '群番号', '群コード', '群名'])
            col_birth_date = self._find_column_index(headers, ['生年月日', '生年月', 'BIRTH', 'BTHD', '誕生日'])
            col_lact = self._find_column_index(headers, ['産次', 'LACT', 'LACTATION', '経産'])
            col_clvd = self._find_column_index(headers, ['最終分娩', '分娩', '分娩日', '分娩月日', 'CLVD', 'CALV', 'CALVING'])
            col_last_ai = self._find_column_index(headers, ['最終AI', '最終授精', 'LAST_AI', 'LASTAI', '授精日', 'AI日', 'AI', '最終AI日', '最終AI日付'])
            col_ai_count = self._find_column_index(headers, ['授精回数', 'AI回数', 'AICNT', 'AI_COUNT', '授精数'])
            col_dam = self._find_column_index(headers, ['母牛識別番', '母牛', '母', 'DAM', '母牛識別番号', '母牛ID'])
            
            # デバッグ: 検出された列をログに出力
            logger.info(f"検出された列: JPN10={col_jpn10}, 品種={col_brd}, 群={col_pen}, 生年月日={col_birth_date}, 産次={col_lact}, 分娩={col_clvd}, 最終AI={col_last_ai}, 授精回数={col_ai_count}, 母牛={col_dam}")
            logger.info(f"CSVヘッダー: {headers}")
            
            # データ行を取得
            data_rows = []
            num_cols = len(df.columns)
            
            def get_value(row, col_idx: Optional[int], default='') -> str:
                """行から値を安全に取得"""
                if col_idx is None or col_idx >= num_cols:
                    return default
                try:
                    val = row.iloc[col_idx]
                    if pd.isna(val):
                        return default
                    val_str = str(val).strip()
                    # Excelの「########」表示をスキップ
                    if '#' in val_str or val_str == '########':
                        return default
                    return val_str
                except:
                    return default
            
            for idx, row in df.iterrows():
                # 空行をスキップ
                if row.isna().all():
                    continue
                
                # JPN10を取得（必須）
                jpn10 = get_value(row, col_jpn10)
                if not jpn10 or len(jpn10) != 10 or not jpn10.isdigit():
                    continue  # JPN10が無効な場合はスキップ
                
                # JPN10の6-9桁目（1-indexed、0-indexedで5-9桁目）をIDとして使用
                # 例: 1544209805 -> 0980
                cow_id = jpn10[5:9].zfill(4)
                
                # 各項目を取得（ない場合は空欄）。品種・群はCSVに列が無い場合にデフォルト値を全頭に設定
                brd = get_value(row, col_brd) if col_brd is not None else 'ホルスタイン'
                pen = get_value(row, col_pen) if col_pen is not None else '1'
                birth_date = self._normalize_date(get_value(row, col_birth_date)) or ''
                
                lact_str = get_value(row, col_lact)
                lact = ''
                if lact_str:
                    try:
                        lact = int(float(lact_str))
                    except:
                        pass
                
                clvd = self._normalize_date(get_value(row, col_clvd)) or ''
                
                last_ai_date = self._normalize_date(get_value(row, col_last_ai)) or ''
                
                ai_count_str = get_value(row, col_ai_count)
                ai_count = ''
                if ai_count_str:
                    try:
                        ai_count = int(float(ai_count_str))
                    except:
                        pass
                
                dam = get_value(row, col_dam) or ''
                
                # 妊娠状態（最終授精月日がある場合は「妊娠鑑定待ち」をデフォルト、ない場合は空）
                pregnant_status = ''
                if last_ai_date:
                    pregnant_status = '妊娠鑑定待ち'
                
                data_rows.append({
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'brd': brd,
                    'pen': pen,
                    'birth_date': birth_date,
                    'lact': lact,
                    'clvd': clvd,
                    'last_ai_date': last_ai_date,
                    'ai_count': ai_count,
                    'dam': dam,
                    'pregnant_status': pregnant_status
                })
            
            self.data_rows = data_rows
            self._update_preview()
            
        except Exception as e:
            logger.error(f"CSVパースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")
    
    def _update_preview(self):
        """プレビューを更新"""
        # 既存のアイテムをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 既存個体数と新規個体数をカウント
        existing_count = 0
        new_count = 0
        
        # データを表示
        for row_data in self.data_rows:
            jpn10 = row_data.get('jpn10', '')
            
            # 既存の個体をチェック
            existing_cows = []
            try:
                existing_cows = self.db.get_cows_by_jpn10(jpn10)
            except:
                pass
            
            if existing_cows:
                status = "既存"
                existing_count += 1
            else:
                status = "新規"
                new_count += 1
            
            # 値を表示用にフォーマット
            cow_id = row_data.get('cow_id', '')
            brd = row_data.get('brd', '')
            pen = row_data.get('pen', '')
            birth_date = row_data.get('birth_date', '')
            lact = row_data.get('lact', '')
            clvd = row_data.get('clvd', '')
            last_ai_date = row_data.get('last_ai_date', '')
            ai_count = row_data.get('ai_count', '')
            dam = row_data.get('dam', '')
            pregnant_status = row_data.get('pregnant_status', '')
            
            self.tree.insert('', tk.END, values=(
                cow_id,
                jpn10,
                brd,
                pen,
                birth_date,
                lact if lact != '' else '',
                clvd,
                last_ai_date,
                ai_count if ai_count != '' else '',
                dam,
                pregnant_status
            ), tags=(status,))
        
        # 既存個体の色を設定（グレー）、新規個体は黒
        self.tree.tag_configure('既存', foreground='gray')
        self.tree.tag_configure('新規', foreground='black')
        
        # 情報ラベルを更新（総件数、既存件数、導入予定件数）
        total_count = len(self.data_rows)
        self.info_label.config(
            text=f"データ件数: {total_count}件 | "
                 f"既存個体（スキップ）: {existing_count}件 | "
                 f"導入予定: {new_count}件（ダブルクリックで編集可能）"
        )
    
    def _on_cell_double_click(self, event):
        """セルのダブルクリックで編集"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        # クリックされた列を取得
        column = self.tree.identify_column(event.x)
        if not column:
            return
        
        column_index = int(column.replace('#', '')) - 1  # 0ベースのインデックス
        if column_index < 0 or column_index >= len(self.tree['columns']):
            return
        
        column_name = self.tree['columns'][column_index]
        
        # 妊娠状態列の場合は選択ダイアログを表示（最終授精月日がある場合のみ）
        if column_name == 'pregnant_status':
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                row_data = self.data_rows[item_index]
                if row_data.get('last_ai_date'):
                    self._edit_pregnant_status(item)
                else:
                    messagebox.showinfo("情報", "最終授精月日が入力されていないため、妊娠状態を選択できません。")
            return
        
        # 現在の値を取得
        values = list(self.tree.item(item, 'values'))
        current_value = values[column_index] if column_index < len(values) else ''
        
        # 編集ダイアログを表示
        edit_window = tk.Toplevel(self.window)
        edit_window.title(f"{self.tree.heading(column_name, 'text')}を編集")
        edit_window.geometry("300x100")
        
        ttk.Label(edit_window, text=f"新しい値を入力してください:").pack(pady=10)
        
        entry = ttk.Entry(edit_window, width=30)
        entry.pack(pady=5)
        entry.insert(0, str(current_value))
        entry.focus()
        
        def save_edit():
            new_value = entry.get().strip()
            
            # 日付列の場合は正規化
            if column_name in ['birth_date', 'clvd', 'last_ai_date']:
                normalized = self._normalize_date(new_value)
                if normalized:
                    new_value = normalized
                elif new_value:
                    messagebox.showerror("エラー", "日付の形式が正しくありません（YYYY-MM-DD形式で入力してください）")
                    return
            
            # 数値列の場合は検証
            if column_name in ['lact', 'ai_count']:
                if new_value:
                    try:
                        int(new_value)
                    except:
                        messagebox.showerror("エラー", "数値を入力してください")
                        return
            
            # 値を更新
            values[column_index] = new_value
            self.tree.item(item, values=values)
            
            # データ行も更新
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                self.data_rows[item_index][column_name] = new_value
                
                # 最終授精日が入力された場合、妊娠状態を「妊娠鑑定待ち」に更新
                if column_name == 'last_ai_date' and new_value:
                    if not self.data_rows[item_index].get('pregnant_status'):
                        self.data_rows[item_index]['pregnant_status'] = '妊娠鑑定待ち'
                        values[-1] = '妊娠鑑定待ち'
                        self.tree.item(item, values=values)
                # 最終授精日が削除された場合、妊娠状態をクリア
                elif column_name == 'last_ai_date' and not new_value:
                    self.data_rows[item_index]['pregnant_status'] = ''
                    values[-1] = ''
                    self.tree.item(item, values=values)
            
            edit_window.destroy()
        
        ttk.Button(edit_window, text="保存", command=save_edit).pack(pady=5)
        entry.bind('<Return>', lambda e: save_edit())
    
    def _edit_pregnant_status(self, item):
        """妊娠状態を編集"""
        # 現在の値を取得
        values = list(self.tree.item(item, 'values'))
        current_status = values[-1] if values else ''
        
        # 選択ダイアログを表示
        status_window = tk.Toplevel(self.window)
        status_window.title("妊娠状態を選択")
        status_window.geometry("250x200")
        
        ttk.Label(status_window, text="妊娠状態を選択してください:").pack(pady=10)
        
        status_var = tk.StringVar(value=current_status)
        
        statuses = ['妊娠鑑定待ち', '受胎なし', '妊娠', '乾乳']
        for status in statuses:
            ttk.Radiobutton(
                status_window,
                text=status,
                variable=status_var,
                value=status
            ).pack(anchor=tk.W, padx=20, pady=5)
        
        def save_status():
            new_status = status_var.get()
            values[-1] = new_status
            self.tree.item(item, values=values)
            
            # データ行も更新
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                self.data_rows[item_index]['pregnant_status'] = new_status
            
            status_window.destroy()
        
        ttk.Button(status_window, text="保存", command=save_status).pack(pady=10)
    
    def _execute_import(self):
        """取り込みを実行"""
        if not self.data_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return
        
        # 導入予定の個体数をカウント
        import_count = 0
        for row_data in self.data_rows:
            jpn10 = row_data.get('jpn10', '')
            existing_cows = []
            try:
                existing_cows = self.db.get_cows_by_jpn10(jpn10)
            except:
                pass
            if not existing_cows:
                import_count += 1
        
        result = messagebox.askyesno(
            "確認",
            f"導入予定: {import_count}件の個体を登録しますか？\n"
            f"（既存個体: {len(self.data_rows) - import_count}件はスキップされます）"
        )
        
        if not result:
            return
        
        success_count = 0
        skip_count = 0
        error_count = 0
        error_cows = []
        
        for row_data in self.data_rows:
            jpn10 = row_data.get('jpn10', '')
            cow_id = row_data.get('cow_id', '')
            
            # 既存の個体をチェック
            existing_cows = []
            try:
                existing_cows = self.db.get_cows_by_jpn10(jpn10)
            except:
                pass
            
            if existing_cows:
                skip_count += 1
                logger.info(f"既存個体をスキップ: jpn10={jpn10}")
                continue
            
            # 新規個体を登録
            try:
                # RCを決定
                clvd = row_data.get('clvd', '')
                last_ai_date = row_data.get('last_ai_date', '')
                pregnant_status = row_data.get('pregnant_status', '')
                
                # RC決定ロジック
                initial_rc = RuleEngine.RC_OPEN  # デフォルトは空胎
                if pregnant_status == '妊娠':
                    initial_rc = RuleEngine.RC_PREGNANT  # 5: 妊娠中
                elif pregnant_status == '乾乳':
                    initial_rc = RuleEngine.RC_DRY  # 6: 乾乳中
                elif last_ai_date:
                    if pregnant_status == '妊娠鑑定待ち':
                        initial_rc = RuleEngine.RC_BRED  # 3: 授精後
                    elif pregnant_status == '受胎なし':
                        initial_rc = RuleEngine.RC_OPEN  # 4: 空胎
                elif clvd and not last_ai_date:
                    initial_rc = RuleEngine.RC_FRESH  # 2: 分娩後
                
                # 個体データを作成（品種・群はCSVまたはプレビューで編集した値を使用）
                brd_val = (row_data.get('brd') or '').strip()
                pen_val = (row_data.get('pen') or '').strip()
                cow_data = {
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'brd': brd_val or None,
                    'bthd': row_data.get('birth_date') or None,
                    'entr': datetime.now().strftime('%Y-%m-%d'),
                    'lact': int(row_data.get('lact', 0)) if row_data.get('lact') else 0,
                    'clvd': clvd or None,
                    'rc': initial_rc,
                    'pen': pen_val or None,
                    'frm': None
                }
                
                cow_auto_id = self.db.insert_cow(cow_data)
                
                # 導入イベントを作成
                intro_json = {
                    "birth_date": row_data.get('birth_date'),
                    "lactation": row_data.get('lact', 0),
                    "calving_date": clvd,
                    "last_ai_date": last_ai_date,
                    "ai_count": row_data.get('ai_count', 0),
                    "dam": row_data.get('dam'),
                    "source": "csv_import"
                }
                
                intro_event = {
                    "cow_auto_id": cow_auto_id,
                    "event_number": RuleEngine.EVENT_IN,
                    "event_date": datetime.now().strftime('%Y-%m-%d'),
                    "json_data": intro_json,
                    "note": "CSVから取り込み"
                }
                
                event_id = self.db.insert_event(intro_event)
                self.rule_engine.on_event_added(event_id)
                
                # 分娩月日が入力された場合、分娩イベントを作成
                if clvd:
                    calv_event = {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_CALV,
                        "event_date": clvd,
                        "json_data": {},
                        "note": "CSV取り込み時の分娩"
                    }
                    calv_event_id = self.db.insert_event(calv_event)
                    self.rule_engine.on_event_added(calv_event_id)
                
                # 最終授精日が入力された場合、AIイベントを作成
                if last_ai_date:
                    # 授精回数（空欄の場合は1）
                    ai_count = row_data.get('ai_count', '')
                    if ai_count == '':
                        ai_count = 1
                    else:
                        try:
                            ai_count = int(ai_count)
                        except:
                            ai_count = 1
                    
                    # AIイベントのjson_data（受胎なしはO、妊娠・乾乳はP）
                    # 乾乳の場合は最終授精で妊娠しているためPとして記録し、受胎日・分娩予定日等を連動させる
                    ai_json_data = {}
                    if pregnant_status == '受胎なし':
                        ai_json_data['result'] = 'O'
                    elif pregnant_status == '妊娠' or pregnant_status == '乾乳':
                        ai_json_data['result'] = 'P'
                    
                    ai_event = {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_AI,
                        "event_date": last_ai_date,
                        "json_data": ai_json_data,
                        "note": f"授精回数: {ai_count}"  # NOTEに授精回数を反映
                    }
                    ai_event_id = self.db.insert_event(ai_event)
                    self.rule_engine.on_event_added(ai_event_id)
                    
                    # 妊娠状態に応じて追加イベントを作成
                    if pregnant_status == '妊娠':
                        # 妊娠プラスイベントを作成（受胎日・分娩予定日等を確定）
                        preg_event = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": RuleEngine.EVENT_PDP,
                            "event_date": last_ai_date,
                            "json_data": {},
                            "note": "CSV取り込み時の妊娠"
                        }
                        preg_event_id = self.db.insert_event(preg_event)
                        self.rule_engine.on_event_added(preg_event_id)
                    elif pregnant_status == '乾乳':
                        # 乾乳の場合も最終授精で妊娠しているため、妊娠プラスイベントを先に作成し
                        # 受胎日・分娩予定日等を連動して確定してから乾乳イベントを作成
                        preg_event = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": RuleEngine.EVENT_PDP,
                            "event_date": last_ai_date,
                            "json_data": {},
                            "note": "CSV取り込み時の妊娠（乾乳）"
                        }
                        preg_event_id = self.db.insert_event(preg_event)
                        self.rule_engine.on_event_added(preg_event_id)
                        # 乾乳イベントを作成
                        dry_event = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": RuleEngine.EVENT_DRY,
                            "event_date": datetime.now().strftime('%Y-%m-%d'),
                            "json_data": {},
                            "note": "CSV取り込み時の乾乳"
                        }
                        dry_event_id = self.db.insert_event(dry_event)
                        self.rule_engine.on_event_added(dry_event_id)
                
                # RuleEngineで状態を更新
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_cows.append({
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'reason': f'登録エラー: {str(e)}'
                })
                logger.error(f"個体登録エラー: cow_id={cow_id}, jpn10={jpn10}, エラー={e}", exc_info=True)
        
        # 結果を表示
        result_msg = (
            f"取り込みが完了しました\n"
            f"成功: {success_count}件\n"
            f"スキップ: {skip_count}件\n"
            f"エラー: {error_count}件"
        )
        
        if error_cows:
            result_msg += f"\n\nエラー個体が{len(error_cows)}件あります。詳細を表示しますか？"
            show_detail = messagebox.askyesno("完了", result_msg)
            if show_detail:
                self._show_error_cows_window(error_cows)
        else:
            messagebox.showinfo("完了", result_msg)
        
        # プレビューを更新
        self._update_preview()
    
    def _show_error_cows_window(self, error_cows: List[Dict[str, Any]]):
        """エラー個体の一覧を表示するウィンドウ"""
        error_window = tk.Toplevel(self.window)
        error_window.title("エラー個体一覧")
        error_window.geometry("600x400")
        
        # フレーム
        main_frame = ttk.Frame(error_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 説明ラベル
        desc_label = ttk.Label(
            main_frame,
            text=f"登録エラーが発生した個体一覧（{len(error_cows)}件）",
            font=("", 10, "bold")
        )
        desc_label.pack(pady=(0, 10))
        
        # ツリービュー
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        tree = ttk.Treeview(
            tree_frame,
            columns=('cow_id', 'jpn10', 'reason'),
            show='headings',
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        tree.heading('cow_id', text='ID')
        tree.heading('jpn10', text='個体識別番号')
        tree.heading('reason', text='エラー理由')
        
        tree.column('cow_id', width=100)
        tree.column('jpn10', width=150)
        tree.column('reason', width=300)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # データを追加
        for error_cow in error_cows:
            tree.insert('', tk.END, values=(
                error_cow.get('cow_id', ''),
                error_cow.get('jpn10', ''),
                error_cow.get('reason', '')
            ))
        
        # ボタン
        ttk.Button(main_frame, text="閉じる", command=error_window.destroy).pack(pady=10)
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()
