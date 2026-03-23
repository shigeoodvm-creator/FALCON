"""
FALCON2 - 乳検データ取り込みウィンドウ
既存の個体に対して乳検CSVファイルからデータを取り込む
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple
import pandas as pd
import logging
import subprocess
import platform

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)


class MilkTestImportWindow:
    """乳検データ取り込みウィンドウ"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            db_handler: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
        """
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        
        self.window = tk.Toplevel(parent)
        self.window.title("乳検データ取り込み")
        self.window.geometry("800x600")
        
        self.csv_path: Optional[Path] = None
        self.test_date: Optional[str] = None
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
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=60, state='readonly').pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(file_frame, text="参照...", command=self._select_csv_file).pack(side=tk.LEFT)
        
        # プレビューエリア
        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # ツリービュー（プレビュー用）
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(
            tree_frame,
            columns=('cow_id', 'jpn10', 'milk_yield', 'fat', 'protein', 'scc', 'ls', 'status'),
            show='headings',
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        self.tree.heading('cow_id', text='ID')
        self.tree.heading('jpn10', text='個体識別番号')
        self.tree.heading('milk_yield', text='乳量')
        self.tree.heading('fat', text='乳脂率')
        self.tree.heading('protein', text='乳蛋白率')
        self.tree.heading('scc', text='SCC')
        self.tree.heading('ls', text='LS')
        self.tree.heading('status', text='状態')
        
        self.tree.column('cow_id', width=80)
        self.tree.column('jpn10', width=120)
        self.tree.column('milk_yield', width=80)
        self.tree.column('fat', width=80)
        self.tree.column('protein', width=80)
        self.tree.column('scc', width=80)
        self.tree.column('ls', width=80)
        self.tree.column('status', width=150)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
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
            title="乳検CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_path = Path(file_path)
            self.file_path_var.set(str(self.csv_path))
            self._parse_csv()
    
    def _parse_csv(self):
        """CSVファイルをパース（farm_creator.pyのロジックを再利用）"""
        if not self.csv_path or not self.csv_path.exists():
            messagebox.showerror("エラー", "CSVファイルが見つかりません")
            return
        
        try:
            # CSVを読み込む（Shift-JIS）
            lines = []
            with open(self.csv_path, 'r', encoding='shift_jis') as f:
                for line in f:
                    lines.append(line.strip().split(','))
            
            # DataFrameに変換
            max_cols = max(len(row) for row in lines) if lines else 0
            for row in lines:
                while len(row) < max_cols:
                    row.append('')
            
            df = pd.DataFrame(lines, dtype=str)
            
            # 検定日を取得
            test_date = None
            if len(df) > 3:
                date_str = df.iloc[3, 0] if len(df.columns) > 0 else None
                if pd.notna(date_str):
                    date_str = str(date_str).strip()
                    try:
                        date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                        test_date = date_obj.strftime('%Y-%m-%d')
                    except:
                        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d']:
                            try:
                                date_obj = datetime.strptime(date_str, fmt)
                                test_date = date_obj.strftime('%Y-%m-%d')
                                break
                            except:
                                pass
            
            if not test_date:
                messagebox.showerror("エラー", "検定日が見つかりません")
                return
            
            self.test_date = test_date
            
            # ヘッダー行を取得（7行目）
            header_row_idx = 7
            if len(df) <= header_row_idx:
                messagebox.showerror("エラー", "ヘッダー行が見つかりません")
                return
            
            headers = df.iloc[header_row_idx].tolist()
            
            # ヘッダーから列位置を検索する関数
            def find_column_index(keywords: List[str]) -> Optional[int]:
                """ヘッダー行からキーワードを含む列のインデックスを検索"""
                for idx, header in enumerate(headers):
                    if pd.notna(header):
                        header_str = str(header).strip()
                        for keyword in keywords:
                            if keyword in header_str:
                                return idx
                return None
            
            # 列位置を指定（固定位置をフォールバックとして使用）
            def excel_col_to_index(col_name: str) -> int:
                result = 0
                for char in col_name:
                    result = result * 26 + (ord(char.upper()) - ord('A') + 1)
                return result - 1
            
            col_jpn10 = 0
            col_cow_id = 1
            
            # ヘッダーから列位置を検索（当日値の列）
            col_milk_yield = find_column_index(['乳量', '当日']) or excel_col_to_index('D')
            col_fat = find_column_index(['脂肪', '当日']) or excel_col_to_index('G')
            col_protein = find_column_index(['蛋白', '当日']) or excel_col_to_index('L')
            col_snf = find_column_index(['無脂', '当日']) or excel_col_to_index('J')
            col_scc = find_column_index(['体細胞', '当日']) or excel_col_to_index('Q')
            col_ls = find_column_index(['スコア', '当日']) or excel_col_to_index('X')
            col_mun = find_column_index(['MUN', '当日']) or excel_col_to_index('T')
            col_bhb = find_column_index(['BHB', '当日']) or excel_col_to_index('AA')
            
            # FA関連（後半の列にある可能性が高い）
            col_denovo_fa = find_column_index(['denovo', 'FA', '当日'])
            col_preformed_fa = find_column_index(['preformed', 'FA', '当日'])
            col_mixed_fa = find_column_index(['mixed', 'FA', '当日'])
            col_denovo_milk = find_column_index(['denovo', 'Milk', '当日'])
            
            # データ行を取得
            data_rows = []
            for idx in range(header_row_idx + 1, len(df)):
                row = df.iloc[idx].tolist()
                
                if not any(pd.notna(val) and str(val).strip() for val in row[:5]):
                    continue
                
                jpn10 = str(row[col_jpn10]).strip() if len(row) > col_jpn10 and pd.notna(row[col_jpn10]) else None
                if not jpn10 or len(jpn10) != 10 or not jpn10.isdigit():
                    continue
                
                cow_id_4digit = str(row[col_cow_id]).strip() if len(row) > col_cow_id and pd.notna(row[col_cow_id]) else None
                if cow_id_4digit and len(cow_id_4digit) >= 4:
                    cow_id = cow_id_4digit[-4:]
                else:
                    cow_id = jpn10[-4:]
                
                def get_float_value(col_idx: int) -> Optional[float]:
                    if col_idx is None or len(row) <= col_idx:
                        return None
                    val = row[col_idx]
                    if pd.notna(val):
                        try:
                            return float(str(val).strip())
                        except:
                            pass
                    return None
                
                def get_int_value(col_idx: int) -> Optional[int]:
                    if col_idx is None or len(row) <= col_idx:
                        return None
                    val = row[col_idx]
                    if pd.notna(val):
                        try:
                            val_str = str(val).strip()
                            val_str = val_str.replace('*', '').replace('**', '').replace('***', '')
                            if val_str:
                                return int(float(val_str))
                        except:
                            pass
                    return None
                
                milk_yield = get_float_value(col_milk_yield)
                fat = get_float_value(col_fat)
                protein = get_float_value(col_protein)
                snf = get_float_value(col_snf)
                scc = get_int_value(col_scc)
                ls_float = get_float_value(col_ls)
                ls = int(ls_float) if ls_float is not None else None
                mun = get_float_value(col_mun)
                bhb = get_float_value(col_bhb)
                denovo_fa = get_float_value(col_denovo_fa)
                preformed_fa = get_float_value(col_preformed_fa)
                mixed_fa = get_float_value(col_mixed_fa)
                denovo_milk = get_float_value(col_denovo_milk)
                
                data_rows.append({
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'milk_yield': milk_yield,
                    'fat': fat,
                    'protein': protein,
                    'snf': snf,
                    'scc': scc,
                    'ls': ls,
                    'mun': mun,
                    'bhb': bhb,
                    'denovo_fa': denovo_fa,
                    'preformed_fa': preformed_fa,
                    'mixed_fa': mixed_fa,
                    'denovo_milk': denovo_milk,
                })
            
            self.data_rows = data_rows
            self._update_preview()
            self.info_label.config(text=f"検定日: {test_date}, データ件数: {len(data_rows)}件")
            
        except Exception as e:
            logger.error(f"CSVパースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")
    
    def _update_preview(self):
        """プレビューを更新"""
        # 既存のアイテムをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # データを表示
        for row_data in self.data_rows:
            cow_id = row_data.get('cow_id', '')
            jpn10 = row_data.get('jpn10', '')
            
            # 既存の個体をチェック
            cow = self.db.get_cow_by_id(cow_id)
            if not cow:
                # JPN10で検索
                try:
                    cows = self.db.get_cows_by_jpn10(jpn10)
                    if cows:
                        cow = cows[0]
                except:
                    cow = None
            
            status = "新規" if not cow else "既存"
            
            self.tree.insert('', tk.END, values=(
                cow_id,
                jpn10,
                row_data.get('milk_yield') or '',
                row_data.get('fat') or '',
                row_data.get('protein') or '',
                row_data.get('scc') or '',
                row_data.get('ls') or '',
                status
            ))
    
    def _execute_import(self):
        """取り込みを実行"""
        if not self.test_date or not self.data_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return
        
        result = messagebox.askyesno(
            "確認",
            f"検定日 {self.test_date} の乳検データを {len(self.data_rows)} 件取り込みますか？\n"
            "既に同じ検定日のデータがある場合はスキップされます。"
        )
        
        if not result:
            return
        
        success_count = 0
        skip_count = 0
        error_count = 0
        error_cows = []  # エラー個体のリスト
        
        for row_data in self.data_rows:
            cow_id = row_data.get('cow_id')
            jpn10 = row_data.get('jpn10')
            
            # 個体を検索
            cow = self.db.get_cow_by_id(cow_id)
            if not cow:
                # JPN10で検索
                try:
                    cows = self.db.get_cows_by_jpn10(jpn10)
                    if cows:
                        cow = cows[0]
                except:
                    cow = None
            
            if not cow:
                error_count += 1
                error_cows.append({
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'reason': 'マスタに個体がありません'
                })
                logger.warning(f"個体が見つかりません: cow_id={cow_id}, jpn10={jpn10}")
                continue
            
            cow_auto_id = cow.get('auto_id')
            
            # 既存の乳検イベントをチェック
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            existing_milk_test = False
            for event in events:
                if event.get('event_number') == RuleEngine.EVENT_MILK_TEST:
                    if event.get('event_date') == self.test_date:
                        existing_milk_test = True
                        break
            
            if existing_milk_test:
                skip_count += 1
                continue
            
            # 乳検イベントを作成
            json_data = {}
            if row_data.get('milk_yield') is not None:
                json_data['milk_yield'] = row_data['milk_yield']
            if row_data.get('fat') is not None:
                json_data['fat'] = row_data['fat']
            if row_data.get('protein') is not None:
                json_data['protein'] = row_data['protein']
            if row_data.get('snf') is not None:
                json_data['snf'] = row_data.get('snf')
            if row_data.get('scc') is not None:
                json_data['scc'] = row_data['scc']
            if row_data.get('ls') is not None:
                json_data['ls'] = row_data['ls']
            if row_data.get('mun') is not None:
                json_data['mun'] = row_data.get('mun')
            if row_data.get('bhb') is not None:
                json_data['bhb'] = row_data.get('bhb')
            if row_data.get('denovo_fa') is not None:
                json_data['denovo_fa'] = row_data.get('denovo_fa')
            if row_data.get('preformed_fa') is not None:
                json_data['preformed_fa'] = row_data.get('preformed_fa')
            if row_data.get('mixed_fa') is not None:
                json_data['mixed_fa'] = row_data.get('mixed_fa')
            if row_data.get('denovo_milk') is not None:
                json_data['denovo_milk'] = row_data.get('denovo_milk')
            
            try:
                milk_event_data = {
                    'cow_auto_id': cow_auto_id,
                    'event_number': RuleEngine.EVENT_MILK_TEST,
                    'event_date': self.test_date,
                    'json_data': json_data if json_data else None,
                    'note': 'CSVから取り込み'
                }
                
                event_id = self.db.insert_event(milk_event_data)
                
                # RuleEngineで状態を更新
                self.rule_engine.on_event_added(event_id)
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_cows.append({
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'reason': f'イベント作成エラー: {str(e)}'
                })
                logger.error(f"乳検イベント作成エラー: cow_id={cow_id}, エラー={e}", exc_info=True)
        
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
            text=f"マスタに個体がなく、吸い込むことができなかった個体一覧（{len(error_cows)}件）",
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
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 左側：保存・印刷ボタン
        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side=tk.LEFT)
        
        ttk.Button(
            left_button_frame, 
            text="テキストファイルとして保存", 
            command=lambda: self._save_error_cows_to_file(error_cows, error_window)
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            left_button_frame, 
            text="印刷", 
            command=lambda: self._print_error_cows(error_cows, error_window)
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # 右側：閉じるボタン
        ttk.Button(button_frame, text="閉じる", command=error_window.destroy).pack(side=tk.RIGHT)
    
    def _save_error_cows_to_file(self, error_cows: List[Dict[str, Any]], parent_window: tk.Toplevel):
        """エラー個体の一覧をテキストファイルとして保存"""
        # 保存先を選択
        file_path = filedialog.asksaveasfilename(
            title="エラー個体一覧を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                # ヘッダー
                f.write("=" * 60 + "\n")
                f.write("エラー個体一覧\n")
                f.write("=" * 60 + "\n")
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"エラー件数: {len(error_cows)}件\n")
                f.write("=" * 60 + "\n\n")
                
                # データ
                f.write(f"{'ID':<10} {'個体識別番号':<15} {'エラー理由':<30}\n")
                f.write("-" * 60 + "\n")
                
                for error_cow in error_cows:
                    cow_id = error_cow.get('cow_id', '')
                    jpn10 = error_cow.get('jpn10', '')
                    reason = error_cow.get('reason', '')
                    f.write(f"{cow_id:<10} {jpn10:<15} {reason:<30}\n")
            
            messagebox.showinfo("保存完了", f"エラー個体一覧を保存しました:\n{file_path}")
            
        except Exception as e:
            logger.error(f"ファイル保存エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ファイルの保存に失敗しました:\n{e}")
    
    def _print_error_cows(self, error_cows: List[Dict[str, Any]], parent_window: tk.Toplevel):
        """エラー個体の一覧を印刷"""
        # 一時ファイルを作成
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8')
        temp_path = temp_file.name
        
        try:
            # テキストファイルに書き込み
            with open(temp_path, 'w', encoding='utf-8') as f:
                # ヘッダー
                f.write("=" * 60 + "\n")
                f.write("エラー個体一覧\n")
                f.write("=" * 60 + "\n")
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"エラー件数: {len(error_cows)}件\n")
                f.write("=" * 60 + "\n\n")
                
                # データ
                f.write(f"{'ID':<10} {'個体識別番号':<15} {'エラー理由':<30}\n")
                f.write("-" * 60 + "\n")
                
                for error_cow in error_cows:
                    cow_id = error_cow.get('cow_id', '')
                    jpn10 = error_cow.get('jpn10', '')
                    reason = error_cow.get('reason', '')
                    f.write(f"{cow_id:<10} {jpn10:<15} {reason:<30}\n")
            
            temp_file.close()
            
            # Windowsの場合：メモ帳で開いて印刷
            if platform.system() == 'Windows':
                try:
                    # メモ帳で開く（印刷ダイアログ付き）
                    subprocess.Popen(['notepad', '/p', temp_path], shell=True)
                    messagebox.showinfo("印刷", "印刷ダイアログが開きます。\n印刷を完了してください。")
                except Exception as e:
                    logger.error(f"印刷エラー: {e}", exc_info=True)
                    # フォールバック：メモ帳で開くだけ
                    try:
                        subprocess.Popen(['notepad', temp_path], shell=True)
                        messagebox.showinfo("印刷", "メモ帳でファイルを開きました。\nファイル > 印刷 から印刷してください。")
                    except Exception as e2:
                        messagebox.showerror("エラー", f"印刷に失敗しました:\n{e2}\n\n一時ファイル: {temp_path}")
            else:
                # Windows以外の場合：ファイルを開くだけ
                try:
                    if platform.system() == 'Darwin':  # macOS
                        subprocess.Popen(['open', temp_path])
                    else:  # Linux
                        subprocess.Popen(['xdg-open', temp_path])
                    messagebox.showinfo("印刷", "ファイルを開きました。\n印刷機能から印刷してください。")
                except Exception as e:
                    messagebox.showinfo("情報", f"ファイルを開けませんでした。\n一時ファイル: {temp_path}\n\n手動で開いて印刷してください。")
        
        except Exception as e:
            logger.error(f"印刷エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"印刷に失敗しました:\n{e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()

