"""
FALCON2 - レポート表表示ウィンドウ
分析結果の表を表示する専用ウインドウ（DC305 Reports画面風）
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Any, List, Optional
from datetime import datetime
import logging


class ReportTableWindow:
    """レポート表表示ウィンドウ（Toplevel）"""
    
    def __init__(self, parent: tk.Tk, report_title: str, 
                 columns: List[str], rows: List[List[Any]],
                 conditions: Optional[str] = None,
                 period: Optional[str] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            report_title: レポート名
            columns: カラム名のリスト
            rows: 行データのリスト（各要素はカラム数と一致する必要がある）
            conditions: 集計条件（オプション）
            period: 期間（オプション）
        """
        self.parent = parent
        self.report_title = report_title
        self.columns = columns
        self.rows = rows
        self._original_rows = list(rows)
        self._display_rows = list(rows)
        self._sort_state = {col: None for col in columns}
        self.conditions = conditions
        self.period = period
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title(report_title)
        self.window.geometry("900x600")
        
        self._create_widgets()
        self._load_data()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ヘッダー情報
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # レポート名
        title_label = ttk.Label(
            header_frame,
            text=self.report_title,
            font=("", 12, "bold")
        )
        title_label.pack(anchor=tk.W)
        
        # 集計条件
        if self.conditions:
            condition_label = ttk.Label(
                header_frame,
                text=f"集計条件：{self.conditions}",
                font=("", 9)
            )
            condition_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 期間
        if self.period:
            period_label = ttk.Label(
                header_frame,
                text=f"期間：{self.period}",
                font=("", 9)
            )
            period_label.pack(anchor=tk.W, pady=(2, 0))
        
        # 区切り線
        separator = ttk.Separator(main_frame, orient=tk.HORIZONTAL)
        separator.pack(fill=tk.X, pady=(0, 10))
        
        # 表（Treeview）
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # スクロールバー
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview
        self.tree = ttk.Treeview(
            table_frame,
            columns=self.columns,
            show='headings',
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        
        # カラム設定
        for col in self.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self._on_header_click(c))
            # 数値カラムは右寄せ、それ以外は左寄せ
            if col == "産次" or col.isdigit() or "月" in col:
                self.tree.column(col, anchor=tk.E, width=80)  # 右寄せ
            else:
                self.tree.column(col, anchor=tk.W, width=100)  # 左寄せ
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        excel_button = ttk.Button(
            button_frame,
            text="Excelコピー",
            command=self._copy_to_excel
        )
        excel_button.pack(side=tk.LEFT, padx=5)
        
        csv_button = ttk.Button(
            button_frame,
            text="CSV保存",
            command=self._save_csv
        )
        csv_button.pack(side=tk.LEFT, padx=5)
        
        close_button = ttk.Button(
            button_frame,
            text="閉じる",
            command=self.window.destroy
        )
        close_button.pack(side=tk.RIGHT, padx=5)
    
    def _load_data(self):
        """データをテーブルにロード"""
        # 既存のデータをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # データを追加
        for row in self._display_rows:
            # 数値を文字列に変換
            row_str = [str(val) for val in row]
            self.tree.insert('', tk.END, values=row_str)

    def _on_header_click(self, column: str):
        """列ヘッダークリックでソート状態を切り替え"""
        current_state = self._sort_state.get(column)
        if current_state is None:
            # 他列はリセット
            for col in self._sort_state:
                self._sort_state[col] = None
            self._sort_state[column] = 'asc'
        elif current_state == 'asc':
            self._sort_state[column] = 'desc'
        else:
            self._sort_state[column] = None
        self._refresh_table()

    def _refresh_table(self):
        """ソート状態に応じて表示を更新"""
        sort_column = None
        sort_direction = None
        for col, state in self._sort_state.items():
            if state is not None:
                sort_column = col
                sort_direction = state
                break
        
        if not sort_column or sort_direction is None:
            self._display_rows = list(self._original_rows)
        else:
            col_index = self.columns.index(sort_column)
            reverse = (sort_direction == 'desc')
            self._display_rows = sorted(
                self._original_rows,
                key=lambda row: self._get_sort_key(row[col_index] if col_index < len(row) else None),
                reverse=reverse
            )
        self._load_data()

    @staticmethod
    def _get_sort_key(value: Any):
        """ソート用キー（数値優先）"""
        if value is None:
            return (1, "")
        value_str = str(value).strip()
        if value_str == "":
            return (1, "")
        try:
            num = float(value_str.replace(",", ""))
            return (0, num)
        except ValueError:
            return (0, value_str.lower())
    
    def _copy_to_excel(self):
        """Excelコピー（タブ区切りTSV）"""
        try:
            # ヘッダー行
            lines = ['\t'.join(self.columns)]
            
            # データ行
            for row in self._display_rows:
                row_str = [str(val) for val in row]
                lines.append('\t'.join(row_str))
            
            # クリップボードにコピー
            text = '\n'.join(lines)
            self.window.clipboard_clear()
            self.window.clipboard_append(text)
            self.window.update()  # クリップボードを更新
            
            messagebox.showinfo("情報", "Excelコピー用データ（タブ区切り）をクリップボードにコピーしました")
            
        except Exception as e:
            logging.error(f"Excelコピーエラー: {e}")
            messagebox.showerror("エラー", f"Excelコピーに失敗しました: {e}")
    
    def _save_csv(self):
        """CSV保存（UTF-8）"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )
            
            if not filename:
                return
            
            # UTF-8で保存
            with open(filename, 'w', encoding='utf-8-sig') as f:  # BOM付きUTF-8（Excel互換）
                # ヘッダー行
                f.write(','.join(self.columns) + '\n')
                
                # データ行
            for row in self._display_rows:
                    row_str = [str(val) for val in row]
                    f.write(','.join(row_str) + '\n')
            
            messagebox.showinfo("情報", f"CSVファイルを保存しました: {filename}")
            
        except Exception as e:
            logging.error(f"CSV保存エラー: {e}")
            messagebox.showerror("エラー", f"CSV保存に失敗しました: {e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()

