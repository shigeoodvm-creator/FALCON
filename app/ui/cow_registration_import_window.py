"""
FALCON2 - 新規個体登録取り込みウィンドウ
CSV / Excel ファイルから新規個体を一括登録する
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
from modules.cow_registration_excel_template import prompt_save_bulk_cow_registration_template

logger = logging.getLogger(__name__)


class CowRegistrationImportWindow:
    """新規個体登録取り込みウィンドウ（CSV / Excel 対応）"""

    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine, farm_path: Path):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("新規一括導入")
        self.window.geometry("1200x700")

        self.file_path: Optional[Path] = None
        self.data_rows: List[Dict[str, Any]] = []

        self._create_widgets()

    # ------------------------------------------------------------------ #
    #  UI 構築
    # ------------------------------------------------------------------ #

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ファイル選択
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択（CSV / Excel）", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=70, state='readonly').pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(file_frame, text="参照...", command=self._select_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(file_frame, text="テンプレートを出力", command=self._export_template).pack(side=tk.LEFT)

        # プレビュー
        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー（編集可能）", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        columns = ('cow_id', 'jpn10', 'brd', 'pen', 'birth_date', 'lact',
                   'clvd', 'last_ai_date', 'ai_count', 'dam', 'pregnant_status', 'intro_date')
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show='headings',
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )

        col_defs = [
            ('cow_id',          'ID（4桁）',        80),
            ('jpn10',           'JPN10（個体識別番号）', 130),
            ('brd',             '品種',             100),
            ('pen',             '群',               60),
            ('birth_date',      '生年月日',          100),
            ('lact',            '産次',              50),
            ('clvd',            '分娩月日',          100),
            ('last_ai_date',    '最終授精月日',       110),
            ('ai_count',        '授精回数',           70),
            ('dam',             '母牛識別番号',       120),
            ('pregnant_status', '妊娠状態',          110),
            ('intro_date',      '導入日',            100),
        ]
        for col, heading, width in col_defs:
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        self.tree.bind('<Double-Button-1>', self._on_cell_double_click)

        # 情報表示
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        self.info_label = ttk.Label(info_frame, text="CSVまたはExcelファイルを選択してください")
        self.info_label.pack(side=tk.LEFT)

        # ボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="取り込み実行", command=self._execute_import).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    # ------------------------------------------------------------------ #
    #  ファイル選択
    # ------------------------------------------------------------------ #

    def _select_file(self):
        file_path = filedialog.askopenfilename(
            title="個体一括導入ファイルを選択",
            filetypes=[
                ("対応ファイル", "*.csv *.xlsx *.xls"),
                ("CSVファイル", "*.csv"),
                ("Excelファイル", "*.xlsx *.xls"),
                ("すべてのファイル", "*.*"),
            ]
        )
        if not file_path:
            return

        self.file_path = Path(file_path)
        self.file_path_var.set(str(self.file_path))

        ext = self.file_path.suffix.lower()
        if ext in ('.xlsx', '.xls'):
            self._parse_excel()
        else:
            self._parse_csv()

    # ------------------------------------------------------------------ #
    #  ユーティリティ
    # ------------------------------------------------------------------ #

    def _normalize_date(self, date_str: str) -> Optional[str]:
        if not date_str or pd.isna(date_str):
            return None
        date_str = str(date_str).strip()
        if not date_str or '#' in date_str:
            return None

        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y%m%d', '%m/%d/%Y', '%d/%m/%Y'):
            try:
                return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
            except Exception:
                continue

        try:
            obj = pd.to_datetime(date_str, errors='coerce')
            if pd.notna(obj):
                return obj.strftime('%Y-%m-%d')
        except Exception:
            pass
        return None

    def _find_column_index(self, headers: List[str], keywords: List[str]) -> Optional[int]:
        # 完全一致優先
        for idx, header in enumerate(headers):
            if pd.notna(header):
                h = str(header).strip().replace(' ', '').replace('\u3000', '')
                for kw in keywords:
                    k = kw.replace(' ', '').replace('\u3000', '')
                    if h == k or h.upper() == k.upper():
                        return idx
        # 部分一致
        for idx, header in enumerate(headers):
            if pd.notna(header):
                h = str(header).strip().upper()
                for kw in keywords:
                    if kw.upper() in h or h in kw.upper():
                        return idx
        return None

    @staticmethod
    def _get_value(row, col_idx: Optional[int], num_cols: int, default: str = '') -> str:
        if col_idx is None or col_idx >= num_cols:
            return default
        try:
            val = row.iloc[col_idx]
            if pd.isna(val):
                return default
            s = str(val).strip()
            if not s or s.lower() in ('nan', 'none'):
                return default
            return s.replace('########', '')
        except Exception:
            return default

    def _rows_from_df(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """DataFrameから共通のデータ行リストを生成する"""
        headers = df.columns.tolist()
        num_cols = len(df.columns)
        get = lambda row, col, default='': self._get_value(row, col, num_cols, default)

        col_jpn10      = self._find_column_index(headers, ['JPN10', '個体識別番号', '個体識別', '識別番号', 'JPN'])
        col_brd        = self._find_column_index(headers, ['品種', 'BRD', '品種名'])
        col_pen        = self._find_column_index(headers, ['群(PEN)', '群', 'PEN', '群番号', '群コード', '群名'])
        col_birth_date = self._find_column_index(headers, ['生年月日', 'BIRTH', 'BTHD', '誕生日'])
        col_lact       = self._find_column_index(headers, ['産次', 'LACT', 'LACTATION'])
        col_clvd       = self._find_column_index(headers, ['最終分娩日', '分娩月日', '分娩日', '分娩', 'CLVD', 'CALV'])
        col_last_ai    = self._find_column_index(headers, ['最終授精日', '最終AI日', '最終AI', 'LAST_AI', 'LASTAI', '授精日'])
        col_ai_sire    = self._find_column_index(headers, ['最終授精SIRE', 'SIRE', '種雄牛'])
        col_ai_count   = self._find_column_index(headers, ['授精回数', 'AI回数', 'AI_COUNT', 'AICNT'])
        col_pregnant   = self._find_column_index(headers, ['妊娠状態', '妊娠', 'PREGNANT'])
        col_dam        = self._find_column_index(headers, ['母牛識別番号', '母牛識別番', '母牛', 'DAM'])
        col_intro_date = self._find_column_index(headers, ['導入日', '入牧日', 'INTRO'])
        col_memo       = self._find_column_index(headers, ['メモ', 'NOTE', '備考'])

        today = datetime.now().strftime('%Y-%m-%d')
        data_rows = []

        for _, row in df.iterrows():
            if row.isna().all():
                continue

            jpn10 = get(row, col_jpn10)
            if not jpn10 or len(jpn10) != 10 or not jpn10.isdigit():
                continue

            cow_id = jpn10[5:9].zfill(4)

            brd = get(row, col_brd) if col_brd is not None else 'ホルスタイン'
            pen = get(row, col_pen) if col_pen is not None else '1'
            birth_date = self._normalize_date(get(row, col_birth_date)) or ''

            lact_str = get(row, col_lact)
            lact = ''
            if lact_str:
                try:
                    lact = int(float(lact_str))
                except Exception:
                    pass

            clvd          = self._normalize_date(get(row, col_clvd)) or ''
            last_ai_date  = self._normalize_date(get(row, col_last_ai)) or ''
            last_ai_sire  = get(row, col_ai_sire) or ''

            ai_count_str = get(row, col_ai_count)
            ai_count = ''
            if ai_count_str:
                try:
                    ai_count = int(float(ai_count_str))
                except Exception:
                    pass

            pregnant_status = get(row, col_pregnant)
            if not pregnant_status and last_ai_date:
                pregnant_status = '妊娠鑑定待ち'

            dam        = get(row, col_dam) or ''
            intro_date = self._normalize_date(get(row, col_intro_date)) or today
            memo       = get(row, col_memo) or ''

            data_rows.append({
                'cow_id':          cow_id,
                'jpn10':           jpn10,
                'brd':             brd,
                'pen':             pen,
                'birth_date':      birth_date,
                'lact':            lact,
                'clvd':            clvd,
                'last_ai_date':    last_ai_date,
                'last_ai_sire':    last_ai_sire,
                'ai_count':        ai_count,
                'dam':             dam,
                'pregnant_status': pregnant_status,
                'intro_date':      intro_date,
                'memo':            memo,
            })

        return data_rows

    # ------------------------------------------------------------------ #
    #  CSV パース
    # ------------------------------------------------------------------ #

    def _parse_csv(self):
        if not self.file_path or not self.file_path.exists():
            messagebox.showerror("エラー", "ファイルが見つかりません")
            return
        try:
            df = None
            for enc in ('shift_jis', 'utf-8', 'cp932'):
                try:
                    df = pd.read_csv(self.file_path, encoding=enc, dtype=str)
                    break
                except Exception:
                    continue
            if df is None:
                messagebox.showerror("エラー", "CSVファイルの読み込みに失敗しました")
                return

            self.data_rows = self._rows_from_df(df)
            self._update_preview()

        except Exception as e:
            logger.error(f"CSVパースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")

    # ------------------------------------------------------------------ #
    #  Excel パース
    # ------------------------------------------------------------------ #

    def _parse_excel(self):
        if not self.file_path or not self.file_path.exists():
            messagebox.showerror("エラー", "ファイルが見つかりません")
            return
        try:
            # "個体情報"シートを優先、なければ最初のシート
            xl = pd.ExcelFile(self.file_path, engine='openpyxl')
            sheet = '個体情報' if '個体情報' in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(self.file_path, sheet_name=sheet, dtype=str, engine='openpyxl')

            self.data_rows = self._rows_from_df(df)
            self._update_preview()

        except ImportError:
            messagebox.showerror(
                "エラー",
                "Excelを読み込むには openpyxl が必要です。\n"
                "コマンドプロンプトで: pip install openpyxl"
            )
        except Exception as e:
            logger.error(f"Excelパースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"Excelファイルの読み込みに失敗しました: {e}")

    # ------------------------------------------------------------------ #
    #  テンプレート出力
    # ------------------------------------------------------------------ #

    def _export_template(self):
        prompt_save_bulk_cow_registration_template(parent=self.window)

    # ------------------------------------------------------------------ #
    #  プレビュー
    # ------------------------------------------------------------------ #

    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        existing_count = 0
        new_count = 0

        for row_data in self.data_rows:
            jpn10 = row_data.get('jpn10', '')
            try:
                existing_cows = self.db.get_cows_by_jpn10(jpn10)
            except Exception:
                existing_cows = []

            if existing_cows:
                tag = '既存'
                existing_count += 1
            else:
                tag = '新規'
                new_count += 1

            lact = row_data.get('lact', '')
            ai_count = row_data.get('ai_count', '')
            self.tree.insert('', tk.END, values=(
                row_data.get('cow_id', ''),
                jpn10,
                row_data.get('brd', ''),
                row_data.get('pen', ''),
                row_data.get('birth_date', ''),
                lact if lact != '' else '',
                row_data.get('clvd', ''),
                row_data.get('last_ai_date', ''),
                ai_count if ai_count != '' else '',
                row_data.get('dam', ''),
                row_data.get('pregnant_status', ''),
                row_data.get('intro_date', ''),
            ), tags=(tag,))

        self.tree.tag_configure('既存', foreground='gray')
        self.tree.tag_configure('新規', foreground='black')

        total = len(self.data_rows)
        self.info_label.config(
            text=f"データ件数: {total}件 | "
                 f"既存個体（スキップ）: {existing_count}件 | "
                 f"導入予定: {new_count}件（ダブルクリックで編集可能）"
        )

    # ------------------------------------------------------------------ #
    #  セル編集
    # ------------------------------------------------------------------ #

    def _on_cell_double_click(self, event):
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return

        column = self.tree.identify_column(event.x)
        if not column:
            return
        col_idx = int(column.replace('#', '')) - 1
        if col_idx < 0 or col_idx >= len(self.tree['columns']):
            return
        col_name = self.tree['columns'][col_idx]

        if col_name == 'pregnant_status':
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                if self.data_rows[item_index].get('last_ai_date'):
                    self._edit_pregnant_status(item)
                else:
                    messagebox.showinfo("情報", "最終授精月日が入力されていないため、妊娠状態を選択できません。")
            return

        values = list(self.tree.item(item, 'values'))
        current_value = values[col_idx] if col_idx < len(values) else ''

        edit_window = tk.Toplevel(self.window)
        edit_window.title(f"{self.tree.heading(col_name, 'text')}を編集")
        edit_window.geometry("300x100")

        ttk.Label(edit_window, text="新しい値を入力してください:").pack(pady=10)
        entry = ttk.Entry(edit_window, width=30)
        entry.pack(pady=5)
        entry.insert(0, str(current_value))
        entry.focus()

        def save_edit():
            new_value = entry.get().strip()

            if col_name in ('birth_date', 'clvd', 'last_ai_date', 'intro_date'):
                normalized = self._normalize_date(new_value)
                if normalized:
                    new_value = normalized
                elif new_value:
                    messagebox.showerror("エラー", "日付の形式が正しくありません（YYYY-MM-DD形式）")
                    return

            if col_name in ('lact', 'ai_count'):
                if new_value:
                    try:
                        int(new_value)
                    except ValueError:
                        messagebox.showerror("エラー", "数値を入力してください")
                        return

            values[col_idx] = new_value
            self.tree.item(item, values=values)

            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                self.data_rows[item_index][col_name] = new_value

                if col_name == 'last_ai_date':
                    if new_value and not self.data_rows[item_index].get('pregnant_status'):
                        self.data_rows[item_index]['pregnant_status'] = '妊娠鑑定待ち'
                        values[-2] = '妊娠鑑定待ち'  # pregnant_status は intro_date の手前
                        self.tree.item(item, values=values)
                    elif not new_value:
                        self.data_rows[item_index]['pregnant_status'] = ''
                        values[-2] = ''
                        self.tree.item(item, values=values)

            edit_window.destroy()

        ttk.Button(edit_window, text="保存", command=save_edit).pack(pady=5)
        entry.bind('<Return>', lambda e: save_edit())

    def _edit_pregnant_status(self, item):
        values = list(self.tree.item(item, 'values'))
        # pregnant_status は columns の index 10
        ps_col_idx = list(self.tree['columns']).index('pregnant_status')
        current_status = values[ps_col_idx] if ps_col_idx < len(values) else ''

        status_window = tk.Toplevel(self.window)
        status_window.title("妊娠状態を選択")
        status_window.geometry("250x200")

        ttk.Label(status_window, text="妊娠状態を選択してください:").pack(pady=10)
        status_var = tk.StringVar(value=current_status)

        for s in ('妊娠鑑定待ち', '受胎なし', '妊娠', '乾乳'):
            ttk.Radiobutton(status_window, text=s, variable=status_var, value=s).pack(anchor=tk.W, padx=20, pady=4)

        def save_status():
            new_status = status_var.get()
            values[ps_col_idx] = new_status
            self.tree.item(item, values=values)
            item_index = self.tree.index(item)
            if 0 <= item_index < len(self.data_rows):
                self.data_rows[item_index]['pregnant_status'] = new_status
            status_window.destroy()

        ttk.Button(status_window, text="保存", command=save_status).pack(pady=10)

    # ------------------------------------------------------------------ #
    #  取り込み実行
    # ------------------------------------------------------------------ #

    def _execute_import(self):
        if not self.data_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return

        import_count = sum(
            1 for row in self.data_rows
            if not self._get_existing_cow(row.get('jpn10', ''))
        )

        if not messagebox.askyesno(
            "確認",
            f"導入予定: {import_count}件の個体を登録しますか？\n"
            f"（既存個体: {len(self.data_rows) - import_count}件はスキップされます）"
        ):
            return

        today = datetime.now().strftime('%Y-%m-%d')
        success_count = skip_count = error_count = 0
        error_cows = []

        for row_data in self.data_rows:
            jpn10  = row_data.get('jpn10', '')
            cow_id = row_data.get('cow_id', '')

            if self._get_existing_cow(jpn10):
                skip_count += 1
                logger.info(f"既存個体をスキップ: jpn10={jpn10}")
                continue

            try:
                clvd            = row_data.get('clvd', '')
                last_ai_date    = row_data.get('last_ai_date', '')
                last_ai_sire    = row_data.get('last_ai_sire', '')
                pregnant_status = row_data.get('pregnant_status', '')
                intro_date      = row_data.get('intro_date') or today

                # RC 決定
                if pregnant_status == '妊娠':
                    initial_rc = RuleEngine.RC_PREGNANT
                elif pregnant_status == '乾乳':
                    initial_rc = RuleEngine.RC_DRY
                elif last_ai_date and pregnant_status == '妊娠鑑定待ち':
                    initial_rc = RuleEngine.RC_BRED
                elif last_ai_date and pregnant_status == '受胎なし':
                    initial_rc = RuleEngine.RC_OPEN
                elif clvd and not last_ai_date:
                    initial_rc = RuleEngine.RC_FRESH
                else:
                    initial_rc = RuleEngine.RC_OPEN

                brd_val = (row_data.get('brd') or '').strip()
                pen_val = (row_data.get('pen') or '').strip()

                cow_data = {
                    'cow_id': cow_id,
                    'jpn10':  jpn10,
                    'brd':    brd_val or None,
                    'bthd':   row_data.get('birth_date') or None,
                    'entr':   intro_date,
                    'lact':   int(row_data.get('lact', 0)) if row_data.get('lact') != '' else 0,
                    'clvd':   clvd or None,
                    'rc':     initial_rc,
                    'pen':    pen_val or None,
                    'frm':    None,
                }
                cow_auto_id = self.db.insert_cow(cow_data)

                # 導入イベント
                src = 'excel_import' if self.file_path and self.file_path.suffix.lower() in ('.xlsx', '.xls') else 'csv_import'
                intro_event = {
                    "cow_auto_id":  cow_auto_id,
                    "event_number": RuleEngine.EVENT_IN,
                    "event_date":   intro_date,
                    "json_data": {
                        "birth_date":    row_data.get('birth_date'),
                        "lactation":     row_data.get('lact', 0),
                        "calving_date":  clvd,
                        "last_ai_date":  last_ai_date,
                        "ai_count":      row_data.get('ai_count', 0),
                        "dam":           row_data.get('dam'),
                        "source":        src,
                    },
                    "note": row_data.get('memo') or "一括取り込み",
                }
                event_id = self.db.insert_event(intro_event)
                self.rule_engine.on_event_added(event_id)

                # 分娩イベント（baseline）
                if clvd:
                    calv_id = self.db.insert_event({
                        "cow_auto_id":  cow_auto_id,
                        "event_number": RuleEngine.EVENT_CALV,
                        "event_date":   clvd,
                        "json_data":    {"baseline_calving": True},
                        "note":         "取り込み時の分娩（baseline）",
                    })
                    self.rule_engine.on_event_added(calv_id)

                # AI イベント
                if last_ai_date:
                    ai_count_val = row_data.get('ai_count', '')
                    try:
                        ai_count_val = int(ai_count_val) if ai_count_val != '' else 1
                    except Exception:
                        ai_count_val = 1

                    ai_json: Dict[str, Any] = {"ai_count": ai_count_val}
                    if last_ai_sire:
                        ai_json["sire"] = last_ai_sire
                    if pregnant_status == '受胎なし':
                        ai_json['result'] = 'O'
                    elif pregnant_status in ('妊娠', '乾乳'):
                        ai_json['result'] = 'P'

                    ai_id = self.db.insert_event({
                        "cow_auto_id":  cow_auto_id,
                        "event_number": RuleEngine.EVENT_AI,
                        "event_date":   last_ai_date,
                        "json_data":    ai_json,
                        "note":         f"授精回数: {ai_count_val}",
                    })
                    self.rule_engine.on_event_added(ai_id)

                    # 妊娠プラスイベント
                    if pregnant_status in ('妊娠', '乾乳'):
                        preg_id = self.db.insert_event({
                            "cow_auto_id":  cow_auto_id,
                            "event_number": RuleEngine.EVENT_PDP,
                            "event_date":   last_ai_date,
                            "json_data":    {"twin": False},
                            "note":         "取り込み時の妊娠",
                        })
                        self.rule_engine.on_event_added(preg_id)

                    # 乾乳イベント
                    if pregnant_status == '乾乳':
                        dry_id = self.db.insert_event({
                            "cow_auto_id":  cow_auto_id,
                            "event_number": RuleEngine.EVENT_DRY,
                            "event_date":   intro_date,
                            "json_data":    {},
                            "note":         "取り込み時の乾乳",
                        })
                        self.rule_engine.on_event_added(dry_id)

                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                success_count += 1

            except Exception as e:
                error_count += 1
                error_cows.append({'cow_id': cow_id, 'jpn10': jpn10, 'reason': str(e)})
                logger.error(f"個体登録エラー: cow_id={cow_id}, jpn10={jpn10}: {e}", exc_info=True)

        result_msg = (
            f"取り込みが完了しました\n"
            f"成功: {success_count}件\n"
            f"スキップ: {skip_count}件\n"
            f"エラー: {error_count}件"
        )

        if error_cows:
            result_msg += f"\n\nエラー個体が {len(error_cows)} 件あります。詳細を表示しますか？"
            if messagebox.askyesno("完了", result_msg):
                self._show_error_cows_window(error_cows)
        else:
            messagebox.showinfo("完了", result_msg)

        self._update_preview()

    def _get_existing_cow(self, jpn10: str):
        try:
            return self.db.get_cows_by_jpn10(jpn10)
        except Exception:
            return []

    # ------------------------------------------------------------------ #
    #  エラー個体一覧ウィンドウ
    # ------------------------------------------------------------------ #

    def _show_error_cows_window(self, error_cows: List[Dict[str, Any]]):
        w = tk.Toplevel(self.window)
        w.title("エラー個体一覧")
        w.geometry("600x400")

        main_frame = ttk.Frame(w, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text=f"登録エラーが発生した個体一覧（{len(error_cows)}件）",
                  font=("", 10, "bold")).pack(pady=(0, 10))

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        sy = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        sx.pack(side=tk.BOTTOM, fill=tk.X)

        tree = ttk.Treeview(tree_frame, columns=('cow_id', 'jpn10', 'reason'),
                             show='headings', yscrollcommand=sy.set, xscrollcommand=sx.set)
        tree.heading('cow_id', text='ID')
        tree.heading('jpn10',  text='個体識別番号')
        tree.heading('reason', text='エラー理由')
        tree.column('cow_id', width=100)
        tree.column('jpn10',  width=150)
        tree.column('reason', width=300)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.config(command=tree.yview)
        sx.config(command=tree.xview)

        for ec in error_cows:
            tree.insert('', tk.END, values=(ec.get('cow_id', ''), ec.get('jpn10', ''), ec.get('reason', '')))

        ttk.Button(main_frame, text="閉じる", command=w.destroy).pack(pady=10)

    # ------------------------------------------------------------------ #

    def show(self):
        self.window.focus_set()
