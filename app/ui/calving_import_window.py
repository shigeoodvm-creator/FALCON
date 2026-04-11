"""
分娩イベント CSV/Excel 取り込み（繁殖検診フロー用）
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from ui.ai_import_window import (
    _find_column_index,
    _normalize_date_cell,
    _extract_jpn10,
    _to_cow_id,
)

logger = logging.getLogger(__name__)

_CALV_DIFF_TEXT = {
    "自然分娩": 1,
    "介助": 2,
    "難産": 3,
    "獣医師による難産": 4,
    "帝王切開": 5,
}


def _parse_calving_difficulty(raw: str) -> Optional[int]:
    if not raw or not str(raw).strip():
        return None
    s = str(raw).strip()
    if s.isdigit():
        v = int(s)
        if 1 <= v <= 5:
            return v
    return _CALV_DIFF_TEXT.get(s)


def _parse_stillborn(raw: str) -> bool:
    if raw is None:
        return False
    s = str(raw).strip().lower()
    if s in ("1", "true", "yes", "y", "はい", "死産"):
        return True
    return False


class CalvingImportWindow:
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine, farm_path: Path):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("分娩データ取り込み")
        self.window.geometry("1100x620")

        self.file_path: Optional[Path] = None
        self.data_rows: List[Dict[str, Any]] = []

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択（Excel / CSV）", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=55, state="readonly").pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(file_frame, text="参照...", command=self._select_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(file_frame, text="テンプレート出力", command=self._export_template).pack(side=tk.LEFT)

        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        cols = ("cow_id", "jpn10", "date", "diff", "calves", "status")
        self.tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        self.tree.heading("cow_id", text="ID")
        self.tree.heading("jpn10", text="個体識別番号")
        self.tree.heading("date", text="分娩日")
        self.tree.heading("diff", text="難易度")
        self.tree.heading("calves", text="子牛概要")
        self.tree.heading("status", text="状態")
        for c in cols:
            self.tree.column(c, width=120)
        self.tree.column("calves", width=280)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)

        self.info_label = ttk.Label(main_frame, text="ファイルを選択してください")
        self.info_label.pack(anchor=tk.W)

        bf = ttk.Frame(main_frame)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="取り込み実行", command=self._execute_import).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bf, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    def _export_template(self):
        from modules.reproduction_flow_templates import write_calving_template

        path = filedialog.asksaveasfilename(
            parent=self.window,
            title="分娩テンプレートの保存先",
            defaultextension=".xlsx",
            filetypes=[("Excel（推奨）", "*.xlsx"), ("CSV", "*.csv"), ("すべて", "*.*")],
            initialfile="繁殖検診_分娩テンプレート.xlsx",
        )
        if not path:
            return
        try:
            write_calving_template(Path(path))
            messagebox.showinfo("完了", f"テンプレートを保存しました:\n{path}")
        except Exception as e:
            logger.error(e, exc_info=True)
            messagebox.showerror("エラー", str(e))

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="分娩データファイルを選択",
            filetypes=[("Excel / CSV", "*.xlsx;*.xls;*.csv"), ("Excel", "*.xlsx;*.xls"), ("CSV", "*.csv"), ("すべて", "*.*")],
        )
        if path:
            self.file_path = Path(path)
            self.file_path_var.set(str(self.file_path))
            self._parse_file()

    def _parse_file(self):
        if not self.file_path or not self.file_path.exists():
            messagebox.showerror("エラー", "ファイルが見つかりません")
            return
        try:
            if self.file_path.suffix.lower() == ".csv":
                self._parse_csv()
            elif self.file_path.suffix.lower() in (".xlsx", ".xls"):
                self._parse_excel()
            else:
                messagebox.showerror("エラー", "対応形式は CSV / .xlsx / .xls です")
        except Exception as e:
            logger.error(e, exc_info=True)
            messagebox.showerror("エラー", f"読み込みに失敗しました:\n{e}")

    def _read_csv_lines(self) -> List[List[str]]:
        import csv as csv_mod

        for enc in ("utf-8", "utf-8-sig", "shift_jis", "cp932"):
            try:
                with open(self.file_path, "r", encoding=enc, newline="") as f:
                    return [row for row in csv_mod.reader(f)]
            except UnicodeDecodeError:
                continue
        raise ValueError("文字コードを判別できませんでした")

    def _parse_csv(self):
        lines = self._read_csv_lines()
        if not lines:
            messagebox.showerror("エラー", "CSVが空です")
            return
        max_cols = max(len(row) for row in lines)
        for row in lines:
            while len(row) < max_cols:
                row.append("")
        headers = [str(h).strip() for h in lines[0]]
        self._build_data_rows(headers, lines[1:])

    def _parse_excel(self):
        import pandas as pd

        try:
            df = pd.read_excel(self.file_path, header=0, sheet_name=0)
        except Exception:
            df = pd.read_excel(self.file_path, header=0, sheet_name=0, engine="openpyxl")
        if df is None or df.empty:
            self.data_rows = []
            self._update_preview()
            return

        def _cell(v):
            if v is None or (hasattr(pd, "isna") and pd.isna(v)):
                return ""
            if hasattr(v, "strftime"):
                try:
                    return v.strftime("%Y-%m-%d")
                except Exception:
                    pass
            if isinstance(v, (int, float)) and 20000 <= v <= 50000:
                try:
                    d = datetime(1899, 12, 30) + timedelta(days=int(v))
                    return d.strftime("%Y-%m-%d")
                except Exception:
                    pass
            return str(v).strip()

        headers = [str(h).strip() if h is not None else "" for h in df.columns.tolist()]
        rows = []
        for _, r in df.iterrows():
            line = [_cell(v) for v in r.tolist()]
            if line and str(line[0]).strip().startswith("#"):
                continue
            rows.append(line)
        self._build_data_rows(headers, rows)

    def _collect_calf(self, row: List[str], idx: int, headers: List[str]) -> Optional[Dict[str, Any]]:
        """1-based calf index: look for calf{N}_breed columns or legacy fixed positions."""
        keys_breed = [f"calf{idx}_breed", f"子牛{idx}品種", f"子牛{idx}・品種"]
        keys_sex = [f"calf{idx}_sex", f"子牛{idx}性別", f"子牛{idx}・性別"]
        keys_sb = [f"calf{idx}_stillborn", f"子牛{idx}死産", f"子牛{idx}・死産"]
        cb = _find_column_index(headers, keys_breed)
        cs = _find_column_index(headers, keys_sex)
        cst = _find_column_index(headers, keys_sb)
        if cb is None and cs is None:
            return None
        breed = ""
        sex = ""
        sb = False
        if cb is not None and cb < len(row):
            breed = str(row[cb]).strip()
        if cs is not None and cs < len(row):
            sex = str(row[cs]).strip()
        if cst is not None and cst < len(row):
            sb = _parse_stillborn(row[cst])
        if not breed and not sex and not sb:
            return None
        return {"breed": breed, "sex": sex, "stillborn": sb}

    def _build_data_rows(self, headers: List[str], rows: List[List[str]]):
        col_id = _find_column_index(headers, ["cow_id", "牛のID", "個体ID", "ID"])
        col_jpn10 = _find_column_index(headers, ["個体識別番号", "jpn10", "JPN10", "10桁"])
        col_date = _find_column_index(headers, ["分娩日", "日付", "event_date"])
        col_diff = _find_column_index(
            headers,
            ["分娩の難易度", "難易度", "分娩難易", "calving_difficulty"],
        )
        col_note = _find_column_index(headers, ["備考", "note", "メモ"])

        if col_date is None and len(headers) >= 2:
            col_date = 1
        if col_date is None:
            messagebox.showerror("エラー", "分娩日（event_date）列が見つかりません")
            return
        if col_id is None and col_jpn10 is None:
            col_id = 0
        if col_jpn10 is None:
            col_jpn10 = col_id

        default_year = datetime.now().year
        data_rows: List[Dict[str, Any]] = []
        for row in rows:
            if not any(row):
                continue
            if row and str(row[0]).strip().startswith("#"):
                continue
            jpn10_raw = row[col_jpn10] if col_jpn10 is not None and col_jpn10 < len(row) else ""
            jpn10 = _extract_jpn10(jpn10_raw)
            cow_id_raw = row[col_id] if col_id is not None and col_id < len(row) else ""
            cow_id = _to_cow_id(cow_id_raw, jpn10)
            if not cow_id and not jpn10:
                continue
            if not cow_id and jpn10:
                cow_id = jpn10[-4:] if len(jpn10) >= 4 else jpn10.zfill(4)
            date_val = row[col_date] if col_date < len(row) else ""
            event_date = _normalize_date_cell(date_val, default_year)
            if not event_date:
                continue

            diff_raw = ""
            if col_diff is not None and col_diff < len(row):
                diff_raw = str(row[col_diff]).strip()
            diff = _parse_calving_difficulty(diff_raw)

            note = ""
            if col_note is not None and col_note < len(row):
                note = str(row[col_note]).strip()

            calves = []
            for i in range(1, 5):
                c = self._collect_calf(row, i, headers)
                if c:
                    calves.append(c)

            data_rows.append(
                {
                    "cow_id": cow_id,
                    "jpn10": jpn10 or "",
                    "event_date": event_date,
                    "calving_difficulty": diff,
                    "note": note,
                    "calves": calves,
                }
            )
        self.data_rows = data_rows
        self._update_preview()
        self.info_label.config(text=f"データ件数: {len(self.data_rows)} 件")

    def _calf_summary(self, calves: List[Dict[str, Any]]) -> str:
        if not calves:
            return ""
        parts = []
        for c in calves:
            parts.append(f"{c.get('breed','')}/{c.get('sex','')}/{'死' if c.get('stillborn') else '生'}")
        return " | ".join(parts)

    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.data_rows:
            cow = self._resolve_cow(row.get("cow_id"), row.get("jpn10"))
            status = "FALCONにありません" if not cow else "取り込み対象"
            diff = row.get("calving_difficulty")
            diff_s = str(diff) if diff is not None else ""
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("cow_id", ""),
                    row.get("jpn10", ""),
                    row.get("event_date", ""),
                    diff_s,
                    self._calf_summary(row.get("calves") or []),
                    status,
                ),
            )

    def _resolve_cow(self, cow_id: str, jpn10: str):
        if jpn10 and len(re.sub(r"\D", "", jpn10)) == 10:
            try:
                cows = self.db.get_cows_by_jpn10(jpn10)
                if cows:
                    return cows[0]
            except Exception:
                pass
        if cow_id:
            return self.db.get_cow_by_id(cow_id)
        return None

    def _execute_import(self):
        if not self.data_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return
        if not messagebox.askyesno("確認", f"分娩イベントを {len(self.data_rows)} 件登録しますか？"):
            return

        ok = 0
        skip = 0
        errors: List[Dict[str, str]] = []
        for row in self.data_rows:
            cow = self._resolve_cow(row.get("cow_id", ""), row.get("jpn10", ""))
            if not cow:
                errors.append({"id": row.get("cow_id", ""), "reason": "個体なし"})
                continue
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID")
            if not cow_auto_id:
                errors.append({"id": row.get("cow_id", ""), "reason": "auto_idなし"})
                continue
            event_date = row.get("event_date", "")
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            dup = any(
                e.get("event_number") == RuleEngine.EVENT_CALV and e.get("event_date") == event_date
                for e in events
            )
            if dup:
                skip += 1
                continue

            jd: Dict[str, Any] = {}
            if row.get("calving_difficulty") is not None:
                jd["calving_difficulty"] = row["calving_difficulty"]
            if row.get("calves"):
                jd["calves"] = row["calves"]

            note = (row.get("note") or "").strip()
            try:
                eid = self.db.insert_event(
                    {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_CALV,
                        "event_date": event_date,
                        "json_data": jd if jd else None,
                        "note": note if note else "CSV一括取り込み（分娩）",
                    }
                )
                self.rule_engine.on_event_added(eid)
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                ok += 1
            except Exception as e:
                logger.error(e, exc_info=True)
                errors.append({"id": row.get("cow_id", ""), "reason": str(e)})

        msg = f"完了: 成功 {ok} 件 / スキップ {skip} 件 / エラー {len(errors)} 件"
        messagebox.showinfo("取り込み", msg)

    def show(self):
        self.window.focus_set()
