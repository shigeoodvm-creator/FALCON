"""
売却・死亡廃用 CSV/Excel 取り込み
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
from ui.ai_import_window import _find_column_index, _normalize_date_cell, _extract_jpn10, _to_cow_id

logger = logging.getLogger(__name__)


def _parse_exit_kind(raw: str) -> Optional[int]:
    if not raw or not str(raw).strip():
        return None
    s = str(raw).strip().upper()
    if s in ("205", "SOLD", "売却"):
        return RuleEngine.EVENT_SOLD
    if s in ("206", "DEAD", "死亡", "死廃", "廃用", "死亡・廃用"):
        return RuleEngine.EVENT_DEAD
    s_lower = str(raw).strip()
    if "売" in s_lower or "却" in s_lower:
        return RuleEngine.EVENT_SOLD
    if "死" in s_lower or "廃" in s_lower:
        return RuleEngine.EVENT_DEAD
    return None


class ExitImportWindow:
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine, farm_path: Path):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("退出データ取り込み")
        self.window.geometry("900x580")

        self.file_path: Optional[Path] = None
        self.data_rows: List[Dict[str, Any]] = []

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ff = ttk.LabelFrame(main_frame, text="ファイル選択（Excel / CSV）", padding=10)
        ff.pack(fill=tk.X, pady=(0, 10))
        self.file_path_var = tk.StringVar()
        ttk.Entry(ff, textvariable=self.file_path_var, width=55, state="readonly").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(ff, text="参照...", command=self._select_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(ff, text="テンプレート出力", command=self._export_template).pack(side=tk.LEFT)

        pf = ttk.LabelFrame(main_frame, text="プレビュー", padding=10)
        pf.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        tf = ttk.Frame(pf)
        tf.pack(fill=tk.BOTH, expand=True)
        sy = ttk.Scrollbar(tf, orient=tk.VERTICAL)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx = ttk.Scrollbar(tf, orient=tk.HORIZONTAL)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        cols = ("cow_id", "jpn10", "date", "kind", "note", "status")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", yscrollcommand=sy.set, xscrollcommand=sx.set)
        for c, t in zip(
            cols,
            ["ID", "個体識別番号", "日付", "種別", "備考", "状態"],
        ):
            self.tree.heading(c, text=t)
        for c in cols:
            self.tree.column(c, width=110)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.config(command=self.tree.yview)
        sx.config(command=self.tree.xview)

        self.info_label = ttk.Label(main_frame, text="ファイルを選択してください")
        self.info_label.pack(anchor=tk.W)

        bf = ttk.Frame(main_frame)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="取り込み実行", command=self._execute_import).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(bf, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    def _export_template(self):
        from modules.reproduction_flow_templates import write_exit_template

        path = filedialog.asksaveasfilename(
            parent=self.window,
            title="退出テンプレートの保存先",
            defaultextension=".xlsx",
            filetypes=[("Excel（推奨）", "*.xlsx"), ("CSV", "*.csv"), ("すべて", "*.*")],
            initialfile="繁殖検診_退出テンプレート.xlsx",
        )
        if not path:
            return
        try:
            write_exit_template(Path(path))
            messagebox.showinfo("完了", f"テンプレートを保存しました:\n{path}")
        except Exception as e:
            logger.error(e, exc_info=True)
            messagebox.showerror("エラー", str(e))

    def _select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel / CSV", "*.xlsx;*.xls;*.csv"), ("すべて", "*.*")],
        )
        if path:
            self.file_path = Path(path)
            self.file_path_var.set(str(self.file_path))
            self._parse_file()

    def _parse_file(self):
        if not self.file_path or not self.file_path.exists():
            return
        try:
            if self.file_path.suffix.lower() == ".csv":
                self._parse_csv()
            else:
                self._parse_excel()
        except Exception as e:
            logger.error(e, exc_info=True)
            messagebox.showerror("エラー", str(e))

    def _read_csv(self) -> List[List[str]]:
        import csv as csv_mod

        for enc in ("utf-8", "utf-8-sig", "shift_jis", "cp932"):
            try:
                with open(self.file_path, "r", encoding=enc, newline="") as f:
                    return [r for r in csv_mod.reader(f)]
            except UnicodeDecodeError:
                continue
        raise ValueError("文字コードを判別できませんでした")

    def _parse_csv(self):
        lines = self._read_csv()
        if not lines:
            return
        max_cols = max(len(r) for r in lines)
        for r in lines:
            while len(r) < max_cols:
                r.append("")
        headers = [str(h).strip() for h in lines[0]]
        self._build_rows(headers, lines[1:])

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
        self._build_rows(headers, rows)

    def _build_rows(self, headers: List[str], rows: List[List[str]]):
        col_id = _find_column_index(headers, ["cow_id", "牛のID", "個体ID", "ID"])
        col_j = _find_column_index(headers, ["個体識別番号", "jpn10", "JPN10", "10桁"])
        col_d = _find_column_index(headers, ["退出日", "日付", "event_date"])
        col_k = _find_column_index(headers, ["退出の種別", "exit_kind", "種別", "区分"])
        col_n = _find_column_index(headers, ["備考", "note", "メモ"])

        if col_d is None:
            messagebox.showerror(
                "エラー",
                "「退出日」または「日付」に相当する列が見つかりません。\n1行目の列名を変更していないか確認してください。",
            )
            return
        if col_id is None and col_j is None:
            col_id = 0
        if col_j is None:
            col_j = col_id

        default_year = datetime.now().year
        out: List[Dict[str, Any]] = []
        for row in rows:
            if not any(row):
                continue
            if row and str(row[0]).strip().startswith("#"):
                continue
            jpn10 = _extract_jpn10(row[col_j] if col_j < len(row) else "")
            cow_id = _to_cow_id(row[col_id] if col_id is not None and col_id < len(row) else "", jpn10)
            if not cow_id and jpn10:
                cow_id = jpn10[-4:] if len(jpn10) >= 4 else jpn10.zfill(4)
            if not cow_id and not jpn10:
                continue

            ed = _normalize_date_cell(row[col_d] if col_d < len(row) else "", default_year)
            if not ed:
                continue

            kind_raw = row[col_k] if col_k is not None and col_k < len(row) else ""
            evn = _parse_exit_kind(str(kind_raw))
            if evn is None:
                continue

            note = ""
            if col_n is not None and col_n < len(row):
                note = str(row[col_n]).strip()

            out.append(
                {
                    "cow_id": cow_id,
                    "jpn10": jpn10 or "",
                    "event_date": ed,
                    "event_number": evn,
                    "note": note,
                }
            )

        self.data_rows = out
        self._update_preview()
        self.info_label.config(text=f"データ件数: {len(self.data_rows)} 件")

    def _kind_label(self, n: int) -> str:
        if n == RuleEngine.EVENT_SOLD:
            return "売却"
        if n == RuleEngine.EVENT_DEAD:
            return "死亡廃用"
        return str(n)

    def _update_preview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in self.data_rows:
            cow = self._resolve_cow(r.get("cow_id", ""), r.get("jpn10", ""))
            st = "FALCONにありません" if not cow else "取り込み対象"
            self.tree.insert(
                "",
                tk.END,
                values=(
                    r.get("cow_id", ""),
                    r.get("jpn10", ""),
                    r.get("event_date", ""),
                    self._kind_label(r.get("event_number")),
                    r.get("note", ""),
                    st,
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
            messagebox.showwarning("警告", "データがありません")
            return
        if not messagebox.askyesno("確認", f"{len(self.data_rows)} 件の退出イベントを登録しますか？"):
            return

        ok = 0
        skip = 0
        err = 0
        for r in self.data_rows:
            cow = self._resolve_cow(r.get("cow_id", ""), r.get("jpn10", ""))
            if not cow:
                err += 1
                continue
            aid = cow.get("auto_id") or cow.get("auto_ID")
            if not aid:
                err += 1
                continue
            en = r["event_number"]
            dt = r["event_date"]
            evs = self.db.get_events_by_cow(aid, include_deleted=False)
            if any(e.get("event_number") == en and e.get("event_date") == dt for e in evs):
                skip += 1
                continue
            note = (r.get("note") or "").strip()
            try:
                eid = self.db.insert_event(
                    {
                        "cow_auto_id": aid,
                        "event_number": en,
                        "event_date": dt,
                        "json_data": None,
                        "note": note if note else "CSV一括取り込み（退出）",
                    }
                )
                self.rule_engine.on_event_added(eid)
                self.rule_engine._recalculate_and_update_cow(aid)
                ok += 1
            except Exception as e:
                logger.error(e, exc_info=True)
                err += 1

        messagebox.showinfo("完了", f"成功: {ok} / スキップ: {skip} / エラー: {err}")

    def show(self):
        self.window.focus_set()
