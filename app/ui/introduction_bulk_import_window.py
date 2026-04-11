"""
導入（外部）CSV/Excel 一括取り込み
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
from modules.introduction_registration import register_introduction_cow
from ui.ai_import_window import _find_column_index, _extract_jpn10, _to_cow_id

logger = logging.getLogger(__name__)


def _normalize_date(raw: Any, default_year: Optional[int] = None) -> Optional[str]:
    from ui.ai_import_window import _normalize_date_cell

    return _normalize_date_cell(raw, default_year)


def _map_pregnant(raw: str) -> str:
    """個体一括導入テンプレのプルダウン（受胎なし・妊娠鑑定待ち・妊娠・乾乳）と整合"""
    if not raw or not str(raw).strip():
        return ""
    s = str(raw).strip()
    m = {
        "妊娠鑑定待ち": "waiting",
        "受胎なし": "",
        "妊娠中": "pregnant",
        "妊娠": "pregnant",
        "乾乳中": "dry",
        "乾乳": "dry",
        "waiting": "waiting",
        "pregnant": "pregnant",
        "dry": "dry",
    }
    return m.get(s, s if s in ("", "waiting", "pregnant", "dry") else "")


class IntroductionBulkImportWindow:
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, rule_engine: RuleEngine, farm_path: Path):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("導入データ一括取り込み")
        self.window.geometry("1000x600")

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
        cols = ("jpn10", "intro", "reg", "breed", "status")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.heading("jpn10", text="JPN10")
        self.tree.heading("intro", text="導入日")
        self.tree.heading("reg", text="登録日")
        self.tree.heading("breed", text="品種")
        self.tree.heading("status", text="状態")
        for c in cols:
            self.tree.column(c, width=160)
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
        from modules.reproduction_flow_templates import write_introduction_template

        path = filedialog.asksaveasfilename(
            parent=self.window,
            title="導入テンプレートの保存先",
            defaultextension=".xlsx",
            filetypes=[("Excel（推奨）", "*.xlsx"), ("CSV", "*.csv"), ("すべて", "*.*")],
            initialfile="繁殖検診_導入テンプレート.xlsx",
        )
        if not path:
            return
        try:
            write_introduction_template(Path(path))
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

    def _get(self, row: List[str], col: Optional[int]) -> str:
        if col is None or col >= len(row):
            return ""
        return str(row[col]).strip()

    def _build_rows(self, headers: List[str], rows: List[List[str]]):
        c_intro = _find_column_index(headers, ["導入日", "introduction_date"])
        c_reg = _find_column_index(headers, ["登録日", "registration_date"])
        c_jpn = _find_column_index(headers, ["個体識別番号", "jpn10", "JPN10"])
        c_id = _find_column_index(headers, ["牛のID", "cow_id", "個体ID", "ID"])
        c_breed = _find_column_index(headers, ["品種", "breed"])
        c_birth = _find_column_index(headers, ["生年月日", "birth_date"])
        c_lact = _find_column_index(headers, ["産次", "lact"])
        c_clvd = _find_column_index(headers, ["最終分娩日", "clvd", "分娩月日"])
        c_lai = _find_column_index(headers, ["最終授精日", "last_ai_date"])
        c_las = _find_column_index(
            headers,
            ["最終授精SIRE", "最終授精の種雄牛", "last_ai_sire", "SIRE"],
        )
        c_ac = _find_column_index(headers, ["授精回数", "ai_count"])
        c_ps = _find_column_index(
            headers,
            ["妊娠状態", "繁殖の状態", "pregnant_status", "受胎"],
        )
        c_pen = _find_column_index(headers, ["群(PEN)", "群", "pen", "PEN"])
        c_tag = _find_column_index(headers, ["タグ", "tag", "TAG"])
        c_memo = _find_column_index(headers, ["メモ", "note"])

        if c_intro is None:
            messagebox.showerror("エラー", "「導入日」の列が見つかりません。1行目の列名を変更していないか確認してください。")
            return
        if c_jpn is None:
            messagebox.showerror(
                "エラー",
                "「個体識別番号(JPN10)」の列が見つかりません。1行目の列名を変更していないか確認してください。",
            )
            return

        default_year = datetime.now().year
        out: List[Dict[str, Any]] = []
        for row in rows:
            if not any(row):
                continue
            if row and str(row[0]).strip().startswith("#"):
                continue
            intro = _normalize_date(self._get(row, c_intro), default_year)
            if not intro:
                continue

            jpn10 = _extract_jpn10(self._get(row, c_jpn))
            if not jpn10 or len(re.sub(r"\D", "", jpn10)) != 10:
                continue

            reg_raw = self._get(row, c_reg)
            registration_date = _normalize_date(reg_raw, default_year) if reg_raw else intro

            cow_id = self._get(row, c_id)
            if not cow_id:
                cow_id = jpn10[5:9].zfill(4)
            else:
                cow_id = _to_cow_id(cow_id, jpn10)

            lact_s = self._get(row, c_lact)
            lact = int(lact_s) if lact_s.isdigit() else 0

            clvd = _normalize_date(self._get(row, c_clvd), default_year) if c_clvd is not None else None
            if self._get(row, c_clvd) and not clvd:
                clvd = None

            last_ai = _normalize_date(self._get(row, c_lai), default_year) if c_lai is not None else None
            if c_lai is not None and self._get(row, c_lai) and not last_ai:
                last_ai = None

            ai_count_s = self._get(row, c_ac)
            ai_count = int(ai_count_s) if ai_count_s.isdigit() else 0

            ps = _map_pregnant(self._get(row, c_ps))
            last_ai_sire = self._get(row, c_las) if c_las is not None else ""

            rc = RuleEngine.RC_OPEN
            if ps == "pregnant":
                rc = RuleEngine.RC_PREGNANT
            elif ps == "dry":
                rc = RuleEngine.RC_DRY
            elif ps == "waiting" or (ps == "" and last_ai):
                rc = RuleEngine.RC_BRED

            birth = _normalize_date(self._get(row, c_birth), default_year) if c_birth is not None else None
            if c_birth is not None and self._get(row, c_birth) and not birth:
                birth = None

            pen_raw = self._get(row, c_pen)
            pen_val = None
            if pen_raw:
                pen_val = pen_raw.split(":")[0].strip() if ":" in pen_raw else pen_raw

            tag_raw = self._get(row, c_tag) if c_tag is not None else ""
            if not str(tag_raw).strip() and c_memo is not None:
                tag_raw = self._get(row, c_memo)

            out.append(
                {
                    "introduction_date": intro,
                    "registration_date": registration_date or intro,
                    "cow_id": cow_id,
                    "jpn10": jpn10,
                    "breed": self._get(row, c_breed) or None,
                    "birth_date": birth,
                    "lact": lact,
                    "clvd": clvd,
                    "last_ai_date": last_ai,
                    "last_ai_sire": last_ai_sire or None,
                    "ai_count": ai_count,
                    "pregnant_status": ps,
                    "rc": rc,
                    "pen": pen_val,
                    "tag": (str(tag_raw).strip() or None),
                }
            )

        self.data_rows = out
        self._update_preview()
        self.info_label.config(text=f"データ件数: {len(self.data_rows)} 件")

    def _update_preview(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in self.data_rows:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    r.get("jpn10", ""),
                    r.get("introduction_date", ""),
                    r.get("registration_date", ""),
                    r.get("breed") or "",
                    "取り込み対象",
                ),
            )

    def _execute_import(self):
        if not self.data_rows:
            messagebox.showwarning("警告", "データがありません")
            return
        if not messagebox.askyesno("確認", f"{len(self.data_rows)} 頭を導入登録しますか？"):
            return

        ok = 0
        err = 0
        for r in self.data_rows:
            try:
                register_introduction_cow(
                    self.db,
                    self.rule_engine,
                    {
                        "cow_id": r["cow_id"],
                        "jpn10": r["jpn10"],
                        "breed": r.get("breed"),
                        "birth_date": r.get("birth_date"),
                        "lact": r.get("lact", 0),
                        "clvd": r.get("clvd"),
                        "last_ai_date": r.get("last_ai_date"),
                        "last_ai_sire": r.get("last_ai_sire"),
                        "ai_count": r.get("ai_count", 0),
                        "pregnant_status": r.get("pregnant_status", ""),
                        "rc": r.get("rc"),
                        "pen": r.get("pen"),
                        "tag": r.get("tag"),
                    },
                    r["introduction_date"],
                    r["registration_date"],
                    intro_note="CSV一括導入",
                )
                ok += 1
            except Exception as e:
                logger.error(e, exc_info=True)
                err += 1

        messagebox.showinfo("完了", f"成功: {ok} 頭 / エラー: {err} 頭")

    def show(self):
        self.window.focus_set()
