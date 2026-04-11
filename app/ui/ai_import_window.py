"""
FALCON2 - AIデータ取り込みウィンドウ
Excel/CSVから個体ID（JPN10またはID）・日付・SIRE・授精師を読み取り、AI/ETイベントとして登録する。
授精種類は直近の繁殖処置と「〇日後AI」設定から自動振り分け。
FALCONに存在しない個体は一覧表示し、テキストファイルに出力可能（乳検吸い込みと同仕様）。
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List, Tuple
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.insemination_type_resolver import compute_suggested_insemination_type_code

logger = logging.getLogger(__name__)


def _find_column_index(headers: List[str], keywords: List[str]) -> Optional[int]:
    """ヘッダー行からキーワードに部分一致する列のインデックスを返す（大文字小文字は区別しない）"""
    for kw in keywords:
        kw_norm = kw.strip()
        kw_lower = kw_norm.lower()
        for idx, h in enumerate(headers):
            if h is None:
                continue
            h_str = str(h).strip()
            if not h_str:
                continue
            if kw_norm in h_str:
                return idx
            if kw_norm.isascii() and kw_lower in h_str.lower():
                return idx
    return None


def _normalize_date_cell(raw: Any, default_year: Optional[int] = None) -> Optional[str]:
    """
    セル値を YYYY-MM-DD に正規化する。
    対応例: 2026-01-07, 2026/1/7, 1月7日, 1/7, 20260107, datetime, Excelシリアル値
    """
    if raw is None:
        return None
    # datetime / date オブジェクト（Excel が日付として読んだ場合など）
    if hasattr(raw, "strftime"):
        try:
            return raw.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            pass
    # Excel の日付シリアル値（数値）
    if isinstance(raw, (int, float)):
        try:
            if 20000 <= raw <= 50000:  # 1970年代〜2030年代程度
                d = datetime(1899, 12, 30) + timedelta(days=int(raw))
                return d.strftime("%Y-%m-%d")
        except (ValueError, TypeError, OSError):
            pass
    s = str(raw).strip()
    if not s:
        return None
    # 数値文字列（Excel シリアルが str になった場合）
    try:
        n = float(s)
        if 20000 <= n <= 50000:
            d = datetime(1899, 12, 30) + timedelta(days=int(n))
            return d.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    year = default_year or datetime.now().year
    # YYYY-MM-DD / YYYY/MM/DD
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$", s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except (ValueError, TypeError):
            pass
    # YYYYMMDD
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except (ValueError, TypeError):
            pass
    # 1月7日 / 1/7
    m = re.match(r"^(\d{1,2})月(\d{1,2})日$", s)
    if m:
        try:
            mo, d = int(m.group(1)), int(m.group(2))
            return f"{year:04d}-{mo:02d}-{d:02d}"
        except (ValueError, TypeError):
            pass
    m = re.match(r"^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$", s)
    if m:
        try:
            mo = int(m.group(1))
            d = int(m.group(2))
            y = int(m.group(3)) if m.group(3) else year
            if y < 100:
                y += 2000 if y < 50 else 1900
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except (ValueError, TypeError):
            pass
    # datetime の isoformat など
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(s[:10], fmt[:10])
            return dt.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
    return None


def _extract_jpn10(raw: Any) -> Optional[str]:
    """10桁個体識別番号を抽出"""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    if len(digits) == 10:
        return digits
    if len(digits) > 10:
        return digits[-10:]
    return None


def _to_cow_id(raw: Any, jpn10_fallback: Optional[str] = None) -> str:
    """4桁の cow_id を取得。raw が4桁以上なら下4桁、なければ JPN10 の下4桁"""
    if raw is not None and str(raw).strip():
        s = re.sub(r"\D", "", str(raw).strip())
        if len(s) >= 4:
            return s[-4:]
        if len(s) > 0:
            return s.zfill(4)
    if jpn10_fallback and len(jpn10_fallback) >= 4:
        return jpn10_fallback[-4:]
    return ""


class AIImportWindow:
    """AIデータ取り込みウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        db_handler: DBHandler,
        rule_engine: RuleEngine,
        farm_path: Path,
    ):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("AIデータ取り込み")
        self.window.geometry("1020x620")

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

        self.tree = ttk.Treeview(
            tree_frame,
            columns=(
                "cow_id",
                "jpn10",
                "date",
                "sire",
                "technician",
                "kind",
                "insem",
                "note",
                "status",
            ),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        self.tree.heading("cow_id", text="ID")
        self.tree.heading("jpn10", text="個体識別番号")
        self.tree.heading("date", text="日付")
        self.tree.heading("sire", text="SIRE")
        self.tree.heading("technician", text="授精師")
        self.tree.heading("kind", text="AI/ET")
        self.tree.heading("insem", text="授精種類")
        self.tree.heading("note", text="備考")
        self.tree.heading("status", text="状態")
        for col in ("cow_id", "jpn10", "date", "sire", "technician", "kind", "insem", "note", "status"):
            self.tree.column(col, width=82)
        self.tree.column("note", width=100)
        self.tree.column("status", width=120)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)

        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        self.info_label = ttk.Label(info_frame, text="Excel または CSV ファイルを選択してください")
        self.info_label.pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="取り込み実行", command=self._execute_import).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="AIデータファイルを選択",
            filetypes=[
                ("Excel / CSV", "*.xlsx;*.xls;*.csv"),
                ("Excel", "*.xlsx;*.xls"),
                ("CSV", "*.csv"),
                ("すべて", "*.*"),
            ],
        )
        if path:
            self.file_path = Path(path)
            self.file_path_var.set(str(self.file_path))
            self._parse_file()

    def _parse_file(self):
        if not self.file_path or not self.file_path.exists():
            messagebox.showerror("エラー", "ファイルが見つかりません")
            return
        suffix = self.file_path.suffix.lower()
        try:
            if suffix == ".csv":
                self._parse_csv()
            elif suffix in (".xlsx", ".xls"):
                self._parse_excel()
            else:
                messagebox.showerror("エラー", "対応形式は CSV / .xlsx / .xls です")
                return
        except Exception as e:
            logger.error(f"AI取込パースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")

    def _read_csv_to_lines(self) -> List[List[str]]:
        for enc in ("utf-8", "utf-8-sig", "shift_jis", "cp932"):
            try:
                with open(self.file_path, "r", encoding=enc, newline="") as f:
                    lines = [row for row in __import__("csv").reader(f)]
                return lines
            except UnicodeDecodeError:
                continue
        raise ValueError("文字コードを判別できませんでした")

    def _export_template(self):
        from modules.reproduction_flow_templates import write_ai_et_template

        path = filedialog.asksaveasfilename(
            parent=self.window,
            title="AI/ET テンプレートの保存先",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel（推奨）", "*.xlsx"),
                ("CSV", "*.csv"),
                ("すべて", "*.*"),
            ],
            initialfile="繁殖検診_AI_ETテンプレート.xlsx",
        )
        if not path:
            return
        try:
            write_ai_et_template(Path(path))
            messagebox.showinfo("完了", f"テンプレートを保存しました:\n{path}")
        except Exception as e:
            logger.error(f"テンプレート出力エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", str(e))

    def _parse_csv(self):
        lines = self._read_csv_to_lines()
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
        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("エラー", "Excel を読むには pandas が必要です")
            return
        try:
            df = pd.read_excel(self.file_path, header=0, sheet_name=0)
        except Exception as e:
            try:
                df = pd.read_excel(self.file_path, header=0, sheet_name=0, engine="openpyxl")
            except Exception as e2:
                messagebox.showerror("エラー", f"Excel の読み込みに失敗しました:\n{e2}")
                return
        if df is None or df.empty:
            self.data_rows = []
            self._update_preview()
            self.info_label.config(text="データ件数: 0 件（シートにデータがありません）")
            return
        headers = [str(h).strip() if h is not None and (not hasattr(h, "__iter__") or isinstance(h, str)) else "" for h in df.columns.tolist()]
        # セル値を文字列に統一（日付は YYYY-MM-DD に変換してから渡す）
        def _cell_to_str(v):
            if v is None or (hasattr(pd, "isna") and pd.isna(v)):
                return ""
            if hasattr(v, "strftime"):
                try:
                    return v.strftime("%Y-%m-%d")
                except (ValueError, TypeError):
                    pass
            if isinstance(v, (int, float)) and 20000 <= v <= 50000:
                try:
                    d = datetime(1899, 12, 30) + timedelta(days=int(v))
                    return d.strftime("%Y-%m-%d")
                except (ValueError, TypeError, OSError):
                    pass
            return str(v).strip()
        rows = []
        for _, r in df.iterrows():
            line = [_cell_to_str(v) for v in r.tolist()]
            if line and str(line[0]).strip().startswith("#"):
                continue
            rows.append(line)
        self._build_data_rows(headers, rows)

    def _build_data_rows(self, headers: List[str], rows: List[List[Any]]):
        col_id = _find_column_index(
            headers,
            ["cow_id", "牛のID", "個体ID", "動物ID", "ID"],
        )
        col_jpn10 = _find_column_index(
            headers,
            ["個体識別番号", "JPN10", "Official ID", "個体識別", "識別番号", "jpn10"],
        )
        col_date = _find_column_index(
            headers,
            ["授精日", "日付", "AI日", "event_date", "date", "DATE"],
        )
        col_sire = _find_column_index(headers, ["種雄牛", "SIRE", "雄牛", " sire"])
        col_technician = _find_column_index(
            headers,
            ["授精師コード", "授精師", "technician", "技師"],
        )
        col_et = _find_column_index(headers, ["胚移植", "ET"])
        col_insem = _find_column_index(
            headers,
            ["授精種類コード", "insemination_type_code", "授精種類"],
        )
        col_note = _find_column_index(headers, ["備考", "note", "メモ"])

        # 列が見つからない場合は先頭列をID・2列目を日付としてフォールバック（4桁IDのみのExcel対応）
        if col_date is None and len(headers) >= 2:
            col_date = 1
        if col_date is None:
            messagebox.showerror(
                "エラー",
                "「授精日」または「日付」に相当する列が見つかりません。\n"
                "1行目の列名を変えていないか確認してください。",
            )
            return
        if col_id is None and col_jpn10 is None:
            if len(headers) >= 1:
                col_id = 0  # 先頭列をIDとして使用（4桁ID照合用）
                col_jpn10 = None
            else:
                messagebox.showerror("エラー", "「ID」または「JPN10」（個体識別番号）列が見つかりません")
                return

        # JPN10 列がない場合は ID 列で照合（4桁のIDでFALCONと整合）
        if col_jpn10 is None:
            col_jpn10 = col_id

        default_year = datetime.now().year
        data_rows = []
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
            sire = ""
            if col_sire is not None and col_sire < len(row):
                sire = str(row[col_sire]).strip()
            technician = ""
            if col_technician is not None and col_technician < len(row):
                technician = str(row[col_technician]).strip()
            is_et = False
            if col_et is not None and col_et < len(row):
                et_val = str(row[col_et]).strip()
                is_et = bool(et_val and et_val.lower() not in ("0", "否", "なし", "-", ""))

            insem = ""
            if col_insem is not None and col_insem < len(row):
                insem = str(row[col_insem]).strip()
            note = ""
            if col_note is not None and col_note < len(row):
                note = str(row[col_note]).strip()

            data_rows.append({
                "cow_id": cow_id,
                "jpn10": jpn10 or "",
                "event_date": event_date,
                "sire": sire,
                "technician": technician,
                "is_et": is_et,
                "insemination_type_code": insem,
                "note": note,
            })
        self.data_rows = data_rows
        self._update_preview()
        self.info_label.config(text=f"データ件数: {len(self.data_rows)} 件")

    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.data_rows:
            cow = self._resolve_cow(row.get("cow_id"), row.get("jpn10"))
            status = "FALCONにありません" if not cow else "吸い込み対象"
            kind = "ET" if row.get("is_et") else "AI"
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("cow_id", ""),
                    row.get("jpn10", ""),
                    row.get("event_date", ""),
                    row.get("sire", ""),
                    row.get("technician", ""),
                    kind,
                    row.get("insemination_type_code", ""),
                    row.get("note", ""),
                    status,
                ),
            )

    def _resolve_cow(self, cow_id: str, jpn10: str):
        """個体を検索。JPN10 があれば JPN10 で、なければ 4桁ID（cow_id）で照合する。"""
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
        ok = messagebox.askyesno(
            "確認",
            f"AI/ET データを {len(self.data_rows)} 件取り込みますか？\n"
            "同一個体・同日の既存AI/ETはスキップし、授精種類は繁殖処置の「〇日後AI」から自動判定します。",
        )
        if not ok:
            return

        success_count = 0
        skip_count = 0
        error_cows: List[Dict[str, Any]] = []

        for row in self.data_rows:
            cow = self._resolve_cow(row.get("cow_id", ""), row.get("jpn10", ""))
            if not cow:
                error_cows.append({
                    "cow_id": row.get("cow_id", ""),
                    "jpn10": row.get("jpn10", ""),
                    "reason": "FALCONに個体がありません",
                })
                continue

            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID")
            if not cow_auto_id:
                error_cows.append({
                    "cow_id": row.get("cow_id", ""),
                    "jpn10": row.get("jpn10", ""),
                    "reason": "個体の auto_id が取得できません",
                })
                continue

            event_date = row.get("event_date", "")
            if not event_date:
                continue

            # 同一個体・同日の AI/ET があればスキップ
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            already = False
            for ev in events:
                if ev.get("event_number") in (RuleEngine.EVENT_AI, RuleEngine.EVENT_ET):
                    if ev.get("event_date") == event_date:
                        already = True
                        break
            if already:
                skip_count += 1
                continue

            event_number = RuleEngine.EVENT_ET if row.get("is_et") else RuleEngine.EVENT_AI
            json_data = {}
            if row.get("sire"):
                json_data["sire"] = str(row["sire"]).strip()
            if row.get("technician"):
                json_data["technician_code"] = str(row["technician"]).strip()

            manual_insem = (row.get("insemination_type_code") or "").strip()
            if manual_insem:
                code_part = manual_insem
                if "：" in code_part:
                    code_part = code_part.split("：", 1)[0].strip()
                elif ":" in code_part:
                    code_part = code_part.split(":", 1)[0].strip()
                json_data["insemination_type_code"] = code_part
            else:
                type_code = compute_suggested_insemination_type_code(
                    self.db, self.farm_path, cow_auto_id, event_date
                )
                if type_code is not None:
                    json_data["insemination_type_code"] = type_code

            note_text = (row.get("note") or "").strip()
            try:
                event_data = {
                    "cow_auto_id": cow_auto_id,
                    "event_number": event_number,
                    "event_date": event_date,
                    "json_data": json_data if json_data else None,
                    "note": note_text if note_text else "Excel/CSVから取り込み",
                }
                event_id = self.db.insert_event(event_data)
                self.rule_engine.on_event_added(event_id)
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                success_count += 1
            except Exception as e:
                logger.error(f"AIイベント作成エラー: {e}", exc_info=True)
                error_cows.append({
                    "cow_id": row.get("cow_id", ""),
                    "jpn10": row.get("jpn10", ""),
                    "reason": f"イベント作成エラー: {str(e)}",
                })

        msg = f"取り込みが完了しました\n成功: {success_count}件\nスキップ: {skip_count}件\nエラー: {len(error_cows)}件"
        if error_cows:
            msg += "\n\nFALCONに存在しない個体（またはエラー）が一覧にあります。詳細を表示しますか？"
            if messagebox.askyesno("完了", msg):
                self._show_error_cows_window(error_cows)
        else:
            messagebox.showinfo("完了", msg)
        self._update_preview()

    def _show_error_cows_window(self, error_cows: List[Dict[str, Any]]):
        err_win = tk.Toplevel(self.window)
        err_win.title("取り込みできなかった個体一覧")
        err_win.geometry("650x420")

        main_frame = ttk.Frame(err_win, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            main_frame,
            text=f"ファイルに含まれるが FALCON に存在しない個体など（{len(error_cows)}件）",
            font=("", 10, "bold"),
        ).pack(pady=(0, 10))

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree = ttk.Treeview(
            tree_frame,
            columns=("cow_id", "jpn10", "reason"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
        )
        tree.heading("cow_id", text="ID")
        tree.heading("jpn10", text="個体識別番号")
        tree.heading("reason", text="理由")
        tree.column("cow_id", width=100)
        tree.column("jpn10", width=150)
        tree.column("reason", width=280)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=tree.yview)
        for ec in error_cows:
            tree.insert("", tk.END, values=(ec.get("cow_id", ""), ec.get("jpn10", ""), ec.get("reason", "")))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            btn_frame,
            text="テキストファイルとして保存",
            command=lambda: self._save_error_cows_to_file(error_cows, err_win),
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="閉じる", command=err_win.destroy).pack(side=tk.RIGHT)

    def _save_error_cows_to_file(self, error_cows: List[Dict[str, Any]], parent: tk.Toplevel):
        path = filedialog.asksaveasfilename(
            title="取り込みできなかった個体一覧を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("=" * 60 + "\n")
                f.write("AI吸い込み エラー個体一覧（FALCONに存在しない個体など）\n")
                f.write("=" * 60 + "\n")
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"エラー件数: {len(error_cows)}件\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"{'ID':<10} {'個体識別番号':<15} {'理由':<30}\n")
                f.write("-" * 60 + "\n")
                for ec in error_cows:
                    f.write(
                        f"{ec.get('cow_id', ''):<10} {ec.get('jpn10', ''):<15} {ec.get('reason', ''):<30}\n"
                    )
            messagebox.showinfo("保存完了", f"エラー個体一覧を保存しました:\n{path}")
        except Exception as e:
            logger.error(f"保存エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ファイルの保存に失敗しました:\n{e}")

    def show(self):
        self.window.focus_set()
