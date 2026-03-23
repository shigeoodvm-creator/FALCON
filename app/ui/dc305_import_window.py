"""
FALCON2 - DC305形式Excel イベント取り込みウィンドウ
列: A=ID, B=イベント名, C=DIM, D=年月日, E=Remark, F=R欄(O,P,R), G=無視, H=授精種類, I,J=無視
イベント: SOLD→売却, FRESH→分娩, OK→フレッシュ/繁殖検査, BRED/OPEN/PREG→AI(+妊鑑), DNB→繁殖停止, DIED→死亡
吸い込み日を note に記録し、一括削除可能。
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List, Tuple
import json
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)

DC305_NOTE_PREFIX = "DC305取込:"

# Excel列（0-based）
COL_ID = 0
COL_EVENT_NAME = 1
COL_DIM = 2
COL_DATE = 3
COL_REMARK = 4
COL_R_OPR = 5
COL_IGNORE_G = 6
COL_INSEM_TYPE = 7
COL_IGNORE_I = 8
COL_IGNORE_J = 9

EVENT_NAME_MAP = {
    "SOLD": "SOLD",
    "FRESH": "FRESH",
    "OK": "OK",
    "BRED": "BRED",
    "OPEN": "OPEN",
    "PREG": "PREG",
    "DNB": "DNB",
    "DIED": "DIED",
}


def _normalize_date_cell(raw: Any, default_year: Optional[int] = None) -> Optional[str]:
    """セル値を YYYY-MM-DD に正規化"""
    if raw is None:
        return None
    if hasattr(raw, "strftime"):
        try:
            return raw.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            pass
    if isinstance(raw, (int, float)):
        try:
            if 20000 <= raw <= 50000:
                d = datetime(1899, 12, 30) + timedelta(days=int(raw))
                return d.strftime("%Y-%m-%d")
        except (ValueError, TypeError, OSError):
            pass
    s = str(raw).strip()
    if not s:
        return None
    try:
        n = float(s)
        if 20000 <= n <= 50000:
            d = datetime(1899, 12, 30) + timedelta(days=int(n))
            return d.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    year = default_year or datetime.now().year
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(s[:10], fmt[:10])
            return dt.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            continue
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$", s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except (ValueError, TypeError):
            pass
    return None


def _to_cow_id(raw: Any) -> str:
    """4桁の cow_id を取得"""
    if raw is None or not str(raw).strip():
        return ""
    s = re.sub(r"\D", "", str(raw).strip())
    if len(s) >= 4:
        return s[-4:]
    return s.zfill(4) if s else ""


def _parse_r_opr(raw: Any) -> Tuple[str, str, str]:
    """F列（R欄）を 右・左・その他 に分割。カンマ/タブで分割し先頭3要素を右・左・その他とする"""
    if raw is None:
        return "", "", ""
    s = str(raw).strip()
    if not s:
        return "", "", ""
    parts = re.split(r"[, \t]+", s, maxsplit=2)
    right = parts[0] if len(parts) > 0 else ""
    left = parts[1] if len(parts) > 1 else ""
    other = parts[2] if len(parts) > 2 else ""
    return right, left, other


def _classify_ok_events(rows: List[Dict[str, Any]]) -> None:
    """
    同一牛・日付順で並べ、OK を「分娩後の初回→フレッシュ(300)、それ以外→繁殖検査(301)」に振り分け。
    rows は cow_id, event_name, event_date を持ち、event_name が 'OK' のものに
    falcon_event_number を 300 または 301 でセットする。他イベントは呼び出し元でセット済み想定。
    """
    # (cow_id, event_date) でソート
    sorted_rows = sorted(rows, key=lambda r: (r.get("cow_id", ""), r.get("event_date", "")))
    per_cow_last_fresh: Dict[str, str] = {}
    per_cow_seen_ok_since_fresh: Dict[str, bool] = {}

    for row in sorted_rows:
        cow_id = row.get("cow_id", "")
        event_name = row.get("event_name", "").strip().upper()
        event_date = row.get("event_date", "")

        if event_name == "FRESH":
            per_cow_last_fresh[cow_id] = event_date
            per_cow_seen_ok_since_fresh[cow_id] = False
            continue
        if event_name != "OK":
            continue

        first_ok_after_fresh = not per_cow_seen_ok_since_fresh.get(cow_id, True)
        row["falcon_event_number"] = RuleEngine.EVENT_FCHK if first_ok_after_fresh else RuleEngine.EVENT_REPRO
        per_cow_seen_ok_since_fresh[cow_id] = True


class DC305ImportWindow:
    """DC305形式Excel イベント取り込みウィンドウ"""

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
        self.insemination_types: Dict[str, str] = {}
        self._load_insemination_types()

        self.window = tk.Toplevel(parent)
        self.window.title("DC305からイベント吸い込み")
        self.window.geometry("920x620")

        self.file_path: Optional[Path] = None
        self.parsed_rows: List[Dict[str, Any]] = []
        self.insem_type_mapping: Dict[str, str] = {}  # ExcelのH列値 -> FALCON授精種類コード

        self._create_widgets()

    def _load_insemination_types(self):
        """insemination_settings.json から授精種類を読み込む"""
        p = self.farm_path / "insemination_settings.json"
        if p.exists():
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.insemination_types = data.get("insemination_types", {})
            except Exception as e:
                logger.warning(f"授精設定読み込み: {e}")
                self.insemination_types = {}

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="Excelファイル（DC305形式）", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=72, state="readonly").pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(file_frame, text="参照...", command=self._select_file).pack(side=tk.LEFT)

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
            columns=("cow_id", "event_name", "event_date", "remark", "falcon_type", "status"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        self.tree.heading("cow_id", text="ID")
        self.tree.heading("event_name", text="イベント名")
        self.tree.heading("event_date", text="日付")
        self.tree.heading("remark", text="Remark")
        self.tree.heading("falcon_type", text="FALCONイベント")
        self.tree.heading("status", text="状態")
        for col in ("cow_id", "event_name", "event_date", "remark", "falcon_type", "status"):
            self.tree.column(col, width=100)
        self.tree.column("status", width=140)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)

        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        self.info_label = ttk.Label(info_frame, text="DC305形式のExcel（列A=ID, B=イベント名, D=日付）を選択してください")
        self.info_label.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="取り込み実行", command=self._execute_import).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="DC305形式Excelを選択",
            filetypes=[
                ("Excel", "*.xlsx;*.xls"),
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
            self.parsed_rows = []
            self._update_preview()
            self.info_label.config(text="データ件数: 0 件")
            return

        default_year = datetime.now().year
        rows = []
        for _, r in df.iterrows():
            arr = r.tolist()
            while len(arr) <= COL_IGNORE_J:
                arr.append(None)
            cow_id = _to_cow_id(arr[COL_ID] if COL_ID < len(arr) else None)
            event_name_raw = arr[COL_EVENT_NAME] if COL_EVENT_NAME < len(arr) else None
            event_name = str(event_name_raw).strip().upper() if event_name_raw is not None else ""
            event_date = _normalize_date_cell(arr[COL_DATE] if COL_DATE < len(arr) else None, default_year)
            if not cow_id or not event_name or not event_date:
                continue
            if event_name not in EVENT_NAME_MAP:
                continue
            remark = ""
            if COL_REMARK < len(arr) and arr[COL_REMARK] is not None:
                remark = str(arr[COL_REMARK]).strip()
            r_right, r_left, r_other = _parse_r_opr(arr[COL_R_OPR] if COL_R_OPR < len(arr) else None)
            insem_type_raw = arr[COL_INSEM_TYPE] if COL_INSEM_TYPE < len(arr) else None
            insem_type_excel = str(insem_type_raw).strip() if insem_type_raw is not None else ""

            row = {
                "cow_id": cow_id,
                "event_name": event_name,
                "event_date": event_date,
                "remark": remark,
                "right_ovary": r_right,
                "left_ovary": r_left,
                "other_findings": r_other,
                "insem_type_excel": insem_type_excel,
            }
            if event_name == "OK":
                row["falcon_event_number"] = None
            elif event_name == "SOLD":
                row["falcon_event_number"] = RuleEngine.EVENT_SOLD
            elif event_name == "FRESH":
                row["falcon_event_number"] = RuleEngine.EVENT_CALV
            elif event_name == "BRED":
                row["falcon_event_number"] = RuleEngine.EVENT_AI
            elif event_name == "OPEN":
                row["falcon_event_number"] = RuleEngine.EVENT_AI
            elif event_name == "PREG":
                row["falcon_event_number"] = RuleEngine.EVENT_AI
            elif event_name == "DNB":
                row["falcon_event_number"] = RuleEngine.EVENT_STOPR
            elif event_name == "DIED":
                row["falcon_event_number"] = RuleEngine.EVENT_DEAD
            else:
                continue
            rows.append(row)

        _classify_ok_events(rows)
        self.parsed_rows = rows
        self._update_preview()
        self.info_label.config(text=f"データ件数: {len(self.parsed_rows)} 件")

    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        name_map = {
            RuleEngine.EVENT_SOLD: "売却",
            RuleEngine.EVENT_CALV: "分娩",
            RuleEngine.EVENT_FCHK: "フレッシュ",
            RuleEngine.EVENT_REPRO: "繁殖検査",
            RuleEngine.EVENT_AI: "AI",
            RuleEngine.EVENT_STOPR: "繁殖停止",
            RuleEngine.EVENT_DEAD: "死亡",
        }
        for row in self.parsed_rows:
            fn = row.get("falcon_event_number")
            falcon_type = name_map.get(fn, str(fn))
            cow = self._resolve_cow(row.get("cow_id", ""))
            status = "FALCONにありません" if not cow else "吸い込み対象"
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("cow_id", ""),
                    row.get("event_name", ""),
                    row.get("event_date", ""),
                    (row.get("remark", "") or "")[:20],
                    falcon_type,
                    status,
                ),
            )

    def _resolve_cow(self, cow_id: str) -> Optional[Dict[str, Any]]:
        """4桁IDで個体を検索"""
        if not cow_id:
            return None
        return self.db.get_cow_by_id(cow_id)

    def _show_insem_mapping_dialog(self) -> bool:
        """BRED/OPEN/PREG に登場する授精種類（H列）をFALCONの授精種類にマッピングするダイアログ"""
        unique_excel = set()
        for row in self.parsed_rows:
            if row.get("event_name") in ("BRED", "OPEN", "PREG") and row.get("insem_type_excel"):
                unique_excel.add(row["insem_type_excel"])
        if not unique_excel:
            return True

        dialog = tk.Toplevel(self.window)
        dialog.title("授精種類のマッピング")
        dialog.geometry("480x320")
        dialog.transient(self.window)
        dialog.grab_set()
        # 最前面に表示
        dialog.lift()
        dialog.attributes("-topmost", True)
        dialog.after(100, lambda: dialog.attributes("-topmost", False))
        try:
            dialog.focus_force()
        except tk.TclError:
            pass

        ttk.Label(
            dialog,
            text="Excelの授精種類（B欄）をFALCONの授精種類に割り当ててください。",
            font=("", 10),
        ).pack(pady=(15, 10), padx=15)
        frame = ttk.Frame(dialog, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)
        combos = {}
        falcon_options = [f"{code}: {name}" for code, name in sorted(self.insemination_types.items(), key=lambda x: (str(x[0])))]
        if not falcon_options:
            falcon_options = ["（授精種類が未設定です）"]
        for i, excel_val in enumerate(sorted(unique_excel)):
            ttk.Label(frame, text=f"Excel「{excel_val}」→").grid(row=i, column=0, sticky=tk.W, padx=(0, 8), pady=4)
            var = tk.StringVar()
            cb = ttk.Combobox(frame, textvariable=var, values=falcon_options, width=28, state="readonly")
            if falcon_options and falcon_options[0] != "（授精種類が未設定です）":
                cb.set(falcon_options[0])
            cb.grid(row=i, column=1, sticky=tk.W, pady=4)
            combos[excel_val] = (var, cb)

        result = [False]

        def on_ok():
            for excel_val, (var, cb) in combos.items():
                val = var.get().strip()
                if val and ":" in val:
                    code = val.split(":", 1)[0].strip()
                    self.insem_type_mapping[excel_val] = code
                else:
                    self.insem_type_mapping[excel_val] = ""
            result[0] = True
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_f = ttk.Frame(dialog)
        btn_f.pack(pady=15)
        ttk.Button(btn_f, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_f, text="キャンセル", command=on_cancel).pack(side=tk.LEFT)
        dialog.wait_window()
        return result[0]

    def _event_already_exists(self, cow_auto_id: int, event_number: int, event_date: str) -> bool:
        """同一個体・同日・同一イベント種別が既に存在するか"""
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for ev in events:
            if ev.get("event_number") == event_number and ev.get("event_date") == event_date:
                return True
        return False

    def _cow_has_ai_or_et_after(self, cow_auto_id: int, after_date: str) -> bool:
        """当該個体に指定日より後のAIまたはETイベントが存在するか"""
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for ev in events:
            if ev.get("event_number") in (RuleEngine.EVENT_AI, RuleEngine.EVENT_ET):
                if (ev.get("event_date") or "") > after_date:
                    return True
        return False

    def _execute_import(self):
        if not self.parsed_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return
        ai_rows = [r for r in self.parsed_rows if r.get("event_name") in ("BRED", "OPEN", "PREG")]
        if ai_rows and not self._show_insem_mapping_dialog():
            return

        import_note = f"{DC305_NOTE_PREFIX}{datetime.now().strftime('%Y-%m-%d')}"
        # 確認ダイアログを最前面に表示
        self.window.lift()
        self.window.attributes("-topmost", True)
        self.window.after(100, lambda: self.window.attributes("-topmost", False))
        try:
            self.window.focus_force()
        except tk.TclError:
            pass
        ok = messagebox.askyesno(
            "確認",
            f"DC305イベントを {len(self.parsed_rows)} 行分取り込みますか？\n"
            "（OPENは、その後にAI/ETがなければ妊鑑−も登録。PREGはAI+妊鑑+で2件登録）\n"
            f"吸い込み日は「{import_note}」で記録され、後から一括削除できます。",
        )
        if not ok:
            return

        # 処理中に他アプリにフォーカスが移っている場合に備え、完了時に前面へ
        def bring_parent_to_front():
            self.window.lift()
            self.window.attributes("-topmost", True)
            self.window.after(100, lambda: self.window.attributes("-topmost", False))
            try:
                self.window.focus_force()
            except tk.TclError:
                pass

        success = 0
        skip_and_errors: List[Dict[str, Any]] = []

        for row in self.parsed_rows:
            cow_id = row.get("cow_id", "")
            cow = self._resolve_cow(cow_id)
            if not cow:
                skip_and_errors.append({
                    "cow_id": cow_id,
                    "event_name": row.get("event_name", ""),
                    "event_date": row.get("event_date", ""),
                    "reason": "FALCONに個体がありません",
                })
                continue
            cow_auto_id = cow.get("auto_id") or cow.get("auto_ID")
            if not cow_auto_id:
                skip_and_errors.append({
                    "cow_id": cow_id,
                    "event_name": row.get("event_name", ""),
                    "event_date": row.get("event_date", ""),
                    "reason": "auto_id が取得できません",
                })
                continue

            event_name = row.get("event_name", "")
            event_date = row.get("event_date", "")
            remark = row.get("remark", "")
            falcon_event_number = row.get("falcon_event_number")

            skip_reason = None
            if event_name == "SOLD" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_SOLD, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "FRESH" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_CALV, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "OK" and self._event_already_exists(cow_auto_id, falcon_event_number, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "BRED" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_AI, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "OPEN" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_AI, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "PREG" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_AI, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "DNB" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_STOPR, event_date):
                skip_reason = "既存のためスキップ"
            elif event_name == "DIED" and self._event_already_exists(cow_auto_id, RuleEngine.EVENT_DEAD, event_date):
                skip_reason = "既存のためスキップ"
            if skip_reason:
                skip_and_errors.append({
                    "cow_id": cow_id,
                    "event_name": event_name,
                    "event_date": event_date,
                    "reason": skip_reason,
                })
                continue

            try:
                if event_name == "SOLD":
                    self._insert_simple(cow_auto_id, RuleEngine.EVENT_SOLD, event_date, remark, import_note)
                    success += 1
                elif event_name == "FRESH":
                    self._insert_simple(cow_auto_id, RuleEngine.EVENT_CALV, event_date, remark, import_note)
                    success += 1
                elif event_name == "OK":
                    self._insert_repro_check(cow_auto_id, falcon_event_number, event_date, remark, import_note)
                    success += 1
                elif event_name == "BRED":
                    self._insert_ai(cow_auto_id, event_date, row, import_note)
                    success += 1
                elif event_name == "OPEN":
                    self._insert_ai(cow_auto_id, event_date, row, import_note)
                    success += 1
                    # 当該個体で、このOPEN日付より後にAI/ETがなければ妊鑑−を登録
                    has_later_in_file = any(
                        r for r in self.parsed_rows
                        if r.get("cow_id") == cow_id
                        and r.get("event_name") in ("BRED", "OPEN", "PREG")
                        and (r.get("event_date") or "") > event_date
                    )
                    if not has_later_in_file and not self._cow_has_ai_or_et_after(cow_auto_id, event_date):
                        self._insert_simple(
                            cow_auto_id,
                            RuleEngine.EVENT_PDN,
                            (datetime.strptime(event_date, "%Y-%m-%d") + timedelta(days=30)).strftime("%Y-%m-%d"),
                            "",
                            import_note,
                        )
                        success += 1
                elif event_name == "PREG":
                    self._insert_ai(cow_auto_id, event_date, row, import_note)
                    self._insert_simple(
                        cow_auto_id,
                        RuleEngine.EVENT_PDP,
                        (datetime.strptime(event_date, "%Y-%m-%d") + timedelta(days=30)).strftime("%Y-%m-%d"),
                        "",
                        import_note,
                    )
                    success += 2
                elif event_name == "DNB":
                    self._insert_simple(cow_auto_id, RuleEngine.EVENT_STOPR, event_date, remark, import_note)
                    success += 1
                elif event_name == "DIED":
                    self._insert_simple(cow_auto_id, RuleEngine.EVENT_DEAD, event_date, remark, import_note)
                    success += 1
            except Exception as e:
                logger.error(f"DC305取込エラー: {e}", exc_info=True)
                skip_and_errors.append({
                    "cow_id": row.get("cow_id", ""),
                    "event_name": event_name,
                    "event_date": event_date,
                    "reason": f"イベント作成エラー: {str(e)}",
                })

        skip_count = len([r for r in skip_and_errors if r.get("reason") == "既存のためスキップ"])
        error_count = len(skip_and_errors) - skip_count
        msg = f"取り込み完了\n成功: {success} 件\nスキップ: {skip_count} 件\nエラー・個体なし: {error_count} 件"
        bring_parent_to_front()
        if skip_and_errors:
            msg += "\n\nスキップ・エラー一覧を表示しますか？"
            if messagebox.askyesno("完了", msg):
                self._show_skip_and_errors_window(skip_and_errors)
        else:
            messagebox.showinfo("完了", msg)
        try:
            if self.window.winfo_exists():
                self._update_preview()
        except tk.TclError:
            pass

    def _insert_simple(
        self,
        cow_auto_id: int,
        event_number: int,
        event_date: str,
        note: str,
        import_note: str,
    ):
        event_data = {
            "cow_auto_id": cow_auto_id,
            "event_number": event_number,
            "event_date": event_date,
            "json_data": {"other": note} if note else None,
            "note": import_note,
        }
        event_id = self.db.insert_event(event_data)
        self.rule_engine.on_event_added(event_id)
        self.rule_engine._recalculate_and_update_cow(cow_auto_id)

    def _insert_repro_check(
        self,
        cow_auto_id: int,
        event_number: int,
        event_date: str,
        note: str,
        import_note: str,
    ):
        json_data = {}
        if note:
            json_data["other"] = note
        event_data = {
            "cow_auto_id": cow_auto_id,
            "event_number": event_number,
            "event_date": event_date,
            "json_data": json_data if json_data else None,
            "note": import_note,
        }
        event_id = self.db.insert_event(event_data)
        self.rule_engine.on_event_added(event_id)
        self.rule_engine._recalculate_and_update_cow(cow_auto_id)

    def _insert_ai(self, cow_auto_id: int, event_date: str, row: Dict[str, Any], import_note: str):
        json_data = {}
        if row.get("remark"):
            json_data["sire"] = str(row["remark"]).strip().upper()
        right = (row.get("right_ovary") or "").strip().upper()
        left = row.get("left_ovary", "")
        other = row.get("other_findings", "")
        # F列（R欄）の第1要素を結果コード（O/P/R）として解釈。"-"や空は未鑑定とする
        if right in ("O", "P", "R"):
            json_data["outcome"] = right
            # 結果コードの場合は right_ovary_findings には入れない
        elif right in ("-", "") or not right:
            json_data["_dc305_no_result"] = True
            # 未鑑定のため outcome は付けない（RuleEngine で O に上書きされないようにする）
        else:
            json_data["right_ovary_findings"] = row.get("right_ovary", "").strip()
        if left:
            json_data["left_ovary_findings"] = left
        if other:
            json_data["other"] = other
        excel_insem = row.get("insem_type_excel", "")
        if excel_insem and excel_insem in self.insem_type_mapping:
            code = self.insem_type_mapping[excel_insem]
            if code:
                json_data["insemination_type_code"] = code
        event_data = {
            "cow_auto_id": cow_auto_id,
            "event_number": RuleEngine.EVENT_AI,
            "event_date": event_date,
            "json_data": json_data if json_data else None,
            "note": import_note,
        }
        event_id = self.db.insert_event(event_data)
        self.rule_engine.on_event_added(event_id)
        self.rule_engine._recalculate_and_update_cow(cow_auto_id)

    def _show_skip_and_errors_window(self, skip_and_errors: List[Dict[str, Any]]):
        """スキップ・エラー一覧を別ウィンドウで表示（乳検・AI吸い込みと同様）。テキスト保存可能。"""
        # 親はメインウィンドウを使用（DC305ウィンドウが既に閉じられている場合でも表示できるようにする）
        parent = self.parent
        try:
            if self.window.winfo_exists():
                parent = self.window
        except tk.TclError:
            pass
        win = tk.Toplevel(parent)
        win.title("スキップ・エラー一覧（DC305取込）")
        win.geometry("720x420")
        # 起動直後に最前面に表示（気づきやすくする）
        win.lift()
        win.attributes("-topmost", True)
        win.after(100, lambda: win.attributes("-topmost", False))
        try:
            win.focus_force()
        except tk.TclError:
            pass

        main_frame = ttk.Frame(win, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            main_frame,
            text=f"個体が存在しない・既存のためスキップ・エラー（{len(skip_and_errors)}件）",
            font=("", 10, "bold"),
        ).pack(pady=(0, 10))

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        tree = ttk.Treeview(
            tree_frame,
            columns=("cow_id", "event_name", "event_date", "reason"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        tree.heading("cow_id", text="ID")
        tree.heading("event_name", text="イベント名")
        tree.heading("event_date", text="日付")
        tree.heading("reason", text="理由")
        tree.column("cow_id", width=80)
        tree.column("event_name", width=100)
        tree.column("event_date", width=100)
        tree.column("reason", width=280)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)

        for r in skip_and_errors:
            tree.insert(
                "",
                tk.END,
                values=(
                    r.get("cow_id", ""),
                    r.get("event_name", ""),
                    r.get("event_date", ""),
                    r.get("reason", ""),
                ),
            )

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            btn_frame,
            text="テキストファイルとして保存",
            command=lambda: self._save_skip_and_errors_to_file(skip_and_errors, win),
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="閉じる", command=win.destroy).pack(side=tk.LEFT)

    def _save_skip_and_errors_to_file(
        self, skip_and_errors: List[Dict[str, Any]], parent_window: tk.Toplevel
    ):
        """スキップ・エラー一覧をテキストファイルに保存"""
        path = filedialog.asksaveasfilename(
            title="スキップ・エラー一覧を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("=" * 70 + "\n")
                f.write("DC305取込 スキップ・エラー一覧\n")
                f.write("=" * 70 + "\n")
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"件数: {len(skip_and_errors)} 件\n")
                f.write("=" * 70 + "\n\n")
                f.write(f"{'ID':<10} {'イベント名':<12} {'日付':<12} {'理由':<40}\n")
                f.write("-" * 70 + "\n")
                for r in skip_and_errors:
                    f.write(
                        f"{r.get('cow_id', ''):<10} {r.get('event_name', ''):<12} {r.get('event_date', ''):<12} {r.get('reason', ''):<40}\n"
                    )
            messagebox.showinfo("保存完了", f"スキップ・エラー一覧を保存しました:\n{path}")
        except Exception as e:
            logger.error(f"保存エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ファイルの保存に失敗しました:\n{e}")

    def show(self):
        self.window.focus_set()
