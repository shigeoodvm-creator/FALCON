"""
FALCON2 - データ出力ウィンドウ
FALCONイベントデータ / ゲノムデータをCSV出力する。
"""

import csv
import json
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from modules.rule_engine import RuleEngine


class DataExportWindow:
    """データ出力ウィンドウ（FALCONイベント / ゲノム）"""

    def __init__(
        self,
        parent: tk.Tk,
        db_handler,
        event_dictionary_path: Optional[Path],
        mode: str,
    ):
        self.parent = parent
        self.db = db_handler
        self.event_dictionary_path = Path(event_dictionary_path) if event_dictionary_path else None
        self.mode = mode  # "falcon_events" or "genome"
        self.event_dictionary = self._load_event_dictionary()

        self.window = tk.Toplevel(parent)
        self.window.configure(bg="#f5f5f5")
        self.window.minsize(640, 500)
        self.window.geometry("760x620")
        self.window.title("FALCONイベントデータ出力" if self.mode == "falcon_events" else "ゲノムデータ出力")

        self._build_ui()

    def _load_event_dictionary(self) -> Dict[str, Dict[str, Any]]:
        paths_to_try: List[Path] = []
        if self.event_dictionary_path:
            paths_to_try.append(self.event_dictionary_path)
        from constants import CONFIG_DEFAULT_DIR
        paths_to_try.append(CONFIG_DEFAULT_DIR / "event_dictionary.json")

        for p in paths_to_try:
            if not p.exists():
                continue
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                continue
        return {}

    def _build_ui(self):
        df = "Meiryo UI"

        header = tk.Frame(self.window, bg="#f5f5f5", padx=24, pady=20)
        header.pack(fill=tk.X)
        icon = "\u2b06\ufe0f" if self.mode == "falcon_events" else "\U0001F9EC"
        title = "FALCONイベントデータ"
        subtitle = "期間とイベントを指定してCSV出力します（2層CSV推奨）"
        if self.mode == "genome":
            title = "ゲノムデータ"
            subtitle = "期間内のゲノムイベント（602）をCSV出力します"
        icon_font = ("Segoe UI Emoji", 24) if platform.system() == "Windows" else (df, 24)
        tk.Label(header, text=icon, font=icon_font, bg="#f5f5f5", fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        tf = tk.Frame(header, bg="#f5f5f5")
        tf.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(tf, text=title, font=(df, 16, "bold"), bg="#f5f5f5", fg="#263238").pack(anchor=tk.W)
        tk.Label(tf, text=subtitle, font=(df, 10), bg="#f5f5f5", fg="#607d8b").pack(anchor=tk.W)

        body = tk.Frame(self.window, bg="#f5f5f5", padx=24, pady=8)
        body.pack(fill=tk.BOTH, expand=True)

        period_card = tk.Frame(body, bg="#f5f5f5", padx=16, pady=12, highlightbackground="#e0e7ef", highlightthickness=1)
        period_card.pack(fill=tk.X, pady=(0, 8))
        tk.Label(period_card, text="期間", font=(df, 11, "bold"), bg="#f5f5f5", fg="#263238").grid(row=0, column=0, sticky="w", pady=(0, 8), columnspan=4)

        today = datetime.now().date()
        default_start = (today - timedelta(days=365)).strftime("%Y-%m-%d")
        default_end = today.strftime("%Y-%m-%d")

        tk.Label(period_card, text="開始日", font=(df, 9), bg="#f5f5f5", fg="#546e7a").grid(row=1, column=0, sticky="w")
        self.start_var = tk.StringVar(value=default_start)
        ttk.Entry(period_card, textvariable=self.start_var, width=14).grid(row=1, column=1, padx=(8, 16), sticky="w")

        tk.Label(period_card, text="終了日", font=(df, 9), bg="#f5f5f5", fg="#546e7a").grid(row=1, column=2, sticky="w")
        self.end_var = tk.StringVar(value=default_end)
        ttk.Entry(period_card, textvariable=self.end_var, width=14).grid(row=1, column=3, padx=(8, 0), sticky="w")

        tk.Label(period_card, text="形式: YYYY-MM-DD", font=(df, 8), bg="#f5f5f5", fg="#90a4ae").grid(
            row=2, column=0, sticky="w", columnspan=4, pady=(6, 0)
        )

        if self.mode == "falcon_events":
            self._build_falcon_event_options(body)
        else:
            self._build_genome_options(body)

        footer = tk.Frame(self.window, bg="#f5f5f5", pady=18)
        footer.pack(fill=tk.X)
        tk.Button(
            footer,
            text="出力実行",
            font=(df, 10),
            bg="#3949ab",
            fg="#ffffff",
            activebackground="#303f9f",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=16,
            pady=8,
            cursor="hand2",
            command=self._on_execute,
        ).pack(side=tk.LEFT, padx=(24, 8))
        tk.Button(
            footer,
            text="閉じる",
            font=(df, 10),
            bg="#fafafa",
            fg="#546e7a",
            activebackground="#eceff1",
            relief=tk.FLAT,
            padx=16,
            pady=8,
            highlightbackground="#b0bec5",
            highlightthickness=1,
            cursor="hand2",
            command=self.window.destroy,
        ).pack(side=tk.LEFT)

    def _build_falcon_event_options(self, parent: tk.Frame):
        df = "Meiryo UI"
        card = tk.Frame(parent, bg="#f5f5f5", padx=16, pady=12, highlightbackground="#e0e7ef", highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        top = tk.Frame(card, bg="#f5f5f5")
        top.pack(fill=tk.X)
        tk.Label(top, text="イベント種類（複数選択）", font=(df, 11, "bold"), bg="#f5f5f5", fg="#263238").pack(side=tk.LEFT)
        ttk.Button(top, text="全選択", command=self._select_all_events).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(top, text="全解除", command=self._clear_all_events).pack(side=tk.RIGHT)

        list_wrap = tk.Frame(card, bg="#f5f5f5")
        list_wrap.pack(fill=tk.BOTH, expand=True, pady=(8, 8))
        scrollbar = ttk.Scrollbar(list_wrap, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.event_listbox = tk.Listbox(
            list_wrap,
            selectmode=tk.MULTIPLE,
            exportselection=False,
            yscrollcommand=scrollbar.set,
            height=12,
        )
        self.event_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.event_listbox.yview)

        self._event_numbers: List[int] = []
        items = self._get_available_events()
        for ev_no, name_jp, alias in items:
            label = f"{ev_no}: {name_jp}" if not alias else f"{ev_no}: {name_jp} ({alias})"
            self.event_listbox.insert(tk.END, label)
            self._event_numbers.append(ev_no)
        self._select_all_events()

        type_frame = tk.Frame(card, bg="#f5f5f5")
        type_frame.pack(fill=tk.X, pady=(2, 0))
        tk.Label(type_frame, text="出力形式", font=(df, 9), bg="#f5f5f5", fg="#546e7a").pack(side=tk.LEFT, padx=(0, 10))
        self.falcon_format_var = tk.StringVar(value="two_layer")
        ttk.Radiobutton(type_frame, text="2層CSV（推奨）", value="two_layer", variable=self.falcon_format_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(type_frame, text="共通CSVのみ", value="main_only", variable=self.falcon_format_var).pack(side=tk.LEFT)

    def _build_genome_options(self, parent: tk.Frame):
        card = tk.Frame(parent, bg="#f5f5f5", padx=16, pady=12, highlightbackground="#e0e7ef", highlightthickness=1)
        card.pack(fill=tk.X, pady=(0, 8))
        self.genome_latest_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(card, text="個体ごとに最新イベントのみ出力", variable=self.genome_latest_only).pack(anchor=tk.W)
        tk.Label(
            card,
            text="オフにすると期間内の全ゲノムイベントを出力します",
            font=("Meiryo UI", 8),
            bg="#f5f5f5",
            fg="#90a4ae",
        ).pack(anchor=tk.W, pady=(4, 0))

    def _get_available_events(self) -> List[Tuple[int, str, str]]:
        rows: List[Tuple[int, str, str]] = []
        for k, v in (self.event_dictionary or {}).items():
            if not isinstance(v, dict):
                continue
            if v.get("deprecated", False):
                continue
            try:
                ev_no = int(k)
            except (TypeError, ValueError):
                continue
            name_jp = str(v.get("name_jp") or v.get("alias") or ev_no)
            alias = str(v.get("alias") or "")
            rows.append((ev_no, name_jp, alias))
        rows.sort(key=lambda x: x[0])
        return rows

    def _select_all_events(self):
        if not hasattr(self, "event_listbox"):
            return
        self.event_listbox.select_set(0, tk.END)

    def _clear_all_events(self):
        if not hasattr(self, "event_listbox"):
            return
        self.event_listbox.selection_clear(0, tk.END)

    def _parse_date_range(self) -> Optional[Tuple[str, str]]:
        s = self.start_var.get().strip()
        e = self.end_var.get().strip()
        try:
            sd = datetime.strptime(s, "%Y-%m-%d").date()
            ed = datetime.strptime(e, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showwarning("入力エラー", "開始日・終了日は YYYY-MM-DD 形式で入力してください。")
            return None
        if sd > ed:
            messagebox.showwarning("入力エラー", "開始日は終了日以前にしてください。")
            return None
        return sd.strftime("%Y-%m-%d"), ed.strftime("%Y-%m-%d")

    def _on_execute(self):
        date_range = self._parse_date_range()
        if not date_range:
            return
        start_date, end_date = date_range

        if self.mode == "falcon_events":
            self._export_falcon_events(start_date, end_date)
        else:
            self._export_genome(start_date, end_date)

    def _event_labels(self, event_number: int) -> Tuple[str, str, str]:
        ev = (self.event_dictionary or {}).get(str(event_number), {})
        alias = str(ev.get("alias") or "")
        name_jp = str(ev.get("name_jp") or alias or event_number)
        category = str(ev.get("category") or "")
        return alias, name_jp, category

    def _field_label(self, event_number: int, key: str) -> str:
        ev = (self.event_dictionary or {}).get(str(event_number), {})
        fields = ev.get("input_fields") or []
        if isinstance(fields, list):
            for f in fields:
                if not isinstance(f, dict):
                    continue
                if str(f.get("key") or "") == key:
                    return str(f.get("label") or key)
        return key

    @staticmethod
    def _to_cell_string(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, (dict, list)):
            try:
                return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
            except Exception:
                return str(value)
        return str(value)

    @staticmethod
    def _try_float(value: Any) -> Optional[float]:
        if value is None:
            return None
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        try:
            return float(str(value).replace(",", "").strip())
        except Exception:
            return None

    @staticmethod
    def _try_bool(value: Any) -> Optional[bool]:
        if isinstance(value, bool):
            return value
        s = str(value).strip().lower()
        if s in ("true", "1", "yes", "y", "on"):
            return True
        if s in ("false", "0", "no", "n", "off"):
            return False
        return None

    def _export_falcon_events(self, start_date: str, end_date: str):
        if not hasattr(self, "event_listbox"):
            return
        selected_indices = list(self.event_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("入力エラー", "イベント種類を1つ以上選択してください。")
            return
        selected_numbers = {self._event_numbers[i] for i in selected_indices if i < len(self._event_numbers)}

        events = self.db.get_events_by_period(start_date, end_date, include_deleted=False)
        target_events = [e for e in events if int(e.get("event_number") or -1) in selected_numbers]
        if not target_events:
            messagebox.showinfo("情報", "条件に一致するイベントがありません。")
            return

        main_rows: List[Dict[str, Any]] = []
        field_rows: List[Dict[str, Any]] = []

        for ev in target_events:
            event_number = int(ev.get("event_number") or 0)
            alias, name_jp, category = self._event_labels(event_number)
            cow_auto_id = ev.get("cow_auto_id")
            cow = self.db.get_cow_by_auto_id(cow_auto_id) if cow_auto_id else None

            main_rows.append(
                {
                    "event_id": ev.get("id"),
                    "event_date": ev.get("event_date"),
                    "cow_auto_id": cow_auto_id,
                    "cow_id": (cow or {}).get("cow_id", ""),
                    "jpn10": (cow or {}).get("jpn10", ""),
                    "event_number": event_number,
                    "event_alias": alias,
                    "event_name_jp": name_jp,
                    "category": category,
                    "note": ev.get("note", ""),
                    "event_dim": ev.get("event_dim"),
                    "event_lact": ev.get("event_lact"),
                    "deleted": ev.get("deleted", 0),
                }
            )

            json_data = ev.get("json_data") or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            if not isinstance(json_data, dict):
                json_data = {}

            for key, value in json_data.items():
                num_value = self._try_float(value)
                bool_value = self._try_bool(value)
                field_rows.append(
                    {
                        "event_id": ev.get("id"),
                        "event_date": ev.get("event_date"),
                        "cow_auto_id": cow_auto_id,
                        "cow_id": (cow or {}).get("cow_id", ""),
                        "jpn10": (cow or {}).get("jpn10", ""),
                        "event_number": event_number,
                        "event_alias": alias,
                        "event_name_jp": name_jp,
                        "field_key": key,
                        "field_label": self._field_label(event_number, key),
                        "value_raw": self._to_cell_string(value),
                        "value_num": "" if num_value is None else num_value,
                        "value_bool": "" if bool_value is None else int(bool_value),
                    }
                )

        default_name = f"falcon_events_{start_date}_{end_date}.csv"
        main_path = filedialog.asksaveasfilename(
            title="出力先を選択",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not main_path:
            return

        main_columns = [
            "event_id",
            "event_date",
            "cow_auto_id",
            "cow_id",
            "jpn10",
            "event_number",
            "event_alias",
            "event_name_jp",
            "category",
            "note",
            "event_dim",
            "event_lact",
            "deleted",
        ]
        self._write_csv(Path(main_path), main_columns, main_rows)

        if self.falcon_format_var.get() == "two_layer":
            field_columns = [
                "event_id",
                "event_date",
                "cow_auto_id",
                "cow_id",
                "jpn10",
                "event_number",
                "event_alias",
                "event_name_jp",
                "field_key",
                "field_label",
                "value_raw",
                "value_num",
                "value_bool",
            ]
            fields_path = Path(main_path).with_name(Path(main_path).stem + "_fields.csv")
            self._write_csv(fields_path, field_columns, field_rows)
            messagebox.showinfo(
                "完了",
                f"出力しました。\n\n"
                f"- 共通CSV: {main_path}\n"
                f"- 項目CSV: {fields_path}\n"
                f"\nイベント件数: {len(main_rows)}件\n項目件数: {len(field_rows)}件",
            )
        else:
            messagebox.showinfo("完了", f"出力しました。\n{main_path}\n\nイベント件数: {len(main_rows)}件")

    def _export_genome(self, start_date: str, end_date: str):
        events = self.db.get_events_by_number_and_period(
            RuleEngine.EVENT_GENOMIC,
            start_date,
            end_date,
            include_deleted=False,
        )
        if not events:
            messagebox.showinfo("情報", "条件に一致するゲノムイベントがありません。")
            return

        if self.genome_latest_only.get():
            latest: Dict[Any, Dict[str, Any]] = {}
            for ev in events:
                cow_auto_id = ev.get("cow_auto_id")
                if cow_auto_id is None:
                    continue
                current = latest.get(cow_auto_id)
                if current is None:
                    latest[cow_auto_id] = ev
                    continue
                current_key = (str(current.get("event_date") or ""), int(current.get("id") or 0))
                new_key = (str(ev.get("event_date") or ""), int(ev.get("id") or 0))
                if new_key > current_key:
                    latest[cow_auto_id] = ev
            target_events = list(latest.values())
        else:
            target_events = events

        all_keys = set()
        parsed_rows: List[Tuple[Dict[str, Any], Dict[str, Any], Dict[str, Any]]] = []
        for ev in target_events:
            json_data = ev.get("json_data") or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            if not isinstance(json_data, dict):
                json_data = {}
            all_keys.update(json_data.keys())
            cow = self.db.get_cow_by_auto_id(ev.get("cow_auto_id")) if ev.get("cow_auto_id") else {}
            parsed_rows.append((ev, json_data, cow or {}))

        genome_keys = sorted(all_keys)
        rows: List[Dict[str, Any]] = []
        for ev, genome_json, cow in parsed_rows:
            row: Dict[str, Any] = {
                "event_id": ev.get("id"),
                "event_date": ev.get("event_date"),
                "cow_auto_id": ev.get("cow_auto_id"),
                "cow_id": cow.get("cow_id", ""),
                "jpn10": cow.get("jpn10", ""),
            }
            for k in genome_keys:
                row[k] = self._to_cell_string(genome_json.get(k))
            rows.append(row)

        default_name = f"genome_data_{start_date}_{end_date}.csv"
        output_path = filedialog.asksaveasfilename(
            title="出力先を選択",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if not output_path:
            return

        columns = ["event_id", "event_date", "cow_auto_id", "cow_id", "jpn10"] + genome_keys
        self._write_csv(Path(output_path), columns, rows)
        mode_label = "個体ごと最新" if self.genome_latest_only.get() else "期間内全件"
        messagebox.showinfo("完了", f"出力しました。\n{output_path}\n\n出力モード: {mode_label}\n件数: {len(rows)}件")

    def _write_csv(self, path: Path, columns: List[str], rows: List[Dict[str, Any]]):
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)

    def show(self):
        self.window.focus_set()

