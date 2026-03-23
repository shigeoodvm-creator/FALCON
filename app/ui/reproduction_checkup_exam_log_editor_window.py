"""
FALCON2 - 検診ログ編集ウィンドウ
繁殖検診フローから開き、直近4か月分の日誌（notes/ratings）を一覧表示し、
右クリックで編集・削除する。編集時は別ウィンドウで信号色・文言を変更可能。
"""

from __future__ import annotations

import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List, Callable

from ui.reproduction_checkup_flow_window import EXAM_LOG_NOTES_FILENAME, EXAM_LOG_RATING_LABELS

VISIT_NOTES_LINES = 8
# 評価は検診入力画面と同じく 緑(good)・黄(normal)・赤(caution) で選択
RATING_VALUES = [
    ("good", "#c8e6c9", "#2e7d32"),    # 緑
    ("normal", "#fff9c4", "#f9a825"),  # 黄
    ("caution", "#ffcdd2", "#c62828"), # 赤
]
RATING_LABELS = ("◎", "○", "△")  # good, normal, caution の表示名（ツールチップ的な簡易表示）

logger = logging.getLogger(__name__)


class ExamLogEditDialog:
    """検診ログ1件を編集するダイアログ（信号色・文言の変更）"""

    def __init__(
        self,
        parent: tk.Widget,
        date_str: str,
        day_data: Dict[str, Any],
        on_save: Callable[[str, List[str], Dict[str, str]], None],
    ):
        self.date_str = date_str
        self.on_save = on_save
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"検診ログの編集 — {date_str}")
        # 文字や行間に少し余裕を持たせたサイズ
        self.dialog.geometry("520x560")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self.visit_note_entries: List[tk.Entry] = []
        self.visit_ratings: Dict[str, Optional[str]] = {}
        self.visit_rating_buttons: Dict[str, Dict[str, tuple]] = {}
        self._rating_colors: Dict[tuple, tuple] = {}

        self._build(day_data)
        self._center()

    def _build(self, day_data: Dict[str, Any]):
        """検診入力画面の「記録」と同じ構成でレイアウトする。左: 部門・評価、右: 気づいた点。"""

        # ダイアログ全体の背景
        self.dialog.configure(bg="#f5f5f5")

        # 本体フレーム（tk.Frame にして Canvas 描画まわりの不具合を避ける）
        main = tk.Frame(self.dialog, bg="#f5f5f5")
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        frame = ttk.LabelFrame(main, text="記録", padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        inner = ttk.Frame(frame)
        inner.pack(fill=tk.BOTH, expand=True)
        # 左列が圧縮されないように列幅を確保（検診入力画面と同様）
        inner.grid_columnconfigure(0, minsize=260, weight=0)
        inner.grid_columnconfigure(1, weight=1)

        # 信号色（検診入力画面と同じ緑・黄・赤）
        rating_colors = RATING_VALUES  # (value, light, dark) のタプル
        label_width = 8
        row_pady = 2
        memo_row_height = 26  # 左右の行高を揃えつつ、文字が切れない程度に余裕

        # 左ブロック: 部門・評価（8行）
        left_block = ttk.Frame(inner)
        left_block.grid(row=0, column=0, sticky=tk.NW, padx=(0, 12))
        ttk.Label(
            left_block,
            text="部門・評価（緑・黄・赤で選択）",
            font=("", 9)
        ).pack(anchor=tk.W, pady=(0, 4))

        signal_row_width = 230
        for key, label in EXAM_LOG_RATING_LABELS:
            self.visit_ratings[key] = None
            row = tk.Frame(left_block, height=memo_row_height, width=signal_row_width, bg="#ffffff")
            row.pack(anchor=tk.W, pady=row_pady, fill=tk.X)
            row.pack_propagate(False)
            ttk.Label(
                row,
                text=f"{label}:",
                width=label_width,
                anchor=tk.E
            ).pack(side=tk.LEFT, padx=(0, 4))
            btns = {}
            for val, light_color, dark_color in rating_colors:
                canv = tk.Canvas(row, width=22, height=20, bg="#f5f5f5", highlightthickness=0)
                canv.pack(side=tk.LEFT, padx=1)
                oval = canv.create_oval(2, 2, 20, 18, fill=light_color, outline="#9e9e9e", width=1)
                canv.configure(cursor="hand2")
                canv.bind("<Button-1>", lambda e, k=key, v=val: self._on_rating_click(k, v))
                # good/normal/caution ごとの色テーブルを保存（_update_rating_style で使用）
                self._rating_colors[(key, val)] = (light_color, dark_color)
                btns[val] = (canv, oval)
            self.visit_rating_buttons[key] = btns

        # 右ブロック: 気づいた点（箇条書き）8行
        notes_block = ttk.Frame(inner)
        notes_block.grid(row=0, column=1, sticky=tk.NSEW)
        inner.grid_rowconfigure(0, weight=1)
        ttk.Label(
            notes_block,
            text="気づいた点（箇条書き）",
            font=("", 9)
        ).pack(anchor=tk.W, pady=(0, 4))

        self.visit_note_entries.clear()
        for _ in range(VISIT_NOTES_LINES):
            row_note = tk.Frame(notes_block, height=memo_row_height, bg="#ffffff")
            row_note.pack(fill=tk.X, pady=row_pady)
            row_note.pack_propagate(False)
            entry = tk.Entry(row_note, font=("", 10))
            entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.visit_note_entries.append(entry)

        # 既存データの反映
        notes = day_data.get("notes") or []
        if not isinstance(notes, list):
            notes = []
        for i, ent in enumerate(self.visit_note_entries):
            if i < len(notes) and notes[i]:
                ent.insert(0, str(notes[i]))
        ratings = day_data.get("ratings") or {}
        for key in self.visit_ratings:
            v = ratings.get(key)
            self.visit_ratings[key] = v if v in ("good", "normal", "caution") else None
            self._update_rating_style(key)

        # 下部ボタン
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btn_frame, text="保存", command=self._do_save, width=10).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="キャンセル", command=self.dialog.destroy, width=10).pack(side=tk.LEFT)

    def _on_rating_click(self, key: str, value: str):
        if self.visit_ratings.get(key) == value:
            self.visit_ratings[key] = None
        else:
            self.visit_ratings[key] = value
        self._update_rating_style(key)

    def _update_rating_style(self, key: str):
        btns = self.visit_rating_buttons.get(key, {})
        current = self.visit_ratings.get(key)
        for val, (canv, oval) in btns.items():
            light, dark = self._rating_colors.get((key, val), ("#eee", "#999"))
            fill = dark if current == val else light
            outline = "#424242" if current == val else "#9e9e9e"
            canv.itemconfig(oval, fill=fill, outline=outline, width=2 if current == val else 1)

    def _do_save(self):
        notes = []
        for ent in self.visit_note_entries:
            notes.append(ent.get().strip() if ent else "")
        ratings = {k: v for k, v in self.visit_ratings.items() if v}
        self.on_save(self.date_str, notes[:VISIT_NOTES_LINES], ratings)
        self.dialog.destroy()

    def _center(self):
        self.dialog.update_idletasks()
        w = self.dialog.winfo_width()
        h = self.dialog.winfo_height()
        x = max(0, (self.dialog.winfo_screenwidth() // 2) - (w // 2))
        y = max(0, (self.dialog.winfo_screenheight() // 2) - (h // 2))
        self.dialog.geometry(f"+{x}+{y}")


class ReproductionCheckupExamLogEditorWindow:
    """検診ログ編集ウィンドウ（一覧のみ・右クリックで編集・削除）"""

    def __init__(
        self,
        parent: tk.Widget,
        farm_path: Path,
        flow_window: Optional[Any] = None,
    ):
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.flow_window = flow_window  # レポート表示用

        self.window = tk.Toplevel(parent)
        self.window.title("検診ログの編集")
        self.window.geometry("680x520")
        self.window.minsize(560, 400)

        self.data: Dict[str, Dict[str, Any]] = {}

        self._create_widgets()
        self._load_data()
        self._refresh_list()

        self.window.update_idletasks()
        x = max(0, (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2))
        y = max(0, (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2))
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main = ttk.Frame(self.window, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        list_frame = ttk.LabelFrame(main, text="日付一覧（直近4か月）", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True)
        list_inner = ttk.Frame(list_frame)
        list_inner.pack(fill=tk.BOTH, expand=True)
        scroll_y = ttk.Scrollbar(list_inner)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree = ttk.Treeview(
            list_inner,
            columns=("date", "preview", "ratings"),
            show="headings",
            height=20,
            yscrollcommand=scroll_y.set,
        )
        self.tree.heading("date", text="日付")
        self.tree.heading("preview", text="メモ（先頭）")
        self.tree.heading("ratings", text="評価")
        # 列幅を余裕持たせて文字切れを防ぐ（評価は8部門分のテキストが入る）
        self.tree.column("date", width=100)
        self.tree.column("preview", width=260)
        self.tree.column("ratings", width=280)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.config(command=self.tree.yview)

        # 右クリックで編集・削除メニュー（Windows/Linux と Mac 対応）
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Button-2>", self._on_right_click)

        # データがないときの案内
        self.empty_label = ttk.Label(
            main,
            text="検診ログのデータがありません。\n検診入力で記録を保存するとここに表示されます。",
            font=("", 10),
            foreground="gray",
        )

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(14, 0))
        if self.flow_window:
            ttk.Button(btn_frame, text="レポートを表示", command=self._output_report, width=16).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="閉じる", command=self.window.destroy, width=12).pack(side=tk.LEFT)

    def _on_right_click(self, event):
        """一覧行の右クリックでコンテキストメニュー（編集・削除）を表示"""
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        item = self.tree.item(row_id)
        tags = item.get("tags") or ()
        values = item.get("values") or ()
        date_str = tags[0] if tags else (str(values[0]).strip() if values else None)
        if not date_str:
            return
        menu = tk.Menu(self.tree, tearoff=0)
        menu.add_command(
            label="編集",
            command=lambda: self._open_edit_dialog(date_str),
        )
        menu.add_command(
            label="削除",
            command=lambda: self._delete_log(date_str),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _open_edit_dialog(self, date_str: str):
        day = self.data.get(date_str)
        if not isinstance(day, dict):
            day = {}
        def on_save(d: str, notes: List[str], ratings: Dict[str, str]):
            if d not in self.data:
                self.data[d] = {}
            self.data[d]["notes"] = notes
            self.data[d]["ratings"] = ratings
            self.data[d]["updated_at"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            self._save_data()
            self._refresh_list()
            messagebox.showinfo("保存", "保存しました。")
        ExamLogEditDialog(self.window, date_str, day, on_save)

    def _delete_log(self, date_str: str):
        if not messagebox.askyesno("削除の確認", f"{date_str} の検診ログを削除しますか？"):
            return
        if date_str in self.data:
            del self.data[date_str]
            self._save_data()
            self._refresh_list()
            messagebox.showinfo("削除", "削除しました。")

    def _load_data(self):
        path = self.farm_path / EXAM_LOG_NOTES_FILENAME
        self.data = {}
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
                if not isinstance(self.data, dict):
                    self.data = {}
            except Exception as e:
                logger.warning("検診ログ読み込みエラー: %s", e)

    def _save_data(self):
        path = self.farm_path / EXAM_LOG_NOTES_FILENAME
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error("検診ログ保存エラー: %s", e)
            messagebox.showerror("エラー", f"保存に失敗しました。\n{e}")

    def _dates_last_4_months(self) -> List[str]:
        end_date = datetime.now().date()
        start_date = end_date - timedelta(days=4 * 31)
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
        out = []
        for date_str in self.data.keys():
            if not isinstance(self.data[date_str], dict):
                continue
            try:
                if start_str <= date_str <= end_str:
                    out.append(date_str)
            except (TypeError, ValueError):
                continue
        out.sort()
        return out

    def _refresh_list(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        dates = self._dates_last_4_months()
        if not dates:
            self.empty_label.pack(pady=20)
            return
        self.empty_label.pack_forget()
        for date_str in dates:
            day = self.data[date_str]
            notes = day.get("notes") or []
            if not isinstance(notes, list):
                notes = []
            preview = ""
            for line in notes:
                if line and str(line).strip():
                    preview = (str(line).strip())[:32]
                    if len(str(line).strip()) > 32:
                        preview += "…"
                    break
            ratings = day.get("ratings") or {}
            rating_parts = []
            for key, label in EXAM_LOG_RATING_LABELS:
                v = ratings.get(key)
                if v == "good":
                    rating_parts.append(f"{label}青")
                elif v == "normal":
                    rating_parts.append(f"{label}黄")
                elif v == "caution":
                    rating_parts.append(f"{label}赤")
            self.tree.insert("", tk.END, values=(date_str, preview, " ".join(rating_parts)), tags=(date_str,))

    def _output_report(self):
        if self.flow_window and hasattr(self.flow_window, "_output_exam_log_html"):
            self.flow_window._output_exam_log_html()

    def show(self):
        self.window.focus_set()
