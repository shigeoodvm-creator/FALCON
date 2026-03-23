"""
FALCON2 - SIRE一覧ウィンドウ
これまで使用されたSIREをSIRE名順に一覧表示し、
種別（乳用種レギュラー／乳用種♀／F1／黒毛和種／不明その他）を管理（後継牛確保の見通し用）。
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, List, Set, Callable

import tkinter as tk
from tkinter import ttk, messagebox

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.sire_list_opts import (
    SIRE_TYPE_ORDER,
    SIRE_TYPE_LABELS_JA,
    ja_label_to_sire_type,
    sire_opts_to_type,
    sire_type_to_stored_dict,
)

logger = logging.getLogger(__name__)

SIRE_LIST_FILENAME = "sire_list.json"


def get_known_sire_names(db: DBHandler, farm_path: Path) -> Set[str]:
    """既存のSIRE名の集合を返す（sire_list.jsonのキー＋全AI/ETイベントから抽出）"""
    known: Set[str] = set()
    path = Path(farm_path) / SIRE_LIST_FILENAME
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                known.update(str(k).strip() for k in data.keys() if k)
        except Exception as e:
            logger.warning(f"SIRE一覧: 既存SIRE読み込みエラー: {e}")
    try:
        for event in db.get_events_by_number(RuleEngine.EVENT_AI, include_deleted=False) + db.get_events_by_number(RuleEngine.EVENT_ET, include_deleted=False):
            sire = _extract_sire_from_json(event.get("json_data"))
            if sire:
                known.add(sire)
    except Exception as e:
        logger.warning(f"SIRE一覧: イベントからSIRE取得エラー: {e}")
    return known


def _extract_sire_from_json(json_data: Any) -> Optional[str]:
    """イベントのjson_dataからSIRE名を取得。空やNoneの場合はNone。"""
    if json_data is None:
        return None
    if isinstance(json_data, str):
        try:
            json_data = json.loads(json_data)
        except Exception:
            return None
    if not isinstance(json_data, dict):
        return None
    sire = json_data.get("sire") or json_data.get("sire_name") or json_data.get("SIRE")
    if sire is None or (isinstance(sire, str) and not sire.strip()):
        return None
    return str(sire).strip()


def show_sire_confirm_dialog(parent: tk.Tk, farm_path: Path, sire_name: str) -> bool:
    """
    新規SIREの種別を入力するダイアログを表示し、OKならsire_list.jsonに保存する。
    Returns:
        True=保存した, False=キャンセル
    """
    d = SireConfirmDialog(parent, farm_path, sire_name)
    return d.show()


class SireConfirmDialog:
    """新規SIRE入力時の種別選択ダイアログ（5択・ラジオ）"""

    def __init__(self, parent: tk.Tk, farm_path: Path, sire_name: str):
        self.farm_path = Path(farm_path)
        self.sire_name = sire_name.strip()
        self._settings_path = self.farm_path / SIRE_LIST_FILENAME
        self._saved = False

        self.window = tk.Toplevel(parent)
        self.window.title("SIREの登録")
        self.window.configure(bg="#f5f5f5")

        f = ttk.Frame(self.window, padding=16)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text=f"新規SIRE「{self.sire_name}」の特性を登録します。", font=("", 10)).pack(anchor=tk.W, pady=(0, 12))
        self.v_type = tk.StringVar(value=SIRE_TYPE_ORDER[0])
        rb_frame = ttk.Frame(f)
        rb_frame.pack(fill=tk.X, pady=4)
        for key in SIRE_TYPE_ORDER:
            ttk.Radiobutton(
                rb_frame,
                text=SIRE_TYPE_LABELS_JA[key],
                variable=self.v_type,
                value=key,
            ).pack(anchor=tk.W, pady=2)
        btn_row = ttk.Frame(f)
        btn_row.pack(fill=tk.X, pady=(16, 0))
        ttk.Button(btn_row, text="OK", command=self._on_ok).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_row, text="キャンセル", command=self.window.destroy).pack(side=tk.RIGHT, padx=4)

        self.window.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.window.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _on_ok(self) -> None:
        self.farm_path.mkdir(parents=True, exist_ok=True)
        data: Dict[str, Any] = {}
        if self._settings_path.exists():
            try:
                with open(self._settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict):
                    data = {}
            except Exception as e:
                logger.warning(f"SIRE確認: 読み込みエラー: {e}")
        data[self.sire_name] = sire_type_to_stored_dict(self.v_type.get())
        try:
            with open(self._settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._saved = True
        except Exception as e:
            logger.error(f"SIRE確認: 保存エラー: {e}")
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}", parent=self.window)
            return
        self.window.destroy()

    def show(self) -> bool:
        self.window.wait_window()
        return self._saved


class SireListWindow:
    """SIRE一覧ウィンドウ（使用済みSIREを名順表示、種別の編集）"""

    def __init__(
        self,
        parent: tk.Tk,
        db_handler: DBHandler,
        farm_path: Path,
        on_saved: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.db = db_handler
        self.farm_path = Path(farm_path)
        self.on_saved = on_saved

        self._settings_path = self.farm_path / SIRE_LIST_FILENAME
        self._sire_type_by_name: Dict[str, str] = {}  # sire_name -> sire_type キー
        self._sire_names: List[str] = []
        self._type_display_vars: Dict[str, tk.StringVar] = {}  # Combobox 用（日本語ラベル）

        self.window = tk.Toplevel(parent)
        self.window.title("SIRE一覧")
        self.window.geometry("600x560")
        self.window.configure(bg="#f5f5f5")

        self._load_used_sires()
        self._load_settings()
        self._create_widgets()
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _load_used_sires(self) -> None:
        """AI/ETイベントから使用済みSIRE名を重複排除・SIRE名順で取得"""
        seen: Set[str] = set()
        try:
            events_200 = self.db.get_events_by_number(RuleEngine.EVENT_AI, include_deleted=False)
            events_201 = self.db.get_events_by_number(RuleEngine.EVENT_ET, include_deleted=False)
        except Exception as e:
            logger.warning(f"SIRE一覧: イベント取得エラー: {e}")
            events_200 = []
            events_201 = []
        for event in events_200 + events_201:
            j = event.get("json_data")
            sire = _extract_sire_from_json(j)
            if sire:
                seen.add(sire)
        self._sire_names = sorted(seen)

    def _load_settings(self) -> None:
        """sire_list.json から種別を読み込み"""
        self._sire_type_by_name = {}
        if not self._settings_path.exists():
            return
        try:
            with open(self._settings_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, dict):
                return
            for sire_name, opts in raw.items():
                if not isinstance(opts, dict):
                    continue
                name = str(sire_name).strip()
                # 後方互換: f1_kurowa があれば f1 と kurowa に展開してから種別推定
                f1_kurowa = opts.get("f1_kurowa")
                if f1_kurowa is not None:
                    v = bool(f1_kurowa)
                    merged = {
                        "f1": v,
                        "kurowa": v,
                        "female": bool(opts.get("female", False)),
                        "sire_type": opts.get("sire_type"),
                    }
                    self._sire_type_by_name[name] = sire_opts_to_type(merged)
                else:
                    self._sire_type_by_name[name] = sire_opts_to_type(opts)
        except Exception as e:
            logger.warning(f"SIRE一覧: 設定読み込みエラー: {e}")

    def _save_settings(self) -> None:
        """現在の種別を sire_list.json に保存"""
        self.farm_path.mkdir(parents=True, exist_ok=True)
        out: Dict[str, Any] = {}
        for sire_name, var in self._type_display_vars.items():
            st = ja_label_to_sire_type(var.get())
            out[sire_name] = sire_type_to_stored_dict(st)
        try:
            with open(self._settings_path, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"SIRE一覧: 設定保存エラー: {e}")
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")

    def _create_widgets(self) -> None:
        _df = "Meiryo UI"
        bg = "#f5f5f5"

        # このウィンドウ用の ttk スタイル
        try:
            style = ttk.Style(self.window)
            style.configure("TLabel", font=(_df, 10))
            style.configure("TButton", font=(_df, 10))
            style.configure("TCombobox", font=(_df, 10))
        except tk.TclError:
            pass

        self.main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        self.main_container.grid_rowconfigure(2, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # ヘッダー（アイコン・タイトル・サブタイトル）
        header = tk.Frame(self.main_container, bg=bg)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        tk.Label(header, text="\U0001F402", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="SIRE一覧", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(title_frame, text="使用済みSIREの一覧と種別の設定", font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)

        # テーブル領域
        table_frame = ttk.Frame(self.main_container)
        table_frame.grid(row=2, column=0, sticky=tk.NSEW, pady=(0, 10))
        table_frame.grid_rowconfigure(1, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        COL_SIRE = 280
        COL_KIND = 260
        inner_minsize = COL_SIRE + COL_KIND + 32
        combo_values = [SIRE_TYPE_LABELS_JA[k] for k in SIRE_TYPE_ORDER]

        # ヘッダー行（スクロール外で固定）
        header_frame = tk.Frame(table_frame, bg=bg)
        header_frame.grid(row=0, column=0, sticky=tk.EW)
        header_frame.columnconfigure(0, minsize=COL_SIRE)
        header_frame.columnconfigure(1, minsize=COL_KIND)
        tk.Label(header_frame, text="SIRE名", font=(_df, 10, "bold"), bg=bg, fg="#37474f").grid(row=0, column=0, padx=4, pady=6, sticky=tk.W)
        tk.Label(header_frame, text="種別", font=(_df, 10, "bold"), bg=bg, fg="#37474f").grid(row=0, column=1, padx=4, pady=6, sticky=tk.W)

        # データ行用のCanvas＋スクロール
        can = tk.Canvas(table_frame, highlightthickness=0, bg=bg)
        scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=can.yview)
        inner = ttk.Frame(can)
        inner.bind("<Configure>", lambda e: can.configure(scrollregion=can.bbox("all")))
        self._inner_window_id = can.create_window((0, 0), window=inner, anchor="nw")
        can.configure(yscrollcommand=scroll.set)

        def _on_canvas_configure(e):
            w = max(e.width, inner_minsize)
            can.itemconfig(self._inner_window_id, width=w)

        can.bind("<Configure>", _on_canvas_configure)

        inner.columnconfigure(0, minsize=COL_SIRE)
        inner.columnconfigure(1, minsize=COL_KIND)

        if not self._sire_names:
            ttk.Label(inner, text="（AI/ETで使用されたSIREがありません）", font=(_df, 9)).grid(row=0, column=0, columnspan=2, padx=4, pady=8, sticky=tk.W)
        for i, sire_name in enumerate(self._sire_names):
            st = self._sire_type_by_name.get(sire_name, SIRE_TYPE_ORDER[0])
            disp = tk.StringVar(value=SIRE_TYPE_LABELS_JA[st])
            self._type_display_vars[sire_name] = disp
            display_name = sire_name if len(sire_name) <= 40 else sire_name[:40] + "..."
            ttk.Label(inner, text=display_name).grid(row=i, column=0, padx=4, pady=3, sticky=tk.W)
            cb = ttk.Combobox(
                inner,
                textvariable=disp,
                values=combo_values,
                state="readonly",
                width=22,
            )
            cb.grid(row=i, column=1, padx=4, pady=3, sticky=tk.W)
        def _on_mousewheel(event):
            can.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        def _on_mousewheel_linux(event, direction):
            can.yview_scroll(direction, "units")
            return "break"

        can.bind("<MouseWheel>", _on_mousewheel)
        inner.bind("<MouseWheel>", _on_mousewheel)
        can.bind("<Button-4>", lambda e: _on_mousewheel_linux(e, -1))
        can.bind("<Button-5>", lambda e: _on_mousewheel_linux(e, 1))
        inner.bind("<Button-4>", lambda e: _on_mousewheel_linux(e, -1))
        inner.bind("<Button-5>", lambda e: _on_mousewheel_linux(e, 1))

        can.grid(row=1, column=0, sticky=tk.NSEW)
        scroll.grid(row=1, column=1, sticky=tk.NS)

        # ボタン
        btn_frame = ttk.Frame(self.main_container)
        btn_frame.grid(row=3, column=0, pady=(0, 10))
        ttk.Button(btn_frame, text="閉じる", command=self.window.destroy, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="保存", command=self._on_save, width=12).pack(side=tk.LEFT, padx=5)

    def _on_save(self) -> None:
        """保存ボタン"""
        self._save_settings()
        if self.on_saved:
            try:
                self.on_saved()
            except Exception as e:
                logger.warning(f"SIRE一覧: 保存後コールバックでエラー: {e}")
        messagebox.showinfo("完了", "SIRE一覧の設定を保存しました。")

    def show(self) -> None:
        self.window.wait_window()
