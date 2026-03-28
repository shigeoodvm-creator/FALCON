"""
FALCON2 - ゲノムレポートウィンドウ
生年月日期間・総合指標・選択ゲノム項目5つでレポートを生成し、HTMLでブラウザに出力。
DBに格納されたゲノムイベント（602）のjson_dataを利用。
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Callable
from datetime import datetime
import json
import logging
import tempfile
import webbrowser

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.genome_trait_descriptions import get_trait_description, get_keys_grouped_by_department
from modules.app_settings_manager import get_app_settings_manager

logger = logging.getLogger(__name__)

# 他ウィンドウと統一したテーマ（農場設定・辞書ウィンドウと同じ）
_THEME = {
    "bg": "#f5f5f5",
    "card_highlight": "#e0e7ef",
    "icon_fg": "#3949ab",
    "title_fg": "#263238",
    "subtitle_fg": "#607d8b",
    "desc_fg": "#78909c",
    "btn_primary_bg": "#3949ab",
    "btn_primary_fg": "#ffffff",
    "btn_secondary_bg": "#fafafa",
    "btn_secondary_fg": "#546e7a",
    "btn_secondary_bd": "#b0bec5",
    "instruction_bg": "#eceff1",
    "instruction_fg": "#546e7a",
}

# UI用フォント（統一してゴシック感を抑え、視認性を確保）
UI_FONT = ("Meiryo UI", 10)
UI_FONT_BOLD = ("Meiryo UI", 10, "bold")

# 総合指標（GDWP$・GNM$・GTPI）。GDPRは繁殖指標のため候補リストで選択可能。
COMPOSITE_INDEX_KEYS = ["GDWP$", "GNM$", "GTPI"]

# 下位25%を「良い」とする指標（低いほど良い）
LOWER_IS_BETTER_KEYS = {"GSCS", "GRFI", "GSCE", "GDCE", "GSSB", "GDSB", "GBVDV Results"}


def _parse_date(s: Optional[str]) -> Optional[datetime]:
    """YYYY-MM-DD または YYYY/MM/DD を datetime に変換"""
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    try:
        return datetime.strptime(s[:10].replace("/", "-"), "%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def _safe_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _get_genome_items_with_display(item_dict_path: Optional[Path]) -> Tuple[List[str], Dict[str, str]]:
    """
    項目辞書から総合指標以外のゲノム項目一覧と表示名マップを取得。
    Returns:
        (other_genome_keys_sorted, key -> display_name)
    """
    others = []
    display_names: Dict[str, str] = {}
    if item_dict_path and item_dict_path.exists():
        try:
            with open(item_dict_path, encoding="utf-8") as f:
                item_dict = json.load(f)
        except Exception as e:
            logger.warning(f"項目辞書読み込みエラー: {e}")
            item_dict = {}
    else:
        item_dict = {}

    for key, defn in item_dict.items():
        if not isinstance(defn, dict):
            continue
        if defn.get("category") != "GENOME":
            continue
        if key in COMPOSITE_INDEX_KEYS:
            continue
        if defn.get("data_type") == "float" or defn.get("type") == "source":
            others.append(key)
            display_names[key] = defn.get("display_name", key)

    # 用語解説の name_ja を優先
    for key in others:
        desc = get_trait_description(key)
        if desc.get("name_ja"):
            display_names[key] = desc["name_ja"]

    others.sort(key=lambda x: (display_names.get(x, x)))
    return others, display_names


def _get_sire_from_events(events: List[Dict]) -> Optional[str]:
    """AI/ETイベントのうち最新の json_data.sire を返す"""
    ai_et = [e for e in events if e.get("event_number") in (200, 201) and e.get("event_date")]
    if not ai_et:
        return None
    ai_et.sort(key=lambda e: (e.get("event_date", ""), e.get("id", 0)), reverse=True)
    j = ai_et[0].get("json_data") or {}
    if isinstance(j, str):
        try:
            j = json.loads(j)
        except Exception:
            j = {}
    sire = j.get("sire") or j.get("SIRE")
    if sire is not None:
        return str(sire).strip()
    return None


def load_genome_report_data(
    db: DBHandler,
    date_from: Optional[str],
    date_to: Optional[str],
    composite_key: str,
    additional_keys: List[str],
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, float]], Optional[str]]:
    """
    生年月日でフィルタし、各牛の最新ゲノムイベントから行データを組み立てる。
    """
    dt_from = _parse_date(date_from) if date_from else None
    dt_to = _parse_date(date_to) if date_to else None

    all_cows = db.get_all_cows()
    rows = []
    # 順位付けに使うのは composite_key のみ。表には3つの総合指数すべてを表示するため行に全総合指数を入れる。
    keys_for_row = list(COMPOSITE_INDEX_KEYS) + list(additional_keys)
    keys_for_stats = list(COMPOSITE_INDEX_KEYS) + list(additional_keys)

    for cow in all_cows:
        bthd = cow.get("bthd")
        if bthd:
            bthd_dt = _parse_date(bthd)
            if bthd_dt:
                if dt_from and bthd_dt < dt_from:
                    continue
                if dt_to and bthd_dt > dt_to:
                    continue
        else:
            if dt_from or dt_to:
                continue

        cow_auto_id = cow.get("auto_id")
        if not cow_auto_id:
            continue

        events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
        genome_ev = next((e for e in events if e.get("event_number") == RuleEngine.EVENT_GENOMIC), None)
        if not genome_ev:
            continue

        j = genome_ev.get("json_data") or {}
        if isinstance(j, str):
            try:
                j = json.loads(j)
            except Exception:
                j = {}

        row = {
            "cow_id": cow.get("cow_id", ""),
            "jpn10": cow.get("jpn10", ""),
            "sire": _get_sire_from_events(events) or "",
            "bthd": bthd or "",
        }
        has_any = False
        for k in keys_for_row:
            v = _safe_float(j.get(k))
            row[k] = v
            if v is not None:
                has_any = True
        if not has_any:
            continue
        if row.get(composite_key) is None:
            continue
        rows.append(row)

    if not rows:
        return [], {}, "対象期間内にゲノムデータがある個体がありません。"

    stats_by_column = {}
    for k in keys_for_stats:
        vals = [row[k] for row in rows if row.get(k) is not None]
        vals = [v for v in vals if v is not None]
        if not vals:
            stats_by_column[k] = {"mean": None, "min": None, "max": None, "p25": None, "p75": None}
            continue
        n = len(vals)
        s = sorted(vals)
        mean = sum(vals) / n
        stats_by_column[k] = {
            "mean": mean,
            "min": min(vals),
            "max": max(vals),
            "p25": s[max(0, (n - 1) * 25 // 100)],
            "p75": s[min(n - 1, (n - 1) * 75 // 100)],
        }

    return rows, stats_by_column, None


class GenomeReportWindow:
    """ゲノムレポート設定ウィンドウ（HTML出力）"""

    MAX_ADDITIONAL = 5

    def __init__(
        self,
        parent: tk.Tk,
        db: DBHandler,
        item_dict_path: Optional[Path] = None,
        farm_name: str = "",
        on_close: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.db = db
        self.item_dict_path = Path(item_dict_path) if item_dict_path else None
        self.farm_name = farm_name or ""
        self._on_close_callback = on_close  # 閉じるときに呼ぶコールバック（メインウィンドウの参照クリア用）

        self.window = tk.Toplevel(parent)
        self.window.title("ゲノムレポート")
        self.window.geometry("1000x640")
        self.window.minsize(820, 520)
        self.window.configure(bg=_THEME["bg"])
        self._apply_font_style()

        self.composite_var = tk.StringVar(value="GDWP$")
        self.date_from_var = tk.StringVar()
        self.date_to_var = tk.StringVar()
        self.show_trend_var = tk.BooleanVar(value=True)
        self.show_avg_line_var = tk.BooleanVar(value=True)
        self.search_var = tk.StringVar()

        other_keys, self.item_display_names = _get_genome_items_with_display(self.item_dict_path)
        self.other_genome_keys = other_keys
        self._selected_list: List[str] = []  # 選択した項目（最大5つ、順序保持）

        self._load_saved_state()
        self._create_widgets(other_keys)
        self.window.protocol("WM_DELETE_WINDOW", self._on_window_close)

    def _apply_font_style(self):
        """ウィンドウ内のフォントを Meiryo UI で統一する。"""
        try:
            style = ttk.Style(self.window)
            style.configure(".", font=UI_FONT)
            style.configure("TLabel", font=UI_FONT)
            style.configure("TButton", font=UI_FONT)
            style.configure("TCheckbutton", font=UI_FONT)
            style.configure("TRadiobutton", font=UI_FONT)
            style.configure("TLabelframe.Label", font=UI_FONT)
            style.configure("TLabelframe", background=_THEME["bg"])
        except Exception:
            pass

    def _build_header(self):
        """他ウィンドウと統一したヘッダー（アイコン・タイトル・サブタイトル）"""
        header = tk.Frame(self.window, bg=_THEME["bg"], pady=16, padx=24)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="\U0001f4ca", font=("Meiryo UI", 22),
            bg=_THEME["bg"], fg=_THEME["icon_fg"]
        ).pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=_THEME["bg"])
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(
            title_frame, text="ゲノムレポート",
            font=("Meiryo UI", 16, "bold"), bg=_THEME["bg"], fg=_THEME["title_fg"]
        ).pack(anchor=tk.W)
        tk.Label(
            title_frame, text="生年月日期間・総合指標・ゲノム項目を選んでHTMLレポートを生成",
            font=("Meiryo UI", 10), bg=_THEME["bg"], fg=_THEME["subtitle_fg"]
        ).pack(anchor=tk.W)

    def _load_saved_state(self):
        """前回の選択・設定をアプリ設定から復元する。"""
        settings = get_app_settings_manager()
        saved_keys = settings.get("genome_report_selected_keys", [])
        if isinstance(saved_keys, list):
            valid = [k for k in saved_keys if k in self.other_genome_keys][: self.MAX_ADDITIONAL]
            self._selected_list[:] = valid
        comp = settings.get("genome_report_composite_index")
        if comp and comp in COMPOSITE_INDEX_KEYS:
            self.composite_var.set(comp)
        date_from = settings.get("genome_report_date_from")
        if date_from is not None and isinstance(date_from, str):
            self.date_from_var.set(date_from)
        date_to = settings.get("genome_report_date_to")
        if date_to is not None and isinstance(date_to, str):
            self.date_to_var.set(date_to)
        if settings.get("genome_report_show_trend") is not None:
            self.show_trend_var.set(bool(settings.get("genome_report_show_trend")))
        if settings.get("genome_report_show_avg_line") is not None:
            self.show_avg_line_var.set(bool(settings.get("genome_report_show_avg_line")))

    def _save_state(self):
        """現在の選択・設定をアプリ設定に保存する（次回ウィンドウで復元）。"""
        settings = get_app_settings_manager()
        settings.set("genome_report_selected_keys", list(self._selected_list))
        settings.set("genome_report_composite_index", self.composite_var.get())
        settings.set("genome_report_date_from", self.date_from_var.get().strip() or None)
        settings.set("genome_report_date_to", self.date_to_var.get().strip() or None)
        settings.set("genome_report_show_trend", self.show_trend_var.get())
        settings.set("genome_report_show_avg_line", self.show_avg_line_var.get())

    def _on_window_close(self):
        """ウィンドウ閉じる際に設定を保存してから破棄する。"""
        self._save_state()
        if self._on_close_callback:
            try:
                self._on_close_callback()
            except Exception:
                pass
        self.window.destroy()

    def _create_widgets(self, other_keys: List[str]):
        self._build_header()

        main = ttk.Frame(self.window, padding=0)
        main.pack(fill=tk.BOTH, expand=True)

        # 設定カード（期間・総合指数・散布図を1ブロックに）
        settings_card = tk.Frame(main, bg=_THEME["bg"], padx=24, pady=0)
        settings_card.pack(fill=tk.X)
        card_inner = tk.Frame(
            settings_card,
            bg=_THEME["card_highlight"],
            highlightbackground=_THEME["card_highlight"],
            highlightthickness=1,
            padx=16,
            pady=12,
        )
        card_inner.pack(fill=tk.X)

        top_row = tk.Frame(card_inner, bg=_THEME["card_highlight"])
        top_row.pack(fill=tk.X)
        top_row.columnconfigure(0, weight=3)
        top_row.columnconfigure(1, weight=2)
        top_row.columnconfigure(2, weight=2)

        period_f = ttk.LabelFrame(top_row, text="生年月日で期間を設定", padding=6)
        period_f.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 10))
        pr = ttk.Frame(period_f)
        pr.pack(fill=tk.X)
        ttk.Label(pr, text="開始").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(pr, textvariable=self.date_from_var, width=11).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(pr, text="終了").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(pr, textvariable=self.date_to_var, width=11).pack(side=tk.LEFT)
        ttk.Label(pr, text="(YYYY-MM-DD)", font=UI_FONT).pack(side=tk.LEFT, padx=(6, 0))

        comp_f = ttk.LabelFrame(top_row, text="総合指数（1つ選択）", padding=6)
        comp_f.grid(row=0, column=1, sticky=tk.NSEW, padx=(0, 10))
        for key in COMPOSITE_INDEX_KEYS:
            ttk.Radiobutton(
                comp_f,
                text=key,
                variable=self.composite_var,
                value=key,
            ).pack(anchor=tk.W)

        opt_f = ttk.LabelFrame(top_row, text="散布図", padding=6)
        opt_f.grid(row=0, column=2, sticky=tk.NSEW)
        ttk.Checkbutton(opt_f, text="近似曲線", variable=self.show_trend_var).pack(anchor=tk.W)
        ttk.Checkbutton(opt_f, text="平均線", variable=self.show_avg_line_var).pack(anchor=tk.W)

        # 操作説明（控えめな色でテーマ統一）
        instruction_bar = tk.Frame(main, bg=_THEME["instruction_bg"], padx=24, pady=8)
        instruction_bar.pack(fill=tk.X, pady=(12, 8))
        tk.Label(
            instruction_bar,
            text="左の候補からダブルクリックで選択に追加（最大5つ）。右の項目をダブルクリックで解除。",
            font=UI_FONT,
            bg=_THEME["instruction_bg"],
            fg=_THEME["instruction_fg"],
        ).pack(anchor=tk.W)

        paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        def _set_initial_sash():
            try:
                paned.update_idletasks()
                w = paned.winfo_width()
                if w > 100:
                    paned.sashpos(0, w // 2)
            except Exception:
                pass

        self.window.after(150, _set_initial_sash)

        # 左: 候補（検索 + 部門別一覧、各項目にⓘで説明表示）
        left = ttk.Frame(paned, padding=(0, 4))
        paned.add(left, weight=1)

        ttk.Label(left, text="候補（ダブルクリックで選択に追加）", font=UI_FONT_BOLD).pack(anchor=tk.W, pady=(0, 4))
        search_f = ttk.Frame(left)
        search_f.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(search_f, text="検索:").pack(side=tk.LEFT, padx=(0, 6))
        search_entry = ttk.Entry(search_f, textvariable=self.search_var, width=20, font=UI_FONT)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.search_var.trace_add("write", lambda *a: self._filter_and_refresh())

        list_container = ttk.Frame(left)
        list_container.pack(fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(list_container)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.card_canvas = tk.Canvas(list_container, yscrollcommand=scroll.set, highlightthickness=0, bg=_THEME["bg"])
        self.card_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=self.card_canvas.yview)

        self.card_inner = ttk.Frame(self.card_canvas)
        self.card_frame_id = self.card_canvas.create_window((0, 0), window=self.card_inner, anchor=tk.NW)
        self.card_inner.bind("<Configure>", self._on_card_inner_configure)
        self.card_canvas.bind("<Configure>", self._on_canvas_configure)
        list_container.bind("<MouseWheel>", self._on_mousewheel)
        list_container.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
        list_container.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))
        self.card_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.card_inner.bind("<MouseWheel>", self._on_mousewheel)
        self._bind_mousewheel_linux()

        self._card_widgets: List[Tuple[str, tk.Frame, tk.Widget]] = []  # (key, frame, info_btn)
        self._department_headers = []
        self._keys_by_department = []
        grouped = get_keys_grouped_by_department(other_keys)
        self._build_candidate_list(grouped)
        self._bind_mousewheel_to_cards()

        # 右: 選択した項目（5つ）— 枠線付き・空き時は中央にプレースホルダー表示
        right = ttk.Frame(paned, padding=(16, 4))
        paned.add(right, weight=1)

        self.selected_title_label = ttk.Label(right, text="選択した項目（0 / 5つ）", font=UI_FONT_BOLD)
        self.selected_title_label.pack(anchor=tk.W, pady=(0, 6))

        # 選択エリアをカード風に（テーマ色で統一）
        selected_border = tk.Frame(right, bg=_THEME["card_highlight"], padx=2, pady=2, highlightbackground=_THEME["btn_secondary_bd"], highlightthickness=1)
        selected_border.pack(fill=tk.BOTH, expand=True)
        self.selected_container = tk.Frame(selected_border, bg=_THEME["bg"], padx=12, pady=12)
        self.selected_container.pack(fill=tk.BOTH, expand=True)
        self._selected_labels = []

        self._refresh_selected_panel()

        # フッター: メイン操作と補助を1行に（他ウィンドウと統一）
        footer = tk.Frame(main, bg=_THEME["bg"], pady=16)
        footer.pack(fill=tk.X)
        btn_primary = tk.Button(
            footer, text="レポートをHTMLで開く", font=UI_FONT,
            bg=_THEME["btn_primary_bg"], fg=_THEME["btn_primary_fg"],
            activebackground="#303f9f", activeforeground="#ffffff",
            relief=tk.FLAT, padx=20, pady=10, cursor="hand2",
            command=self._run_report,
        )
        btn_primary.pack(side=tk.LEFT, padx=(24, 8))
        tk.Button(
            footer, text="用語解説を一覧で開く", font=UI_FONT,
            bg=_THEME["btn_secondary_bg"], fg=_THEME["btn_secondary_fg"],
            activebackground="#eceff1", relief=tk.FLAT, padx=16, pady=10,
            highlightbackground=_THEME["btn_secondary_bd"], highlightthickness=1,
            cursor="hand2", command=self._open_glossary_window,
        ).pack(side=tk.LEFT)

    def _build_candidate_list(self, grouped: List[Tuple[str, List[str]]]):
        """左パネル: 部門別の候補一覧。各項目はラベル+ⓘ。ダブルクリックで選択に追加。"""
        for w in self.card_inner.winfo_children():
            w.destroy()
        self._card_widgets.clear()
        self._department_headers = []
        self._keys_by_department = list(grouped)

        for dept_idx, (dept_name, keys) in enumerate(grouped):
            header_frame = ttk.Frame(self.card_inner, padding=(0, 6 if dept_idx > 0 else 0, 0, 2))
            header_frame.pack(fill=tk.X)
            ttk.Label(header_frame, text=f"■ {dept_name}", font=UI_FONT_BOLD).pack(anchor=tk.W)
            self._department_headers.append((dept_name, header_frame))

            for key in keys:
                display = self.item_display_names.get(key, key)
                frame = ttk.Frame(self.card_inner, padding=2)
                frame.pack(fill=tk.X, pady=1)
                lbl = ttk.Label(frame, text=f"  {display}", cursor="hand2")
                lbl.pack(side=tk.LEFT, anchor=tk.W)
                info_btn = ttk.Button(frame, text="ⓘ", width=2, command=lambda k=key: self._show_item_description_popup(k))
                info_btn.pack(side=tk.RIGHT, padx=(4, 0))
                lbl.bind("<Double-1>", lambda e, k=key: self._on_candidate_double_click(k))
                frame.bind("<Double-1>", lambda e, k=key: self._on_candidate_double_click(k))
                self._card_widgets.append((key, frame, info_btn))

    def _on_candidate_double_click(self, key: str):
        """候補をダブルクリック → 選択に追加（最大5つ）"""
        if key in self._selected_list:
            return
        if len(self._selected_list) >= self.MAX_ADDITIONAL:
            messagebox.showinfo("ゲノムレポート", f"選択は最大{self.MAX_ADDITIONAL}つまでです。右の項目をダブルクリックで解除できます。")
            return
        self._selected_list.append(key)
        self._refresh_selected_panel()
        self._filter_and_refresh()

    def _refresh_selected_panel(self):
        """右パネル「選択した項目」を描画し直す。空き時は中央にプレースホルダーを表示。"""
        n = len(self._selected_list)
        self.selected_title_label.config(text=f"選択した項目（{n} / {self.MAX_ADDITIONAL}つ）")
        for w in self.selected_container.winfo_children():
            w.destroy()
        self._selected_labels.clear()

        if self._selected_list:
            for idx, key in enumerate(self._selected_list):
                display = self.item_display_names.get(key, key)
                frame = ttk.Frame(self.selected_container, padding=4)
                frame.pack(fill=tk.X, pady=2)
                ttk.Label(frame, text=f"{idx + 1}. {display}", font=UI_FONT).pack(side=tk.LEFT, anchor=tk.W)
                frame.bind("<Double-1>", lambda e, k=key: self._on_selected_double_click(k))
                for c in frame.winfo_children():
                    c.bind("<Double-1>", lambda e, k=key: self._on_selected_double_click(k))
                self._selected_labels.append((key, frame))
        else:
            # 空き時: 中央にアイコン＋案内テキスト（テーマ色で統一）
            placeholder = tk.Frame(self.selected_container, bg=_THEME["bg"])
            placeholder.pack(fill=tk.BOTH, expand=True)
            inner = tk.Frame(placeholder, bg=_THEME["bg"])
            inner.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            tk.Label(inner, text="\U0001f4cb", font=("Meiryo UI", 28), bg=_THEME["bg"], fg=_THEME["icon_fg"]).pack(pady=(0, 8))
            tk.Label(
                inner,
                text="ダブルクリックでここに追加",
                font=UI_FONT,
                bg=_THEME["bg"],
                fg=_THEME["desc_fg"],
            ).pack()

    def _on_selected_double_click(self, key: str):
        """選択項目をダブルクリック → 解除"""
        if key in self._selected_list:
            self._selected_list.remove(key)
            self._refresh_selected_panel()
            self._filter_and_refresh()

    def _show_item_description_popup(self, key: str):
        """項目横のⓘクリックで説明をポップアップ表示（メリハリ付き）"""
        desc = get_trait_description(key)
        name_ja = desc.get("name_ja", key)
        text_desc = (desc.get("desc") or "").strip()
        tip = (desc.get("tip") or "").strip()
        direction = desc.get("direction", "higher")
        d_map = {"higher": "高いほど良い", "lower": "低いほど良い", "neutral": "色分けなし"}
        dir_ja = d_map.get(direction, "")

        pop = tk.Toplevel(self.window)
        pop.title(f"説明 — {name_ja}")
        pop.geometry("400x260")
        pop.minsize(340, 200)
        f = ttk.Frame(pop, padding=12)
        f.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(f, wrap=tk.WORD, font=UI_FONT, padx=10, pady=10, state=tk.NORMAL, spacing1=4, spacing2=2)
        txt.pack(fill=tk.BOTH, expand=True)

        # タグでメリハリを付ける
        txt.tag_configure("title", font=("Meiryo UI", 12, "bold"), foreground="#1565c0", spacing3=6)
        txt.tag_configure("body", foreground="#37474f", spacing3=4)
        txt.tag_configure("tip", foreground="#2e7d32", spacing3=4)
        txt.tag_configure("meta_label", foreground="#546e7a", font=("Meiryo UI", 9))
        txt.tag_configure("meta_value", foreground="#1565c0", font=("Meiryo UI", 9, "bold"))

        txt.insert(tk.END, f"{name_ja}\n\n", "title")
        if text_desc:
            txt.insert(tk.END, f"{text_desc}\n\n", "body")
        if tip:
            txt.insert(tk.END, f"→ {tip}\n\n", "tip")
        txt.insert(tk.END, "レポートでの色分け: ", "meta_label")
        txt.insert(tk.END, f"{dir_ja}\n", "meta_value")

        txt.config(state=tk.DISABLED)
        ttk.Button(f, text="閉じる", command=pop.destroy).pack(pady=(10, 0))

    def _bind_mousewheel_to_cards(self):
        """リスト内の各カード・部門見出しにマウスホイールをバインド"""
        for _dept, header_frame in self._department_headers:
            header_frame.bind("<MouseWheel>", self._on_mousewheel)
            header_frame.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
            header_frame.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))
            for child in header_frame.winfo_children():
                child.bind("<MouseWheel>", self._on_mousewheel)
                child.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
                child.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))
        for _key, frame, info_btn in self._card_widgets:
            frame.bind("<MouseWheel>", self._on_mousewheel)
            frame.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
            frame.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))
            info_btn.bind("<MouseWheel>", self._on_mousewheel)
            info_btn.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
            info_btn.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))
            for child in frame.winfo_children():
                child.bind("<MouseWheel>", self._on_mousewheel)
                child.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
                child.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))

    def _open_glossary_window(self):
        """用語解説を別ウィンドウで開く（部門ごと・フォント統一）"""
        win = tk.Toplevel(self.window)
        win.title("ゲノム項目 用語解説")
        win.geometry("520x520")
        win.minsize(400, 320)
        f = ttk.Frame(win, padding=12)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="各ゲノム指標の説明です。部門ごとにまとめて表示しています。", font=UI_FONT).pack(anchor=tk.W, pady=(0, 8))
        txt = tk.Text(f, wrap=tk.WORD, font=UI_FONT, padx=8, pady=8, state=tk.NORMAL)
        scr = ttk.Scrollbar(f)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scr.pack(side=tk.RIGHT, fill=tk.Y)
        txt.config(yscrollcommand=scr.set)
        scr.config(command=txt.yview)
        d_map = {"higher": "高いほど良い", "lower": "低いほど良い", "neutral": "色分けなし"}
        grouped = get_keys_grouped_by_department(list(COMPOSITE_INDEX_KEYS) + list(self.other_genome_keys))
        for dept_name, keys in grouped:
            txt.insert(tk.END, f"\n■ {dept_name}\n", "dept")
            for key in keys:
                desc = get_trait_description(key)
                name_ja = desc.get("name_ja", key)
                text_desc = (desc.get("desc") or "").strip()
                tip = (desc.get("tip") or "").strip()
                direction = desc.get("direction", "higher")
                dir_ja = d_map.get(direction, "")
                txt.insert(tk.END, f"  【{name_ja}】\n", "title")
                if text_desc:
                    txt.insert(tk.END, f"    {text_desc}\n\n", "body")
                if tip:
                    txt.insert(tk.END, f"    → {tip}\n\n", "tip")
                txt.insert(tk.END, f"    レポートでの色分け: {dir_ja}\n\n", "meta")
        txt.config(state=tk.DISABLED)
        txt.tag_configure("dept", font=UI_FONT_BOLD)
        txt.tag_configure("title", font=UI_FONT_BOLD)
        ttk.Button(f, text="閉じる", command=win.destroy).pack(pady=(12, 0))

    def _filter_and_refresh(self):
        q = (self.search_var.get() or "").strip().lower()
        visible_keys = set()
        for key, frame, _ in self._card_widgets:
            if key in self._selected_list:
                frame.pack_forget()
                continue
            display = (self.item_display_names.get(key, key) + key).lower()
            if not q or q in display:
                frame.pack(fill=tk.X, pady=1)
                visible_keys.add(key)
            else:
                frame.pack_forget()
        for dept_name, header_frame in self._department_headers:
            keys_in_dept = next((ks for d, ks in self._keys_by_department if d == dept_name), [])
            if any(k in visible_keys for k in keys_in_dept):
                header_frame.pack(fill=tk.X)
            else:
                header_frame.pack_forget()

    def _on_card_inner_configure(self, ev):
        self.card_canvas.configure(scrollregion=self.card_canvas.bbox("all"))

    def _on_canvas_configure(self, ev):
        self.card_canvas.itemconfig(self.card_frame_id, width=ev.width)

    def _on_mousewheel(self, event):
        """マウスホイールでゲノム項目リストをスクロール"""
        if event.delta:
            # Windows: delta は 120 または -120
            self.card_canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _bind_mousewheel_linux(self):
        """Linux では Button-4/5 でホイールスクロール"""
        for w in (self.card_canvas, self.card_inner):
            w.bind("<Button-4>", lambda e: self.card_canvas.yview_scroll(-3, "units"))
            w.bind("<Button-5>", lambda e: self.card_canvas.yview_scroll(3, "units"))

    def _get_selected_additional(self) -> List[str]:
        return list(self._selected_list)

    def _run_report(self):
        additional = self._get_selected_additional()
        if len(additional) < 1:
            messagebox.showwarning("ゲノムレポート", "他のゲノム項目を1つ以上選択してください。（現在 0 つ）")
            return
        if len(additional) > self.MAX_ADDITIONAL:
            messagebox.showwarning("ゲノムレポート", f"選択は最大{self.MAX_ADDITIONAL}つまでです。（現在 {len(additional)} つ）")
            return
        composite = self.composite_var.get()
        if composite not in COMPOSITE_INDEX_KEYS:
            messagebox.showwarning("ゲノムレポート", "総合指数を選択してください。")
            return

        rows, stats, err = load_genome_report_data(
            self.db,
            self.date_from_var.get().strip() or None,
            self.date_to_var.get().strip() or None,
            composite,
            additional,
        )
        if err:
            messagebox.showerror("ゲノムレポート", err)
            return

        # 表示名: 用語解説の name_ja を優先、なければ item_display_names
        trait_display_names: Dict[str, str] = {}
        for k in [composite] + additional + list(COMPOSITE_INDEX_KEYS):
            trait_display_names[k] = get_trait_description(k).get("name_ja") or self.item_display_names.get(k, k)

        try:
            from modules.genome_report_html import build_genome_report_html
            html_content = build_genome_report_html(
                farm_name=self.farm_name,
                date_from=self.date_from_var.get().strip() or None,
                date_to=self.date_to_var.get().strip() or None,
                rows=rows,
                stats=stats,
                composite_key=composite,
                additional_keys=additional,
                show_trend=self.show_trend_var.get(),
                show_avg_line=self.show_avg_line_var.get(),
                trait_display_names=trait_display_names,
            )
        except Exception as e:
            logger.error(f"HTML生成エラー: {e}", exc_info=True)
            messagebox.showerror("ゲノムレポート", f"HTMLの生成に失敗しました:\n{e}")
            return

        try:
            from modules.report_cow_bridge import DEFAULT_PORT as _report_open_cow_port
        except ImportError:
            _report_open_cow_port = 51985
        html_content = html_content.replace(
            "</head>",
            "<script>var FALCON_OPEN_COW_PORT = " + str(_report_open_cow_port) + ";</script>\n</head>",
        )

        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", suffix=".html", delete=False) as f:
                f.write(html_content)
                path = f.name
            webbrowser.open(f"file://{path}")
            self._save_state()
        except Exception as e:
            logger.error(f"HTML保存・ブラウザ起動エラー: {e}", exc_info=True)
            messagebox.showerror("ゲノムレポート", f"ファイルの保存またはブラウザの起動に失敗しました:\n{e}")

    def show(self):
        self.window.lift()
        self.window.focus_force()
