"""
FALCON2 - CowHistoryWindow（個体カードのイベント履歴専用）

個体カードと並べて表示する、縦長のイベント履歴。
parent が Toplevel のときは別ウィンドウ、Frame のときはそのフレーム内に埋め込む。
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Union, Tuple
import json
import logging
import warnings
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from matplotlib.transforms import blended_transform_factory

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.event_display import (
    format_insemination_event,
    format_calving_event,
    format_reproduction_check_event,
)
from modules.app_settings_manager import get_app_settings_manager
from ui.cow_card import CowCard


def _setup_japanese_font() -> None:
    """利用可能な日本語フォントを検出して matplotlib に設定"""
    preferred_fonts = [
        "MS Gothic",
        "MS PGothic",
        "Meiryo",
        "Yu Gothic",
        "MS UI Gothic",
        "Noto Sans CJK JP",
    ]

    try:
        available_fonts = [f.name for f in font_manager.fontManager.ttflist]
        selected_font = None
        for font_name in preferred_fonts:
            if font_name in available_fonts:
                selected_font = font_name
                break

        # 見つからない場合はデフォルトのまま（フォールバック）
        if selected_font:
            plt.rcParams["font.family"] = selected_font

        # マイナス記号の文字化け対策
        plt.rcParams["axes.unicode_minus"] = False

        # matplotlib のフォント警告を抑制（見た目崩れの回避が主目的）
        warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib.font_manager")
        warnings.filterwarnings("ignore", message=".*findfont.*", category=UserWarning)
    except Exception:
        # フォント設定が失敗しても致命的にしない
        return


_setup_japanese_font()


class CowHistoryWindow:
    """個体カード用のイベント履歴（別ウィンドウまたは埋め込み）"""

    def __init__(
        self,
        parent: Union[tk.Toplevel, tk.Frame, ttk.Frame],
        cow_card: CowCard,
        db_handler: DBHandler,
    ):
        """
        初期化

        Args:
            parent: 親。Toplevel のときは別ウィンドウ、Frame のときはその中に埋め込む
            cow_card: 既存の CowCard インスタンス
            db_handler: DBHandler インスタンス
        """
        self.parent = parent
        self.cow_card = cow_card
        self.db = db_handler
        # 親が Toplevel でなければ埋め込み（ttk.Frame 等）
        self._is_embedded = not isinstance(parent, tk.Toplevel)

        # CowCard から共有する情報（load_cow 時に都度更新する）
        self.event_dictionary: Dict[str, Dict[str, Any]] = getattr(
            cow_card, "event_dictionary", {}
        )
        self.technicians_dict: Dict[str, str] = getattr(
            cow_card, "technicians_dict", {}
        )
        self.insemination_types_dict: Dict[str, str] = getattr(
            cow_card, "insemination_types_dict", {}
        )

        try:
            self.app_settings = cow_card.app_settings
        except AttributeError:
            self.app_settings = get_app_settings_manager()

        self.cow_auto_id: Optional[int] = None
        self._configured_color_tags: set[str] = set()

        if self._is_embedded:
            # 埋め込み: parent をコンテナとして使う
            self.window = parent
            container = ttk.Frame(parent)
            container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        else:
            # 別ウィンドウ
            self.window = tk.Toplevel(parent)
            self.window.title("イベント履歴")
            self.window.configure(bg="#f5f5f5")
            self.window.protocol("WM_DELETE_WINDOW", self._on_close)
            container = ttk.Frame(self.window)
            container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ヘッダー（タイトル + 繁殖のみ表示）
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 5))

        font_size = self.app_settings.get_font_size()
        ttk.Label(
            header_frame,
            text="イベント履歴",
            font=("Meiryo UI", font_size, "bold"),
        ).pack(side=tk.LEFT)

        self.reproduction_filter_var = tk.BooleanVar(value=False)
        reproduction_checkbox = ttk.Checkbutton(
            header_frame,
            text="繁殖のみ表示",
            variable=self.reproduction_filter_var,
            command=self._on_reproduction_filter_changed,
        )
        reproduction_checkbox.pack(side=tk.LEFT, padx=(12, 0))

        # Tkinter のラッパーは、参照が切れるとGCで destroy されることがあるため、
        # ウィジェット参照を attributes として保持する
        self.content_paned = ttk.PanedWindow(container, orient=tk.VERTICAL)
        self.content_paned.pack(fill=tk.BOTH, expand=True)

        self.tree_frame = ttk.Frame(self.content_paned)
        self.graph_frame = ttk.Frame(self.content_paned)
        # グラフを右ペイン下半分として見やすくする
        #（イベント履歴は上側に縮める）
        self.content_paned.add(self.tree_frame, weight=1)
        self.content_paned.add(self.graph_frame, weight=1)

        # Treeview
        columns = ("date", "dim", "event", "note")
        event_font_family = "MS Gothic"
        try:
            style = ttk.Style()
            style.configure(
                "History.Treeview",
                font=(event_font_family, font_size),
            )
            style.configure(
                "History.Treeview.Heading",
                font=(event_font_family, font_size),
            )
        except tk.TclError:
            pass

        self.event_tree = ttk.Treeview(
            self.tree_frame,
            columns=columns,
            show="headings",
            # パン配置による高さ調整の邪魔にならないよう表示行数を抑える
            height=20,
            style="History.Treeview",
        )

        self.event_tree.heading("date", text="日付")
        self.event_tree.heading("dim", text="DIM")
        self.event_tree.heading("event", text="イベント")
        self.event_tree.heading("note", text="NOTE")

        char_width = max(8, int(font_size * 0.6))
        column_spacing = char_width
        date_width = 11 * char_width + column_spacing
        dim_width = 5 * char_width + column_spacing
        event_width = 14 * char_width + column_spacing
        note_width = max(280, 18 * char_width)

        self.event_tree.column(
            "date", width=date_width, stretch=False, anchor="w", minwidth=date_width
        )
        self.event_tree.column(
            "dim", width=dim_width, stretch=False, anchor="e", minwidth=dim_width
        )
        self.event_tree.column(
            "event",
            width=event_width,
            stretch=False,
            anchor="w",
            minwidth=event_width,
        )
        self.event_tree.column(
            "note", width=note_width, stretch=True, anchor="w", minwidth=note_width
        )

        scrollbar = ttk.Scrollbar(
            self.tree_frame, orient=tk.VERTICAL, command=self.event_tree.yview
        )
        self.event_tree.configure(yscrollcommand=scrollbar.set)

        self.event_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 右クリックで編集・削除メニュー
        # - ButtonRelease: Windows でもコンテキストメニューが確実に出るようリリースで表示
        # - Button-3/2 のプレス: 先に行を選択（リリース前の見た目用）
        self.event_tree.bind("<Button-3>", self._on_event_right_button_press)
        self.event_tree.bind("<Button-2>", self._on_event_right_button_press)
        self.event_tree.bind("<ButtonRelease-3>", self._on_event_context_menu)
        self.event_tree.bind("<ButtonRelease-2>", self._on_event_context_menu)
        # ダブルクリックで編集（右クリックと同じ可否）
        self.event_tree.bind("<Double-Button-1>", self._on_event_double_click)

        # 今産次の乳検推移（DIM横軸・左=乳量、右=リニアスコア）
        ttk.Label(
            self.graph_frame,
            text="乳量・リニアスコア推移（今産次）",
            font=("Meiryo UI", max(9, font_size - 1), "bold"),
        ).pack(anchor="w", pady=(0, 4))
        self.graph_canvas_frame = ttk.Frame(self.graph_frame)
        self.graph_canvas_frame.pack(fill=tk.BOTH, expand=True)
        # 右ペイン下側（グラフ領域）で読みやすい高さに調整
        self.figure = Figure(figsize=(6, 4.5), dpi=100)
        self.ax_milk = self.figure.add_subplot(111)
        self.ax_ls = None
        self.graph_canvas = FigureCanvasTkAgg(self.figure, self.graph_canvas_frame)
        self.graph_canvas.draw()
        self.graph_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    @staticmethod
    def _parse_tree_event_id(row_id: str) -> Optional[int]:
        """Treeview の iid（str(event.id)）からイベント ID を取得"""
        if not row_id:
            return None
        try:
            return int(str(row_id).strip())
        except (TypeError, ValueError):
            return None

    def _on_event_right_button_press(self, event):
        """右ボタン押下時に行を選択（メニューはリリースで表示）"""
        try:
            row_id = self.event_tree.identify_row(event.y)
            if not row_id:
                return
            self.event_tree.selection_set(row_id)
            self.event_tree.focus(row_id)
        except Exception as e:
            logging.exception("CowHistoryWindow._on_event_right_button_press: %s", e)

    def _on_event_double_click(self, event):
        """行のダブルクリックで編集（システム生成・deprecated のみブロック）"""
        try:
            row_id = self.event_tree.identify_row(event.y)
            if not row_id:
                return
            self.event_tree.selection_set(row_id)
            self.event_tree.focus(row_id)
            event_id = self._parse_tree_event_id(row_id)
            if event_id is None:
                return
            event_data = self.db.get_event_by_id(event_id)
            if event_data is None:
                return
            json_data_raw = event_data.get("json_data")
            json_data = json_data_raw if isinstance(json_data_raw, dict) else {}
            event_number = event_data.get("event_number")
            if isinstance(json_data, dict) and json_data.get("system_generated", False):
                event_str = str(event_number)
                if event_str in self.event_dictionary and self.event_dictionary[
                    event_str
                ].get("deprecated", False):
                    messagebox.showinfo(
                        "編集できません",
                        "このイベントはシステム生成のため、編集・削除はできません。",
                    )
                    return
            self._on_edit_event(event_id)
        except Exception as e:
            logging.exception("CowHistoryWindow._on_event_double_click: %s", e)

    def _on_event_context_menu(self, event):
        """イベント行の右クリック（リリース）でコンテキストメニュー（編集・削除）を表示"""
        try:
            row_id = self.event_tree.identify_row(event.y)
            if not row_id:
                return
            self.event_tree.selection_set(row_id)
            self.event_tree.focus(row_id)
            event_id = self._parse_tree_event_id(row_id)
            if event_id is None:
                return
            event_data = self.db.get_event_by_id(event_id)
            if event_data is None:
                return
            json_data_raw = event_data.get("json_data")
            json_data = json_data_raw if isinstance(json_data_raw, dict) else {}
            event_number = event_data.get("event_number")
            if isinstance(json_data, dict) and json_data.get("system_generated", False):
                event_str = str(event_number)
                if event_str in self.event_dictionary and self.event_dictionary[
                    event_str
                ].get("deprecated", False):
                    messagebox.showinfo(
                        "編集・削除できません",
                        "このイベントはシステム生成のため、編集・削除はできません。",
                    )
                    return
            menu = tk.Menu(self.event_tree, tearoff=0)
            menu.add_command(
                label="編集",
                command=lambda eid=event_id: self._on_edit_event(eid),
            )
            menu.add_command(
                label="削除",
                command=lambda eid=event_id: self._on_delete_event(eid),
            )
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        except Exception as e:
            logging.exception("CowHistoryWindow._on_event_context_menu: %s", e)

    def _on_edit_event(self, event_id: int):
        """編集メニュー選択時：CowCard の編集処理を呼び、閉じたら履歴を再表示"""
        if hasattr(self.cow_card, "_on_edit_event"):
            self.cow_card._on_edit_event(event_id)
        if self.cow_auto_id:
            self._display_events()
            self._draw_milk_ls_graph()

    def _on_delete_event(self, event_id: int):
        """削除メニュー選択時：CowCard の削除処理を呼び、履歴を再表示"""
        if hasattr(self.cow_card, "_on_delete_event"):
            self.cow_card._on_delete_event(event_id)
        if self.cow_auto_id:
            self._display_events()
            self._draw_milk_ls_graph()

    def _on_close(self):
        """ウィンドウクローズ時の処理（別ウィンドウの場合のみ）"""
        if not self._is_embedded and self.window.winfo_exists():
            self.window.destroy()

    def show(self):
        """ウィンドウを前面に表示（別ウィンドウの場合のみ）"""
        if not self._is_embedded and self.window.winfo_exists():
            self.window.deiconify()
            self.window.lift()
            self.window.focus_set()

    def load_cow(self, cow_auto_id: int):
        """
        指定された個体のイベント履歴を表示

        Args:
            cow_auto_id: 牛の auto_id
        """
        self.cow_auto_id = cow_auto_id

        # CowCard 側で辞書が更新されている可能性があるため、都度同期する
        self.event_dictionary = getattr(self.cow_card, "event_dictionary", {})
        self.technicians_dict = getattr(self.cow_card, "technicians_dict", {})
        self.insemination_types_dict = getattr(
            self.cow_card, "insemination_types_dict", {}
        )

        # タイトルにも個体IDを表示（別ウィンドウの場合のみ）
        if not self._is_embedded:
            try:
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                if cow:
                    cow_id = cow.get("cow_id", "")
                    jpn10 = cow.get("jpn10", "")
                    if jpn10:
                        self.window.title(f"イベント履歴 - {cow_id} ({jpn10})")
                    else:
                        self.window.title(f"イベント履歴 - {cow_id}")
            except Exception:
                logging.exception("CowHistoryWindow.load_cow: タイトル更新に失敗しました。")

        self._display_events()
        self._draw_milk_ls_graph()

    def _on_reproduction_filter_changed(self):
        """繁殖フィルターチェックボックスの変更時の処理"""
        if self.cow_auto_id:
            self._display_events()

    @staticmethod
    def _to_float(value: Any) -> Optional[float]:
        """値をfloatに変換（不可の場合はNone）"""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        try:
            text = str(value).strip()
            if not text:
                return None
            return float(text)
        except (TypeError, ValueError):
            return None

    def _ai_et_overlay_rows(
        self,
        events_asc: List[Dict[str, Any]],
        latest_calving_date: Optional[str],
        current_lact: Optional[Any],
    ) -> List[Tuple[int, bool]]:
        """
        今産次の AI/ET を (DIM, outcome が P か) のリストで返す。
        受胎は rule_engine が json_data.outcome に付与する 'P' で判定する。
        """
        rows: List[Tuple[int, bool]] = []
        for event in events_asc:
            en = event.get("event_number")
            if en not in (RuleEngine.EVENT_AI, RuleEngine.EVENT_ET):
                continue
            event_date = event.get("event_date")
            event_lact = event.get("event_lact")
            if current_lact is not None and event_lact is not None:
                include_event = int(event_lact) == int(current_lact)
            elif latest_calving_date and event_date:
                include_event = event_date >= latest_calving_date
            else:
                include_event = False
            if not include_event:
                continue

            dim_value = event.get("event_dim")
            if dim_value is None:
                dim_value = self._calculate_dim_at_event_date(events_asc, event_date or "")
            if dim_value is None:
                continue
            try:
                dim_int = int(dim_value)
            except (TypeError, ValueError):
                continue

            jd = event.get("json_data")
            if isinstance(jd, str):
                try:
                    jd = json.loads(jd)
                except Exception:
                    jd = {}
            if not isinstance(jd, dict):
                jd = {}
            outcome = (jd.get("outcome") or "").strip().upper()
            is_p = outcome == "P"
            rows.append((dim_int, is_p))
        rows.sort(key=lambda x: x[0])
        return rows

    def _plot_ai_et_markers(self, ai_rows: List[Tuple[int, bool]]) -> None:
        """
        AI/ET を縦線ではなく、グラフ上側（axes 座標で上三分の一の帯の中央付近）にプロットする。
        受胎（outcome=P）は ☆（*）、その他は △（^）。
        """
        if not ai_rows:
            return
        # 上三分の一: y_axes ∈ [2/3, 1] の中央 = 5/6
        y_axes = 5.0 / 6.0
        trans = blended_transform_factory(self.ax_milk.transData, self.ax_milk.transAxes)
        xs_p = [d for d, p in ai_rows if p]
        xs_o = [d for d, p in ai_rows if not p]
        if xs_p:
            self.ax_milk.plot(
                xs_p,
                [y_axes] * len(xs_p),
                transform=trans,
                linestyle="none",
                marker="*",
                markersize=12,
                markerfacecolor="#00acc1",
                markeredgewidth=0,
                clip_on=False,
                zorder=4,
            )
            # 受胎AI/ET（☆）には DIM を数値で併記する
            for d in xs_p:
                self.ax_milk.text(
                    d,
                    min(0.98, y_axes + 0.04),
                    str(int(d)),
                    transform=trans,
                    ha="center",
                    va="bottom",
                    fontsize=7,
                    color="#00838f",
                    clip_on=False,
                    zorder=5,
                )
        if xs_o:
            self.ax_milk.plot(
                xs_o,
                [y_axes] * len(xs_o),
                transform=trans,
                linestyle="none",
                marker="^",
                markersize=9,
                markerfacecolor="#5c6bc0",
                markeredgewidth=0,
                clip_on=False,
                zorder=4,
            )

    @staticmethod
    def _ai_et_legend_handles(ai_rows: List[Tuple[int, bool]]) -> Tuple[List[Line2D], List[str]]:
        """AI/ET 凡例用のハンドル（マーカーのみ）"""
        leg_h: List[Line2D] = []
        leg_l: List[str] = []
        if any(r[1] for r in ai_rows):
            leg_h.append(
                Line2D(
                    [0],
                    [0],
                    linestyle="none",
                    marker="*",
                    markersize=12,
                    markerfacecolor="#00acc1",
                    markeredgewidth=0,
                )
            )
            leg_l.append("受胎AI/ET")
        if any(not r[1] for r in ai_rows):
            leg_h.append(
                Line2D(
                    [0],
                    [0],
                    linestyle="none",
                    marker="^",
                    markersize=9,
                    markerfacecolor="#5c6bc0",
                    markeredgewidth=0,
                )
            )
            leg_l.append("AI/ET")
        return leg_h, leg_l

    def _draw_milk_ls_graph(self) -> None:
        """今産次の乳量・リニアスコア推移グラフを描画（欠損はスキップ）。AI/ET は縦線で重ねる。"""
        self.figure.clear()
        self.ax_milk = self.figure.add_subplot(111)
        self.ax_ls = self.ax_milk.twinx()

        if not self.cow_auto_id:
            self.ax_milk.text(
                0.5, 0.5, "個体が未選択です", ha="center", va="center", transform=self.ax_milk.transAxes
            )
            self.graph_canvas.draw_idle()
            return

        try:
            cow = self.db.get_cow_by_auto_id(self.cow_auto_id) or {}
            current_lact = cow.get("lact")
            events_all = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
            events_asc = sorted(
                events_all,
                key=lambda e: ((e.get("event_date") or ""), (e.get("id") or 0)),
            )

            latest_calving_date = None
            for event in events_asc:
                if event.get("event_number") == RuleEngine.EVENT_CALV:
                    event_date = event.get("event_date")
                    if event_date and (latest_calving_date is None or event_date > latest_calving_date):
                        latest_calving_date = event_date

            milk_points: List[Tuple[int, float]] = []
            ls_points: List[Tuple[int, float]] = []
            for event in events_asc:
                if event.get("event_number") != RuleEngine.EVENT_MILK_TEST:
                    continue

                event_date = event.get("event_date")
                event_lact = event.get("event_lact")
                include_event = True

                # 今産次のみ: event_lact が使える場合は厳密一致、欠損時は最新分娩以降で代替
                if current_lact is not None and event_lact is not None:
                    include_event = int(event_lact) == int(current_lact)
                elif latest_calving_date and event_date:
                    include_event = event_date >= latest_calving_date
                else:
                    include_event = False

                if not include_event:
                    continue

                dim_value = event.get("event_dim")
                if dim_value is None:
                    dim_value = self._calculate_dim_at_event_date(events_asc, event_date or "")
                if dim_value is None:
                    continue
                try:
                    dim_int = int(dim_value)
                except (TypeError, ValueError):
                    continue

                json_data = event.get("json_data")
                if not isinstance(json_data, dict):
                    json_data = {}

                milk = self._to_float(json_data.get("milk_yield"))
                ls = self._to_float(json_data.get("ls"))

                if milk is not None:
                    milk_points.append((dim_int, milk))
                if ls is not None:
                    ls_points.append((dim_int, ls))

            milk_points.sort(key=lambda x: x[0])
            ls_points.sort(key=lambda x: x[0])

            ai_rows = self._ai_et_overlay_rows(events_asc, latest_calving_date, current_lact)

            if not milk_points and not ls_points:
                self.ax_milk.set_xlabel("DIM")
                self.ax_milk.set_xlim(0, 400)
                self.ax_milk.set_xticks(list(range(0, 401, 50)))
                self.ax_milk.set_ylim(0, 60)
                self.ax_ls.set_ylim(0, 8)
                self.ax_milk.grid(True, color="#e0e0e0", linewidth=0.8)
                self.ax_milk.set_axisbelow(True)
                self.ax_milk.set_title("乳量・リニアスコア推移（今産次）", fontsize=10)
                if ai_rows:
                    self._plot_ai_et_markers(ai_rows)
                    leg_h, leg_l = self._ai_et_legend_handles(ai_rows)
                    self.ax_milk.legend(leg_h, leg_l, loc="upper right", fontsize=8)
                    self.ax_milk.text(
                        0.5,
                        0.5,
                        "今産次の乳検データがありません",
                        ha="center",
                        va="center",
                        transform=self.ax_milk.transAxes,
                        fontsize=10,
                        color="#757575",
                    )
                else:
                    self.ax_milk.text(
                        0.5,
                        0.5,
                        "今産次の乳検データがありません",
                        ha="center",
                        va="center",
                        transform=self.ax_milk.transAxes,
                    )
                self.figure.tight_layout()
                self.graph_canvas.draw_idle()
                return

            handles = []
            labels = []
            if milk_points:
                x_milk = [p[0] for p in milk_points]
                y_milk = [p[1] for p in milk_points]
                milk_line = self.ax_milk.plot(
                    x_milk,
                    y_milk,
                    color="#1e88e5",
                    marker="o",
                    linewidth=1.8,
                    markersize=4,
                    label="乳量(kg)",
                    zorder=3,
                )[0]
                self.ax_milk.set_ylabel("乳量(kg)", color="#1e88e5")
                self.ax_milk.tick_params(axis="y", colors="#1e88e5")
                handles.append(milk_line)
                labels.append("乳量(kg)")
            else:
                self.ax_milk.set_ylabel("乳量(kg)", color="#1e88e5")
                self.ax_milk.tick_params(axis="y", colors="#1e88e5")

            if ls_points:
                x_ls = [p[0] for p in ls_points]
                y_ls = [p[1] for p in ls_points]
                ls_line = self.ax_ls.plot(
                    x_ls,
                    y_ls,
                    color="#fb8c00",
                    marker="s",
                    linewidth=1.6,
                    markersize=4,
                    label="リニアスコア",
                    zorder=3,
                )[0]
                self.ax_ls.set_ylabel("リニアスコア", color="#fb8c00")
                self.ax_ls.tick_params(axis="y", colors="#fb8c00")
                handles.append(ls_line)
                labels.append("リニアスコア")
            else:
                self.ax_ls.set_ylabel("リニアスコア", color="#fb8c00")
                self.ax_ls.tick_params(axis="y", colors="#fb8c00")

            self.ax_milk.set_xlabel("DIM")
            # 横軸（DIM）を固定：個体ごとの差分でスケールが変わらないようにする
            self.ax_milk.set_xlim(0, 400)
            self.ax_milk.set_xticks(list(range(0, 401, 50)))
            # 縦軸を固定：乳量(左)=0-60kg, リニアスコア(右)=0-8
            self.ax_milk.set_ylim(0, 60)
            self.ax_ls.set_ylim(0, 8)
            self.ax_milk.grid(True, color="#e0e0e0", linewidth=0.8)
            self.ax_milk.set_axisbelow(True)
            self.ax_milk.set_title("乳量・リニアスコア推移（今産次）", fontsize=10)

            # 今産次の AI/ET：上三分の一付近に ☆（受胎=P）／△（その他）
            self._plot_ai_et_markers(ai_rows)
            ah, al = self._ai_et_legend_handles(ai_rows)
            handles.extend(ah)
            labels.extend(al)

            if handles:
                self.ax_milk.legend(handles, labels, loc="upper right", fontsize=8)

            self.figure.tight_layout()
            self.graph_canvas.draw_idle()
        except Exception as e:
            logging.exception("CowHistoryWindow._draw_milk_ls_graph: %s", e)
            self.figure.clear()
            self.ax_milk = self.figure.add_subplot(111)
            self.ax_milk.text(
                0.5, 0.5, "グラフの描画に失敗しました", ha="center", va="center", transform=self.ax_milk.transAxes
            )
            self.graph_canvas.draw_idle()

    # ======== イベント表示ロジック（CowCard とほぼ同じ） ========

    @staticmethod
    def _calculate_dim_at_event_date(
        events: List[Dict[str, Any]], event_date: str
    ) -> Optional[int]:
        """
        指定されたイベント日付時点でのDIM（分娩後日数）を計算
        """
        if not event_date:
            return None

        try:
            from datetime import datetime

            event_dt = datetime.strptime(event_date, "%Y-%m-%d")

            latest_calving_date = None
            for event in events:
                if event.get("event_number") == RuleEngine.EVENT_CALV:
                    calving_date = event.get("event_date")
                    if calving_date:
                        try:
                            calving_dt = datetime.strptime(calving_date, "%Y-%m-%d")
                            if calving_dt <= event_dt:
                                if (
                                    latest_calving_date is None
                                    or calving_dt > latest_calving_date
                                ):
                                    latest_calving_date = calving_dt
                        except ValueError:
                            continue

            if latest_calving_date:
                dim = (event_dt - latest_calving_date).days
                return dim if dim >= 0 else None

            return None
        except (ValueError, TypeError) as e:
            logging.warning(
                "[CowHistoryWindow] DIM計算エラー: event_date=%s, error=%s",
                event_date,
                e,
            )
            return None

    def _ensure_color_tag(self, color: str, event_number: int) -> str:
        """
        色に対応するタグを確保（まだ設定されていない場合は設定する）
        """
        tag_name = color
        if tag_name not in self._configured_color_tags:
            font_family = "MS Gothic"
            font_size = self.app_settings.get_font_size()
            if event_number == RuleEngine.EVENT_CALV:
                event_font = (font_family, font_size, "bold")
                self.event_tree.tag_configure(
                    tag_name, foreground=color, font=event_font
                )
            else:
                event_font = (font_family, font_size)
                self.event_tree.tag_configure(
                    tag_name, foreground=color, font=event_font
                )
            self._configured_color_tags.add(tag_name)
        return tag_name

    def _get_event_display_color(self, event_number: int) -> str:
        """
        イベントの表示色を決定
        """
        event_str = str(event_number)
        event_dict = self.event_dictionary.get(event_str, {})

        display_color = event_dict.get("display_color")
        if display_color:
            return display_color

        category = event_dict.get("category", "")

        if category == "CALVING":
            return "#0066cc"
        elif category == "PREGNANCY":
            outcome = event_dict.get("outcome", "")
            if outcome == "NEGATIVE":
                return "#cc0000"
            else:
                return "#008000"
        elif category == "REPRODUCTION":
            return "#000000"

        return "#000000"

    def _get_event_name(self, event_number: int) -> str:
        """
        イベント番号からイベント名を取得（CowCard と同様の短縮ロジック）
        """
        event_str = str(event_number)
        name_jp = None

        if event_str in self.event_dictionary:
            name_jp = self.event_dictionary[event_str].get("name_jp")

        if not name_jp:
            default_names = {
                RuleEngine.EVENT_AI: "AI",
                RuleEngine.EVENT_ET: "ET",
                RuleEngine.EVENT_CALV: "分娩",
                RuleEngine.EVENT_DRY: "乾乳",
                RuleEngine.EVENT_STOPR: "繁殖停止",
                RuleEngine.EVENT_SOLD: "売却",
                RuleEngine.EVENT_DEAD: "死亡・淘汰",
                RuleEngine.EVENT_PDN: "妊娠鑑定マイナス",
                RuleEngine.EVENT_PDP: "妊娠鑑定プラス",
                RuleEngine.EVENT_PDP2: "妊娠鑑定プラス（検診以外）",
                RuleEngine.EVENT_ABRT: "流産",
                RuleEngine.EVENT_PAGN: "PAGマイナス",
                RuleEngine.EVENT_PAGP: "PAGプラス",
                RuleEngine.EVENT_MILK_TEST: "乳検",
                RuleEngine.EVENT_MOVE: "群変更",
            }
            name_jp = default_names.get(event_number, f"イベント{event_number}")

        if event_number == RuleEngine.EVENT_PDN:
            return "妊鑑－"
        elif event_number in (RuleEngine.EVENT_PDP, RuleEngine.EVENT_PDP2):
            return "妊鑑＋"
        elif event_number == 300:
            if name_jp and "フレッシュチェック" in name_jp:
                return "フレチェック"
        elif event_number == 601:
            return "乳検"

        return name_jp or f"イベント{event_number}"

    def _display_events(self):
        """イベント履歴を表示（event_date DESC順）"""
        if not self.cow_auto_id:
            return

        try:
            # 既存のアイテムをクリア
            for item in self.event_tree.get_children():
                self.event_tree.delete(item)

            # イベントを取得（既にevent_date DESC順でソート済み）
            events = self.db.get_events_by_cow(
                self.cow_auto_id, include_deleted=False
            )

            # 繁殖関連イベントのみ表示フィルターが有効な場合はフィルタリング
            if self.reproduction_filter_var.get():
                events = [
                    event
                    for event in events
                    if event.get("event_number") is not None
                    and (
                        200 <= event.get("event_number") < 300
                        or 300 <= event.get("event_number") < 400
                    )
                ]

            displayed_count = 0

            for event in events:
                try:
                    event_date = event.get("event_date", "")
                    event_number = event.get("event_number")

                    if event_number is None:
                        logging.warning(
                            "[CowHistoryWindow] Skipping event with None event_number: event_id=%s",
                            event.get("id"),
                        )
                        continue

                    event_name = self._get_event_name(event_number)

                    json_data = event.get("json_data")
                    if json_data is None:
                        json_data = {}

                    note = ""

                    # AI/ET イベント
                    if event_number in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
                        outcome = json_data.get("outcome")
                        detail_text = format_insemination_event(
                            json_data,
                            self.technicians_dict,
                            self.insemination_types_dict,
                            outcome,
                        )
                        if detail_text:
                            note = detail_text.lstrip("　").strip() or detail_text.strip()

                    # 分娩イベント
                    elif event_number == RuleEngine.EVENT_CALV:
                        calv_def = self.event_dictionary.get(
                            str(RuleEngine.EVENT_CALV), {}
                        )
                        diff_labels = calv_def.get("calving_difficulty", {})
                        detail_text = format_calving_event(json_data, diff_labels)
                        note = detail_text or ""

                    # 乳検イベント
                    elif event_number == RuleEngine.EVENT_MILK_TEST:
                        milk_yield = json_data.get("milk_yield")
                        if milk_yield is not None:
                            note = f"乳量{milk_yield}kg"

                    # REPRODUCTION カテゴリ
                    if event_number not in [
                        RuleEngine.EVENT_AI,
                        RuleEngine.EVENT_ET,
                        RuleEngine.EVENT_CALV,
                        RuleEngine.EVENT_MILK_TEST,
                    ]:
                        event_dict = self.event_dictionary.get(str(event_number), {})
                        is_breeding = (
                            event_number in [300, 301, RuleEngine.EVENT_PDN]
                            or event_dict.get("category") == "REPRODUCTION"
                        )

                        if is_breeding:
                            detail_text = format_reproduction_check_event(json_data)
                            if detail_text:
                                if note:
                                    note = f"{detail_text} | {note}"
                                else:
                                    note = detail_text

                        if note:
                            original_note = note
                            while note.startswith("　") and len(note) > len("　"):
                                note = note[1:]
                            if not note or note.strip() == "":
                                note = original_note.strip()
                            else:
                                note = note.strip()

                    # その他のイベントで、REPRODUCTION カテゴリでもない場合は既存 note を使用
                    elif event_number not in [
                        RuleEngine.EVENT_AI,
                        RuleEngine.EVENT_ET,
                        RuleEngine.EVENT_CALV,
                        RuleEngine.EVENT_MILK_TEST,
                    ]:
                        event_dict = self.event_dictionary.get(str(event_number), {})
                        is_breeding = (
                            event_number in [300, 301, RuleEngine.EVENT_PDN]
                            or event_dict.get("category") == "REPRODUCTION"
                        )
                        if not is_breeding:
                            note = event.get("note", "") or ""
                        if note:
                            original_note = note
                            while note.startswith("　") and len(note) > len("　"):
                                note = note[1:]
                            if not note or note.strip() == "":
                                note = original_note.strip()
                            else:
                                note = note.strip()

                    display_color = self._get_event_display_color(event_number)
                    color_tag = self._ensure_color_tag(display_color, event_number)

                    event_id = event.get("id")
                    if event_id is None:
                        logging.warning(
                            "[CowHistoryWindow] Skipping event with None event_id: event_number=%s",
                            event_number,
                        )
                        continue

                    dim_value = self._calculate_dim_at_event_date(events, event_date)
                    dim_display = str(dim_value) if dim_value is not None else ""

                    self.event_tree.insert(
                        "",
                        "end",
                        iid=str(event_id),
                        values=(event_date, dim_display, event_name, note),
                        tags=(color_tag,),
                    )
                    displayed_count += 1

                except Exception as e:
                    logging.error(
                        "[CowHistoryWindow] Error processing event: %s, event_id=%s",
                        e,
                        event.get("id"),
                    )
                    import traceback

                    traceback.print_exc()
                    continue

            logging.debug(
                "[CowHistoryWindow] displayed events = %d", displayed_count
            )
        except Exception as e:
            import traceback

            logging.error(
                "ERROR: CowHistoryWindow._display_events で例外が発生しました: %s", e
            )
            traceback.print_exc()

