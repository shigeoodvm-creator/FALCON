"""
FALCON2 - CowHistoryWindow（個体カードのイベント履歴専用）

個体カードと並べて表示する、縦長のイベント履歴。
parent が Toplevel のときは別ウィンドウ、Frame のときはそのフレーム内に埋め込む。
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List, Union
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.event_display import (
    format_insemination_event,
    format_calving_event,
    format_reproduction_check_event,
)
from modules.app_settings_manager import get_app_settings_manager
from ui.cow_card import CowCard


class CowHistoryWindow:
    """個体カード用のイベント履歴（別ウィンドウまたは埋め込み）"""

    def __init__(
        self,
        parent: Union[tk.Toplevel, tk.Frame],
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
            container,
            columns=columns,
            show="headings",
            height=35,
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
            container, orient=tk.VERTICAL, command=self.event_tree.yview
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

    def _on_delete_event(self, event_id: int):
        """削除メニュー選択時：CowCard の削除処理を呼び、履歴を再表示"""
        if hasattr(self.cow_card, "_on_delete_event"):
            self.cow_card._on_delete_event(event_id)
        if self.cow_auto_id:
            self._display_events()

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

    def _on_reproduction_filter_changed(self):
        """繁殖フィルターチェックボックスの変更時の処理"""
        if self.cow_auto_id:
            self._display_events()

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

