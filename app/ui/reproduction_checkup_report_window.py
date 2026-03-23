"""
FALCON2 - 繁殖検診レポートウインドウ
年間必要妊娠頭数の計算を表示する
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager


def _annual_required_pregnancies(
    parous_count: int,
    target_calving_interval_days: float,
    replacement_rate_percent: float,
    abortion_rate_percent: float,
) -> float:
    """
    年間必要妊娠頭数を計算する。

    年間必要妊娠頭数 ＝ 経産牛頭数 × (1 - 更新率) × 365 / 目標分娩間隔 × (1.0 + 流産率)

    Args:
        parous_count: 想定経産牛頭数
        target_calving_interval_days: 目標分娩間隔（日）
        replacement_rate_percent: 更新率（%）（例: 30）
        abortion_rate_percent: 流産率（%）（例: 10）

    Returns:
        年間必要妊娠頭数（小数点付き）
    """
    if target_calving_interval_days <= 0:
        return 0.0
    replacement = replacement_rate_percent / 100.0
    abortion = abortion_rate_percent / 100.0
    return (
        parous_count
        * (1.0 - replacement)
        * 365.0
        / target_calving_interval_days
        * (1.0 + abortion)
    )


class ReproductionCheckupReportWindow:
    """繁殖検診レポートウインドウ（年間必要妊娠頭数）"""

    # 農場設定ウィンドウなどと統一したデザイン定数
    _FONT = "Meiryo UI"
    _BG = "#f5f5f5"
    _CARD_HIGHLIGHT = "#e0e7ef"
    _ICON_FG = "#3949ab"
    _TITLE_FG = "#263238"
    _SUBTITLE_FG = "#607d8b"
    _DESC_FG = "#78909c"
    _BTN_PRIMARY_BG = "#3949ab"
    _BTN_PRIMARY_FG = "#ffffff"
    _BTN_SECONDARY_BG = "#fafafa"
    _BTN_SECONDARY_FG = "#546e7a"
    _BTN_SECONDARY_BD = "#b0bec5"

    def __init__(
        self,
        parent: tk.Widget,
        db: DBHandler,
        rule_engine: RuleEngine,
        farm_path: Path,
    ):
        self.parent = parent
        self.db = db
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        self.settings_manager = SettingsManager(self.farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("必要妊娠シミュレーション（年間必要妊娠頭数）")
        self.window.geometry("540x520")
        self.window.configure(bg=self._BG)
        self.window.minsize(460, 420)

        self._parous_var = tk.StringVar()
        self._result_var = tk.StringVar()
        self._formula_var = tk.StringVar()
        self._interval_var = tk.StringVar()
        self._replacement_var = tk.StringVar()
        self._abortion_var = tk.StringVar()

        self._create_widgets()
        self._refresh()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _get_current_parous_count(self) -> int:
        """現状の経産牛頭数（現存・LACT>=1）を取得"""
        all_cows = self.db.get_all_cows()
        existing = [
            c for c in all_cows
            if not self._is_cow_disposed(c.get("auto_id"))
        ]
        lactated = [
            c for c in existing
            if c.get("lact") is not None and (int(c.get("lact", 0)) if c.get("lact") is not None else 0) >= 1
        ]
        return len(lactated)

    def _is_cow_disposed(self, cow_auto_id) -> bool:
        if cow_auto_id is None:
            return True
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for ev in events:
            if ev.get("event_number") in [self.rule_engine.EVENT_SOLD, self.rule_engine.EVENT_DEAD]:
                return True
        return False

    def _get_first_parity_ratio_percent(self) -> Optional[float]:
        """現状の初産割合（1産/経産牛）を%で返す。経産牛0頭の場合はNone"""
        all_cows = self.db.get_all_cows()
        existing = [c for c in all_cows if not self._is_cow_disposed(c.get("auto_id"))]
        lactated = [
            c for c in existing
            if c.get("lact") is not None and (int(c.get("lact", 0)) if c.get("lact") is not None else 0) >= 1
        ]
        total = len(lactated)
        if total == 0:
            return None
        first_parity = sum(1 for c in lactated if (int(c.get("lact", 0)) if c.get("lact") is not None else 0) == 1)
        return round(100.0 * first_parity / total, 1)

    def _get_goal_values(self):
        """農場目標設定から目標分娩間隔・更新率・流産率を取得"""
        goals = self.settings_manager.get("farm_goals", {}) or {}
        calving_interval = goals.get("calving_interval_days")
        if calving_interval is None:
            calving_interval = 420
        replacement = goals.get("first_lactation_ratio")
        if replacement is None:
            replacement = 30
        abortion = goals.get("abortion_rate_percent")
        if abortion is None:
            abortion = 10
        return calving_interval, replacement, abortion

    def _save_goals_to_settings(self) -> bool:
        """現在の入力値（目標分娩間隔・更新率・流産率）を農場目標に保存。有効な場合True。"""
        try:
            interval = int(self._interval_var.get().strip())
            replacement = int(self._replacement_var.get().strip())
            abortion = int(self._abortion_var.get().strip())
        except ValueError:
            return False
        if interval <= 0 or replacement < 0 or replacement > 100 or abortion < 0 or abortion > 100:
            return False
        goals = self.settings_manager.get("farm_goals", {}) or {}
        goals["calving_interval_days"] = interval
        goals["first_lactation_ratio"] = replacement
        goals["abortion_rate_percent"] = abortion
        self.settings_manager.set("farm_goals", goals)
        return True

    def _refresh(self):
        """表示を更新"""
        current_parous = self._get_current_parous_count()
        goal_interval, goal_replacement, goal_abortion = self._get_goal_values()

        if self._parous_var.get().strip() == "":
            self._parous_var.set(str(current_parous))

        # 目標分娩間隔・更新率・流産率は入力欄から取得（無効なら農場目標の値を使用）
        try:
            target_interval = int(self._interval_var.get().strip())
            if target_interval <= 0:
                target_interval = goal_interval
        except ValueError:
            target_interval = goal_interval
        try:
            replacement = int(self._replacement_var.get().strip())
            if replacement < 0 or replacement > 100:
                replacement = goal_replacement
        except ValueError:
            replacement = goal_replacement
        try:
            abortion = int(self._abortion_var.get().strip())
            if abortion < 0 or abortion > 100:
                abortion = goal_abortion
        except ValueError:
            abortion = goal_abortion

        try:
            n = int(self._parous_var.get().strip())
        except ValueError:
            n = current_parous

        if n < 0 or target_interval <= 0:
            self._result_var.set("—")
            self._formula_var.set("想定経産牛頭数・目標分娩間隔を正しく入力してください。")
            return

        result = _annual_required_pregnancies(n, float(target_interval), float(replacement), float(abortion))
        self._result_var.set(f"{result:.1f} 頭")

        formula_text = (
            f"年間必要妊娠頭数 ＝ {n} × (1 － {replacement}％) × 365 ÷ {int(target_interval)} × (1 ＋ {abortion}％)\n"
            f"                 ＝ {n} × {1 - replacement/100:.2f} × 365 ÷ {target_interval:.0f} × {1 + abortion/100:.2f} ＝ {result:.1f} 頭"
        )
        self._formula_var.set(formula_text)

    def _on_recalc(self):
        """保存ボタン: 入力値を農場目標に保存し表示を更新"""
        if self._save_goals_to_settings():
            self._refresh()
        else:
            messagebox.showwarning("入力確認", "目標分娩間隔（正の整数）、更新率・流産率（0～100の整数）を入力してください。")
            self._refresh()

    def _build_header(self):
        """統一デザインのヘッダー（アイコン・タイトル・サブタイトル）"""
        header = tk.Frame(self.window, bg=self._BG, pady=16, padx=24)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="\U0001f4ca", font=(self._FONT, 24), bg=self._BG, fg=self._ICON_FG
        ).pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=self._BG)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(
            title_frame,
            text="必要妊娠シミュレーション",
            font=(self._FONT, 16, "bold"),
            bg=self._BG,
            fg=self._TITLE_FG,
        ).pack(anchor=tk.W)
        tk.Label(
            title_frame,
            text="想定経産牛頭数と農場目標設定に基づき、年間必要妊娠頭数を計算します。",
            font=(self._FONT, 10),
            bg=self._BG,
            fg=self._SUBTITLE_FG,
        ).pack(anchor=tk.W)

    def _create_widgets(self):
        self._build_header()

        main = tk.Frame(self.window, bg=self._BG, padx=24, pady=8)
        main.pack(fill=tk.BOTH, expand=True)

        # 入力エリア（カード風）
        form_card = tk.Frame(
            main,
            bg=self._BG,
            padx=16,
            pady=14,
            highlightbackground=self._CARD_HIGHLIGHT,
            highlightthickness=1,
        )
        form_card.pack(fill=tk.X, pady=(0, 12))

        def add_row(parent_card, label_text, var, unit_text, ref_text):
            row = tk.Frame(parent_card, bg=self._BG)
            row.pack(fill=tk.X, pady=6)
            tk.Label(
                row, text=label_text, font=(self._FONT, 10), width=16, anchor=tk.W, bg=self._BG, fg=self._TITLE_FG
            ).pack(side=tk.LEFT, padx=(0, 8))
            entry = ttk.Entry(row, textvariable=var, width=10)
            entry.pack(side=tk.LEFT, padx=(0, 8))
            tk.Label(row, text=unit_text, font=(self._FONT, 10), bg=self._BG, fg=self._TITLE_FG).pack(side=tk.LEFT)
            tk.Label(
                row, text=ref_text, font=(self._FONT, 9), bg=self._BG, fg=self._DESC_FG
            ).pack(side=tk.LEFT, padx=(12, 0))
            return entry

        current_parous = self._get_current_parous_count()
        parous_entry = add_row(form_card, "想定経産牛頭数", self._parous_var, "頭", f"（参考：現状 {current_parous} 頭）")
        parous_entry.bind("<KeyRelease>", lambda e: self._refresh())

        target_interval, replacement, abortion = self._get_goal_values()
        self._interval_var.set(str(int(target_interval)))
        self._replacement_var.set(str(int(replacement)))
        self._abortion_var.set(str(int(abortion)))
        ref_first = self._get_first_parity_ratio_percent()
        ref_first_str = (
            f"（参考：現状の初産割合 {ref_first}％）" if ref_first is not None else "（参考：経産牛0頭のため算出不可）"
        )

        interval_entry = add_row(
            form_card,
            "目標分娩間隔",
            self._interval_var,
            "日",
            "（変更は農場の目標設定に反映されます）",
        )
        interval_entry.bind("<KeyRelease>", lambda e: self._refresh())

        replacement_entry = add_row(form_card, "更新率", self._replacement_var, "％", ref_first_str)
        replacement_entry.bind("<KeyRelease>", lambda e: self._refresh())

        abortion_entry = add_row(
            form_card,
            "流産率",
            self._abortion_var,
            "％",
            "（変更は農場の目標設定に反映されます）",
        )
        abortion_entry.bind("<KeyRelease>", lambda e: self._refresh())

        # 計算結果（強調表示）
        result_card = tk.Frame(
            main,
            bg=self._CARD_HIGHLIGHT,
            padx=14,
            pady=12,
            highlightbackground=self._CARD_HIGHLIGHT,
            highlightthickness=1,
        )
        result_card.pack(fill=tk.X, pady=(4, 8))
        tk.Label(
            result_card,
            text="年間必要妊娠頭数",
            font=(self._FONT, 11, "bold"),
            bg=self._CARD_HIGHLIGHT,
            fg=self._TITLE_FG,
        ).pack(anchor=tk.W)
        result_label = tk.Label(
            result_card,
            textvariable=self._result_var,
            font=(self._FONT, 20, "bold"),
            fg=self._ICON_FG,
            bg=self._CARD_HIGHLIGHT,
        )
        result_label.pack(anchor=tk.W)

        # 計算式
        tk.Label(
            main,
            text="計算式",
            font=(self._FONT, 10, "bold"),
            bg=self._BG,
            fg=self._TITLE_FG,
        ).pack(anchor=tk.W, pady=(8, 4))
        formula_label = tk.Label(
            main,
            textvariable=self._formula_var,
            font=(self._FONT, 9),
            fg=self._SUBTITLE_FG,
            bg=self._BG,
            justify=tk.LEFT,
        )
        formula_label.pack(anchor=tk.W)

        # フッター（統一ボタン）
        footer = tk.Frame(self.window, bg=self._BG, pady=16)
        footer.pack(fill=tk.X)
        tk.Button(
            footer,
            text="保存",
            font=(self._FONT, 10),
            bg=self._BTN_PRIMARY_BG,
            fg=self._BTN_PRIMARY_FG,
            activebackground="#303f9f",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._on_recalc,
        ).pack(side=tk.LEFT, padx=(24, 8))
        tk.Button(
            footer,
            text="閉じる",
            font=(self._FONT, 10),
            bg=self._BTN_SECONDARY_BG,
            fg=self._BTN_SECONDARY_FG,
            activebackground="#eceff1",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            highlightbackground=self._BTN_SECONDARY_BD,
            highlightthickness=1,
            cursor="hand2",
            command=self.window.destroy,
        ).pack(side=tk.LEFT, padx=4)

    def show(self):
        self.window.wait_window()
