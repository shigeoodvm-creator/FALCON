"""
FALCON2 - 目標設定ウインドウ
農場の目標値を管理する
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional

from settings_manager import SettingsManager


class FarmGoalSettingsWindow:
    """目標設定ウインドウ"""

    DEFAULTS: Dict[str, Any] = {
        "parous_cow_count": None,          # 経産牛頭数
        "conception_rate": 40,             # 受胎率（%）
        "first_service_conception_rate": 40,
        "second_service_conception_rate": 40,
        "insemination_rate": 60,           # 授精実施率（%）
        "pregnancy_rate": 20,              # 妊娠率（%）
        "pregnant_cow_ratio": 50,          # 妊娠牛割合（%）
        "calving_interval_days": 420,      # 目標分娩間隔（日）（年間必要妊娠頭数計算にも使用）
        "first_lactation_ratio": 30,       # 初産割合（更新率）（%）
        "abortion_rate_percent": 10,      # 流産率（%）- 年間必要妊娠頭数計算用
        "herd_linear_score": 2.5,          # 牛群リニアスコア
        "individual_scc": 200,             # 個体体細胞（千）
        "milk_100d_heifer": 30,            # 分娩後100日以内乳量（初産）
        "milk_100d_parous": 40,            # 分娩後100日以内乳量（経産）
        "average_milk_yield": 32           # 平均乳量（kg）
    }

    FIELDS = [
        ("parous_cow_count", "経産牛頭数", "", "int"),
        ("conception_rate", "受胎率", "%", "percent"),
        ("first_service_conception_rate", "初回授精受胎率", "%", "percent"),
        ("second_service_conception_rate", "二回授精受胎率", "%", "percent"),
        ("insemination_rate", "授精実施率", "%", "percent"),
        ("pregnancy_rate", "妊娠率", "%", "percent"),
        ("pregnant_cow_ratio", "妊娠牛割合", "%", "percent"),
        ("calving_interval_days", "目標分娩間隔", "日", "int"),
        ("first_lactation_ratio", "初産割合（更新率）", "%", "percent"),
        ("abortion_rate_percent", "流産率", "%", "percent"),
        ("herd_linear_score", "牛群リニアスコア", "", "float"),
        ("individual_scc", "個体体細胞（SCC）", "千", "int"),
        ("milk_100d_heifer", "分娩後100日以内乳量（初産）", "kg", "float"),
        ("milk_100d_parous", "分娩後100日以内乳量（経産）", "kg", "float"),
        ("average_milk_yield", "平均乳量", "kg", "float"),
    ]

    def __init__(self, parent: tk.Tk, farm_path: Path):
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_manager = SettingsManager(self.farm_path)

        self.window = tk.Toplevel(parent)
        self.window.title("目標設定")
        self.window.geometry("520x620")

        self.vars: Dict[str, tk.StringVar] = {}

        self._create_widgets()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _load_values(self) -> Dict[str, Any]:
        stored = self.settings_manager.get("farm_goals", {}) or {}
        values = {}
        for key, default in self.DEFAULTS.items():
            value = stored.get(key)
            values[key] = default if value is None else value
        return values

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(
            main_frame,
            text="目標設定",
            font=("", 14, "bold")
        )
        title_label.pack(pady=(0, 10))

        info_label = ttk.Label(
            main_frame,
            text="目標値を入力してください（後でレポート参照に使用）",
            font=("", 9)
        )
        info_label.pack(pady=(0, 10))

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True)

        values = self._load_values()

        for row, (key, label, unit, _dtype) in enumerate(self.FIELDS):
            ttk.Label(form_frame, text=label).grid(row=row, column=0, sticky=tk.W, padx=5, pady=4)
            var = tk.StringVar()
            current_val = values.get(key)
            var.set("" if current_val is None else str(current_val))
            entry = ttk.Entry(form_frame, textvariable=var, width=18)
            entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=4)
            if unit:
                ttk.Label(form_frame, text=unit, width=6).grid(row=row, column=2, sticky=tk.W, padx=5, pady=4)
            # 参考値の表示（デフォルト値）
            default_val = self.DEFAULTS.get(key)
            if default_val is not None:
                ref_text = f"参考：{default_val}{unit}"
            else:
                ref_text = "参考：-"
            ref_label = ttk.Label(form_frame, text=ref_text, font=("", 9), foreground="gray")
            ref_label.grid(row=row, column=3, sticky=tk.W, padx=(10, 5), pady=4)
            self.vars[key] = var

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        save_btn = ttk.Button(
            button_frame,
            text="保存",
            command=self._on_save,
            width=12
        )
        save_btn.pack(side=tk.LEFT, padx=5)

        close_btn = ttk.Button(
            button_frame,
            text="閉じる",
            command=self.window.destroy,
            width=12
        )
        close_btn.pack(side=tk.LEFT, padx=5)

    def _parse_value(self, value: str, dtype: str) -> Optional[float]:
        value = value.strip()
        if value == "":
            return None
        if dtype in ("int", "percent"):
            if not value.replace(".", "", 1).isdigit():
                return None
            return int(float(value))
        try:
            return float(value)
        except ValueError:
            return None

    def _on_save(self):
        settings: Dict[str, Any] = {}
        for key, _label, _unit, dtype in self.FIELDS:
            raw = self.vars[key].get()
            parsed = self._parse_value(raw, dtype)
            if raw.strip() != "" and parsed is None:
                messagebox.showwarning("警告", f"{_label} は数値で入力してください")
                return
            settings[key] = parsed

        self.settings_manager.set("farm_goals", settings)
        messagebox.showinfo("完了", "目標設定を保存しました")

    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
