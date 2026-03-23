"""
FALCON2 - 農場設定ウインドウ
農場設定のメインウィンドウ（辞書・設定ダイアログと同一デザイン）
"""

import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import Optional

from ui.insemination_settings_window import InseminationSettingsWindow
from ui.pen_settings_window import PenSettingsWindow
from ui.reproduction_checkup_settings_window import ReproductionCheckupSettingsWindow
from ui.reproduction_treatment_settings_window import ReproductionTreatmentSettingsWindow
from ui.drug_settings_window import DrugSettingsWindow
from ui.farm_goal_settings_window import FarmGoalSettingsWindow


class FarmSettingsWindow:
    """農場設定ウインドウ（辞書・設定と統一イメージ）"""
    
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
    
    def __init__(self, parent: tk.Tk, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（MainWindow）
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("農場設定")
        self.window.geometry("520x580")
        self.window.configure(bg=self._BG)
        self.window.minsize(480, 400)
        
        # UI作成（統一デザイン）
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _build_header(self, icon_char: str, title_text: str, subtitle_text: str):
        """統一デザインのヘッダー（アイコン・タイトル・サブタイトル）"""
        header = tk.Frame(self.window, bg=self._BG, pady=20, padx=24)
        header.pack(fill=tk.X)
        tk.Label(header, text=icon_char, font=(self._FONT, 24), bg=self._BG, fg=self._ICON_FG).pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=self._BG)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text=title_text, font=(self._FONT, 16, "bold"), bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)
        tk.Label(title_frame, text=subtitle_text, font=(self._FONT, 10), bg=self._BG, fg=self._SUBTITLE_FG).pack(anchor=tk.W)
    
    def _build_card(self, parent: tk.Frame, icon_char: str, title_text: str, desc_text: str, btn_label: str, command):
        """統一デザインのカード（アイコン・タイトル・説明・ボタン）"""
        card = tk.Frame(parent, bg=self._BG, padx=16, pady=16, highlightbackground=self._CARD_HIGHLIGHT, highlightthickness=1)
        card.pack(fill=tk.X, pady=6, padx=24)
        tk.Label(card, text=icon_char, font=(self._FONT, 22), bg=self._BG, fg=self._ICON_FG).pack(side=tk.LEFT, padx=(0, 14), pady=4)
        text_frame = tk.Frame(card, bg=self._BG)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(text_frame, text=title_text, font=(self._FONT, 11, "bold"), bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)
        tk.Label(text_frame, text=desc_text, font=(self._FONT, 9), bg=self._BG, fg=self._DESC_FG).pack(anchor=tk.W)
        btn = tk.Button(card, text=btn_label, font=(self._FONT, 10), bg=self._BTN_PRIMARY_BG, fg=self._BTN_PRIMARY_FG,
                       activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
                       padx=16, pady=8, cursor="hand2", command=command)
        btn.pack(side=tk.RIGHT, padx=(10, 0))
    
    def _build_footer(self):
        """統一デザインのフッター（閉じるボタン）"""
        footer = tk.Frame(self.window, bg=self._BG, pady=20)
        footer.pack(fill=tk.X)
        tk.Button(footer, text="閉じる", font=(self._FONT, 10), bg=self._BTN_SECONDARY_BG, fg=self._BTN_SECONDARY_FG,
                  activebackground="#eceff1", relief=tk.FLAT, padx=24, pady=10,
                  highlightbackground=self._BTN_SECONDARY_BD, highlightthickness=1, cursor="hand2",
                  command=self._on_close).pack()
    
    def _create_widgets(self):
        """ウィジェットを作成（辞書・設定ダイアログと同一レイアウト）。一覧はスクロール可能。"""
        self._build_header("\U0001f3e0", "農場設定", "設定項目を選択してください")
        
        # スクロール可能なコンテナ（全カードが小さなウィンドウでも見えるように）
        container = tk.Frame(self.window, bg=self._BG)
        container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(container, bg=self._BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container)
        
        content = tk.Frame(canvas, bg=self._BG, padx=0, pady=8)
        content_window = canvas.create_window((0, 0), window=content, anchor=tk.NW)
        
        def _on_content_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(event):
            canvas.itemconfig(content_window, width=event.width)
        content.bind("<Configure>", _on_content_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=canvas.yview)
        
        # マウスホイールでスクロール（このウィンドウにフォーカスがあるとき）
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        self._build_card(content, "\U0001f4c4", "授精設定", "授精種類・授精師などの授精関連を設定します", "開く", self._on_insemination_settings)
        self._build_card(content, "\U0001f4cb", "PEN設定", "群（PEN）のコードと名前を設定します", "開く", self._on_pen_settings)
        self._build_card(content, "\U0001f4dd", "繁殖検診設定", "繁殖検診の項目とフローを設定します", "開く", self._on_reproduction_checkup_settings)
        self._build_card(content, "\U0001f48a", "繁殖処置設定", "繁殖処置の種類とコードを設定します", "開く", self._on_reproduction_treatment_settings)
        self._build_card(content, "\U0001f9ea", "農場薬品設定", "農場で使用する薬品を設定します", "開く", self._on_drug_settings)
        self._build_card(content, "\U0001f3af", "目標設定", "農場の目標・KPIを設定します", "開く", self._on_farm_goal_settings)
        
        self._build_footer()
    
    def _on_insemination_settings(self):
        """授精設定ボタンをクリック"""
        insemination_window = InseminationSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        insemination_window.show()

    def _on_pen_settings(self):
        """PEN設定ボタンをクリック"""
        pen_window = PenSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        pen_window.show()
    
    def _on_reproduction_checkup_settings(self):
        """繁殖検診設定ボタンをクリック"""
        repro_checkup_window = ReproductionCheckupSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        repro_checkup_window.show()
    
    def _on_reproduction_treatment_settings(self):
        """繁殖処置設定ボタンをクリック"""
        treatment_window = ReproductionTreatmentSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        treatment_window.show()

    def _on_drug_settings(self):
        """農場薬品設定ボタンをクリック"""
        drug_window = DrugSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        drug_window.show()

    def _on_farm_goal_settings(self):
        """農場目標設定ボタンをクリック"""
        goal_window = FarmGoalSettingsWindow(
            parent=self.window,
            farm_path=self.farm_path
        )
        goal_window.show()
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        try:
            self.window.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
