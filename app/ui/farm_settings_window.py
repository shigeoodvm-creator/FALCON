"""
FALCON2 - 農場設定ウインドウ
農場設定のメインウィンドウ
"""

import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import Optional

from ui.insemination_settings_window import InseminationSettingsWindow
from ui.pen_settings_window import PenSettingsWindow


class FarmSettingsWindow:
    """農場設定ウインドウ"""
    
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
        self.window.geometry("500x400")
        self.window.transient(parent)
        self.window.grab_set()
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        title_label = ttk.Label(
            main_frame,
            text="農場設定",
            font=("", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 説明
        desc_label = ttk.Label(
            main_frame,
            text="設定項目を選択してください",
            font=("", 9)
        )
        desc_label.pack(pady=(0, 20))
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        # 授精設定ボタン
        insemination_btn = ttk.Button(
            button_frame,
            text="授精設定",
            command=self._on_insemination_settings,
            width=25
        )
        insemination_btn.pack(pady=10, padx=20)

        # PEN設定ボタン
        pen_btn = ttk.Button(
            button_frame,
            text="PEN設定",
            command=self._on_pen_settings,
            width=25
        )
        pen_btn.pack(pady=10, padx=20)
        
        # その他の設定ボタン（将来の拡張用）
        # TODO: 他の設定項目を追加
        
        # 閉じるボタン
        close_btn = ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close,
            width=25
        )
        close_btn.pack(pady=10, padx=20)
    
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
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()



