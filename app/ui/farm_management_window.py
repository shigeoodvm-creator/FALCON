"""
FALCON2 - 農場管理ウィンドウ
農場の切り替えと新規作成を管理
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Callable

from ui.farm_selector import FarmSelectorWindow
from ui.farm_create_window import FarmCreateWindow


class FarmManagementWindow:
    """農場管理ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, current_farm_path: Optional[Path] = None, 
                 on_farm_changed: Optional[Callable[[Path], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            current_farm_path: 現在開いている農場のパス
            on_farm_changed: 農場が変更された時のコールバック関数（farm_path を引数に取る）
        """
        self.parent = parent
        self.current_farm_path = current_farm_path
        self.on_farm_changed = on_farm_changed
        
        self.window = tk.Toplevel(parent)
        self.window.title("農場管理")
        self.window.minsize(400, 320)
        self.window.geometry("450x320")
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.window,
            text="農場管理",
            font=("", 14, "bold")
        )
        title_label.pack(pady=20)
        
        # 現在の農場表示
        if self.current_farm_path:
            current_farm_label = ttk.Label(
                self.window,
                text=f"現在の農場: {self.current_farm_path.name}",
                font=("", 10)
            )
            current_farm_label.pack(pady=10)
        else:
            current_farm_label = ttk.Label(
                self.window,
                text="現在の農場: なし",
                font=("", 10),
                foreground="gray"
            )
            current_farm_label.pack(pady=10)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=30)
        
        # 農場切り替えボタン
        switch_farm_btn = ttk.Button(
            button_frame,
            text="農場切り替え",
            command=self._on_switch_farm,
            width=20
        )
        switch_farm_btn.pack(pady=10)
        
        # 農場新規作成ボタン
        create_farm_btn = ttk.Button(
            button_frame,
            text="農場新規作成",
            command=self._on_create_farm,
            width=20
        )
        create_farm_btn.pack(pady=10)
        
        # 閉じるボタン
        close_btn = ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close,
            width=20
        )
        close_btn.pack(pady=10)
    
    def _on_switch_farm(self):
        """農場切り替えボタンをクリック"""
        def on_farm_selected(farm_path: Path):
            """農場が選択された時のコールバック"""
            if self.on_farm_changed:
                self.on_farm_changed(farm_path)
            # ウィンドウを閉じる（切り替え処理はコールバック内で行われる）
            self.window.destroy()
        
        # 農場選択ウィンドウを表示
        farm_selector = FarmSelectorWindow(
            parent=self.window,
            on_farm_selected=on_farm_selected
        )
        farm_selector.show()
    
    def _on_create_farm(self):
        """農場新規作成ボタンをクリック"""
        def on_farm_created(farm_path: Path):
            """農場が作成された時のコールバック"""
            if self.on_farm_changed:
                self.on_farm_changed(farm_path)
            # ウィンドウを閉じる（切り替え処理はコールバック内で行われる）
            self.window.destroy()
        
        # 新規農場作成ウィンドウを表示
        create_window = FarmCreateWindow(
            parent=self.window,
            on_farm_created=on_farm_created
        )
        created_path = create_window.show()
        
        # 作成された場合は自動的に切り替える（コールバックはFarmCreateWindow内で呼ばれる）
        if created_path and self.on_farm_changed:
            self.on_farm_changed(created_path)
            self.window.destroy()
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
