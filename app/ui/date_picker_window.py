"""
FALCON2 - 統一カレンダーコンポーネント
左右の矢印ボタンで月を変更できるカレンダー
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
from typing import Optional, Callable
from calendar import monthrange


class DatePickerWindow:
    """統一カレンダーウィンドウ"""
    
    def __init__(self, parent: tk.Tk, initial_date: Optional[str] = None, 
                 on_date_selected: Optional[Callable[[str], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            initial_date: 初期日付（YYYY-MM-DD形式、Noneの場合は今日）
            on_date_selected: 日付選択時のコールバック関数（date_strを引数に取る）
        """
        self.parent = parent
        self.on_date_selected = on_date_selected
        self.selected_date: Optional[str] = None
        
        # 初期日付の設定
        if initial_date:
            try:
                self.current_date = datetime.strptime(initial_date, '%Y-%m-%d')
            except:
                self.current_date = datetime.now()
        else:
            self.current_date = datetime.now()
        
        self.window = tk.Toplevel(parent)
        self.window.title("日付選択")
        self.window.geometry("300x350")
        
        self._create_widgets()
        self._update_calendar()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # コントロールフレーム（年月選択と矢印ボタン）
        control_frame = ttk.Frame(self.window)
        control_frame.pack(pady=10)
        
        # 年と月の変数
        self.year_var = tk.IntVar(value=self.current_date.year)
        self.month_var = tk.IntVar(value=self.current_date.month)
        
        # 前月ボタン（左矢印）
        prev_btn = ttk.Button(
            control_frame,
            text="◀",
            command=self._prev_month,
            width=3
        )
        prev_btn.pack(side=tk.LEFT, padx=5)
        
        # 年Spinbox
        year_spinbox = ttk.Spinbox(
            control_frame,
            from_=2000,
            to=2100,
            textvariable=self.year_var,
            width=6,
            command=self._update_calendar
        )
        year_spinbox.pack(side=tk.LEFT, padx=5)
        year_spinbox.bind('<Return>', lambda e: self._update_calendar())
        
        # 月Spinbox
        month_spinbox = ttk.Spinbox(
            control_frame,
            from_=1,
            to=12,
            textvariable=self.month_var,
            width=4,
            command=self._update_calendar
        )
        month_spinbox.pack(side=tk.LEFT, padx=5)
        month_spinbox.bind('<Return>', lambda e: self._update_calendar())
        
        # 次月ボタン（右矢印）
        next_btn = ttk.Button(
            control_frame,
            text="▶",
            command=self._next_month,
            width=3
        )
        next_btn.pack(side=tk.LEFT, padx=5)
        
        # 年月変更時の更新
        self.year_var.trace('w', lambda *args: self._update_calendar())
        self.month_var.trace('w', lambda *args: self._update_calendar())
        
        # カレンダー表示フレーム
        self.calendar_frame = ttk.Frame(self.window)
        self.calendar_frame.pack(padx=10, pady=10)
        
        # 閉じるボタン
        close_btn = ttk.Button(
            self.window,
            text="閉じる",
            command=self._on_close,
            width=10
        )
        close_btn.pack(pady=10)
    
    def _prev_month(self):
        """前月に移動"""
        current_year = self.year_var.get()
        current_month = self.month_var.get()
        
        new_month = current_month - 1
        if new_month < 1:
            new_month = 12
            current_year -= 1
        
        self.year_var.set(current_year)
        self.month_var.set(new_month)
        self._update_calendar()
    
    def _next_month(self):
        """次月に移動"""
        current_year = self.year_var.get()
        current_month = self.month_var.get()
        
        new_month = current_month + 1
        if new_month > 12:
            new_month = 1
            current_year += 1
        
        self.year_var.set(current_year)
        self.month_var.set(new_month)
        self._update_calendar()
    
    def _update_calendar(self):
        """カレンダーを更新"""
        # 既存のウィジェットを削除
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        year = self.year_var.get()
        month = self.month_var.get()
        
        # 曜日ヘッダー
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        for i, day in enumerate(weekdays):
            label = ttk.Label(
                self.calendar_frame,
                text=day,
                width=4,
                anchor="center"
            )
            label.grid(row=0, column=i, padx=1, pady=1)
        
        # 月の最初の日と最後の日を取得
        first_day = datetime(year, month, 1)
        days_in_month = monthrange(year, month)[1]
        
        # 最初の日の曜日（月曜日=0）
        first_weekday = (first_day.weekday()) % 7
        
        # 日付ボタンを配置
        row = 1
        col = first_weekday
        
        for day in range(1, days_in_month + 1):
            date_str = f"{year}-{month:02d}-{day:02d}"
            btn = ttk.Button(
                self.calendar_frame,
                text=str(day),
                width=4,
                command=lambda d=date_str: self._select_date(d)
            )
            btn.grid(row=row, column=col, padx=1, pady=1)
            
            col += 1
            if col > 6:
                col = 0
                row += 1
    
    def _select_date(self, date_str: str):
        """日付を選択"""
        self.selected_date = date_str
        if self.on_date_selected:
            self.on_date_selected(date_str)
        self.window.destroy()
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示して待機"""
        self.window.wait_window()
        return self.selected_date
