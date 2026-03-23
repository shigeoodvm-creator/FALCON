"""
FALCON2 - 繁殖検診設定ウインドウ
繁殖検診の設定を管理するウインドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional
import json

from settings_manager import SettingsManager


class ReproductionCheckupSettingsWindow:
    """繁殖検診設定ウインドウ"""
    
    # 検診区分定義
    CHECKUP_CATEGORIES = {
        # 経産牛
        'fresh': {
            'name': 'フレッシュチェック', 
            'default_days': 30, 
            'type': 'days',
            'description': '分娩後N日以上経過した個体'
        },
        'repro1': {
            'name': '繁殖検査１', 
            'default_days': None, 
            'type': 'none',
            'description': 'フレッシュチェック後、AI/ET履歴がない個体'
        },
        'repro2': {
            'name': '繁殖検査２', 
            'default_days': None, 
            'type': 'none',
            'description': 'AI/ET後、妊娠鑑定マイナスがあり、現在Openの個体'
        },
        'preg': {
            'name': '妊娠鑑定', 
            'default_days': 30, 
            'type': 'days',
            'description': '最終授精日からN日以上経過した個体'
        },
        'repreg': {
            'name': '再妊娠鑑定', 
            'default_days': 60, 
            'type': 'days',
            'description': '妊娠中で、最終授精日からN日以上経過した個体'
        },
        'preg2': {
            'name': '任意妊娠鑑定', 
            'default_days': None, 
            'type': 'days_optional',
            'description': '妊娠中で、最終授精日からN日以上経過（任意設定）'
        },
        'due_over': {
            'name': '分娩予定超過', 
            'default_days': 14, 
            'type': 'days',
            'description': '分娩予定日からN日以上超過した個体'
        },
        'check': {
            'name': 'チェック', 
            'default_days': None, 
            'type': 'none',
            'description': '妊娠プラス後、妊娠マイナスがなく、その後AI/ETがある個体'
        },
        # 未経産牛
        'heifer_repro1': {
            'name': '育成繁殖１', 
            'default_age_months': 12, 
            'type': 'age_months',
            'description': '月齢Nヶ月以上で、AI/ET履歴がない個体'
        },
        'heifer_repro2': {
            'name': '育成繁殖２', 
            'default_days': None, 
            'type': 'none',
            'description': 'AI/ET後、妊娠鑑定マイナスがあり、現在Openの個体'
        },
        'heifer_preg': {
            'name': '育成妊娠鑑定', 
            'default_days': 30, 
            'type': 'days',
            'description': '最終授精日からN日以上経過した個体'
        },
        'heifer_repreg': {
            'name': '育成再妊娠鑑定', 
            'default_days': 60, 
            'type': 'days',
            'description': '妊娠中で、最終授精日からN日以上経過した個体'
        },
    }
    
    def __init__(self, parent: tk.Widget, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_manager = SettingsManager(self.farm_path)
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖検診設定")
        self.window.geometry("600x700")
        
        # 設定を読み込む
        self.settings = self._load_settings()
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _load_settings(self) -> Dict[str, Any]:
        """設定を読み込む"""
        settings = self.settings_manager.get('repro_checkup_settings', {})
        
        # デフォルト設定をマージ
        default_settings = {}
        for key, category in self.CHECKUP_CATEGORIES.items():
            default_settings[key] = {
                'enabled': True,
                'days': category.get('default_days'),
                'age_months': category.get('default_age_months')
            }
        
        # 既存設定とデフォルトをマージ
        for key in default_settings:
            if key not in settings:
                settings[key] = default_settings[key]
            else:
                # 既存設定に不足している項目を追加
                for sub_key in default_settings[key]:
                    if sub_key not in settings[key]:
                        settings[key][sub_key] = default_settings[key][sub_key]
        
        return settings
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        title_label = ttk.Label(
            main_frame,
            text="繁殖検診設定",
            font=("", 14, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # スクロール可能なフレーム
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 経産牛セクション
        parous_frame = ttk.LabelFrame(scrollable_frame, text="経産牛", padding=10)
        parous_frame.pack(fill=tk.X, pady=(0, 10))
        
        parous_categories = [
            'fresh', 'repro1', 'repro2', 'preg', 'repreg', 
            'preg2', 'due_over', 'check'
        ]
        
        for key in parous_categories:
            self._create_category_row(parous_frame, key)
        
        # 未経産牛セクション
        heifer_frame = ttk.LabelFrame(scrollable_frame, text="未経産牛", padding=10)
        heifer_frame.pack(fill=tk.X, pady=(0, 10))
        
        heifer_categories = [
            'heifer_repro1', 'heifer_repro2', 'heifer_preg', 'heifer_repreg'
        ]
        
        for key in heifer_categories:
            self._create_category_row(heifer_frame, key)
        
        # スクロールバーとCanvasの配置
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        save_btn = ttk.Button(
            button_frame,
            text="保存",
            command=self._on_save,
            width=15
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self._on_cancel,
            width=15
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def _create_category_row(self, parent: ttk.Frame, key: str):
        """検診区分の行を作成"""
        category = self.CHECKUP_CATEGORIES[key]
        name = category['name']
        cat_type = category['type']
        description = category.get('description', '')
        
        # メイン行フレーム（縦方向に配置）
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=3)
        
        # 上段：チェックボックスと入力欄
        top_frame = ttk.Frame(row_frame)
        top_frame.pack(fill=tk.X, pady=(0, 2))
        
        # 有効/無効チェックボックス
        enabled_var = tk.BooleanVar(value=self.settings[key].get('enabled', True))
        enabled_check = ttk.Checkbutton(
            top_frame,
            text=name,
            variable=enabled_var
        )
        enabled_check.pack(side=tk.LEFT, padx=5)
        
        # 設定値を保持
        self.settings[key]['_enabled_var'] = enabled_var
        
        # 日数/月齢入力
        if cat_type == 'days':
            ttk.Label(top_frame, text="基準日数:").pack(side=tk.LEFT, padx=(20, 5))
            days_var = tk.StringVar(value=str(self.settings[key].get('days', '') or ''))
            days_entry = ttk.Entry(top_frame, textvariable=days_var, width=10)
            days_entry.pack(side=tk.LEFT, padx=5)
            self.settings[key]['_days_var'] = days_var
        elif cat_type == 'days_optional':
            ttk.Label(top_frame, text="基準日数（任意）:").pack(side=tk.LEFT, padx=(20, 5))
            days_var = tk.StringVar(value=str(self.settings[key].get('days', '') or ''))
            days_entry = ttk.Entry(top_frame, textvariable=days_var, width=10)
            days_entry.pack(side=tk.LEFT, padx=5)
            self.settings[key]['_days_var'] = days_var
        elif cat_type == 'age_months':
            ttk.Label(top_frame, text="基準月齢:").pack(side=tk.LEFT, padx=(20, 5))
            age_var = tk.StringVar(value=str(self.settings[key].get('age_months', '') or ''))
            age_entry = ttk.Entry(top_frame, textvariable=age_var, width=10)
            age_entry.pack(side=tk.LEFT, padx=5)
            self.settings[key]['_age_months_var'] = age_var
        
        # 下段：説明ラベル（小さなフォント、グレー色、左にインデント）
        if description:
            desc_label = ttk.Label(
                row_frame,
                text=f"  ※ {description}",
                font=("", 8),
                foreground="gray"
            )
            desc_label.pack(side=tk.LEFT, padx=(30, 5), anchor=tk.W)
    
    def _on_save(self):
        """保存ボタンをクリック"""
        # 設定値を収集
        saved_settings = {}
        
        for key in self.CHECKUP_CATEGORIES:
            category = self.CHECKUP_CATEGORIES[key]
            cat_type = category['type']
            
            setting = {
                'enabled': self.settings[key]['_enabled_var'].get()
            }
            
            if cat_type in ('days', 'days_optional'):
                days_str = self.settings[key]['_days_var'].get().strip()
                if days_str:
                    try:
                        setting['days'] = int(days_str)
                    except ValueError:
                        messagebox.showerror("エラー", f"{category['name']}の基準日数が無効です。")
                        return
                else:
                    setting['days'] = None
            
            if cat_type == 'age_months':
                age_str = self.settings[key]['_age_months_var'].get().strip()
                if age_str:
                    try:
                        setting['age_months'] = float(age_str)
                    except ValueError:
                        messagebox.showerror("エラー", f"{category['name']}の基準月齢が無効です。")
                        return
                else:
                    setting['age_months'] = None
            
            saved_settings[key] = setting
        
        # 設定を保存
        self.settings_manager.set('repro_checkup_settings', saved_settings)
        
        messagebox.showinfo("完了", "設定を保存しました。")
        self.window.destroy()
    
    def _on_cancel(self):
        """キャンセルボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

