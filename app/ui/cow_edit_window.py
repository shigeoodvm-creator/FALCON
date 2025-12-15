"""
FALCON2 - 個体編集ウィンドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable, Dict, Any
from datetime import datetime

from db.db_handler import DBHandler


class CowEditWindow:
    """個体編集ウィンドウ"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 cow_auto_id: int,
                 on_saved: Optional[Callable[[], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            cow_auto_id: 編集する牛の auto_id
            on_saved: 保存時のコールバック
        """
        self.db = db_handler
        self.cow_auto_id = cow_auto_id
        self.on_saved = on_saved
        
        # 牛データを取得
        self.cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not self.cow:
            raise ValueError(f"個体が見つかりません: auto_id={cow_auto_id}")
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("個体編集")
        self.window.geometry("500x400")
        self.window.transient(parent)
        self.window.grab_set()
        
        # 変数
        self.cow_id_var = tk.StringVar(value=self.cow.get('cow_id', ''))
        self.jpn10_var = tk.StringVar(value=self.cow.get('jpn10', ''))
        self.brd_var = tk.StringVar(value=self.cow.get('brd', 'ホルスタイン'))
        self.bthd_var = tk.StringVar(value=self.cow.get('bthd', '') or '')
        self.entr_var = tk.StringVar(value=self.cow.get('entr', '') or '')
        self.lact_var = tk.StringVar(value=str(self.cow.get('lact', 0) or 0))
        self.clvd_var = tk.StringVar(value=self.cow.get('clvd', '') or '')
        self.pen_var = tk.StringVar(value=self.cow.get('pen', 'Lactating'))
        
        self._create_widgets()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.window,
            text="個体情報を編集",
            font=("", 12, "bold")
        )
        title_label.pack(pady=10)
        
        # メインフレーム
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 拡大4桁ID
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="拡大4桁ID:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.cow_id_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 個体識別番号
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="個体識別番号:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.jpn10_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 品種
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="品種:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        brd_combo = ttk.Combobox(row, textvariable=self.brd_var, width=27, values=["ホルスタイン", "ジャージー", "その他"])
        brd_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 生年月日
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="生年月日:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.bthd_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(row, text="(YYYY-MM-DD)", font=("", 8)).pack(side=tk.LEFT, padx=5)
        
        # 入荷日
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="入荷日:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.entr_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(row, text="(YYYY-MM-DD)", font=("", 8)).pack(side=tk.LEFT, padx=5)
        
        # 産次
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="産次:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.lact_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 最終分娩日
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="最終分娩日:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.clvd_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(row, text="(YYYY-MM-DD)", font=("", 8)).pack(side=tk.LEFT, padx=5)
        
        # 群
        row = ttk.Frame(main_frame)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="群:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        pen_combo = ttk.Combobox(row, textvariable=self.pen_var, width=27, values=["Lactating", "Dry", "Heifer", "その他"])
        pen_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 注意事項
        note_label = ttk.Label(
            main_frame,
            text="※ 産次・最終分娩日・繁殖コードは通常、イベントから自動計算されます。",
            font=("", 8),
            foreground="gray"
        )
        note_label.pack(pady=10)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=10)
        
        save_button = ttk.Button(
            button_frame,
            text="保存",
            command=self._on_save,
            width=12
        )
        save_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self._on_cancel,
            width=12
        )
        cancel_button.pack(side=tk.LEFT, padx=5)
    
    def _on_save(self):
        """保存処理"""
        # 入力値を取得
        cow_id = self.cow_id_var.get().strip()
        jpn10 = self.jpn10_var.get().strip()
        brd = self.brd_var.get().strip()
        bthd = self.bthd_var.get().strip() or None
        entr = self.entr_var.get().strip() or None
        lact_str = self.lact_var.get().strip()
        clvd = self.clvd_var.get().strip() or None
        pen = self.pen_var.get().strip()
        
        # バリデーション
        if not cow_id:
            messagebox.showerror("エラー", "拡大4桁IDを入力してください")
            return
        
        if not jpn10:
            messagebox.showerror("エラー", "個体識別番号を入力してください")
            return
        
        if len(jpn10) != 10 or not jpn10.isdigit():
            messagebox.showerror("エラー", "個体識別番号は10桁の数字である必要があります")
            return
        
        # 産次を整数に変換
        try:
            lact = int(lact_str) if lact_str else 0
        except ValueError:
            messagebox.showerror("エラー", "産次は数値である必要があります")
            return
        
        # 日付形式のチェック（簡易）
        if bthd and not self._is_valid_date(bthd):
            messagebox.showerror("エラー", "生年月日の形式が正しくありません (YYYY-MM-DD)")
            return
        
        if entr and not self._is_valid_date(entr):
            messagebox.showerror("エラー", "入荷日の形式が正しくありません (YYYY-MM-DD)")
            return
        
        if clvd and not self._is_valid_date(clvd):
            messagebox.showerror("エラー", "最終分娩日の形式が正しくありません (YYYY-MM-DD)")
            return
        
        # 更新データを準備
        cow_data = {
            'cow_id': cow_id,
            'jpn10': jpn10,
            'brd': brd,
            'bthd': bthd,
            'entr': entr,
            'lact': lact,
            'clvd': clvd,
            'rc': self.cow.get('rc'),  # 繁殖コードは変更しない（RuleEngineで管理）
            'pen': pen,
            'frm': self.cow.get('frm')
        }
        
        # 更新実行
        try:
            self.db.update_cow(self.cow_auto_id, cow_data)
            messagebox.showinfo("完了", "個体情報を更新しました")
            self.window.destroy()
            
            if self.on_saved:
                self.on_saved()
        except Exception as e:
            messagebox.showerror("エラー", f"更新に失敗しました:\n{e}")
    
    def _is_valid_date(self, date_str: str) -> bool:
        """日付形式の簡易チェック"""
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            return True
        except ValueError:
            return False
    
    def _on_cancel(self):
        """キャンセル"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()




