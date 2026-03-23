"""
FALCON2 - 新規農場作成ウィンドウ
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from typing import Optional, Callable

from modules.farm_creator import FarmCreator
from constants import FARMS_ROOT


class FarmCreateWindow:
    """新規農場作成ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, on_farm_created: Optional[Callable] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            on_farm_created: 農場作成時のコールバック関数（farm_path を引数に取る）
        """
        self.parent = parent
        self.on_farm_created = on_farm_created
        self.created_farm_path: Optional[Path] = None
        
        self.window = tk.Toplevel(parent)
        self.window.title("新規農場作成")
        self.window.geometry("500x400")
        
        # 変数
        self.farm_name_var = tk.StringVar()
        self.csv_path_var = tk.StringVar()
        self.template_farm_var = tk.StringVar()
        
        self._create_widgets()
        self._load_template_farms()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = tk.Label(
            self.window,
            text="新規農場を作成",
            font=("", 12, "bold")
        )
        title_label.pack(pady=10)
        
        # メインフレーム
        main_frame = tk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 農場名
        name_frame = tk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(name_frame, text="農場名:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        name_entry = tk.Entry(name_frame, textvariable=self.farm_name_var, width=30)
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # テンプレート農場（オプション）
        template_frame = tk.Frame(main_frame)
        template_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(template_frame, text="テンプレート農場:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.template_farm_var,
            width=27,
            state="readonly"
        )
        template_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.template_combo = template_combo
        
        # CSVファイル選択
        csv_frame = tk.Frame(main_frame)
        csv_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(csv_frame, text="乳検速報CSV:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        csv_entry = tk.Entry(csv_frame, textvariable=self.csv_path_var, width=20)
        csv_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        csv_browse_btn = tk.Button(
            csv_frame,
            text="参照...",
            command=self._browse_csv,
            width=8
        )
        csv_browse_btn.pack(side=tk.LEFT)
        
        # 説明テキスト
        info_text = tk.Text(main_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        info_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        info_content = """【新規農場作成について】

・農場名を入力してください
・テンプレート農場を選択すると、設定を引き継げます（オプション）
・乳検速報CSVを指定すると、CSVから個体データとイベントを自動登録します

【CSV取り込み時の処理】
・CSVに含まれるすべての個体を登録
・産次・分娩月日はCSVの値をそのまま使用
・分娩イベントと乳検イベントを自動作成
・状態を自動計算して反映"""
        
        info_text.config(state=tk.NORMAL)
        info_text.insert(tk.END, info_content)
        info_text.config(state=tk.DISABLED)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=10)
        
        create_button = tk.Button(
            button_frame,
            text="作成",
            command=self._on_create,
            width=12
        )
        create_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = tk.Button(
            button_frame,
            text="キャンセル",
            command=self._on_cancel,
            width=12
        )
        cancel_button.pack(side=tk.LEFT, padx=5)
    
    def _load_template_farms(self):
        """テンプレート農場リストをロード"""
        farms_root = FARMS_ROOT

        if not farms_root.exists():
            self.template_combo['values'] = []
            return

        # FARMS_ROOT 配下のフォルダを取得
        farms = []
        for item in farms_root.iterdir():
            if item.is_dir():
                db_path = item / "farm.db"
                if db_path.exists():
                    farms.append(item.name)
        
        # コンボボックスに追加（先頭に「なし」を追加）
        values = ["（なし）"] + sorted(farms)
        self.template_combo['values'] = values
        self.template_combo.current(0)  # 「なし」を選択
    
    def _browse_csv(self):
        """CSVファイルを選択"""
        file_path = filedialog.askopenfilename(
            title="乳検速報CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        
        if file_path:
            self.csv_path_var.set(file_path)
    
    def _on_create(self):
        """農場を作成"""
        farm_name = self.farm_name_var.get().strip()
        
        if not farm_name:
            messagebox.showerror("エラー", "農場名を入力してください")
            return
        
        # 農場パスを決定
        farm_path = FARMS_ROOT / farm_name
        
        # 既存チェック
        if farm_path.exists() and (farm_path / "farm.db").exists():
            if not messagebox.askyesno(
                "確認",
                f"農場「{farm_name}」は既に存在します。\n上書きしますか？"
            ):
                return
        
        # CSVパス
        csv_path = None
        csv_path_str = self.csv_path_var.get().strip()
        if csv_path_str:
            csv_path = Path(csv_path_str)
            if not csv_path.exists():
                messagebox.showerror("エラー", f"CSVファイルが見つかりません: {csv_path}")
                return
        
        # テンプレート農場パス
        template_farm_path = None
        template_name = self.template_farm_var.get()
        if template_name and template_name != "（なし）":
            template_farm_path = FARMS_ROOT / template_name
            if not (template_farm_path / "farm.db").exists():
                messagebox.showerror("エラー", f"テンプレート農場が見つかりません: {template_farm_path}")
                return
        
        # 作成処理
        try:
            creator = FarmCreator(farm_path)
            creator.create_farm(
                farm_name=farm_name,
                csv_path=csv_path,
                template_farm_path=template_farm_path
            )
            
            self.created_farm_path = farm_path
            messagebox.showinfo("完了", f"農場「{farm_name}」を作成しました")
            
            self.window.destroy()
            
            if self.on_farm_created:
                self.on_farm_created(farm_path)
        
        except Exception as e:
            messagebox.showerror("エラー", f"農場の作成に失敗しました:\n{e}")
            import traceback
            traceback.print_exc()
    
    def _on_cancel(self):
        """キャンセル"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
        return self.created_farm_path









































