"""
FALCON2 - 農場選択ウィンドウ
設計書 第14章 14.2 参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Callable

from ui.farm_create_window import FarmCreateWindow


class FarmSelectorWindow:
    """農場選択ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, on_farm_selected: Optional[Callable] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            on_farm_selected: 農場選択時のコールバック関数（farm_path を引数に取る）
        """
        self.parent = parent
        self.on_farm_selected = on_farm_selected
        self.selected_farm_path: Optional[Path] = None
        
        self.window = tk.Toplevel(parent)
        self.window.title("農場選択")
        self.window.geometry("400x300")
        self.window.transient(parent)
        self.window.grab_set()
        
        self._create_widgets()
        self._load_farms()
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = tk.Label(
            self.window,
            text="農場を選択してください",
            font=("", 12, "bold")
        )
        title_label.pack(pady=10)
        
        # 農場リスト
        list_frame = tk.Frame(self.window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.farm_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("", 11)
        )
        self.farm_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.farm_listbox.yview)
        
        # ダブルクリックで選択
        self.farm_listbox.bind("<Double-Button-1>", self._on_farm_double_click)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=10)
        
        open_button = tk.Button(
            button_frame,
            text="農場を開く",
            command=self._on_open_farm,
            width=12
        )
        open_button.pack(side=tk.LEFT, padx=5)
        
        create_button = tk.Button(
            button_frame,
            text="新規農場作成",
            command=self._on_create_farm,
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
    
    def _load_farms(self):
        """農場リストをロード"""
        farms_root = Path("C:/FARMS")
        
        if not farms_root.exists():
            farms_root.mkdir(parents=True, exist_ok=True)
            print(f"農場フォルダを作成: {farms_root}")
            # 空のリストを表示
            self.farm_listbox.delete(0, tk.END)
            self.farm_listbox.insert(tk.END, "(農場が見つかりません)")
            return
        
        # C:/FARMS 配下のフォルダを取得
        farms = []
        for item in farms_root.iterdir():
            if item.is_dir():
                db_path = item / "farm.db"
                if db_path.exists():
                    farms.append(item.name)
                    print(f"農場を発見: {item.name}")
                else:
                    print(f"farm.dbが見つかりません: {item}")
        
        # リストボックスに追加
        self.farm_listbox.delete(0, tk.END)
        if farms:
            for farm in sorted(farms):
                self.farm_listbox.insert(tk.END, farm)
        else:
            # 農場が見つからない場合のメッセージ
            self.farm_listbox.insert(tk.END, "(農場が見つかりません)")
            print("農場が見つかりません。CSVファイルからデモ農場を作成してください:")
            print("  python main.py --demo-csv <CSVファイルのパス>")
    
    def _on_farm_double_click(self, event):
        """農場のダブルクリック処理"""
        self._on_open_farm()
    
    def _on_open_farm(self):
        """農場を開く"""
        selection = self.farm_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "農場を選択してください")
            return
        
        farm_name = self.farm_listbox.get(selection[0])
        
        # "(農場が見つかりません)" が選択された場合は処理しない
        if farm_name == "(農場が見つかりません)":
            messagebox.showinfo(
                "情報",
                "デモ農場を作成するには、以下のコマンドを実行してください:\n\n"
                "python main.py --demo-csv <CSVファイルのパス>"
            )
            return
        
        farm_path = Path("C:/FARMS") / farm_name
        
        if not (farm_path / "farm.db").exists():
            messagebox.showerror("エラー", f"農場データが見つかりません: {farm_path}")
            return
        
        self.selected_farm_path = farm_path
        self.window.destroy()
        
        if self.on_farm_selected:
            self.on_farm_selected(farm_path)
    
    def _on_create_farm(self):
        """新規農場作成"""
        def on_farm_created(farm_path: Path):
            """農場作成時のコールバック"""
            # 農場リストを更新
            self._load_farms()
            
            # 作成された農場を選択状態にする
            farm_name = farm_path.name
            for i in range(self.farm_listbox.size()):
                if self.farm_listbox.get(i) == farm_name:
                    self.farm_listbox.selection_clear(0, tk.END)
                    self.farm_listbox.selection_set(i)
                    self.farm_listbox.see(i)
                    break
        
        # 新規農場作成ウィンドウを表示
        create_window = FarmCreateWindow(self.window, on_farm_created=on_farm_created)
        created_path = create_window.show()
        
        # 農場が作成された場合は、その農場を選択して開く
        if created_path:
            self.selected_farm_path = created_path
            self.window.destroy()
            
            if self.on_farm_selected:
                self.on_farm_selected(created_path)
    
    def _on_cancel(self):
        """キャンセル"""
        self.window.destroy()
        self.parent.quit()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
        return self.selected_farm_path

