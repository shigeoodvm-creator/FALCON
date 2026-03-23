"""
FALCON2 - 農場選択ウィンドウ
設計書 第14章 14.2 参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Callable

from ui.farm_create_window import FarmCreateWindow
from constants import FARMS_ROOT

# デザイン定数
_BG = "#f5f5f5"
_CARD_BG = "#ffffff"
_CARD_BORDER = "#e0e0e0"
_ACCENT = "#3949ab"
_ACCENT_HOVER = "#303f9f"
_TEXT_PRIMARY = "#263238"
_TEXT_SECONDARY = "#607d8b"
_SELECT_BG = "#e8eaf6"
_FONT = "Meiryo UI"


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
        self.window.geometry("480x460")
        self.window.minsize(420, 400)
        self.window.configure(bg=_BG)
        
        self._create_widgets()
        self._load_farms()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        # 最前面に表示（ルートが非表示でも、他ウィンドウがある場合に備える）
        self.window.lift()
        self.window.focus_force()

    def _create_widgets(self):
        """ウィジェットを作成（視覚的階層・カード風リスト・プライマリボタン）"""
        main_container = tk.Frame(self.window, bg=_BG, padx=28, pady=24)
        main_container.pack(fill=tk.BOTH, expand=True)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # ヘッダー：タイトル＋説明（余白をやや広めに）
        header = tk.Frame(main_container, bg=_BG)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 20))
        title_frame = tk.Frame(header, bg=_BG)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="農場選択", font=(_FONT, 17, "bold"), bg=_BG, fg=_TEXT_PRIMARY).pack(anchor=tk.W)
        tk.Label(title_frame, text="農場を選択してください", font=(_FONT, 10), bg=_BG, fg=_TEXT_SECONDARY).pack(anchor=tk.W)
        
        # 農場リスト用カード（白背景・枠線で囲む）
        list_card = tk.Frame(main_container, bg=_CARD_BORDER, padx=1, pady=1)
        list_card.grid(row=1, column=0, sticky=tk.NSEW, pady=(0, 20))
        list_card.grid_columnconfigure(0, weight=1)
        list_card.grid_rowconfigure(0, weight=1)
        
        list_inner = tk.Frame(list_card, bg=_CARD_BG, padx=0, pady=0)
        list_inner.grid(row=0, column=0, sticky=tk.NSEW)
        list_inner.grid_columnconfigure(0, weight=1)
        list_inner.grid_rowconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(list_inner)
        scrollbar.grid(row=0, column=1, sticky=tk.NS, pady=2)
        
        self.farm_listbox = tk.Listbox(
            list_inner,
            yscrollcommand=scrollbar.set,
            font=(_FONT, 11),
            selectbackground=_SELECT_BG,
            selectforeground=_TEXT_PRIMARY,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
            bg=_CARD_BG,
            fg=_TEXT_PRIMARY,
            activestyle=tk.NONE
        )
        self.farm_listbox.grid(row=0, column=0, sticky=tk.NSEW, padx=(8, 2), pady=6)
        scrollbar.config(command=self.farm_listbox.yview)
        
        self.farm_listbox.bind("<Double-Button-1>", self._on_farm_double_click)
        
        # フッター：区切り線＋ボタン（プライマリを強調）
        footer = tk.Frame(main_container, bg=_BG)
        footer.grid(row=2, column=0, sticky=tk.EW, pady=(4, 0))
        footer.grid_columnconfigure(0, weight=1)
        
        sep = tk.Frame(footer, height=1, bg=_CARD_BORDER)
        sep.grid(row=0, column=0, sticky=tk.EW, pady=(0, 16))
        
        btn_frame = tk.Frame(footer, bg=_BG)
        btn_frame.grid(row=1, column=0, sticky=tk.E)
        
        # キャンセル・新規作成は ttk、メインは tk でプライマリ色
        ttk.Button(
            btn_frame,
            text="キャンセル",
            command=self._on_cancel,
            width=11
        ).pack(side=tk.RIGHT, padx=(8, 0))
        
        ttk.Button(
            btn_frame,
            text="新規農場作成",
            command=self._on_create_farm,
            width=12
        ).pack(side=tk.RIGHT, padx=(8, 0))
        
        open_btn = tk.Button(
            btn_frame,
            text="農場を開く",
            command=self._on_open_farm,
            font=(_FONT, 10),
            bg=_ACCENT,
            fg="white",
            activebackground=_ACCENT_HOVER,
            activeforeground="white",
            relief=tk.FLAT,
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2"
        )
        open_btn.pack(side=tk.RIGHT)
        open_btn.bind("<Enter>", lambda e: open_btn.configure(bg=_ACCENT_HOVER))
        open_btn.bind("<Leave>", lambda e: open_btn.configure(bg=_ACCENT))
    
    def _load_farms(self):
        """農場リストをロード"""
        farms_root = FARMS_ROOT

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
        
        farm_path = FARMS_ROOT / farm_name
        
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

