"""
FALCON2 - 授精設定ウインドウ
授精師コード・授精種類コードの管理
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Optional
import json
import logging


class InseminationSettingsWindow:
    """授精設定ウインドウ"""
    
    def __init__(self, parent: tk.Tk, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（FarmSettingsWindow）
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_file = self.farm_path / "insemination_settings.json"
        
        # 設定データ
        self.technicians: Dict[str, str] = {}  # {"1": "園田", "2": "NOSAI北見"}
        self.insemination_types: Dict[str, str] = {}  # {"1": "自然発情", "2": "CIDR"}
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("授精設定")
        self.window.geometry("700x600")
        self.window.transient(parent)
        self.window.grab_set()
        
        # 設定をロード
        self._load_settings()
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _load_settings(self):
        """設定をロード"""
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.technicians = data.get('technicians', {})
                    self.insemination_types = data.get('insemination_types', {})
            except Exception as e:
                logging.error(f"授精設定ファイル読み込みエラー: {e}")
                messagebox.showerror("エラー", f"設定ファイルの読み込みに失敗しました: {e}")
                self.technicians = {}
                self.insemination_types = {}
        else:
            # ファイルが存在しない場合は空の初期データ
            self.technicians = {}
            self.insemination_types = {}
    
    def _save_settings(self):
        """設定を保存"""
        data = {
            "technicians": self.technicians,
            "insemination_types": self.insemination_types
        }
        
        try:
            self.farm_path.mkdir(parents=True, exist_ok=True)
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logging.info("授精設定を保存しました")
            return True
        except Exception as e:
            logging.error(f"授精設定ファイル保存エラー: {e}")
            messagebox.showerror("エラー", f"設定ファイルの保存に失敗しました: {e}")
            return False
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 授精師コード設定 ==========
        technician_frame = ttk.LabelFrame(main_frame, text="授精師コード設定", padding=10)
        technician_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview
        technician_tree_frame = ttk.Frame(technician_frame)
        technician_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        technician_scrollbar = ttk.Scrollbar(technician_tree_frame)
        technician_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.technician_tree = ttk.Treeview(
            technician_tree_frame,
            columns=("code", "name"),
            show="headings",
            yscrollcommand=technician_scrollbar.set,
            height=8
        )
        self.technician_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        technician_scrollbar.config(command=self.technician_tree.yview)
        
        self.technician_tree.heading("code", text="コード")
        self.technician_tree.heading("name", text="表示名")
        self.technician_tree.column("code", width=100)
        self.technician_tree.column("name", width=200)
        
        # 右クリックメニュー
        technician_menu = tk.Menu(self.window, tearoff=0)
        technician_menu.add_command(label="編集", command=self._edit_technician)
        technician_menu.add_command(label="削除", command=self._delete_technician)
        self.technician_tree.bind("<Button-3>", lambda e: self._show_context_menu(e, technician_menu))
        
        # ボタンフレーム
        technician_button_frame = ttk.Frame(technician_frame)
        technician_button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(
            technician_button_frame,
            text="追加",
            command=self._add_technician
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            technician_button_frame,
            text="編集",
            command=self._edit_technician
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            technician_button_frame,
            text="削除",
            command=self._delete_technician
        ).pack(side=tk.LEFT, padx=5)
        
        # ========== 授精種類コード設定 ==========
        type_frame = ttk.LabelFrame(main_frame, text="授精種類コード設定", padding=10)
        type_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview
        type_tree_frame = ttk.Frame(type_frame)
        type_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        type_scrollbar = ttk.Scrollbar(type_tree_frame)
        type_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.type_tree = ttk.Treeview(
            type_tree_frame,
            columns=("code", "name"),
            show="headings",
            yscrollcommand=type_scrollbar.set,
            height=8
        )
        self.type_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        type_scrollbar.config(command=self.type_tree.yview)
        
        self.type_tree.heading("code", text="コード")
        self.type_tree.heading("name", text="表示名")
        self.type_tree.column("code", width=100)
        self.type_tree.column("name", width=200)
        
        # 右クリックメニュー
        type_menu = tk.Menu(self.window, tearoff=0)
        type_menu.add_command(label="編集", command=self._edit_type)
        type_menu.add_command(label="削除", command=self._delete_type)
        self.type_tree.bind("<Button-3>", lambda e: self._show_context_menu(e, type_menu))
        
        # ボタンフレーム
        type_button_frame = ttk.Frame(type_frame)
        type_button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(
            type_button_frame,
            text="追加",
            command=self._add_type
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            type_button_frame,
            text="編集",
            command=self._edit_type
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            type_button_frame,
            text="削除",
            command=self._delete_type
        ).pack(side=tk.LEFT, padx=5)
        
        # ========== OK / キャンセル ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame,
            text="OK",
            command=self._on_ok
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame,
            text="キャンセル",
            command=self._on_cancel
        ).pack(side=tk.RIGHT, padx=5)
        
        # データを表示
        self._refresh_trees()
    
    def _refresh_trees(self):
        """Treeviewを更新"""
        # 授精師コード
        for item in self.technician_tree.get_children():
            self.technician_tree.delete(item)
        
        for code, name in sorted(self.technicians.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
            self.technician_tree.insert("", tk.END, values=(code, name))
        
        # 授精種類コード
        for item in self.type_tree.get_children():
            self.type_tree.delete(item)
        
        for code, name in sorted(self.insemination_types.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
            self.type_tree.insert("", tk.END, values=(code, name))
    
    def _show_context_menu(self, event, menu: tk.Menu):
        """右クリックメニューを表示"""
        item = event.widget.selection()[0] if event.widget.selection() else None
        if item:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
    
    def _add_technician(self):
        """授精師コードを追加"""
        dialog = tk.Toplevel(self.window)
        dialog.title("授精師コード追加")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        ttk.Label(dialog, text="コード:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        code_entry = ttk.Entry(dialog, width=30)
        code_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="表示名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def on_ok():
            code = code_entry.get().strip()
            name = name_entry.get().strip()
            
            if not code:
                messagebox.showwarning("警告", "コードを入力してください")
                return
            
            if not name:
                messagebox.showwarning("警告", "表示名を入力してください")
                return
            
            if code in self.technicians:
                messagebox.showwarning("警告", "このコードは既に登録されています")
                return
            
            self.technicians[code] = name
            self._refresh_trees()
            dialog.destroy()
        
        ttk.Button(dialog, text="OK", command=on_ok).grid(row=2, column=0, columnspan=2, pady=10)
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _edit_technician(self):
        """授精師コードを編集"""
        selection = self.technician_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "編集する項目を選択してください")
            return
        
        item = self.technician_tree.item(selection[0])
        code = item['values'][0]
        current_name = item['values'][1]
        
        dialog = tk.Toplevel(self.window)
        dialog.title("授精師コード編集")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        ttk.Label(dialog, text="コード:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        code_label = ttk.Label(dialog, text=code)
        code_label.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="表示名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.insert(0, current_name)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def on_ok():
            name = name_entry.get().strip()
            
            if not name:
                messagebox.showwarning("警告", "表示名を入力してください")
                return
            
            self.technicians[code] = name
            self._refresh_trees()
            dialog.destroy()
        
        ttk.Button(dialog, text="OK", command=on_ok).grid(row=2, column=0, columnspan=2, pady=10)
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _delete_technician(self):
        """授精師コードを削除"""
        selection = self.technician_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "削除する項目を選択してください")
            return
        
        item = self.technician_tree.item(selection[0])
        code = item['values'][0]
        name = item['values'][1]
        
        if messagebox.askyesno("確認", f"授精師コード「{code}: {name}」を削除しますか？"):
            del self.technicians[code]
            self._refresh_trees()
    
    def _add_type(self):
        """授精種類コードを追加"""
        dialog = tk.Toplevel(self.window)
        dialog.title("授精種類コード追加")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        ttk.Label(dialog, text="コード:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        code_entry = ttk.Entry(dialog, width=30)
        code_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(dialog, text="表示名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def on_ok():
            code = code_entry.get().strip()
            name = name_entry.get().strip()
            
            if not code:
                messagebox.showwarning("警告", "コードを入力してください")
                return
            
            if not name:
                messagebox.showwarning("警告", "表示名を入力してください")
                return
            
            if code in self.insemination_types:
                messagebox.showwarning("警告", "このコードは既に登録されています")
                return
            
            self.insemination_types[code] = name
            self._refresh_trees()
            dialog.destroy()
        
        ttk.Button(dialog, text="OK", command=on_ok).grid(row=2, column=0, columnspan=2, pady=10)
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _edit_type(self):
        """授精種類コードを編集"""
        selection = self.type_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "編集する項目を選択してください")
            return
        
        item = self.type_tree.item(selection[0])
        code = item['values'][0]
        current_name = item['values'][1]
        
        dialog = tk.Toplevel(self.window)
        dialog.title("授精種類コード編集")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        ttk.Label(dialog, text="コード:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        code_label = ttk.Label(dialog, text=code)
        code_label.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(dialog, text="表示名:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.insert(0, current_name)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def on_ok():
            name = name_entry.get().strip()
            
            if not name:
                messagebox.showwarning("警告", "表示名を入力してください")
                return
            
            self.insemination_types[code] = name
            self._refresh_trees()
            dialog.destroy()
        
        ttk.Button(dialog, text="OK", command=on_ok).grid(row=2, column=0, columnspan=2, pady=10)
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _delete_type(self):
        """授精種類コードを削除"""
        selection = self.type_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "削除する項目を選択してください")
            return
        
        item = self.type_tree.item(selection[0])
        code = item['values'][0]
        name = item['values'][1]
        
        if messagebox.askyesno("確認", f"授精種類コード「{code}: {name}」を削除しますか？"):
            del self.insemination_types[code]
            self._refresh_trees()
    
    def _on_ok(self):
        """OKボタンクリック時の処理"""
        if self._save_settings():
            messagebox.showinfo("完了", "設定を保存しました")
            self.window.destroy()
    
    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()



