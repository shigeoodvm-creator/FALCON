"""
FALCON2 - イベント辞書編集ウィンドウ
event_dictionary.json のイベント定義を編集
"""

import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
from pathlib import Path
from typing import Dict, Any, Optional, Callable
import json
import logging
import re

logger = logging.getLogger(__name__)


class EventDictionaryEditWindow:
    """イベント辞書編集ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, event_number: str, 
                 event_data: Dict[str, Any],
                 event_dictionary_path: Path,
                 on_saved: Optional[Callable[[], None]] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            event_number: イベント番号（文字列）
            event_data: イベントデータ
            event_dictionary_path: event_dictionary.json のパス
            on_saved: 保存完了時のコールバック
        """
        self.parent = parent
        self.event_number = event_number
        self.original_event_data = event_data.copy()
        self.event_dictionary_path = event_dictionary_path
        self.on_saved = on_saved
        
        # ウィンドウを作成
        self.window = tk.Toplevel(parent)
        self.window.title(f"イベント辞書編集 - {event_number}")
        self.window.geometry("700x480")
        
        # 編集用のデータ（コピー）
        self.edited_data = {
            'alias': event_data.get('alias', ''),
            'name_jp': event_data.get('name_jp', ''),
            'category': event_data.get('category', ''),
            'input_code': event_data.get('input_code'),
            'input_fields': event_data.get('input_fields', []).copy() if event_data.get('input_fields') else [],
            'deprecated': event_data.get('deprecated', False),
            'display_color': event_data.get('display_color', ''),
            'outcome': event_data.get('outcome', '')  # 既存の値も保持
        }
        
        self._create_widgets()
        self._populate_fields()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # イベント番号（編集不可）
        info_frame = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text="イベント番号:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(info_frame, text=self.event_number, font=("", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # alias
        ttk.Label(info_frame, text="エイリアス:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.alias_entry = ttk.Entry(info_frame, width=30)
        self.alias_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # name_jp
        ttk.Label(info_frame, text="日本語名:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.name_jp_entry = ttk.Entry(info_frame, width=30)
        self.name_jp_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # category
        ttk.Label(info_frame, text="カテゴリ:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_combo = ttk.Combobox(
            info_frame,
            values=["REPRODUCTION", "PRODUCTION", "HEALTH", "MILK_TEST", "MANAGEMENT", "TASK"],
            width=27,
            state="readonly"
        )
        self.category_combo.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        # input_code
        ttk.Label(info_frame, text="入力コード:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.input_code_entry = ttk.Entry(info_frame, width=30)
        self.input_code_entry.grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        
        # deprecated
        self.deprecated_var = tk.BooleanVar()
        deprecated_check = ttk.Checkbutton(
            info_frame,
            text="非推奨（deprecated）",
            variable=self.deprecated_var
        )
        deprecated_check.grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        # display_color（表示色）
        ttk.Label(info_frame, text="表示色:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        color_frame = ttk.Frame(info_frame)
        color_frame.grid(row=6, column=1, sticky=tk.W, padx=5, pady=5)
        
        # HEXコード入力Entry
        self.display_color_entry = ttk.Entry(color_frame, width=15)
        self.display_color_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # カラーピッカーボタン
        color_picker_btn = ttk.Button(
            color_frame,
            text="色を選択",
            command=self._on_color_picker
        )
        color_picker_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 色プレビュー（小さなフレーム）
        self.color_preview = tk.Frame(color_frame, width=30, height=20, relief=tk.SUNKEN, borderwidth=2)
        self.color_preview.pack(side=tk.LEFT)
        self.color_preview.pack_propagate(False)
        
        # HEXコード入力時の検証とプレビュー更新
        self.display_color_entry.bind('<KeyRelease>', self._on_color_entry_changed)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
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
    
    def _populate_fields(self):
        """フィールドに値を設定"""
        self.alias_entry.insert(0, self.edited_data.get('alias', ''))
        self.name_jp_entry.insert(0, self.edited_data.get('name_jp', ''))
        
        category = self.edited_data.get('category', '')
        if category:
            self.category_combo.set(category)
        
        input_code = self.edited_data.get('input_code')
        if input_code is not None:
            self.input_code_entry.insert(0, str(input_code))
        
        self.deprecated_var.set(self.edited_data.get('deprecated', False))
        
        # display_colorを初期表示
        display_color = self.edited_data.get('display_color', '')
        if display_color:
            self.display_color_entry.insert(0, display_color)
            self._update_color_preview(display_color)
    
    def _on_color_picker(self):
        """カラーピッカーを開く"""
        # 現在のHEXコードを取得
        current_color = self.display_color_entry.get().strip()
        # デフォルト色を設定（現在の色があれば使用）
        initial_color = current_color if self._is_valid_hex_color(current_color) else None
        
        # カラーピッカーを開く
        color = colorchooser.askcolor(
            title="色を選択",
            color=initial_color
        )
        
        if color and color[1]:  # color[1]はHEXコード（例: "#ff0000"）
            hex_color = color[1]
            # Entryに設定
            self.display_color_entry.delete(0, tk.END)
            self.display_color_entry.insert(0, hex_color)
            # プレビューを更新
            self._update_color_preview(hex_color)
    
    def _on_color_entry_changed(self, event=None):
        """HEXコード入力時の検証とプレビュー更新"""
        hex_color = self.display_color_entry.get().strip()
        if hex_color:
            if self._is_valid_hex_color(hex_color):
                self._update_color_preview(hex_color)
            else:
                # 無効な色の場合はプレビューをクリア（グレーアウト）
                self.color_preview.config(bg="lightgray")
        else:
            # 空の場合はプレビューをクリア
            self.color_preview.config(bg="white")
    
    def _is_valid_hex_color(self, color_str: str) -> bool:
        """HEXカラーコードの形式を検証"""
        if not color_str:
            return False
        # #RRGGBB または #RGB 形式をチェック
        pattern = r'^#[0-9A-Fa-f]{6}$'
        return bool(re.match(pattern, color_str))
    
    def _update_color_preview(self, hex_color: str):
        """色プレビューを更新"""
        if self._is_valid_hex_color(hex_color):
            self.color_preview.config(bg=hex_color)
        else:
            self.color_preview.config(bg="white")
    
    def _on_save(self):
        """保存処理"""
        # 入力値を取得
        alias = self.alias_entry.get().strip()
        name_jp = self.name_jp_entry.get().strip()
        category = self.category_combo.get()
        input_code_str = self.input_code_entry.get().strip()
        deprecated = self.deprecated_var.get()
        
        # バリデーション
        if not alias:
            messagebox.showerror("エラー", "エイリアスを入力してください")
            return
        
        if not name_jp:
            messagebox.showerror("エラー", "日本語名を入力してください")
            return
        
        if not category:
            messagebox.showerror("エラー", "カテゴリを選択してください")
            return
        
        # input_codeの処理
        input_code = None
        if input_code_str:
            try:
                input_code = int(input_code_str)
            except ValueError:
                messagebox.showerror("エラー", "入力コードは整数で入力してください")
                return
        
        # 入力フィールドは開発側で管理するため、既存の値をそのまま維持する
        input_fields = self.original_event_data.get('input_fields', [])
        if isinstance(input_fields, list):
            input_fields = input_fields.copy()
        else:
            input_fields = []
        
        # display_colorを取得・検証
        display_color = self.display_color_entry.get().strip()
        if display_color:
            # HEXコードの形式を検証
            if not self._is_valid_hex_color(display_color):
                messagebox.showerror("エラー", "表示色は#RRGGBB形式（例: #cc0000）で入力してください")
                return
        
        # 編集データを構築
        updated_data = {
            'alias': alias,
            'name_jp': name_jp,
            'category': category,
            'input_code': input_code,
            'input_fields': input_fields,
            'deprecated': deprecated
        }
        
        # outcomeが既に存在する場合は保持
        if self.edited_data.get('outcome'):
            updated_data['outcome'] = self.edited_data.get('outcome')
        
        # display_colorの処理は、既存データとマージ後に実行
        
        # event_dictionary.json を読み込んで更新
        try:
            with open(self.event_dictionary_path, 'r', encoding='utf-8') as f:
                event_dictionary = json.load(f)
            
            # 既存のイベントデータを取得（存在しない場合は空dict）
            existing_data = event_dictionary.get(self.event_number, {})
            
            # 既存データをマージ（outcomeなどの他のフィールドを保持）
            merged_data = existing_data.copy()
            merged_data.update(updated_data)
            
            # display_colorの処理
            if display_color:
                # display_colorを設定
                merged_data['display_color'] = display_color
            elif 'display_color' in merged_data:
                # display_colorが空の場合は削除
                del merged_data['display_color']
            
            # イベントを更新
            event_dictionary[self.event_number] = merged_data
            
            # 保存（フォルダが存在しない場合は作成）
            if self.event_dictionary_path:
                self.event_dictionary_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.event_dictionary_path, 'w', encoding='utf-8') as f:
                json.dump(event_dictionary, f, ensure_ascii=False, indent=2)
            
            # ログに「event_dictionary updated: event_number=XXX」を出力
            logger.info(f"event_dictionary updated: event_number={self.event_number}")
            
            messagebox.showinfo("完了", "イベント辞書を保存しました")
            
            # コールバックを呼び出し
            if self.on_saved:
                self.on_saved()
            
            self.window.destroy()
            
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")
    
    def _on_cancel(self):
        """キャンセル"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

