"""
FALCON2 - イベント辞書一覧ウィンドウ
登録されているイベントの一覧を表示
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Optional, Dict, Any, Tuple
import json
import logging

logger = logging.getLogger(__name__)


class EventDictionaryWindow:
    """イベント辞書一覧ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            event_dictionary_path: event_dictionary.json のパス
        """
        self.parent = parent
        self.event_dict_path = event_dictionary_path
        self.event_dictionary: Dict[str, Any] = {}
        self.selected_item_id: Optional[str] = None  # 右クリックで選択されたアイテムID
        
        # ウィンドウを作成
        self.window = tk.Toplevel(parent)
        self.window.title("イベント辞書")
        self.window.geometry("800x600")
        self.window.transient(parent)
        
        self._load_event_dictionary()
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
        else:
            # 農場フォルダ側の辞書が存在しない場合は空辞書
            # （同期処理で作成されるはずだが、念のため）
            print(f"警告: event_dictionary.json が見つかりません: {self.event_dict_path}")
            self.event_dictionary = {}
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.window,
            text="イベント辞書一覧",
            font=("", 14, "bold")
        )
        title_label.pack(pady=10)
        
        # 説明
        desc_label = ttk.Label(
            self.window,
            text="登録されているイベントの一覧です",
            font=("", 9)
        )
        desc_label.pack(pady=(0, 10))
        
        # テーブルフレーム
        table_frame = ttk.Frame(self.window)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # ツリービュー（テーブル表示）
        columns = ("event_number", "alias", "name_jp", "category")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        
        # カラムヘッダー
        self.tree.heading("event_number", text="イベント番号")
        self.tree.heading("alias", text="エイリアス")
        self.tree.heading("name_jp", text="日本語名")
        self.tree.heading("category", text="カテゴリ")
        
        # カラム幅
        self.tree.column("event_number", width=120, anchor=tk.CENTER)
        self.tree.column("alias", width=150, anchor=tk.W)
        self.tree.column("name_jp", width=250, anchor=tk.W)
        self.tree.column("category", width=150, anchor=tk.W)
        
        # スクロールバー
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # 配置
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # データを挿入
        self._populate_tree()
        
        # 詳細表示フレーム
        detail_frame = ttk.LabelFrame(self.window, text="詳細情報", padding=10)
        detail_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.detail_text = tk.Text(detail_frame, height=6, wrap=tk.WORD, state=tk.DISABLED)
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        
        # 選択イベント
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        
        # 右クリックメニュー
        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(label="編集", command=self._on_edit_selected)
        self.context_menu.add_command(label="削除", command=self._on_delete_selected)
        
        # 右クリックイベント
        self.tree.bind('<Button-3>', self._on_right_click)
        
        # デバッグ情報フレーム（画面下部）
        debug_frame = ttk.LabelFrame(self.window, text="デバッグ情報", padding=5)
        debug_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 読み込み元JSONのフルパス
        dict_path_str = str(self.event_dict_path) if self.event_dict_path else "(未指定)"
        path_label = ttk.Label(
            debug_frame,
            text=f"読み込み元: {dict_path_str}",
            font=("", 8),
            foreground="gray"
        )
        path_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # 表示中イベント数（deprecatedを除く）
        visible_count = sum(
            1 for event_data in self.event_dictionary.values()
            if not event_data.get("deprecated", False)
        )
        total_count = len(self.event_dictionary)
        count_label = ttk.Label(
            debug_frame,
            text=f"表示中イベント数: {visible_count} / 総イベント数: {total_count}",
            font=("", 8),
            foreground="gray"
        )
        count_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=10)
        
        close_button = ttk.Button(
            button_frame,
            text="閉じる",
            command=self.window.destroy,
            width=15
        )
        close_button.pack(padx=5)
    
    def _populate_tree(self):
        """ツリービューにデータを挿入"""
        # イベント番号でソート
        sorted_events = sorted(
            self.event_dictionary.items(),
            key=lambda x: int(x[0]) if x[0].isdigit() else 9999
        )
        
        for event_number, event_data in sorted_events:
            # 非推奨のイベントはスキップ（必要に応じて表示も可）
            if event_data.get("deprecated", False):
                continue
            
            alias = event_data.get("alias", "")
            name_jp = event_data.get("name_jp", "")
            category = event_data.get("category", "")
            
            # ツリーに挿入
            self.tree.insert(
                "",
                tk.END,
                values=(event_number, alias, name_jp, category),
                tags=(event_number,)
            )
    
    def _on_select(self, event):
        """イベント選択時の処理"""
        selection = self.tree.selection()
        if not selection:
            self.selected_item_id = None
            return
        
        # 選択されたアイテムIDを保存
        self.selected_item_id = selection[0]
        
        item = self.tree.item(selection[0])
        event_number = item['tags'][0] if item['tags'] else None
        
        if event_number and event_number in self.event_dictionary:
            event_data = self.event_dictionary[event_number]
            self._show_detail(event_number, event_data)
    
    def _on_right_click(self, event):
        """右クリック時の処理（コンテキストメニューを表示）"""
        # event.y を使って Treeview.identify_row(event.y) を取得
        row_id = self.tree.identify_row(event.y)
        if row_id:
            # 該当行があれば、その行を selection_set する
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
            # 選択されたアイテムIDを保存
            self.selected_item_id = row_id
            # コンテキストメニューを表示
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
        else:
            # アイテム以外をクリックした場合は選択をクリア
            self.selected_item_id = None
    
    def _get_selected_event(self) -> Tuple[Optional[str], Optional[Dict[str, Any]]]:
        """選択されたイベントを取得"""
        # tree.selection() が空でないことを確認
        selection = self.tree.selection()
        if not selection:
            return None, None
        
        # 選択された行の values から event_number を取得
        # values の想定: (values[0] = event_number)
        item = self.tree.item(selection[0])
        values = item.get('values', [])
        
        if not values or len(values) == 0:
            return None, None
        
        # values[0] から event_number を取得
        event_number = str(values[0])
        
        # event_dictionary[event_number] を取得する
        if event_number in self.event_dictionary:
            event_data = self.event_dictionary[event_number]
            return event_number, event_data
        
        return None, None
    
    def _on_edit_selected(self):
        """選択されたイベントを編集"""
        event_number, event_data = self._get_selected_event()
        if not event_number or not event_data:
            messagebox.showwarning("警告", "イベントを選択してください")
            return
        
        # deprecated=true のイベントは編集不可
        if event_data.get('deprecated', False):
            messagebox.showwarning(
                "警告",
                f"イベント番号 {event_number} は非推奨（deprecated）のため編集できません。"
            )
            return
        
        # 編集後、選択状態をクリア
        self.selected_item_id = None
        self._open_edit_window(event_number, event_data)
    
    def _on_delete_selected(self):
        """選択されたイベントを削除"""
        event_number, event_data = self._get_selected_event()
        if not event_number or not event_data:
            messagebox.showwarning("警告", "イベントを選択してください")
            return
        
        # 確認ダイアログ
        name_jp = event_data.get('name_jp', event_number)
        result = messagebox.askyesno(
            "確認",
            f"イベント「{name_jp} (番号: {event_number})」を削除しますか？\n\n"
            "削除すると、このイベントは非推奨（deprecated）としてマークされます。"
        )
        
        if result:
            # 削除後、選択状態をクリア
            self.selected_item_id = None
            self._delete_event(event_number)
    
    def _open_edit_window(self, event_number: str, event_data: Dict[str, Any]):
        """編集ウィンドウを開く"""
        if not self.event_dict_path:
            messagebox.showerror("エラー", "event_dictionary.json のパスが設定されていません")
            return
        
        from ui.event_dictionary_edit_window import EventDictionaryEditWindow
        
        def on_saved():
            """保存後のコールバック（辞書を再読み込みして表示を更新）"""
            self._load_event_dictionary()
            # ツリービューをクリアして再構築
            for item in self.tree.get_children():
                self.tree.delete(item)
            self._populate_tree()
        
        edit_window = EventDictionaryEditWindow(
            parent=self.window,
            event_number=event_number,
            event_data=event_data,
            event_dictionary_path=self.event_dict_path,
            on_saved=on_saved
        )
        edit_window.show()
    
    def _delete_event(self, event_number: str):
        """イベントを削除（deprecatedフラグをtrueに設定）"""
        if not self.event_dict_path:
            messagebox.showerror("エラー", "event_dictionary.json のパスが設定されていません")
            return
        
        try:
            # event_dictionary.json を読み込む
            with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                event_dictionary = json.load(f)
            
            # deprecatedフラグを設定
            if event_number in event_dictionary:
                event_dictionary[event_number]['deprecated'] = True
                
                # 保存
                with open(self.event_dict_path, 'w', encoding='utf-8') as f:
                    json.dump(event_dictionary, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("完了", f"イベント番号 {event_number} を削除しました（非推奨としてマーク）")
                
                # 表示を更新
                self._load_event_dictionary()
                # ツリービューをクリアして再構築
                for item in self.tree.get_children():
                    self.tree.delete(item)
                self._populate_tree()
            else:
                messagebox.showerror("エラー", f"イベント番号 {event_number} が見つかりません")
                
        except Exception as e:
            messagebox.showerror("エラー", f"削除に失敗しました: {e}")
    
    def _show_detail(self, event_number: str, event_data: Dict[str, Any]):
        """詳細情報を表示"""
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete(1.0, tk.END)
        
        detail_lines = [
            f"イベント番号: {event_number}",
            f"エイリアス: {event_data.get('alias', '')}",
            f"日本語名: {event_data.get('name_jp', '')}",
            f"カテゴリ: {event_data.get('category', '')}",
        ]
        
        # 入力コードがある場合
        if 'input_code' in event_data:
            detail_lines.append(f"入力コード: {event_data['input_code']}")
        
        # 入力フィールドがある場合
        if 'input_fields' in event_data and event_data['input_fields']:
            # 入力項目のラベルをカンマ区切りで取得
            field_labels = []
            for field in event_data['input_fields']:
                label = field.get('label', field.get('key', ''))
                field_labels.append(label)
            
            if field_labels:
                detail_lines.append(f"\n入力項目: {', '.join(field_labels)}")
            
            # 詳細情報も表示（オプション）
            detail_lines.append("\n詳細:")
            for field in event_data['input_fields']:
                key = field.get('key', '')
                datatype = field.get('datatype', '')
                label = field.get('label', key)
                detail_lines.append(f"  - {label} ({key}): {datatype}")
        else:
            detail_lines.append("\n入力項目: なし")
        
        detail_text = "\n".join(detail_lines)
        self.detail_text.insert(1.0, detail_text)
        self.detail_text.config(state=tk.DISABLED)
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()
        self.window.grab_set()

