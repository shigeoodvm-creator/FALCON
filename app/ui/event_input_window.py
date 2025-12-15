"""
FALCON2 - EventInputWindow
イベント入力ウィンドウ（動的フォーム生成）
設計書 第18章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable
from pathlib import Path
import json
from datetime import datetime
import re
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


class EventInputWindow:
    """イベント入力ウィンドウ"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, 
                 rule_engine: RuleEngine,
                 cow_auto_id: Optional[int],
                 event_dictionary_path: Path,
                 on_saved: Optional[Callable[[int], None]] = None,
                 farm_path: Optional[Path] = None,
                 edit_event_id: Optional[int] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            db_handler: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
            cow_auto_id: 対象牛の auto_id（Noneの場合は後で入力欄から解決）
            event_dictionary_path: event_dictionary.json のパス
            on_saved: 保存完了時のコールバック（cow_auto_id を引数に取る）
            farm_path: 農場フォルダのパス（enum型参照用、将来的に使用）
            edit_event_id: 編集モードの場合のイベントID（Noneの場合は新規作成）
        """
        self.db = db_handler
        self.rule_engine = rule_engine
        self.cow_auto_id = cow_auto_id
        self.event_dict_path = event_dictionary_path
        self.on_saved = on_saved
        self.farm_path = farm_path
        self.edit_event_id = edit_event_id
        self.edit_event_data = None
        
        # event_dictionary を読み込む
        self.event_dictionary: Dict[str, Dict[str, Any]] = {}
        self._load_event_dictionary()
        
        # 現在選択中のイベント
        self.selected_event: Optional[Dict[str, Any]] = None
        self.selected_event_number: Optional[int] = None
        
        # 動的フォームのウィジェット
        self.field_widgets: Dict[str, tk.Widget] = {}
        
        # 編集モードの場合、既存イベントデータを読み込む
        if self.edit_event_id:
            self.edit_event_data = self.db.get_event_by_id(self.edit_event_id)
            if not self.edit_event_data:
                messagebox.showerror("エラー", "編集対象のイベントが見つかりません")
                return
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("イベント編集" if self.edit_event_id else "イベント入力")
        self.window.geometry("500x600")
        self.window.transient(parent)
        self.window.grab_set()
        
        self._create_widgets()
        
        # 編集モードの場合、既存データをフォームに設定
        if self.edit_event_id and self.edit_event_data:
            self._load_event_data()
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                messagebox.showerror("エラー", f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
    
    def _load_insemination_settings(self) -> Optional[Dict[str, Any]]:
        """授精設定をロード"""
        if not self.farm_path:
            return None
        
        settings_file = self.farm_path / "insemination_settings.json"
        if not settings_file.exists():
            return None
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"授精設定ファイル読み込みエラー: {e}")
            return None
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== イベント選択 ==========
        event_frame = ttk.LabelFrame(main_frame, text="イベント選択", padding=10)
        event_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(event_frame, text="イベント:").pack(side=tk.LEFT, padx=5)
        
        # イベント一覧を準備（deprecated でないもの）
        event_list = []
        self.event_map = {}  # name_jp -> event_number
        
        for event_num_str, event_data in self.event_dictionary.items():
            if event_data.get('deprecated', False):
                continue
            name_jp = event_data.get('name_jp', f'イベント{event_num_str}')
            event_list.append(name_jp)
            self.event_map[name_jp] = int(event_num_str)
        
        # Combobox
        self.event_combo = ttk.Combobox(
            event_frame,
            values=event_list,
            state='readonly',
            width=30
        )
        self.event_combo.pack(side=tk.LEFT, padx=5)
        self.event_combo.bind('<<ComboboxSelected>>', self._on_event_selected)
        # Enterキーで日付入力欄に移動
        self.event_combo.bind('<Return>', lambda e: self.date_entry.focus())
        
        # ========== 入力フォーム領域 ==========
        self.form_frame = ttk.LabelFrame(main_frame, text="入力項目", padding=10)
        self.form_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # ========== 共通項目 ==========
        common_frame = ttk.LabelFrame(main_frame, text="共通項目", padding=10)
        common_frame.pack(fill=tk.X, pady=(0, 10))
        
        # event_date（必須）
        ttk.Label(common_frame, text="日付*:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_entry = ttk.Entry(common_frame, width=20)
        self.date_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        # デフォルト値を今日の日付に設定
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_entry.insert(0, today)
        # Enterキーで最初の入力項目に移動
        self.date_entry.bind('<Return>', lambda e: self._move_from_date_field(e))
        
        # note（任意）
        ttk.Label(common_frame, text="備考:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.note_entry = ttk.Entry(common_frame, width=40)
        self.note_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        # メモ欄でEnterキーを押すと最初の入力項目に戻る
        self.note_entry.bind('<Return>', lambda e: self._move_from_note_field(e))
        
        # ========== ボタン ==========
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ok_button = ttk.Button(
            button_frame,
            text="OK",
            command=self._on_ok
        )
        ok_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self._on_cancel
        )
        cancel_button.pack(side=tk.LEFT, padx=5)
    
    def _on_event_selected(self, event):
        """イベント選択時の処理"""
        selected_name = self.event_combo.get()
        if not selected_name or selected_name not in self.event_map:
            return
        
        event_number = self.event_map[selected_name]
        self.selected_event_number = event_number
        self.selected_event = self.event_dictionary.get(str(event_number))
        
        # 入力フォームを再生成
        self._create_input_form()
    
    def _create_input_form(self):
        """入力フォームを動的生成"""
        # 既存のフォームをクリア
        for widget in self.form_frame.winfo_children():
            widget.destroy()
        self.field_widgets.clear()
        
        if not self.selected_event:
            return
        
        input_fields = self.selected_event.get('input_fields', [])
        
        if not input_fields:
            ttk.Label(
                self.form_frame,
                text="入力項目はありません",
                foreground="gray"
            ).pack(pady=10)
            return
        
        # 授精設定をロード（AIイベント用）
        insemination_settings = self._load_insemination_settings()
        
        # 各入力フィールドを生成
        entry_widgets = []  # Enterキーで移動する順序を保持
        for i, field in enumerate(input_fields):
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            label_text = field.get('label', key)
            
            # ラベル
            ttk.Label(
                self.form_frame,
                text=f"{label_text}:"
            ).grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)
            
            # 入力ウィジェット
            # AIイベント（200）のtechnician_codeとinsemination_type_codeはComboboxを使用
            if self.selected_event_number == 200:
                if key == 'technician_code':
                    # 授精師コードのCombobox
                    values = []
                    if insemination_settings:
                        technicians = insemination_settings.get('technicians', {})
                        for code, name in sorted(technicians.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
                            values.append(f"{code}：{name}")
                    
                    combo = ttk.Combobox(
                        self.form_frame,
                        values=values,
                        width=28,
                        state='readonly' if values else 'normal'
                    )
                    combo.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                    self.field_widgets[key] = combo
                    entry_widgets.append(combo)
                    continue
                elif key == 'insemination_type_code':
                    # 授精種類コードのCombobox
                    values = []
                    if insemination_settings:
                        insemination_types = insemination_settings.get('insemination_types', {})
                        for code, name in sorted(insemination_types.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
                            values.append(f"{code}：{name}")
                    
                    combo = ttk.Combobox(
                        self.form_frame,
                        values=values,
                        width=28,
                        state='readonly' if values else 'normal'
                    )
                    combo.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                    self.field_widgets[key] = combo
                    entry_widgets.append(combo)
                    continue
            
            # その他のフィールドは従来通り
            if datatype == 'int':
                entry = ttk.Entry(self.form_frame, width=20)
                entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                self.field_widgets[key] = entry
                entry_widgets.append(entry)
            elif datatype == 'float':
                entry = ttk.Entry(self.form_frame, width=20)
                entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                self.field_widgets[key] = entry
                entry_widgets.append(entry)
            elif datatype == 'str':
                entry = ttk.Entry(self.form_frame, width=30)
                entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                self.field_widgets[key] = entry
                entry_widgets.append(entry)
            else:
                # その他の型は文字列として扱う
                entry = ttk.Entry(self.form_frame, width=30)
                entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                self.field_widgets[key] = entry
                entry_widgets.append(entry)
            
            # 繁殖検査（301）、フレッシュチェック（300）、妊娠鑑定マイナス（302）の特定フィールドに半角大文字変換機能を追加
            if self.selected_event_number in [300, 301, 302]:
                if key in ['treatment', 'uterine_findings', 'left_ovary_findings', 'right_ovary_findings', 'other']:
                    entry.bind('<KeyRelease>', lambda e, widget=entry: self._convert_to_uppercase(widget))
        
        # Enterキーで次のフィールドに移動する機能を追加（Entryのみ）
        for i, widget in enumerate(entry_widgets):
            if isinstance(widget, ttk.Entry):
                if i < len(entry_widgets) - 1:
                    # 最後のフィールド以外は次のフィールドに移動
                    next_widget = entry_widgets[i + 1]
                    widget.bind('<Return>', lambda e, next=next_widget: self._move_to_next_field(e, next))
                else:
                    # 最後のフィールドはメモ欄に移動
                    widget.bind('<Return>', lambda e: self._move_to_note_field(e))
    
    def _validate_input(self) -> bool:
        """入力値を検証"""
        # 日付の検証
        event_date = self.date_entry.get().strip()
        if not event_date:
            messagebox.showerror("エラー", "日付を入力してください")
            self.date_entry.focus()
            return False
        
        # 日付形式の検証
        try:
            datetime.strptime(event_date, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("エラー", "日付の形式が正しくありません（YYYY-MM-DD形式）")
            self.date_entry.focus()
            return False
        
        # イベント選択の検証
        if not self.selected_event:
            messagebox.showerror("エラー", "イベントを選択してください")
            return False
        
        # 入力フィールドの検証
        input_fields = self.selected_event.get('input_fields', [])
        for field in input_fields:
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            widget = self.field_widgets.get(key)
            
            if widget:
                value = widget.get().strip()
                
                # 必須チェック（必要に応じて追加）
                # 型チェック
                if value:
                    if datatype == 'int':
                        try:
                            int(value)
                        except ValueError:
                            messagebox.showerror(
                                "エラー",
                                f"{field.get('label', key)} は整数で入力してください"
                            )
                            widget.focus()
                            return False
                    elif datatype == 'float':
                        try:
                            float(value)
                        except ValueError:
                            messagebox.showerror(
                                "エラー",
                                f"{field.get('label', key)} は数値で入力してください"
                            )
                            widget.focus()
                            return False
        
        return True
    
    def _load_event_data(self):
        """編集モードの場合、既存イベントデータをフォームに読み込む"""
        if not self.edit_event_data:
            return
        
        try:
            # イベント番号を設定
            event_number = self.edit_event_data.get('event_number')
            event_str = str(event_number)
            
            # event_dictionaryからイベント名を取得
            if event_str in self.event_dictionary:
                event_name = self.event_dictionary[event_str].get('name_jp')
                if event_name and event_name in self.event_map:
                    self.event_combo.set(event_name)
                    self._on_event_selected(None)
                    # _on_event_selected で self.selected_event が設定される
            
            # 日付を設定
            event_date = self.edit_event_data.get('event_date', '')
            if event_date:
                self.date_entry.delete(0, tk.END)
                self.date_entry.insert(0, event_date)
            
            # 備考を設定
            note = self.edit_event_data.get('note', '')
            if note:
                self.note_entry.delete(0, tk.END)
                self.note_entry.insert(0, note)
            
            # json_dataをinput_fieldsに展開
            # _on_event_selected が呼ばれていることを確認
            json_data = self.edit_event_data.get('json_data', {})
            if json_data and self.selected_event:
                input_fields = self.selected_event.get('input_fields', [])
                for field in input_fields:
                    key = field.get('key')
                    widget = self.field_widgets.get(key)
                    
                    if widget and key in json_data:
                        value = json_data[key]
                        
                        if isinstance(widget, ttk.Combobox):
                            # Comboboxの場合は、コードと名前の組み合わせを検索
                            if isinstance(value, (int, str)):
                                value_str = str(value)
                                # リストから該当する項目を探す
                                for item in widget['values']:
                                    if item.startswith(f"{value_str}："):
                                        widget.set(item)
                                        break
                                else:
                                    # 見つからない場合は値そのものを設定
                                    widget.set(value_str)
                        else:
                            # Entryの場合は値そのものを設定
                            widget.delete(0, tk.END)
                            widget.insert(0, str(value))
                            
        except Exception as e:
            logging.error(f"ERROR: _load_event_data で例外が発生しました: {e}")
            import traceback
            traceback.print_exc()
    
    def _on_ok(self):
        """OKボタンクリック時の処理"""
        if not self._validate_input():
            return
        
        # 入力値を取得
        event_date = self.date_entry.get().strip()
        note = self.note_entry.get().strip()
        
        # json_data を構築
        json_data = {}
        input_fields = self.selected_event.get('input_fields', [])
        for field in input_fields:
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            widget = self.field_widgets.get(key)
            
            if widget:
                if isinstance(widget, ttk.Combobox):
                    # Comboboxの場合は選択値を取得
                    value = widget.get().strip()
                    if value:
                        # 「1：園田」形式からコード部分のみを抽出
                        if '：' in value:
                            code = value.split('：')[0]
                            json_data[key] = code
                        else:
                            json_data[key] = value
                else:
                    # Entryの場合は従来通り
                    value = widget.get().strip()
                    if value:
                        # 型変換
                        if datatype == 'int':
                            json_data[key] = int(value)
                        elif datatype == 'float':
                            json_data[key] = float(value)
                        else:
                            json_data[key] = value
        
        # event_number を取得
        selected_name = self.event_combo.get()
        event_number = self.event_map[selected_name]
        
        try:
            # cow_auto_idがNoneの場合はエラー（将来的には入力欄から解決する実装を追加）
            if self.cow_auto_id is None:
                messagebox.showerror("エラー", "対象牛を指定してください")
                return
            
            # 編集モードの場合
            if self.edit_event_id:
                # 既存レコードをUPDATE
                event_data = {
                    'event_number': event_number,
                    'event_date': event_date,
                    'json_data': json_data if json_data else None,
                    'note': note if note else None
                }
                
                self.db.update_event(self.edit_event_id, event_data)
                
                # RuleEngine を呼ぶ
                self.rule_engine.on_event_updated(self.edit_event_id)
                
                logging.info(f"Event updated: event_id={self.edit_event_id}")
            else:
                # 新規作成の場合
                event_data = {
                    'cow_auto_id': self.cow_auto_id,
                    'event_number': event_number,
                    'event_date': event_date,
                    'json_data': json_data if json_data else None,
                    'note': note if note else None
                }
                
                event_id = self.db.insert_event(event_data)
                
                # RuleEngine を呼ぶ
                self.rule_engine.on_event_added(event_id)
                
                logging.info(f"Event created: event_id={event_id}")
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            # ウィンドウを閉じる
            self.window.destroy()
            
            messagebox.showinfo("完了", "イベントを保存しました")
            
        except Exception as e:
            logging.error(f"ERROR: _on_ok で例外が発生しました: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("エラー", f"イベントの保存に失敗しました: {e}")
    
    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.window.destroy()
    
    def _convert_to_uppercase(self, widget: tk.Widget):
        """
        入力されたテキストを半角大文字に変換（全角日本語は保持）
        
        Args:
            widget: Entryウィジェット
        """
        if not isinstance(widget, tk.Entry):
            return
        
        current_text = widget.get()
        cursor_pos = widget.index(tk.INSERT)
        
        # 全角文字を半角に変換し、英数字を大文字に変換
        # 全角英数字 → 半角大文字
        # 全角カタカナ・ひらがな・漢字はそのまま
        converted = ""
        for char in current_text:
            # 全角英数字を半角大文字に変換
            if ord('Ａ') <= ord(char) <= ord('Ｚ'):
                converted += chr(ord(char) - ord('Ａ') + ord('A'))
            elif ord('ａ') <= ord(char) <= ord('ｚ'):
                converted += chr(ord(char) - ord('ａ') + ord('A'))
            elif ord('０') <= ord(char) <= ord('９'):
                converted += chr(ord(char) - ord('０') + ord('0'))
            # 半角英数字を大文字に変換
            elif 'a' <= char <= 'z':
                converted += char.upper()
            elif 'A' <= char <= 'Z' or '0' <= char <= '9':
                converted += char
            # その他の文字（日本語など）はそのまま
            else:
                converted += char
        
        # テキストが変更された場合のみ更新
        if converted != current_text:
            widget.delete(0, tk.END)
            widget.insert(0, converted)
            # カーソル位置を復元（可能な範囲で）
            try:
                widget.icursor(min(cursor_pos, len(converted)))
            except:
                pass
    
    def _move_to_next_field(self, event, next_widget: tk.Widget):
        """
        Enterキーで次のフィールドに移動
        
        Args:
            event: イベントオブジェクト
            next_widget: 次のウィジェット
        """
        next_widget.focus()
        return "break"  # デフォルトのEnterキー動作を防ぐ
    
    def _move_to_note_field(self, event):
        """
        メモ欄に移動
        
        Args:
            event: イベントオブジェクト
        """
        self.note_entry.focus()
        return "break"
    
    def _move_from_date_field(self, event):
        """
        日付入力欄から最初の入力項目に移動
        
        Args:
            event: イベントオブジェクト
        """
        if self.field_widgets:
            # 最初の入力項目に移動
            first_widget = list(self.field_widgets.values())[0]
            first_widget.focus()
        return "break"
    
    def _move_from_note_field(self, event):
        """
        メモ欄から最初の入力項目に戻る
        
        Args:
            event: イベントオブジェクト
        """
        if self.field_widgets:
            # 最初の入力項目に移動
            first_widget = list(self.field_widgets.values())[0]
            first_widget.focus()
        return "break"
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

