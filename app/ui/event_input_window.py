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
from datetime import datetime, timedelta
import re
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.reproduction_checkup_billing import REPRO_CHECKUP_EVENT_NUMBERS


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
        self.window.geometry("500x400")  # 初期サイズを少し小さく（後で動的調整）
        self.window.minsize(500, 300)  # 最小サイズを設定
        
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
    
    def _load_treatment_settings(self) -> Optional[Dict[str, Any]]:
        """繁殖処置設定をロード"""
        if not self.farm_path:
            return None
        
        settings_file = self.farm_path / "reproduction_treatment_settings.json"
        if not settings_file.exists():
            return None
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"繁殖処置設定ファイル読み込みエラー: {e}")
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
        
        # ラベルの幅を統一するため、列の設定
        common_frame.columnconfigure(0, minsize=100)  # ラベル列の最小幅
        common_frame.columnconfigure(1, weight=1)  # 入力欄列は伸縮可能
        
        # event_date（必須）
        ttk.Label(common_frame, text="日付*:").grid(row=0, column=0, sticky=tk.E, padx=(5, 10), pady=5)
        self.date_entry = ttk.Entry(common_frame, width=20)
        self.date_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        # デフォルト値を今日の日付に設定
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_entry.insert(0, today)
        # 日付入力欄からフォーカスが外れたときに自動変換
        self.date_entry.bind('<FocusOut>', lambda e: self._parse_and_convert_date())
        # Enterキーで最初の入力項目に移動（日付変換も実行）
        self.date_entry.bind('<Return>', lambda e: self._on_date_return())
        
        # note（任意）
        ttk.Label(common_frame, text="備考:").grid(row=1, column=0, sticky=tk.E, padx=(5, 10), pady=5)
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
        
        # ラベルの幅を統一するため、列の設定
        self.form_frame.columnconfigure(0, minsize=100)  # ラベル列の最小幅
        self.form_frame.columnconfigure(1, weight=1)  # 入力欄列は伸縮可能
        
        if not self.selected_event:
            return
        
        input_fields = self.selected_event.get('input_fields', [])
        
        if not input_fields:
            ttk.Label(
                self.form_frame,
                text="入力項目はありません",
                foreground="gray"
            ).pack(pady=10)
            # ウィンドウサイズを調整
            self._adjust_window_size()
            return
        
        # 授精設定をロード（AIイベント用）
        insemination_settings = self._load_insemination_settings()
        
        # 繁殖処置設定をロード（フレッシュチェック、繁殖検査、妊娠鑑定マイナス用）
        treatment_settings = None
        if self.selected_event_number in [300, 301, 302]:
            treatment_settings = self._load_treatment_settings()
        
        # 各入力フィールドを生成
        entry_widgets = []  # Enterキーで移動する順序を保持
        for i, field in enumerate(input_fields):
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            label_text = field.get('label', key)
            
            # ラベル（右揃えで統一）
            ttk.Label(
                self.form_frame,
                text=f"{label_text}:"
            ).grid(row=i, column=0, sticky=tk.E, padx=(5, 10), pady=5)
            
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
            
            # event_reference型のフィールド（AI/ETイベント選択）
            if datatype == 'event_reference':
                event_types = field.get('event_types', [200, 201])  # デフォルトはAI/ET
                # 対象牛のAI/ETイベントを取得
                ai_et_events = []
                if self.cow_auto_id:
                    all_events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
                    for event in all_events:
                        if event.get('event_number') in event_types:
                            event_id = event.get('id')
                            event_date = event.get('event_date', '')
                            event_number = event.get('event_number')
                            event_type_name = 'AI' if event_number == 200 else 'ET'
                            # json_dataからSIREを取得
                            json_data = event.get('json_data') or {}
                            if isinstance(json_data, str):
                                try:
                                    json_data = json.loads(json_data)
                                except:
                                    json_data = {}
                            sire = json_data.get('sire', '')
                            display_text = f"{event_date} ({event_type_name})"
                            if sire:
                                display_text += f" - SIRE: {sire}"
                            ai_et_events.append((event_id, display_text))
                
                # 日付順にソート（降順：最新が先頭）
                ai_et_events.sort(key=lambda x: x[1], reverse=True)
                
                values = [f"{event_id}: {display}" for event_id, display in ai_et_events]
                combo = ttk.Combobox(
                    self.form_frame,
                    values=values,
                    width=40,
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
                # 処置入力欄の場合、プルダウン（Combobox）を使用
                if key == 'treatment':
                    # 処置コード入力欄が存在する場合は、処置内容は読み取り専用にする
                    if 'treatment_code' in self.field_widgets:
                        entry.config(state='readonly')
                    else:
                        # 処置コード入力欄が存在しない場合は、プルダウンまたは直接入力可能なComboboxを使用
                        # 既存のEntryを削除してComboboxに置き換え
                        entry.destroy()
                        
                        # プルダウンの値を準備
                        values = []
                        if treatment_settings:
                            treatments = treatment_settings.get('treatments', {})
                            for code, treatment_data in sorted(treatments.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 999):
                                name = treatment_data.get('name', '')
                                if code and name:
                                    values.append(f"{code}：{name}")
                        
                        # Comboboxを作成
                        combo = ttk.Combobox(
                            self.form_frame,
                            values=values,
                            width=28,
                            state='normal'  # 直接入力も可能
                        )
                        combo.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
                        self.field_widgets[key] = combo
                        # entry_widgetsの該当インデックスを更新
                        if i < len(entry_widgets):
                            entry_widgets[i] = combo
                        else:
                            entry_widgets.append(combo)
                        
                        # 選択時に処置内容のみを入力欄に設定
                        def on_treatment_selected(event):
                            selected = combo.get()
                            if selected and '：' in selected:
                                # "1：WPG" 形式から "WPG" のみを抽出
                                parts = selected.split('：', 1)
                                if len(parts) == 2:
                                    treatment_name = parts[1].strip()
                                    combo.set(treatment_name)
                        
                        combo.bind('<<ComboboxSelected>>', on_treatment_selected)
                        
                        # 大文字変換機能も追加
                        combo.bind('<KeyRelease>', lambda e, widget=combo: self._convert_to_uppercase_combobox(widget))
                        continue
                elif key in ['uterine_findings', 'left_ovary_findings', 'right_ovary_findings', 'other']:
                    entry.bind('<KeyRelease>', lambda e, widget=entry: self._convert_to_uppercase(widget))
        
        # Enterキーで次のフィールドに移動する機能を追加（EntryとCombobox）
        for i, widget in enumerate(entry_widgets):
            if isinstance(widget, (ttk.Entry, ttk.Combobox)):
                if i < len(entry_widgets) - 1:
                    # 最後のフィールド以外は次のフィールドに移動
                    next_widget = entry_widgets[i + 1]
                    widget.bind('<Return>', lambda e, next=next_widget: self._move_to_next_field(e, next))
                else:
                    # 最後のフィールドはメモ欄に移動
                    widget.bind('<Return>', lambda e: self._move_to_note_field(e))
        
        # ウィンドウサイズを調整
        self._adjust_window_size()
    
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
                        field_datatype = field.get('datatype', 'str')
                        
                        if isinstance(widget, ttk.Combobox):
                            # event_reference型の場合はevent_idで検索
                            if field_datatype == 'event_reference':
                                if isinstance(value, (int, str)):
                                    event_id = int(value) if isinstance(value, str) and value.isdigit() else value
                                    # リストから該当するevent_idの項目を探す
                                    for item in widget['values']:
                                        if item.startswith(f"{event_id}:"):
                                            widget.set(item)
                                            break
                            else:
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
            
            # ウィンドウサイズを調整
            self._adjust_window_size()
                            
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
                        # event_reference型の場合はevent_idを抽出
                        if datatype == 'event_reference':
                            # 「123: 2024-01-01 (AI) - SIRE: ABC123」形式からevent_idを抽出
                            if ':' in value:
                                event_id_str = value.split(':')[0].strip()
                                try:
                                    json_data[key] = int(event_id_str)
                                except ValueError:
                                    pass
                            continue
                        # 処置フィールドの場合、「1：WPG」形式から「WPG」のみを抽出
                        if key == 'treatment' and '：' in value:
                            parts = value.split('：', 1)
                            if len(parts) == 2:
                                treatment_name = parts[1].strip()
                                json_data[key] = treatment_name
                            else:
                                json_data[key] = value
                        # その他のCombobox（授精師コードなど）はコード部分のみを抽出
                        elif '：' in value:
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
        
        # AI/ETのとき新規SIREなら種別の登録ダイアログを表示
        if event_number in (RuleEngine.EVENT_AI, RuleEngine.EVENT_ET) and self.farm_path:
            sire = (json_data.get('sire') or json_data.get('sire_name') or '').strip()
            if sire:
                from ui.sire_list_window import get_known_sire_names, show_sire_confirm_dialog
                known = get_known_sire_names(self.db, self.farm_path)
                if sire not in known:
                    show_sire_confirm_dialog(self.window, self.farm_path, sire)
        
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
                
                # 繁殖検診イベントで同じ日付に既存がある場合は警告
                if event_number in REPRO_CHECKUP_EVENT_NUMBERS and self.cow_auto_id is not None:
                    events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
                    date_str = (event_date or "")[:10]
                    same_day_repro = [
                        e for e in events
                        if ((e.get("event_date") or "")[:10] == date_str
                            and e.get("event_number") in REPRO_CHECKUP_EVENT_NUMBERS)
                    ]
                    if same_day_repro:
                        if not messagebox.askyesno(
                            "重複の可能性",
                            "同じ日付で既に繁殖検診イベントが登録されています。"
                            "重複の可能性があります。\n\n登録しますか？",
                        ):
                            return
                
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
    
    def _convert_to_uppercase_combobox(self, widget: tk.Widget):
        """
        Comboboxの入力されたテキストを半角大文字に変換（全角日本語は保持）
        
        Args:
            widget: Comboboxウィジェット
        """
        if not isinstance(widget, ttk.Combobox):
            return
        
        current_text = widget.get()
        if not current_text:
            return
        
        cursor_pos = widget.index(tk.INSERT)
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
            widget.set(converted)
            # カーソル位置を復元（可能な範囲で）
            try:
                widget.icursor(min(cursor_pos, len(converted)))
            except:
                pass
    
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
    
    def _on_treatment_code_changed(self, widget: tk.Widget, treatment_settings: Optional[Dict[str, Any]]):
        """
        処置コード入力時に処置内容を自動入力
        
        Args:
            widget: 処置コード入力欄のウィジェット
            treatment_settings: 繁殖処置設定
        """
        if not isinstance(widget, tk.Entry):
            return
        
        treatment_code = widget.get().strip()
        if not treatment_code or not treatment_settings:
            return
        
        # 処置設定から処置内容を取得
        treatments = treatment_settings.get('treatments', {})
        treatment = treatments.get(treatment_code)
        
        if treatment:
            treatment_name = treatment.get('name', '')
            # 処置内容入力欄に自動入力
            if 'treatment' in self.field_widgets:
                treatment_widget = self.field_widgets['treatment']
                if isinstance(treatment_widget, tk.Entry):
                    treatment_widget.config(state='normal')
                    treatment_widget.delete(0, tk.END)
                    treatment_widget.insert(0, treatment_name)
                    treatment_widget.config(state='readonly')
    
    def _on_treatment_input_changed(self, widget: tk.Widget, treatment_settings: Optional[Dict[str, Any]]):
        """
        処置入力欄にコードを入力した場合、自動的に処置内容に変換
        
        Args:
            widget: 処置入力欄のウィジェット
            treatment_settings: 繁殖処置設定
        """
        if not isinstance(widget, tk.Entry):
            return
        
        input_value = widget.get().strip()
        if not input_value:
            return
        
        # 入力値が数字のみの場合、処置コードとして扱う
        if input_value.isdigit():
            if not treatment_settings:
                logging.warning("_on_treatment_input_changed: treatment_settingsがNoneです")
                # 大文字変換のみ実行
                self._convert_to_uppercase(widget)
                return
            
            treatments = treatment_settings.get('treatments', {})
            logging.info(f"_on_treatment_input_changed: input_value={input_value}, treatments={list(treatments.keys())}")
            
            # キーを文字列に統一して検索
            treatment = None
            for key, value in treatments.items():
                if str(key) == input_value:
                    treatment = value
                    break
            
            if treatment:
                treatment_name = treatment.get('name', '')
                logging.info(f"_on_treatment_input_changed: 処置内容を取得しました: {treatment_name}")
                # 処置内容に変換
                cursor_pos = widget.index(tk.INSERT)
                widget.delete(0, tk.END)
                widget.insert(0, treatment_name)
                # カーソル位置を最後に移動
                widget.icursor(tk.END)
            else:
                logging.warning(f"_on_treatment_input_changed: 処置コード '{input_value}' が見つかりません")
                # 見つからない場合は大文字変換のみ実行
                self._convert_to_uppercase(widget)
        else:
            # 数字以外の場合は大文字変換のみ実行
            self._convert_to_uppercase(widget)
    
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
    
    def _parse_and_convert_date(self) -> bool:
        """
        日付入力欄の値をパースして自動変換
        
        Returns:
            変換に成功した場合はTrue
        """
        date_str = self.date_entry.get().strip()
        if not date_str:
            return False
        
        # 既にYYYY-MM-DD形式の場合はそのまま
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            return True
        except ValueError:
            pass
        
        # 日付をパースして変換
        try:
            parsed_date = self._parse_date_input(date_str)
            if parsed_date:
                # 変換された日付を入力欄に設定
                self.date_entry.delete(0, tk.END)
                self.date_entry.insert(0, parsed_date.strftime('%Y-%m-%d'))
                return True
        except Exception as e:
            # パースエラーは無視（ユーザーが入力中の場合もある）
            pass
        
        return False
    
    def _parse_date_input(self, date_str: str) -> Optional[datetime]:
        """
        日付入力をパースしてdatetimeオブジェクトに変換
        
        ルール:
        - 数字だけ（例：1, 10）→ 直近の日にち
          - 本日1月9日なら、1→1月1日、10→前年12月10日
        - 月日（例：02-02, 2/2, 2-2）→ 直近の月日
          - 本日1月9日で、入力が02-02であれば、昨年の2月2日
        
        Args:
            date_str: 日付文字列
        
        Returns:
            パースされたdatetimeオブジェクト、パースできない場合はNone
        """
        date_str = date_str.strip()
        if not date_str:
            return None
        
        today = datetime.now()
        current_year = today.year
        current_month = today.month
        current_day = today.day
        
        # 数字だけの場合（1-31の範囲）
        if date_str.isdigit():
            day = int(date_str)
            if 1 <= day <= 31:
                # 今月の該当日を試す
                try:
                    candidate = datetime(current_year, current_month, day)
                    if candidate <= today:
                        # 今日以前なら今月
                        return candidate
                    else:
                        # 今日より後なら前月
                        if current_month == 1:
                            return datetime(current_year - 1, 12, day)
                        else:
                            return datetime(current_year, current_month - 1, day)
                except ValueError:
                    # 無効な日付（例：2月30日）の場合は前月を試す
                    if current_month == 1:
                        try:
                            return datetime(current_year - 1, 12, day)
                        except ValueError:
                            return None
                    else:
                        try:
                            return datetime(current_year, current_month - 1, day)
                        except ValueError:
                            return None
        
        # 月日の形式（MM-DD, M-D, MM/DD, M/D）
        # ハイフンまたはスラッシュで区切られている
        separators = ['-', '/', '‐', '－']  # 全角ハイフンも対応
        for sep in separators:
            if sep in date_str:
                parts = date_str.split(sep)
                if len(parts) == 2:
                    try:
                        month = int(parts[0].strip())
                        day = int(parts[1].strip())
                        
                        if 1 <= month <= 12 and 1 <= day <= 31:
                            # 今年の該当日を試す
                            try:
                                candidate = datetime(current_year, month, day)
                                if candidate <= today:
                                    # 今日以前なら今年
                                    return candidate
                                else:
                                    # 今日より後なら昨年
                                    return datetime(current_year - 1, month, day)
                            except ValueError:
                                # 無効な日付（例：2月30日）の場合は昨年を試す
                                try:
                                    return datetime(current_year - 1, month, day)
                                except ValueError:
                                    return None
                    except ValueError:
                        continue
        
        return None
    
    def _on_date_return(self, event=None):
        """
        日付入力欄でEnterキーが押されたときの処理
        """
        # 日付を変換
        self._parse_and_convert_date()
        # 最初の入力項目に移動
        if self.field_widgets:
            first_widget = list(self.field_widgets.values())[0]
            first_widget.focus()
        if event:
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
    
    def _adjust_window_size(self):
        """ウィンドウサイズを入力欄の数に応じて調整"""
        try:
            # ウィンドウの現在の幅を取得
            current_width = self.window.winfo_width()
            if current_width < 100:
                current_width = 500  # デフォルト幅
            
            # 必要な高さを計算
            # 基本高さ: イベント選択 + 共通項目 + ボタン + パディング
            base_height = 200
            
            # 入力項目の数に応じて高さを追加
            input_fields = []
            if self.selected_event:
                input_fields = self.selected_event.get('input_fields', [])
            
            # 各行の高さ（ラベル + パディング）を約30pxと仮定
            field_height = 30
            form_height = len(input_fields) * field_height
            
            # 最小高さを確保
            min_height = 300
            calculated_height = base_height + form_height + 50  # 余裕を持たせる
            window_height = max(min_height, calculated_height)
            
            # 最大高さを制限（画面からはみ出さないように）
            max_height = 800
            window_height = min(window_height, max_height)
            
            # ウィンドウサイズを更新
            self.window.geometry(f"{current_width}x{window_height}")
            
            # スクロールが必要な場合はスクロール可能にする
            if window_height >= max_height:
                # CanvasとScrollbarを追加（必要に応じて）
                pass  # 現時点では最大高さで制限
        except Exception as e:
            logging.warning(f"ウィンドウサイズ調整エラー: {e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

