"""
FALCON2 - EventInputWindow
イベント入力ウィンドウ（辞書駆動型・左右2カラム構成）
設計書 第18章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable, List
from pathlib import Path
import json
from datetime import datetime, timedelta
import re
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.event_display import format_insemination_event, format_calving_event
from settings_manager import SettingsManager


def normalize_date(date_str: str) -> Optional[str]:
    """
    日付文字列を正規化（YYYY-MM-DD形式に変換）
    
    対応形式：
    - 空欄 → 本日
    - YYYY/MM/DD → そのまま採用
    - M/D → 当年の日付
    - D → 直近の D 日
    
    Args:
        date_str: 入力された日付文字列
    
    Returns:
        正規化された日付（YYYY-MM-DD形式）、不正な場合はNone
    """
    date_str = date_str.strip()
    
    # 空欄 → 本日
    if not date_str:
        return datetime.now().strftime('%Y-%m-%d')
    
    # YYYY/MM/DD または YYYY-MM-DD
    if '/' in date_str or '-' in date_str:
        parts = date_str.replace('-', '/').split('/')
        if len(parts) == 3:
            try:
                year = int(parts[0])
                month = int(parts[1])
                day = int(parts[2])
                
                # 年が2桁の場合は補完
                if year < 100:
                    current_year = datetime.now().year
                    century = (current_year // 100) * 100
                    year = century + year
                
                date_obj = datetime(year, month, day)
                return date_obj.strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                return None
    
    # M/D 形式
    if '/' in date_str:
        parts = date_str.split('/')
        if len(parts) == 2:
            try:
                month = int(parts[0])
                day = int(parts[1])
                current_year = datetime.now().year
                date_obj = datetime(current_year, month, day)
                return date_obj.strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                return None
    
    # D 形式（直近の D 日）
    try:
        day = int(date_str)
        if 1 <= day <= 31:
            today = datetime.now()
            # 今月の該当日を試す
            try:
                date_obj = datetime(today.year, today.month, day)
                if date_obj <= today:
                    return date_obj.strftime('%Y-%m-%d')
            except ValueError:
                pass
            
            # 先月の該当日を試す
            try:
                if today.month == 1:
                    date_obj = datetime(today.year - 1, 12, day)
                else:
                    date_obj = datetime(today.year, today.month - 1, day)
                if date_obj <= today:
                    return date_obj.strftime('%Y-%m-%d')
            except ValueError:
                pass
    except ValueError:
        pass
    
    return None


class EventInputWindow:
    """イベント入力ウィンドウ（左右2カラム構成）"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, 
                 rule_engine: RuleEngine,
                 event_dictionary_path: Path,
                 cow_auto_id: Optional[int] = None,  # Noneの場合はメニュー起動
                 event_id: Optional[int] = None,  # 編集時は指定（後方互換性のため残す）
                 on_saved: Optional[Callable[[int], None]] = None,
                 farm_path: Optional[Path] = None,  # 農場パス（SettingsManager用）
                 edit_event_id: Optional[int] = None):  # 編集時のイベントID（event_id の別名）
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            db_handler: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
            event_dictionary_path: event_dictionary.json のパス
            cow_auto_id: 対象牛の auto_id（Noneの場合はメニュー起動、牛ID入力欄を表示）
            event_id: 編集時のイベントID（Noneの場合は新規、後方互換性のため残す）
            on_saved: 保存完了時のコールバック（cow_auto_id を引数に取る）
            farm_path: 農場パス（SettingsManager用、Noneの場合はDBから推測）
            edit_event_id: 編集時のイベントID（event_id の別名、優先される）
        """
        self.db = db_handler
        self.rule_engine = rule_engine
        self.cow_auto_id = cow_auto_id  # Noneの場合は後で入力欄から解決
        self.event_dict_path = event_dictionary_path
        # edit_event_id が指定された場合はそれを優先、なければ event_id を使用
        self.event_id = edit_event_id if edit_event_id is not None else event_id
        self.on_saved = on_saved
        
        # SettingsManagerを初期化（farm_pathが指定されていない場合はDBから推測）
        if farm_path is None and cow_auto_id is not None:
            # cow_auto_idから農場パスを推測
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                frm = cow.get('frm')
                if frm:
                    farm_path = Path(f"C:/FARMS/{frm}")
        
        if farm_path:
            self.settings_manager = SettingsManager(farm_path)
        else:
            # デフォルトパスを使用（後で設定可能）
            self.settings_manager = None
        
        # event_dictionary を読み込む
        self.event_dictionary: Dict[str, Dict[str, Any]] = {}
        self.input_code_map: Dict[int, List[str]] = {}  # input_code -> [event_number, ...]
        self._load_event_dictionary()
        
        # 現在選択中のイベント
        self.selected_event: Optional[Dict[str, Any]] = None
        self.selected_event_number: Optional[int] = None
        
        # 動的フォームのウィジェット
        self.field_widgets: Dict[str, tk.Widget] = {}
        # 分娩専用UI用の状態
        self.calving_difficulty_var: Optional[tk.StringVar] = None
        self.calving_child_count_var: Optional[tk.IntVar] = None
        self.calving_calf_vars: List[Dict[str, Any]] = []
        self.calving_block_container: Optional[ttk.Frame] = None
        self._editing_event_json: Dict[str, Any] = {}
        
        # 授精設定データ（insemination_settings.jsonからロード）
        self.technicians: Dict[str, str] = {}  # {"1": "園田", "2": "NOSAI北見"}
        self.insemination_types: Dict[str, str] = {}  # {"1": "自然発情", "2": "CIDR"}
        self.pen_settings: Dict[str, str] = {}  # {"1": "Lact1", ...}
        
        # 授精設定をロード
        self._load_insemination_settings()
        # PEN設定をロード（将来のMOVEイベント等で利用）
        self._load_pen_settings()
        
        # イベント候補リスト（絞り込み用）
        self.event_candidates: List[Dict[str, Any]] = []  # [{'event_number': int, 'name_jp': str}, ...]
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("イベント入力" if self.event_id is None else "イベント編集")
        self.window.geometry("1000x700")
        self.window.transient(parent)
        self.window.grab_set()
        
        self._create_widgets()
        
        # 編集時は既存データを読み込む
        if self.event_id is not None:
            self._load_event_data_for_edit()
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path is not None and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
                
                # input_code マップを作成
                for event_num_str, event_data in self.event_dictionary.items():
                    if event_data.get('deprecated', False):
                        continue
                    input_code = event_data.get('input_code')
                    if input_code is not None:
                        if input_code not in self.input_code_map:
                            self.input_code_map[input_code] = []
                        self.input_code_map[input_code].append(event_num_str)
            except Exception as e:
                messagebox.showerror("エラー", f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
        else:
            messagebox.showerror("エラー", f"event_dictionary.json が見つかりません: {self.event_dict_path}")
            self.event_dictionary = {}
    
    def _load_insemination_settings(self):
        """insemination_settings.json をロード"""
        if not self.settings_manager:
            return
        
        farm_path = self.settings_manager.farm_path
        settings_file = farm_path / "insemination_settings.json"
        
        if not settings_file.exists():
            logging.warning(f"insemination_settings.json が見つかりません: {settings_file}")
            return
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            self.technicians = settings.get('technicians', {})
            self.insemination_types = settings.get('insemination_types', {})
            
            logging.debug(f"授精設定をロード: technicians={len(self.technicians)}, insemination_types={len(self.insemination_types)}")
        except Exception as e:
            logging.error(f"insemination_settings.json 読み込みエラー: {e}")
            self.technicians = {}
            self.insemination_types = {}

    def _load_pen_settings(self):
        """PEN設定をロード"""
        if not self.settings_manager:
            return
        try:
            # farm_settings.json からロード
            self.pen_settings = self.settings_manager.load_pen_settings()
        except Exception as e:
            logging.error(f"PEN設定の読み込みに失敗しました: {e}")
            self.pen_settings = {}
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインPanedWindow（左右分割）
        main_paned = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========== 左カラム：イベント入力フォーム ==========
        left_frame = ttk.Frame(main_paned, width=500)
        main_paned.add(left_frame, weight=1)
        
        # スクロール可能なフレーム（左カラム全体）
        left_canvas = tk.Canvas(left_frame)
        left_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas)
        
        left_scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        
        left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ========== 1. 牛ID入力欄（メニュー起動時のみ） ==========
        if self.cow_auto_id is None:
            cow_id_frame = ttk.LabelFrame(left_scrollable_frame, text="牛ID入力", padding=10)
            cow_id_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(cow_id_frame, text="牛ID*:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
            self.cow_id_entry = ttk.Entry(cow_id_frame, width=20)
            self.cow_id_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
            self.cow_id_entry.bind('<KeyRelease>', self._on_cow_id_changed)
            # Enterキーで牛IDを解決してイベント番号入力欄に移動
            self.cow_id_entry.bind('<Return>', lambda e: self._on_cow_id_enter())
            self.cow_id_entry.bind('<FocusOut>', lambda e: self._on_cow_id_enter())
            ttk.Label(cow_id_frame, text="(拡大4桁ID または JPN10)", font=("", 8), foreground="gray").grid(
                row=0, column=2, sticky=tk.W, padx=5
            )
            
            # 対象牛表示ラベル（初期は非表示）
            self.cow_info_label = ttk.Label(cow_id_frame, text="", foreground="blue", font=("", 9, "bold"))
            self.cow_info_label.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        # ========== 2. 対象牛表示ラベル（個体カード起動時） ==========
        else:
            cow_info_frame = ttk.Frame(left_scrollable_frame)
            cow_info_frame.pack(fill=tk.X, pady=(0, 10))
            self.cow_info_label = ttk.Label(
                cow_info_frame,
                text="",
                foreground="blue",
                font=("", 9, "bold")
            )
            self.cow_info_label.pack(side=tk.LEFT, padx=5)
            self._show_cow_info()
        
        # ========== 3. イベント番号入力欄 ==========
        event_frame = ttk.LabelFrame(left_scrollable_frame, text="イベント選択", padding=10)
        event_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(event_frame, text="イベント番号:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.event_number_entry = ttk.Entry(event_frame, width=15)
        self.event_number_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.event_number_entry.bind('<KeyRelease>', self._on_event_number_changed)
        # Enter/Tabキーでイベント番号を確定してから日付入力欄に移動
        self.event_number_entry.bind('<Return>', self._on_event_number_enter)
        self.event_number_entry.bind('<Tab>', self._on_event_number_tab)
        ttk.Label(event_frame, text="(数字を入力すると候補が絞り込まれます)", font=("", 8), foreground="gray").grid(
            row=0, column=2, sticky=tk.W, padx=5
        )
        
        # ========== 4. イベント候補プルダウン ==========
        ttk.Label(event_frame, text="イベント候補:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.event_combo = ttk.Combobox(event_frame, width=40, state="readonly")
        self.event_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5, columnspan=2)
        self.event_combo.bind('<<ComboboxSelected>>', self._on_event_combo_selected)
        
        # ========== 5. 日付入力欄 ==========
        date_frame = ttk.LabelFrame(left_scrollable_frame, text="共通項目", padding=10)
        date_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(date_frame, text="日付*:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_entry = ttk.Entry(date_frame, width=20)
        self.date_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        # デフォルト値を今日の日付に設定
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_entry.insert(0, today)
        ttk.Label(date_frame, text="(YYYY/MM/DD, M/D, D 形式可)", font=("", 8), foreground="gray").grid(
            row=0, column=2, sticky=tk.W, padx=5
        )
        # 日付入力欄でEnterキーを押すとイベント番号入力欄に移動（イベント未選択時）または最初の入力項目に移動
        self.date_entry.bind('<Return>', lambda e: self._move_from_date_field(e))
        
        # ========== 6. イベント詳細入力欄（動的） ==========
        self.form_frame = ttk.LabelFrame(left_scrollable_frame, text="入力項目", padding=10)
        self.form_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # ========== 7. メモ欄 ==========
        ttk.Label(date_frame, text="メモ:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.note_entry = ttk.Entry(date_frame, width=40)
        self.note_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5, columnspan=2)
        # メモ欄でEnterキーを押すとOKボタンにフォーカス（または最初の入力項目に戻る）
        self.note_entry.bind('<Return>', lambda e: self._move_from_note_field(e))
        
        # ========== 8. 保存 / キャンセル ==========
        button_frame = ttk.Frame(left_scrollable_frame)
        button_frame.pack(fill=tk.X)
        
        ok_button = ttk.Button(button_frame, text="OK", command=self._on_ok)
        ok_button.pack(side=tk.LEFT, padx=5)
        
        if self.event_id:
            delete_button = ttk.Button(button_frame, text="削除", command=self._on_delete)
            delete_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = ttk.Button(button_frame, text="キャンセル", command=self._on_cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # ========== 右カラム：イベント履歴表示 ==========
        right_frame = ttk.Frame(main_paned, width=400)
        main_paned.add(right_frame, weight=1)
        
        history_frame = ttk.LabelFrame(right_frame, text="イベント履歴", padding=10)
        history_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview
        columns = ('date', 'event', 'info')
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show='headings', height=20)
        
        self.history_tree.heading('date', text='日付')
        self.history_tree.heading('event', text='イベント')
        self.history_tree.heading('info', text='簡易情報')
        
        self.history_tree.column('date', width=100)
        self.history_tree.column('event', width=150)
        self.history_tree.column('info', width=200)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 右クリックメニューを追加
        self.history_context_menu = tk.Menu(self.window, tearoff=0)
        self.history_context_menu.add_command(label="編集", command=self._on_edit_event_from_history)
        self.history_context_menu.add_separator()
        self.history_context_menu.add_command(label="削除", command=self._on_delete_event_from_history)
        
        # 右クリックイベントをバインド（Windows/Linux/Mac対応）
        self.history_tree.bind('<Button-3>', self._on_history_right_click)  # Windows/Linux
        self.history_tree.bind('<Button-2>', self._on_history_right_click)  # Mac
        # 念のため、Control+クリックも対応
        self.history_tree.bind('<Control-Button-1>', self._on_history_right_click)
        
        # 履歴表示は cow_auto_id が確定している場合のみ
        if self.cow_auto_id is not None:
            self._load_event_history()
        else:
            # メニュー起動時は「牛IDを入力してください」と表示（tagsを付けない）
            self.history_tree.insert('', 'end', values=("", "牛IDを入力してください", ""), tags=('no_event',))
    
    def _on_event_number_changed(self, event):
        """イベント番号入力変更時の処理（候補を絞り込む・自動選択しない）"""
        # 特殊キー（Enter, Tab, Shift等）の場合は処理しない
        if event.keysym in ('Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R'):
            return
        
        event_number_str = self.event_number_entry.get().strip()
        
        if not event_number_str:
            # 空欄の場合は全候補を表示（自動選択しない）
            self._update_event_candidates("", auto_select=False)
            return
        
        # 入力内容に応じて候補を更新（自動選択しない）
        self._update_event_candidates(event_number_str, auto_select=False)
    
    def _on_event_number_enter(self, event):
        """イベント番号入力欄でEnterキーが押された時の処理（確定）"""
        event_number_str = self.event_number_entry.get().strip()
        
        if event_number_str:
            # イベント番号を確定
            self._select_event_by_number(event_number_str)
        
        # 日付入力欄に移動
        self.date_entry.focus()
        return "break"
    
    def _on_event_number_tab(self, event):
        """イベント番号入力欄でTabキーが押された時の処理（確定）"""
        event_number_str = self.event_number_entry.get().strip()
        
        if event_number_str:
            # イベント番号を確定
            self._select_event_by_number(event_number_str)
        
        # 通常のTab動作（次のウィジェットに移動）を許可
        return None
    
    def _select_event_by_number(self, event_number_str: str):
        """
        イベント番号文字列からイベントを選択
        
        Args:
            event_number_str: イベント番号文字列（例："200"）
        """
        if not event_number_str:
            return
        
        # 数字の場合のみ処理
        if not event_number_str.isdigit():
            return
        
        try:
            event_number = int(event_number_str)
            event_num_str = str(event_number)
            
            # event_dictionary に存在するか確認
            if event_num_str in self.event_dictionary:
                self._select_event(event_num_str)
                # Combobox も更新
                event_data = self.event_dictionary[event_num_str]
                name_jp = event_data.get('name_jp', f'イベント{event_number}')
                self.event_combo.set(f"{event_number}: {name_jp}")
        except (ValueError, KeyError):
            pass
    
    def _update_event_candidates(self, search_str: str, auto_select: bool = False):
        """
        イベント候補を更新（番号・alias・name_jp の前方一致で絞り込み）
        
        Args:
            search_str: 検索文字列（数字または文字列）
                       - 数字の場合: event_number の前方一致
                       - 文字列の場合: alias または name_jp の前方一致
            auto_select: 自動選択するか（False の場合は候補表示のみ）
        """
        candidates = []
        
        if not search_str:
            # 空欄の場合は全イベントを候補に
            for event_num_str, event_data in self.event_dictionary.items():
                if event_data.get('deprecated', False):
                    continue
                event_number = int(event_num_str)
                name_jp = event_data.get('name_jp', f'イベント{event_number}')
                candidates.append({
                    'event_number': event_number,
                    'name_jp': name_jp,
                    'event_num_str': event_num_str
                })
        else:
            # 検索文字列で絞り込み
            search_lower = search_str.lower()
            
            for event_num_str, event_data in self.event_dictionary.items():
                if event_data.get('deprecated', False):
                    continue
                
                event_number = int(event_num_str)
                name_jp = event_data.get('name_jp', f'イベント{event_number}')
                alias = event_data.get('alias', '')
                
                # 数字入力の場合: event_number の前方一致
                if search_str.isdigit():
                    if event_num_str.startswith(search_str):
                        candidates.append({
                            'event_number': event_number,
                            'name_jp': name_jp,
                            'event_num_str': event_num_str
                        })
                else:
                    # 文字入力の場合: alias または name_jp の前方一致（大文字小文字を区別しない）
                    alias_match = alias.lower().startswith(search_lower) if alias else False
                    name_jp_match = name_jp.lower().startswith(search_lower) if name_jp else False
                    
                    if alias_match or name_jp_match:
                        candidates.append({
                            'event_number': event_number,
                            'name_jp': name_jp,
                            'event_num_str': event_num_str
                        })
        
        # 候補を番号順にソート
        candidates.sort(key=lambda x: x['event_number'])
        
        # プルダウンに表示する形式: "203: 乾乳"
        combo_values = [f"{c['event_number']}: {c['name_jp']}" for c in candidates]
        self.event_combo['values'] = combo_values
        self.event_candidates = candidates
        
        # 自動選択しない（候補表示のみ）
        # ユーザーは Enter/Tab で確定するか、Combobox から選択する
        if not auto_select:
            # 候補がある場合は最初の候補を表示（選択はしない）
            if candidates:
                # Combobox の表示のみ更新（選択はしない）
                pass
            else:
                # 候補がない場合
                self.event_combo.set("")
                self.selected_event = None
                self.selected_event_number = None
                self._create_input_form()
        else:
            # 旧来の動作（後方互換性のため残す）
            if len(candidates) == 1:
                self.event_combo.set(combo_values[0])
                self._select_event(candidates[0]['event_num_str'])
            elif len(candidates) > 1:
                self.event_combo.set("")
                self.selected_event = None
                self.selected_event_number = None
                self._create_input_form()
            else:
                self.event_combo.set("")
                self.selected_event = None
                self.selected_event_number = None
                self._create_input_form()
    
    def _open_combo_dropdown(self):
        """
        Comboboxのドロップダウンを自動的に開く
        
        注意: ttk.Comboboxは直接ドロップダウンを開くメソッドがないため、
        event_generate("<Down>")を使用する。
        ただし、Comboboxにフォーカスがない場合は動作しないため、
        一時的にフォーカスを移してから実行する。
        """
        try:
            # Comboboxにフォーカスを移す
            self.event_combo.focus_set()
            # ドロップダウンを開く（<Down>キーイベントを生成）
            self.event_combo.event_generate("<Down>")
            # 元のEntryにフォーカスを戻す（ユーザーが続けて入力できるように）
            self.window.after(50, lambda: self.event_number_entry.focus_set())
        except Exception as e:
            # エラーが発生しても処理を続行
            print(f"ドロップダウン自動表示エラー: {e}")
    
    def _on_event_combo_selected(self, event):
        """イベント候補プルダウンで選択された時の処理"""
        selection = self.event_combo.get()
        if not selection:
            return
        
        # "203: 乾乳" 形式から番号を抽出
        try:
            event_number = int(selection.split(':')[0].strip())
            event_num_str = str(event_number)
            self._select_event(event_num_str)
        except (ValueError, IndexError):
            pass
    
    def _on_cow_id_changed(self, event):
        """牛ID入力変更時の処理（リアルタイム検証は行わない）"""
        pass
    
    def _on_cow_id_enter(self):
        """牛ID入力欄でEnterキー押下時またはフォーカスアウト時"""
        if not hasattr(self, 'cow_id_entry'):
            return
        
        cow_id_str = self.cow_id_entry.get().strip()
        if not cow_id_str:
            return
        
        # 牛IDを解決
        resolved_auto_id = self._resolve_cow_id(cow_id_str)
        if resolved_auto_id:
            self.cow_auto_id = resolved_auto_id
            # 対象牛情報を表示
            self._show_cow_info()
            # イベント履歴を更新
            self._load_event_history()
            # フォーカスをイベント番号入力欄に移動
            self.event_number_entry.focus()
        else:
            messagebox.showerror("エラー", f"牛が見つかりません: {cow_id_str}")
            self.cow_id_entry.focus()
    
    def _resolve_cow_id(self, cow_id_str: str) -> Optional[int]:
        """
        牛ID文字列から auto_id を解決
        
        Args:
            cow_id_str: 拡大4桁ID または 個体識別番号10桁
        
        Returns:
            auto_id、見つからない場合はNone
        """
        cow_id_str = cow_id_str.strip()
        
        # 拡大4桁IDで検索
        if cow_id_str.isdigit() and len(cow_id_str) <= 4:
            padded_id = cow_id_str.zfill(4)
            cow = self.db.get_cow_by_id(padded_id)
            if cow:
                return cow.get('auto_id')
        
        # 個体識別番号10桁で検索
        if cow_id_str.isdigit() and len(cow_id_str) == 10:
            cows = self.db.get_all_cows()
            for cow in cows:
                if cow.get('jpn10') == cow_id_str:
                    return cow.get('auto_id')
        
        return None
    
    def _show_cow_info(self):
        """対象牛情報を表示"""
        if self.cow_auto_id is None:
            return
        
        cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
        if cow:
            cow_id = cow.get('cow_id', '')
            info_text = f"対象牛：{cow_id}"
            
            # 既存の情報表示があれば更新
            if hasattr(self, 'cow_info_label'):
                self.cow_info_label.config(text=info_text, foreground="blue")
    
    def _select_event(self, event_num_str: str):
        """
        イベントを選択
        
        Args:
            event_num_str: イベント番号（文字列）
        """
        event_data = self.event_dictionary.get(event_num_str)
        if not event_data or event_data.get('deprecated', False):
            return
        
        self.selected_event = event_data
        self.selected_event_number = int(event_num_str)
        
        # 入力フォームを再生成
        self._create_input_form()
    
    def _create_input_form(self):
        """入力フォームを動的生成"""
        # 既存のフォームをクリア
        for widget in self.form_frame.winfo_children():
            widget.destroy()
        self.field_widgets.clear()
        
        if not self.selected_event:
            ttk.Label(
                self.form_frame,
                text="イベントを選択してください",
                foreground="gray"
            ).pack(pady=10)
            return
        
        event_number = self.selected_event_number

        # 分娩イベント（202）の特別処理
        if event_number == RuleEngine.EVENT_CALV:
            self._create_calving_form()
            return
        
        # AI/ETイベント（200/201）の特別処理
        if event_number in [200, 201]:  # AI or ET
            self._create_ai_et_form()
            return
        
        # 通常のイベント処理
        input_fields = self.selected_event.get('input_fields', [])
        
        if not input_fields:
            ttk.Label(
                self.form_frame,
                text="入力項目はありません",
                foreground="gray"
            ).pack(pady=10)
            return
        
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
            entry = ttk.Entry(self.form_frame, width=30)
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
            self.field_widgets[key] = entry
            entry_widgets.append(entry)
            
            # 繁殖検査（301）、フレッシュチェック（300）、妊娠鑑定マイナス（302）の特定フィールドに半角大文字変換機能を追加
            if event_number in [300, 301, 302] and key in ['treatment', 'uterine_findings', 'left_ovary_findings', 'right_ovary_findings', 'other']:
                # 入力時にリアルタイムで大文字に変換
                entry.bind('<KeyPress>', lambda e, widget=entry: self._convert_to_uppercase_on_input(e, widget))
                entry.bind('<KeyRelease>', lambda e, widget=entry: self._convert_to_uppercase(widget))
        
        # Enterキーで次のフィールドに移動する機能を追加
        for i, entry in enumerate(entry_widgets):
            if i < len(entry_widgets) - 1:
                # 最後のフィールド以外は次のフィールドに移動
                entry.bind('<Return>', lambda e, next_widget=entry_widgets[i + 1]: self._move_to_next_field(e, next_widget))
            else:
                # 最後のフィールドはメモ欄に移動（メモ欄がある場合）
                entry.bind('<Return>', lambda e: self._move_to_note_field(e))
    
    # ========== 分娩専用UI ==========
    def _get_calving_difficulty_options(self) -> List[Dict[str, str]]:
        """event_dictionaryの定義から分娩難易度の候補を取得"""
        default_map = {
            "1": "自然分娩",
            "2": "介助",
            "3": "難産",
            "4": "獣医師による難産",
            "5": "帝王切開"
        }
        event_def = self.event_dictionary.get(str(RuleEngine.EVENT_CALV), {})
        diff_map = event_def.get("calving_difficulty") or default_map
        
        options = []
        for code in sorted(diff_map.keys(), key=lambda x: int(x) if str(x).isdigit() else 999):
            options.append({"code": str(code), "label": diff_map.get(code, str(code))})
        return options

    def _get_calving_breed_options(self) -> List[str]:
        """子牛品種の候補"""
        return ["ホルスタイン", "ジャージー", "その他乳用種", "F1", "黒毛和種", "その他肉用種"]

    def _create_calving_form(self):
        """分娩イベント専用フォームを生成"""
        # 既存の分娩用状態を初期化
        self.calving_difficulty_var = tk.StringVar(value="")
        self.calving_child_count_var = tk.IntVar(value=1)
        self.calving_calf_vars = []
        
        container = ttk.Frame(self.form_frame)
        container.pack(fill=tk.BOTH, expand=True)
        
        # 難易度
        diff_frame = ttk.Frame(container)
        diff_frame.pack(fill=tk.X, pady=5)
        ttk.Label(diff_frame, text="分娩難易度:").pack(side=tk.LEFT, padx=5)
        diff_options = self._get_calving_difficulty_options()
        diff_values = [f"{opt['code']}: {opt['label']}" for opt in diff_options]
        diff_combo = ttk.Combobox(
            diff_frame,
            textvariable=self.calving_difficulty_var,
            values=diff_values,
            state="readonly",
            width=30
        )
        if diff_values:
            diff_combo.set(diff_values[0])
        diff_combo.pack(side=tk.LEFT, padx=5)
        
        # 頭数
        count_frame = ttk.Frame(container)
        count_frame.pack(fill=tk.X, pady=5)
        ttk.Label(count_frame, text="子牛頭数:").pack(side=tk.LEFT, padx=5)
        for count, label in [(1, "単子"), (2, "双子"), (3, "三つ子")]:
            rb = ttk.Radiobutton(
                count_frame,
                text=label,
                value=count,
                variable=self.calving_child_count_var,
                command=lambda c=count: self._update_calf_blocks(c)
            )
            rb.pack(side=tk.LEFT, padx=5)
        
        # 子牛ブロックコンテナ
        self.calving_block_container = ttk.Frame(container)
        self.calving_block_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 初期状態として単子を生成
        self._update_calf_blocks(1)

    def _update_calf_blocks(self, count: int):
        """子牛入力ブロックを頭数に応じて生成"""
        # 既存のウィジェットをクリア
        for widget in self.calving_block_container.winfo_children():
            widget.destroy()
        self.calving_calf_vars = []
        
        breeds = self._get_calving_breed_options()
        
        for idx in range(count):
            block = ttk.LabelFrame(self.calving_block_container, text=f"子牛{idx + 1}", padding=5)
            block.pack(fill=tk.X, pady=5)
            
            # 品種
            breed_var = tk.StringVar(value="")
            ttk.Label(block, text="品種:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
            breed_combo = ttk.Combobox(block, textvariable=breed_var, values=breeds, state="readonly", width=25)
            breed_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
            
            # 性別
            sex_var = tk.StringVar(value="")
            ttk.Label(block, text="性別:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
            sex_frame = ttk.Frame(block)
            sex_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
            ttk.Radiobutton(sex_frame, text="オス", value="M", variable=sex_var).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Radiobutton(sex_frame, text="メス", value="F", variable=sex_var).pack(side=tk.LEFT)
            
            # 死産
            stillborn_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(block, text="死産", variable=stillborn_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
            
            self.calving_calf_vars.append({
                "breed_var": breed_var,
                "sex_var": sex_var,
                "stillborn_var": stillborn_var
            })
        
        # 画面上の頭数表示を同期（Radiobuttonの状態も更新）
        if self.calving_child_count_var:
            self.calving_child_count_var.set(count)

    def _load_calving_data(self, json_data: Dict[str, Any]):
        """既存の分娩イベントデータをUIに反映"""
        if not self.calving_block_container:
            return
        
        # 難易度
        diff = json_data.get("calving_difficulty")
        if self.calving_difficulty_var and diff is not None:
            options = self._get_calving_difficulty_options()
            diff_map = {opt["code"]: opt for opt in options}
            diff_str = str(diff)
            if diff_str in diff_map:
                self.calving_difficulty_var.set(f"{diff_str}: {diff_map[diff_str]['label']}")
        
        # 子牛
        calves = json_data.get("calves") or []
        count = min(max(len(calves), 1), 3)
        self._update_calf_blocks(count)
        
        for idx, calf in enumerate(calves[:count]):
            vars_dict = self.calving_calf_vars[idx]
            if calf.get("breed"):
                vars_dict["breed_var"].set(calf["breed"])
            if calf.get("sex"):
                vars_dict["sex_var"].set(str(calf["sex"]).upper())
            vars_dict["stillborn_var"].set(bool(calf.get("stillborn", False)))

    def _validate_calving_input(self) -> bool:
        """分娩入力の検証"""
        # baseline分娩（インポート由来など）の編集時は既存値を尊重
        if self.event_id is not None and self._editing_event_json.get("baseline_calving") and not self._editing_event_json.get("calves"):
            return True
        
        # 難易度
        if self.calving_difficulty_var:
            diff_val = self.calving_difficulty_var.get().strip()
            if not diff_val:
                messagebox.showerror("エラー", "分娩難易度を選択してください")
                return False
        
        # 子牛ブロック
        if not self.calving_calf_vars:
            messagebox.showerror("エラー", "子牛情報を入力してください")
            return False
        
        for idx, vars_dict in enumerate(self.calving_calf_vars, start=1):
            breed = vars_dict["breed_var"].get().strip()
            sex = vars_dict["sex_var"].get().strip()
            if not breed:
                messagebox.showerror("エラー", f"子牛{idx}の品種を選択してください")
                return False
            if not sex:
                messagebox.showerror("エラー", f"子牛{idx}の性別を選択してください")
                return False
        return True

    def _collect_calving_json(self) -> Dict[str, Any]:
        """分娩イベントのjson_dataを構築"""
        json_data: Dict[str, Any] = {}
        
        # 既存のjson_dataに残しておきたい値（baseline_calvingなど）を引き継ぐ
        base = {}
        if self._editing_event_json:
            for k, v in self._editing_event_json.items():
                if k not in ["calving_difficulty", "calves"]:
                    base[k] = v
        json_data.update(base)
        
        # 難易度
        diff_val = None
        if self.calving_difficulty_var:
            raw = self.calving_difficulty_var.get()
            if "：" in raw:
                raw = raw.split("：", 1)[0]
            elif ":" in raw:
                raw = raw.split(":", 1)[0]
            diff_val = raw.strip()
            if diff_val.isdigit():
                diff_val = int(diff_val)
        if diff_val is not None:
            json_data["calving_difficulty"] = diff_val
        
        # 子牛情報
        calves = []
        for vars_dict in self.calving_calf_vars:
            breed = vars_dict["breed_var"].get().strip()
            sex = vars_dict["sex_var"].get().strip()
            stillborn = bool(vars_dict["stillborn_var"].get())
            if not breed and not sex and not stillborn:
                continue
            calves.append({
                "breed": breed,
                "sex": sex,
                "stillborn": stillborn
            })
        json_data["calves"] = calves
        
        return json_data

    def _handle_calving_followups(self, calving_event_id: int, calving_date: str, json_data: Dict[str, Any]):
        """分娩保存後の派生処理（導入イベント自動生成など）"""
        calves = json_data.get("calves") or []
        if not calves:
            return
        
        dairy_breeds = {"ホルスタイン", "ジャージー", "その他乳用種"}
        
        eligible: List[Dict[str, Any]] = []
        for idx, calf in enumerate(calves, start=1):
            if calf.get("stillborn"):
                continue
            if str(calf.get("sex")).upper() != "F":
                continue
            if calf.get("breed") not in dairy_breeds:
                continue
            eligible.append({"idx": idx, "data": calf})
        
        if not eligible:
            return
        
        mother = self.db.get_cow_by_auto_id(self.cow_auto_id) if self.cow_auto_id else None
        mother_id = mother.get("cow_id") if mother else None
        mother_jpn10 = mother.get("jpn10") if mother else None
        
        created = 0
        
        for item in eligible:
            idx = item["idx"]
            calf = item["data"]
            
            # 双子・三つ子の場合は登録確認を表示（デフォルトNO）
            if len(calves) > 1:
                calf_desc = f"{calf.get('breed', '')}{'♀' if str(calf.get('sex')).upper() == 'F' else calf.get('sex', '')}"
                confirm = messagebox.askyesno(
                    "確認",
                    f"この子牛を牛群に登録しますか？\n{calf_desc}",
                    default='no'
                )
                if not confirm:
                    continue
            
            intro_json = {
                "calf_index": idx,
                "calf_breed": calf.get("breed"),
                "calf_sex": calf.get("sex"),
                "calf_birth_date": calving_date,
                "calf_mother_id": mother_id,
                "calf_mother_jpn10": mother_jpn10,
                "calving_event_id": calving_event_id,
                "source": "auto_from_calving"
            }
            
            intro_event = {
                "cow_auto_id": self.cow_auto_id,
                "event_number": 600,  # 導入イベント
                "event_date": calving_date,
                "json_data": intro_json,
                "note": f"自動生成: 子牛{idx}登録"
            }
            intro_event_id = self.db.insert_event(intro_event)
            self.rule_engine.on_event_added(intro_event_id)
            created += 1
        
        if created:
            messagebox.showinfo("導入イベント", f"{created}件の導入イベントを自動生成しました。")
    
    def _create_ai_et_form(self):
        """AI/ETイベント専用の入力フォームを作成（Comboboxのみ）"""
        row = 0
        
        # 1. SIRE（文字列入力）
        ttk.Label(self.form_frame, text="SIRE:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        sire_entry = ttk.Entry(self.form_frame, width=30)
        sire_entry.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        self.field_widgets['sire'] = sire_entry
        # Enterキーで授精師コード入力欄に移動
        sire_entry.bind('<Return>', lambda e: self.field_widgets.get('technician_code', sire_entry).focus())
        row += 1
        
        # 2. 授精師コード（Comboboxのみ）
        ttk.Label(self.form_frame, text="授精師コード:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Comboboxのvaluesを生成
        technician_values = [
            f"{code}：{name}"
            for code, name in sorted(self.technicians.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999)
        ]
        
        technician_combo = ttk.Combobox(
            self.form_frame,
            values=technician_values,
            width=28,
            state='normal'  # editable にしてキーボード入力可能に
        )
        technician_combo.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        self.field_widgets['technician_code'] = technician_combo
        # キー入力で候補を絞り込み
        technician_combo.bind('<KeyRelease>', lambda e: self._on_technician_key(e))
        # Enter/Tabキーで確定して次のフィールドに移動
        technician_combo.bind('<Return>', lambda e: self._on_technician_enter(e))
        technician_combo.bind('<Tab>', lambda e: self._on_technician_tab(e))
        row += 1
        
        # 3. 授精種類コード（Comboboxのみ）
        ttk.Label(self.form_frame, text="授精種類コード:").grid(row=row, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Comboboxのvaluesを生成
        insemination_type_values = [
            f"{code}：{name}"
            for code, name in sorted(self.insemination_types.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999)
        ]
        
        insemination_type_combo = ttk.Combobox(
            self.form_frame,
            values=insemination_type_values,
            width=28,
            state='normal'  # editable にしてキーボード入力可能に
        )
        insemination_type_combo.grid(row=row, column=1, sticky=tk.W, padx=5, pady=5)
        self.field_widgets['insemination_type_code'] = insemination_type_combo
        # キー入力で候補を絞り込み
        insemination_type_combo.bind('<KeyRelease>', lambda e: self._on_insemination_type_key(e))
        # Enter/Tabキーで確定して次のフィールドに移動
        insemination_type_combo.bind('<Return>', lambda e: self._on_insemination_type_enter(e))
        insemination_type_combo.bind('<Tab>', lambda e: self._on_insemination_type_tab(e))
    
    def _on_technician_key(self, event):
        """授精師コード入力時の候補絞り込み"""
        # 特殊キー（Enter, Tab, Shift等）の場合は処理しない
        if event.keysym in ('Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Up', 'Down'):
            return
        
        text = event.widget.get().strip().upper()
        self._filter_combobox_values(event.widget, text, self.technicians)
    
    def _on_technician_enter(self, event):
        """授精師コードでEnterキーが押された時の処理（確定）"""
        text = event.widget.get().strip().upper()
        matches = self._filter_combobox_values(event.widget, text, self.technicians)
        
        # 完全一致が1件のみなら自動確定
        if len(matches) == 1:
            event.widget.set(matches[0])
        
        # 次のフィールドに移動
        self.field_widgets.get('insemination_type_code', event.widget).focus()
        return "break"
    
    def _on_technician_tab(self, event):
        """授精師コードでTabキーが押された時の処理（確定）"""
        text = event.widget.get().strip().upper()
        matches = self._filter_combobox_values(event.widget, text, self.technicians)
        
        # 完全一致が1件のみなら自動確定
        if len(matches) == 1:
            event.widget.set(matches[0])
        
        # 通常のTab動作を許可
        return None
    
    def _on_insemination_type_key(self, event):
        """授精種類コード入力時の候補絞り込み"""
        # 特殊キー（Enter, Tab, Shift等）の場合は処理しない
        if event.keysym in ('Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Up', 'Down'):
            return
        
        text = event.widget.get().strip().upper()
        self._filter_combobox_values(event.widget, text, self.insemination_types)
    
    def _on_insemination_type_enter(self, event):
        """授精種類コードでEnterキーが押された時の処理（確定）"""
        text = event.widget.get().strip().upper()
        matches = self._filter_combobox_values(event.widget, text, self.insemination_types)
        
        # 完全一致が1件のみなら自動確定
        if len(matches) == 1:
            event.widget.set(matches[0])
        
        # メモ欄に移動
        self._move_to_note_field(event)
        return "break"
    
    def _on_insemination_type_tab(self, event):
        """授精種類コードでTabキーが押された時の処理（確定）"""
        text = event.widget.get().strip().upper()
        matches = self._filter_combobox_values(event.widget, text, self.insemination_types)
        
        # 完全一致が1件のみなら自動確定
        if len(matches) == 1:
            event.widget.set(matches[0])
        
        # 通常のTab動作を許可
        return None
    
    def _filter_combobox_values(self, combo_widget, prefix: str, source_dict: Dict[str, str]) -> List[str]:
        """
        Combobox の候補を prefix で絞り込む
        
        Args:
            combo_widget: Combobox ウィジェット
            prefix: 検索プレフィックス（コードまたは名称の先頭）
            source_dict: ソース辞書（{"code": "name", ...}）
        
        Returns:
            マッチした候補リスト（"code：name" 形式）
        """
        if not prefix:
            # 空欄の場合は全候補を表示
            all_values = [
                f"{code}：{name}"
                for code, name in sorted(source_dict.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999)
            ]
            combo_widget['values'] = all_values
            return all_values
        
        matches = []
        prefix_upper = prefix.upper()
        
        for code, name in source_dict.items():
            code_str = str(code).upper()
            name_str = str(name).upper()
            value_str = f"{code}：{name}"
            
            # コードまたは名称の先頭が一致するか
            if code_str.startswith(prefix_upper) or name_str.startswith(prefix_upper):
                matches.append(value_str)
        
        # ソート（コード順）
        matches.sort(key=lambda x: int(x.split('：')[0]) if x.split('：')[0].isdigit() else 999)
        
        # Combobox の候補を更新
        combo_widget['values'] = matches
        
        return matches
    
    def _set_widget_value(self, w, value: str) -> None:
        """
        Entry または Text ウィジェットに安全に値を設定する
        
        Args:
            w: ウィジェット（tk.Entry, ttk.Entry, tk.Text など）
            value: 設定する値（None の場合は空文字列）
        """
        s = "" if value is None else str(value)
        
        # Text widget の判定
        if isinstance(w, tk.Text):
            try:
                w.delete("1.0", "end")
                w.insert("1.0", s)
                return
            except Exception:
                pass
        
        # Entry widget (tk.Entry / ttk.Entry)
        try:
            w.delete(0, tk.END)
            w.insert(0, s)  # Entry は必ず index=0
        except Exception:
            # 最後の手段：何もしない（クラッシュ回避）
            pass
    
    def _load_event_data_for_edit(self):
        """編集時に既存イベントデータを読み込む"""
        # イベントを取得（削除済みも含む）
        event = self.db.get_event_by_id(self.event_id, include_deleted=True)
        
        if not event:
            messagebox.showerror(
                "エラー",
                "編集対象のイベントが見つかりません"
            )
            self.window.destroy()
            return
        
        # cow_auto_id が設定されていない場合は、イベントから取得
        if self.cow_auto_id is None:
            self.cow_auto_id = event.get('cow_auto_id')
        
        # ========== 1. 共通項目を入力欄に反映 ==========
        event_number = event.get('event_number')
        event_date = event.get('event_date', '')
        note = event.get('note', '')
        
        # イベント番号を Entry に反映
        self.event_number_entry.delete(0, tk.END)
        if event_number:
            self.event_number_entry.insert(0, str(event_number))
        
        # イベントを選択（これにより input_fields が設定される）
        self._select_event(str(event_number))
        
        # イベント候補 Combobox に反映
        if self.selected_event:
            event_name = self.selected_event.get('name_jp', f'イベント{event_number}')
            self.event_combo.set(f"{event_number}: {event_name}")
        
        # 日付を設定
        self.date_entry.delete(0, tk.END)
        if event_date:
            self.date_entry.insert(0, event_date)
        
        # メモを設定
        self.note_entry.delete(0, tk.END)
        if note:
            self.note_entry.insert(0, note)
        
        # ========== 2. json_data から入力フィールドを設定 ==========
        json_data = event.get('json_data') or {}
        # 編集時の元データを保持（baseline_calving等を引き継ぐため）
        self._editing_event_json = json_data.copy()
        
        # 分娩イベントは専用ロジックで復元
        if event_number == RuleEngine.EVENT_CALV:
            self._load_calving_data(json_data)
            return
        
        # AIイベント（200, 201）の場合は専用項目を明示的に処理
        if event_number in [200, 201]:  # AI, ET
            # SIRE
            sire_widget = self.field_widgets.get('sire')
            if sire_widget:
                sire_value = json_data.get('sire', '')
                self._set_widget_value(sire_widget, sire_value)
            
            # 授精師コード
            tech_widget = self.field_widgets.get('technician_code')
            if tech_widget:
                tech_code = json_data.get('technician_code')
                if tech_code:
                    # コードから名称を取得して「code：name」形式で設定
                    name = self.technicians.get(str(tech_code), '')
                    if name:
                        tech_widget.set(f"{tech_code}：{name}")
                    else:
                        tech_widget.set(str(tech_code))
                else:
                    tech_widget.set('')
            
            # 授精種類コード
            type_widget = self.field_widgets.get('insemination_type_code')
            if type_widget:
                type_code = json_data.get('insemination_type_code')
                if type_code:
                    # コードから名称を取得して「code：name」形式で設定
                    name = self.insemination_types.get(str(type_code), '')
                    if name:
                        type_widget.set(f"{type_code}：{name}")
                    else:
                        type_widget.set(str(type_code))
                else:
                    type_widget.set('')
        
        # ========== 3. その他の入力フィールドを設定（input_fields に基づく） ==========
        input_fields = self.selected_event.get('input_fields', [])
        for field in input_fields:
            key = field.get('key')
            
            # AIイベント固有項目は既に処理済みなのでスキップ
            if key in ('sire', 'technician_code', 'insemination_type_code'):
                continue
            
            widget = self.field_widgets.get(key)
            if not widget:
                continue
            
            value = json_data.get(key)
            if value is None:
                continue
            
            # Combobox の場合
            if isinstance(widget, ttk.Combobox):
                # コードから名称を取得して「code：name」形式で設定
                if key == 'technician_code':
                    name = self.technicians.get(str(value), '')
                    if name:
                        widget.set(f"{value}：{name}")
                    else:
                        widget.set(str(value))
                elif key == 'insemination_type_code':
                    name = self.insemination_types.get(str(value), '')
                    if name:
                        widget.set(f"{value}：{name}")
                    else:
                        widget.set(str(value))
                else:
                    # その他の Combobox は値そのまま
                    widget.set(str(value))
            else:
                # Entry または Text の場合（ヘルパーを使用）
                self._set_widget_value(widget, str(value))
    
    def _load_event_history(self, select_event_id: Optional[int] = None):
        """イベント履歴を表示（event_date DESC順）

        Args:
            select_event_id: 追加・更新直後に選択状態にするイベントID
        """
        if self.cow_auto_id is None:
            return
        
        # 既存のアイテムをクリア
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # イベントを取得（既にevent_date DESC順でソート済み）
        events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
        
        if not events:
            # イベント履歴がない場合は、tagsを付けない（メニューを表示しない）
            self.history_tree.insert('', 'end', values=("", "イベント履歴がありません", ""), tags=('no_event',))
            return
        
        # Treeviewに追加
        for event in events:
            event_date = event.get('event_date', '')
            event_number = event.get('event_number')
            
            # event_number から日本語名を取得
            if event_number is not None:
                event_num_str = str(event_number)
                event_def = self.event_dictionary.get(event_num_str)
                if event_def:
                    display_name = event_def.get("name_jp", f"イベント{event_number}")
                else:
                    display_name = f"イベント{event_number}"
            else:
                display_name = "イベント不明"
            
            # 簡易情報（json_dataの主要項目を表示）
            json_data = event.get('json_data') or {}
            info_parts = []
            
            # AI/ETイベント（200/201）の特別表示
            if event_number in [200, 201]:
                # 授精師・授精種類の辞書を取得（SettingsManagerから取得、なければ空辞書）
                tech_dict = {}
                type_dict = {}
                if self.settings_manager:
                    tech_dict = self.settings_manager.get_inseminator_codes()
                    type_dict = self.settings_manager.get_insemination_type_codes()
                
                # 共通関数を使用して表示文字列を生成
                # technicians_dict / insemination_types_dict は必ず渡す（空辞書でも可）
                detail_text = format_insemination_event(json_data, tech_dict, type_dict)
                if detail_text:
                    info_parts.append(detail_text)
            elif event_number == RuleEngine.EVENT_CALV:
                calv_def = self.event_dictionary.get(str(RuleEngine.EVENT_CALV), {})
                diff_labels = calv_def.get("calving_difficulty", {})
                detail_text = format_calving_event(json_data, diff_labels)
                if detail_text:
                    info_parts.append(detail_text)
            # 繁殖検査（301）、妊娠鑑定マイナス（302）、フレッシュチェック（300）の特別表示
            elif event_number in [300, 301, 302]:
                # 処置（新しいキー名と古いキー名の両方をサポート）
                treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
                treatment_display = treatment if treatment else '-'
                info_parts.append(treatment_display)
                
                # 子宮所見（新しいキー名と古いキー名の両方をサポート）
                uterine = (json_data.get('uterine_findings') or 
                          json_data.get('uterus_findings') or 
                          json_data.get('uterus_finding') or 
                          json_data.get('uterus', ''))
                uterine_display = f"子宮{uterine}" if uterine else "子宮-"
                info_parts.append(uterine_display)
                
                # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
                right_ovary = (json_data.get('right_ovary_findings') or 
                              json_data.get('rightovary_findings') or 
                              json_data.get('rightovary_finding') or 
                              json_data.get('right_ovary', '') or
                              json_data.get('rightovary', ''))
                right_display = f"右{right_ovary}" if right_ovary else "右-"
                info_parts.append(right_display)
                
                # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
                left_ovary = (json_data.get('left_ovary_findings') or 
                             json_data.get('leftovary_findings') or 
                             json_data.get('leftovary_finding') or 
                             json_data.get('left_ovary', '') or
                             json_data.get('leftovary', ''))
                left_display = f"左{left_ovary}" if left_ovary else "左-"
                info_parts.append(left_display)
                
                # その他（新しいキー名と古いキー名の両方をサポート）
                other = (json_data.get('other') or 
                        json_data.get('other_info') or 
                        json_data.get('other_findings', ''))
                other_display = other if other else '-'
                info_parts.append(other_display)
            else:
                # 通常のイベント
                if 'sire' in json_data:
                    info_parts.append(f"SIRE:{json_data['sire']}")
                if 'milk_yield' in json_data:
                    info_parts.append(f"乳量:{json_data['milk_yield']}")
                if 'to_pen' in json_data:
                    info_parts.append(f"→{json_data['to_pen']}")
            
            # 繁殖検査、妊娠鑑定マイナス、フレッシュチェックの場合は「/」で区切る
            if event_number in [300, 301, 302]:
                info = "/".join(info_parts) if info_parts else ""
            else:
                info = " / ".join(info_parts) if info_parts else ""
            
            # イベントIDをtagsに保存
            event_id = event.get('id')
            item_id = self.history_tree.insert('', 'end', values=(event_date, display_name, info), tags=(f"event_{event_id}",))
            
            # 追加・更新したイベントを自動選択
            if select_event_id is not None and event_id == select_event_id:
                self.history_tree.selection_set(item_id)
                self.history_tree.see(item_id)
    
    def _validate_input(self) -> bool:
        """入力値を検証"""
        # 牛IDの検証（メニュー起動時のみ）
        if self.cow_auto_id is None:
            if not hasattr(self, 'cow_id_entry'):
                messagebox.showerror("エラー", "牛ID入力欄が見つかりません")
                return False
            
            cow_id_str = self.cow_id_entry.get().strip()
            if not cow_id_str:
                messagebox.showerror("エラー", "牛IDを入力してください")
                self.cow_id_entry.focus()
                return False
            
            resolved_auto_id = self._resolve_cow_id(cow_id_str)
            if not resolved_auto_id:
                messagebox.showerror("エラー", f"牛が見つかりません: {cow_id_str}")
                self.cow_id_entry.focus()
                return False
            
            self.cow_auto_id = resolved_auto_id
        
        # 日付の検証
        date_str = self.date_entry.get().strip()
        normalized_date = normalize_date(date_str)
        if not normalized_date:
            messagebox.showerror("エラー", "日付の形式が正しくありません")
            self.date_entry.focus()
            return False
        
        # イベント選択の検証
        if not self.selected_event or not self.selected_event_number:
            messagebox.showerror("エラー", "イベントを選択してください")
            return False
        
        # 分娩イベントは専用の検証を行う
        if self.selected_event_number == RuleEngine.EVENT_CALV:
            return self._validate_calving_input()
        
        # 入力フィールドの検証
        input_fields = self.selected_event.get('input_fields', [])
        for field in input_fields:
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            widget = self.field_widgets.get(key)
            
            if widget:
                value = widget.get().strip()
                
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
    
    def _on_ok(self):
        """OKボタンクリック時の処理"""
        if not self._validate_input():
            return
        
        # 日付を正規化
        date_str = self.date_entry.get().strip()
        normalized_date = normalize_date(date_str)
        
        # メモを取得
        note = self.note_entry.get().strip()
        
        # json_data を必ず新規構築（新規・編集で同じロジック）
        if self.selected_event_number == RuleEngine.EVENT_CALV:
            json_data = self._collect_calving_json()
        else:
            json_data = {}
            
            # 1. input_fields から動的に処理
            input_fields = self.selected_event.get('input_fields', [])
            
            for field in input_fields:
                key = field.get('key')
                datatype = field.get('datatype', 'str')
                widget = self.field_widgets.get(key)
                
                if not widget:
                    continue
                
                value = widget.get().strip()
                if not value:
                    continue
                
                # Combobox の場合、「コード：名称」形式からコード部分のみを抽出
                if isinstance(widget, ttk.Combobox):
                    if "：" in value:
                        code = value.split("：", 1)[0]
                    else:
                        code = value
                    json_data[key] = code
                else:
                    # Entry の場合、型変換
                    if datatype == 'int':
                        json_data[key] = int(value)
                    elif datatype == 'float':
                        json_data[key] = float(value)
                    else:
                        json_data[key] = value
            
            # 2. AIイベント（200, 201）の場合は専用項目を必ず詰める
            if self.selected_event_number in [200, 201]:  # AI, ET
                # SIRE
                sire_widget = self.field_widgets.get('sire')
                if sire_widget:
                    sire_value = sire_widget.get().strip()
                    if sire_value:
                        json_data['sire'] = sire_value
                    else:
                        # 空の場合は None を設定（削除の意味）
                        json_data['sire'] = None
                
                # 授精師コード
                tech_widget = self.field_widgets.get('technician_code')
                if tech_widget:
                    tech_value = tech_widget.get().strip()
                    if tech_value:
                        # 「コード：名称」形式からコード部分のみを抽出
                        if "：" in tech_value:
                            tech_code = tech_value.split("：", 1)[0]
                        else:
                            tech_code = tech_value
                        json_data['technician_code'] = tech_code
                    else:
                        # 空の場合は None を設定（削除の意味）
                        json_data['technician_code'] = None
                
                # 授精種類コード
                type_widget = self.field_widgets.get('insemination_type_code')
                if type_widget:
                    type_value = type_widget.get().strip()
                    if type_value:
                        # 「コード：名称」形式からコード部分のみを抽出
                        if "：" in type_value:
                            type_code = type_value.split("：", 1)[0]
                        else:
                            type_code = type_value
                        json_data['insemination_type_code'] = type_code
                    else:
                        # 空の場合は None を設定（削除の意味）
                        json_data['insemination_type_code'] = None
        
        # 3. json_data を正規化（必ず dict にする、空の場合は {}）
        # None や空文字列の値は削除
        json_data = {k: v for k, v in json_data.items() if v is not None and v != ""}
        
        # 空の場合は {} を保存（None ではない）
        if not json_data:
            json_data = {}
        
        # デバッグログ
        logging.debug(f"Event saved: event_number={self.selected_event_number}, json_data={json_data}")
        
        try:
            saved_event_id: Optional[int] = None
            saved_event_number = self.selected_event_number
            
            if self.event_id:
                # 更新（新規・編集で同じ json_data を使用）
                event_data = {
                    'event_number': self.selected_event_number,
                    'event_date': normalized_date,
                    'json_data': json_data,  # 正規化済み（必ず dict）
                    'note': note if note else None
                }
                self.db.update_event(self.event_id, event_data)
                self.rule_engine.on_event_updated(self.event_id)
                saved_event_id = self.event_id
            else:
                # 新規追加（新規・編集で同じ json_data を使用）
                event_data = {
                    'cow_auto_id': self.cow_auto_id,
                    'event_number': self.selected_event_number,
                    'event_date': normalized_date,
                    'json_data': json_data,  # 正規化済み（必ず dict）
                    'note': note if note else None
                }
                event_id = self.db.insert_event(event_data)
                self.rule_engine.on_event_added(event_id)
                saved_event_id = event_id
                
                # 分娩時の派生処理（子牛導入イベントなど）
                if self.selected_event_number == RuleEngine.EVENT_CALV:
                    self._handle_calving_followups(event_id, normalized_date, json_data)
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            # 履歴を更新（直前のイベントを選択）
            if self.cow_auto_id:
                self._load_event_history(select_event_id=saved_event_id)

            # 編集モードを解除し、次入力のためにフォームを初期化
            self.event_id = None
            self._editing_event_json = {}
            self._reset_form_for_next_input(keep_date=normalized_date)

            logging.info(f"Event saved (continuous mode): cow_auto_id={self.cow_auto_id} event_number={saved_event_number}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"イベントの保存に失敗しました: {e}")

    def _reset_form_for_next_input(self, keep_date: Optional[str] = None):
        """
        連続入力用にフォームを初期化（牛IDと日付は保持）

        Args:
            keep_date: 日付欄に残す値（省略時は現在値を維持）
        """
        # イベント選択状態をクリア
        self.selected_event = None
        self.selected_event_number = None
        self.event_candidates = []
        self._editing_event_json = {}

        # 入力欄リセット（牛IDはそのまま）
        self.event_number_entry.delete(0, tk.END)
        self.event_combo.set("")

        if keep_date:
            self.date_entry.delete(0, tk.END)
            self.date_entry.insert(0, keep_date)

        self.note_entry.delete(0, tk.END)

        # 動的フォームを初期化
        self._create_input_form()

        # フォーカスをイベント番号へ
        self.event_number_entry.focus_set()

        logging.info("EventInputWindow reset for next input")
    
    def _on_delete(self):
        """削除ボタンクリック時の処理"""
        if not self.event_id:
            return
        
        result = messagebox.askyesno("確認", "このイベントを削除しますか？")
        if not result:
            return
        
        try:
            # 論理削除
            self.db.delete_event(self.event_id, soft_delete=True)
            self.rule_engine.on_event_deleted(self.event_id)
            
            # イベント履歴を更新
            if self.cow_auto_id:
                self._load_event_history()
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            # ウィンドウを閉じる
            self.window.destroy()
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.window.destroy()
    
    def _convert_to_uppercase_on_input(self, event, widget: tk.Widget):
        """
        入力時にリアルタイムで半角大文字に変換（全角日本語は保持）
        
        Args:
            event: キーイベント
            widget: Entryウィジェット
        """
        if not isinstance(widget, tk.Entry):
            return
        
        # 入力された文字を取得
        char = event.char
        if not char or len(char) != 1:
            return
        
        # 小文字の半角英字を大文字に変換
        if 'a' <= char <= 'z':
            # 現在のカーソル位置を取得
            cursor_pos = widget.index(tk.INSERT)
            current_text = widget.get()
            
            # カーソル位置に大文字を挿入
            new_text = current_text[:cursor_pos] + char.upper() + current_text[cursor_pos:]
            widget.delete(0, tk.END)
            widget.insert(0, new_text)
            widget.icursor(cursor_pos + 1)
            
            # デフォルトの文字入力をキャンセル
            return "break"
    
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
        日付入力欄から次のフィールドに移動
        イベントが選択されている場合は最初の入力項目、そうでない場合はイベント番号入力欄
        
        Args:
            event: イベントオブジェクト
        """
        if self.selected_event and self.field_widgets:
            # 最初の入力項目に移動
            first_widget = list(self.field_widgets.values())[0]
            first_widget.focus()
        else:
            # イベントが選択されていない場合はイベント番号入力欄に戻る
            self.event_number_entry.focus()
        return "break"
    
    def _move_from_note_field(self, event):
        """
        メモ欄から最初の入力項目に戻る
        
        Args:
            event: イベントオブジェクト
        """
        if self.selected_event and self.field_widgets:
            # 最初の入力項目に移動
            first_widget = list(self.field_widgets.values())[0]
            first_widget.focus()
        return "break"
    
    def _move_to_next_ai_et_field(self, event, next_row: int):
        """
        AI/ETイベントのEnterキーで次のフィールドに移動
        
        Args:
            event: イベントオブジェクト
            next_row: 次の行番号
        """
        # 次の行のウィジェットを探す
        children = self.form_frame.grid_slaves(row=next_row, column=1)
        if children:
            next_widget = children[0]
            if isinstance(next_widget, tk.Widget):
                next_widget.focus()
                return "break"
        # 次の行が見つからない場合はメモ欄に移動
        self.note_entry.focus()
        return "break"
    
    def _on_history_right_click(self, event):
        """イベント履歴で右クリックされた時の処理"""
        # クリックされたアイテムを選択
        item = self.history_tree.identify_row(event.y)
        if not item:
            return
        
        # アイテムを選択
        self.history_tree.selection_set(item)
        
        # tagsを確認（イベントIDが保存されているか）
        tags = self.history_tree.item(item, 'tags')
        if not tags:
            # tagsがない場合はメニューを表示しない
            return
        
        # 'no_event'タグが付いている場合はメニューを表示しない
        if 'no_event' in tags:
            return
        
        # 'event_'で始まるタグがあるか確認
        has_event_tag = any(tag.startswith('event_') for tag in tags)
        if not has_event_tag:
            # イベントIDが保存されていないアイテムはメニューを表示しない
            return
        
        # メニューを表示
        try:
            self.history_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.history_context_menu.grab_release()
    
    def _on_edit_event_from_history(self):
        """イベント履歴から編集を選択した時の処理"""
        selected_items = self.history_tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        tags = self.history_tree.item(item, 'tags')
        if not tags:
            return
        
        # tagsからイベントIDを取得（"event_123"形式）
        event_id = None
        for tag in tags:
            if tag.startswith('event_'):
                event_id_str = tag.replace('event_', '')
                try:
                    event_id = int(event_id_str)
                    break
                except ValueError:
                    continue
        
        if event_id is None:
            return
        
        # 親ウィンドウを取得
        parent = self.window.master if hasattr(self.window, 'master') else None
        if parent is None:
            # masterがない場合は、ウィンドウの親を探す
            try:
                parent_name = self.window.winfo_parent()
                if parent_name:
                    parent = self.window.nametowidget(parent_name)
            except:
                pass
        
        # farm_pathを取得
        farm_path = None
        if hasattr(self, 'settings_manager') and self.settings_manager:
            farm_path = getattr(self.settings_manager, 'farm_path', None)
        
        # 現在のウィンドウを閉じる
        self.window.destroy()
        
        # 編集用のイベント入力ウィンドウを開く
        from ui.event_input import EventInputWindow
        edit_window = EventInputWindow(
            parent=parent,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=self.cow_auto_id,
            event_id=event_id,
            on_saved=self.on_saved,
            farm_path=farm_path
        )
        edit_window.show()
    
    def _on_delete_event_from_history(self):
        """イベント履歴から削除を選択した時の処理"""
        selected_items = self.history_tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        tags = self.history_tree.item(item, 'tags')
        if not tags:
            return
        
        # tagsからイベントIDを取得（"event_123"形式）
        event_id = None
        for tag in tags:
            if tag.startswith('event_'):
                event_id_str = tag.replace('event_', '')
                try:
                    event_id = int(event_id_str)
                    break
                except ValueError:
                    continue
        
        if event_id is None:
            return
        
        # 確認ダイアログ
        result = messagebox.askyesno("確認", "このイベントを削除しますか？")
        if not result:
            return
        
        try:
            # 論理削除
            self.db.delete_event(event_id, soft_delete=True)
            self.rule_engine.on_event_deleted(event_id)
            
            # イベント履歴を更新
            self._load_event_history()
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

