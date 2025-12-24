"""
FALCON2 - CowCard（個体カード）
1頭の牛の情報を表示する（表示専用）
設計書 第11章・第13章参照
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, Callable, List
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.event_display import format_insemination_event, format_calving_event
from ui.event_input import EventInputWindow
from settings_manager import SettingsManager

# EventDetailWindow は後で実装（今はダミー）
try:
    from ui.event_detail_window import EventDetailWindow
except ImportError:
    # EventDetailWindow が存在しない場合はダミークラス
    class EventDetailWindow:
        def __init__(self, *args, **kwargs):
            pass
        def show(self):
            pass


class CowCard:
    """個体カード（表示専用）"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler, 
                 formula_engine: FormulaEngine,
                 rule_engine: Optional[RuleEngine] = None,
                 event_dictionary_path: Optional[Path] = None,
                 item_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス（イベント追加ボタン用）
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス
        """
        try:
            print("DEBUG: CowCard.__init__ 開始")
            
            self.db = db_handler
            self.formula_engine = formula_engine
            self.rule_engine = rule_engine  # None でもOK
            self.event_dict_path = event_dictionary_path
            self.item_dict_path = item_dictionary_path
            self.calculated_items_path = Path("config/cow_card_calculated_items.json")
            self.calculated_item_codes: List[str] = self._load_calculated_item_codes()
            self.cow_auto_id: Optional[int] = None
            self.settings_manager: Optional[SettingsManager] = None  # 後で初期化
            
            # 授精設定辞書（code -> name の形式で保持）
            self.technicians_dict: Dict[str, str] = {}  # {"1": "Sonoda"}
            self.insemination_types_dict: Dict[str, str] = {}  # {"1": "自然発情"}
            # PEN設定（code -> name）
            self.pen_settings: Dict[str, str] = {}
            
            # event_dictionary を読み込む
            self.event_dictionary: Dict[str, Dict[str, Any]] = {}
            self._load_event_dictionary()
            
            # 使用される色タグを追跡（動的にタグを設定するため）
            self._configured_color_tags: set = set()
            
            # 繁殖コード表示名（RuleEngineがNoneでも動作するように）
            self.rc_names = {
                1: "Fresh（分娩後）",
                2: "Bred（授精後）",
                3: "Pregnant（妊娠中）",
                4: "Dry（乾乳中）",
                5: "Open（空胎）",
                6: "Stopped（繁殖停止）"
            }
            
            # UI作成
            self.frame = ttk.Frame(parent)
            self._create_widgets()
            
            # イベント追加ボタン用のコールバック
            self.on_event_saved_callback: Optional[Callable] = None
            
            # 最終日付計算用のコールバック（latest_calving, latest_ai, latest_milk, latest_any を通知）
            self.on_latest_dates_calculated_callback: Optional[Callable[[Optional[str], Optional[str], Optional[str], Optional[str]], None]] = None
            
            print("DEBUG: CowCard.__init__ 完了")
            
        except Exception as e:
            import traceback
            print("ERROR: CowCard.__init__ で例外が発生しました:")
            traceback.print_exc()
            raise
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
    
    @staticmethod
    def get_latest_event_date(events: List[Dict[str, Any]], target_numbers: List[int]) -> Optional[str]:
        """
        指定されたイベント番号の中で最新のevent_dateを取得
        
        Args:
            events: イベントリスト
            target_numbers: 対象となるイベント番号のリスト
        
        Returns:
            最新のevent_date（YYYY-MM-DD形式）、存在しない場合はNone
        """
        dates = [
            e["event_date"]
            for e in events
            if e.get("event_number") in target_numbers and e.get("event_date")
        ]
        return max(dates) if dates else None
    
    @staticmethod
    def get_latest_any_event_date(events: List[Dict[str, Any]]) -> Optional[str]:
        """
        全イベントの中で最新のevent_dateを取得
        
        Args:
            events: イベントリスト
        
        Returns:
            最新のevent_date（YYYY-MM-DD形式）、存在しない場合はNone
        """
        dates = [e["event_date"] for e in events if e.get("event_date")]
        return max(dates) if dates else None
    
    def _get_event_name(self, event_number: int) -> str:
        """
        イベント番号からイベント名を取得
        
        Args:
            event_number: イベント番号
        
        Returns:
            イベント名（日本語）
        """
        event_str = str(event_number)
        if event_str in self.event_dictionary:
            name_jp = self.event_dictionary[event_str].get('name_jp')
            if name_jp:
                return name_jp
        
        # event_dictionary.jsonにない場合のフォールバック
        # RuleEngineの定義に基づくデフォルト名
        default_names = {
            RuleEngine.EVENT_AI: "AI",
            RuleEngine.EVENT_ET: "ET",
            RuleEngine.EVENT_CALV: "分娩",
            RuleEngine.EVENT_DRY: "乾乳",
            RuleEngine.EVENT_STOPR: "繁殖停止",
            RuleEngine.EVENT_SOLD: "売却",
            RuleEngine.EVENT_DEAD: "死亡・淘汰",
            RuleEngine.EVENT_PDN: "妊娠鑑定マイナス",
            RuleEngine.EVENT_PDP: "妊娠鑑定プラス",
            RuleEngine.EVENT_PDP2: "妊娠鑑定プラス（検診以外）",
            RuleEngine.EVENT_ABRT: "流産",
            RuleEngine.EVENT_PAGN: "PAGマイナス",
            RuleEngine.EVENT_PAGP: "PAGプラス",
            RuleEngine.EVENT_MILK_TEST: "乳検",
            RuleEngine.EVENT_MOVE: "群変更"
        }
        
        return default_names.get(event_number, f'イベント{event_number}')
    
    
    def _ensure_color_tag(self, color: str, event_number: int) -> str:
        """
        色に対応するタグを確保（まだ設定されていない場合は設定する）
        
        Args:
            color: 色名（Tkinterが解釈できる色名または#RRGGBB形式）
            event_number: イベント番号（分娩イベントの場合は太字にするため）
        
        Returns:
            タグ名
        """
        # タグ名は色名を使用（#RRGGBB形式の場合はそのまま使用）
        tag_name = color
        
        # まだ設定されていない場合は設定
        if tag_name not in self._configured_color_tags:
            # 分娩イベントの場合は太字も設定
            if event_number == RuleEngine.EVENT_CALV:
                self.event_tree.tag_configure(tag_name, foreground=color, font=('', 9, 'bold'))
            else:
                self.event_tree.tag_configure(tag_name, foreground=color)
            self._configured_color_tags.add(tag_name)
        
        return tag_name
    
    def _get_event_display_color(self, event_number: int) -> str:
        """
        イベントの表示色を決定
        
        【優先順位】
        1. event_dictionary.json の display_color が指定されていればそれを使用
        2. category / outcome に応じてデフォルト色を使用
           - CALVING → #0066cc（青）
           - PREGNANCY + outcome=NEGATIVE → #cc0000（赤）
           - PREGNANCY + outcome=POSITIVE（または未指定） → #008000（緑）
           - BREEDING → #000000（黒）
        3. それ以外は #000000（黒）
        
        Args:
            event_number: イベント番号
        
        Returns:
            色（#RRGGBB形式）
        """
        # イベント辞書から情報を取得
        event_str = str(event_number)
        event_dict = self.event_dictionary.get(event_str, {})
        
        # 1. display_color が指定されていればそれを使用（最優先）
        display_color = event_dict.get('display_color')
        if display_color:
            return display_color
        
        # 2. category / outcome に応じてデフォルト色を使用
        category = event_dict.get('category', '')
        
        if category == "CALVING":
            return "#0066cc"  # 青
        elif category == "PREGNANCY":
            outcome = event_dict.get('outcome', '')
            if outcome == "NEGATIVE":
                return "#cc0000"  # 赤
            else:
                return "#008000"  # 緑（POSITIVE または未指定）
        elif category == "BREEDING":
            return "#000000"  # 黒
        
        # 3. それ以外は黒
        return "#000000"
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ========== イベント追加ボタン ==========
        if self.rule_engine:
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.add_event_button = ttk.Button(
                button_frame,
                text="イベント追加",
                command=self._on_add_event,
                width=15
            )
            self.add_event_button.pack(side=tk.RIGHT, padx=5)
        
        # ========== 上：基本情報 ==========
        basic_frame = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        basic_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.basic_labels = {}
        basic_fields = [
            ('cow_id', '拡大4桁ID'),
            ('jpn10', '個体識別番号'),
            ('brd', '品種'),
            ('bthd', '生年月日'),
            ('lact', '産次'),
            ('rc', '繁殖コード'),
            ('pen', '群')
        ]
        
        for i, (key, label) in enumerate(basic_fields):
            row = i // 2
            col = (i % 2) * 2
            
            ttk.Label(basic_frame, text=f"{label}:").grid(
                row=row, column=col, sticky=tk.W, padx=5, pady=2
            )
            value_label = ttk.Label(basic_frame, text="", foreground="blue")
            value_label.grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)
            self.basic_labels[key] = value_label
        
        # ========== 中：計算項目 ==========
        calc_frame = ttk.LabelFrame(main_frame, text="計算項目", padding=10)
        calc_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.calc_items_frame = ttk.Frame(calc_frame)
        self.calc_items_frame.pack(fill=tk.X, expand=True)
        self.calc_item_widgets: List[Dict[str, Any]] = []
        self.refresh_calculated_items({})
        
        # ========== 下：イベント履歴 ==========
        event_frame = ttk.LabelFrame(main_frame, text="イベント履歴", padding=10)
        event_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview
        columns = ('date', 'event', 'note')
        self.event_tree = ttk.Treeview(event_frame, columns=columns, show='headings', height=10)
        
        # カラム設定
        self.event_tree.heading('date', text='日付')
        self.event_tree.heading('event', text='イベント')
        self.event_tree.heading('note', text='NOTE')
        
        self.event_tree.column('date', width=100)
        self.event_tree.column('event', width=150)
        self.event_tree.column('note', width=300)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(event_frame, orient=tk.VERTICAL, command=self.event_tree.yview)
        self.event_tree.configure(yscrollcommand=scrollbar.set)
        
        # 配置
        self.event_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ダブルクリックイベント
        self.event_tree.bind('<Double-Button-1>', self._on_event_double_click)
        
        # 右クリックイベント
        self.event_tree.bind('<Button-3>', self._on_event_right_click)
    
    def load_cow(self, cow_auto_id: int):
        """
        牛の情報を読み込んで表示
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            self.cow_auto_id = cow_auto_id
            
            # cow データを取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                print(f"WARNING: cow が見つかりません: auto_id={cow_auto_id}")
                return
            
            # 基本情報を表示
            self._display_basic_info(cow)
            
            # 計算項目を表示（FormulaEngine を1回だけ呼ぶ）
            try:
                calculated = self.formula_engine.calculate(cow_auto_id)
                self.refresh_calculated_items(calculated)
            except Exception as e:
                print(f"WARNING: FormulaEngine.calculate でエラー: {e}")
                import traceback
                traceback.print_exc()
            
            # イベント履歴を表示
            self._display_events(cow_auto_id)
            
            # 最終日付を計算してコールバックで通知
            self._calculate_and_notify_latest_dates(cow_auto_id)
            
        except Exception as e:
            import traceback
            print(f"ERROR: load_cow で例外が発生しました: {e}")
            traceback.print_exc()
            raise
    
    def _display_basic_info(self, cow: Dict[str, Any]):
        """基本情報を表示"""
        # cow_id
        self.basic_labels['cow_id'].config(text=cow.get('cow_id', ''))
        
        # jpn10（個体識別番号）- 表示のみ
        jpn10 = cow.get('jpn10', '')
        self.basic_labels['jpn10'].config(text=jpn10)
        
        # brd（品種）
        self.basic_labels['brd'].config(text=cow.get('brd', ''))
        
        # bthd（生年月日）
        self.basic_labels['bthd'].config(text=cow.get('bthd', ''))
        
        # lact（産次）
        lact = cow.get('lact')
        self.basic_labels['lact'].config(text=str(lact) if lact is not None else '')
        
        # rc（繁殖コード）
        rc = cow.get('rc')
        rc_text = self.rc_names.get(rc, f'RC{rc}') if rc is not None else ''
        self.basic_labels['rc'].config(text=rc_text)
        
        # pen（群）
        self.basic_labels['pen'].config(text=cow.get('pen', ''))
    
    def refresh_calculated_items(self, calculated: Optional[Dict[str, Any]] = None):
        """
        計算項目表示を再描画（ウィンドウ再生成なし）
        """
        try:
            if calculated is None:
                if self.cow_auto_id is None:
                    calculated = {}
                else:
                    calculated = self.formula_engine.calculate(self.cow_auto_id)
            self._render_calculated_items(calculated or {})
        except Exception as e:
            import traceback
            logging.error(f"ERROR: refresh_calculated_items で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _render_calculated_items(self, calculated: Dict[str, Any]):
        """計算項目エリアのUIを再構築"""
        # 既存ウィジェットをクリア
        for row in self.calc_item_widgets:
            try:
                row_frame = row.get("frame")
                if row_frame is not None:
                    row_frame.destroy()
            except tk.TclError:
                pass
        self.calc_item_widgets = []
        
        # 表示用リスト（末尾に空行を追加）
        display_codes = list(self.calculated_item_codes)
        display_codes.append(None)  # 空行
        
        for idx, item_code in enumerate(display_codes):
            row_frame = ttk.Frame(self.calc_items_frame)
            row_frame.grid(row=idx, column=0, sticky=tk.W, pady=2)
            
            name = self._get_item_display_name(item_code)
            name_label = ttk.Label(row_frame, text=f"{name}:")
            name_label.grid(row=0, column=0, sticky=tk.W, padx=5)
            
            # 計算値を取得（None の場合は None、それ以外はそのまま）
            raw_value = calculated.get(item_code) if item_code else None
            # デバッグログ
            if item_code:
                logging.info(
                    f"[CowCard] calc item={item_code} value={raw_value} type={type(raw_value)}"
                )
            
            value_text = self._format_calc_value(item_code, raw_value)
            value_label = ttk.Label(row_frame, text=value_text, foreground="green")
            value_label.grid(row=0, column=1, sticky=tk.W, padx=5)
            
            # editableな項目の場合はダブルクリックで編集可能
            if item_code:
                item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
                item_def = item_dict.get(item_code, {})
                if item_def.get("editable", False):
                    # ダブルクリックで編集
                    for widget in (row_frame, name_label, value_label):
                        widget.bind(
                            '<Double-Button-1>',
                            lambda e, code=item_code: self._on_edit_calc_item(code)
                        )
                        # 編集可能であることを示すカーソル
                        widget.bind('<Enter>', lambda e, w=widget: w.config(cursor='hand2'))
                        widget.bind('<Leave>', lambda e, w=widget: w.config(cursor=''))
            
            # 右クリックメニュー
            for widget in (row_frame, name_label, value_label):
                widget.bind(
                    '<Button-3>',
                    lambda e, idx=idx, code=item_code: self._on_calc_item_right_click(e, idx, code)
                )
            
            self.calc_item_widgets.append(
                {
                    "frame": row_frame,
                    "name_label": name_label,
                    "value_label": value_label,
                    "item_code": item_code,
                }
            )
    
    def _on_calc_item_right_click(self, event, index: int, item_code: Optional[str]):
        """計算項目行の右クリックメニュー"""
        try:
            menu = tk.Menu(self.frame, tearoff=0)
            if item_code:
                menu.add_command(label="項目を変更", command=lambda: self._replace_calc_item(index))
                menu.add_command(label="項目を削除", command=lambda: self._remove_calc_item(index, item_code))
            else:
                menu.add_command(label="項目を追加", command=lambda: self._add_calc_item(index))
            
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_calc_item_right_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _add_calc_item(self, index: int):
        """空行に項目を追加"""
        selected = self._select_item_from_dictionary()
        if not selected:
            return
        if selected in self.calculated_item_codes:
            messagebox.showinfo("情報", "同じ項目がすでに追加されています。")
            return
        self.calculated_item_codes.insert(index, selected)
        self._save_calculated_item_codes()
        logging.info(f"CowCard calculated item added: {selected}")
        self.refresh_calculated_items()
    
    def _replace_calc_item(self, index: int):
        """既存行の項目を差し替え"""
        if index < 0 or index >= len(self.calculated_item_codes):
            return
        current = self.calculated_item_codes[index]
        selected = self._select_item_from_dictionary(initial_code=current)
        if not selected or selected == current:
            return
        if selected in self.calculated_item_codes:
            messagebox.showinfo("情報", "同じ項目がすでに追加されています。")
            return
        self.calculated_item_codes[index] = selected
        self._save_calculated_item_codes()
        logging.info(f"CowCard calculated item replaced: {current} -> {selected}")
        self.refresh_calculated_items()
    
    def _remove_calc_item(self, index: int, item_code: str):
        """項目を削除"""
        if index < 0 or index >= len(self.calculated_item_codes):
            return
        if self.calculated_item_codes[index] != item_code:
            return
        del self.calculated_item_codes[index]
        self._save_calculated_item_codes()
        logging.info(f"CowCard calculated item removed: {item_code}")
        self.refresh_calculated_items()
    
    def _on_edit_calc_item(self, item_code: str):
        """editableな計算項目を編集"""
        if not self.cow_auto_id:
            messagebox.showwarning("警告", "対象牛が設定されていません")
            return
        
        # 現在の値を取得
        calculated = self.formula_engine.calculate(self.cow_auto_id)
        current_value = calculated.get(item_code)
        
        # 項目定義を取得
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        item_def = item_dict.get(item_code, {})
        display_name = item_def.get("display_name") or item_def.get("label") or item_code
        data_type = item_def.get("data_type", "str")
        
        # 編集ダイアログを表示
        dialog = tk.Toplevel(self.frame)
        dialog.title(f"{display_name}を編集")
        dialog.geometry("300x150")
        dialog.transient(self.frame.winfo_toplevel())
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"{display_name}:").pack(pady=10)
        
        # 入力フィールド
        entry = ttk.Entry(dialog, width=20)
        entry.pack(pady=5)
        
        # 現在の値を設定
        if current_value is not None:
            entry.insert(0, str(current_value))
        entry.select_range(0, tk.END)
        entry.focus()
        
        def on_ok():
            try:
                value_str = entry.get().strip()
                
                # データ型に応じて変換
                if data_type == "int":
                    value = int(value_str) if value_str else None
                elif data_type == "float":
                    value = float(value_str) if value_str else None
                else:
                    value = value_str if value_str else None
                
                # item_valueテーブルに保存
                if value is not None:
                    self.db.set_item_value(self.cow_auto_id, item_code, value)
                    logging.info(f"CowCard editable calc item updated: {item_code}={value}")
                else:
                    # Noneの場合は削除（item_valueテーブルから削除する必要があるが、現在のDBHandlerには削除メソッドがない）
                    # とりあえず空文字列を保存
                    self.db.set_item_value(self.cow_auto_id, item_code, "")
                    logging.info(f"CowCard editable calc item cleared: {item_code}")
                
                # 表示を更新
                self.refresh_calculated_items()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("エラー", f"無効な値です。{data_type}型の値を入力してください。")
        
        def on_cancel():
            dialog.destroy()
        
        # EnterキーでOK
        entry.bind('<Return>', lambda e: on_ok())
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
    
    def _select_item_from_dictionary(self, initial_code: Optional[str] = None) -> Optional[str]:
        """Item Dictionary から計算項目を選択"""
        selectable = self._get_selectable_items()
        if not selectable:
            messagebox.showerror("エラー", "選択可能な計算項目がありません。item_dictionary.json を確認してください。")
            return None
        
        dialog = tk.Toplevel(self.frame)
        dialog.title("計算項目を選択")
        dialog.geometry("400x400")
        dialog.transient(self.frame.winfo_toplevel())
        dialog.grab_set()
        
        ttk.Label(dialog, text="計算項目を選択してください").pack(pady=10)
        
        listbox = tk.Listbox(dialog, height=15)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10)
        
        display_items = []
        for code, name in selectable:
            display_items.append(f"{name} ({code})")
            listbox.insert(tk.END, display_items[-1])
            if initial_code and code == initial_code:
                listbox.select_set(tk.END)
        
        selected_code: Optional[str] = None
        
        def on_ok():
            nonlocal selected_code
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("警告", "項目を選択してください。")
                return
            idx = sel[0]
            selected_code = selectable[idx][0]
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        listbox.bind('<Double-Button-1>', lambda e: on_ok())
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return selected_code
    
    def _get_selectable_items(self) -> List[tuple]:
        """選択可能な計算項目の一覧を取得"""
        items = []
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        for code, data in item_dict.items():
            item_type = data.get("type")
            if item_type not in ("calculated", "calc"):
                continue
            if data.get("visible", True) is False:
                continue
            if not data.get("formula") and not data.get("source"):
                continue
            # label, display_name, name_jp の順で取得
            display_name = data.get("label") or data.get("display_name") or data.get("name_jp") or code
            items.append((code, display_name))
        items.sort(key=lambda x: x[1])
        return items
    
    def _get_item_display_name(self, item_code: Optional[str]) -> str:
        """表示名を取得（item_dictionaryに未定義なら item_code を使用）"""
        if not item_code:
            return "（空行）"
        item_dict = getattr(self.formula_engine, "item_dictionary", {}) or {}
        if item_code in item_dict:
            data = item_dict[item_code]
            # label, display_name, name_jp の順で取得
            return data.get("label") or data.get("display_name") or data.get("name_jp") or item_code
        # 既存の固定表示名への後方互換
        fallback_names = {
            'DIM': '分娩後日数',
            'DPAI': '最終AI後日数',
            'DUE': '分娩予定日',
            'LAST_MILK_YIELD': '最新乳量',
            'LAST_FAT': '最新乳脂肪',
            'LAST_PROTEIN': '最新乳タンパク',
            'LAST_SNF': '最新無脂固形分',
            'LAST_SCC': '最新体細胞数',
            'LAST_LS': '最新LS',
            'LAST_MUN': '最新MUN',
            'LAST_BHB': '最新BHB',
            'LAST_DENOVO_FA': '最新DeNovo FA',
            'LAST_PREFORMED_FA': '最新Preformed FA',
            'LAST_MIXED_FA': '最新Mixed FA',
            'LAST_DENOVO_MILK': '最新DeNovo Milk'
        }
        return fallback_names.get(item_code, item_code)
    
    def _format_calc_value(self, item_code: Optional[str], value: Any) -> str:
        """
        値の表示フォーマット
        
        Args:
            item_code: 項目コード
            value: 値（None の場合は空文字列を返す、0 / 0.0 は有効な値として表示）
        
        Returns:
            フォーマットされた文字列
        """
        # None のみを空表示にする（0 / 0.0 は有効な値として表示）
        if item_code is None or value is None:
            return ""
        
        # 数値型の場合はそのまま表示（0 / 0.0 も含む）
        if item_code in ['DIM', 'DPAI']:
            return f"{value}日"
        # 削除された項目（LAST_MILK_YIELD, LAST_FAT, LAST_PROTEIN, LAST_SNF, LAST_SCC等）は
        # item_dictionaryから削除されているため、通常はここに来ないが、
        # 念のため存在チェックを追加（削除された項目はそのまま文字列として表示）
        # その他の項目も 0 / 0.0 を含めて表示
        return str(value)
    
    def _load_calculated_item_codes(self) -> List[str]:
        """表示定義リストをロード（無ければデフォルト生成）"""
        # 削除された項目（LAST_MILK_YIELD, LAST_FAT, LAST_PROTEIN, LAST_SCC等）は
        # item_dictionaryから削除されているため、デフォルトから除外
        default_codes = [
            "DIM",
            "DPAI",
            "DUE",
        ]
        path = self.calculated_items_path
        try:
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list) and all(isinstance(x, str) for x in data):
                    return data
            # 無い場合はデフォルトを保存して返す
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(default_codes, f, ensure_ascii=False, indent=2)
            return default_codes
        except Exception as e:
            logging.error(f"calculated_items 読み込みエラー: {e}")
            return default_codes
    
    def _save_calculated_item_codes(self):
        """表示定義リストを保存"""
        try:
            self.calculated_items_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.calculated_items_path, 'w', encoding='utf-8') as f:
                json.dump(self.calculated_item_codes, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"calculated_items 保存エラー: {e}")
    
    def _display_events(self, cow_auto_id: int):
        """イベント履歴を表示（event_date DESC順）"""
        try:
            # SettingsManagerを初期化（まだ初期化されていない場合）
            if self.settings_manager is None:
                self._init_settings_manager(cow_auto_id)
            
            # 既存のアイテムをクリア
            for item in self.event_tree.get_children():
                self.event_tree.delete(item)
            
            # イベントを取得（既にevent_date DESC順でソート済み）
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            logging.debug(f"[CowCard] total events fetched = {len(events)}")
            
            # 表示されたイベント数をカウント
            displayed_count = 0
            
            # Treeviewに追加（すべてのイベントを表示）
            for event in events:
                try:
                    event_date = event.get('event_date', '')
                    event_number = event.get('event_number')
                    
                    # event_numberがNoneの場合はスキップ
                    if event_number is None:
                        logging.warning(f"[CowCard] Skipping event with None event_number: event_id={event.get('id')}")
                        continue
                    
                    event_name = self._get_event_name(event_number)
                    # json_dataはDBから取得したものを一切加工せず、そのまま保持する
                    json_data = event.get('json_data')
                    
                    # json_dataがNoneの場合は空辞書に変換（表示処理用）
                    if json_data is None:
                        json_data = {}
                    
                    # noteを初期化（AI/ETイベント以外は既存のnoteを使用）
                    note = ''
                    
                    # AI/ETイベント（200, 201）の場合は詳細情報を備考欄に追加
                    # 【必須ルール】event["memo"]やevent["note"]は一切使わず、必ずjson_dataから再生成する
                    if event_number in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
                        # 共通関数を使用して表示文字列を生成（毎回json_dataから再生成）
                        # technicians_dict / insemination_types_dict は必ず渡す（空辞書でも可）
                        note = format_insemination_event(
                            json_data,
                            self.technicians_dict,          # code -> name
                            self.insemination_types_dict    # code -> name
                        ) or ''  # Noneの場合は空文字列
                    
                    # 分娩イベントの表示
                    elif event_number == RuleEngine.EVENT_CALV:
                        calv_def = self.event_dictionary.get(str(RuleEngine.EVENT_CALV), {})
                        diff_labels = calv_def.get("calving_difficulty", {})
                        note = format_calving_event(json_data, diff_labels) or ''
                    
                    # BREEDINGカテゴリイベントの場合は詳細情報を備考欄に追加
                    # イベント番号で判定（300, 301, 302）またはcategoryで判定
                    event_dict = self.event_dictionary.get(str(event_number), {})
                    is_breeding = (event_number in [300, 301, RuleEngine.EVENT_PDN] or 
                                   event_dict.get('category') == 'BREEDING')
                    
                    if is_breeding:
                        # 有効な値のみをリスト化（空文字、None、"-"は除外）
                        valid_parts = []
                        
                        # 処置（新しいキー名と古いキー名の両方をサポート）
                        treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
                        if treatment and str(treatment).strip() and str(treatment).strip() != '-':
                            valid_parts.append(str(treatment).strip())
                        
                        # 子宮所見（新しいキー名と古いキー名の両方をサポート）
                        uterine = (json_data.get('uterine_findings') or 
                                  json_data.get('uterus_findings') or 
                                  json_data.get('uterus_finding') or 
                                  json_data.get('uterus', ''))
                        if uterine and str(uterine).strip() and str(uterine).strip() != '-':
                            valid_parts.append(f"子宮{str(uterine).strip()}")
                        
                        # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
                        left_ovary = (json_data.get('left_ovary_findings') or 
                                     json_data.get('leftovary_findings') or 
                                     json_data.get('leftovary_finding') or 
                                     json_data.get('left_ovary', '') or
                                     json_data.get('leftovary', ''))
                        if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
                            valid_parts.append(f"左{str(left_ovary).strip()}")
                        
                        # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
                        right_ovary = (json_data.get('right_ovary_findings') or 
                                      json_data.get('rightovary_findings') or 
                                      json_data.get('rightovary_finding') or 
                                      json_data.get('right_ovary', '') or
                                      json_data.get('rightovary', ''))
                        if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
                            valid_parts.append(f"右{str(right_ovary).strip()}")
                        
                        # remark（新しいキー名と古いキー名の両方をサポート）
                        # まず remark を取得、なければ other 系を確認
                        remark = json_data.get('remark')
                        if not remark:
                            # remark が無い場合は other 系を確認
                            remark = (json_data.get('other') or 
                                     json_data.get('other_info') or 
                                     json_data.get('other_findings', ''))
                        if remark and str(remark).strip() and str(remark).strip() != '-':
                            valid_parts.append(str(remark).strip())
                        
                        # 有効な項目のみを「  」で区切って表示
                        if valid_parts:
                            detail_text = "  ".join(valid_parts)
                            # 既存の備考がある場合は結合
                            if note:
                                note = f"{detail_text} | {note}"
                            else:
                                note = detail_text
                
                    # 乳検イベント（601）の場合は詳細情報を備考欄に追加
                    elif event_number == RuleEngine.EVENT_MILK_TEST:
                        detail_parts = []
                        # 指定された順序で表示
                        if json_data.get('milk_yield') is not None:
                            detail_parts.append(f"乳量{json_data['milk_yield']}kg")
                        if json_data.get('fat') is not None:
                            detail_parts.append(f"乳脂率{json_data['fat']}%")
                        if json_data.get('snf') is not None:
                            detail_parts.append(f"無脂固形分{json_data['snf']}%")
                        if json_data.get('protein') is not None:
                            detail_parts.append(f"蛋白率{json_data['protein']}%")
                        if json_data.get('scc') is not None:
                            detail_parts.append(f"体細胞{json_data['scc']:,}")
                        if json_data.get('mun') is not None:
                            detail_parts.append(f"MUN{json_data['mun']}")
                        if json_data.get('ls') is not None:
                            detail_parts.append(f"体細胞スコア{json_data['ls']}")
                        if json_data.get('bhb') is not None:
                            detail_parts.append(f"BHB{json_data['bhb']}")
                        if json_data.get('denovo_fa') is not None:
                            detail_parts.append(f"デノボFA{json_data['denovo_fa']}")
                        if json_data.get('preformed_fa') is not None:
                            detail_parts.append(f"プレフォームFA{json_data['preformed_fa']}")
                        if json_data.get('mixed_fa') is not None:
                            detail_parts.append(f"ミックスFA{json_data['mixed_fa']}")
                        if json_data.get('denovo_milk') is not None:
                            detail_parts.append(f"デノボMilk{json_data['denovo_milk']}")
                        
                        if detail_parts:
                            detail_text = "  ".join(detail_parts)
                            # 既存の備考がある場合は結合
                            if note:
                                note = f"{note} | {detail_text}"
                            else:
                                note = detail_text
                    else:
                        # AI/ET/繁殖検査/乳検以外のイベントは既存のnoteを使用
                        note = event.get('note', '') or ''
                
                    # イベントの表示色を決定
                    display_color = self._get_event_display_color(event_number)
                    
                    # 色タグを動的に設定（まだ設定されていない場合）
                    color_tag = self._ensure_color_tag(display_color, event_number)
                    
                    # イベントIDを取得
                    event_id = event.get('id')
                    if event_id is None:
                        logging.warning(f"[CowCard] Skipping event with None event_id: event_number={event_number}")
                        continue
                    
                    # 安全確認用ログ（Treeviewにセットする直前）
                    logging.info(
                        "CowCard display event_id=%s number=%s json=%s display=%s",
                        event_id,
                        event_number,
                        json.dumps(json_data, ensure_ascii=False),
                        note
                    )
                    
                    # Treeviewに追加（iid に event_id を紐づける）
                    # DB由来イベントは iid = str(event_id) とする
                    self.event_tree.insert(
                        '',
                        'end',
                        iid=str(event_id),
                        values=(event_date, event_name, note),
                        tags=(color_tag,)
                    )
                    displayed_count += 1
                        
                except Exception as e:
                    # 個別のイベント処理でエラーが発生しても、他のイベントは表示を続ける
                    logging.error(f"[CowCard] Error processing event: {e}, event_id={event.get('id')}")
                    import traceback
                    traceback.print_exc()
                    continue
            
            logging.debug(f"[CowCard] displayed events = {displayed_count}")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _display_events で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _calculate_and_notify_latest_dates(self, cow_auto_id: int):
        """
        最終日付を計算してコールバックで通知
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            # イベントを取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            # 最終分娩日を計算（EVENT_CALV = 202、baselineも含む）
            latest_calving = self.get_latest_event_date(events, [RuleEngine.EVENT_CALV])
            
            # 最終AI日を計算（AI/ET系イベント: 200, 201, 203, 204など）
            # ユーザー要求では「200, 201, 203, 204 など」とあるが、一般的には200(AI), 201(ET)が対象
            # 203(DRY), 204(STOPR)は繁殖関連だがAI/ETではないため、200, 201のみを対象とする
            latest_ai = self.get_latest_event_date(events, [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET])
            
            # 最終乳検日を計算（EVENT_MILK_TEST = 601）
            latest_milk = self.get_latest_event_date(events, [RuleEngine.EVENT_MILK_TEST])
            
            # 最終イベント日を計算（全イベント対象）
            latest_any = self.get_latest_any_event_date(events)
            
            # コールバックがあれば呼び出す
            if self.on_latest_dates_calculated_callback:
                self.on_latest_dates_calculated_callback(latest_calving, latest_ai, latest_milk, latest_any)
                
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _calculate_and_notify_latest_dates で例外が発生しました: {e}")
            traceback.print_exc()
            # エラー時もコールバックを呼ぶ（Noneを渡す）
            if self.on_latest_dates_calculated_callback:
                self.on_latest_dates_calculated_callback(None, None, None, None)
    
    def _init_settings_manager(self, cow_auto_id: int):
        """
        SettingsManagerを初期化し、授精設定辞書を読み込む
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                frm = cow.get('frm')
                if frm:
                    farm_path = Path(f"C:/FARMS/{frm}")
                    self.settings_manager = SettingsManager(farm_path)
                    # insemination_settings.json から直接読み込む（EventInputWindowと同じ方法）
                    self._load_insemination_settings(farm_path)
                    # farm_settings.json から PEN 設定をロード
                    try:
                        self.pen_settings = self.settings_manager.load_pen_settings()
                    except Exception as e:
                        logging.error(f"PEN設定の読み込みに失敗しました: {e}")
                        self.pen_settings = {}
                else:
                    # frm が無い場合は空辞書を設定
                    self.technicians_dict = {}
                    self.insemination_types_dict = {}
                    self.pen_settings = {}
            else:
                # cow が見つからない場合は空辞書を設定
                self.technicians_dict = {}
                self.insemination_types_dict = {}
                self.pen_settings = {}
        except Exception as e:
            print(f"WARNING: SettingsManagerの初期化に失敗しました: {e}")
            self.settings_manager = None
            # エラー時も空辞書を設定（クラッシュを防ぐ）
            self.technicians_dict = {}
            self.insemination_types_dict = {}
    
    def _load_insemination_settings(self, farm_path: Path):
        """
        insemination_settings.json をロード（EventInputWindowと同じ方法）
        
        Args:
            farm_path: 農場フォルダパス
        """
        settings_file = farm_path / "insemination_settings.json"
        
        if not settings_file.exists():
            logging.warning(f"insemination_settings.json が見つかりません: {settings_file}")
            self.technicians_dict = {}
            self.insemination_types_dict = {}
            return
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            self.technicians_dict = settings.get('technicians', {})
            self.insemination_types_dict = settings.get('insemination_types', {})
            
            logging.debug(f"授精設定をロード: technicians={len(self.technicians_dict)}, insemination_types={len(self.insemination_types_dict)}")
            logging.debug(f"technicians_dict sample: {list(self.technicians_dict.items())[:3] if self.technicians_dict else 'empty'}")
            logging.debug(f"insemination_types_dict sample: {list(self.insemination_types_dict.items())[:3] if self.insemination_types_dict else 'empty'}")
        except Exception as e:
            logging.error(f"insemination_settings.json 読み込みエラー: {e}")
            self.technicians_dict = {}
            self.insemination_types_dict = {}
    
    def _on_event_double_click(self, event):
        """イベントのダブルクリック処理（詳細表示）"""
        try:
            selection = self.event_tree.selection()
            if selection:
                row_id = selection[0]
                # iid から直接 event_id を取得
                # DBイベント以外（baseline 等）は数値でない iid を使うため、isdigit() でチェック
                if not row_id.isdigit():
                    # ダブルクリック不可イベント（baseline 等）の場合は何もしない
                    return
                
                event_id = int(row_id)
                # イベント詳細ウィンドウを表示
                detail_window = EventDetailWindow(
                    self.frame,
                    self.db,
                    event_id,
                    event_dictionary_path=self.event_dict_path
                )
                detail_window.show()
        except Exception as e:
            import traceback
            print(f"ERROR: _on_event_double_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_event_right_click(self, event):
        """イベントの右クリック処理（コンテキストメニュー表示）"""
        try:
            # 右クリック位置の行を取得
            row_id = self.event_tree.identify_row(event.y)
            if not row_id:
                return
            
            # 行を選択・フォーカス
            self.event_tree.selection_set(row_id)
            self.event_tree.focus(row_id)
            
            # iid から直接 event_id を取得
            # DBイベント以外（baseline 等）は数値でない iid を使うため、isdigit() でチェック
            if not row_id.isdigit():
                # 削除不可イベント（baseline 等）の場合はメニューを表示しない
                return
            
            event_id = int(row_id)
            logging.debug(f"Right-click event_id={event_id}")
            
            # イベントデータを取得して削除可否をチェック
            event_data = self.db.get_event_by_id(event_id)
            if event_data is None:
                # イベントが見つからない場合はメニューを表示しない
                logging.warning(f"右クリック対象イベントが取得できません: event_id={event_id}")
                return
            
            # json_data を安全に取得（None / str / dict のいずれでも安全に処理）
            json_data_raw = event_data.get('json_data')
            # json_data が dict でない場合は空dictとして扱う
            if not isinstance(json_data_raw, dict):
                json_data = {}
            else:
                json_data = json_data_raw
            
            event_number = event_data.get('event_number')
            
            # baseline（CSVから作成された分娩イベント）のチェック
            is_baseline = (event_number == RuleEngine.EVENT_CALV and 
                          json_data.get('baseline_calving', False))
            
            # system_generated かつ deprecated のイベントのチェック（dict の場合のみ）
            is_system_deprecated = False
            if isinstance(json_data, dict) and json_data.get('system_generated', False):
                event_str = str(event_number)
                if event_str in self.event_dictionary:
                    if self.event_dictionary[event_str].get('deprecated', False):
                        is_system_deprecated = True
            
            # 削除不可イベントの場合はメニューを表示しない
            if is_baseline or is_system_deprecated:
                return
            
            # コンテキストメニューを作成
            menu = tk.Menu(self.frame, tearoff=0)
            menu.add_command(label="編集", command=lambda: self._on_edit_event(event_id))
            menu.add_command(label="削除", command=lambda: self._on_delete_event(event_id))
            
            # メニューを表示
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
                
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_event_right_click で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_edit_event(self, event_id: int):
        """イベント編集処理"""
        try:
            if not self.cow_auto_id:
                messagebox.showerror("エラー", "対象牛が設定されていません")
                return
            
            if not self.rule_engine:
                messagebox.showerror("エラー", "RuleEngineが設定されていません")
                return
            
            # イベントデータを取得
            event = self.db.get_event_by_id(event_id)
            if not event:
                messagebox.showerror("エラー", "イベントが見つかりません")
                return
            
            logging.debug(f"Editing event_id={event_id}")
            
            # 農場パスを取得
            farm_path = None
            if self.cow_auto_id:
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    frm = cow.get('frm')
                    if frm:
                        farm_path = Path(f"C:/FARMS/{frm}")
            
            # EventInputWindowを編集モードで開く
            event_window = EventInputWindow(
                parent=self.frame.winfo_toplevel(),
                db_handler=self.db,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path,
                cow_auto_id=self.cow_auto_id,
                on_saved=self._on_event_saved,
                farm_path=farm_path,
                edit_event_id=event_id  # 編集モード
            )
            event_window.show()
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_edit_event で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"イベントの編集に失敗しました: {e}")
    
    def _on_delete_event(self, event_id: int):
        """イベント削除処理"""
        try:
            # イベントデータを取得
            event = self.db.get_event_by_id(event_id)
            if event is None:
                messagebox.showwarning(
                    "削除不可",
                    "イベント情報が取得できません"
                )
                return
            
            # 削除制限チェック
            # json_data を安全に取得（None / str / dict のいずれでも安全に処理）
            json_data_raw = event.get('json_data')
            # json_data が dict でない場合は空dictとして扱う
            if not isinstance(json_data_raw, dict):
                json_data = {}
            else:
                json_data = json_data_raw
            
            event_number = event.get('event_number')
            
            # baseline（CSVから作成された分娩イベント）のチェック
            if event_number == RuleEngine.EVENT_CALV and json_data.get('baseline_calving', False):
                messagebox.showerror("エラー", "このイベントは削除できません（baselineイベント）")
                return
            
            # deprecated=true かつ system_generated=true のイベントのチェック
            # 注：event_dictionary.jsonのdeprecatedフラグとjson_dataのsystem_generatedフラグを確認
            # system_generated 判定は dict の場合のみ行う
            if isinstance(json_data, dict) and json_data.get('system_generated', False):
                # event_dictionary.jsonでdeprecatedかどうか確認
                event_str = str(event_number)
                if event_str in self.event_dictionary:
                    if self.event_dictionary[event_str].get('deprecated', False):
                        messagebox.showerror("エラー", "このイベントは削除できません（システム生成イベント）")
                        return
            
            # 確認ダイアログ
            result = messagebox.askyesno(
                "確認",
                "このイベントを削除しますか？\n（この操作は元に戻せません）"
            )
            
            if not result:
                return
            
            logging.info(f"Deleting event_id={event_id}")
            
            # イベントを削除（物理削除）
            self.db.delete_event(event_id, soft_delete=False)
            
            # RuleEngineを再実行
            if self.rule_engine:
                self.rule_engine.on_event_deleted(event_id)
            
            # 表示を更新（load_cow内で日付も再計算される）
            self.refresh()
            
            # 外部コールバックがあれば呼ぶ
            if self.on_event_saved_callback:
                self.on_event_saved_callback(self.cow_auto_id)
            
            # 削除後のイベント件数をログ出力
            events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
            logging.info(f"Event deleted: event_id={event_id}, remaining events: {len(events)}")
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            import traceback
            logging.error(f"ERROR: _on_delete_event で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def refresh(self):
        """表示を更新"""
        if self.cow_auto_id:
            self.load_cow(self.cow_auto_id)
    
    def _on_add_event(self):
        """イベント追加ボタンクリック時の処理"""
        if not self.cow_auto_id:
            print("WARNING: cow_auto_id が設定されていません")
            return
        
        if not self.rule_engine:
            print("WARNING: rule_engine が設定されていません")
            return
        
        try:
            # EventInputWindow を開く（cow_auto_id を指定）
            # 農場パスを取得（SettingsManager用）
            farm_path = None
            if self.cow_auto_id:
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    frm = cow.get('frm')
                    if frm:
                        farm_path = Path(f"C:/FARMS/{frm}")
            
            event_window = EventInputWindow(
                parent=self.frame.winfo_toplevel(),
                db_handler=self.db,
                rule_engine=self.rule_engine,
                event_dictionary_path=self.event_dict_path,
                cow_auto_id=self.cow_auto_id,
                on_saved=self._on_event_saved,
                farm_path=farm_path
            )
            event_window.show()
        except Exception as e:
            import traceback
            print(f"ERROR: _on_add_event で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_event_saved(self, cow_auto_id: int):
        """
        イベント保存後のコールバック
        
        Args:
            cow_auto_id: イベントが保存された牛の auto_id
        """
        try:
            # 個体カードを更新
            if self.cow_auto_id == cow_auto_id:
                self.refresh()
            
            # 外部コールバックがあれば呼ぶ
            if self.on_event_saved_callback:
                self.on_event_saved_callback(cow_auto_id)
        except Exception as e:
            import traceback
            print(f"ERROR: _on_event_saved で例外が発生しました: {e}")
            traceback.print_exc()
    
    def set_on_event_saved(self, callback: Callable[[int], None]):
        """
        イベント保存時のコールバックを設定
        
        Args:
            callback: コールバック関数（cow_auto_id を引数に取る）
        """
        self.on_event_saved_callback = callback
    
    def set_on_latest_dates_calculated(self, callback: Callable[[Optional[str], Optional[str], Optional[str], Optional[str]], None]):
        """
        最終日付計算時のコールバックを設定
        
        Args:
            callback: コールバック関数（latest_calving, latest_ai, latest_milk, latest_any を引数に取る）
        """
        self.on_latest_dates_calculated_callback = callback
    
    def get_widget(self) -> ttk.Frame:
        """ウィジェットを取得"""
        return self.frame

