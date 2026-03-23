"""
FALCON2 - イベント詳細ウィンドウ
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any, Optional, Union
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.event_display import format_insemination_event, build_ai_et_event_note
from constants import FARMS_ROOT


class EventDetailWindow:
    """イベント詳細ウィンドウ"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 event_id: int,
                 event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            event_id: イベントID
            event_dictionary_path: event_dictionary.json のパス
        """
        self.db = db_handler
        self.event_id = event_id
        self.event_dict_path = event_dictionary_path
        
        # event_dictionary を読み込む
        self.event_dictionary: Dict[str, Dict[str, Any]] = {}
        self._load_event_dictionary()
        
        # イベントを取得
        self.event = self._get_event_by_id(event_id)
        if not self.event:
            raise ValueError(f"イベントが見つかりません: event_id={event_id}")
        
        # 授精設定辞書を読み込む（AI/ETイベント用）
        self.technicians_dict: Dict[str, str] = {}
        self.insemination_types_dict: Dict[str, str] = {}
        self._load_insemination_settings()
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("イベント詳細")
        self.window.geometry("600x500")
        
        self._create_widgets()
        self._display_event()
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
    
    def _get_event_by_id(self, event_id: int) -> Optional[Dict[str, Any]]:
        """イベントIDでイベントを取得"""
        conn = self.db.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM event WHERE id = ? AND deleted = 0", (event_id,))
        row = cursor.fetchone()
        
        if row:
            event = dict(row)
            # json_data をパース
            if event.get('json_data'):
                try:
                    event['json_data'] = json.loads(event['json_data'])
                except:
                    event['json_data'] = {}
            return event
        
        return None
    
    def _load_insemination_settings(self):
        """授精設定をロード（AI/ETイベント用）"""
        try:
            # イベントからcow_auto_idを取得
            cow_auto_id = self.event.get('cow_auto_id')
            if not cow_auto_id:
                logging.warning("EventDetailWindow: cow_auto_id not found")
                return
            
            # cowデータを取得してfrmを取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                logging.warning(f"EventDetailWindow: cow not found: cow_auto_id={cow_auto_id}")
                return
            
            frm = cow.get('frm')
            if not frm:
                logging.warning("EventDetailWindow: frm not found")
                return
            
            # 農場パスを構築
            farm_path = FARMS_ROOT / frm
            settings_file = farm_path / "insemination_settings.json"
            
            if not settings_file.exists():
                logging.warning(f"EventDetailWindow: insemination_settings.json not found: {settings_file}")
                return
            
            # 設定ファイルを読み込む
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            self.technicians_dict = settings.get('technicians', {})
            self.insemination_types_dict = settings.get('insemination_types', {})
            
            logging.debug(
                f"EventDetailWindow: Loaded insemination settings: "
                f"technicians={len(self.technicians_dict)}, "
                f"insemination_types={len(self.insemination_types_dict)}"
            )
        except Exception as e:
            logging.error(f"EventDetailWindow: Failed to load insemination settings: {e}")
            self.technicians_dict = {}
            self.insemination_types_dict = {}
    
    def _get_event_name(self, event_number: int) -> str:
        """イベント番号からイベント名を取得"""
        event_str = str(event_number)
        if event_str in self.event_dictionary:
            name_jp = self.event_dictionary[event_str].get('name_jp')
            if name_jp:
                return name_jp
        
        # デフォルト名
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
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_label = ttk.Label(
            self.window,
            text="イベント詳細",
            font=("", 12, "bold")
        )
        title_label.pack(pady=10)
        
        # メインフレーム
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 基本情報フレーム
        basic_frame = ttk.LabelFrame(main_frame, text="基本情報", padding=10)
        basic_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.basic_labels = {}
        basic_fields = [
            ('event_date', '日付'),
            ('event_name', 'イベント'),
            ('note', '備考')
        ]
        
        for i, (key, label) in enumerate(basic_fields):
            row = ttk.Frame(basic_frame)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=f"{label}:", width=15, anchor=tk.W).pack(side=tk.LEFT)
            value_label = ttk.Label(row, text="", foreground="blue")
            value_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.basic_labels[key] = value_label
        
        # 詳細情報フレーム（乳検イベントなど）
        self.detail_frame = ttk.LabelFrame(main_frame, text="詳細情報", padding=10)
        self.detail_frame.pack(fill=tk.BOTH, expand=True)
        
        # 詳細情報表示用のテキストエリア
        self.detail_text = tk.Text(
            self.detail_frame,
            wrap=tk.WORD,
            font=("", 10),
            state=tk.DISABLED,
            height=15
        )
        self.detail_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # スクロールバー
        detail_scrollbar = ttk.Scrollbar(self.detail_frame, orient=tk.VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scrollbar.set)
        detail_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=10)
        
        close_button = ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close,
            width=12
        )
        close_button.pack(side=tk.LEFT, padx=5)
    
    def _display_event(self):
        """イベント情報を表示"""
        event_number = self.event.get('event_number')
        event_date = self.event.get('event_date', '')
        note = self.event.get('note', '')
        json_data = self.event.get('json_data', {})
        
        # 基本情報
        self.basic_labels['event_date'].config(text=event_date)
        self.basic_labels['event_name'].config(text=self._get_event_name(event_number))
        self.basic_labels['note'].config(text=note if note else '(なし)')
        
        # 詳細情報（イベントタイプに応じて表示）
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete(1.0, tk.END)
        
        if event_number == RuleEngine.EVENT_MILK_TEST:
            # 乳検イベントの詳細
            self._display_milk_test_details(json_data)
        else:
            # その他のイベント
            if json_data:
                self._display_json_data(json_data)
            else:
                self.detail_text.insert(tk.END, "詳細情報はありません。")
        
        self.detail_text.config(state=tk.DISABLED)
    
    def _display_milk_test_details(self, json_data: Dict[str, Any]):
        """乳検イベントの詳細を表示"""
        lines = []
        
        # 指定された順序で表示
        if json_data.get('milk_yield') is not None:
            lines.append(f"乳量: {json_data['milk_yield']} kg")
        if json_data.get('fat') is not None:
            lines.append(f"乳脂率: {json_data['fat']} %")
        if json_data.get('snf') is not None:
            lines.append(f"無脂固形分: {json_data['snf']} %")
        if json_data.get('protein') is not None:
            lines.append(f"蛋白率: {json_data['protein']} %")
        if json_data.get('scc') is not None:
            lines.append(f"体細胞数: {json_data['scc']:,}")
        if json_data.get('mun') is not None:
            lines.append(f"乳中尿素窒素（MUN）: {json_data['mun']}")
        if json_data.get('ls') is not None:
            lines.append(f"体細胞スコア（リニアスコア）: {json_data['ls']}")
        if json_data.get('bhb') is not None:
            lines.append(f"BHB: {json_data['bhb']}")
        if json_data.get('denovo_fa') is not None:
            lines.append(f"デノボFA: {json_data['denovo_fa']}")
        if json_data.get('preformed_fa') is not None:
            lines.append(f"プレフォームFA: {json_data['preformed_fa']}")
        if json_data.get('mixed_fa') is not None:
            lines.append(f"ミックスFA: {json_data['mixed_fa']}")
        if json_data.get('denovo_milk') is not None:
            lines.append(f"デノボMilk: {json_data['denovo_milk']}")
        
        if lines:
            self.detail_text.insert(tk.END, "\n".join(lines))
        else:
            self.detail_text.insert(tk.END, "乳検データはありません。")
    
    def _format_insemination_display(self, json_data: Dict[str, Any]) -> str:
        """
        AI/ETイベントの表示文字列を生成（共通関数を使用）
        
        Args:
            json_data: イベントの json_data
        
        Returns:
            表示文字列（例: "14H16102　　Sonoda　　１　P"）
        """
        # 共通関数を使用（CowCardイベント履歴と同じロジック）
        from db.db_handler import DBHandler
        from modules.formula_engine import FormulaEngine
        
        # formula_engineとdb_handlerはイベント詳細ウィンドウでは使用しない（受胎ステータス計算なし）
        note = build_ai_et_event_note(
            self.event,
            self.technicians_dict,
            self.insemination_types_dict,
            formula_engine=None,  # イベント詳細では受胎ステータス計算をスキップ
            db_handler=None
        )
        
        return note
    
    def _display_json_data(self, json_data: Dict[str, Any]):
        """JSONデータを表示"""
        event_number = self.event.get('event_number')
        
        # AI/ETイベントの場合は特別な表示形式（共通関数を使用）
        if event_number in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
            # 共通関数を使用（CowCardイベント履歴と同じロジック）
            # イベント詳細ウィンドウでは受胎ステータス計算をスキップ（formula_engine=None）
            display_text = build_ai_et_event_note(
                self.event,
                self.technicians_dict,
                self.insemination_types_dict,
                formula_engine=None,  # イベント詳細では受胎ステータス計算をスキップ
                db_handler=None
            )
            if display_text:
                self.detail_text.insert(tk.END, display_text)
            else:
                self.detail_text.insert(tk.END, "詳細情報はありません。")
        else:
            # その他のイベントは従来通り
            lines = []
            for key, value in json_data.items():
                if value is not None:
                    lines.append(f"{key}: {value}")
            
            if lines:
                self.detail_text.insert(tk.END, "\n".join(lines))
            else:
                self.detail_text.insert(tk.END, "詳細情報はありません。")
    
    def _on_close(self):
        """閉じる"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

