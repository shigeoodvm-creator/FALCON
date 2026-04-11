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
from modules.app_settings_manager import get_app_settings_manager
from modules.reproduction_checkup_billing import REPRO_CHECKUP_EVENT_NUMBERS
from settings_manager import SettingsManager
from constants import FARMS_ROOT


def normalize_date(date_str: str) -> Optional[str]:
    """
    日付文字列を正規化（YYYY-MM-DD形式に変換）
    
    対応形式：
    - 空欄 → 本日
    - YYYY/MM/DD または YYYY-MM-DD → そのまま採用
    - M/D または M-D → 当年の日付
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
    
    # M/D または M-D 形式（未来日なら前年に補正）
    if '/' in date_str or '-' in date_str:
        parts = date_str.replace('-', '/').split('/')
        if len(parts) == 2:
            try:
                month = int(parts[0])
                day = int(parts[1])
                today = datetime.now()
                date_obj = datetime(today.year, month, day)
                # 未来日なら前年に補正（常に現在より前の日付を採用）
                if date_obj > today:
                    date_obj = datetime(today.year - 1, month, day)
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

    # 同じ種類のイベント入力ウィンドウを再利用するためのアクティブインスタンス
    _active_window: Optional["EventInputWindow"] = None
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler, 
                 rule_engine: RuleEngine,
                 event_dictionary_path: Path,
                 cow_auto_id: Optional[int] = None,  # Noneの場合はメニュー起動
                 event_id: Optional[int] = None,  # 編集時は指定（後方互換性のため残す）
                 on_saved: Optional[Callable[[int], None]] = None,
                 farm_path: Optional[Path] = None,  # 農場パス（SettingsManager用）
                 edit_event_id: Optional[int] = None,  # 編集時のイベントID（event_id の別名）
                 allowed_event_numbers: Optional[List[int]] = None,  # 許可するイベント番号のリスト（Noneの場合は全イベント）
                 default_event_number: Optional[int] = None):  # デフォルトで選択するイベント番号（Noneの場合は選択しない）
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
            allowed_event_numbers: 許可するイベント番号のリスト（Noneの場合は全イベント）
            default_event_number: デフォルトで選択するイベント番号（Noneの場合は選択しない）
        """
        self.db = db_handler
        self.rule_engine = rule_engine
        self.cow_auto_id = cow_auto_id  # Noneの場合は後で入力欄から解決
        self.event_dict_path = event_dictionary_path
        # edit_event_id が指定された場合はそれを優先、なければ event_id を使用
        self.event_id = edit_event_id if edit_event_id is not None else event_id
        self.on_saved = on_saved
        # 許可するイベント番号のリスト（Noneの場合は全イベント）
        self.allowed_event_numbers = allowed_event_numbers
        # デフォルトで選択するイベント番号（Noneの場合は選択しない）
        self.default_event_number = default_event_number
        
        # SettingsManagerを初期化（farm_pathが指定されていない場合はDBから推測）
        if farm_path is None and cow_auto_id is not None:
            # cow_auto_idから農場パスを推測
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                frm = cow.get('frm')
                if frm:
                    farm_path = FARMS_ROOT / frm
        if farm_path is None and hasattr(self.db, 'db_path') and self.db.db_path:
            # 個体カード等から開いた場合、DBの親フォルダを農場パスとする（新規SIRE登録等に必要）
            farm_path = Path(self.db.db_path).parent
        
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
        # 繁殖処置設定データ
        self.treatments: Dict[str, Dict[str, Any]] = {}  # {"1": {"code": "1", "name": "WPG", ...}}
        
        # 授精設定をロード
        self._load_insemination_settings()
        # PEN設定をロード（将来のMOVEイベント等で利用）
        self._load_pen_settings()
        # 繁殖処置設定をロード
        self._load_treatment_settings()
        
        # イベント候補リスト（絞り込み用）
        self.event_candidates: List[Dict[str, Any]] = []  # [{'event_number': int, 'name_jp': str}, ...]
        
        # ウィンドウ作成（他ウィンドウと同一デザイン）
        self.window = tk.Toplevel(parent)
        self.window.title(self._get_window_title())
        self.window.geometry("1020x820")
        self.window.minsize(900, 680)
        self.window.configure(bg="#f5f5f5")
        # 閉じるときにアクティブインスタンスをクリアする
        self.window.protocol("WM_DELETE_WINDOW", self._close)
        
        self._create_widgets()

        # 自分自身をアクティブインスタンスとして登録
        try:
            EventInputWindow._active_window = self
        except Exception:
            pass

        # 編集モードで開いた場合、既存イベントデータをフォームに読み込み、履歴を表示
        if self.event_id is not None:
            self._load_event_data_for_edit()
            self._load_event_history()

    @classmethod
    def open_or_focus(cls,
                      parent: tk.Tk,
                      db_handler: DBHandler,
                      rule_engine: RuleEngine,
                      event_dictionary_path: Path,
                      cow_auto_id: Optional[int] = None,
                      event_id: Optional[int] = None,
                      on_saved: Optional[Callable[[int], None]] = None,
                      farm_path: Optional[Path] = None,
                      edit_event_id: Optional[int] = None,
                      allowed_event_numbers: Optional[List[int]] = None,
                      default_event_number: Optional[int] = None) -> "EventInputWindow":
        """
        同じ種類のイベント入力ウィンドウが既に開いていればそれを前面に出し、
        開いていなければ新規に作成して返す。
        """
        existing = cls._active_window
        if existing is not None:
            window = getattr(existing, "window", None)
            try:
                if window is not None and window.winfo_exists():
                    # 編集モードで呼ばれた場合は既存ウィンドウを編集用に切り替え
                    if edit_event_id is not None:
                        existing.event_id = edit_event_id
                        if cow_auto_id is not None:
                            existing.cow_auto_id = cow_auto_id
                        if on_saved is not None:
                            existing.on_saved = on_saved
                        if farm_path is not None:
                            existing.farm_path = farm_path
                        existing.window.title("イベント編集")
                        existing._load_event_data_for_edit()
                        existing._load_event_history()
                    window.deiconify()
                    window.lift()
                    window.focus_set()
                    return existing
            except Exception:
                # 参照切れなどは無視して新規作成へ
                pass

        # 新規作成
        return cls(
            parent=parent,
            db_handler=db_handler,
            rule_engine=rule_engine,
            event_dictionary_path=event_dictionary_path,
            cow_auto_id=cow_auto_id,
            event_id=event_id,
            on_saved=on_saved,
            farm_path=farm_path,
            edit_event_id=edit_event_id,
            allowed_event_numbers=allowed_event_numbers,
            default_event_number=default_event_number,
        )

    def _is_exit_mode(self) -> bool:
        """売却・死亡廃用（退出）専用で開かれているか"""
        if self.allowed_event_numbers is None or len(self.allowed_event_numbers) != 2:
            return False
        return set(self.allowed_event_numbers) == {RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD}

    def _is_ai_et_only_mode(self) -> bool:
        """繁殖検診フロー等で AI(200)/ET(201) のみ許可で開かれているか"""
        if self.allowed_event_numbers is None or len(self.allowed_event_numbers) != 2:
            return False
        return set(self.allowed_event_numbers) == {RuleEngine.EVENT_AI, RuleEngine.EVENT_ET}

    def _get_window_title(self) -> str:
        """ウィンドウ・ヘッダー用タイトル（退出モード時は「退出」に統一）"""
        if self.event_id is not None:
            return "イベント編集"
        if self._is_exit_mode():
            return "退出"
        return "イベント入力"

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

    def _load_treatment_settings(self):
        """農場設定＞繁殖処置設定（reproduction_treatment_settings.json）をロード。処置＝WPG/CIDR/GN 等の一覧。"""
        if not self.settings_manager:
            return
        
        farm_path = self.settings_manager.farm_path
        settings_file = farm_path / "reproduction_treatment_settings.json"
        
        if not settings_file.exists():
            logging.info(f"reproduction_treatment_settings.json が見つかりません: {settings_file}")
            return
        
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            treatments_raw = settings.get('treatments', {})
            # キーを文字列に統一
            self.treatments = {}
            for key, value in treatments_raw.items():
                self.treatments[str(key)] = value
            
            logging.debug(f"繁殖処置設定をロード: treatments={len(self.treatments)}")
        except Exception as e:
            logging.error(f"reproduction_treatment_settings.json 読み込みエラー: {e}")
            self.treatments = {}
    
    def _get_treatment_ai_days_map(self) -> Dict[str, int]:
        """
        繁殖処置設定から「処置名 → AIまでの日数」のマップを構築
        
        Returns:
            {"WPG": 3, "CIDR": 7, ...} のような辞書
        """
        result: Dict[str, int] = {}
        for code, treatment in self.treatments.items():
            name = treatment.get('name', '').strip()
            if not name:
                continue
            protocols = treatment.get('protocols_by_position') or treatment.get('protocols', [])
            if protocols and isinstance(protocols, list):
                for p in protocols:
                    if p and isinstance(p, dict) and str(p.get('instruction', '')).strip().upper() == 'AI':
                        days = p.get('days')
                        if days is not None and str(days).strip() != '':
                            try:
                                result[name.upper()] = int(days)
                            except (ValueError, TypeError):
                                pass
                        break
        return result
    
    # 授精種類判定で「直近の繁殖関連」に使うイベント（AI, ET, フレッシュ, 繁殖検査, 妊鑑-）
    _LATEST_REPRO_FOR_INSEM_TYPE = (200, 201, 300, 301, 302)

    def _get_most_recent_reproduction_or_ai_et_before_date(self, events: List[Dict[str, Any]], before_date: str) -> Optional[Dict[str, Any]]:
        """AI日付より前の直近の繁殖関連イベントを取得（AI, ET, フレッシュ, 繁殖検査, 妊鑑- のいずれか）"""
        try:
            before_dt = datetime.strptime(before_date, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        for ev in events:
            ev_date = ev.get('event_date', '')
            if not ev_date:
                continue
            try:
                ev_dt = datetime.strptime(ev_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            if ev_dt >= before_dt:
                continue
            if ev.get('event_number') in self._LATEST_REPRO_FOR_INSEM_TYPE:
                return ev
        return None

    # 所見であり処置でない値（農場設定＞繁殖処置設定の「処置」は WPG/CIDR/GN 等のみ。
    # 子宮OK・NS 等は所見で、入力時の子宮・右・左・その他は所見。これらが処置欄に入っている場合は
    # 授精種類判定では処置なしとして自然発情とする）
    _FINDING_VALUES_NOT_TREATMENT = frozenset({'', '-', 'NS', 'N.S.', '子宮OK', '処置なし', 'なし', '無'})

    def _get_treatment_name_from_event(self, ev: Dict[str, Any]) -> Optional[str]:
        """
        イベントの json_data から処置名を取得する。
        処置＝農場設定＞繁殖処置設定で設定される項目（WPG, CIDR, GN, PG, E2 等）のみ。
        子宮OK・NS 等は所見（子宮・右・左・その他は所見欄）であり処置ではないため、
        処置欄に所見が入っている場合は None を返し授精種類は自然発情とする。
        """
        json_data = ev.get('json_data') or {}
        treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
        treatment = str(treatment).strip() if treatment else ''
        if not treatment or treatment == '-':
            return None
        if treatment.isdigit():
            t = self.treatments.get(str(treatment), {})
            name = (t.get('name', treatment) or treatment).strip() if isinstance(t, dict) else treatment.strip()
            if not name or name in self._FINDING_VALUES_NOT_TREATMENT:
                return None
            return name
        if treatment in self._FINDING_VALUES_NOT_TREATMENT:
            return None
        return treatment.strip() or None

    def _get_treatment_ai_days_map(self) -> Dict[str, int]:
        """処置名（大文字）→ 〇日後AI のマップを構築"""
        result: Dict[str, int] = {}
        for treatment in (self.treatments or {}).values():
            if not isinstance(treatment, dict):
                continue
            name = (treatment.get('name') or '').strip()
            if not name:
                continue
            protocols = treatment.get('protocols_by_position') or treatment.get('protocols') or []
            if not isinstance(protocols, list):
                continue
            for p in protocols:
                if not p or not isinstance(p, dict):
                    continue
                if str(p.get('instruction', '')).strip().upper() == 'AI':
                    days = p.get('days')
                    if days is not None and str(days).strip() != '':
                        try:
                            result[name.upper()] = int(days)
                        except (ValueError, TypeError):
                            pass
                    break
        return result
    
    def _find_natural_estrus_insemination_type(self) -> Optional[str]:
        """
        授精種類設定から A:自然発情 に該当する項目を検索
        コード "A" を優先、なければ「自然発情」を含む項目を返す
        
        Returns:
            "code：name" 形式の表示文字列、該当なければ None
        """
        fallback = None
        for code, name in self.insemination_types.items():
            code_str = str(code).strip().upper()
            name_str = str(name).strip()
            if '自然発情' in name_str:
                if fallback is None:
                    fallback = f"{code}：{name}"
                if code_str == 'A':
                    return f"{code}：{name}"
        return fallback
    
    def _compute_suggested_insemination_type(self) -> Optional[str]:
        """
        AI/ET日付入力時点で、直近の繁殖関連イベントから授精種類を推定する。
        処置＝農場設定＞繁殖処置設定の項目（WPG, CIDR 等）。所見＝子宮・右・左・その他（子宮OK, NS 等は所見で NOTE にも出る）。
        1) 直近の繁殖関連（AI, ET, フレッシュ, 繁殖検査, 妊鑑-）を1件確定
        2) 直近がAI/ET → 自然発情
        3) 直近が繁殖検査/妊鑑-で処置なし、または処置が授精種類とリンクしていない（例: GN）→ 自然発情
        4) 直近が繁殖検査/妊鑑-で処置が授精種類とリンクし、〇日後AIと日付が±1日で一致 → その授精種類
        5) 上記以外（〇日後AIと一致しない等）→ 空欄（None）
        """
        if self.cow_auto_id is None:
            return None
        if self.selected_event_number not in (200, 201):
            return None
        date_str = self.date_entry.get().strip() if hasattr(self, 'date_entry') else ''
        ai_date = normalize_date(date_str)
        if not ai_date:
            return None
        events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
        natural_estrus = self._find_natural_estrus_insemination_type()
        latest_repro = self._get_most_recent_reproduction_or_ai_et_before_date(events, ai_date)
        if not latest_repro:
            return natural_estrus

        # 2) 直近がAI/ET → 自然発情
        if latest_repro.get('event_number') in (200, 201):
            return natural_estrus

        # 直近がフレッシュ(300)/繁殖検査(301)/妊鑑-(302)の場合
        if latest_repro.get('event_number') not in (300, 301, 302):
            return natural_estrus

        treatment_name = self._get_treatment_name_from_event(latest_repro)
        treatment_date = latest_repro.get('event_date', '')
        if not treatment_date:
            return natural_estrus

        # 処置なし → 自然発情
        if not treatment_name:
            return natural_estrus

        # 処置が授精種類とリンクしているか
        treatment_upper = treatment_name.strip().upper()
        matched_display = None
        for code, name in self.insemination_types.items():
            name_upper = (str(name).strip()).upper()
            if treatment_upper == name_upper or treatment_upper in name_upper:
                matched_display = f"{code}：{name}"
                break
        if matched_display is None:
            # 3) リンクしていない（例: GN）→ 自然発情
            return natural_estrus

        # 〇日後AIが設定されているか
        ai_days_map = self._get_treatment_ai_days_map()
        ai_days = ai_days_map.get(treatment_upper)
        if ai_days is None:
            # 処置が「自然発情」に一致している場合は自然発情とする（〇日後AIは不要）
            if natural_estrus and '自然発情' in (matched_display or ''):
                return natural_estrus
            # 6) 〇日後AIなし → 空欄
            return None

        # 5) 入力日付が 処置日 + 〇日 ±1 の範囲か
        try:
            td = datetime.strptime(treatment_date, '%Y-%m-%d')
            day_low = (td + timedelta(days=ai_days - 1)).strftime('%Y-%m-%d')
            day_high = (td + timedelta(days=ai_days + 1)).strftime('%Y-%m-%d')
            if day_low <= ai_date <= day_high:
                return matched_display
        except (ValueError, TypeError):
            pass
        # 6) 日付が一致しない → 空欄
        return None
    
    def _try_auto_fill_insemination_type(self) -> None:
        """AI/ET入力時、授精種類を自動表示（ユーザーは修正可能）。空欄の場合は''をセット"""
        if 'insemination_type_code' not in self.field_widgets:
            return
        try:
            self.window.update_idletasks()
        except tk.TclError:
            pass
        suggested = self._compute_suggested_insemination_type()
        widget = self.field_widgets['insemination_type_code']
        if hasattr(widget, 'get') and hasattr(widget, 'set'):
            widget.set(suggested if suggested else '')
    
    def _create_widgets(self):
        """ウィジェットを作成（他ウィンドウと同一デザイン）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        # 罫線を出さないスタイル（このウィンドウ用の LabelFrame）
        try:
            _style = ttk.Style()
            _style.configure("EventInput.TLabelframe", borderwidth=0)
            _style.configure("EventInput.TLabelframe.Label", font=(_df, 10, "bold"))
            # 左カラム入力欄用：フォントをやや大きくして見やすく、幅を統一してプルダウンが使えるようにする
            try:
                _input_font_size = max(11, get_app_settings_manager().get_font_size() + 3)
            except Exception:
                _input_font_size = 12
            _style.configure("EventInput.TEntry", font=(_df, _input_font_size))
            _style.configure("EventInput.TCombobox", font=(_df, _input_font_size))
            _style.configure("EventInput.Hint.TLabel", font=(_df, 8), foreground="#78909c")
        except tk.TclError:
            pass
        # 入力欄の統一幅（文字数）と入力列の最小幅（px）。伸ばしすぎないことでプルダウンが常に使える
        _ENTRY_WIDTH = 24
        _ENTRY_COLUMN_MINSIZE = 300
        self._entry_width = _ENTRY_WIDTH
        self._entry_col_minsize = _ENTRY_COLUMN_MINSIZE
        
        main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # ヘッダー（タイトル左・対象牛＋OK/キャンセル右でボタンが切れないように）
        header = tk.Frame(main_container, bg=bg, pady=12)
        header.pack(fill=tk.X)
        _header_icon = "\U0001f6aa" if self._is_exit_mode() else "\u270d\ufe0f"
        tk.Label(header, text=_header_icon, font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        title_text = self._get_window_title()
        tk.Label(title_frame, text=title_text, font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        _subtitle = (
            "売却・死亡廃用の入力を行います。牛ID・イベントを選択し、日付や項目を入力して保存します。"
            if self._is_exit_mode()
            else "牛ID・イベントを選択し、日付や項目を入力して保存します"
        )
        tk.Label(title_frame, text=_subtitle, font=(_df, 10), bg=bg, fg="#607d8b").pack(anchor=tk.W)
        # ヘッダー右：対象牛（個体指定時）＋ OK / 削除 / キャンセル
        header_right = tk.Frame(header, bg=bg)
        header_right.pack(side=tk.RIGHT, padx=(10, 0))

        # 対象牛表示ブロック
        # 小さい「対象牛」ラベル + 大きなID で、このカードがどの牛かを一目で分かるようにする
        cow_header_block = tk.Frame(header_right, bg=bg)
        cow_header_block.pack(side=tk.LEFT, padx=(0, 16))
        tk.Label(
            cow_header_block,
            text="対象牛",
            font=(_df, 9),
            bg=bg,
            fg="#607d8b",
        ).pack(anchor=tk.E)
        self.header_cow_label = tk.Label(
            cow_header_block,
            text="",
            font=(_df, 18, "bold"),  # 個体カードのIDと同等のサイズ感
            bg=bg,
            fg="#1565c0",
        )
        self.header_cow_label.pack(anchor=tk.E)

        ok_btn = ttk.Button(header_right, text="OK", command=self._on_ok, width=8)
        ok_btn.pack(side=tk.LEFT, padx=(0, 5))
        if self.event_id:
            ttk.Button(header_right, text="削除", command=self._on_delete, width=8).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(header_right, text="キャンセル", command=self._on_cancel, width=10).pack(side=tk.LEFT)
        
        content_area = tk.Frame(main_container, bg=bg)
        content_area.pack(fill=tk.BOTH, expand=True)
        
        # イベント情報表示（特定のイベントに固定されている場合、または退出モード）
        # 表示は日本語名のみ（alias の CALV 等は出さない）
        if self._is_exit_mode():
            event_info_frame = tk.Frame(content_area, bg="#e8eaf6", pady=12, padx=16)
            event_info_frame.pack(fill=tk.X, pady=(0, 8))
            event_info_label = tk.Label(
                event_info_frame,
                text="退出（売却・死亡廃用）",
                font=(_df, 15, "bold"),
                bg="#e8eaf6",
                fg="#283593",
            )
            event_info_label.pack(anchor=tk.W)
        elif self._is_ai_et_only_mode():
            event_info_frame = tk.Frame(content_area, bg="#e8eaf6", pady=12, padx=16)
            event_info_frame.pack(fill=tk.X, pady=(0, 8))
            event_info_label = tk.Label(
                event_info_frame,
                text="AI/ET",
                font=(_df, 15, "bold"),
                bg="#e8eaf6",
                fg="#283593",
            )
            event_info_label.pack(anchor=tk.W)
        elif self.allowed_event_numbers is not None and len(self.allowed_event_numbers) == 1:
            event_number = self.allowed_event_numbers[0]
            event_num_str = str(event_number)
            event_data = self.event_dictionary.get(event_num_str, {})
            name_jp = event_data.get('name_jp', f'イベント{event_number}')
            
            event_info_frame = tk.Frame(content_area, bg="#e8eaf6", pady=12, padx=16)
            event_info_frame.pack(fill=tk.X, pady=(0, 8))
            event_info_label = tk.Label(
                event_info_frame,
                text=name_jp,
                font=(_df, 15, "bold"),
                bg="#e8eaf6",
                fg="#283593"
            )
            event_info_label.pack(anchor=tk.W)
        
        # メイン左右レイアウト（罫線なし）
        main_split = ttk.Frame(content_area)
        main_split.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左カラム：イベント入力フォーム ==========
        left_frame = ttk.Frame(main_split, width=520)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        left_frame.pack_propagate(False)  # width を維持
        
        # 左カラムを縦方向に分割（スクロール可能な領域とボタン領域）
        left_container = ttk.Frame(left_frame)
        left_container.pack(fill=tk.BOTH, expand=True)
        
        # スクロール可能なフレーム（入力項目用）
        left_canvas = tk.Canvas(left_container, highlightthickness=0, bg="#fafafa")
        left_scrollbar = ttk.Scrollbar(left_container, orient=tk.VERTICAL, command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas)
        
        left_scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        
        left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        def _on_left_mousewheel(evt):
            if evt.delta:
                left_canvas.yview_scroll(int(-evt.delta / 60), "units")
            return "break"

        def _on_linux_scroll_up(_evt):
            left_canvas.yview_scroll(-3, "units")
            return "break"

        def _on_linux_scroll_down(_evt):
            left_canvas.yview_scroll(3, "units")
            return "break"

        left_canvas.bind("<MouseWheel>", _on_left_mousewheel)
        left_canvas.bind("<Button-4>", _on_linux_scroll_up)
        left_canvas.bind("<Button-5>", _on_linux_scroll_down)
        left_scrollable_frame.bind("<MouseWheel>", _on_left_mousewheel)
        left_scrollable_frame.bind("<Button-4>", _on_linux_scroll_up)
        left_scrollable_frame.bind("<Button-5>", _on_linux_scroll_down)

        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ラベル列を統一幅にして入力欄を縦に揃える（全フレームで column 0 を同じ幅に）
        _label_col_minsize = 100
        
        # ========== 1. 牛ID入力欄（メニュー起動時のみ） ==========
        if self.cow_auto_id is None:
            cow_id_frame = ttk.LabelFrame(left_scrollable_frame, text="牛ID入力", padding=(12, 10), style="EventInput.TLabelframe")
            cow_id_frame.pack(fill=tk.X, pady=(0, 14))
            cow_id_frame.columnconfigure(0, minsize=_label_col_minsize)
            cow_id_frame.columnconfigure(1, weight=0, minsize=_ENTRY_COLUMN_MINSIZE)
            
            ttk.Label(cow_id_frame, text="牛ID*:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
            self.cow_id_entry = ttk.Entry(cow_id_frame, width=_ENTRY_WIDTH, style="EventInput.TEntry")
            self.cow_id_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
            self.cow_id_entry.bind('<KeyRelease>', self._on_cow_id_changed)
            self.cow_id_entry.bind('<Return>', lambda e: self._on_cow_id_enter())
            ttk.Label(cow_id_frame, text="拡大4桁ID または JPN10", style="EventInput.Hint.TLabel").grid(
                row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(0, 4)
            )
            
            # 対象牛表示ラベル（初期は非表示）
            # → 少し大きめ・青字で、このカードがどの牛かを強調
            self.cow_info_label = ttk.Label(
                cow_id_frame,
                text="",
                foreground="#1565c0",
                font=(_df, 12, "bold"),
            )
            self.cow_info_label.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
            
            # 候補リスト用のフレーム
            self.cow_candidate_frame = ttk.Frame(cow_id_frame)
            self.cow_candidate_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W+tk.E, padx=5, pady=5)
            
            # 候補リスト用のTreeview
            candidate_columns = ('cow_id', 'jpn10')
            self.cow_candidate_tree = ttk.Treeview(
                self.cow_candidate_frame,
                columns=candidate_columns,
                show='headings',
                height=5
            )
            self.cow_candidate_tree.heading('cow_id', text='牛ID')
            self.cow_candidate_tree.heading('jpn10', text='個体識別番号')
            
            self.cow_candidate_tree.column('cow_id', width=100)
            self.cow_candidate_tree.column('jpn10', width=200)
            
            # スクロールバー
            candidate_scrollbar = ttk.Scrollbar(
                self.cow_candidate_frame,
                orient=tk.VERTICAL,
                command=self.cow_candidate_tree.yview
            )
            self.cow_candidate_tree.configure(yscrollcommand=candidate_scrollbar.set)
            
            self.cow_candidate_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            candidate_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # ダブルクリックで選択
            self.cow_candidate_tree.bind('<Double-Button-1>', self._on_cow_candidate_selected)
            
            # 初期状態では候補リストを非表示
            self.cow_candidate_frame.grid_remove()
        
        # ========== 2. 対象牛表示ラベル（個体カード起動時） ==========
        else:
            cow_info_frame = tk.Frame(left_scrollable_frame, bg="#e3f2fd", padx=14, pady=10)
            cow_info_frame.pack(fill=tk.X, pady=(0, 14))
            self.cow_info_label = tk.Label(
                cow_info_frame,
                text="",
                font=(_df, 15, "bold"),
                fg="#1565c0",
                bg="#e3f2fd"
            )
            self.cow_info_label.pack(side=tk.LEFT)
            self._show_cow_info()
        
        # ========== 3. イベント番号入力欄 ==========
        event_frame = ttk.LabelFrame(left_scrollable_frame, text="イベント選択", padding=(12, 10), style="EventInput.TLabelframe")
        event_frame.pack(fill=tk.X, pady=(0, 14))
        event_frame.columnconfigure(0, minsize=_label_col_minsize)
        event_frame.columnconfigure(1, weight=0, minsize=_ENTRY_COLUMN_MINSIZE)
        
        ttk.Label(event_frame, text="イベント番号:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        self.event_number_entry = ttk.Entry(event_frame, width=_ENTRY_WIDTH, style="EventInput.TEntry")
        self.event_number_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.event_number_entry.bind('<KeyRelease>', self._on_event_number_changed)
        self.event_number_entry.bind('<Return>', self._on_event_number_enter)
        self.event_number_entry.bind('<Tab>', self._on_event_number_tab)
        _event_num_hint = (
            f"AI：{RuleEngine.EVENT_AI}、ET：{RuleEngine.EVENT_ET}"
            if self._is_ai_et_only_mode()
            else "例: 1xx=体況/生産, 2xx=繁殖・分娩・管理, 3xx=繁殖検査/妊娠,\n4xx=疾病, 6xx=導入/群管理/タスク"
        )
        ttk.Label(
            event_frame,
            text=_event_num_hint,
            style="EventInput.Hint.TLabel"
        ).grid(
            row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(0, 4)
        )
        
        # ========== 4. イベント候補プルダウン ==========
        ttk.Label(event_frame, text="イベント候補:").grid(row=2, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        self.event_combo = ttk.Combobox(event_frame, width=_ENTRY_WIDTH, state="readonly", style="EventInput.TCombobox")
        self.event_combo.grid(row=2, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.event_combo.bind('<<ComboboxSelected>>', self._on_event_combo_selected)
        
        # イベント選択欄を非表示にする条件：allowed_event_numbersが1つだけの場合
        if self.allowed_event_numbers is not None and len(self.allowed_event_numbers) == 1:
            # イベント選択欄を非表示
            event_frame.pack_forget()
        
        # ========== 5. 日付入力欄 ==========
        date_frame = ttk.LabelFrame(left_scrollable_frame, text="共通項目", padding=(12, 10), style="EventInput.TLabelframe")
        date_frame.pack(fill=tk.X, pady=(0, 14))
        date_frame.columnconfigure(0, minsize=_label_col_minsize)
        date_frame.columnconfigure(1, weight=0, minsize=_ENTRY_COLUMN_MINSIZE)
        
        ttk.Label(date_frame, text="日付*:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        self.date_entry = ttk.Entry(date_frame, width=_ENTRY_WIDTH, style="EventInput.TEntry")
        self.date_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        today = datetime.now().strftime('%Y-%m-%d')
        self.date_entry.insert(0, today)
        ttk.Label(date_frame, text="YYYY/MM/DD または M/D, D", style="EventInput.Hint.TLabel").grid(
            row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(0, 4)
        )
        self.date_entry.bind('<Return>', lambda e: self._move_from_date_field(e))
        self.date_entry.bind('<FocusOut>', lambda e: self._try_auto_fill_insemination_type())
        
        # ========== 6. イベント詳細入力欄（動的） ==========
        self.form_frame = ttk.LabelFrame(left_scrollable_frame, text="入力項目", padding=(12, 10), style="EventInput.TLabelframe")
        self.form_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 14))
        
        # ========== 7. メモ欄 ==========
        ttk.Label(date_frame, text="メモ:").grid(row=2, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        self.note_entry = ttk.Entry(date_frame, width=_ENTRY_WIDTH, style="EventInput.TEntry")
        self.note_entry.grid(row=2, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        # メモ欄でEnterキーを押すとOKボタンにフォーカス（または最初の入力項目に戻る）
        self.note_entry.bind('<Return>', lambda e: self._move_from_note_field(e))
        
        # ========== 右カラム：イベント履歴表示（フォントは MS ゴシック） ==========
        right_frame = ttk.Frame(main_split, width=400)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        history_frame = ttk.LabelFrame(right_frame, text="イベント履歴", padding=10, style="EventInput.TLabelframe")
        history_frame.pack(fill=tk.BOTH, expand=True)
        
        # イベント履歴用スタイル（MS Gothic）
        try:
            _font_size = get_app_settings_manager().get_font_size()
        except Exception:
            _font_size = 10
        try:
            style = ttk.Style()
            style.configure("EventInputHistory.Treeview", font=("MS Gothic", _font_size))
            style.configure("EventInputHistory.Treeview.Heading", font=("MS Gothic", _font_size))
        except tk.TclError:
            pass
        # Treeview
        columns = ('date', 'event', 'info')
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show='headings', height=20, style='EventInputHistory.Treeview')
        
        self.history_tree.heading('date', text='日付')
        self.history_tree.heading('event', text='イベント')
        self.history_tree.heading('info', text='NOTE')
        
        self.history_tree.column('date', width=100)
        self.history_tree.column('event', width=80, stretch=False)
        self.history_tree.column('info', width=200, stretch=True)
        
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
        
        # 初期状態で全イベントを候補として表示
        self._update_event_candidates("", auto_select=False)
        
        # デフォルトで選択するイベント番号が指定されている場合は自動選択
        if self.default_event_number is not None:
            # 許可されたイベント番号のリストがある場合は確認
            if self.allowed_event_numbers is None or self.default_event_number in self.allowed_event_numbers:
                # イベント番号を文字列に変換して選択
                default_event_str = str(self.default_event_number)
                if default_event_str in self.event_dictionary:
                    self._select_event_by_number(default_event_str)
    
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
                # Combobox も更新（イベント番号：日本語名の形式）
                event_data = self.event_dictionary[event_num_str]
                name_jp = event_data.get('name_jp', f'イベント{event_number}')
                self.event_combo.set(f"{event_number}：{name_jp}")
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
            # 空欄の場合は全イベントを候補に（許可されたイベントのみ）
            for event_num_str, event_data in self.event_dictionary.items():
                if event_data.get('deprecated', False):
                    continue
                event_number = int(event_num_str)
                # 許可されたイベント番号のリストがある場合はフィルタリング
                if self.allowed_event_numbers is not None:
                    if event_number not in self.allowed_event_numbers:
                        continue
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
                # 許可されたイベント番号のリストがある場合はフィルタリング
                if self.allowed_event_numbers is not None:
                    if event_number not in self.allowed_event_numbers:
                        continue
                
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
        
        # プルダウンに表示する形式: イベント番号：日本語名
        combo_values = [f"{c['event_number']}：{c['name_jp']}" for c in candidates]
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
        
        # 選択された値からイベント番号を特定
        # 表示形式は「イベント番号：日本語名」なので、「：」で分割してイベント番号を取得
        # event_candidatesの順序とComboboxのvaluesの順序は一致している
        try:
            current_values = list(self.event_combo['values'])
            if selection in current_values:
                index = current_values.index(selection)
                if 0 <= index < len(self.event_candidates):
                    event_num_str = self.event_candidates[index]['event_num_str']
                    self._select_event(event_num_str)
            else:
                # 直接入力された場合（「：」で分割してイベント番号を取得）
                if "：" in selection or ":" in selection:
                    separator = "：" if "：" in selection else ":"
                    parts = selection.split(separator, 1)
                    if parts[0].strip().isdigit():
                        event_num_str = parts[0].strip()
                        self._select_event(event_num_str)
        except (ValueError, IndexError, KeyError):
            pass
    
    def _on_cow_id_changed(self, event):
        """牛ID入力変更時の処理（リアルタイムで候補を絞り込む）"""
        # 特殊キー（Enter, Tab, Shift等）の場合は処理しない
        if event.keysym in ('Return', 'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Up', 'Down'):
            return
        
        cow_id_str = self.cow_id_entry.get().strip()
        
        if not cow_id_str:
            # 空欄の場合は候補リストを非表示
            self.cow_candidate_frame.grid_remove()
            return
        
        # 数字のみの場合は前方一致で検索
        if cow_id_str.isdigit():
            candidates = self.db.search_cows_by_id_prefix(cow_id_str, limit=50)
            self._update_cow_candidates(candidates)
        else:
            # 数字以外が含まれている場合は候補リストを非表示
            self.cow_candidate_frame.grid_remove()
    
    def _update_cow_candidates(self, candidates: List[Dict[str, Any]]):
        """
        牛の候補リストを更新
        
        Args:
            candidates: 候補牛のリスト
        """
        # 既存のアイテムをクリア
        for item in self.cow_candidate_tree.get_children():
            self.cow_candidate_tree.delete(item)
        
        if not candidates:
            # 候補がない場合は非表示
            self.cow_candidate_frame.grid_remove()
            return
        
        # 候補を表示
        for cow in candidates:
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            
            # auto_idをtagsに保存
            auto_id = cow.get('auto_id')
            self.cow_candidate_tree.insert(
                '',
                'end',
                values=(cow_id, jpn10),
                tags=(f"cow_{auto_id}",)
            )
        
        # 候補リストを表示
        self.cow_candidate_frame.grid()
    
    def _on_cow_candidate_selected(self, event):
        """候補リストでダブルクリックされた時の処理（またはEnterキーで自動選択時）"""
        # eventがNoneの場合はEnterキーで自動選択された場合
        if event is not None:
            selected_items = self.cow_candidate_tree.selection()
            if not selected_items:
                return
            item = selected_items[0]
        else:
            # Enterキーで自動選択された場合
            candidates = self.cow_candidate_tree.get_children()
            if not candidates:
                return
            item = candidates[0]
        
        tags = self.cow_candidate_tree.item(item, 'tags')
        
        if not tags:
            return
        
        # tagsからauto_idを取得
        auto_id = None
        for tag in tags:
            if tag.startswith('cow_'):
                auto_id_str = tag.replace('cow_', '')
                try:
                    auto_id = int(auto_id_str)
                    break
                except ValueError:
                    continue
        
        if auto_id is None:
            return
        
        # 牛IDを解決
        cow = self.db.get_cow_by_auto_id(auto_id)
        if cow:
            self.cow_auto_id = auto_id
            
            # SettingsManagerを初期化（まだ初期化されていない場合）
            if self.settings_manager is None:
                frm = cow.get('frm')
                if frm:
                    farm_path = FARMS_ROOT / frm
                    self.settings_manager = SettingsManager(farm_path)
                    # 授精設定・繁殖処置設定を再読み込み
                    self._load_insemination_settings()
                    self._load_treatment_settings()
            
            # 牛ID入力欄に値を設定
            cow_id = cow.get('cow_id', '')
            self.cow_id_entry.delete(0, tk.END)
            self.cow_id_entry.insert(0, cow_id)
            
            # 対象牛情報を表示
            self._show_cow_info()
            # イベント履歴を更新
            self._load_event_history()
            # 候補リストを非表示
            self.cow_candidate_frame.grid_remove()
            # 授精種類を自動表示（AI/ET選択時）
            self.window.after_idle(self._try_auto_fill_insemination_type)
            # イベントが選択されていない場合、デフォルトで選択するイベント番号があれば選択
            if self.selected_event_number is None and self.default_event_number is not None:
                # 許可されたイベント番号のリストがある場合は確認
                if self.allowed_event_numbers is None or self.default_event_number in self.allowed_event_numbers:
                    # イベント番号を文字列に変換して選択
                    default_event_str = str(self.default_event_number)
                    if default_event_str in self.event_dictionary:
                        self._select_event_by_number(default_event_str)
            
            # フォーカスをイベント番号入力欄に移動
            self.event_number_entry.focus()
    
    def _on_cow_id_enter(self):
        """牛ID入力欄でEnterキー押下時"""
        if not hasattr(self, 'cow_id_entry'):
            return
        
        cow_id_str = self.cow_id_entry.get().strip()
        if not cow_id_str:
            return
        
        # 候補リストが表示されている場合は、候補が1件だけなら自動選択
        if hasattr(self, 'cow_candidate_frame') and self.cow_candidate_frame.winfo_viewable():
            candidates = self.cow_candidate_tree.get_children()
            if len(candidates) == 1:
                # 候補が1件だけの場合は自動選択
                self.cow_candidate_tree.selection_set(candidates[0])
                self._on_cow_candidate_selected(None)
                return
            elif len(candidates) > 1:
                # 候補が複数ある場合はエラーを出さず、候補リストから選択を促す
                return
        
        # 候補リストが表示されていない場合、または候補がない場合は通常の解決処理
        resolved_auto_id = self._resolve_cow_id(cow_id_str)
        if resolved_auto_id:
            self.cow_auto_id = resolved_auto_id
            
            # SettingsManagerを初期化（まだ初期化されていない場合）
            if self.settings_manager is None:
                cow = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow:
                    frm = cow.get('frm')
                    if frm:
                        farm_path = FARMS_ROOT / frm
                        self.settings_manager = SettingsManager(farm_path)
                        # 授精設定・繁殖処置設定を再読み込み
                        self._load_insemination_settings()
                        self._load_treatment_settings()
            
            # 対象牛情報を表示
            self._show_cow_info()
            # イベント履歴を更新
            self._load_event_history()
            # 候補リストを非表示
            if hasattr(self, 'cow_candidate_frame'):
                self.cow_candidate_frame.grid_remove()
            # 授精種類を自動表示（AI/ET選択時）
            self.window.after_idle(self._try_auto_fill_insemination_type)
            # フォーカスをイベント番号入力欄に移動
            self.event_number_entry.focus()
        else:
            # 候補リストが表示されていない場合のみエラーを表示
            if not (hasattr(self, 'cow_candidate_frame') and self.cow_candidate_frame.winfo_viewable()):
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
            # ヘッダー右上：大きなIDのみ表示（左に小さく「対象牛」ラベル）
            if hasattr(self, 'header_cow_label'):
                self.header_cow_label.config(text=cow_id, fg="#1565c0")
            # 入力欄直下のラベルや情報枠では「対象牛：XXXX」とテキスト付きで表示
            info_text = f"対象牛：{cow_id}"
            if hasattr(self, 'cow_info_label'):
                self.cow_info_label.config(text=info_text, foreground="#1565c0")
    
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
        # 入力項目エリアのラベル列を統一幅に。入力列は伸ばさず固定幅でプルダウンを有効に
        self.form_frame.columnconfigure(0, minsize=110)
        self.form_frame.columnconfigure(1, weight=0, minsize=self._entry_col_minsize)
        
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
        
        # 導入イベント（600）の特別処理
        if event_number == 600:  # 導入イベント
            self._create_introduction_form()
            return
        
        # 通常のイベント処理
        input_fields = self.selected_event.get('input_fields', [])
        
        # 妊娠鑑定プラスイベント（303, 304, 307）の場合、双子・♀判定チェックボックスが存在しない場合は追加
        if event_number in [303, 304, 307]:  # PDP, PDP2, PAGP
            has_twin_field = any(field.get('key') == 'twin' for field in input_fields)
            if not has_twin_field:
                input_fields.append({
                    'key': 'twin',
                    'label': '双子',
                    'datatype': 'bool'
                })
            has_female_judgment_field = any(field.get('key') == 'female_judgment' for field in input_fields)
            if not has_female_judgment_field:
                input_fields.append({
                    'key': 'female_judgment',
                    'label': '♀',
                    'datatype': 'bool'
                })
            has_male_judgment_field = any(field.get('key') == 'male_judgment' for field in input_fields)
            if not has_male_judgment_field:
                input_fields.append({
                    'key': 'male_judgment',
                    'label': '♂',
                    'datatype': 'bool'
                })
        
        if not input_fields:
            ttk.Label(
                self.form_frame,
                text="入力項目はありません",
                foreground="gray"
            ).pack(pady=10)
            return
        
        # 各入力フィールドを生成
        entry_widgets = []  # Enterキーで移動する順序を保持
        row_index = 0
        treatment_settings = self.treatments if event_number in [300, 301, 302] else {}
        for i, field in enumerate(input_fields):
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            label_text = field.get('label', key)
            
            # チェックボックスの場合
            if datatype in ['bool', 'boolean']:
                # チェックボックス用の変数
                var = tk.BooleanVar()
                # Checkbuttonウィジェットへの参照も保持するため、varを保存
                # （実際のウィジェットはCheckbuttonだが、値の取得はvarから行う）
                self.field_widgets[key] = var
                
                # チェックボックス
                checkbox = ttk.Checkbutton(
                    self.form_frame,
                    text=label_text,
                    variable=var
                )
                checkbox.grid(row=row_index, column=0, columnspan=2, sticky=tk.W, padx=(5, 8), pady=(6, 2))
                # Checkbuttonウィジェットへの参照も保存（フォーカス移動用）
                # key_checkboxというキーで保存
                self.field_widgets[f"{key}_checkbox"] = checkbox
                row_index += 1
                continue

            # choices 付き文字列（例: BLV結果＝未検査/陰性/陽性）
            choices = field.get("choices")
            if choices and datatype == "str" and isinstance(choices, list) and len(choices) > 0:
                ttk.Label(
                    self.form_frame,
                    text=f"{label_text}:",
                ).grid(row=row_index, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
                combo = ttk.Combobox(
                    self.form_frame,
                    values=list(choices),
                    width=self._entry_width,
                    state="readonly",
                    style="EventInput.TCombobox",
                )
                combo.set(choices[0])
                combo.grid(row=row_index, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
                self.field_widgets[key] = combo
                entry_widgets.append(combo)
                row_index += 1
                continue
            
            # ラベル
            ttk.Label(
                self.form_frame,
                text=f"{label_text}:"
            ).grid(row=row_index, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
            
            # 入力ウィジェット
            if event_number in [300, 301, 302] and key == 'treatment':
                # 処置入力はプルダウン（Combobox）＋直接入力
                values = []
                for code, treatment_data in sorted(treatment_settings.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 999):
                    name = treatment_data.get('name', '')
                    if code and name:
                        values.append(f"{code}：{name}")
                
                combo = ttk.Combobox(
                    self.form_frame,
                    values=values,
                    width=self._entry_width,
                    state='normal',  # 直接入力も可能
                    style="EventInput.TCombobox"
                )
                combo.grid(row=row_index, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
                self.field_widgets[key] = combo
                entry_widgets.append(combo)
                
                # 選択時に処置内容のみを入力欄に設定
                def on_treatment_selected(event):
                    selected = combo.get()
                    if selected and '：' in selected:
                        parts = selected.split('：', 1)
                        if len(parts) == 2:
                            treatment_name = parts[1].strip()
                            combo.set(treatment_name)
                
                combo.bind('<<ComboboxSelected>>', on_treatment_selected)
                # 大文字変換 + 数字入力時の自動変換
                def on_treatment_key(event):
                    value = combo.get().strip()
                    if value.isdigit():
                        treatment = self.treatments.get(str(value))
                        if treatment:
                            treatment_name = treatment.get('name', '')
                            if treatment_name:
                                combo.set(treatment_name)
                                return
                    self._convert_to_uppercase_combobox(combo)
                
                combo.bind('<KeyRelease>', on_treatment_key)
                
                row_index += 1
                continue

            entry = ttk.Entry(self.form_frame, width=self._entry_width, style="EventInput.TEntry")
            entry.grid(row=row_index, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
            self.field_widgets[key] = entry
            entry_widgets.append(entry)
            
            # 繁殖検査（301）、フレッシュチェック（300）、妊娠鑑定マイナス（302）の特定フィールドに半角大文字変換機能を追加
            if event_number in [300, 301, 302] and key in ['treatment', 'uterine_findings', 'left_ovary_findings', 'right_ovary_findings', 'other']:
                # 入力時にリアルタイムで大文字に変換
                entry.bind('<KeyPress>', lambda e, widget=entry: self._convert_to_uppercase_on_input(e, widget))
                entry.bind('<KeyRelease>', lambda e, widget=entry: self._convert_to_uppercase(widget))
            
            row_index += 1
        
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
        return ["ホルスタイン", "ジャージー", "その他乳用種", "F1", "黒毛和種", "その他肉用種", "不明"]

    def _create_calving_form(self):
        """分娩イベント専用フォームを生成"""
        # 既存の分娩用状態を初期化
        self.calving_difficulty_var = tk.StringVar(value="")
        self.calving_child_count_var = tk.IntVar(value=1)
        self.calving_calf_vars = []
        
        container = ttk.Frame(self.form_frame)
        container.pack(fill=tk.BOTH, expand=True)
        _calving_label_minsize = 100
        
        # 難易度
        diff_frame = ttk.Frame(container)
        diff_frame.pack(fill=tk.X, pady=5)
        diff_frame.columnconfigure(0, minsize=_calving_label_minsize)
        diff_frame.columnconfigure(1, weight=0, minsize=self._entry_col_minsize)
        ttk.Label(diff_frame, text="分娩難易度:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        diff_options = self._get_calving_difficulty_options()
        diff_values = [f"{opt['code']}: {opt['label']}" for opt in diff_options]
        diff_combo = ttk.Combobox(
            diff_frame,
            textvariable=self.calving_difficulty_var,
            values=diff_values,
            state="readonly",
            width=self._entry_width,
            style="EventInput.TCombobox"
        )
        if diff_values:
            diff_combo.set(diff_values[0])
        diff_combo.grid(row=0, column=1, sticky=tk.EW, padx=(0, 5), pady=2)
        
        # 頭数
        count_frame = ttk.Frame(container)
        count_frame.pack(fill=tk.X, pady=5)
        count_frame.columnconfigure(0, minsize=_calving_label_minsize)
        count_frame.columnconfigure(1, weight=0, minsize=self._entry_col_minsize)
        ttk.Label(count_frame, text="子牛頭数:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        rb_frame = ttk.Frame(count_frame)
        rb_frame.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        for count, label in [(1, "単子"), (2, "双子"), (3, "三つ子")]:
            rb = ttk.Radiobutton(
                rb_frame,
                text=label,
                value=count,
                variable=self.calving_child_count_var,
                command=lambda c=count: self._update_calf_blocks(c)
            )
            rb.pack(side=tk.LEFT, padx=(0, 10))
        
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
        
        _block_label_minsize = 80
        for idx in range(count):
            block = ttk.LabelFrame(self.calving_block_container, text=f"子牛{idx + 1}", padding=5)
            block.pack(fill=tk.X, pady=5)
            block.columnconfigure(0, minsize=_block_label_minsize)
            block.columnconfigure(1, weight=0, minsize=self._entry_col_minsize)
            
            # 品種
            breed_var = tk.StringVar(value="")
            ttk.Label(block, text="品種:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=2)
            breed_combo = ttk.Combobox(block, textvariable=breed_var, values=breeds, state="readonly", width=self._entry_width, style="EventInput.TCombobox")
            breed_combo.grid(row=0, column=1, sticky=tk.EW, padx=(0, 5), pady=2)
            
            # 性別
            sex_var = tk.StringVar(value="")
            ttk.Label(block, text="性別:").grid(row=1, column=0, sticky=tk.W, padx=(5, 8), pady=2)
            sex_frame = ttk.Frame(block)
            sex_frame.grid(row=1, column=1, sticky=tk.W, padx=(0, 5), pady=2)
            ttk.Radiobutton(sex_frame, text="オス", value="M", variable=sex_var).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Radiobutton(sex_frame, text="メス", value="F", variable=sex_var).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Radiobutton(sex_frame, text="不明", value="U", variable=sex_var).pack(side=tk.LEFT)
            
            # 死産
            stillborn_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(block, text="死産", variable=stillborn_var).grid(row=2, column=1, sticky=tk.W, padx=(0, 5), pady=2)
            
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

    def _validate_introduction_input(self) -> bool:
        """導入イベント入力の検証"""
        input_fields = self.selected_event.get('input_fields', [])
        
        for field in input_fields:
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            is_required = field.get('required', False)
            widget = self.field_widgets.get(key)
            
            if not widget:
                continue
            
            value = widget.get().strip()
            
            # 必須項目のチェック
            if is_required and not value:
                label = field.get('label', key)
                messagebox.showerror("エラー", f"{label}は必須項目です")
                widget.focus()
                return False
            
            # jpn10の形式チェック（10桁の数字）
            if key == 'jpn10' and value:
                if not value.isdigit() or len(value) != 10:
                    messagebox.showerror("エラー", "個体識別番号は10桁の数字で入力してください")
                    widget.focus()
                    return False
            
            # 登録日（YYYY-MM-DD）
            if key == 'registration_date' and value:
                if not normalize_date(value):
                    messagebox.showerror(
                        "エラー",
                        "登録日（繁殖指標の起点）は YYYY-MM-DD 形式で入力してください。",
                    )
                    widget.focus()
                    return False
            
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

    def _collect_introduction_json(self) -> Dict[str, Any]:
        """導入イベント（600）のjson_dataを構築"""
        json_data: Dict[str, Any] = {}
        
        # 既存のjson_dataに残しておきたい値を引き継ぐ
        base = {}
        if self._editing_event_json:
            base = self._editing_event_json.copy()
        json_data.update(base)
        
        # input_fieldsから入力値を取得
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
            
            # PENプルダウンの表示値からコードを取得
            if key == 'pen':
                if hasattr(self, '_pen_display_to_code') and value in self._pen_display_to_code:
                    value = self._pen_display_to_code[value]
                elif "：" in value:
                    value = value.split("：", 1)[0].strip()
                elif ":" in value:
                    value = value.split(":", 1)[0].strip()
            
            # 型変換
            if datatype == 'int':
                try:
                    json_data[key] = int(value)
                except ValueError:
                    pass  # 変換失敗時はスキップ
            elif datatype == 'float':
                try:
                    json_data[key] = float(value)
                except ValueError:
                    pass  # 変換失敗時はスキップ
            else:
                json_data[key] = value
        
        return json_data

    def _handle_calving_followups(self, calving_event_id: int, calving_date: str, json_data: Dict[str, Any]):
        """
        分娩保存後の派生処理（導入イベント自動生成など）
        
        パターン1: 農場にいる母牛が乳用種のメスを産出し、その子牛がそのまま農場に登録される場合
        - 分娩イベントで乳用種（ホルスタイン、ジャージー、その他乳用種）かつメス（かつ多子の場合ではその片方がオスではない）の場合、導入イベントが発動
        - 母牛は自動で登録される（dam_jpn10）
        - SIREも自動で登録される（受胎している授精もしくはETイベントのSIRE）
        - 導入月日＝生年月日が自動入力
        - 品種も自動入力
        - PENは100：Heiferに自動入力
        """
        calves = json_data.get("calves") or []
        if not calves:
            return
        
        dairy_breeds = {"ホルスタイン", "ジャージー", "その他乳用種"}
        
        # 多子の場合、片方がオスでない（つまりメス）を確認
        eligible: List[Dict[str, Any]] = []
        for idx, calf in enumerate(calves, start=1):
            if calf.get("stillborn"):
                continue
            # メス（F）のみ対象
            if str(calf.get("sex")).upper() != "F":
                continue
            # 乳用種のみ対象
            if calf.get("breed") not in dairy_breeds:
                continue
            eligible.append({"idx": idx, "data": calf})
        
        if not eligible:
            return
        
        mother = self.db.get_cow_by_auto_id(self.cow_auto_id) if self.cow_auto_id else None
        if not mother:
            return
        
        mother_id = mother.get("cow_id")
        mother_jpn10 = mother.get("jpn10")
        
        # 受胎している授精もしくはETイベントのSIREを取得
        # 分娩日より前のAI/ETイベントで、その後に妊娠鑑定プラスがあるものを探す
        events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
        conception_sire = None
        
        # 分娩日より前のイベントを日付順にソート（降順：新しい順）
        events_before_calving = [
            e for e in events
            if e.get('event_date') and e.get('event_date') < calving_date
        ]
        events_before_calving.sort(key=lambda x: (x.get('event_date', ''), x.get('id', 0)), reverse=True)
        
        # 妊娠鑑定プラスイベントを探す
        preg_event = None
        for event in events_before_calving:
            if event.get('event_number') in [RuleEngine.EVENT_PDP, RuleEngine.EVENT_PDP2, RuleEngine.EVENT_PAGP]:
                preg_event = event
                break
        
        # 妊娠鑑定プラスイベントの前のAI/ETイベントを探す
        if preg_event:
            preg_date = preg_event.get('event_date')
            for event in events_before_calving:
                if event.get('event_date') and event.get('event_date') <= preg_date:
                    if event.get('event_number') in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
                        event_json = event.get('json_data') or {}
                        if isinstance(event_json, str):
                            try:
                                event_json = json.loads(event_json)
                            except:
                                event_json = {}
                        conception_sire = event_json.get('sire')
                        if conception_sire:
                            break
        
        created = 0
        
        for item in eligible:
            idx = item["idx"]
            calf = item["data"]
            calf_breed = calf.get("breed")
            
            # JPN10入力ダイアログを表示
            jpn10 = self._show_jpn10_input_dialog(
                calf_breed=calf_breed,
                calf_index=idx,
                birth_date=calving_date
            )
            
            if not jpn10:
                # ユーザーがキャンセルした場合、この子牛はスキップ
                continue
            
            # JPN10からcow_idを自動生成（6-9桁目）
            cow_id = jpn10[5:9].zfill(4)
            
            # 既に同じJPN10の牛が存在するか確認
            existing_cow = self.db.get_cow_by_id(cow_id)
            if existing_cow and existing_cow.get('jpn10') == jpn10:
                # 既存の牛が存在する場合、その牛に対して導入イベントを作成
                new_cow_auto_id = existing_cow.get('auto_id')
            else:
                # 新しい牛を登録
                new_cow_data = {
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'brd': calf_breed,
                    'bthd': calving_date,  # 生年月日
                    'entr': calving_date,   # 導入月日＝生年月日
                    'lact': 0,
                    'clvd': None,
                    'rc': RuleEngine.RC_OPEN,  # 初期状態は空胎
                    'pen': '100',  # PENは100：Heiferに自動入力
                    'frm': mother.get('frm') if mother else None
                }
                new_cow_auto_id = self.db.insert_cow(new_cow_data)
            
            # 導入イベントのjson_dataを作成
            intro_json = {
                "dam_jpn10": mother_jpn10,  # 母牛のJPN10（自動登録）
                "sire": conception_sire,     # SIRE（自動登録：受胎している授精もしくはETイベントのSIRE）
                "birth_date": calving_date,   # 導入月日＝生年月日（自動入力）
                "breed": calf_breed,         # 品種（自動入力）
                "pen": "100",                # PENは100：Heiferに自動入力
                "source": "auto_from_calving",  # 自動生成フラグ
                "calf_index": idx,
                "calf_sex": calf.get("sex"),
                "calving_event_id": calving_event_id
            }
            
            # 新しい牛に対して導入イベントを作成
            intro_event = {
                "cow_auto_id": new_cow_auto_id,
                "event_number": RuleEngine.EVENT_IN,  # 導入イベント
                "event_date": calving_date,
                "json_data": intro_json,
                "note": f"自動生成: 子牛{idx}導入（{calf_breed}♀）"
            }
            intro_event_id = self.db.insert_event(intro_event)
            self.rule_engine.on_event_added(intro_event_id)
            created += 1
        
        if created:
            messagebox.showinfo("導入イベント", f"{created}件の導入イベントを自動生成しました。")
    
    def _show_jpn10_input_dialog(self, calf_breed: str, calf_index: int, birth_date: str) -> Optional[str]:
        """
        JPN10入力ダイアログを表示
        
        Args:
            calf_breed: 子牛の品種
            calf_index: 子牛のインデックス
            birth_date: 生年月日
        
        Returns:
            JPN10（10桁の文字列）、キャンセル時はNone
        """
        dialog = tk.Toplevel(self.window)
        dialog.title("子牛登録")
        dialog.geometry("400x200")
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = {'jpn10': None}
        
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 説明文
        info_text = f"子牛{calf_index}（{calf_breed}♀）の個体識別番号（10桁）を入力してください。"
        ttk.Label(
            main_frame,
            text=info_text,
            font=("", 10)
        ).pack(pady=(0, 10))
        
        # JPN10入力欄
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(pady=10)
        
        ttk.Label(input_frame, text="JPN10:").pack(side=tk.LEFT, padx=5)
        jpn10_entry = ttk.Entry(input_frame, width=15)
        jpn10_entry.pack(side=tk.LEFT, padx=5)
        jpn10_entry.focus_set()
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def on_ok():
            jpn10_value = jpn10_entry.get().strip()
            if not jpn10_value:
                messagebox.showwarning("警告", "JPN10を入力してください。")
                return
            if not jpn10_value.isdigit() or len(jpn10_value) != 10:
                messagebox.showwarning("警告", "JPN10は10桁の数字で入力してください。")
                return
            result['jpn10'] = jpn10_value
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # EnterキーでOK
        jpn10_entry.bind('<Return>', lambda e: on_ok())
        
        # ダイアログが閉じられるまで待機
        dialog.wait_window()
        
        return result.get('jpn10')
    
    def _create_introduction_form(self):
        """導入イベント（600）専用の入力フォームを作成"""
        input_fields = self.selected_event.get('input_fields', [])
        
        # デバッグログ
        logging.debug(f"導入イベント: selected_event={self.selected_event}")
        logging.debug(f"導入イベント: input_fields={input_fields}")
        
        if not input_fields:
            logging.warning(f"導入イベント: input_fieldsが空です。selected_event={self.selected_event}")
            ttk.Label(
                self.form_frame,
                text="入力項目はありません",
                foreground="gray"
            ).pack(pady=10)
            return
        
        entry_widgets = []  # Enterキーで移動する順序を保持
        
        for i, field in enumerate(input_fields):
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            label_text = field.get('label', key)
            is_required = field.get('required', False)
            
            # ラベル（必須項目には*を付ける）
            label_text_display = f"{label_text}:" + ("*" if is_required else "")
            ttk.Label(
                self.form_frame,
                text=label_text_display
            ).grid(row=i, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
            
            # 入力ウィジェット
            if key in ('brd', 'breed'):
                # 導入イベントの品種は導入専用ウィンドウと同じプルダウンにする
                breed_options = ["ホルスタイン", "ジャージー", "その他"]
                entry = ttk.Combobox(
                    self.form_frame,
                    values=breed_options,
                    state="readonly",
                    width=self._entry_width,
                    style="EventInput.TCombobox"
                )
                entry.grid(row=i, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
            elif key == 'pen':
                # 導入イベントの群はPEN設定をプルダウンにする
                pen_options = []
                self._pen_display_to_code = {}
                if self.pen_settings:
                    def _pen_sort_key(item):
                        code = item[0]
                        return int(code) if str(code).isdigit() else 9999
                    for code, name in sorted(self.pen_settings.items(), key=_pen_sort_key):
                        display = f"{code}：{name}" if name else str(code)
                        pen_options.append(display)
                        self._pen_display_to_code[display] = str(code)
                entry = ttk.Combobox(
                    self.form_frame,
                    values=pen_options,
                    state="readonly",
                    width=self._entry_width,
                    style="EventInput.TCombobox"
                )
                entry.grid(row=i, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
            else:
                entry = ttk.Entry(self.form_frame, width=self._entry_width, style="EventInput.TEntry")
                entry.grid(row=i, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
            self.field_widgets[key] = entry
            entry_widgets.append(entry)
            
            # jpn10入力時に拡大4桁IDを自動取得する処理を追加
            if key == 'jpn10':
                entry.bind('<KeyRelease>', self._on_jpn10_changed)
                entry.bind('<FocusOut>', self._on_jpn10_focus_out)
            
            # calving_dateまたはdam_jpn10入力時に父（SIRE）を自動入力
            if key in ['calving_date', 'dam_jpn10']:
                entry.bind('<FocusOut>', self._on_calving_or_dam_changed)
            
            # 説明文がある場合は表示
            description = field.get('description')
            if description:
                ttk.Label(
                    self.form_frame,
                    text=description,
                    font=("", 8),
                    foreground="gray"
                ).grid(row=i, column=2, sticky=tk.W, padx=(0, 5), pady=5)
        
        # Enterキーで次のフィールドに移動する機能を追加
        for i, entry in enumerate(entry_widgets):
            if i < len(entry_widgets) - 1:
                # 最後のフィールド以外は次のフィールドに移動
                entry.bind('<Return>', lambda e, next_widget=entry_widgets[i + 1]: self._move_to_next_field(e, next_widget))
            else:
                # 最後のフィールドはメモ欄に移動
                entry.bind('<Return>', lambda e: self._move_to_note_field(e))
        
        # フォーム作成後に、既存データから自動入力（編集モードの場合）
        if self.event_id:
            self._auto_fill_sire_from_calving()
    
    def _on_jpn10_changed(self, event):
        """jpn10入力変更時の処理（拡大4桁IDを自動取得）"""
        jpn10_widget = event.widget
        jpn10_value = jpn10_widget.get().strip()
        
        # 10桁の数字が入力された場合、拡大4桁IDを自動取得
        if jpn10_value.isdigit() and len(jpn10_value) == 10:
            # jpn10の6-9桁目を拡大4桁IDとして使用
            cow_id = jpn10_value[5:9].zfill(4)
            
            # 牛ID入力欄が存在する場合は自動入力
            if hasattr(self, 'cow_id_entry'):
                current_cow_id = self.cow_id_entry.get().strip()
                if not current_cow_id:
                    self.cow_id_entry.delete(0, tk.END)
                    self.cow_id_entry.insert(0, cow_id)
    
    def _on_jpn10_focus_out(self, event):
        """jpn10入力欄からフォーカスが外れた時の処理"""
        self._on_jpn10_changed(event)
    
    def _on_calving_or_dam_changed(self, event):
        """calving_dateまたはdam_jpn10入力変更時の処理（父（SIRE）を自動入力）"""
        # 分娩イベントに紐づいた導入イベントかどうかを確認
        calving_date_widget = self.field_widgets.get('calving_date')
        dam_jpn10_widget = self.field_widgets.get('dam_jpn10')
        sire_widget = self.field_widgets.get('sire')
        
        if not calving_date_widget or not dam_jpn10_widget or not sire_widget:
            return
        
        calving_date = calving_date_widget.get().strip()
        dam_jpn10 = dam_jpn10_widget.get().strip()
        current_sire = sire_widget.get().strip()
        
        # 既にSIREが入力されている場合は自動入力しない
        if current_sire:
            return
        
        # calving_dateとdam_jpn10の両方が入力されている場合のみ自動入力
        if not calving_date or not dam_jpn10:
            return
        
        # 母牛を検索
        all_cows = self.db.get_all_cows()
        dam_cow = None
        for cow in all_cows:
            if cow.get('jpn10') == dam_jpn10:
                dam_cow = cow
                break
        
        if not dam_cow:
            return
        
        dam_auto_id = dam_cow.get('auto_id')
        if not dam_auto_id:
            return
        
        # 母牛のイベント履歴を取得
        dam_events = self.db.get_events_by_cow(dam_auto_id, include_deleted=False)
        
        # 分娩日を取得
        try:
            calving_date_obj = datetime.strptime(calving_date, '%Y-%m-%d')
        except (ValueError, TypeError):
            # 日付形式が不正な場合は処理しない
            return
        
        # 分娩日より前のAI/ETイベントを取得
        ai_et_events = [
            e for e in dam_events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
        ]
        
        if not ai_et_events:
            return
        
        # 分娩日より前のAI/ETイベントをフィルタリング
        before_calving_ai_et = [
            e for e in ai_et_events
            if e.get('event_date', '') < calving_date
        ]
        
        if not before_calving_ai_et:
            return
        
        # 妊娠鑑定プラスイベントを取得
        preg_events = [
            e for e in dam_events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
            and e.get('event_date')
            and e.get('event_date', '') < calving_date
        ]
        
        # 受胎したAI/ETイベントを特定
        conception_ai_event = None
        
        if preg_events:
            # 最初の妊娠イベントを取得
            sorted_preg = sorted(
                preg_events,
                key=lambda e: (e.get('event_date', ''), e.get('id', 0))
            )
            first_preg = sorted_preg[0]
            first_preg_date = first_preg.get('event_date')
            
            # 最初の妊娠イベントより前のAI/ETイベントを取得
            ai_et_before_preg = [
                e for e in before_calving_ai_et
                if e.get('event_date', '') <= first_preg_date
            ]
            
            if ai_et_before_preg:
                # 最新のAI/ETイベントを取得
                conception_ai_event = sorted(
                    ai_et_before_preg,
                    key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                    reverse=True
                )[0]
        else:
            # 妊娠イベントがない場合、分娩日より前の最新のAI/ETイベントを使用
            conception_ai_event = sorted(
                before_calving_ai_et,
                key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                reverse=True
            )[0]
        
        if conception_ai_event:
            # AI/ETイベントのjson_dataからSIREを取得
            ai_json_data = conception_ai_event.get('json_data') or {}
            if isinstance(ai_json_data, str):
                try:
                    ai_json_data = json.loads(ai_json_data)
                except:
                    ai_json_data = {}
            
            sire_value = ai_json_data.get('sire')
            if sire_value:
                # SIREを自動入力（大文字に変換）
                sire_widget.delete(0, tk.END)
                sire_widget.insert(0, str(sire_value).upper())
    
    def _auto_fill_sire_from_calving(self):
        """導入イベントのフォーム作成時に、既存データから父（SIRE）を自動入力"""
        if self.selected_event_number != 600:
            return
        
        # 既存のイベントデータを取得（編集モードの場合のみ）
        if self.event_id:
            event = self.db.get_event_by_id(self.event_id)
            if not event:
                return
            
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # calving_dateとdam_jpn10を取得
            calving_date = json_data.get('calving_date', '').strip()
            dam_jpn10 = json_data.get('dam_jpn10', '').strip()
            
            # 既にSIREが入力されている場合は自動入力しない
            current_sire = json_data.get('sire', '').strip()
            if current_sire:
                return
            
            # calving_dateとdam_jpn10の両方が存在する場合のみ自動入力
            if calving_date and dam_jpn10:
                # 入力欄に値を設定（まだ設定されていない場合）
                calving_date_widget = self.field_widgets.get('calving_date')
                dam_jpn10_widget = self.field_widgets.get('dam_jpn10')
                
                if calving_date_widget and not calving_date_widget.get().strip():
                    calving_date_widget.insert(0, calving_date)
                if dam_jpn10_widget and not dam_jpn10_widget.get().strip():
                    dam_jpn10_widget.insert(0, dam_jpn10)
                
                # 自動入力ロジックを実行
                self._on_calving_or_dam_changed(None)
    
    def _create_ai_et_form(self):
        """AI/ETイベント専用の入力フォームを作成（Comboboxのみ）"""
        row = 0
        
        # 1. SIRE（文字列入力）
        ttk.Label(self.form_frame, text="SIRE:").grid(row=row, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        sire_entry = ttk.Entry(self.form_frame, width=self._entry_width, style="EventInput.TEntry")
        sire_entry.grid(row=row, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.field_widgets['sire'] = sire_entry
        
        # SIRE入力時に大文字に変換
        def on_sire_key_release(event):
            """SIRE入力時に大文字に変換"""
            current_value = sire_entry.get()
            cursor_pos = sire_entry.index(tk.INSERT)
            # 大文字に変換
            upper_value = current_value.upper()
            if current_value != upper_value:
                sire_entry.delete(0, tk.END)
                sire_entry.insert(0, upper_value)
                # カーソル位置を復元（大文字変換で位置が変わらないように）
                try:
                    sire_entry.icursor(min(cursor_pos, len(upper_value)))
                except:
                    pass
        
        sire_entry.bind('<KeyRelease>', on_sire_key_release)
        
        # Enterキーで授精師コード入力欄に移動
        sire_entry.bind('<Return>', lambda e: self.field_widgets.get('technician_code', sire_entry).focus())
        
        row += 1
        
        # SIRE候補リストを表示（SIRE入力欄の下に配置）
        self._create_sire_candidates(sire_entry, row, 0)
        
        row += 1
        
        # 2. 授精師コード（Comboboxのみ）
        ttk.Label(self.form_frame, text="授精師コード:").grid(row=row, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        
        # Comboboxのvaluesを生成
        technician_values = [
            f"{code}：{name}"
            for code, name in sorted(self.technicians.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999)
        ]
        
        technician_combo = ttk.Combobox(
            self.form_frame,
            values=technician_values,
            width=self._entry_width,
            state='normal',  # editable にしてキーボード入力可能に
            style="EventInput.TCombobox"
        )
        technician_combo.grid(row=row, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.field_widgets['technician_code'] = technician_combo
        # キー入力で候補を絞り込み
        technician_combo.bind('<KeyRelease>', lambda e: self._on_technician_key(e))
        # Enter/Tabキーで確定して次のフィールドに移動
        technician_combo.bind('<Return>', lambda e: self._on_technician_enter(e))
        technician_combo.bind('<Tab>', lambda e: self._on_technician_tab(e))
        row += 1
        
        # 3. 授精種類コード（Comboboxのみ）
        ttk.Label(self.form_frame, text="授精種類コード:").grid(row=row, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        
        # Comboboxのvaluesを生成
        insemination_type_values = [
            f"{code}：{name}"
            for code, name in sorted(self.insemination_types.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999)
        ]
        
        insemination_type_combo = ttk.Combobox(
            self.form_frame,
            values=insemination_type_values,
            width=self._entry_width,
            state='normal',  # editable にしてキーボード入力可能に
            style="EventInput.TCombobox"
        )
        insemination_type_combo.grid(row=row, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.field_widgets['insemination_type_code'] = insemination_type_combo
        # キー入力で候補を絞り込み
        insemination_type_combo.bind('<KeyRelease>', lambda e: self._on_insemination_type_key(e))
        # Enter/Tabキーで確定して次のフィールドに移動
        insemination_type_combo.bind('<Return>', lambda e: self._on_insemination_type_enter(e))
        insemination_type_combo.bind('<Tab>', lambda e: self._on_insemination_type_tab(e))
        # 牛ID・日付に基づき授精種類を自動表示（個体カード起動時も同様に動作）
        self.window.after_idle(self._try_auto_fill_insemination_type)
        # フォーム描画完了後に再実行（親ウィンドウの描画待ちが必要な場合に備える）
        if self.cow_auto_id is not None:
            self.window.after(150, self._try_auto_fill_insemination_type)
    
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
        Entry、Text、またはBooleanVarウィジェットに安全に値を設定する
        
        Args:
            w: ウィジェット（tk.Entry, ttk.Entry, tk.Text, tk.BooleanVar など）
            value: 設定する値（None の場合は空文字列、BooleanVarの場合はbool）
        """
        # BooleanVar の判定
        if isinstance(w, tk.BooleanVar):
            try:
                if value is None:
                    w.set(False)
                elif isinstance(value, bool):
                    w.set(value)
                elif isinstance(value, str):
                    # 文字列の場合は "true", "True", "1" などを True に変換
                    w.set(value.lower() in ('true', '1', 'yes', 'on'))
                else:
                    w.set(bool(value))
                return
            except Exception:
                pass
        
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
            self._close()
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
            self.event_combo.set(event_name)
        
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
        if not isinstance(json_data, dict):
            json_data = {}
        if event_number == RuleEngine.EVENT_BLV:
            if not str(json_data.get("blv_result") or "").strip():
                bp = json_data.get("blv_positive")
                if bp is True:
                    json_data = {**json_data, "blv_result": "陽性"}
                elif bp is False:
                    json_data = {**json_data, "blv_result": "陰性"}
                else:
                    json_data = {**json_data, "blv_result": "未検査"}
        if event_number == RuleEngine.EVENT_IN and self.cow_auto_id:
            if not str(json_data.get('tag') or '').strip():
                tag_iv = self.db.get_item_value(self.cow_auto_id, 'TAG')
                if tag_iv:
                    json_data = {**json_data, 'tag': tag_iv}
            if not str(json_data.get('registration_date') or '').strip():
                cow_row = self.db.get_cow_by_auto_id(self.cow_auto_id)
                if cow_row and cow_row.get('entr'):
                    json_data = {**json_data, 'registration_date': cow_row.get('entr')}
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
                # SIREの値を大文字に変換
                if sire_value:
                    sire_value = str(sire_value).upper()
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
            
            # BooleanVar（チェックボックス）の場合
            if isinstance(widget, tk.BooleanVar):
                widget.set(bool(value))
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
                elif key == 'pen':
                    name = self.pen_settings.get(str(value), '')
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
                
                # イベント名の短縮処理（個体カードと統一）
                if event_number == RuleEngine.EVENT_PDN:
                    # 妊娠鑑定マイナス → 妊鑑－
                    display_name = "妊鑑－"
                elif event_number == RuleEngine.EVENT_PDP or event_number == RuleEngine.EVENT_PDP2:
                    # 妊娠鑑定プラス → 妊鑑＋
                    display_name = "妊鑑＋"
                elif event_number == 300:
                    # フレッシュチェック → フレチェック
                    if "フレッシュチェック" in display_name:
                        display_name = "フレチェック"
            else:
                display_name = "イベント不明"
            
            # 簡易情報（json_dataの主要項目を表示）
            json_data = event.get('json_data') or {}
            info_parts = []
            
            # AI/ETイベント（200/201）の特別表示
            if event_number in [200, 201]:
                # json_dataからoutcomeを取得（P/O/R/N）
                outcome = json_data.get('outcome')
                
                # 授精師・授精種類の辞書を取得（既にロード済みのself.technicians / self.insemination_typesを使用）
                # 共通関数を使用して表示文字列を生成
                # technicians_dict / insemination_types_dict は必ず渡す（空辞書でも可）
                detail_text = format_insemination_event(
                    json_data,
                    self.technicians,          # code -> name
                    self.insemination_types,   # code -> name
                    outcome                    # 受胎ステータス（P/O/R/N）
                )
                if detail_text:
                    info = detail_text
                else:
                    info = ""
            elif event_number == RuleEngine.EVENT_CALV:
                calv_def = self.event_dictionary.get(str(RuleEngine.EVENT_CALV), {})
                diff_labels = calv_def.get("calving_difficulty", {})
                detail_text = format_calving_event(json_data, diff_labels)
                if detail_text:
                    info = detail_text
                else:
                    info = ""
            # 繁殖検査（301）、妊娠鑑定マイナス（302）、フレッシュチェック（300）の特別表示
            elif event_number in [300, 301, 302]:
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
                
                # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
                right_ovary = (json_data.get('right_ovary_findings') or 
                              json_data.get('rightovary_findings') or 
                              json_data.get('rightovary_finding') or 
                              json_data.get('right_ovary', '') or
                              json_data.get('rightovary', ''))
                if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
                    valid_parts.append(f"右{str(right_ovary).strip()}")
                
                # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
                left_ovary = (json_data.get('left_ovary_findings') or 
                             json_data.get('leftovary_findings') or 
                             json_data.get('leftovary_finding') or 
                             json_data.get('left_ovary', '') or
                             json_data.get('leftovary', ''))
                if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
                    valid_parts.append(f"左{str(left_ovary).strip()}")
                
                # その他（新しいキー名と古いキー名の両方をサポート）
                other = (json_data.get('other') or 
                        json_data.get('other_info') or 
                        json_data.get('other_findings', ''))
                if other and str(other).strip() and str(other).strip() != '-':
                    valid_parts.append(str(other).strip())
                
                # 有効な項目のみを「  」（2つのスペース）で区切って表示（個体カードと同じ仕様）
                info = "  ".join(valid_parts) if valid_parts else ""
            else:
                # 通常のイベント
                if 'sire' in json_data:
                    info_parts.append(f"SIRE:{json_data['sire']}")
                if 'milk_yield' in json_data:
                    info_parts.append(f"乳量:{json_data['milk_yield']}")
                if 'to_pen' in json_data:
                    info_parts.append(f"→{json_data['to_pen']}")
                
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
        # 導入イベント（600）の場合は、IDが存在しない個体でもOK（新規個体が対象）
        is_introduction_event = (self.selected_event_number == 600)
        
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
                # 導入イベントの場合は、IDが存在しない個体でもOK（エラーを出さない）
                if is_introduction_event:
                    # 導入イベントの場合、IDが存在しない個体でもOK
                    # cow_auto_idはNoneのまま（後で_on_okで処理）
                    pass
                else:
                    # その他のイベントの場合はエラー
                    messagebox.showerror("エラー", f"牛が見つかりません: {cow_id_str}")
                    self.cow_id_entry.focus()
                    return False
            else:
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
        
        # 導入イベント（600）は専用の検証を行う
        if self.selected_event_number == 600:
            return self._validate_introduction_input()
        
        # 入力フィールドの検証
        input_fields = self.selected_event.get('input_fields', [])
        for field in input_fields:
            key = field.get('key')
            datatype = field.get('datatype', 'str')
            widget = self.field_widgets.get(key)
            
            if widget:
                # BooleanVar（チェックボックス）の場合は特別な処理
                if isinstance(widget, tk.BooleanVar):
                    # bool型は検証不要（常に有効な値）
                    continue
                
                # EntryやTextウィジェットの場合
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
    
    @staticmethod
    def _parse_date_safe(date_str: Optional[str]) -> Optional[datetime]:
        if not date_str:
            return None
        try:
            return datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
        except (ValueError, TypeError):
            return None
    
    def _get_events_for_warning_check(self, cow_auto_id: int) -> List[Dict[str, Any]]:
        """警告判定用にイベント一覧を取得（編集中イベントは除外）。"""
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        result: List[Dict[str, Any]] = []
        for event in events:
            event_id = event.get("id")
            if self.event_id is not None and event_id == self.event_id:
                continue
            if not event.get("event_date"):
                continue
            result.append(event)
        result.sort(key=lambda e: (e.get("event_date", ""), e.get("id", 0)))
        return result
    
    def _build_suspicious_input_warnings(
        self,
        cow_auto_id: int,
        event_number: int,
        event_date: str,
        json_data: Dict[str, Any],
    ) -> List[str]:
        """入力内容の注意喚起メッセージを生成（保存は可能）。"""
        warnings: List[str] = []
        target_dt = self._parse_date_safe(event_date)
        if target_dt is None:
            return warnings

        # 0) 未来日付チェック
        if target_dt.date() > datetime.now().date():
            warnings.append(f"イベント日（{event_date[:10]}）が本日より先の日付です。")

        events = self._get_events_for_warning_check(cow_auto_id)
        events_on_or_before = [
            e for e in events
            if (e.get("event_date") or "")[:10] <= event_date
        ]
        
        ai_et_numbers = {RuleEngine.EVENT_AI, RuleEngine.EVENT_ET}
        preg_pos_numbers = {RuleEngine.EVENT_PDP, RuleEngine.EVENT_PDP2, RuleEngine.EVENT_PAGP}
        preg_check_numbers = preg_pos_numbers | {RuleEngine.EVENT_PDN, RuleEngine.EVENT_PAGN}
        
        latest_ai_et = None
        for e in reversed(events_on_or_before):
            if e.get("event_number") in ai_et_numbers:
                latest_ai_et = e
                break
        
        # 1) 授精後25日以内の妊娠鑑定（+/-）
        if event_number in preg_check_numbers and latest_ai_et:
            ai_dt = self._parse_date_safe(latest_ai_et.get("event_date"))
            if ai_dt is not None:
                days_after_ai = (target_dt - ai_dt).days
                if 0 <= days_after_ai < 25:
                    warnings.append(
                        f"授精後{days_after_ai}日で妊娠鑑定（+/-）を入力しようとしています。"
                        "通常、授精後25日以内の妊娠鑑定は困難です。"
                    )
        
        # 4) 授精イベントなしで妊娠鑑定（+/-）
        if event_number in preg_check_numbers and latest_ai_et is None:
            if event_number in preg_pos_numbers:
                warnings.append(
                    "授精イベントがない状態で妊娠鑑定プラスを入力しようとしています。"
                    "受胎日を指定するか、不明として扱うか確認してください（不明の場合、分娩予定日も不明になります）。"
                )
            else:
                warnings.append(
                    "授精イベントがない状態で妊娠鑑定マイナスを入力しようとしています。"
                )
        
        # RuleEngineで当日直前の状態を取得（受胎日・妊娠状態判定に利用）
        state_before = None
        try:
            day_before = (target_dt - timedelta(days=1)).strftime("%Y-%m-%d")
            state_before = self.rule_engine.apply_events_until_date(cow_auto_id, day_before)
        except Exception:
            state_before = None
        
        # 2) 受胎日から270日以内の分娩
        if event_number == RuleEngine.EVENT_CALV and state_before:
            conception_date = state_before.get("conception_date")
            conc_dt = self._parse_date_safe(conception_date)
            if conc_dt is not None:
                gest_days = (target_dt - conc_dt).days
                if 0 <= gest_days < 270:
                    warnings.append(
                        f"受胎日から{gest_days}日で分娩を入力しようとしています。"
                        "通常の妊娠期間より短い可能性があります。"
                    )
        
        # 3) 受胎していない状態で分娩
        if event_number == RuleEngine.EVENT_CALV and state_before:
            rc_before = state_before.get("rc")
            if rc_before != RuleEngine.RC_PREGNANT:
                warnings.append(
                    "受胎中（妊娠中）になっていない状態で分娩を入力しようとしています。"
                    "妊娠鑑定プラス未入力などの可能性があります。"
                )
        
        # 5) 月齢11か月以内で授精
        if event_number in ai_et_numbers:
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                bthd_dt = self._parse_date_safe(cow.get("bthd"))
                if bthd_dt is not None and target_dt >= bthd_dt:
                    age_months = (target_dt.year - bthd_dt.year) * 12 + (target_dt.month - bthd_dt.month)
                    if target_dt.day < bthd_dt.day:
                        age_months -= 1
                    if age_months <= 11:
                        warnings.append(
                            f"授精時の月齢が{age_months}か月です。個体取り違いなどがないか確認してください。"
                        )
        
        # 6) 分娩後30日以内の授精
        if event_number in ai_et_numbers:
            latest_calv_dt = None
            for e in reversed(events_on_or_before):
                if e.get("event_number") == RuleEngine.EVENT_CALV:
                    latest_calv_dt = self._parse_date_safe(e.get("event_date"))
                    if latest_calv_dt is not None:
                        break
            if latest_calv_dt is None:
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                if cow:
                    clvd_dt = self._parse_date_safe(cow.get("clvd"))
                    if clvd_dt is not None and clvd_dt <= target_dt:
                        latest_calv_dt = clvd_dt
            if latest_calv_dt is not None:
                days_post_calving = (target_dt - latest_calv_dt).days
                if 0 <= days_post_calving < 30:
                    warnings.append(
                        f"分娩後{days_post_calving}日で授精を入力しようとしています。"
                        "分娩後30日以内の授精の可能性があります。"
                    )
        
        return warnings
    
    def _confirm_suspicious_input_warnings(
        self,
        cow_auto_id: int,
        event_number: int,
        event_date: str,
        json_data: Dict[str, Any],
    ) -> bool:
        warnings = self._build_suspicious_input_warnings(
            cow_auto_id=cow_auto_id,
            event_number=event_number,
            event_date=event_date,
            json_data=json_data,
        )
        if not warnings:
            return True
        
        lines = ["以下の内容は入力ミスの可能性があります。"] + [f"{idx}. {msg}" for idx, msg in enumerate(warnings, start=1)]
        lines.append("")
        lines.append("このまま保存しますか？")
        return messagebox.askyesno("入力内容の確認", "\n".join(lines))
    
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
        elif self.selected_event_number == 600:
            # 導入イベント（600）の場合は専用処理
            json_data = self._collect_introduction_json()
        else:
            json_data = {}
            
            # 1. input_fields から動的に処理
            input_fields = self.selected_event.get('input_fields', [])
            
            # 妊娠鑑定プラスイベント（303, 304, 307）の場合、双子・♀判定チェックボックスが存在しない場合は追加
            if self.selected_event_number in [303, 304, 307]:  # PDP, PDP2, PAGP
                has_twin_field = any(field.get('key') == 'twin' for field in input_fields)
                if not has_twin_field:
                    input_fields.append({
                        'key': 'twin',
                        'label': '双子',
                        'datatype': 'bool'
                    })
                has_female_judgment_field = any(field.get('key') == 'female_judgment' for field in input_fields)
                if not has_female_judgment_field:
                    input_fields.append({
                        'key': 'female_judgment',
                        'label': '♀',
                        'datatype': 'bool'
                    })
                has_male_judgment_field = any(field.get('key') == 'male_judgment' for field in input_fields)
                if not has_male_judgment_field:
                    input_fields.append({
                        'key': 'male_judgment',
                        'label': '♂',
                        'datatype': 'bool'
                    })
            
            for field in input_fields:
                key = field.get('key')
                datatype = field.get('datatype', 'str')
                widget = self.field_widgets.get(key)
                
                if not widget:
                    continue
                
                # BooleanVar（チェックボックス）の場合
                if isinstance(widget, tk.BooleanVar):
                    json_data[key] = widget.get()
                    continue
                
                value = widget.get().strip()
                if not value:
                    continue
                
                # Combobox の場合、「コード：名称」形式からコード部分のみを抽出
                if isinstance(widget, ttk.Combobox):
                    field_choices = field.get("choices")
                    if field_choices and isinstance(field_choices, list) and len(field_choices) > 0:
                        json_data[key] = value
                        continue
                    if key == 'treatment' and "：" in value:
                        # 処置は「1：WPG」形式から「WPG」のみを抽出
                        parts = value.split("：", 1)
                        if len(parts) == 2:
                            json_data[key] = parts[1].strip()
                        else:
                            json_data[key] = value
                    elif "：" in value:
                        # その他のComboboxはコード部分のみを抽出
                        code = value.split("：", 1)[0]
                        json_data[key] = code
                    else:
                        json_data[key] = value
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
                        # SIREの値を大文字に変換して保存
                        json_data['sire'] = sire_value.upper()
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
        
        # 導入イベント: TAG を item_value 用にウィジェットから取得（正規化で空欄が落ちる前に）
        intro_tag_for_item = ""
        if self.selected_event_number == RuleEngine.EVENT_IN:
            tw = self.field_widgets.get('tag')
            intro_tag_for_item = tw.get().strip() if tw else ""
        
        # 導入イベント: 登録日（繁殖指標の起点）。未入力時はイベント日付と同じ
        entr_for_metrics: Optional[str] = None
        if self.selected_event_number == 600:
            entr_for_metrics = normalized_date
            reg_raw = (json_data.get('registration_date') or '').strip()
            if reg_raw:
                entr_for_metrics = normalize_date(reg_raw) or normalized_date
        
        # 3. json_data を正規化（必ず dict にする、空の場合は {}）
        # None や空文字列の値は削除
        json_data = {k: v for k, v in json_data.items() if v is not None and v != ""}
        
        # 空の場合は {} を保存（None ではない）
        if not json_data:
            json_data = {}
        
        # AI/ETイベント保存時はP/O/R/Nの判定はrule_engine.update_insemination_outcomes()で行う
        # ここでは何もしない（on_event_added/on_event_updatedで自動的にoutcomeが設定される）
        
        # AI/ETのとき新規SIREなら種別の登録ダイアログを表示
        if self.selected_event_number in (RuleEngine.EVENT_AI, RuleEngine.EVENT_ET):
            farm_path = self.settings_manager.farm_path if self.settings_manager else None
            if farm_path:
                sire = (json_data.get("sire") or json_data.get("sire_name") or "").strip()
                if sire:
                    from ui.sire_list_window import get_known_sire_names, show_sire_confirm_dialog
                    known = get_known_sire_names(self.db, farm_path)
                    if sire not in known:
                        show_sire_confirm_dialog(self.window, farm_path, sire)
        
        # デバッグログ
        logging.debug(f"Event saved: event_number={self.selected_event_number}, json_data={json_data}")
        
        try:
            editing_event_id_before = self.event_id
            saved_event_id: Optional[int] = None
            saved_event_number = self.selected_event_number
            
            if self.event_id:
                if self.cow_auto_id is not None:
                    if not self._confirm_suspicious_input_warnings(
                        cow_auto_id=self.cow_auto_id,
                        event_number=self.selected_event_number,
                        event_date=normalized_date,
                        json_data=json_data,
                    ):
                        return
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
                
                # 妊娠プラス/マイナスイベント更新時は、rule_engine.update_insemination_outcomes()で自動的にoutcomeが更新される
            else:
                # 新規追加（新規・編集で同じ json_data を使用）
                # 導入イベント（600）の場合、IDが存在しない個体でもOK（新規個体が対象）
                # cow_auto_idがNoneの場合は、入力項目から新規に牛を登録
                if self.cow_auto_id is None and self.selected_event_number == 600:
                    # 導入イベントの場合、IDが存在しない個体でもOK
                    # json_dataから入力項目を取得
                    jpn10_value = json_data.get('jpn10', '').strip()
                    
                    if not jpn10_value:
                        messagebox.showerror("エラー", "個体識別番号（10桁）は必須です")
                        return
                    
                    # jpn10から拡大4桁IDを取得
                    if jpn10_value.isdigit() and len(jpn10_value) == 10:
                        jpn10 = jpn10_value
                        cow_id = jpn10[5:9].zfill(4)  # 6-9桁目を使用
                    else:
                        messagebox.showerror("エラー", "個体識別番号は10桁の数字で入力してください")
                        return
                    
                    # 既に存在するか確認
                    resolved_auto_id = self._resolve_cow_id(cow_id)
                    if resolved_auto_id:
                        # 既に存在する場合はそのauto_idを使用
                        self.cow_auto_id = resolved_auto_id
                        logging.info(f"導入イベント: 既存個体を使用 cow_id={cow_id}, jpn10={jpn10}, auto_id={self.cow_auto_id}")
                    else:
                        # 新規に牛を登録
                        # 農場コードを取得（settings_managerから）
                        frm = None
                        if self.settings_manager:
                            settings = self.settings_manager.load()
                            frm = settings.get('farm_code')
                        
                        # 入力項目から牛データを構築
                        breed = json_data.get('breed')
                        birth_date = json_data.get('birth_date') or json_data.get('calving_date')  # calving_dateもbirth_dateとして使用可能
                        lactation = json_data.get('lactation')
                        reproduction_code = json_data.get('reproduction_code')
                        pen_value = json_data.get('pen')
                        
                        # 新規に牛を登録
                        cow_data = {
                            'cow_id': cow_id,
                            'jpn10': jpn10,
                            'brd': breed if breed else None,
                            'bthd': birth_date if birth_date else None,
                            'entr': entr_for_metrics or normalized_date,
                            'lact': int(lactation) if lactation is not None else None,
                            'clvd': None,
                            'rc': int(reproduction_code) if reproduction_code is not None else None,
                            'pen': pen_value if pen_value else None,
                            'frm': frm
                        }
                        try:
                            self.cow_auto_id = self.db.insert_cow(cow_data)
                            logging.info(f"導入イベント: 新規個体を登録しました cow_id={cow_id}, jpn10={jpn10}, auto_id={self.cow_auto_id}")
                        except Exception as e:
                            # 既に存在する場合は取得
                            logging.warning(f"導入イベント: 個体登録エラー（既存の可能性）: {e}")
                            # 再度検索を試みる
                            resolved_auto_id = self._resolve_cow_id(cow_id)
                            if resolved_auto_id:
                                self.cow_auto_id = resolved_auto_id
                            else:
                                messagebox.showerror("エラー", f"個体の登録に失敗しました: {e}")
                                return
                
                # cow_auto_idがまだNoneの場合はエラー
                if self.cow_auto_id is None:
                    messagebox.showerror("エラー", "牛IDを解決できませんでした")
                    return
                
                if not self._confirm_suspicious_input_warnings(
                    cow_auto_id=self.cow_auto_id,
                    event_number=self.selected_event_number,
                    event_date=normalized_date,
                    json_data=json_data,
                ):
                    return

                # 導入イベントで分娩月日が入力されている場合は分娩イベントを自動作成
                if self.selected_event_number == 600:
                    calving_date_raw = json_data.get('calving_date')
                    if calving_date_raw:
                        normalized_calving_date = normalize_date(str(calving_date_raw))
                        if normalized_calving_date:
                            # 既存の分娩イベントが同日で存在する場合は作成しない
                            events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
                            already_exists = any(
                                e.get('event_number') == RuleEngine.EVENT_CALV and
                                e.get('event_date') == normalized_calving_date
                                for e in events
                            )
                            if not already_exists:
                                calv_event = {
                                    'cow_auto_id': self.cow_auto_id,
                                    'event_number': RuleEngine.EVENT_CALV,
                                    'event_date': normalized_calving_date,
                                    'json_data': {'baseline_calving': True},
                                    'note': '導入イベントから自動作成'
                                }
                                calv_event_id = self.db.insert_event(calv_event)
                                self.rule_engine.on_event_added(calv_event_id)
                
                # 繁殖検診イベント（300,301,302,303,304）で同じ日付に既存がある場合は警告
                if self.selected_event_number in REPRO_CHECKUP_EVENT_NUMBERS:
                    events = self.db.get_events_by_cow(self.cow_auto_id, include_deleted=False)
                    date_str = (normalized_date or "")[:10]
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
                
                # 妊娠プラス/マイナスイベント保存時は、rule_engine.update_insemination_outcomes()で自動的にoutcomeが更新される
            
            # 導入イベント: タグを item_value（TAG）に同期、登録日（entr）を同期
            if saved_event_number == RuleEngine.EVENT_IN and self.cow_auto_id is not None:
                self.db.set_item_value(self.cow_auto_id, "TAG", intro_tag_for_item)
                if entr_for_metrics:
                    try:
                        self.db.update_cow(self.cow_auto_id, {'entr': entr_for_metrics})
                    except Exception as e:
                        logging.warning(f"導入イベント: entr 更新に失敗: {e}")
            
            try:
                from modules.activity_log import record_from_event
                record_from_event(
                    self.db,
                    cow_auto_id=self.cow_auto_id,
                    action="更新" if editing_event_id_before else "登録",
                    event_number=saved_event_number,
                    event_id=saved_event_id,
                    event_date=normalized_date or "",
                    event_dictionary=self.event_dictionary,
                )
            except Exception:
                pass
            
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
            ev_before = self.db.get_event_by_id(self.event_id)
            # 論理削除
            self.db.delete_event(self.event_id, soft_delete=True)
            self.rule_engine.on_event_deleted(self.event_id)
            if ev_before:
                try:
                    from modules.activity_log import record_from_event
                    record_from_event(
                        self.db,
                        cow_auto_id=self.cow_auto_id,
                        action="削除",
                        event_number=ev_before.get("event_number"),
                        event_id=self.event_id,
                        event_date=str(ev_before.get("event_date") or "")[:10],
                        event_dictionary=self.event_dictionary,
                    )
                except Exception:
                    pass
            
            # イベント履歴を更新
            if self.cow_auto_id:
                self._load_event_history()
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            # ウィンドウを閉じる
            self._close()
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self._close()

    def _close(self):
        """ウィンドウを安全に閉じてアクティブインスタンスをクリアする"""
        # アクティブインスタンスの解除
        try:
            if EventInputWindow._active_window is self:
                EventInputWindow._active_window = None
        except Exception:
            pass

        # 実際のウィンドウ破棄
        try:
            if hasattr(self, "window") and self.window is not None:
                # 既に破棄済みの場合の例外は無視
                self.window.destroy()
        except Exception:
            pass
    
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
            # BooleanVarの場合は対応するCheckbuttonを探す
            for key, widget in self.field_widgets.items():
                # _checkboxで終わるキーはスキップ（実際のウィジェットは別のキーで保存）
                if key.endswith('_checkbox'):
                    continue
                if isinstance(widget, tk.BooleanVar):
                    # 対応するCheckbuttonウィジェットを探す
                    checkbox_key = f"{key}_checkbox"
                    if checkbox_key in self.field_widgets:
                        checkbox = self.field_widgets[checkbox_key]
                        checkbox.focus()
                        return "break"
                elif hasattr(widget, 'focus'):
                    # EntryやTextなどのフォーカス可能なウィジェットの場合
                    widget.focus()
                    return "break"
            # フォーカス可能なウィジェットがない場合はメモ欄に移動
            self.note_entry.focus()
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
            # BooleanVarの場合は対応するCheckbuttonを探す
            for key, widget in self.field_widgets.items():
                # _checkboxで終わるキーはスキップ（実際のウィジェットは別のキーで保存）
                if key.endswith('_checkbox'):
                    continue
                if isinstance(widget, tk.BooleanVar):
                    # 対応するCheckbuttonウィジェットを探す
                    checkbox_key = f"{key}_checkbox"
                    if checkbox_key in self.field_widgets:
                        checkbox = self.field_widgets[checkbox_key]
                        checkbox.focus()
                        return "break"
                elif hasattr(widget, 'focus'):
                    # EntryやTextなどのフォーカス可能なウィジェットの場合
                    widget.focus()
                    return "break"
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
        self._close()
        
        # 編集用のイベント入力ウィンドウを開く
        from ui.event_input import EventInputWindow
        edit_window = EventInputWindow.open_or_focus(
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
        
        ev_before = self.db.get_event_by_id(event_id)
        
        # 確認ダイアログ
        result = messagebox.askyesno("確認", "このイベントを削除しますか？")
        if not result:
            return
        
        try:
            # 論理削除
            self.db.delete_event(event_id, soft_delete=True)
            self.rule_engine.on_event_deleted(event_id)
            if ev_before:
                try:
                    from modules.activity_log import record_from_event
                    record_from_event(
                        self.db,
                        cow_auto_id=self.cow_auto_id,
                        action="削除",
                        event_number=ev_before.get("event_number"),
                        event_id=event_id,
                        event_date=str(ev_before.get("event_date") or "")[:10],
                        event_dictionary=self.event_dictionary,
                    )
                except Exception:
                    pass
            
            # イベント履歴を更新
            self._load_event_history()
            
            # MainWindow に通知
            if self.on_saved:
                self.on_saved(self.cow_auto_id)
            
            messagebox.showinfo("完了", "イベントを削除しました")
            
        except Exception as e:
            messagebox.showerror("エラー", f"イベントの削除に失敗しました: {e}")
    
    def _get_recent_sire_candidates(self) -> List[tuple]:
        """
        直近30件のAIイベントからSIREを取得し、頻度順にソート
        
        Returns:
            (SIRE, 頻度) のタプルのリスト（最大6つ、頻度の高い順）
        """
        try:
            # 全イベントを取得（削除済みを除く）
            all_events = self.db.get_all_events(include_deleted=False)
            
            # AIイベント（event_number = 200）のみをフィルタリング
            ai_events = [
                e for e in all_events
                if e.get('event_number') == RuleEngine.EVENT_AI
            ]
            
            # 直近30件に制限（event_date DESC順で既にソートされている）
            recent_ai_events = ai_events[:30]
            
            # SIREの頻度を集計
            sire_counts: Dict[str, int] = {}
            for event in recent_ai_events:
                json_data = event.get('json_data', {})
                if isinstance(json_data, dict):
                    sire = json_data.get('sire')
                    if sire and str(sire).strip():
                        sire_upper = str(sire).strip().upper()
                        sire_counts[sire_upper] = sire_counts.get(sire_upper, 0) + 1
            
            # 頻度の高い順にソート（最大6つ）
            sorted_sires = sorted(sire_counts.items(), key=lambda x: x[1], reverse=True)[:6]
            
            return sorted_sires
            
        except Exception as e:
            logging.error(f"SIRE候補の取得エラー: {e}")
            return []
    
    def _create_sire_candidates(self, sire_entry: tk.Entry, row: int, column: int):
        """
        SIRE候補リストを作成
        
        Args:
            sire_entry: SIRE入力欄のEntryウィジェット
            row: グリッドの行番号
            column: グリッドの列番号
        """
        # 候補を取得
        candidates = self._get_recent_sire_candidates()
        
        if not candidates:
            # 候補がない場合は何も表示しない
            return
        
        # 候補フレームを作成（SIRE入力欄の下に配置）
        candidate_frame = ttk.Frame(self.form_frame)
        candidate_frame.grid(row=row, column=column, columnspan=2, sticky=tk.W, padx=(5, 5), pady=(0, 5))
        
        # 候補ボタンを配置（SIRE名のみ、幅は自動調整）
        for sire, count in candidates:
            # SIRE名の長さに合わせてボタンの幅を設定（最小幅は確保）
            btn_width = max(len(sire) + 2, 8)  # SIRE名の長さ + 余白、最小8文字分
            
            btn = ttk.Button(
                candidate_frame,
                text=sire,  # 頻度は表示しない
                width=btn_width,
                command=lambda s=sire, entry=sire_entry: self._on_sire_candidate_clicked(s, entry)
            )
            btn.pack(side=tk.LEFT, padx=2)
            
            # ダブルクリックでも入力できるようにする
            btn.bind('<Double-Button-1>', lambda e, s=sire, entry=sire_entry: self._on_sire_candidate_clicked(s, entry))
    
    def _on_sire_candidate_clicked(self, sire: str, entry: tk.Entry):
        """
        SIRE候補がクリックされた時の処理
        
        Args:
            sire: 選択されたSIRE
            entry: SIRE入力欄のEntryウィジェット
        """
        entry.delete(0, tk.END)
        entry.insert(0, sire)
        entry.focus()
    
    # _calculate_ai_result()と_update_ai_result_on_pregnancy_event()は削除
    # P/O/R/Nの判定はrule_engine.update_insemination_outcomes()で行い、outcomeとして保存される
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()

