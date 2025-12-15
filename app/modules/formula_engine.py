"""
FALCON2 - FormulaEngine
表示用・分析用の計算項目をオンデマンドで計算
設計書 第10章参照
"""

from typing import Dict, Any, Optional
from datetime import datetime, timedelta
from pathlib import Path
import json
import math
import logging

from db.db_handler import DBHandler


class FormulaEngine:
    """計算項目をオンデマンドで計算するエンジン"""
    
    def __init__(self, db_handler: DBHandler, item_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            item_dictionary_path: item_dictionary.json のパス（Noneの場合はデフォルトパスを使用）
        """
        self.db = db_handler
        self.item_dict_path = item_dictionary_path
        self.item_dictionary: Dict[str, Any] = {}
        self._load_item_dictionary()

    def reload_item_dictionary(self):
        """外部更新後に辞書を再読込する"""
        self._load_item_dictionary()
    
    def _load_item_dictionary(self):
        """item_dictionary.json を読み込む"""
        if self.item_dict_path is None:
            # デフォルトパス（農場フォルダ内のitem_dictionary.jsonを想定）
            # 実際の使用時は農場フォルダから読み込む必要がある
            default_path = Path("docs/item_dictionary.json")
            if default_path.exists():
                self.item_dict_path = default_path
        
        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                print(f"item_dictionary.json 読み込みエラー: {e}")
                self.item_dictionary = {}
        else:
            self.item_dictionary = {}
    
    def calculate(self, cow_auto_id: int) -> Dict[str, Any]:
        """
        牛の計算項目を計算
        
        【重要】副作用なし。DB更新は行わない。
        
        Args:
            cow_auto_id: 牛の auto_id
        
        Returns:
            計算項目の辞書
        """
        # 常に最新の辞書を反映
        self._load_item_dictionary()

        # cow データを取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return {}
        
        # イベント履歴を取得
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        result = {}
        
        # 1. source 定義による項目（event.field 形式）
        source_items = self._calculate_source_items(events)
        logging.debug(f"[FormulaEngine] source_items: {source_items}")
        result.update(source_items)
        
        # 2. formula 定義による calc 項目（DIM、DPAI、DUE も含む）
        formula_items = self._calculate_formula_items(cow, events, result, cow_auto_id)
        logging.debug(f"[FormulaEngine] formula_items: {formula_items}")
        result.update(formula_items)
        
        logging.info(f"[FormulaEngine] calculate result: {result}")
        return result
    
    def _calculate_dim(self, clvd: Optional[str]) -> Optional[int]:
        """
        DIM（分娩後日数）を計算
        
        Args:
            clvd: 最終分娩日（YYYY-MM-DD形式）
        
        Returns:
            DIM（日数）、clvdがNoneの場合はNone
        """
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
            today = datetime.now()
            dim = (today - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_dpai(self, last_ai_date: Optional[str]) -> Optional[int]:
        """
        DPAI（最終AI後日数）を計算
        
        Args:
            last_ai_date: 最終AI日（YYYY-MM-DD形式）
        
        Returns:
            DPAI（日数）、last_ai_dateがNoneの場合はNone
        """
        if not last_ai_date:
            return None
        
        try:
            ai_date = datetime.strptime(last_ai_date, '%Y-%m-%d')
            today = datetime.now()
            dpai = (today - ai_date).days
            return dpai if dpai >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_last_items(self, cow_auto_id: int) -> Dict[str, Any]:
        """
        乳検（601）由来の LAST_ 系項目を計算
        
        item_dictionary.json の source 定義（例: "601.milk_yield"）から
        汎用的に解決する
        
        Args:
            cow_auto_id: 牛の auto_id
        
        Returns:
            LAST_ 系項目の辞書
        """
        # 互換性のために source 処理へ委譲
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        return self._calculate_source_items(events)
    
    def _get_last_ai_date(self, events: list) -> Optional[str]:
        """
        イベント履歴から最新のAI/ET日を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            最新のAI/ET日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
        ]
        
        if not ai_et_events:
            return None
        
        # event_date でソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        return sorted_events[0].get('event_date')
    
    def _calculate_due_date(self, events: list, clvd: Optional[str]) -> Optional[str]:
        """
        分娩予定日を計算
        
        Args:
            events: イベントリスト
            clvd: 最終分娩日（現在の産次での分娩日）
        
        Returns:
            分娩予定日（YYYY-MM-DD形式）、計算できない場合はNone
        """
        # 最新の妊娠鑑定プラスイベントを取得
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
        ]
        
        if not preg_events:
            return None
        
        # 最新の妊娠鑑定プラスイベント
        latest_preg = sorted(
            preg_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        
        preg_date = latest_preg.get('event_date')
        
        # 最新のAI/ETイベントを取得（妊娠鑑定より前）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date', '') <= preg_date
        ]
        
        if not ai_et_events:
            return None
        
        latest_ai_et = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        
        ai_et_date = latest_ai_et.get('event_date')
        event_number = latest_ai_et.get('event_number')
        
        try:
            # ETの場合は7日前、AIの場合は当日が受胎日
            if event_number == 201:  # ET
                conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d') - timedelta(days=7)
            else:  # AI
                conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d')
            
            # 分娩予定日 = 受胎日 + 280日
            due_date = conception_date + timedelta(days=280)
            return due_date.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None

    def _calculate_source_items(self, events: list) -> Dict[str, Any]:
        """
        item_dictionary の source 定義 (event_number.field) を汎用的に解決
        """
        result: Dict[str, Any] = {}
        # event_date が存在するイベントのみをフィルタリング
        events_with_date = [e for e in events if e.get('event_date')]
        # イベントを event_number ごとに最新順へソート
        events_sorted = sorted(
            events_with_date,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        latest_by_number: Dict[int, Dict[str, Any]] = {}
        for ev in events_sorted:
            num = ev.get('event_number')
            if num is None:
                continue
            if num not in latest_by_number:
                latest_by_number[num] = ev

        for item_key, item_def in self.item_dictionary.items():
            source = item_def.get('source')
            if not source or '.' not in source:
                continue
            event_num_str, field_key = source.split('.', 1)
            try:
                event_num = int(event_num_str)
            except ValueError:
                continue
            latest_event = latest_by_number.get(event_num)
            if not latest_event:
                logging.debug(f"[FormulaEngine] {item_key}: イベント{event_num}が見つかりません")
                result[item_key] = None
                continue
            json_data = latest_event.get('json_data') or {}
            # json_data になければトップレベルも参照
            value = json_data.get(field_key)
            if value is None:
                value = latest_event.get(field_key)
            logging.debug(f"[FormulaEngine] {item_key}: source={source}, value={value}, json_data={json_data}")
            result[item_key] = value
        return result

    def latest_event_value(self, event_number: int, key: str) -> Optional[Any]:
        """
        指定されたイベント番号の最新レコードから指定キーの値を取得
        
        Args:
            event_number: イベント番号
            key: json_data 内のキー名
        
        Returns:
            値（存在しない場合は None）
        """
        # 現在の cow_auto_id を取得する必要があるため、
        # このメソッドは _calculate_formula_items 内で呼び出されることを想定
        # cow_auto_id は self._current_cow_auto_id に保存されている
        if not hasattr(self, '_current_cow_auto_id'):
            return None
        
        cow_auto_id = self._current_cow_auto_id
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 指定されたイベント番号のイベントをフィルタリング
        target_events = [
            e for e in events
            if e.get("event_number") == event_number and e.get("event_date")
        ]
        
        if not target_events:
            return None
        
        # event_date でソートして最新を取得
        latest = max(
            target_events,
            key=lambda e: e["event_date"],
            default=None
        )
        
        if not latest:
            return None
        
        # json_data から値を取得
        json_data = latest.get("json_data") or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        return json_data.get(key)

    def _calculate_formula_items(
        self,
        cow: Dict[str, Any],
        events: list,
        current: Dict[str, Any],
        cow_auto_id: int,
    ) -> Dict[str, Any]:
        """
        formula 定義を評価して値を算出
        """
        # latest_event_value で使用するために cow_auto_id を保存
        self._current_cow_auto_id = cow_auto_id
        
        results: Dict[str, Any] = {}
        safe_builtins = {
            "min": min,
            "max": max,
            "sum": sum,
            "len": len,
            "abs": abs,
            "round": round,
            "math": math,
        }
        for item_key, item_def in self.item_dictionary.items():
            formula = item_def.get("formula")
            if not formula:
                continue
            local_ctx = {
                "cow": cow,
                "events": events,
                "result": {**current, **results},  # 逐次計算された値を利用可能に
                "auto_id": cow_auto_id,
                "cow_id": cow.get("cow_id"),
                "latest_event_value": lambda en, k: self.latest_event_value(en, k),  # latest_event_value を利用可能に
                "dim": lambda clvd: self._calculate_dim(clvd),  # DIM 計算関数
                "dpai": lambda evs: self._calculate_dpai(self._get_last_ai_date(evs)),  # DPAI 計算関数
                "due_date": lambda evs, clvd: self._calculate_due_date(evs, clvd),  # DUE 計算関数
            }
            try:
                value = eval(formula, {"__builtins__": safe_builtins}, local_ctx)
                # デバッグログ（LAST_DENOVO_FA の場合のみ）
                if item_key == "LAST_DENOVO_FA":
                    logging.info(
                        f"[Formula] LAST_DENOVO_FA={value} type={type(value)}"
                    )
                results[item_key] = value
            except Exception as e:
                # 計算失敗時は None にして継続
                logging.warning(f"[Formula] 計算失敗: item_key={item_key}, formula={formula}, error={e}")
                results[item_key] = None
        
        # クリーンアップ
        if hasattr(self, '_current_cow_auto_id'):
            delattr(self, '_current_cow_auto_id')
        
        return results
    
    def get_item_value(self, cow_auto_id: int, item_key: str) -> Any:
        """
        特定の計算項目の値を取得
        
        Args:
            cow_auto_id: 牛の auto_id
            item_key: アイテムキー（例: "DIM", "LAST_MILK_YIELD"）
        
        Returns:
            計算値、存在しない場合はNone
        """
        calculated = self.calculate(cow_auto_id)
        return calculated.get(item_key)
    
    def set_item_dictionary_path(self, item_dictionary_path: Path):
        """
        item_dictionary.json のパスを設定（農場切替時など）
        
        Args:
            item_dictionary_path: item_dictionary.json のパス
        """
        self.item_dict_path = item_dictionary_path
        self._load_item_dictionary()

