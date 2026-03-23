"""
FALCON2 - Query Executor（クエリ実行モジュール）
正規化されたクエリからDB集計を実行
"""

import json
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime
from pathlib import Path
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine


class QueryExecutor:
    """正規化されたクエリからDB集計を実行"""
    
    def __init__(self, db_handler: DBHandler, formula_engine: FormulaEngine, 
                 rule_engine: RuleEngine, item_dictionary_path=None, 
                 event_dictionary_path=None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            item_dictionary_path: item_dictionary.json のパス（Noneの場合はデフォルト）
            event_dictionary_path: event_dictionary.json のパス（Noneの場合はデフォルト）
        """
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.item_dictionary_path = item_dictionary_path
        self.event_dictionary_path = event_dictionary_path
        
        # 辞書を読み込む
        self.item_dictionary: Dict[str, Any] = {}
        self.event_dictionary: Dict[str, Any] = {}
        self.event_key_to_number: Dict[str, int] = {}  # event_key (alias) -> event_number
        
        self._load_dictionaries()
    
    def _load_dictionaries(self):
        """item_dictionaryとevent_dictionaryを読み込む"""
        # item_dictionary.json
        if self.item_dictionary_path:
            item_path = Path(self.item_dictionary_path)
        else:
            # デフォルトパス（農場フォルダ内を想定、実際の使用時は適切なパスを指定）
            item_path = Path("config_default/item_dictionary.json")
        
        if item_path.exists():
            try:
                with open(item_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                logging.warning(f"item_dictionary.json読み込みエラー: {e}")
        
        # event_dictionary.json
        if self.event_dictionary_path:
            event_path = Path(self.event_dictionary_path)
        else:
            # デフォルトパス
            event_path = Path("config_default/event_dictionary.json")
        
        if event_path.exists():
            try:
                with open(event_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
                
                # event_key (alias) -> event_number のマッピングを作成
                for event_number_str, event_data in self.event_dictionary.items():
                    if isinstance(event_data, dict):
                        alias = event_data.get("alias")
                        if alias:
                            try:
                                event_number = int(event_number_str)
                                self.event_key_to_number[alias] = event_number
                                # 正規化辞書で使用されるキーもマッピング
                                # 例: "CALVING" -> 202 (CALV alias)
                                if alias == "CALV":
                                    self.event_key_to_number["CALVING"] = event_number
                                elif alias == "PDP":
                                    self.event_key_to_number["PREG_POS"] = event_number
                                elif alias == "PDN":
                                    self.event_key_to_number["PREG_NEG"] = event_number
                                elif alias == "MILK_TEST":
                                    self.event_key_to_number["MILK_TEST"] = event_number
                            except ValueError:
                                pass
            except Exception as e:
                logging.warning(f"event_dictionary.json読み込みエラー: {e}")
    
    def execute(self, normalized_query: Dict[str, Any]) -> Optional[str]:
        """
        正規化されたクエリを実行
        
        Args:
            normalized_query: 正規化されたクエリ辞書
                {
                    "item": item_key or None,
                    "event": event_key or None,
                    "term": term_key or None,
                    "period": {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"} or None,
                    "error": error_message or None
                }
        
        Returns:
            集計結果の文字列、未対応の場合はNone
        """
        # エラーチェック
        if normalized_query.get("error"):
            return None
        
        item = normalized_query.get("item")
        event = normalized_query.get("event")
        term = normalized_query.get("term")
        period = normalized_query.get("period")
        
        # item + AVG → 項目の平均値集計
        if item and term == "AVG":
            return self._execute_item_avg(item, period)
        
        # event + COUNT → イベント件数集計
        if event and term == "COUNT":
            return self._execute_event_count(event, period)
        
        # event + LIST → イベント該当個体抽出
        if event and term == "LIST":
            return self._execute_event_list(event, period)
        
        # SCATTER → 2項目指定必須（未実装、後で追加可能）
        if term == "SCATTER":
            return "未対応：散布図は2項目指定が必要です"
        
        # 未対応クエリ
        return "未対応：このクエリは処理できません"
    
    def _execute_item_avg(self, item_key: str, period: Optional[Dict[str, str]]) -> Optional[str]:
        """
        項目の平均値を集計
        
        Args:
            item_key: 項目キー（例: "DIM", "DIMFAI"）
            period: 期間辞書またはNone
        
        Returns:
            集計結果の文字列
        """
        try:
            # 全牛を取得
            cows = self.db.get_all_cows()
            
            # 期間フィルタリング
            if period:
                start_date = datetime.strptime(period["start"], "%Y-%m-%d")
                end_date = datetime.strptime(period["end"], "%Y-%m-%d")
                # 期間フィルタリングは後で実装（必要に応じて）
            
            values = []
            cow_count = 0
            
            for cow in cows:
                cow_auto_id = cow.get("auto_id")
                if not cow_auto_id:
                    continue
                
                # FormulaEngineで計算（item_keyを指定してその項目のみ計算）
                calculated = self.formula_engine.calculate(cow_auto_id, item_key=item_key)
                value = calculated.get(item_key)
                
                if value is not None:
                    try:
                        # 数値に変換可能な場合のみ集計対象
                        if isinstance(value, (int, float)):
                            values.append(float(value))
                            cow_count += 1
                    except (ValueError, TypeError):
                        pass
            
            if not values:
                return None
            
            # 平均値を計算
            avg_value = sum(values) / len(values)
            
            # 出力フォーマット
            item_display_name = self._get_item_display_name(item_key)
            unit = self._get_item_unit(item_key)
            
            result_lines = []
            result_lines.append(f"＜{item_display_name}＞：{avg_value:.2f}{unit}")
            result_lines.append(f"対象牛：{cow_count}頭")
            
            if period:
                result_lines.append(f"期間：{period['start']} ～ {period['end']}")
            else:
                today = datetime.now().strftime("%Y年%m月%d日")
                result_lines.append(f"基準日：{today}")
            
            return "\n".join(result_lines)
        
        except Exception as e:
            logging.error(f"項目平均値集計エラー: {e}")
            return None
    
    def _execute_event_count(self, event_key: str, period: Optional[Dict[str, str]]) -> Optional[str]:
        """
        イベント件数を集計
        
        Args:
            event_key: イベントキー（例: "AI", "CALVING"）
            period: 期間辞書またはNone
        
        Returns:
            集計結果の文字列
        """
        try:
            # イベントキーからイベント番号を取得
            event_number = self._get_event_number(event_key)
            if event_number is None:
                return None
            
            # イベントを取得
            if period:
                start_date = period["start"]
                end_date = period["end"]
                events = self.db.get_events_by_number_and_period(event_number, start_date, end_date)
            else:
                events = self.db.get_events_by_number(event_number)
            
            count = len(events)
            
            # 出力フォーマット
            event_display_name = self._get_event_display_name(event_key)
            
            result_lines = []
            result_lines.append(f"＜{event_display_name}件数＞：{count}件")
            
            # 対象牛数を計算（重複除去）
            cow_ids = set()
            for event in events:
                cow_auto_id = event.get("cow_auto_id")
                if cow_auto_id:
                    cow_ids.add(cow_auto_id)
            
            result_lines.append(f"対象牛：{len(cow_ids)}頭")
            
            if period:
                result_lines.append(f"期間：{period['start']} ～ {period['end']}")
            else:
                today = datetime.now().strftime("%Y年%m月%d日")
                result_lines.append(f"基準日：{today}")
            
            return "\n".join(result_lines)
        
        except Exception as e:
            logging.error(f"イベント件数集計エラー: {e}")
            return None
    
    def _execute_event_list(self, event_key: str, period: Optional[Dict[str, str]]) -> Optional[str]:
        """
        イベント該当個体を抽出
        
        Args:
            event_key: イベントキー（例: "AI", "CALVING"）
            period: 期間辞書またはNone
        
        Returns:
            抽出結果の文字列
        """
        try:
            # イベントキーからイベント番号を取得
            event_number = self._get_event_number(event_key)
            if event_number is None:
                return None
            
            # イベントを取得
            if period:
                start_date = period["start"]
                end_date = period["end"]
                events = self.db.get_events_by_number_and_period(event_number, start_date, end_date)
            else:
                events = self.db.get_events_by_number(event_number)
            
            if not events:
                return None
            
            # 個体IDを抽出（重複除去）
            cow_ids = []
            seen_cows = set()
            
            for event in events:
                cow_auto_id = event.get("cow_auto_id")
                if cow_auto_id and cow_auto_id not in seen_cows:
                    cow = self.db.get_cow_by_auto_id(cow_auto_id)
                    if cow:
                        cow_id = cow.get("cow_id", "")
                        if cow_id:
                            cow_ids.append(cow_id)
                            seen_cows.add(cow_auto_id)
            
            if not cow_ids:
                return None
            
            # 出力フォーマット
            event_display_name = self._get_event_display_name(event_key)
            
            result_lines = []
            result_lines.append(f"＜{event_display_name}該当個体＞：{len(cow_ids)}頭")
            result_lines.append(f"個体ID：{', '.join(cow_ids)}")
            
            if period:
                result_lines.append(f"期間：{period['start']} ～ {period['end']}")
            
            return "\n".join(result_lines)
        
        except Exception as e:
            logging.error(f"イベント個体抽出エラー: {e}")
            return None
    
    def _get_item_display_name(self, item_key: str) -> str:
        """項目キーから表示名を取得（item_dictionary.jsonから）"""
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            return item_def.get("display_name", item_key)
        return item_key
    
    def _get_item_unit(self, item_key: str) -> str:
        """項目キーから単位を取得"""
        # data_typeに基づいて単位を推測
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            data_type = item_def.get("data_type", "")
            display_name = item_def.get("display_name", "")
            
            # 表示名から単位を推測
            if "日数" in display_name or "日" in display_name:
                return "日"
            elif "月齢" in display_name:
                return "ヶ月"
            elif "回数" in display_name or "回" in display_name:
                return "回"
            elif "率" in display_name or "%" in display_name:
                return "%"
            elif "乳量" in display_name or "kg" in display_name:
                return "kg"
        
        # デフォルト単位マッピング
        units = {
            "DIM": "日",
            "DIMFAI": "日",
            "DAI": "日",
            "CINT": "日",
            "CI": "日",
            "DOPN": "日",
            "MILK": "kg",
            "MFAT": "%",
            "MPROT": "%",
            "SCC": "",
            "LS": "",
            "BCS": "",
            "BRED": "回",
            "AGEFAI": "ヶ月"
        }
        return units.get(item_key, "")
    
    def _get_event_number(self, event_key: str) -> Optional[int]:
        """イベントキーからイベント番号を取得（event_dictionary.jsonから）"""
        # まずevent_key_to_numberマッピングを確認
        if event_key in self.event_key_to_number:
            return self.event_key_to_number[event_key]
        
        # event_dictionaryを直接検索
        for event_number_str, event_data in self.event_dictionary.items():
            if isinstance(event_data, dict):
                alias = event_data.get("alias")
                if alias == event_key:
                    try:
                        return int(event_number_str)
                    except ValueError:
                        pass
        
        return None
    
    def _get_event_display_name(self, event_key: str) -> str:
        """イベントキーから表示名を取得（event_dictionary.jsonから）"""
        # event_dictionaryを検索
        for event_number_str, event_data in self.event_dictionary.items():
            if isinstance(event_data, dict):
                alias = event_data.get("alias")
                if alias == event_key or event_key in [alias, event_data.get("name_jp", "")]:
                    return event_data.get("name_jp", event_key)
        
        # デフォルト表示名マッピング
        display_names = {
            "CALVING": "分娩",
            "AI": "AI",
            "ET": "ET",
            "PREG_POS": "妊娠鑑定プラス",
            "PREG_NEG": "妊娠鑑定マイナス",
            "DRY": "乾乳",
            "MILK_TEST": "乳検",
            "STOPR": "繁殖停止",
            "MAST": "乳房炎",
            "BCS": "BCS",
            "FCHK": "フレッシュチェック",
            "REPRO": "繁殖検査"
        }
        return display_names.get(event_key, event_key)

