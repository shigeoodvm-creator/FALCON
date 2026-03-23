"""
FALCON2 - Query Mappings（集計関数マッピング）
item_key/event_key から DB カラム/計算方法への確定マッピング
"""

from typing import Dict, List, Optional, Set
import json
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class QueryMappings:
    """クエリマッピング管理クラス"""
    
    def __init__(self, item_dictionary_path: Optional[Path] = None,
                 event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            item_dictionary_path: item_dictionary.json のパス
            event_dictionary_path: event_dictionary.json のパス
        """
        self.item_dictionary: Dict[str, Any] = {}
        self.event_dictionary: Dict[str, Any] = {}
        self.event_key_to_numbers: Dict[str, List[int]] = {}  # event_key -> [event_number, ...]
        
        self._load_dictionaries(item_dictionary_path, event_dictionary_path)
    
    def _load_dictionaries(self, item_dict_path: Optional[Path],
                          event_dict_path: Optional[Path]):
        """辞書を読み込む"""
        # item_dictionary.json
        if item_dict_path and item_dict_path.exists():
            try:
                with open(item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                logger.warning(f"item_dictionary.json読み込みエラー: {e}")
        
        # event_dictionary.json
        if event_dict_path and event_dict_path.exists():
            try:
                with open(event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
                
                # event_key -> event_number のマッピングを作成
                for event_number_str, event_data in self.event_dictionary.items():
                    if isinstance(event_data, dict):
                        alias = event_data.get("alias")
                        if alias:
                            try:
                                event_number = int(event_number_str)
                                
                                # alias をキーとして登録
                                if alias not in self.event_key_to_numbers:
                                    self.event_key_to_numbers[alias] = []
                                if event_number not in self.event_key_to_numbers[alias]:
                                    self.event_key_to_numbers[alias].append(event_number)
                                
                                # 正規化辞書で使用されるキーもマッピング
                                if alias == "CALV":
                                    if "CALVING" not in self.event_key_to_numbers:
                                        self.event_key_to_numbers["CALVING"] = []
                                    if event_number not in self.event_key_to_numbers["CALVING"]:
                                        self.event_key_to_numbers["CALVING"].append(event_number)
                                elif alias == "PDP":
                                    if "PREG_POS" not in self.event_key_to_numbers:
                                        self.event_key_to_numbers["PREG_POS"] = []
                                    if event_number not in self.event_key_to_numbers["PREG_POS"]:
                                        self.event_key_to_numbers["PREG_POS"].append(event_number)
                                elif alias == "PDN":
                                    if "PREG_NEG" not in self.event_key_to_numbers:
                                        self.event_key_to_numbers["PREG_NEG"] = []
                                    if event_number not in self.event_key_to_numbers["PREG_NEG"]:
                                        self.event_key_to_numbers["PREG_NEG"].append(event_number)
                                elif alias == "MILK_TEST":
                                    if "MILK_TEST" not in self.event_key_to_numbers:
                                        self.event_key_to_numbers["MILK_TEST"] = []
                                    if event_number not in self.event_key_to_numbers["MILK_TEST"]:
                                        self.event_key_to_numbers["MILK_TEST"].append(event_number)
                            except ValueError:
                                pass
            except Exception as e:
                logger.warning(f"event_dictionary.json読み込みエラー: {e}")
    
    def get_target_filter_sql(self, target: str) -> str:
        """
        対象牛フィルタのSQL条件を取得
        
        Args:
            target: "all", "milking", "parous", "heifer"
        
        Returns:
            SQL WHERE句の条件（例： "rc != 0"）
        """
        if target == "milking":
            # 搾乳牛：rc != 0（繁殖コードが0でない）
            return "rc != 0 AND rc IS NOT NULL"
        elif target == "parous":
            # 経産牛：lact >= 1
            return "lact >= 1 AND lact IS NOT NULL"
        elif target == "heifer":
            # 未経産牛：lact == 0 または lact IS NULL
            return "(lact = 0 OR lact IS NULL)"
        else:
            # "all" の場合は条件なし
            return "1=1"
    
    def get_item_column(self, item_key: str) -> Optional[str]:
        """
        item_key に対応する cow テーブルのカラム名を取得
        
        Args:
            item_key: 項目キー（例： "DIM", "LACT"）
        
        Returns:
            カラム名（存在する場合）、None（計算項目の場合）
        """
        # cowテーブルに直接存在する項目のマッピング
        direct_mapping = {
            "LACT": "lact",
            "RC": "rc",
            "CLVD": "clvd",
            "BRD": "brd",
            "BTHD": "bthd",
            "JPN10": "jpn10",
            "COW_ID": "cow_id",
            "PEN": "pen",
        }
        
        if item_key in direct_mapping:
            return direct_mapping[item_key]
        
        # item_dictionary を確認
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            origin = item_def.get("origin") or item_def.get("type", "")
            
            # origin が "core" または "source" の場合は cow テーブルに存在する可能性
            if origin in ("core", "source"):
                # 小文字に変換してカラム名として使用
                return item_key.lower()
        
        # 計算項目の場合は None
        return None
    
    def is_calc_item(self, item_key: str) -> bool:
        """
        item_key が計算項目かどうかを判定
        
        Args:
            item_key: 項目キー
        
        Returns:
            True: 計算項目（FormulaEngineで計算が必要）
            False: DBカラムから直接取得可能
        """
        column = self.get_item_column(item_key)
        return column is None
    
    def get_event_numbers(self, event_key: str) -> List[int]:
        """
        event_key に対応する event_number のリストを取得
        
        Args:
            event_key: イベントキー（例： "CALVING", "AI"）
        
        Returns:
            event_number のリスト
        """
        return self.event_key_to_numbers.get(event_key, [])
    
    def get_item_label(self, item_key: str) -> str:
        """
        item_key に対応する表示名（label）を取得
        
        Args:
            item_key: 項目キー（例： "DIM", "DAI"）
        
        Returns:
            表示名（存在しない場合は item_key を返す）
        """
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            return item_def.get("display_name") or item_def.get("label") or item_key
        return item_key
    
    def get_item_description(self, item_key: str) -> str:
        """
        item_key に対応する説明（description）を取得
        
        Args:
            item_key: 項目キー（例： "DIM", "DAI"）
        
        Returns:
            説明（存在しない場合は空文字列を返す）
        """
        if item_key in self.item_dictionary:
            item_def = self.item_dictionary[item_key]
            return item_def.get("description") or ""
        return ""

