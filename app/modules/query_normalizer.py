"""
FALCON2 - Query Normalizer（クエリ正規化モジュール）
日本語入力を辞書ベースで正規化し、内部クエリを確定する
AI（ChatGPT API）は一切使用しない
"""

import json
import re
import logging
import unicodedata
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime, date

from modules.text_normalizer import normalize_user_input


class QueryNormalizer:
    """日本語クエリを正規化辞書に基づいて正規化"""
    
    def __init__(self, normalization_dir: Optional[Path] = None):
        """
        初期化
        
        Args:
            normalization_dir: normalizationディレクトリのパス（Noneの場合はデフォルト）
        """
        if normalization_dir is None:
            from constants import NORMALIZATION_DIR
            normalization_dir = NORMALIZATION_DIR
        
        self.normalization_dir = Path(normalization_dir)
        self.items_dict: Dict[str, Any] = {}  # priority と aliases をサポート
        self.items_priority: Dict[str, int] = {}  # item_key -> priority のマッピング
        self.events_dict: Dict[str, List[str]] = {}
        self.terms_dict: Dict[str, List[str]] = {}
        
        self._load_dictionaries()
    
    def _load_dictionaries(self):
        """正規化辞書を読み込む"""
        try:
            # items.json
            items_path = self.normalization_dir / "items.json"
            if items_path.exists():
                with open(items_path, 'r', encoding='utf-8') as f:
                    raw_items_dict = json.load(f)
                    
                    # 後方互換性のため、配列形式と辞書形式の両方をサポート
                    for item_key, item_data in raw_items_dict.items():
                        if isinstance(item_data, dict):
                            # 新しい形式: {"priority": 100, "aliases": [...]}
                            self.items_dict[item_key] = item_data.get("aliases", [])
                            priority = item_data.get("priority", 0)
                            self.items_priority[item_key] = priority
                        else:
                            # 古い形式: ["alias1", "alias2", ...]
                            self.items_dict[item_key] = item_data
                            self.items_priority[item_key] = 0  # デフォルト priority
            else:
                logging.warning(f"items.json not found: {items_path}")
            
            # events.json
            events_path = self.normalization_dir / "events.json"
            if events_path.exists():
                with open(events_path, 'r', encoding='utf-8') as f:
                    self.events_dict = json.load(f)
            else:
                logging.warning(f"events.json not found: {events_path}")
            
            # terms.json
            terms_path = self.normalization_dir / "terms.json"
            if terms_path.exists():
                with open(terms_path, 'r', encoding='utf-8') as f:
                    self.terms_dict = json.load(f)
            else:
                logging.warning(f"terms.json not found: {terms_path}")
        
        except Exception as e:
            logging.error(f"正規化辞書読み込みエラー: {e}")
    
    def normalize_query(self, query: str) -> Dict[str, Any]:
        """
        クエリを正規化
        
        Args:
            query: 日本語クエリ文字列
        
        Returns:
            正規化されたクエリ辞書
            {
                "item": item_key or None,
                "event": event_key or None,
                "term": term_key or None,
                "period": {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"} or None,
                "error": error_message or None
            }
        """
        query = query.strip()
        
        if not query:
            return {"error": "クエリが空です"}
        
        # 語彙ベースの強制マッピング（最優先）
        # 特定の日本語表現は DIM に強制マッピング
        forced_item = self._check_forced_item_mapping(query)
        if forced_item:
            # 強制マッピングが適用された場合は、item を確定して他のマッチングをスキップ
            event = self._match_event(query)
            term = self._match_term(query)
            period = self._parse_period(query)
            
            # イベント・集計語の複数一致チェック
            if event and isinstance(event, list) and len(event) > 1:
                return {"error": f"イベントが複数一致しました: {event}"}
            if term and isinstance(term, list) and len(term) > 1:
                return {"error": f"集計語が複数一致しました: {term}"}
            
            # 単一値に変換
            event_key = event[0] if isinstance(event, list) and len(event) == 1 else (event if event else None)
            term_key = term[0] if isinstance(term, list) and len(term) == 1 else (term if term else None)
            
            logging.info(f"[QueryNormalizer] 強制マッピング適用: '{query}' → item={forced_item}")
            
            return {
                "item": forced_item,
                "event": event_key,
                "term": term_key,
                "period": period,
                "error": None
            }
        
        # 各要素をマッチング
        item_result = self._match_item(query)
        event = self._match_event(query)
        term = self._match_term(query)
        period = self._parse_period(query)
        
        # 項目の複数一致を解決（完全一致が1件の場合は即確定、priority判定をスキップ）
        item_key = None
        ambiguous_candidates = None
        if item_result:
            exact_matches = item_result.get("exact", [])
            prefix_matches = item_result.get("prefix", [])
            partial_matches = item_result.get("partial", [])
            
            # ① 完全一致が1件の場合は即確定（priority判定をスキップ）
            if len(exact_matches) == 1:
                item_key = exact_matches[0]
                logging.info(f"[QueryNormalizer] 完全一致で確定: '{query}' → {item_key}")
            # ② 完全一致が複数の場合、priority で解決
            elif len(exact_matches) > 1:
                resolve_result = self._resolve_item_conflict({"exact": exact_matches, "partial": []}, query)
                if isinstance(resolve_result, dict) and resolve_result.get("error_type") == "AMBIGUOUS_ITEM":
                    return resolve_result
                item_key = resolve_result
            # ③ 完全一致がない場合、前方一致を確認
            elif len(prefix_matches) > 0:
                if len(prefix_matches) == 1:
                    item_key = prefix_matches[0]
                    logging.info(f"[QueryNormalizer] 前方一致で確定: '{query}' → {item_key}")
                else:
                    # 前方一致が複数の場合、priority で解決
                    resolve_result = self._resolve_item_conflict({"exact": [], "partial": prefix_matches}, query)
                    if isinstance(resolve_result, dict) and resolve_result.get("error_type") == "AMBIGUOUS_ITEM":
                        return resolve_result
                    item_key = resolve_result
            # ④ 部分一致を確認
            elif len(partial_matches) > 0:
                if len(partial_matches) == 1:
                    item_key = partial_matches[0]
                    logging.info(f"[QueryNormalizer] 部分一致で確定: '{query}' → {item_key}")
                else:
                    # 部分一致が複数の場合、priority で解決
                    resolve_result = self._resolve_item_conflict({"exact": [], "partial": partial_matches}, query)
                    if isinstance(resolve_result, dict) and resolve_result.get("error_type") == "AMBIGUOUS_ITEM":
                        return resolve_result
                    item_key = resolve_result
        
        # イベント・集計語の複数一致チェック（従来通り）
        if event and isinstance(event, list) and len(event) > 1:
            return {"error": f"イベントが複数一致しました: {event}"}
        if term and isinstance(term, list) and len(term) > 1:
            return {"error": f"集計語が複数一致しました: {term}"}
        
        # 単一値に変換
        event_key = event[0] if isinstance(event, list) and len(event) == 1 else (event if event else None)
        term_key = term[0] if isinstance(term, list) and len(term) == 1 else (term if term else None)
        
        return {
            "item": item_key,
            "event": event_key,
            "term": term_key,
            "period": period,
            "error": None
        }
    
    def _match_item(self, query: str) -> Dict[str, List[str]]:
        """
        項目をマッチング（完全一致最優先、前方一致、部分一致の順）
        スペース非依存：normalized と nospace の両方でマッチング
        
        優先順位：
        ① aliasの完全一致
        ② 前方一致
        ③ 部分一致
        
        Returns:
            {"exact": [...], "prefix": [...], "partial": [...]} の形式
        """
        # クエリを正規化（normalized と nospace の両方を取得）
        normalized_result = normalize_user_input(query)
        normalized_query = normalized_result["normalized"]
        nospace_query = normalized_result["nospace"]
        
        exact_matches = []  # 完全一致
        prefix_matches = []  # 前方一致
        partial_matches = []  # 部分一致
        
        for item_key, synonyms in self.items_dict.items():
            for synonym in synonyms:
                # 同義語も正規化（normalized と nospace の両方を取得）
                synonym_normalized_result = normalize_user_input(synonym)
                normalized_synonym = synonym_normalized_result["normalized"]
                nospace_synonym = synonym_normalized_result["nospace"]
                
                # ① 完全一致チェック（normalized と nospace の両方で）
                if (normalized_synonym == normalized_query or 
                    nospace_synonym == nospace_query or
                    normalized_synonym == nospace_query or
                    nospace_synonym == normalized_query):
                    if item_key not in exact_matches:
                        exact_matches.append(item_key)
                    break
                
                # ② 前方一致チェック（クエリがsynonymで始まる）
                if (normalized_query.startswith(normalized_synonym) or
                    normalized_query.startswith(nospace_synonym) or
                    nospace_query.startswith(normalized_synonym) or
                    nospace_query.startswith(nospace_synonym)):
                    if item_key not in prefix_matches:
                        prefix_matches.append(item_key)
                    break
                
                # ③ 部分一致チェック（normalized と nospace の両方で）
                if (normalized_synonym in normalized_query or 
                    normalized_synonym in nospace_query or
                    nospace_synonym in normalized_query or
                    nospace_synonym in nospace_query or
                    normalized_query in normalized_synonym or
                    normalized_query in nospace_synonym or
                    nospace_query in normalized_synonym or
                    nospace_query in nospace_synonym):
                    if item_key not in partial_matches:
                        partial_matches.append(item_key)
                    break
        
        return {
            "exact": exact_matches,
            "prefix": prefix_matches,
            "partial": partial_matches
        }
    
    def _check_forced_item_mapping(self, query: str) -> Optional[str]:
        """
        語彙ベースの強制マッピングをチェック
        
        特定の日本語表現が含まれる場合は、DIM に強制マッピング
        DUE / CINT との曖昧一致を無視する
        
        Args:
            query: クエリ文字列
        
        Returns:
            強制マッピングされる item_key または None
        """
        # クエリを正規化（スペース非依存の判定のため）
        normalized_result = normalize_user_input(query)
        normalized_query = normalized_result["normalized"]
        nospace_query = normalized_result["nospace"]
        
        # DIM に強制マッピングするキーワードリスト
        dim_force_keywords = [
            "分娩後日数",
            "分娩後 日数",  # スペースあり
            "産後日数",
            "泌乳日数"
        ]
        
        # 各キーワードを正規化してチェック
        for keyword in dim_force_keywords:
            keyword_normalized = normalize_user_input(keyword)
            keyword_normalized_text = keyword_normalized["normalized"]
            keyword_nospace_text = keyword_normalized["nospace"]
            
            # 正規化後のクエリにキーワードが含まれているかチェック
            if (keyword_normalized_text in normalized_query or
                keyword_nospace_text in nospace_query or
                keyword_normalized_text in nospace_query or
                keyword_nospace_text in normalized_query):
                return "DIM"
        
        return None
    
    def _resolve_item_conflict(self, item_result: Dict[str, List[str]], query: str):
        """
        項目の複数一致を priority で解決
        
        Args:
            item_result: _match_item の戻り値 {"exact": [...], "partial": [...]}
            query: 元のクエリ文字列（ログ出力用）
        
        Returns:
            解決された item_key または {"error_type": "AMBIGUOUS_ITEM", ...}（解決不能の場合）
        """
        if not item_result:
            return None
        
        exact_matches = item_result.get("exact", [])
        partial_matches = item_result.get("partial", [])
        
        # ① 完全一致があればそれを優先
        if exact_matches:
            if len(exact_matches) == 1:
                return exact_matches[0]
            
            # 完全一致が複数の場合、priority で解決
            candidates = exact_matches
            match_type = "完全一致"
        else:
            # 完全一致がない場合、部分一致を使用
            if not partial_matches:
                return None
            
            if len(partial_matches) == 1:
                return partial_matches[0]
            
            candidates = partial_matches
            match_type = "部分一致"
        
        # ② priority が最大の item_key を採用
        best_item = None
        best_priority = -1
        priority_ties = []  # priority が同点の候補
        
        for item_key in candidates:
            priority = self.items_priority.get(item_key, 0)
            if priority > best_priority:
                best_priority = priority
                best_item = item_key
                priority_ties = [item_key]
            elif priority == best_priority:
                priority_ties.append(item_key)
        
        # ③ priority が同点の場合、候補情報を返す
        if len(priority_ties) > 1:
            logging.warning(f"[QueryNormalizer] 項目の複数一致で priority が同点: "
                          f"候補={priority_ties}, priority={best_priority}, "
                          f"クエリ='{query}'")
            
            # 正規化されたクエリを取得
            normalized_result = normalize_user_input(query)
            normalized_query = normalized_result["normalized"]
            
            # 候補情報を構築（label と description は QueryMappings で取得する必要があるため、
            # ここでは item_key のみを返し、QueryRouter で補完する）
            return {
                "error_type": "AMBIGUOUS_ITEM",
                "original_query": query,
                "normalized_query": normalized_query,
                "candidate_item_keys": priority_ties
            }
        
        # 解決成功：ログ出力（INFO レベル）
        logging.info(f"[QueryNormalizer] 項目の複数一致を解決: "
                    f"候補={candidates}, 採用={best_item}, "
                    f"priority={best_priority}, マッチタイプ={match_type}, "
                    f"クエリ='{query}'")
        
        return best_item
    
    def _match_event(self, query: str) -> Optional[List[str]]:
        """
        イベントをマッチング（完全一致優先、部分一致も許可）
        スペース非依存：normalized と nospace の両方でマッチング
        
        Returns:
            マッチしたevent_keyのリスト（複数一致の可能性あり）
        """
        # クエリを正規化（normalized と nospace の両方を取得）
        normalized_result = normalize_user_input(query)
        normalized_query = normalized_result["normalized"]
        nospace_query = normalized_result["nospace"]
        
        matched_keys = []
        exact_matches = []  # 完全一致
        
        for event_key, synonyms in self.events_dict.items():
            for synonym in synonyms:
                # 同義語も正規化（normalized と nospace の両方を取得）
                synonym_normalized_result = normalize_user_input(synonym)
                normalized_synonym = synonym_normalized_result["normalized"]
                nospace_synonym = synonym_normalized_result["nospace"]
                
                # 完全一致チェック（normalized と nospace の両方で）
                if (normalized_synonym == normalized_query or 
                    nospace_synonym == nospace_query or
                    normalized_synonym == nospace_query or
                    nospace_synonym == normalized_query):
                    if event_key not in exact_matches:
                        exact_matches.append(event_key)
                    break
                
                # 部分一致チェック（完全一致がない場合のみ、normalized と nospace の両方で）
                if (normalized_synonym in normalized_query or 
                    normalized_synonym in nospace_query or
                    nospace_synonym in normalized_query or
                    nospace_synonym in nospace_query or
                    normalized_query in normalized_synonym or
                    normalized_query in nospace_synonym or
                    nospace_query in normalized_synonym or
                    nospace_query in nospace_synonym):
                    if event_key not in matched_keys:
                        matched_keys.append(event_key)
                    break
        
        # 完全一致があればそれを優先
        if exact_matches:
            return exact_matches
        
        return matched_keys if matched_keys else None
    
    def _match_term(self, query: str) -> Optional[List[str]]:
        """
        集計・修飾語をマッチング（完全一致優先、部分一致も許可）
        スペース非依存：normalized と nospace の両方でマッチング
        
        Returns:
            マッチしたterm_keyのリスト（複数一致の可能性あり）
        """
        # クエリを正規化（normalized と nospace の両方を取得）
        normalized_result = normalize_user_input(query)
        normalized_query = normalized_result["normalized"]
        nospace_query = normalized_result["nospace"]
        
        matched_keys = []
        exact_matches = []  # 完全一致
        
        for term_key, synonyms in self.terms_dict.items():
            for synonym in synonyms:
                # 同義語も正規化（normalized と nospace の両方を取得）
                synonym_normalized_result = normalize_user_input(synonym)
                normalized_synonym = synonym_normalized_result["normalized"]
                nospace_synonym = synonym_normalized_result["nospace"]
                
                # 完全一致チェック（normalized と nospace の両方で）
                if (normalized_synonym == normalized_query or 
                    nospace_synonym == nospace_query or
                    normalized_synonym == nospace_query or
                    nospace_synonym == normalized_query):
                    if term_key not in exact_matches:
                        exact_matches.append(term_key)
                    break
                
                # 部分一致チェック（完全一致がない場合のみ、normalized と nospace の両方で）
                if (normalized_synonym in normalized_query or 
                    normalized_synonym in nospace_query or
                    nospace_synonym in normalized_query or
                    nospace_synonym in nospace_query or
                    normalized_query in normalized_synonym or
                    normalized_query in nospace_synonym or
                    nospace_query in normalized_synonym or
                    nospace_query in nospace_synonym):
                    if term_key not in matched_keys:
                        matched_keys.append(term_key)
                    break
        
        # 完全一致があればそれを優先
        if exact_matches:
            return exact_matches
        
        return matched_keys if matched_keys else None
    
    def _parse_period(self, query: str) -> Optional[Dict[str, str]]:
        """
        期間を解析
        
        例：
        - 「10月」→ YYYY-10-01 ～ YYYY-10-31
        - 「2024年10月」→ 2024-10-01 ～ 2024-10-31
        - 「今年」→ YYYY-01-01 ～ 今日
        - 「昨年」→ (YYYY-1)-01-01 ～ (YYYY-1)-12-31
        
        Returns:
            期間辞書 {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"} または None
        """
        today = date.today()
        current_year = today.year
        current_month = today.month
        
        # 「今年」
        if "今年" in query:
            start_date = date(current_year, 1, 1)
            end_date = today
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        # 「昨年」
        if "昨年" in query or "去年" in query:
            start_date = date(current_year - 1, 1, 1)
            end_date = date(current_year - 1, 12, 31)
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        # 「YYYY年MM月」パターン
        pattern_year_month = re.search(r'(\d{4})年(\d{1,2})月', query)
        if pattern_year_month:
            year = int(pattern_year_month.group(1))
            month = int(pattern_year_month.group(2))
            if 1 <= month <= 12:
                # 月の最初の日
                start_date = date(year, month, 1)
                # 月の最後の日
                if month == 12:
                    end_date = date(year, 12, 31)
                else:
                    end_date = date(year, month + 1, 1) - date.resolution
                return {
                    "start": start_date.strftime("%Y-%m-%d"),
                    "end": end_date.strftime("%Y-%m-%d")
                }
        
        # 「MM月」パターン（年が指定されていない場合は今年）
        pattern_month = re.search(r'(\d{1,2})月', query)
        if pattern_month:
            month = int(pattern_month.group(1))
            if 1 <= month <= 12:
                year = current_year
                # 月の最初の日
                start_date = date(year, month, 1)
                # 月の最後の日
                if month == 12:
                    end_date = date(year, 12, 31)
                else:
                    end_date = date(year, month + 1, 1) - date.resolution
                return {
                    "start": start_date.strftime("%Y-%m-%d"),
                    "end": end_date.strftime("%Y-%m-%d")
                }
        
        # 「YYYY-MM-DD ～ YYYY-MM-DD」パターン
        pattern_date_range = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})\s*[～〜]\s*(\d{4})-(\d{1,2})-(\d{1,2})', query)
        if pattern_date_range:
            start_year = int(pattern_date_range.group(1))
            start_month = int(pattern_date_range.group(2))
            start_day = int(pattern_date_range.group(3))
            end_year = int(pattern_date_range.group(4))
            end_month = int(pattern_date_range.group(5))
            end_day = int(pattern_date_range.group(6))
            
            try:
                start_date = date(start_year, start_month, start_day)
                end_date = date(end_year, end_month, end_day)
                return {
                    "start": start_date.strftime("%Y-%m-%d"),
                    "end": end_date.strftime("%Y-%m-%d")
                }
            except ValueError:
                pass
        
        return None
    
    def _normalize_query_text(self, text: str) -> str:
        """
        クエリテキストを正規化（マッチング用）
        - 全角半角統一
        - 記号除去
        - 空白除去
        - 大文字小文字統一（英数字のみ）
        """
        if not text:
            return ""
        
        # 全角半角統一（全角に統一）
        text = unicodedata.normalize('NFKC', text)
        
        # 記号除去
        text = re.sub(r'[，。、．（）()【】「」『』〈〉《》［］｛｝・\-]', '', text)
        
        # 空白除去
        text = re.sub(r'\s+', '', text)
        
        # 英数字は小文字に統一
        text = text.lower()
        
        return text.strip()

