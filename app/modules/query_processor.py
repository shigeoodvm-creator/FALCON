"""
FALCON2 - Query Processor（クエリ処理統合モジュール）
日本語クエリを正規化し、DB集計を実行する統合処理
"""

from typing import Optional
from pathlib import Path
import logging

from modules.query_normalizer import QueryNormalizer
from modules.query_executor import QueryExecutor
from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine


class QueryProcessor:
    """日本語クエリを処理する統合クラス"""
    
    def __init__(self, db_handler: DBHandler, formula_engine: FormulaEngine,
                 rule_engine: RuleEngine, item_dictionary_path=None,
                 event_dictionary_path=None, normalization_dir=None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            item_dictionary_path: item_dictionary.json のパス
            event_dictionary_path: event_dictionary.json のパス
            normalization_dir: normalizationディレクトリのパス
        """
        self.normalizer = QueryNormalizer(normalization_dir)
        self.executor = QueryExecutor(
            db_handler, formula_engine, rule_engine,
            item_dictionary_path, event_dictionary_path
        )
    
    def process(self, query: str) -> str:
        """
        日本語クエリを処理
        
        Args:
            query: 日本語クエリ文字列
        
        Returns:
            処理結果の文字列（未対応の場合は「未対応」）
        """
        # 正規化
        normalized = self.normalizer.normalize_query(query)
        
        # エラーチェック
        if normalized.get("error"):
            return f"未対応：{normalized['error']}"
        
        # 実行
        result = self.executor.execute(normalized)
        
        if result is None:
            return "未対応：このクエリは処理できません"
        
        return result






















