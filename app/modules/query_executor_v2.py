"""
FALCON2 - Query Executor V2（クエリ実行ラッパー）
db_path と辞書パスから ExecutorV2 を組み立て、ExecutionPlan を実行する
"""

import logging
from pathlib import Path
from typing import Dict, Any, Union, Optional

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.execution_plan import ExecutionPlan
from modules.executor_v2 import ExecutorV2

logger = logging.getLogger(__name__)


class QueryExecutorV2:
    """
    db_path と辞書パスを受け取り、ExecutionPlan を実行するラッパー。
    UI からは db_path / item_dictionary_path / event_dictionary_path のみ渡す想定。
    """

    def __init__(
        self,
        db_path: Optional[Path] = None,
        item_dictionary_path: Optional[Path] = None,
        event_dictionary_path: Optional[Path] = None,
    ):
        self.db_path = Path(db_path) if db_path else None
        self.item_dictionary_path = Path(item_dictionary_path) if item_dictionary_path else None
        self.event_dictionary_path = Path(event_dictionary_path) if event_dictionary_path else None

    def execute(self, dsl_query: Union[ExecutionPlan, int, Any]) -> Dict[str, Any]:
        """
        ExecutionPlan の場合は DB 実行して結果を返す。
        それ以外（例: 個体IDの int）の場合は未対応の結果を返す。
        """
        if isinstance(dsl_query, ExecutionPlan):
            return self._execute_plan(dsl_query)
        if isinstance(dsl_query, int):
            return {
                "success": False,
                "data": None,
                "errors": ["個体IDの場合は一覧から開くか、検索欄にIDを入力してください。"],
                "warnings": [],
            }
        return {
            "success": False,
            "data": None,
            "errors": ["未対応のクエリ形式です。"],
            "warnings": [],
        }

    def _execute_plan(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """ExecutorV2 で ExecutionPlan を実行"""
        if not self.db_path or not self.db_path.exists():
            return {
                "success": False,
                "data": None,
                "errors": ["データベースが見つかりません。"],
                "warnings": [],
            }
        try:
            db_handler = DBHandler(self.db_path)
            try:
                farm_path = self.db_path.parent
                formula_engine = FormulaEngine(
                    db_handler,
                    self.item_dictionary_path or (farm_path / "item_dictionary.json"),
                )
                rule_engine = RuleEngine(db_handler)
                executor = ExecutorV2(
                    db_handler,
                    formula_engine,
                    rule_engine,
                    item_dictionary_path=self.item_dictionary_path,
                    event_dictionary_path=self.event_dictionary_path,
                    farm_path=farm_path,
                )
                return executor.execute(plan)
            finally:
                db_handler.close()
        except Exception as e:
            logger.error(f"QueryExecutorV2 実行エラー: {e}", exc_info=True)
            return {
                "success": False,
                "data": None,
                "errors": [f"実行エラー: {str(e)}"],
                "warnings": [],
            }
