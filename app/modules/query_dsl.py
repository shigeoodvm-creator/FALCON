"""
FALCON2 - Query DSL（クエリ表現型）
QueryRouter が返す UnhandledQuery / AmbiguousQuery および実行計画の型定義
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


@dataclass
class UnhandledQuery:
    """未対応クエリ（理由付き）"""
    reason: str


@dataclass
class AmbiguousQuery:
    """曖昧なクエリ（複数候補）"""
    original_query: str
    candidates: Optional[List[Dict[str, str]]] = field(default_factory=list)
    # candidates: [{"item_key": "...", "label": "...", "description": "..."}, ...]
