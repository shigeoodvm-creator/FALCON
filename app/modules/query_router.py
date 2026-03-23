"""
FALCON2 - QueryRouter（最小構成）
数値のみを牛IDとして解釈し、それ以外はNoneを返す。
"""

import re
import logging
from typing import Optional

logger = logging.getLogger(__name__)


class QueryRouter:
    """
    最小構成のQueryRouter
    - 入力文字列が数値のみの場合のみ牛IDとして解釈
    - それ以外はNoneを返す
    """
    
    def __init__(self, *args, **kwargs):
        """初期化（引数は互換性のため受け取るが使用しない）"""
        logger.info("QueryRouter (minimal) initialized")
    
    def route(self, query: str, **kwargs) -> Optional[int]:
        """
        入力文字列を解釈して牛IDを返す
        
        Args:
            query: 入力文字列
            **kwargs: 互換性のため受け取るが使用しない
        
        Returns:
            int: 数値のみの場合、牛ID（整数）
            None: それ以外の場合
        """
        if query is None:
            return None
        
        # 前後の空白を削除
        text = str(query).strip()
        
        # 空文字列の場合はNone
        if not text:
            return None
        
        # 正規表現で数値のみかチェック
        if re.match(r'^\d+$', text):
            try:
                cow_id = int(text)
                logger.info(f"Cow ID command detected: {cow_id}")
                return cow_id
            except ValueError:
                # int変換に失敗した場合（通常は発生しない）
                logger.debug(f"Failed to convert to int: {text}")
                return None
        else:
            # 数値以外の場合はNoneを返す（ログは出さない）
            return None
 
