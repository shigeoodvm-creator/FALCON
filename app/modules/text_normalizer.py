"""
FALCON2 - Text Normalizer（入力正規化モジュール）
ユーザー入力を正規化し、スペース非依存の解析を可能にする
"""

import re
import unicodedata
from typing import Dict


def normalize_user_input(text: str) -> Dict[str, str]:
    """
    ユーザー入力を正規化
    
    処理内容（順序厳守）：
    1) NFKC 正規化（全角英数 → 半角）
    2) 全角スペース（\u3000）→ 半角スペース
    3) strip()
    4) 連続するスペースを1個に圧縮
    5) 解析用文字列ではスペースを無視できる形も保持
    
    Args:
        text: 元の文字列
    
    Returns:
        {
            "raw": 元の文字列,
            "normalized": 正規化済み文字列（スペースあり）,
            "nospace": スペース除去版
        }
    """
    # 1) NFKC 正規化（全角英数 → 半角）
    normalized = unicodedata.normalize('NFKC', text)
    
    # 2) 全角スペース（\u3000）→ 半角スペース
    normalized = normalized.replace('\u3000', ' ')
    
    # 3) strip()
    normalized = normalized.strip()
    
    # 4) 連続するスペースを1個に圧縮
    normalized = re.sub(r'\s+', ' ', normalized)
    
    # 5) スペース除去版を作成（解析用）
    nospace = normalized.replace(' ', '')
    
    return {
        "raw": text,
        "normalized": normalized,
        "nospace": nospace
    }






















