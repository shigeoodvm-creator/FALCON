"""
FALCON2 - Normalization Dictionary Generator Module
アプリケーション内から呼び出し可能な正規化辞書生成モジュール
"""

import subprocess
import sys
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


def generate_normalization_dict(farm_path: Path = None):
    """
    normalization 辞書を自動生成
    
    Args:
        farm_path: 農場フォルダのパス（Noneの場合はconfig_defaultを使用）
    
    Note:
        scripts/generate_normalization_dict.py を実行して辞書を生成
    """
    try:
        # スクリプトパス
        base_dir = Path(__file__).parent.parent.parent
        script_path = base_dir / "scripts" / "generate_normalization_dict.py"
        
        if not script_path.exists():
            logger.warning(f"normalization辞書生成スクリプトが見つかりません: {script_path}")
            return False
        
        # スクリプトを実行
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(base_dir),
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        if result.returncode == 0:
            logger.info("normalization辞書を自動生成しました")
            return True
        else:
            logger.error(f"normalization辞書生成エラー: {result.stderr}")
            return False
    
    except Exception as e:
        logger.error(f"normalization辞書生成中にエラーが発生しました: {e}")
        return False






















