"""
FALCON2 - 全農場のイベント辞書を同期化するスクリプト
"""

import sys
import logging
from pathlib import Path

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from modules.dictionary_sync import DictionarySync

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def main():
    """全農場のイベント辞書を同期化"""
    logger.info("全農場のイベント辞書同期を開始します...")
    
    # DictionarySyncを初期化
    falcon_root = APP_ROOT
    sync = DictionarySync(falcon_root)
    
    # 全農場の辞書を同期
    results = sync.sync_all_farms()
    
    # 結果を表示
    logger.info("=" * 60)
    logger.info("同期結果:")
    logger.info("=" * 60)
    for farm_name, result in results.items():
        if 'error' in result:
            logger.error(f"  {farm_name}: エラー - {result['error']}")
        else:
            logger.info(f"  {farm_name}: 同期完了")
    
    logger.info("=" * 60)
    logger.info(f"同期対象農場数: {len(results)}")
    logger.info("全農場のイベント辞書同期が完了しました。")


if __name__ == "__main__":
    main()













