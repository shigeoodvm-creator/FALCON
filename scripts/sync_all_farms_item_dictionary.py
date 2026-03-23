"""
FALCON2 - 全農場のitem_dictionary.jsonを同期
config_default/item_dictionary.json（マスター）から全農場のitem_dictionary.jsonに新しい項目を追加
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
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('sync_all_farms_item_dictionary.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)


def main():
    """メイン処理"""
    print("=" * 60)
    print("全農場のitem_dictionary.jsonを同期")
    print("=" * 60)
    print(f"マスター辞書: {APP_ROOT / 'config_default' / 'item_dictionary.json'}")
    print(f"または: {APP_ROOT / 'docs' / 'item_dictionary.json'}")
    print()
    
    # DictionarySyncインスタンスを作成
    sync = DictionarySync(APP_ROOT)
    
    # 全農場の辞書を同期
    farms_root = Path("C:/FARMS")
    print(f"農場フォルダ: {farms_root}")
    print()
    
    if not farms_root.exists():
        print(f"エラー: 農場フォルダが見つかりません: {farms_root}")
        logger.error(f"農場フォルダが見つかりません: {farms_root}")
        return 1
    
    print("全農場のitem_dictionary.jsonを同期中...")
    print()
    
    results = sync.sync_all_farms(farms_root)
    
    print("=" * 60)
    print("同期完了")
    print("=" * 60)
    print(f"対象農場数: {len(results)}")
    print()
    
    success_count = 0
    error_count = 0
    
    for farm_name, result in results.items():
        if 'error' in result:
            print(f"  [ERROR] {farm_name}: エラー - {result['error']}")
            error_count += 1
        else:
            print(f"  [OK] {farm_name}: 同期完了")
            success_count += 1
    
    print()
    print(f"成功: {success_count} 農場")
    if error_count > 0:
        print(f"エラー: {error_count} 農場")
    
    return 0 if error_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

