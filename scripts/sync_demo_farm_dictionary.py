"""
FALCON2 - デモファームのイベント辞書を全農場・メイン辞書に反映
デモファームの event_dictionary.json を config_default と全農場にコピー
"""

import sys
import json
import shutil
import logging
from pathlib import Path

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('sync_demo_farm_dictionary.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)


def main():
    """メイン処理"""
    # デモファームのパス
    demo_farm_path = Path("C:/FARMS/デモファーム")
    demo_event_dict_path = demo_farm_path / "event_dictionary.json"
    demo_item_dict_path = demo_farm_path / "item_dictionary.json"
    
    # メイン辞書のパス
    config_default_dir = APP_ROOT / "config_default"
    main_event_dict_path = config_default_dir / "event_dictionary.json"
    main_item_dict_path = config_default_dir / "item_dictionary.json"
    
    # デモファームの辞書が存在するか確認
    if not demo_event_dict_path.exists():
        logger.error(f"デモファームの event_dictionary.json が見つかりません: {demo_event_dict_path}")
        print(f"エラー: デモファームの event_dictionary.json が見つかりません: {demo_event_dict_path}")
        return 1
    
    if not demo_item_dict_path.exists():
        logger.error(f"デモファームの item_dictionary.json が見つかりません: {demo_item_dict_path}")
        print(f"エラー: デモファームの item_dictionary.json が見つかりません: {demo_item_dict_path}")
        return 1
    
    print("=" * 60)
    print("デモファームの辞書を全農場・メイン辞書に反映")
    print("=" * 60)
    print(f"デモファーム: {demo_farm_path}")
    print()
    
    # 1. メイン辞書（config_default）にコピー
    print("【1】メイン辞書（config_default）を更新...")
    try:
        config_default_dir.mkdir(parents=True, exist_ok=True)
        
        # event_dictionary.json をコピー
        shutil.copy2(demo_event_dict_path, main_event_dict_path)
        logger.info(f"event_dictionary.json copied to: {main_event_dict_path}")
        print(f"  ✓ event_dictionary.json をコピーしました")
        
        # item_dictionary.json をコピー
        shutil.copy2(demo_item_dict_path, main_item_dict_path)
        logger.info(f"item_dictionary.json copied to: {main_item_dict_path}")
        print(f"  ✓ item_dictionary.json をコピーしました")
        
    except Exception as e:
        logger.error(f"メイン辞書の更新に失敗: {e}")
        print(f"  ✗ エラー: {e}")
        return 1
    
    print()
    
    # 2. 全農場の辞書を更新
    print("【2】全農場の辞書を更新...")
    farms_root = Path("C:/FARMS")
    
    if not farms_root.exists():
        logger.warning(f"農場フォルダが見つかりません: {farms_root}")
        print(f"  警告: 農場フォルダが見つかりません: {farms_root}")
        return 0
    
    # デモファームの辞書を読み込む
    try:
        with open(demo_event_dict_path, 'r', encoding='utf-8') as f:
            demo_event_dict = json.load(f)
        with open(demo_item_dict_path, 'r', encoding='utf-8') as f:
            demo_item_dict = json.load(f)
    except Exception as e:
        logger.error(f"デモファームの辞書読み込みエラー: {e}")
        print(f"  ✗ エラー: デモファームの辞書を読み込めませんでした: {e}")
        return 1
    
    # 全農場を取得
    farms = []
    for item in farms_root.iterdir():
        if item.is_dir() and (item / "farm.db").exists():
            farms.append(item)
    
    if not farms:
        print("  農場が見つかりませんでした")
        return 0
    
    print(f"  見つかった農場数: {len(farms)}")
    print()
    
    updated_farms = []
    skipped_farms = []
    error_farms = []
    
    for farm_path in farms:
        farm_name = farm_path.name
        farm_event_dict_path = farm_path / "event_dictionary.json"
        farm_item_dict_path = farm_path / "item_dictionary.json"
        
        print(f"  処理中: {farm_name}...")
        
        try:
            # 農場側の辞書を読み込む（存在しない場合は空辞書）
            farm_event_dict = {}
            if farm_event_dict_path.exists():
                with open(farm_event_dict_path, 'r', encoding='utf-8') as f:
                    farm_event_dict = json.load(f)
            
            farm_item_dict = {}
            if farm_item_dict_path.exists():
                with open(farm_item_dict_path, 'r', encoding='utf-8') as f:
                    farm_item_dict = json.load(f)
            
            # デモファームの辞書で完全に置き換え
            # （既存のカスタマイズは失われるが、デモファームを基準にする）
            farm_event_dict = demo_event_dict.copy()
            farm_item_dict = demo_item_dict.copy()
            
            # 保存
            with open(farm_event_dict_path, 'w', encoding='utf-8') as f:
                json.dump(farm_event_dict, f, ensure_ascii=False, indent=2)
            
            with open(farm_item_dict_path, 'w', encoding='utf-8') as f:
                json.dump(farm_item_dict, f, ensure_ascii=False, indent=2)
            
            logger.info(f"Farm dictionary updated: {farm_path}")
            updated_farms.append(farm_name)
            print(f"    ✓ 更新しました")
            
        except Exception as e:
            logger.error(f"農場 {farm_name} の更新に失敗: {e}")
            error_farms.append((farm_name, str(e)))
            print(f"    ✗ エラー: {e}")
    
    print()
    print("=" * 60)
    print("完了")
    print("=" * 60)
    print(f"更新された農場数: {len(updated_farms)}")
    if updated_farms:
        print(f"  農場: {', '.join(updated_farms)}")
    
    if error_farms:
        print(f"エラーが発生した農場数: {len(error_farms)}")
        for farm_name, error_msg in error_farms:
            print(f"  {farm_name}: {error_msg}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())








































