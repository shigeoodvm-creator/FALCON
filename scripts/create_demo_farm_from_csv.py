"""
FALCON2 - 乳検CSVからデモ農場を作成
"""

import sys
from pathlib import Path

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT))
sys.path.insert(0, str(APP_ROOT / "app"))

from app.modules.farm_creator import FarmCreator


def main():
    """メイン処理"""
    csv_path = Path(r"c:\Users\user\Downloads\検定日速報_1219102_202511.csv")
    
    if not csv_path.exists():
        print(f"エラー: CSVファイルが見つかりません: {csv_path}")
        return 1
    
    print(f"CSVファイル: {csv_path}")
    
    try:
        # FarmCreatorを使用してデモ農場を作成
        farm_path = Path("C:/FARMS/デモファーム")
        print(f"農場パス: {farm_path}")
        
        creator = FarmCreator(farm_path)
        creator.create_farm(
            farm_name="デモファーム",
            csv_path=csv_path,
            template_farm_path=None
        )
        
        print(f"\n完了:")
        print(f"  農場パス: {farm_path}")
        return 0
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())









































