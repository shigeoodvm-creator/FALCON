"""
FALCON2 - 乳検CSVからデモ農場4を作成
CSVファイルを読み込んで、DemoFarm4を作成し、牛・分娩イベント・乳検イベントを登録する

【仕様】
- 産次はCSVの値をそのまま使用（baseline）
- 分娩イベントはbaseline_calvingフラグ付きで作成（産次を増やさない）
- 指定された列位置から正確にデータを読み込む
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
    import argparse
    
    parser = argparse.ArgumentParser(description='乳検CSVからデモ農場4を作成')
    parser.add_argument('csv_path', type=str, nargs='?', help='CSVファイルのパス')
    
    args = parser.parse_args()
    
    # CSVパスが指定されていない場合はデフォルトパスを使用
    if args.csv_path:
        csv_path = Path(args.csv_path)
    else:
        # デフォルトパスを試す
        default_paths = [
            Path(r"c:\Users\user\Downloads\検定日速報_1219102_202511.csv"),
            Path("temp_milk_test.csv"),
        ]
        csv_path = None
        for path in default_paths:
            if path.exists():
                csv_path = path
                break
        
        if csv_path is None:
            print("エラー: CSVファイルが見つかりません。")
            print("使用可能なパス:")
            for path in default_paths:
                print(f"  {path}")
            return 1
    
    if not csv_path.exists():
        print(f"エラー: CSVファイルが見つかりません: {csv_path}")
        return 1
    
    print(f"CSVファイル: {csv_path}")
    
    try:
        # FarmCreatorを使用してデモ農場4を作成
        farm_path = Path("C:/FARMS/DemoFarm4")
        print(f"農場パス: {farm_path}")
        
        creator = FarmCreator(farm_path)
        creator.create_farm(
            farm_name="DemoFarm4",
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









































