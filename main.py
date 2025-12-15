"""
FALCON2 - 牛群管理システム
エントリーポイント（ラッパー）
"""

import sys
import argparse
from pathlib import Path

# app ディレクトリを sys.path に追加
APP_DIR = Path(__file__).parent / "app"
sys.path.insert(0, str(APP_DIR))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='FALCON2 - 牛群管理システム')
    parser.add_argument('--demo-csv', type=str, help='乳検CSVファイルからデモ農場を作成')
    parser.add_argument('--demo-csv2', type=str, help='乳検CSVファイルからデモ農場2を作成（分娩イベント付き）')
    parser.add_argument('--overwrite', action='store_true', help='既存のデモ農場を上書きする')
    
    args = parser.parse_args()
    
    # --demo-csv2 オプションが指定された場合
    if args.demo_csv2:
        csv_path = Path(args.demo_csv2)
        if not csv_path.exists():
            print(f"エラー: CSVファイルが見つかりません: {csv_path}")
            sys.exit(1)
        
        # デモ農場2作成スクリプトを実行
        sys.path.insert(0, str(Path(__file__).parent / "scripts"))
        from create_demo_farm2_from_csv import create_demo_farm2
        
        try:
            create_demo_farm2(csv_path)
            print("デモ農場2の作成が完了しました。")
            sys.exit(0)
        except Exception as e:
            print(f"エラー: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    
    # --demo-csv オプションが指定された場合
    if args.demo_csv:
        csv_path = Path(args.demo_csv)
        if not csv_path.exists():
            print(f"エラー: CSVファイルが見つかりません: {csv_path}")
            sys.exit(1)
        
        # デモ農場作成スクリプトを実行
        sys.path.insert(0, str(Path(__file__).parent / "scripts"))
        from create_demo_farm_from_milk_csv import create_demo_farm
        
        try:
            create_demo_farm(csv_path, overwrite=args.overwrite)
            print("デモ農場の作成が完了しました。")
            sys.exit(0)
        except Exception as e:
            print(f"エラー: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    
    # 通常の起動
    print("DEBUG: C:\\FALCON\\main.py から起動")
    
    # 初回起動時、DemoFarmが存在しない場合に自動実行（オプション）
    # この機能は必要に応じて有効化
    AUTO_CREATE_DEMO = False  # デフォルトは無効
    
    if AUTO_CREATE_DEMO:
        demo_farm_path = Path("C:/FARMS/DemoFarm")
        if not (demo_farm_path / "farm.db").exists():
            print("DemoFarmが存在しないため、自動作成をスキップします。")
            print("デモ農場を作成するには、以下のコマンドを実行してください:")
            print("  python main.py --demo-csv <CSVファイルのパス>")
    
    from main import main
    main()

