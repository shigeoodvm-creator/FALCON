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

from constants import FARMS_ROOT

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='FALCON2 - 牛群管理システム')
    parser.add_argument('--demo-csv', type=str, help='乳検CSVファイルからデモ農場を作成')
    parser.add_argument('--demo-csv2', type=str, help='乳検CSVファイルからデモ農場2を作成（分娩イベント付き）')
    parser.add_argument('--overwrite', action='store_true', help='既存のデモ農場を上書きする')
    parser.add_argument('--set', type=str, help='一括更新（例: ENTR=2024-04-01 : ID>0）')
    parser.add_argument('--farm', type=str, help='農場名（--set使用時は必須）')
    parser.add_argument('--migrate-event-dim-lact', action='store_true', help='全農場のイベントにDIMと産次を追加・再計算')
    parser.add_argument('cow_id', nargs='?', type=str, help='個体ID（4桁の管理番号、例: 0980）')
    
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
    
    # --migrate-event-dim-lact オプションが指定された場合（全農場のイベントにDIMと産次を追加）
    if args.migrate_event_dim_lact:
        sys.path.insert(0, str(Path(__file__).parent / "app"))
        from modules.recalculate_event_dim_lact import migrate_all_farms
        
        try:
            print("全農場のイベントにDIMと産次を追加・再計算中...")
            results = migrate_all_farms()
            
            print("\n=== 処理結果 ===")
            for farm_name, result in results.items():
                if 'error' in result:
                    print(f"{farm_name}: エラー - {result['error']}")
                else:
                    print(f"{farm_name}: 総数={result['total']}, 更新={result['updated']}, スキップ={result['skipped']}, エラー={result['errors']}")
            
            print("\n処理が完了しました。")
            sys.exit(0)
        except Exception as e:
            print(f"エラー: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    
    # --set オプションが指定された場合（一括更新）
    if args.set:
        if not args.farm:
            print("エラー: --set オプションを使用する場合は --farm で農場名を指定してください")
            sys.exit(1)
        
        farm_path = FARMS_ROOT / args.farm
        if not (farm_path / "farm.db").exists():
            print(f"エラー: 農場が見つかりません: {farm_path}")
            sys.exit(1)
        
        # 一括更新を実行
        sys.path.insert(0, str(Path(__file__).parent / "app"))
        from db.db_handler import DBHandler
        from modules.formula_engine import FormulaEngine
        from modules.batch_update import BatchUpdate
        
        try:
            db_handler = DBHandler(farm_path / "farm.db")
            item_dict_path = farm_path / "item_dictionary.json"
            formula_engine = FormulaEngine(db_handler, item_dict_path)
            batch_update = BatchUpdate(db_handler, formula_engine)
            
            # コマンドをパース
            update_items, condition = batch_update.parse_command(args.set)
            
            # 更新を実行
            print(f"一括更新を実行中: {args.set}")
            updated_count = batch_update.execute_update(update_items, condition)
            print(f"一括更新が完了しました: {updated_count}頭を更新")
            
            db_handler.close()
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
    
    # 通常の起動（app/main.py の main を importlib で読み込み、リンターの main と衝突しないようにする）
    import importlib.util
    _spec = importlib.util.spec_from_file_location("app_main", APP_DIR / "main.py")
    if _spec is None or _spec.loader is None:
        raise RuntimeError("app/main.py を読み込めません")
    _app_main = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_app_main)
    _app_main.main(cow_id=args.cow_id if hasattr(args, 'cow_id') else None)

