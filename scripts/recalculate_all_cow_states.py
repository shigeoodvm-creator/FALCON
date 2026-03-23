"""
既存の全牛の状態（lact, clvd, rc, pen）を再計算してDBを更新するスクリプト

このスクリプトは、RuleEngine.apply_events()を使用して全牛の状態を再計算し、
cowテーブルのlact, clvd, rc, penを更新します。
"""
import sys
import sqlite3
from pathlib import Path
import logging

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def recalculate_all_cows(db_path: Path):
    """
    指定されたfarm.dbの全牛の状態を再計算
    
    Args:
        db_path: farm.dbのパス
    """
    # パスを文字列からPathオブジェクトに変換（既にPathの場合はそのまま）
    if isinstance(db_path, str):
        db_path = Path(db_path)
    
    # 絶対パスに変換
    db_path = db_path.resolve()
    
    if not db_path.exists():
        logger.warning(f"スキップ: {db_path} (farm.dbが見つかりません)")
        logger.warning(f"親フォルダ: {db_path.parent}")
        logger.warning(f"親フォルダが存在するか: {db_path.parent.exists()}")
        return 0, 0
    
    logger.info(f"処理中: {db_path.parent.name}")
    
    try:
        db = DBHandler(db_path)
        rule_engine = RuleEngine(db)
        
        # 全牛を取得
        all_cows = db.get_all_cows()
        logger.info(f"  牛数: {len(all_cows)}")
        
        updated_count = 0
        error_count = 0
        
        for cow in all_cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            try:
                # 牛の状態を再計算してcowテーブルを更新
                rule_engine._recalculate_and_update_cow(cow_auto_id)
                updated_count += 1
                
                # 進捗表示（100頭ごと）
                if updated_count % 100 == 0:
                    logger.info(f"  進捗: {updated_count}/{len(all_cows)}頭を処理")
                    
            except Exception as e:
                error_count += 1
                logger.error(f"  エラー (cow_auto_id={cow_auto_id}, cow_id={cow.get('cow_id')}): {e}", exc_info=True)
        
        logger.info(f"  完了: {updated_count}頭を更新, {error_count}件のエラー")
        db.close()
        
        return updated_count, error_count
        
    except Exception as e:
        logger.error(f"  エラー: {e}", exc_info=True)
        return 0, 1

def update_all_farms():
    """全デモファームの牛の状態を再計算"""
    farms = [
        Path("C:/FARMS/デモファーム"),
        Path("C:/FARMS/DemoFarm"),
        Path("C:/FARMS/DemoFarm2"),
        Path("C:/FARMS/DemoFarm3"),
        Path("C:/FARMS/DemoFarm4"),
    ]
    
    total_updated = 0
    total_errors = 0
    
    for farm_path in farms:
        db_path = farm_path / "farm.db"
        updated, errors = recalculate_all_cows(db_path)
        total_updated += updated
        total_errors += errors
        print()
    
    return total_updated, total_errors

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="既存の全牛の状態（lact, clvd, rc, pen）を再計算してDBを更新"
    )
    parser.add_argument(
        "--farm-db",
        type=str,
        help="更新するfarm.dbのパス（指定しない場合は全デモファームを更新）"
    )
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("既存の全牛の状態（lact, clvd, rc, pen）を再計算してDBを更新")
    print("=" * 80)
    print()
    print("このスクリプトは以下を更新します:")
    print("  - lact (産次)")
    print("  - clvd (最終分娩日)")
    print("  - rc (繁殖コード)")
    print("  - pen (群)")
    print()
    
    if args.farm_db:
        # 指定されたfarm.dbのみを更新
        # パスを正しく処理するため、文字列から直接Pathオブジェクトを作成
        farm_db_str = args.farm_db
        farm_db_path = Path(farm_db_str)
        # 絶対パスに変換
        farm_db_path = farm_db_path.resolve()
        print(f"指定されたfarm.dbを更新: {farm_db_path}")
        print(f"パスが存在するか: {farm_db_path.exists()}")
        print()
        total_updated, total_errors = recalculate_all_cows(farm_db_path)
        print()
        print("=" * 80)
        print(f"更新完了: {total_updated}頭を更新, {total_errors}件のエラー")
        print("=" * 80)
    else:
        # 全デモファームを更新
        print("全デモファームを更新します...")
        print()
        total_updated, total_errors = update_all_farms()
        print()
        print("=" * 80)
        print(f"更新完了: {total_updated}頭を更新, {total_errors}件のエラー")
        print("=" * 80)

