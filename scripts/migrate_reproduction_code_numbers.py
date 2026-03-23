"""
既存のDB内の繁殖コード（RC）の番号を新しい定義に置き換えるスクリプト

変更マッピング:
- 1 (RC_FRESH) → 2
- 2 (RC_BRED) → 3
- 3 (RC_PREGNANT) → 5
- 4 (RC_DRY) → 6
- 5 (RC_OPEN) → 4
- 6 (RC_STOPPED) → 1

このスクリプトは、cowテーブルのrcカラムの値を新しい番号に置き換えます。
"""
import sys
import sqlite3
from pathlib import Path
import logging

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 旧コードから新コードへのマッピング
RC_MAPPING = {
    1: 2,  # RC_FRESH: 1 → 2
    2: 3,  # RC_BRED: 2 → 3
    3: 5,  # RC_PREGNANT: 3 → 5
    4: 6,  # RC_DRY: 4 → 6
    5: 4,  # RC_OPEN: 5 → 4
    6: 1,  # RC_STOPPED: 6 → 1
}

def migrate_rc_codes(db_path: Path):
    """
    指定されたfarm.dbの繁殖コードを新しい番号に置き換え
    
    Args:
        db_path: farm.dbのパス
    """
    if not db_path.exists():
        logger.warning(f"スキップ: {db_path} (farm.dbが見つかりません)")
        return 0, 0
    
    logger.info(f"処理中: {db_path.parent.name}")
    
    try:
        db = DBHandler(db_path)
        conn = db.connect()
        cursor = conn.cursor()
        
        # 現在のrcの分布を確認
        cursor.execute("SELECT rc, COUNT(*) as count FROM cow GROUP BY rc")
        current_distribution = {row[0]: row[1] for row in cursor.fetchall() if row[0] is not None}
        logger.info(f"  現在のRC分布: {current_distribution}")
        
        updated_count = 0
        error_count = 0
        
        # 各マッピングを適用
        for old_rc, new_rc in RC_MAPPING.items():
            try:
                cursor.execute("""
                    UPDATE cow 
                    SET rc = ?
                    WHERE rc = ?
                """, (new_rc, old_rc))
                
                rows_affected = cursor.rowcount
                if rows_affected > 0:
                    updated_count += rows_affected
                    logger.info(f"  RC {old_rc} → {new_rc}: {rows_affected}頭を更新")
                    
            except Exception as e:
                error_count += 1
                logger.error(f"  RC {old_rc} → {new_rc} の更新エラー: {e}", exc_info=True)
        
        conn.commit()
        logger.info(f"  完了: {updated_count}頭を更新, {error_count}件のエラー")
        db.close()
        
        return updated_count, error_count
        
    except Exception as e:
        logger.error(f"  エラー: {e}", exc_info=True)
        return 0, 1

def update_all_farms():
    """全デモファームの繁殖コードを更新"""
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
        updated, errors = migrate_rc_codes(db_path)
        total_updated += updated
        total_errors += errors
        print()
    
    return total_updated, total_errors

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="既存のDB内の繁殖コード（RC）の番号を新しい定義に置き換え"
    )
    parser.add_argument(
        "--farm-db",
        type=str,
        help="更新するfarm.dbのパス（指定しない場合は全デモファームを更新）"
    )
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("既存のDB内の繁殖コード（RC）の番号を新しい定義に置き換え")
    print("=" * 80)
    print()
    print("変更マッピング:")
    print("  1 (RC_FRESH) → 2")
    print("  2 (RC_BRED) → 3")
    print("  3 (RC_PREGNANT) → 5")
    print("  4 (RC_DRY) → 6")
    print("  5 (RC_OPEN) → 4")
    print("  6 (RC_STOPPED) → 1")
    print()
    
    if args.farm_db:
        # 指定されたfarm.dbのみを更新
        farm_db_path = Path(args.farm_db)
        print(f"指定されたfarm.dbを更新: {farm_db_path}")
        print()
        total_updated, total_errors = migrate_rc_codes(farm_db_path)
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











