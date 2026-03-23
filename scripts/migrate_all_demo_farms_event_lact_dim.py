"""
C:/FARMS内のすべてのデモファームに対して
event_lact と event_dim カラムを追加し、値を計算して設定する

既存のデータを壊さないように安全に実行します。
"""
import sys
import logging
from pathlib import Path

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def check_column_exists(db_handler: DBHandler, table_name: str, column_name: str) -> bool:
    """
    カラムが存在するかチェック
    
    Args:
        db_handler: DBHandler インスタンス
        table_name: テーブル名
        column_name: カラム名
    
    Returns:
        存在する場合はTrue
    """
    conn = db_handler.connect()
    cursor = conn.cursor()
    
    # SQLiteのスキーマ情報を取得
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = cursor.fetchall()
    
    for col in columns:
        if col[1] == column_name:  # col[1]はカラム名
            return True
    
    return False


def add_columns_if_not_exists(db_handler: DBHandler):
    """
    event_lact と event_dim カラムを追加（存在しない場合のみ）
    
    Args:
        db_handler: DBHandler インスタンス
    
    Returns:
        (event_lact_added, event_dim_added) のタプル
    """
    conn = db_handler.connect()
    cursor = conn.cursor()
    
    event_lact_added = False
    event_dim_added = False
    
    # event_lact カラムを追加
    if not check_column_exists(db_handler, "event", "event_lact"):
        logger.info("Adding event_lact column to event table...")
        cursor.execute("ALTER TABLE event ADD COLUMN event_lact INTEGER NULL")
        conn.commit()
        logger.info("event_lact column added successfully")
        event_lact_added = True
    else:
        logger.info("event_lact column already exists, skipping")
    
    # event_dim カラムを追加
    if not check_column_exists(db_handler, "event", "event_dim"):
        logger.info("Adding event_dim column to event table...")
        cursor.execute("ALTER TABLE event ADD COLUMN event_dim INTEGER NULL")
        conn.commit()
        logger.info("event_dim column added successfully")
        event_dim_added = True
    else:
        logger.info("event_dim column already exists, skipping")
    
    return (event_lact_added, event_dim_added)


def migrate_all_cows(db_handler: DBHandler, rule_engine: RuleEngine):
    """
    全牛のイベントを再計算して event_lact/event_dim を更新
    
    Args:
        db_handler: DBHandler インスタンス
        rule_engine: RuleEngine インスタンス
    
    Returns:
        (updated_count, error_count) のタプル
    """
    conn = db_handler.connect()
    cursor = conn.cursor()
    
    # 全牛のauto_idを取得
    cursor.execute("SELECT auto_id FROM cow ORDER BY auto_id")
    cows = cursor.fetchall()
    
    total_cows = len(cows)
    logger.info(f"Found {total_cows} cows to migrate")
    
    updated_count = 0
    error_count = 0
    
    for i, (cow_auto_id,) in enumerate(cows, 1):
        try:
            logger.info(f"Processing cow auto_id={cow_auto_id} ({i}/{total_cows})...")
            rule_engine.recalculate_events_for_cow(cow_auto_id)
            updated_count += 1
        except Exception as e:
            error_count += 1
            logger.error(f"Error processing cow auto_id={cow_auto_id}: {e}", exc_info=True)
    
    logger.info(f"Migration completed: {updated_count} cows updated, {error_count} errors")
    return (updated_count, error_count)


def migrate_farm(farm_path: Path):
    """
    単一の農場に対してマイグレーションを実行
    
    Args:
        farm_path: 農場フォルダのパス
    
    Returns:
        (success, stats) のタプル。statsは辞書形式
    """
    db_path = farm_path / "farm.db"
    
    if not db_path.exists():
        logger.warning(f"Database file not found: {db_path}, skipping")
        return (False, {"error": "Database not found"})
    
    logger.info("=" * 80)
    logger.info(f"Processing farm: {farm_path.name}")
    logger.info(f"Database: {db_path}")
    logger.info("=" * 80)
    
    try:
        # DBHandlerとRuleEngineを初期化
        db_handler = DBHandler(db_path)
        rule_engine = RuleEngine(db_handler)
        
        # カラムを追加
        event_lact_added, event_dim_added = add_columns_if_not_exists(db_handler)
        
        # 既存データの再計算
        logger.info("Recalculating event_lact and event_dim for all existing events...")
        updated_count, error_count = migrate_all_cows(db_handler, rule_engine)
        
        stats = {
            "event_lact_added": event_lact_added,
            "event_dim_added": event_dim_added,
            "cows_updated": updated_count,
            "cows_errors": error_count
        }
        
        logger.info(f"Farm {farm_path.name} migration completed successfully")
        logger.info(f"Stats: {stats}")
        
        return (True, stats)
        
    except Exception as e:
        logger.error(f"Migration failed for {farm_path.name}: {e}", exc_info=True)
        return (False, {"error": str(e)})


def main():
    """メイン処理"""
    farms_root = Path("C:/FARMS")
    
    if not farms_root.exists():
        logger.error(f"Farms directory not found: {farms_root}")
        sys.exit(1)
    
    # デモファームを検索
    demo_farms = []
    for item in farms_root.iterdir():
        if item.is_dir():
            db_path = item / "farm.db"
            if db_path.exists():
                demo_farms.append(item)
    
    if not demo_farms:
        logger.warning("No demo farms found in C:/FARMS")
        sys.exit(0)
    
    logger.info(f"Found {len(demo_farms)} demo farms to migrate")
    logger.info("Farms: " + ", ".join(f.name for f in demo_farms))
    
    # 各デモファームに対してマイグレーションを実行
    results = []
    for farm_path in sorted(demo_farms):
        success, stats = migrate_farm(farm_path)
        results.append({
            "farm": farm_path.name,
            "success": success,
            "stats": stats
        })
    
    # 結果サマリー
    logger.info("\n" + "=" * 80)
    logger.info("Migration Summary")
    logger.info("=" * 80)
    
    for result in results:
        farm_name = result["farm"]
        success = result["success"]
        stats = result["stats"]
        
        if success:
            logger.info(f"{farm_name}: SUCCESS")
            logger.info(f"  - Cows updated: {stats.get('cows_updated', 0)}")
            logger.info(f"  - Cows errors: {stats.get('cows_errors', 0)}")
            logger.info(f"  - event_lact added: {stats.get('event_lact_added', False)}")
            logger.info(f"  - event_dim added: {stats.get('event_dim_added', False)}")
        else:
            logger.error(f"{farm_name}: FAILED")
            logger.error(f"  - Error: {stats.get('error', 'Unknown error')}")
    
    # 成功/失敗のカウント
    success_count = sum(1 for r in results if r["success"])
    failed_count = len(results) - success_count
    
    logger.info("=" * 80)
    logger.info(f"Total: {len(results)} farms, {success_count} succeeded, {failed_count} failed")
    logger.info("=" * 80)


if __name__ == "__main__":
    main()


















