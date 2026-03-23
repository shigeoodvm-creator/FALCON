"""
FALCON2 - Migration: event_lact と event_dim カラムを追加

Phase A: イベントごとに event_lact / event_dim をDBに保持
"""

import sys
import logging
from pathlib import Path
from datetime import datetime

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
    """
    conn = db_handler.connect()
    cursor = conn.cursor()
    
    # event_lact カラムを追加
    if not check_column_exists(db_handler, "event", "event_lact"):
        logger.info("Adding event_lact column to event table...")
        cursor.execute("ALTER TABLE event ADD COLUMN event_lact INTEGER NULL")
        conn.commit()
        logger.info("event_lact column added successfully")
    else:
        logger.info("event_lact column already exists, skipping")
    
    # event_dim カラムを追加
    if not check_column_exists(db_handler, "event", "event_dim"):
        logger.info("Adding event_dim column to event table...")
        cursor.execute("ALTER TABLE event ADD COLUMN event_dim INTEGER NULL")
        conn.commit()
        logger.info("event_dim column added successfully")
    else:
        logger.info("event_dim column already exists, skipping")


def migrate_all_cows(db_handler: DBHandler, rule_engine: RuleEngine):
    """
    全牛のイベントを再計算して event_lact/event_dim を更新
    
    Args:
        db_handler: DBHandler インスタンス
        rule_engine: RuleEngine インスタンス
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


def main():
    """メイン処理"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Add event_lact and event_dim columns to event table')
    parser.add_argument('db_path', type=str, help='Path to farm.db file')
    parser.add_argument('--skip-recalc', action='store_true', 
                       help='Skip recalculation of existing events (only add columns)')
    
    args = parser.parse_args()
    
    db_path = Path(args.db_path)
    if not db_path.exists():
        logger.error(f"Database file not found: {db_path}")
        sys.exit(1)
    
    logger.info(f"Starting migration for: {db_path}")
    
    try:
        # DBHandlerとRuleEngineを初期化
        db_handler = DBHandler(db_path)
        rule_engine = RuleEngine(db_handler)
        
        # カラムを追加
        add_columns_if_not_exists(db_handler)
        
        # 既存データの再計算（スキップオプションが無い場合）
        if not args.skip_recalc:
            logger.info("Recalculating event_lact and event_dim for all existing events...")
            migrate_all_cows(db_handler, rule_engine)
        else:
            logger.info("Skipping recalculation (--skip-recalc specified)")
        
        logger.info("Migration completed successfully")
        
    except Exception as e:
        logger.error(f"Migration failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()


















