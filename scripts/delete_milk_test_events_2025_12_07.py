"""
FALCON2 - 2025-12-07の乳検イベントを削除
再検証のため、指定日付のすべての個体の乳検イベントを削除する
"""

import sys
import logging
from pathlib import Path
from typing import List, Dict, Any

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def delete_milk_test_events_by_date(farm_path: Path, event_date: str) -> Dict[str, Any]:
    """
    指定日付の乳検イベントを削除
    
    Args:
        farm_path: 農場フォルダのパス
        event_date: 削除するイベントの日付（YYYY-MM-DD形式）
    
    Returns:
        削除結果の辞書
    """
    db_path = farm_path / "farm.db"
    
    if not db_path.exists():
        logger.warning(f"データベースが見つかりません: {db_path}")
        return {'error': f'データベースが見つかりません: {db_path}'}
    
    db = DBHandler(db_path)
    rule_engine = RuleEngine(db)
    
    # 指定日付の乳検イベントを取得
    conn = db.connect()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT id, cow_auto_id 
        FROM event 
        WHERE event_number = ? 
        AND event_date = ? 
        AND deleted = 0
    """, (RuleEngine.EVENT_MILK_TEST, event_date))
    
    events = cursor.fetchall()
    event_list = [dict(row) for row in events]
    
    logger.info(f"削除対象のイベント数: {len(event_list)}")
    
    deleted_count = 0
    error_count = 0
    affected_cows = set()
    
    # 各イベントを削除
    for event in event_list:
        event_id = event['id']
        cow_auto_id = event['cow_auto_id']
        
        try:
            # 論理削除
            db.delete_event(event_id, soft_delete=True)
            
            # RuleEngineで状態を更新
            rule_engine.on_event_deleted(event_id)
            
            deleted_count += 1
            affected_cows.add(cow_auto_id)
            
        except Exception as e:
            error_count += 1
            logger.error(f"イベント削除エラー: event_id={event_id}, エラー={e}")
    
    # 影響を受けた牛の状態を再計算
    for cow_auto_id in affected_cows:
        try:
            rule_engine._recalculate_and_update_cow(cow_auto_id)
        except Exception as e:
            logger.error(f"牛の状態再計算エラー: cow_auto_id={cow_auto_id}, エラー={e}")
    
    db.close()
    
    result = {
        'farm_name': farm_path.name,
        'event_date': event_date,
        'deleted_count': deleted_count,
        'error_count': error_count,
        'affected_cows': len(affected_cows)
    }
    
    logger.info(f"削除完了: {result}")
    return result


def main():
    """メイン処理"""
    import argparse
    
    parser = argparse.ArgumentParser(description='2025-12-07の乳検イベントを削除')
    parser.add_argument('--farm-path', type=str, help='農場フォルダのパス（指定しない場合は全農場）')
    parser.add_argument('--date', type=str, default='2025-12-07', help='削除する日付（デフォルト: 2025-12-07）')
    
    args = parser.parse_args()
    
    event_date = args.date
    
    if args.farm_path:
        # 指定された農場のみ
        farm_path = Path(args.farm_path)
        if not farm_path.exists():
            print(f"エラー: 農場フォルダが見つかりません: {farm_path}")
            return 1
        
        result = delete_milk_test_events_by_date(farm_path, event_date)
        print(f"\n結果:")
        print(f"  農場: {result.get('farm_name')}")
        print(f"  削除日付: {result.get('event_date')}")
        print(f"  削除件数: {result.get('deleted_count')}")
        print(f"  エラー件数: {result.get('error_count')}")
        print(f"  影響を受けた牛数: {result.get('affected_cows')}")
    else:
        # 全農場を対象
        farms_root = Path("C:/FARMS")
        
        if not farms_root.exists():
            print(f"エラー: 農場フォルダが見つかりません: {farms_root}")
            return 1
        
        print(f"全農場の {event_date} の乳検イベントを削除します...")
        print()
        
        results = []
        for farm_dir in farms_root.iterdir():
            if not farm_dir.is_dir():
                continue
            
            db_path = farm_dir / "farm.db"
            if not db_path.exists():
                continue
            
            print(f"処理中: {farm_dir.name}")
            result = delete_milk_test_events_by_date(farm_dir, event_date)
            results.append(result)
        
        print("\n" + "=" * 60)
        print("削除完了")
        print("=" * 60)
        
        total_deleted = sum(r.get('deleted_count', 0) for r in results)
        total_errors = sum(r.get('error_count', 0) for r in results)
        total_cows = sum(r.get('affected_cows', 0) for r in results)
        
        print(f"対象農場数: {len(results)}")
        print(f"合計削除件数: {total_deleted}")
        print(f"合計エラー件数: {total_errors}")
        print(f"合計影響を受けた牛数: {total_cows}")
        print()
        
        for result in results:
            if result.get('deleted_count', 0) > 0:
                print(f"  {result.get('farm_name')}: {result.get('deleted_count')}件削除")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())







