"""
FALCON2 - 既存イベントのDIMと産次再計算モジュール
既存の全イベントに対してDIMと産次を計算・更新
"""

import logging
from pathlib import Path
from typing import List, Dict, Any
from db.db_handler import DBHandler
from modules.event_dim_lact_calculator import EventDimLactCalculator
from constants import FARMS_ROOT

logger = logging.getLogger(__name__)


class RecalculateEventDimLact:
    """既存イベントのDIMと産次を再計算するクラス"""
    
    def __init__(self, db_handler: DBHandler):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
        """
        self.db = db_handler
        self.calculator = EventDimLactCalculator(db_handler)
    
    def recalculate_all_events(self, force: bool = False) -> Dict[str, int]:
        """
        全イベントのDIMと産次を再計算
        
        Args:
            force: Trueの場合は、既に値があるイベントも再計算
        
        Returns:
            統計情報（total, updated, skipped, errors）
        """
        # 全イベントを取得
        conn = self.db.connect()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT id, cow_auto_id, event_number, event_date, event_dim, event_lact
            FROM event
            WHERE deleted = 0
            ORDER BY cow_auto_id, event_date, id
        """)
        
        events = [dict(row) for row in cursor.fetchall()]
        
        total = len(events)
        updated = 0
        skipped = 0
        errors = 0
        
        logger.info(f"全イベントのDIM/産次再計算を開始: {total}件")
        
        for event in events:
            event_id = event.get('id')
            cow_auto_id = event.get('cow_auto_id')
            event_date = event.get('event_date')
            event_number = event.get('event_number')
            existing_dim = event.get('event_dim')
            existing_lact = event.get('event_lact')
            
            # 既に値がある場合はスキップ（force=Trueの場合は再計算）
            if not force and existing_dim is not None and existing_lact is not None:
                skipped += 1
                continue
            
            try:
                # DIMと産次を計算（このイベント自体を含めて計算）
                event_dim, event_lact = self.calculator.calculate_event_dim_and_lact(
                    cow_auto_id, event_date, event_number, exclude_event_id=None
                )
                
                # 更新
                cursor.execute("""
                    UPDATE event
                    SET event_dim = ?, event_lact = ?
                    WHERE id = ?
                """, (event_dim, event_lact, event_id))
                
                updated += 1
                
                if updated % 100 == 0:
                    logger.info(f"進捗: {updated}/{total}件を更新")
                    conn.commit()
                    
            except Exception as e:
                errors += 1
                logger.error(f"イベント再計算エラー: event_id={event_id}, エラー={e}", exc_info=True)
        
        conn.commit()
        
        result = {
            'total': total,
            'updated': updated,
            'skipped': skipped,
            'errors': errors
        }
        
        logger.info(f"全イベントのDIM/産次再計算が完了: {result}")
        return result
    
    def recalculate_events_for_cow(self, cow_auto_id: int, force: bool = False) -> Dict[str, int]:
        """
        特定の牛の全イベントのDIMと産次を再計算
        
        Args:
            cow_auto_id: 牛のauto_id
            force: Trueの場合は、既に値があるイベントも再計算
        
        Returns:
            統計情報（total, updated, skipped, errors）
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        total = len(events)
        updated = 0
        skipped = 0
        errors = 0
        
        conn = self.db.connect()
        cursor = conn.cursor()
        
        for event in events:
            event_id = event.get('id')
            event_date = event.get('event_date')
            event_number = event.get('event_number')
            existing_dim = event.get('event_dim')
            existing_lact = event.get('event_lact')
            
            # 既に値がある場合はスキップ（force=Trueの場合は再計算）
            if not force and existing_dim is not None and existing_lact is not None:
                skipped += 1
                continue
            
            try:
                # DIMと産次を計算（このイベント自体を含めて計算）
                event_dim, event_lact = self.calculator.calculate_event_dim_and_lact(
                    cow_auto_id, event_date, event_number, exclude_event_id=None
                )
                
                # 更新
                cursor.execute("""
                    UPDATE event
                    SET event_dim = ?, event_lact = ?
                    WHERE id = ?
                """, (event_dim, event_lact, event_id))
                
                updated += 1
                    
            except Exception as e:
                errors += 1
                logger.error(f"イベント再計算エラー: event_id={event_id}, エラー={e}", exc_info=True)
        
        conn.commit()
        
        result = {
            'total': total,
            'updated': updated,
            'skipped': skipped,
            'errors': errors
        }
        
        return result


def migrate_all_farms():
    """
    全農場のeventテーブルにevent_dimとevent_lactカラムを追加し、
    既存イベントのDIMと産次を再計算
    
    Returns:
        処理結果の辞書
    """
    farms_root = FARMS_ROOT
    if not farms_root.exists():
        logger.warning(f"農場フォルダが見つかりません: {farms_root}")
        return {}
    
    results = {}
    
    # 全農場を処理
    for farm_dir in farms_root.iterdir():
        if not farm_dir.is_dir():
            continue
        
        db_path = farm_dir / "farm.db"
        if not db_path.exists():
            continue
        
        farm_name = farm_dir.name
        logger.info(f"農場 '{farm_name}' の処理を開始")
        
        try:
            # データベースに接続
            db = DBHandler(db_path)
            
            # テーブル作成（カラム追加も含む）
            db.create_tables()
            
            # 既存イベントの再計算
            recalculator = RecalculateEventDimLact(db)
            result = recalculator.recalculate_all_events(force=False)
            
            results[farm_name] = result
            logger.info(f"農場 '{farm_name}' の処理が完了: {result}")
            
            db.close()
            
        except Exception as e:
            logger.error(f"農場 '{farm_name}' の処理でエラー: {e}", exc_info=True)
            results[farm_name] = {'error': str(e)}
    
    return results








