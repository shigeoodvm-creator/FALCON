"""
FALCON2 - イベント時のDIMと産次計算モジュール
イベント作成・更新時に、そのイベント時点でのDIMと産次を計算して保存
"""

from typing import Dict, Any, Optional, Tuple
from datetime import datetime
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)


class EventDimLactCalculator:
    """イベント時のDIMと産次を計算するクラス"""
    
    def __init__(self, db_handler: DBHandler):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
        """
        self.db = db_handler
        self.rule_engine = RuleEngine(db_handler)
    
    def calculate_event_dim_and_lact(
        self,
        cow_auto_id: int,
        event_date: str,
        event_number: int,
        exclude_event_id: Optional[int] = None
    ) -> Tuple[Optional[int], Optional[int]]:
        """
        イベント時点でのDIMと産次を計算
        
        Args:
            cow_auto_id: 牛のauto_id
            event_date: イベント日（YYYY-MM-DD形式）
            event_number: イベント番号
            exclude_event_id: 計算から除外するイベントID（更新時など）
        
        Returns:
            (event_dim, event_lact) のタプル
            DIMが計算できない場合はNone、産次が計算できない場合はNone
        """
        try:
            # 牛データを取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                return None, None
            
            # イベント日時点までのイベント履歴を取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            # イベント日時点までのイベントのみをフィルタ（除外するイベントIDも除外）
            event_date_dt = datetime.strptime(event_date, '%Y-%m-%d')
            filtered_events = [
                e for e in events
                if datetime.strptime(e.get('event_date', ''), '%Y-%m-%d') < event_date_dt
                or (datetime.strptime(e.get('event_date', ''), '%Y-%m-%d') == event_date_dt
                    and (exclude_event_id is None or e.get('id') != exclude_event_id))
            ]
            
            # イベント日でソート（昇順：古い順）
            filtered_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
            
            # 初期状態を設定
            state = {
                'lact': cow.get('lact') or 0,
                'clvd': cow.get('clvd'),
                'rc': cow.get('rc') or self.rule_engine.RC_OPEN,
                'last_ai_date': None,
                'conception_date': None,
                'due_date': None,
                'dry_date': None,
                'pen': cow.get('pen'),
                'current_lact_clvd': None,  # 現在の産次での分娩日
                'cow_auto_id': cow_auto_id
            }
            
            # イベント日時点までのイベントを適用して状態を計算
            for event in filtered_events:
                self.rule_engine._apply_event(state, event)
            
            # 産次を取得
            event_lact = state.get('lact')
            
            # DIMを計算（現在の産次での分娩日から）
            event_dim = None
            current_lact_clvd = state.get('current_lact_clvd')
            if current_lact_clvd:
                try:
                    clvd_dt = datetime.strptime(current_lact_clvd, '%Y-%m-%d')
                    dim = (event_date_dt - clvd_dt).days
                    event_dim = dim if dim >= 0 else None
                except (ValueError, TypeError):
                    event_dim = None
            
            return event_dim, event_lact
            
        except Exception as e:
            logger.error(f"イベントDIM/産次計算エラー: cow_auto_id={cow_auto_id}, event_date={event_date}, エラー={e}", exc_info=True)
            return None, None
    
    def should_recalculate(
        self,
        event: Dict[str, Any],
        new_event_date: Optional[str] = None,
        new_event_number: Optional[int] = None
    ) -> bool:
        """
        イベントの再計算が必要かどうかを判定
        
        次の産次に繰り上がった場合は除くが、当該産次内で分娩日が変更になったり、
        イベント日そのものが変更された場合は再計算が必要
        
        Args:
            event: 既存のイベントデータ
            new_event_date: 新しいイベント日（更新時）
            new_event_number: 新しいイベント番号（更新時）
        
        Returns:
            再計算が必要な場合はTrue
        """
        # イベント日が変更された場合は再計算
        if new_event_date and new_event_date != event.get('event_date'):
            return True
        
        # イベント番号が変更された場合は再計算
        if new_event_number and new_event_number != event.get('event_number'):
            return True
        
        # 既存イベントでevent_dimやevent_lactがNULLの場合は再計算
        if event.get('event_dim') is None or event.get('event_lact') is None:
            return True
        
        # その他の場合は、イベント日や分娩日の変更を検知する必要がある
        # これはRuleEngineの再計算時に検知される
        return False

