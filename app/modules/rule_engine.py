"""
FALCON2 - RuleEngine
イベント履歴から牛の状態を完全再計算する
設計書 第5章参照
"""

from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from pathlib import Path
import json

from db.db_handler import DBHandler


class RuleEngine:
    """イベント履歴から牛の状態を再計算するエンジン"""
    
    # イベント番号定義（設計書 第3章参照）
    EVENT_AI = 200          # AI（人工授精）
    EVENT_ET = 201          # ET（胚移植）
    EVENT_CALV = 202        # 分娩
    EVENT_DRY = 203         # 乾乳
    EVENT_STOPR = 204       # 繁殖停止
    EVENT_SOLD = 205        # 売却
    EVENT_DEAD = 206        # 死亡・淘汰
    
    EVENT_PDN = 302         # 妊娠鑑定マイナス
    EVENT_PDP = 303         # 妊娠鑑定プラス
    EVENT_PDP2 = 304        # 妊娠鑑定プラス（検診以外）
    EVENT_ABRT = 305        # 流産
    EVENT_PAGN = 306        # PAGマイナス
    EVENT_PAGP = 307        # PAGプラス
    
    EVENT_MILK_TEST = 601   # 乳検
    EVENT_MOVE = 611        # 群変更
    
    # 繁殖コード（RC）定義
    RC_FRESH = 1            # Fresh（分娩後）
    RC_BRED = 2             # Bred（授精後）
    RC_PREGNANT = 3         # Pregnant（妊娠中）
    RC_DRY = 4              # Dry（乾乳中）
    RC_OPEN = 5             # Open（空胎）
    RC_STOPPED = 6          # Stopped（繁殖停止）
    
    def __init__(self, db_handler: DBHandler):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
        """
        self.db = db_handler
    
    def apply_events(self, cow_auto_id: int) -> Dict[str, Any]:
        """
        牛の全イベント履歴を適用して状態を再計算
        
        【重要】差分更新は禁止。常に全イベント履歴から完全再計算する
        
        Args:
            cow_auto_id: 牛の auto_id
        
        Returns:
            計算された状態（lact, clvd, rc, last_ai_date, conception_date, due_date, dry_date, pen）
        """
        # 1. baseline（cowテーブルの初期値）を取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            raise ValueError(f"牛が見つかりません: auto_id={cow_auto_id}")
        
        # 2. 初期状態を設定
        state = {
            'lact': cow.get('lact') or 0,
            'clvd': cow.get('clvd'),
            'rc': cow.get('rc') or self.RC_OPEN,
            'last_ai_date': None,
            'conception_date': None,
            'due_date': None,
            'dry_date': None,
            'pen': cow.get('pen'),
            'current_lact_clvd': None  # 現在の産次での分娩日
        }
        
        # 3. 全イベント履歴を取得（event_date 昇順）
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # event_date でソート（昇順：古い順）
        events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # 4. 各イベントを順に適用
        for event in events:
            self._apply_event(state, event)
        
        return state
    
    def _apply_event(self, state: Dict[str, Any], event: Dict[str, Any]):
        """
        単一イベントを状態に適用
        
        Args:
            state: 現在の状態（更新される）
            event: イベントデータ
        """
        event_number = event.get('event_number')
        event_date = event.get('event_date')
        json_data = event.get('json_data') or {}
        
        # 分娩イベント（設計書 5.3）
        if event_number == self.EVENT_CALV:
            # baseline_calvingフラグがある場合は産次を増やさない（baselineとして扱う）
            if not json_data.get('baseline_calving', False):
                state['lact'] = (state.get('lact') or 0) + 1
            state['clvd'] = event_date
            state['current_lact_clvd'] = event_date
            state['rc'] = self.RC_FRESH
            # 分娩により受胎・分娩予定日はリセット
            state['conception_date'] = None
            state['due_date'] = None
            state['last_ai_date'] = None
        
        # AI イベント（設計書 5.3）
        elif event_number == self.EVENT_AI:
            state['last_ai_date'] = event_date
            state['rc'] = self.RC_BRED
            # 受胎日・分娩予定日は一旦クリア（妊娠鑑定で確定）
            state['conception_date'] = None
            state['due_date'] = None
        
        # ET イベント（設計書 5.3）
        elif event_number == self.EVENT_ET:
            state['last_ai_date'] = event_date
            state['rc'] = self.RC_BRED
            # ETの場合は7日前が受胎日（設計書 10.3.6）
            try:
                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                conception_dt = event_dt - timedelta(days=7)
                state['conception_date'] = conception_dt.strftime('%Y-%m-%d')
                due_dt = conception_dt + timedelta(days=280)
                state['due_date'] = due_dt.strftime('%Y-%m-%d')
            except:
                pass
        
        # 妊娠鑑定プラス（PDP, PDP2, PAGP）（設計書 5.3）
        elif event_number in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
            state['rc'] = self.RC_PREGNANT
            # 受胎日を計算（設計書 10.3.6）
            if state.get('last_ai_date'):
                # 最新のAI/ET日を取得
                ai_date = state['last_ai_date']
                try:
                    ai_dt = datetime.strptime(ai_date, '%Y-%m-%d')
                    # ETの場合は7日前、AIの場合は当日
                    # ここでは簡易的にAIとして扱う（ET判定は別途必要）
                    conception_dt = ai_dt
                    state['conception_date'] = conception_dt.strftime('%Y-%m-%d')
                    due_dt = conception_dt + timedelta(days=280)
                    state['due_date'] = due_dt.strftime('%Y-%m-%d')
                except:
                    pass
        
        # 妊娠鑑定マイナス（PDN, PAGN）
        elif event_number in [self.EVENT_PDN, self.EVENT_PAGN]:
            # 妊娠していないので Open に戻す
            state['rc'] = self.RC_OPEN
            state['conception_date'] = None
            state['due_date'] = None
        
        # 流産（ABRT）
        elif event_number == self.EVENT_ABRT:
            state['rc'] = self.RC_OPEN
            state['conception_date'] = None
            state['due_date'] = None
        
        # 乾乳（DRY）（設計書 5.3）
        elif event_number == self.EVENT_DRY:
            state['rc'] = self.RC_DRY
            state['dry_date'] = event_date
        
        # 群変更（MOVE）
        elif event_number == self.EVENT_MOVE:
            if json_data.get('to_pen'):
                state['pen'] = json_data.get('to_pen')
        
        # 繁殖停止（STOPR）
        elif event_number == self.EVENT_STOPR:
            state['rc'] = self.RC_STOPPED
        
        # 売却・死亡（SOLD, DEAD）
        elif event_number in [self.EVENT_SOLD, self.EVENT_DEAD]:
            # 状態は維持（特別な処理は不要）
            pass
        
        # 乳検（601）は状態に影響しない
        # elif event_number == self.EVENT_MILK_TEST:
        #     pass
    
    def on_event_added(self, event_id: int):
        """
        イベント追加時の処理
        
        Args:
            event_id: 追加されたイベントID
        """
        # イベントを取得
        event = self._get_event_by_id(event_id)
        if not event:
            return
        
        # 対象牛の状態を再計算
        cow_auto_id = event.get('cow_auto_id')
        if cow_auto_id:
            self._recalculate_and_update_cow(cow_auto_id)
    
    def on_event_updated(self, event_id: int):
        """
        イベント更新時の処理
        
        Args:
            event_id: 更新されたイベントID
        """
        # イベントを取得
        event = self._get_event_by_id(event_id)
        if not event:
            return
        
        # 対象牛の状態を再計算
        cow_auto_id = event.get('cow_auto_id')
        if cow_auto_id:
            self._recalculate_and_update_cow(cow_auto_id)
    
    def on_event_deleted(self, event_id: int):
        """
        イベント削除時の処理
        
        Args:
            event_id: 削除されたイベントID
        """
        # 削除前のイベント情報を取得（削除済みも含む）
        event = self._get_event_by_id(event_id, include_deleted=True)
        if not event:
            return
        
        # 対象牛の状態を再計算
        cow_auto_id = event.get('cow_auto_id')
        if cow_auto_id:
            self._recalculate_and_update_cow(cow_auto_id)
    
    def _get_event_by_id(self, event_id: int, include_deleted: bool = False) -> Optional[Dict[str, Any]]:
        """
        イベントIDでイベントを取得
        
        Args:
            event_id: イベントID
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントデータ
        """
        conn = self.db.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("SELECT * FROM event WHERE id = ?", (event_id,))
        else:
            cursor.execute("SELECT * FROM event WHERE id = ? AND deleted = 0", (event_id,))
        
        row = cursor.fetchone()
        if row:
            event = dict(row)
            # json_data をパース
            if event.get('json_data'):
                try:
                    event['json_data'] = json.loads(event['json_data'])
                except:
                    event['json_data'] = {}
            return event
        
        return None
    
    def _recalculate_and_update_cow(self, cow_auto_id: int):
        """
        牛の状態を再計算して cow テーブルを更新
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 状態を再計算
        state = self.apply_events(cow_auto_id)
        
        # cow テーブルを更新
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return
        
        # 更新データを準備
        update_data = {
            'cow_id': cow.get('cow_id'),
            'jpn10': cow.get('jpn10'),
            'brd': cow.get('brd'),
            'bthd': cow.get('bthd'),
            'entr': cow.get('entr'),
            'lact': state.get('lact'),
            'clvd': state.get('clvd'),
            'rc': state.get('rc'),
            'pen': state.get('pen'),
            'frm': cow.get('frm')
        }
        
        # 更新実行
        self.db.update_cow(cow_auto_id, update_data)
    
    def get_ai_count_in_lact(self, cow_auto_id: int, lact: int) -> int:
        """
        指定産次でのAI/ET回数を取得
        
        Args:
            cow_auto_id: 牛の auto_id
            lact: 産次
        
        Returns:
            AI/ET回数
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 現在の産次での分娩日を取得
        current_lact_clvd = None
        for event in sorted(events, key=lambda e: (e.get('event_date', ''), e.get('id', 0))):
            if event.get('event_number') == self.EVENT_CALV:
                current_lact = sum(1 for e in events 
                                 if e.get('event_number') == self.EVENT_CALV 
                                 and (e.get('event_date', '') <= event.get('event_date', '') 
                                      or e.get('id', 0) <= event.get('id', 0)))
                if current_lact == lact:
                    current_lact_clvd = event.get('event_date')
                    break
        
        # 該当産次でのAI/ET回数をカウント
        count = 0
        for event in events:
            if event.get('event_number') in [self.EVENT_AI, self.EVENT_ET]:
                event_date = event.get('event_date')
                if current_lact_clvd and event_date >= current_lact_clvd:
                    count += 1
        
        return count


