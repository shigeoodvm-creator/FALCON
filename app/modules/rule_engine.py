"""
FALCON2 - RuleEngine
イベント履歴から牛の状態を完全再計算する
設計書 第5章参照
"""

from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta, date
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler

logger = logging.getLogger(__name__)


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
    EVENT_FCHK = 300        # フレッシュチェック
    EVENT_REPRO = 301       # 繁殖検査
    EVENT_ABRT = 305        # 流産
    EVENT_PAGN = 306        # PAGマイナス
    EVENT_PAGP = 307        # PAGプラス
    
    EVENT_MILK_TEST = 601   # 乳検
    EVENT_GENOMIC = 602     # ゲノム（未実装）
    EVENT_MOVE = 611        # 群変更
    EVENT_IN = 600          # 導入
    
    # 繁殖コード（RC）定義
    RC_STOPPED = 1          # Stopped（繁殖停止）
    RC_FRESH = 2            # Fresh（分娩後）
    RC_BRED = 3             # Bred（授精後）
    RC_OPEN = 4             # Open（空胎）
    RC_PREGNANT = 5         # Pregnant（妊娠中）
    RC_DRY = 6              # Dry（乾乳中）
    
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
            'current_lact_clvd': None,  # 現在の産次での分娩日
            'cow_auto_id': cow_auto_id  # イベント判定用に保持
        }
        
        # 3. 全イベント履歴を取得（event_date 昇順）
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # event_date でソート（昇順：古い順）
        events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # 4. 各イベントを順に適用
        for event in events:
            self._apply_event(state, event)
        
        return state
    
    def apply_events_until_date(self, cow_auto_id: int, target_date: str) -> Dict[str, Any]:
        """
        イベント日付時点での状態を計算
        
        Args:
            cow_auto_id: 牛の auto_id
            target_date: 対象日付（YYYY-MM-DD形式、この日付以前のイベントを適用）
        
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
            'current_lact_clvd': None,  # 現在の産次での分娩日
            'cow_auto_id': cow_auto_id  # イベント判定用に保持
        }
        
        # 3. 全イベント履歴を取得
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 4. 対象日付以前（同じ日を含む）のイベントのみをフィルタリング
        target_dt = datetime.strptime(target_date, '%Y-%m-%d')
        events_before = []
        for ev in events:
            ev_date_str = ev.get('event_date', '')
            if ev_date_str:
                try:
                    ev_dt = datetime.strptime(ev_date_str, '%Y-%m-%d')
                    if ev_dt <= target_dt:
                        events_before.append(ev)
                except ValueError:
                    continue
        
        # 5. event_date でソート（昇順：古い順）
        events_before.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # 6. 各イベントを順に適用
        for event in events_before:
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
                # 最新のAI/ETイベントを取得してイベントタイプを判定
                cow_auto_id = state.get('cow_auto_id')
                if cow_auto_id:
                    events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                    latest_ai_et_event = None
                    for e in sorted(events, key=lambda x: (x.get('event_date', ''), x.get('id', 0)), reverse=True):
                        if e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]:
                            if e.get('event_date') == ai_date:
                                latest_ai_et_event = e
                                break
                    
                    try:
                        ai_dt = datetime.strptime(ai_date, '%Y-%m-%d')
                        # ETの場合は7日前、AIの場合は当日が受胎日（設計書 10.3.6）
                        if latest_ai_et_event and latest_ai_et_event.get('event_number') == self.EVENT_ET:
                            conception_dt = ai_dt - timedelta(days=7)
                        else:  # AI
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
        イベント追加時の処理（Phase A: event_lact/event_dim も再計算）
        
        Args:
            event_id: 追加されたイベントID
        """
        # イベントを取得
        event = self._get_event_by_id(event_id)
        if not event:
            return
        
        cow_auto_id = event.get('cow_auto_id')
        if not cow_auto_id:
            return
        
        event_number = event.get('event_number')
        # 非baselineの分娩追加時：削除時に産次を戻すため、追加前の産次を baseline_lact に保存
        if event_number == self.EVENT_CALV:
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            if not json_data.get('baseline_calving', False):
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                if cow is not None:
                    current_lact = cow.get('lact')
                    if current_lact is not None:
                        self.db.update_cow(cow_auto_id, {'baseline_lact': current_lact})
        
        # 対象牛の全イベントを再計算（event_lact/event_dim を含む）
        self.recalculate_events_for_cow(cow_auto_id)
        # AI/ETイベントのoutcomeを更新
        self.update_insemination_outcomes(cow_auto_id)
        # 授精カウントを更新
        if event_number == self.EVENT_CALV:
            self.update_insemination_counts_for_cow(cow_auto_id, preserve_existing=True)
        else:
            self.update_insemination_counts_for_cow(cow_auto_id, preserve_existing=False)
        # 牛の状態（lact, clvd, rc, pen）を再計算してcowテーブルを更新
        self._recalculate_and_update_cow(cow_auto_id)
    
    def on_event_updated(self, event_id: int):
        """
        イベント更新時の処理（Phase A: event_lact/event_dim も再計算）
        授精カウントも再計算
        
        Args:
            event_id: 更新されたイベントID
        """
        # イベントを取得
        event = self._get_event_by_id(event_id)
        if not event:
            return
        
        # 対象牛の全イベントを再計算（event_lact/event_dim を含む）
        cow_auto_id = event.get('cow_auto_id')
        event_number = event.get('event_number')
        if cow_auto_id:
            self.recalculate_events_for_cow(cow_auto_id)
            # AI/ETイベントのoutcomeを更新
            self.update_insemination_outcomes(cow_auto_id)
            # 授精カウントを更新
            # 分娩イベントが更新された場合は、既存のAI/ETイベントのカウントは保持（更新しない）
            # 新規のAI/ETイベントのみカウントを設定
            if event_number == self.EVENT_CALV:
                # 分娩イベント更新時：既存のカウントは保持、新規イベントのみ更新
                self.update_insemination_counts_for_cow(cow_auto_id, preserve_existing=True)
            else:
                # AI/ETイベント更新時：全イベントを再計算（日付変更などに対応）
                self.update_insemination_counts_for_cow(cow_auto_id, preserve_existing=False)
            # 牛の状態（lact, clvd, rc, pen）を再計算してcowテーブルを更新
            self._recalculate_and_update_cow(cow_auto_id)
    
    def on_event_deleted(self, event_id: int, cow_auto_id: Optional[int] = None):
        """
        イベント削除時の処理（Phase A: event_lact/event_dim も再計算）
        物理削除の場合は削除後にイベントを取得できないため、呼び出し元で cow_auto_id を渡すこと。
        
        Args:
            event_id: 削除されたイベントID
            cow_auto_id: 対象牛の auto_id（物理削除時は必須。渡されない場合は削除前イベントから取得）
        """
        # 削除前のイベント情報を取得（削除済みも含む）。物理削除の場合は取得できない。
        event = self._get_event_by_id(event_id, include_deleted=True)
        if event:
            target_cow_auto_id = event.get('cow_auto_id')
        else:
            target_cow_auto_id = cow_auto_id
        
        if not target_cow_auto_id:
            return
        
        # 対象牛の全イベントを再計算（event_lact/event_dim を含む）
        self.recalculate_events_for_cow(target_cow_auto_id)
        # AI/ETイベントのoutcomeを更新
        self.update_insemination_outcomes(target_cow_auto_id)
        # 授精カウントを更新
        event_number = event.get('event_number') if event else None
        if event_number == self.EVENT_CALV:
            self.update_insemination_counts_for_cow(target_cow_auto_id, preserve_existing=True)
        else:
            self.update_insemination_counts_for_cow(target_cow_auto_id, preserve_existing=False)
        # 牛の状態（lact, clvd, rc, pen）を再計算してcowテーブルを更新
        self._recalculate_and_update_cow(target_cow_auto_id)
    
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
        
        clvd, rc, pen は apply_events の結果で更新する。
        産次（lact）は recalculate_events_for_cow が既に正しく計算・更新しているため、
        ここでは cow テーブルの現在値（再取得）を使い、apply_events の state.lact で
        上書きしない。さもないと apply_events が「cow.lact + CALVごと+1」で二重カウントし、
        分娩追加で+2・削除後も戻らない不具合になる。
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 状態を再計算（clvd, rc, pen 取得用）
        state = self.apply_events(cow_auto_id)
        
        # cow テーブルを再取得（recalculate_events_for_cow が lact を既に更新済み）
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return
        
        # 更新データを準備
        # lact は recalculate_events_for_cow の結果を保持（state は二重カウントのため使わない）
        update_data = {
            'cow_id': cow.get('cow_id'),
            'jpn10': cow.get('jpn10'),
            'brd': cow.get('brd'),
            'bthd': cow.get('bthd'),
            'entr': cow.get('entr'),
            'lact': cow.get('lact'),
            'clvd': state.get('clvd'),
            'rc': state.get('rc'),
            'pen': state.get('pen'),
            'frm': cow.get('frm')
        }
        
        # 更新実行
        self.db.update_cow(cow_auto_id, update_data)
    
    def calculate_insemination_counts(self, cow_auto_id: int) -> Dict[int, int]:
        """
        牛のAI/ETイベントの授精カウントを計算
        
        ルール:
        - 分娩イベントから順に1, 2, 3とカウント
        - AI/ETイベント間で間隔が8日以内の場合は同じカウント
        - 後からイベントを追加した場合も再計算される
        
        Args:
            cow_auto_id: 牛の auto_id
        
        Returns:
            {event_id: insemination_count, ...} の辞書
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 分娩イベントを取得（日付順、ID順）
        calving_events = [
            e for e in events 
            if e.get('event_number') == self.EVENT_CALV
        ]
        calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # AI/ETイベントを取得（日付順、ID順）
        ai_et_events = [
            e for e in events 
            if e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]
        ]
        ai_et_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        if not ai_et_events:
            return {}
        
        # 分娩イベントごとに授精カウントを計算
        result = {}
        
        # 各分娩イベント以降のAI/ETイベントを処理
        for calving_idx, calving_event in enumerate(calving_events):
            calving_date = calving_event.get('event_date')
            if not calving_date:
                continue
            
            try:
                calving_dt = datetime.strptime(calving_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            
            # この分娩イベント以降のAI/ETイベントを取得
            # 次の分娩イベントより前のもの
            next_calving_date = None
            if calving_idx + 1 < len(calving_events):
                next_calving_date = calving_events[calving_idx + 1].get('event_date')
            
            current_count = 0
            last_insemination_date = None
            
            for ai_et_event in ai_et_events:
                event_id = ai_et_event.get('id')
                event_date = ai_et_event.get('event_date')
                
                if not event_date:
                    continue
                
                try:
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                except (ValueError, TypeError):
                    continue
                
                # この分娩イベント以降、次の分娩イベントより前かどうか
                if event_dt < calving_dt:
                    continue
                
                if next_calving_date:
                    try:
                        next_calving_dt = datetime.strptime(next_calving_date, '%Y-%m-%d')
                        if event_dt >= next_calving_dt:
                            continue
                    except (ValueError, TypeError):
                        pass
                
                # 8日以内の間隔の場合は同じカウント
                if last_insemination_date:
                    try:
                        last_dt = datetime.strptime(last_insemination_date, '%Y-%m-%d')
                        days_diff = (event_dt - last_dt).days
                        if days_diff <= 8:
                            # 同じカウント
                            result[event_id] = current_count
                            continue
                    except (ValueError, TypeError):
                        pass
                
                # カウントを増やす
                current_count += 1
                result[event_id] = current_count
                last_insemination_date = event_date
        
        return result
    
    def update_insemination_counts_for_cow(self, cow_auto_id: int, preserve_existing: bool = False) -> None:
        """
        牛のAI/ETイベントの授精カウントを計算してDBに保存
        
        Args:
            cow_auto_id: 牛の auto_id
            preserve_existing: Trueの場合は既存のカウントを保持（更新しない）、Falseの場合は全イベントを更新
        """
        counts = self.calculate_insemination_counts(cow_auto_id)
        
        for event_id, count in counts.items():
            event = self.db.get_event_by_id(event_id)
            if not event:
                continue
            
            # json_dataを取得・更新
            json_data = event.get('json_data')
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except (json.JSONDecodeError, TypeError):
                    json_data = {}
            elif json_data is None:
                json_data = {}
            
            # 既存のカウントがある場合、preserve_existingがTrueの場合はスキップ（保持）
            if preserve_existing and 'insemination_count' in json_data:
                continue
            
            # 授精カウントを設定
            json_data['insemination_count'] = count
            
            # DBを更新
            self.db.update_event(event_id, {'json_data': json_data})
    
    def update_all_insemination_counts(self) -> None:
        """
        全牛のAI/ETイベントの授精カウントを計算して更新
        """
        # 全牛を取得
        conn = self.db.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT auto_id FROM cow")
        cows = cursor.fetchall()
        
        total_cows = len(cows)
        logger.info(f"[授精カウント] 全{total_cows}頭の授精カウントを更新開始")
        
        for idx, cow_row in enumerate(cows, 1):
            cow_auto_id = cow_row['auto_id']
            try:
                self.update_insemination_counts_for_cow(cow_auto_id)
                if idx % 100 == 0:
                    logger.info(f"[授精カウント] 進捗: {idx}/{total_cows}頭完了")
            except Exception as e:
                logger.error(f"[授精カウント] 牛auto_id={cow_auto_id}の更新エラー: {e}")
        
        logger.info(f"[授精カウント] 全{total_cows}頭の授精カウント更新完了")
    
    def apply_insemination_count_from_item(self, cow_auto_id: int, total_count: int) -> None:
        """
        項目「授精回数」の値を直近産次のAI/ETイベントに反映する。
        ユーザーが個体カードで授精回数を変更した場合、直近のAIをtotal_count、
        その前をtotal_count-1、… と割り当ててDBを更新する。
        
        Args:
            cow_auto_id: 牛の auto_id
            total_count: 授精回数（直近AIの回数。1以上）
        """
        if total_count is None or total_count < 1:
            return
        try:
            total_count = int(total_count)
        except (TypeError, ValueError):
            return
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        calving_events = [
            e for e in events
            if e.get('event_number') == self.EVENT_CALV
        ]
        calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]
        ]
        ai_et_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        if not calving_events or not ai_et_events:
            return
        # 直近産次 = 最後の分娩以降のAI/ET
        latest_calving_date = calving_events[-1].get('event_date')
        if not latest_calving_date:
            return
        try:
            calving_dt = datetime.strptime(latest_calving_date, '%Y-%m-%d')
        except (ValueError, TypeError):
            return
        current_lact_ai_et = []
        for e in ai_et_events:
            ed = e.get('event_date')
            if not ed:
                continue
            try:
                edt = datetime.strptime(ed, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            if edt >= calving_dt:
                current_lact_ai_et.append(e)
        if not current_lact_ai_et:
            return
        # 新しい順（直近が先頭）
        current_lact_ai_et.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)), reverse=True)
        for i, ev in enumerate(current_lact_ai_et):
            count = total_count - i
            if count < 1:
                break
            event_id = ev.get('id')
            if not event_id:
                continue
            event = self.db.get_event_by_id(event_id)
            if not event:
                continue
            json_data = event.get('json_data')
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except (json.JSONDecodeError, TypeError):
                    json_data = {}
            elif json_data is None:
                json_data = {}
            json_data['insemination_count'] = count
            self.db.update_event(event_id, {'json_data': json_data})
        logger.debug(f"[授精カウント] apply_insemination_count_from_item: cow_auto_id={cow_auto_id}, total_count={total_count}, updated {min(len(current_lact_ai_et), total_count)} events")
    
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
    
    def update_insemination_outcomes(self, cow_auto_id: int) -> None:
        """
        AI/ETイベントのoutcome（P/O/R/N）を確定してDBに保存
        
        ルール:
        - P: 妊娠プラス系イベント（PDP/PDP2/PAGP）で直近の未確定AI/ETを確定
        - O: 妊娠マイナス系イベント（PDN/PAGN）で直近の未確定AI/ETを確定
        - R: 8日以内に次のAI/ETが入力された場合、前のAI/ETをRに確定
        - N: P確定後のAI/ETはNに確定
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 1. 牛のbaseline値を取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            logger.warning(f"牛が見つかりません: auto_id={cow_auto_id}")
            return
        
        # 2. 今産次の分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            # 分娩日がない場合は処理しない
            return
        
        # 3. 全イベントを取得（event_date, event_id 昇順）
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # 4. 今産次のAI/ETイベントを抽出（分娩日より後）
        ai_et_events = []
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if event_number in [self.EVENT_AI, self.EVENT_ET] and event_date and event_date > clvd:
                ai_et_events.append(event)
        
        if not ai_et_events:
            return
        
        # 5. 各AI/ETイベントのoutcomeを一旦Noneに初期化（既存のoutcomeをクリア）
        conn = self.db.connect()
        cursor = conn.cursor()
        
        try:
            # 各AI/ETイベントのjson_dataを取得・更新
            outcomes = {}  # {event_id: outcome}
            
            # まず、すべてのAI/ETイベントのoutcomeを一旦Noneに
            for event in ai_et_events:
                event_id = event.get('id')
                json_data = event.get('json_data') or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except:
                        json_data = {}
                
                # outcomeを一旦削除（後で再計算）
                if 'outcome' in json_data:
                    del json_data['outcome']
                    outcomes[event_id] = None
            
            # 6. イベントを時系列で走査してoutcomeを判定
            # 全イベントを時系列順にソート（AI/ET + 妊娠関連イベント）
            all_relevant_events = []
            for event in events:
                event_number = event.get('event_number')
                event_date = event.get('event_date')
                if event_date and event_date > clvd:
                    if event_number in [self.EVENT_AI, self.EVENT_ET]:
                        all_relevant_events.append(('AI_ET', event))
                    elif event_number in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
                        all_relevant_events.append(('PREG_PLUS', event))
                    elif event_number in [self.EVENT_PDN, self.EVENT_PAGN]:
                        all_relevant_events.append(('PREG_MINUS', event))
            
            all_relevant_events.sort(key=lambda x: (x[1].get('event_date', ''), x[1].get('id', 0)))
            
            # 時系列で走査してoutcomeを確定
            for event_type, event in all_relevant_events:
                event_date = event.get('event_date', '')
                
                if event_type == 'PREG_PLUS':
                    # 妊娠プラスイベント: 直近の未確定AI/ETをPに確定
                    # この妊娠プラスイベントより前の、未確定のAI/ETイベントを探す
                    for ai_event in reversed(ai_et_events):
                        ai_event_id = ai_event.get('id')
                        ai_event_date = ai_event.get('event_date', '')
                        
                        if ai_event_date > event_date:
                            continue
                        
                        # 既にoutcomeが確定している場合はスキップ
                        if outcomes.get(ai_event_id):
                            continue
                        
                        # PDP2の場合はai_event_idが指定されている場合がある
                        if event.get('event_number') == self.EVENT_PDP2:
                            preg_json = event.get('json_data') or {}
                            if isinstance(preg_json, str):
                                try:
                                    preg_json = json.loads(preg_json)
                                except:
                                    preg_json = {}
                            specified_ai_id = preg_json.get('ai_event_id')
                            if specified_ai_id and specified_ai_id != ai_event_id:
                                continue
                        
                        # このAI/ETイベントの後に別のAI/ETイベントが存在する場合は、
                        # 次のAI/ETイベントの方が妊娠プラスイベントに対応する可能性があるためスキップ
                        # （後でOに設定される）
                        ai_index = next((i for i, e in enumerate(ai_et_events) if e.get('id') == ai_event_id), -1)
                        if ai_index >= 0 and ai_index < len(ai_et_events) - 1:
                            next_ai = ai_et_events[ai_index + 1]
                            next_date = next_ai.get('event_date', '')
                            # 次のAI/ETが妊娠プラスイベントより前の場合は、次のAI/ETを優先
                            if next_date <= event_date:
                                continue
                        
                        outcomes[ai_event_id] = 'P'
                        break
                
                elif event_type == 'PREG_MINUS':
                    # 妊娠マイナスイベント: 直近の未確定AI/ETをOに確定
                    # この妊娠マイナスイベントより前の、未確定のAI/ETイベントを探す
                    for ai_event in reversed(ai_et_events):
                        ai_event_id = ai_event.get('id')
                        ai_event_date = ai_event.get('event_date', '')
                        
                        if ai_event_date > event_date:
                            continue
                        
                        # 既にoutcomeが確定している場合はスキップ
                        if outcomes.get(ai_event_id):
                            continue
                        
                        # 次のAI/ETイベントより前か確認
                        ai_index = next((i for i, e in enumerate(ai_et_events) if e.get('id') == ai_event_id), -1)
                        if ai_index >= 0 and ai_index < len(ai_et_events) - 1:
                            next_ai = ai_et_events[ai_index + 1]
                            next_date = next_ai.get('event_date', '')
                            # 次のAI/ETが妊娠マイナスイベントより前の場合はスキップ
                            if next_date < event_date:
                                continue
                        
                        # DC305取込で結果未確定（-や空欄）の場合はOを付けない
                        ai_json = ai_event.get('json_data') or {}
                        if isinstance(ai_json, str):
                            try:
                                ai_json = json.loads(ai_json)
                            except Exception:
                                ai_json = {}
                        if ai_json.get('_dc305_no_result'):
                            break
                        
                        outcomes[ai_event_id] = 'O'
                        break
            
            # 7. RとNを判定（妊娠関連イベントの処理後）
            for i, ai_event in enumerate(ai_et_events):
                event_id = ai_event.get('id')
                event_date = ai_event.get('event_date', '')
                
                if not event_id or not event_date:
                    continue
                
                # 既にoutcomeが確定している場合はスキップ
                if outcomes.get(event_id):
                    continue
                
                # N: P確定後のAI/ET
                # このAI/ETより前のAI/ETでPが確定しているか確認
                for prev_ai in ai_et_events[:i]:
                    if outcomes.get(prev_ai.get('id')) == 'P':
                        outcomes[event_id] = 'N'
                        break
                
                # 既にNが確定している場合はスキップ
                if outcomes.get(event_id) == 'N':
                    continue
                
                # R: 8日以内に次のAI/ETが入力された場合
                if i < len(ai_et_events) - 1:
                    next_ai = ai_et_events[i + 1]
                    next_date = next_ai.get('event_date')
                    try:
                        current_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        next_dt = datetime.strptime(next_date, '%Y-%m-%d')
                        days_diff = (next_dt - current_dt).days
                        if days_diff <= 8:
                            outcomes[event_id] = 'R'
                            continue
                    except (ValueError, TypeError):
                        pass
            
            # 8. Oを判定（次のAI/ETイベントが存在する場合）
            # RとNの判定後、次のAI/ETが存在する場合、その前のAI/ETをOに設定
            # （ただし、P、N、Rが既に確定している場合は除く）
            for i, ai_event in enumerate(ai_et_events):
                event_id = ai_event.get('id')
                
                if not event_id:
                    continue
                
                # 既にP、N、Rが確定している場合はスキップ
                current_outcome = outcomes.get(event_id)
                if current_outcome in ['P', 'N', 'R']:
                    continue
                
                # DC305取込で結果未確定（-や空欄）の場合はOを付けない
                ai_json = ai_event.get('json_data') or {}
                if isinstance(ai_json, str):
                    try:
                        ai_json = json.loads(ai_json)
                    except Exception:
                        ai_json = {}
                if ai_json.get('_dc305_no_result'):
                    continue
                
                # 次のAI/ETイベントが存在する場合、このAI/ETをOに設定
                if i < len(ai_et_events) - 1:
                    outcomes[event_id] = 'O'
            
            # 9. 判定結果をevent.json_dataに書き戻す
            for event in ai_et_events:
                event_id = event.get('id')
                outcome = outcomes.get(event_id)
                
                # json_dataを取得
                json_data = event.get('json_data') or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except:
                        json_data = {}
                
                # outcomeを更新
                if outcome:
                    json_data['outcome'] = outcome
                elif 'outcome' in json_data:
                    # outcomeがNoneの場合は削除
                    del json_data['outcome']
                
                # DBに保存
                json_str = json.dumps(json_data, ensure_ascii=False) if json_data else None
                cursor.execute("""
                    UPDATE event 
                    SET json_data = ?
                    WHERE id = ?
                """, (json_str, event_id))
            
            conn.commit()
            logger.info(f"Updated insemination outcomes for cow auto_id={cow_auto_id}: {len([o for o in outcomes.values() if o])} outcomes set")
            
        except Exception as e:
            conn.rollback()
            logger.error(f"Error updating insemination outcomes for cow auto_id={cow_auto_id}: {e}", exc_info=True)
            raise
    
    def recalculate_events_for_cow(self, cow_auto_id: int) -> None:
        """
        Phase A: 対象牛の全イベントについて event_lact と event_dim を再計算してDBに書き込む。
        産次（lact）は baseline_lact + 非baseline分娩イベント数 で算出する。
        分娩削除時は baseline_lact のみになるため元の産次に戻る。
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        # 1. 牛のbaseline値を取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            raise ValueError(f"牛が見つかりません: auto_id={cow_auto_id}")
        
        # 2. 産次は baseline_lact を起点にし、非baseline分娩の数だけ加算（削除時に元に戻るため）
        baseline_lact = cow.get('baseline_lact')
        if baseline_lact is None:
            baseline_lact = cow.get('lact') or 0  # 既存DBで baseline_lact 未設定時
        lact = baseline_lact
        last_calv_date = None  # baselineでない分娩日の最新値
        
        # 3. 全イベントを取得（event_date, event_id 昇順）
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        # 4. 各イベントの event_lact/event_dim を計算してUPDATE
        conn = self.db.connect()
        cursor = conn.cursor()
        
        updated_count = 0
        skipped_count = 0
        
        try:
            for event in events:
                event_id = event.get('id')
                event_date = event.get('event_date')
                event_number = event.get('event_number')
                
                # event_dateがNULLの場合はスキップ
                if not event_date:
                    logger.warning(f"Event id={event_id} has NULL event_date, skipping")
                    skipped_count += 1
                    continue
                
                # json_dataを安全に取得
                json_data = event.get('json_data') or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except:
                        json_data = {}
                
                baseline_calving = json_data.get('baseline_calving', False)
                
                # event_lact を決定
                event_lact = lact  # デフォルトは現在のlact
                
                # event_dim を計算
                event_dim = None
                if last_calv_date:
                    try:
                        calv_dt = datetime.strptime(last_calv_date, '%Y-%m-%d')
                        event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        event_dim = (event_dt - calv_dt).days
                    except ValueError as e:
                        logger.warning(f"Date parse error for event id={event_id}: {e}")
                        event_dim = None
                
                # 分娩イベントの場合の特別処理
                if event_number == self.EVENT_CALV:
                    if baseline_calving:
                        # baseline分娩：lactは増やさないが、last_calv_dateは更新してDIM計算を可能にする
                        event_lact = lact
                        last_calv_date = event_date
                        event_dim = 0  # 分娩当日はDIM=0
                    else:
                        # 通常分娩：lactを+1してからevent_lactに設定
                        lact = lact + 1
                        event_lact = lact
                        last_calv_date = event_date
                        event_dim = 0  # 分娩当日はDIM=0
                
                # eventテーブルをUPDATE
                cursor.execute("""
                    UPDATE event 
                    SET event_lact = ?, event_dim = ?
                    WHERE id = ?
                """, (event_lact, event_dim, event_id))
                
                updated_count += 1
            
            # 5. cowテーブルのlactを更新
            cursor.execute("""
                UPDATE cow 
                SET lact = ?
                WHERE auto_id = ?
            """, (lact, cow_auto_id))
            
            conn.commit()
            
            # 6. ログ出力
            logger.info(
                f"Recalculated events for cow auto_id={cow_auto_id}: "
                f"updated={updated_count}, skipped={skipped_count}, "
                f"final_lact={lact}, last_calv_date={last_calv_date or 'None'}"
            )
            
        except Exception as e:
            conn.rollback()
            logger.error(f"Error recalculating events for cow auto_id={cow_auto_id}: {e}", exc_info=True)
            raise

    def recalculate_events_for_cow_with_lact_override(self, cow_auto_id: int, lact_override: int) -> None:
        """
        個体カードで産次を手動更新した際に、全イベントの event_lact / event_dim を
        その産次に合わせて再計算する。
        - 今産次（最後の分娩以降）: event_lact = lact_override
        - 前産次: event_lact = lact_override - 1
        - 前々産次以降も同様に lact_override - 2, ...
        cow テーブルの lact は呼び出し元で既に更新済みであること。

        Args:
            cow_auto_id: 牛の auto_id
            lact_override: 個体カードで設定した産次（正の整数）
        """
        if lact_override is None or lact_override < 0:
            return
        lact_override = int(lact_override)

        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        events.sort(key=lambda e: (e.get('event_date', '') or '', e.get('id', 0)))

        # 通常分娩（baseline でない）の日付リスト（昇順）
        calving_dates: List[str] = []
        for e in events:
            if e.get('event_number') != self.EVENT_CALV:
                continue
            ed = e.get('event_date')
            if not ed:
                continue
            j = e.get('json_data') or {}
            if isinstance(j, str):
                try:
                    j = json.loads(j)
                except Exception:
                    j = {}
            if j.get('baseline_calving', False):
                continue
            calving_dates.append(ed)
        calving_dates.sort()

        conn = self.db.connect()
        cursor = conn.cursor()
        try:
            for event in events:
                event_id = event.get('id')
                event_date = event.get('event_date')
                if not event_date:
                    continue

                # このイベントより後の分娩の数 → 今産次から何産前か
                n_after = sum(1 for cd in calving_dates if cd > event_date)
                event_lact = max(0, lact_override - n_after)

                # event_dim: 直近の分娩日以降なら DIM
                event_dim = None
                last_calv_on_or_before = None
                for cd in calving_dates:
                    if cd <= event_date:
                        last_calv_on_or_before = cd
                if last_calv_on_or_before:
                    try:
                        calv_dt = datetime.strptime(last_calv_on_or_before, '%Y-%m-%d')
                        event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        event_dim = (event_dt - calv_dt).days
                    except ValueError:
                        event_dim = None

                cursor.execute("""
                    UPDATE event
                    SET event_lact = ?, event_dim = ?
                    WHERE id = ?
                """, (event_lact, event_dim, event_id))

            conn.commit()
            logger.info(
                f"Recalculated event_lact/event_dim for cow auto_id={cow_auto_id} with lact_override={lact_override}"
            )
        except Exception as e:
            conn.rollback()
            logger.error(
                f"Error recalculating events with lact override for cow auto_id={cow_auto_id}: {e}",
                exc_info=True,
            )
            raise

    def fix_null_lact_cows(self) -> int:
        """
        産次（lact）が NULL だが乳検・分娩などのイベントがある個体について、
        イベントから産次を算出し cow.lact および全イベントの event_lact/event_dim を記録し直す。
        recalculate_events_for_cow のロジック（分娩ごとに産次をカウント）で統一する。

        Returns:
            修正した個体数
        """
        cows = self.db.get_all_cows()
        fixed_count = 0
        for cow in cows:
            auto_id = cow.get('auto_id')
            if auto_id is None:
                continue
            lact = cow.get('lact')
            if lact is not None:
                continue
            events = self.db.get_events_by_cow(auto_id, include_deleted=False)
            events_with_date = [e for e in events if e.get('event_date')]
            if not events_with_date:
                continue
            try:
                self.recalculate_events_for_cow(auto_id)
                fixed_count += 1
                logger.info(f"Fixed null lact for cow auto_id={auto_id} (cow_id={cow.get('cow_id')})")
            except Exception as e:
                logger.error(f"Failed to fix null lact for cow auto_id={auto_id}: {e}", exc_info=True)
        return fixed_count

    def sync_event_lact_for_all_cows(self) -> int:
        """
        産次（lact）が設定されている全頭について、全イベントの event_lact/event_dim を
        cow.lact に合わせて記録し直す。分娩後の乳検などで event_lact が取り残されている場合の解消用。

        Returns:
            同期した個体数
        """
        cows = self.db.get_all_cows()
        synced_count = 0
        for cow in cows:
            auto_id = cow.get('auto_id')
            if auto_id is None:
                continue
            lact = cow.get('lact')
            if lact is None or lact < 0:
                continue
            events = self.db.get_events_by_cow(auto_id, include_deleted=False)
            if not events:
                continue
            try:
                self.recalculate_events_for_cow_with_lact_override(auto_id, int(lact))
                synced_count += 1
            except Exception as e:
                logger.warning(f"sync_event_lact for cow auto_id={auto_id}: {e}")
        return synced_count

