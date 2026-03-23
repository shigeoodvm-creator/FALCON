"""
FALCON2 - 繁殖分析（DC305互換）
21日サイクルベースの繁殖指標計算
"""

import logging
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime, timedelta
from dataclasses import dataclass

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.formula_engine import FormulaEngine

logger = logging.getLogger(__name__)


@dataclass
class Cycle:
    """21日サイクル"""
    cycle_number: int  # サイクル番号（1から開始）
    start_date: str    # 開始日（YYYY-MM-DD）
    end_date: str      # 終了日（YYYY-MM-DD）


@dataclass
class ReproResult:
    """繁殖分析結果（1サイクル分）"""
    cycle_number: int
    start_date: str
    end_date: str
    br_el: int      # 繁殖対象個体数
    bred: int       # 授精数
    preg: int       # 妊娠数
    hdr: float      # 発情発見率（%）
    pr: float       # 妊娠率（%）


class ReproductionAnalysis:
    """繁殖分析（DC305互換）"""
    
    CYCLE_DAYS = 21  # 21日サイクル
    DEFAULT_CYCLES = 18  # デフォルトで18サイクル遡る
    EXCLUDE_RECENT_CYCLES = 2  # 直近2サイクルは妊娠未確定のため除外
    DEFAULT_VWP = 50  # デフォルトVWP（日数）
    
    def __init__(self, db: DBHandler, rule_engine: RuleEngine, 
                 formula_engine: FormulaEngine, vwp: int = DEFAULT_VWP):
        """
        初期化
        
        Args:
            db: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
            formula_engine: FormulaEngine インスタンス
            vwp: VWP（Voluntary Waiting Period、日数）
        """
        self.db = db
        self.rule_engine = rule_engine
        self.formula_engine = formula_engine
        self.vwp = vwp
        
        # イベント番号定義
        self.EVENT_AI = 200
        self.EVENT_ET = 201
        self.EVENT_CALV = 202
        self.EVENT_STOPR = 204
        self.EVENT_SOLD = 205
        self.EVENT_DEAD = 206
        self.EVENT_ABRT = 305
        self.EVENT_PDP = 303
        self.EVENT_PDP2 = 304
        self.EVENT_PAGP = 307
        
        # 繁殖コード定義（RuleEngineと同じ値を使用）
        self.RC_STOPPED = RuleEngine.RC_STOPPED  # 1
        self.RC_FRESH = RuleEngine.RC_FRESH  # 2
        self.RC_BRED = RuleEngine.RC_BRED  # 3
        self.RC_OPEN = RuleEngine.RC_OPEN  # 4
        self.RC_PREGNANT = RuleEngine.RC_PREGNANT  # 5
        self.RC_DRY = RuleEngine.RC_DRY  # 6
    
    def generate_cycles(self, period_start: Optional[str], period_end: Optional[str]) -> List[Cycle]:
        """
        21日サイクルを生成
        
        Args:
            period_start: 期間開始（YYYY-MM-DD、Noneの場合はデフォルト）
            period_end: 期間終了（YYYY-MM-DD、Noneの場合は今日）
        
        Returns:
            サイクルリスト（古い順）
        """
        # 終了日を決定
        if period_end:
            end_date = datetime.strptime(period_end, '%Y-%m-%d')
        else:
            end_date = datetime.now()
        
        cycles = []
        
        if period_start:
            # 期間指定がある場合：終了日を起算点として開始日まで21日区切りで生成
            start_date = datetime.strptime(period_start, '%Y-%m-%d')
            current_end = end_date
            cycle_number = 1
            
            while current_end >= start_date:
                cycle_start = current_end - timedelta(days=self.CYCLE_DAYS - 1)
                if cycle_start < start_date:
                    cycle_start = start_date
                
                cycles.insert(0, Cycle(
                    cycle_number=cycle_number,
                    start_date=cycle_start.strftime('%Y-%m-%d'),
                    end_date=current_end.strftime('%Y-%m-%d')
                ))
                
                current_end = cycle_start - timedelta(days=1)
                cycle_number += 1
        else:
            # 期間指定がない場合：デフォルトで18サイクル遡る
            current_end = end_date
            for cycle_num in range(self.DEFAULT_CYCLES, 0, -1):
                cycle_start = current_end - timedelta(days=self.CYCLE_DAYS - 1)
                cycles.insert(0, Cycle(
                    cycle_number=cycle_num,
                    start_date=cycle_start.strftime('%Y-%m-%d'),
                    end_date=current_end.strftime('%Y-%m-%d')
                ))
                current_end = cycle_start - timedelta(days=1)
        
        return cycles
    
    def is_breeding_eligible(self, cow_auto_id: int, cycle: Cycle, 
                            all_events: List[Dict[str, Any]]) -> bool:
        """
        繁殖対象個体かどうかを判定
        
        Args:
            cow_auto_id: 牛の auto_id
            cycle: サイクル
            all_events: 牛の全イベント履歴
        
        Returns:
            True: 繁殖対象、False: 対象外
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')
        
        # サイクル内のイベントを抽出
        cycle_events = [
            e for e in all_events
            if cycle_start <= datetime.strptime(e.get('event_date', ''), '%Y-%m-%d') <= cycle_end
        ]
        
        # サイクル開始前のイベントを抽出
        before_cycle_events = [
            e for e in all_events
            if datetime.strptime(e.get('event_date', ''), '%Y-%m-%d') < cycle_start
        ]
        
        # 1. サイクル開始前に除籍（SOLD/DEAD）されている場合は除外
        for event in before_cycle_events:
            if event.get('event_number') in [self.EVENT_SOLD, self.EVENT_DEAD]:
                return False
        
        # 2. サイクル開始前に繁殖停止が下されている場合は除外
        for event in before_cycle_events:
            if event.get('event_number') == self.EVENT_STOPR:
                return False
        
        # 3. サイクル開始前の授精で妊娠成立している場合は除外
        # （サイクル開始時点で妊娠中の場合）
        last_preg_event = None
        last_ai_before_cycle = None
        for event in sorted(before_cycle_events, key=lambda e: e.get('event_date', ''), reverse=True):
            if event.get('event_number') in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
                last_preg_event = event
                break
            if event.get('event_number') in [self.EVENT_AI, self.EVENT_ET]:
                if not last_ai_before_cycle:
                    last_ai_before_cycle = event
        
        if last_preg_event and last_ai_before_cycle:
            preg_date = datetime.strptime(last_preg_event.get('event_date', ''), '%Y-%m-%d')
            ai_date = datetime.strptime(last_ai_before_cycle.get('event_date', ''), '%Y-%m-%d')
            if preg_date > ai_date:
                # サイクル開始時点で妊娠中
                return False
        
        # 4. VWPを超えているかチェック
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return False
        
        clvd = cow.get('clvd')
        if clvd:
            try:
                clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
                days_since_calving = (cycle_start - clvd_date).days
                if days_since_calving < self.vwp:
                    # VWP内
                    return False
            except (ValueError, TypeError):
                pass
        
        # 5. サイクル開始時点で空胎（未授精、または不受胎確定）かチェック
        # サイクル開始時点より前のイベントのみで状態を計算
        state_before_cycle = self._calculate_state_at_date(cow_auto_id, cycle_start, before_cycle_events)
        rc = state_before_cycle.get('rc')
        if rc == self.RC_STOPPED:
            # サイクル開始時点で繁殖停止
            return False
        
        # 6. 流産イベントによる空胎個体は除外
        for event in cycle_events:
            if event.get('event_number') == self.EVENT_ABRT:
                return False
        
        # 7. 繁殖停止がサイクル途中で下された場合は対象
        # 8. 除籍イベントがサイクル途中で行われた場合は対象
        # （これらは既に上記のチェックで除外されていないため、対象とする）
        
        # 9. 空胎（未授精、または不受胎確定）かチェック
        # サイクル開始時点で妊娠中でないことを確認
        if rc == self.RC_PREGNANT:
            # サイクル開始時点で妊娠中
            return False
        
        return True
    
    def _calculate_state_at_date(self, cow_auto_id: int, target_date: datetime, 
                                 events_before: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        指定日時点の状態を計算（簡易版）
        
        Args:
            cow_auto_id: 牛の auto_id
            target_date: 対象日
            events_before: 対象日より前のイベントリスト
        
        Returns:
            状態辞書（rc, clvd等）
        """
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return {'rc': self.RC_OPEN}
        
        # 初期状態
        state = {
            'rc': cow.get('rc') or self.RC_OPEN,  # デフォルトはOPEN
            'clvd': cow.get('clvd'),
            'last_ai_date': None
        }
        
        # イベントを時系列で適用
        sorted_events = sorted(events_before, key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        
        for event in sorted_events:
            event_number = event.get('event_number')
            
            if event_number == self.EVENT_CALV:
                state['clvd'] = event.get('event_date')
                state['rc'] = self.RC_FRESH
            elif event_number in [self.EVENT_AI, self.EVENT_ET]:
                state['last_ai_date'] = event.get('event_date')
                if state['rc'] == self.RC_FRESH:
                    state['rc'] = self.RC_BRED
                elif state['rc'] == self.RC_OPEN:
                    state['rc'] = self.RC_BRED
            elif event_number in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
                state['rc'] = self.RC_PREGNANT
            elif event_number in [self.EVENT_PDN, self.EVENT_PAGN]:
                state['rc'] = self.RC_OPEN
            elif event_number == self.EVENT_ABRT:
                state['rc'] = self.RC_OPEN
            elif event_number == self.EVENT_STOPR:
                state['rc'] = self.RC_STOPPED
            elif event_number in [self.EVENT_SOLD, self.EVENT_DEAD]:
                # 除籍
                pass
        
        return state
    
    def count_bred(self, cow_auto_id: int, cycle: Cycle, 
                   all_events: List[Dict[str, Any]]) -> int:
        """
        授精数（Bred）を計算
        
        Args:
            cow_auto_id: 牛の auto_id
            cycle: サイクル
            all_events: 牛の全イベント履歴
        
        Returns:
            授精数（Rフラグのイベントは除外）
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')
        
        count = 0
        for event in all_events:
            event_date = event.get('event_date', '')
            if not event_date:
                continue
            
            try:
                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            
            if not (cycle_start <= event_dt <= cycle_end):
                continue
            
            # AI/ETイベントをカウント
            if event.get('event_number') in [self.EVENT_AI, self.EVENT_ET]:
                # Rフラグ（outcome='R'）のイベントは除外
                json_data = event.get('json_data') or {}
                if isinstance(json_data, str):
                    try:
                        json_data = json.loads(json_data)
                    except:
                        json_data = {}
                
                outcome = json_data.get('outcome')
                if outcome != 'R':
                    count += 1
        
        return count
    
    def count_preg(self, cow_auto_id: int, cycle: Cycle, 
                   all_events: List[Dict[str, Any]]) -> int:
        """
        妊娠数（Preg）を計算
        
        Args:
            cow_auto_id: 牛の auto_id
            cycle: サイクル
            all_events: 牛の全イベント履歴
        
        Returns:
            妊娠数（サイクル内で受胎が確定した授精数）
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')
        
        # サイクル内のAI/ETイベントを取得
        ai_et_events = []
        for event in all_events:
            event_date = event.get('event_date', '')
            if not event_date:
                continue
            
            try:
                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            
            if not (cycle_start <= event_dt <= cycle_end):
                continue
            
            if event.get('event_number') in [self.EVENT_AI, self.EVENT_ET]:
                ai_et_events.append(event)
        
        # サイクル内の妊娠確定イベントを取得
        preg_events = []
        for event in all_events:
            event_date = event.get('event_date', '')
            if not event_date:
                continue
            
            try:
                event_dt = datetime.strptime(event_date, '%Y-%m-%d')
            except (ValueError, TypeError):
                continue
            
            if event.get('event_number') in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
                if cycle_start <= event_dt <= cycle_end:
                    preg_events.append(event)
        
        # サイクル内のAI/ETイベントと妊娠確定イベントをマッチング
        preg_count = 0
        for ai_et_event in ai_et_events:
            ai_et_date = datetime.strptime(ai_et_event.get('event_date', ''), '%Y-%m-%d')
            # このAI/ETイベントに対応する妊娠確定イベントを探す
            for preg_event in preg_events:
                preg_date = datetime.strptime(preg_event.get('event_date', ''), '%Y-%m-%d')
                # 妊娠確定がAI/ETの後で、かつ合理的な期間内（例：300日以内）にある場合
                if preg_date > ai_et_date and (preg_date - ai_et_date).days <= 300:
                    preg_count += 1
                    break
        
        return preg_count
    
    def analyze_cycle(self, cycle: Cycle, cows: List[Dict[str, Any]], 
                     exclude_recent: bool = False) -> ReproResult:
        """
        1サイクルの繁殖分析を実行
        
        Args:
            cycle: サイクル
            cows: 全牛リスト
            exclude_recent: 直近サイクルとして除外するか（妊娠数=0とする）
        
        Returns:
            繁殖分析結果
        """
        br_el = 0  # 繁殖対象個体数
        bred = 0   # 授精数
        preg = 0    # 妊娠数
        
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # 全イベント履歴を取得
            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            
            # 繁殖対象個体かどうかを判定
            if self.is_breeding_eligible(cow_auto_id, cycle, all_events):
                br_el += 1
                
                # 授精数をカウント
                bred += self.count_bred(cow_auto_id, cycle, all_events)
                
                # 妊娠数をカウント（直近サイクルの場合は0）
                if not exclude_recent:
                    preg += self.count_preg(cow_auto_id, cycle, all_events)
        
        # 指標計算
        hdr = (bred / br_el * 100) if br_el > 0 else 0.0
        pr = (preg / br_el * 100) if br_el > 0 else 0.0
        
        return ReproResult(
            cycle_number=cycle.cycle_number,
            start_date=cycle.start_date,
            end_date=cycle.end_date,
            br_el=br_el,
            bred=bred,
            preg=preg,
            hdr=round(hdr, 2),
            pr=round(pr, 2)
        )
    
    def analyze(self, period_start: Optional[str], period_end: Optional[str]) -> List[ReproResult]:
        """
        繁殖分析を実行
        
        Args:
            period_start: 期間開始（YYYY-MM-DD）
            period_end: 期間終了（YYYY-MM-DD）
        
        Returns:
            繁殖分析結果リスト（各サイクル分）
        """
        # サイクルを生成
        cycles = self.generate_cycles(period_start, period_end)
        
        # 全牛を取得
        cows = self.db.get_all_cows()
        
        # 各サイクルを分析
        results = []
        for i, cycle in enumerate(cycles):
            # 直近2サイクルは妊娠未確定のため、妊娠数=0とする
            exclude_recent = i >= len(cycles) - self.EXCLUDE_RECENT_CYCLES
            result = self.analyze_cycle(cycle, cows, exclude_recent=exclude_recent)
            results.append(result)
        
        return results

