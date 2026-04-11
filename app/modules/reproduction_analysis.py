"""
FALCON2 - 繁殖分析（DC305互換）
21日サイクルベースの繁殖指標計算
"""

import json
import logging
from typing import Any, Callable, Dict, List, Optional, Tuple
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
    br_el: int          # 繁殖対象個体数
    bred: int           # 授精数
    preg_eligible: int  # 妊娠対象個体数（結果未確定の牛を除く）
    preg: int           # 妊娠数
    loss: int           # 損耗数（流産: outcome='A'）
    hdr: float          # 授精率（%）
    pr: float           # 妊娠率（%）


@dataclass
class DimPrResult:
    """DIMレンジ別妊娠率結果"""
    dim_start: int      # DIMレンジ開始（例: 50）
    dim_end: int        # DIMレンジ終了（9999 = 上限なし）
    br_el: int          # 繁殖対象個体数
    bred: int           # 授精数
    preg_eligible: int  # 妊娠対象個体数
    preg: int           # 妊娠数
    loss: int           # 損耗数（流産）
    hdr: float          # 授精率（%）
    pr: float           # 妊娠率（%）


class ReproductionAnalysis:
    """繁殖分析（DC305互換）"""

    CYCLE_DAYS = 21          # 21日サイクル
    DEFAULT_CYCLES = 18      # デフォルトで18サイクル遡る
    PREG_CONFIRM_DAYS = 28   # サイクル終了から28日未満は妊娠対象=0
    DEFAULT_VWP = 50         # デフォルトVWP（日数）

    def __init__(self, db: DBHandler, rule_engine: RuleEngine,
                 formula_engine: FormulaEngine, vwp: int = DEFAULT_VWP):
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
        self.EVENT_PDN = 302
        self.EVENT_PDP = 303
        self.EVENT_PDP2 = 304
        self.EVENT_ABRT = 305
        self.EVENT_PAGN = 306
        self.EVENT_PAGP = 307

        # 繁殖コード定義（RuleEngineと同じ値を使用）
        self.RC_STOPPED = RuleEngine.RC_STOPPED  # 1
        self.RC_FRESH = RuleEngine.RC_FRESH      # 2
        self.RC_BRED = RuleEngine.RC_BRED        # 3
        self.RC_OPEN = RuleEngine.RC_OPEN        # 4
        self.RC_PREGNANT = RuleEngine.RC_PREGNANT  # 5
        self.RC_DRY = RuleEngine.RC_DRY          # 6

    # ------------------------------------------------------------------
    # サイクル生成
    # ------------------------------------------------------------------

    def generate_cycles(self, period_start: Optional[str], period_end: Optional[str]) -> List[Cycle]:
        """
        21日サイクルを生成

        Args:
            period_start: 期間開始（YYYY-MM-DD、Noneの場合はデフォルト）
            period_end:   期間終了（YYYY-MM-DD、Noneの場合は今日）

        Returns:
            サイクルリスト（古い順）
        """
        if period_end:
            end_date = datetime.strptime(period_end, '%Y-%m-%d')
        else:
            end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        cycles = []

        if period_start:
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

    # ------------------------------------------------------------------
    # ユーティリティ
    # ------------------------------------------------------------------

    def _parse_json_data(self, json_data) -> dict:
        """json_dataフィールドをdictに変換"""
        if isinstance(json_data, dict):
            return json_data
        if isinstance(json_data, str):
            try:
                return json.loads(json_data)
            except Exception:
                return {}
        return {}

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """日付文字列をdatetimeに変換（失敗時はNone）"""
        if not date_str:
            return None
        try:
            return datetime.strptime(date_str, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None

    def _metrics_on_or_after_registration(self, cow: Dict[str, Any], reference_date: datetime) -> bool:
        """
        繁殖指標の集計起点（cow.entr 登録日）以降の基準日か。
        entr が未設定のときは従来どおり True（既存データ互換）。
        """
        entr = (cow.get('entr') or '').strip()
        if not entr:
            return True
        entr_dt = self._parse_date(entr)
        if entr_dt is None:
            return True
        return reference_date >= entr_dt

    def _get_cycle_events(self, all_events: List[Dict[str, Any]],
                          cycle: Cycle) -> Tuple[List, List]:
        """
        サイクル内イベントとサイクル前イベントを返す

        Returns:
            (before_cycle_events, cycle_events)
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')

        before = []
        within = []
        for e in all_events:
            edt = self._parse_date(e.get('event_date', ''))
            if edt is None:
                continue
            if edt < cycle_start:
                before.append(e)
            elif edt <= cycle_end:
                within.append(e)
        return before, within

    # ------------------------------------------------------------------
    # 繁殖対象判定
    # ------------------------------------------------------------------

    def is_breeding_eligible(self, cow_auto_id: int, cycle: Cycle,
                             all_events: List[Dict[str, Any]]) -> bool:
        """
        繁殖対象個体かどうかを判定（BR_EL）

        Args:
            cow_auto_id: 牛の auto_id
            cycle:       サイクル
            all_events:  牛の全イベント履歴

        Returns:
            True: 繁殖対象、False: 対象外
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        before_cycle_events, _ = self._get_cycle_events(all_events, cycle)

        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if cow and not self._metrics_on_or_after_registration(cow, cycle_start):
            return False

        # 1. サイクル開始前に除籍・繁殖停止されている場合は除外
        for event in before_cycle_events:
            if event.get('event_number') in [self.EVENT_SOLD, self.EVENT_DEAD, self.EVENT_STOPR]:
                return False

        # 2. VWPチェック（分娩からの日数 < VWP なら除外）
        if not cow:
            return False

        clvd = cow.get('clvd')
        if clvd:
            clvd_dt = self._parse_date(clvd)
            if clvd_dt is not None:
                days_since_calving = (cycle_start - clvd_dt).days
                if days_since_calving < self.vwp:
                    return False

        # 3. サイクル開始時点の繁殖コードを計算
        state = self._calculate_state_at_date(cow_auto_id, cycle_start, before_cycle_events)
        rc = state.get('rc')
        if rc in [self.RC_STOPPED, self.RC_PREGNANT]:
            return False

        return True

    # ------------------------------------------------------------------
    # 妊娠対象判定（DC305互換）
    # ------------------------------------------------------------------

    def _is_preg_eligible_cow(self, cow_auto_id: int, cycle: Cycle,
                               all_events: List[Dict[str, Any]]) -> bool:
        """
        妊娠対象個体かどうかを判定

        DC305ロジック: 繁殖対象牛のうち、
        「サイクル内でAI実施かつ結果未確定（outcome='N'またはNull）」の牛を除外。
        AI未実施の牛（空胎確定）は含む。
        """
        _, cycle_events = self._get_cycle_events(all_events, cycle)

        cycle_ais = [
            e for e in cycle_events
            if e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]
        ]

        if not cycle_ais:
            # AI未実施 → 空胎確定 → 妊娠対象に含む
            return True

        # AI有り: 結果未確定のAIが1件でもあれば対象外
        for ai in cycle_ais:
            jd = self._parse_json_data(ai.get('json_data'))
            outcome = jd.get('outcome') or 'N'
            if outcome not in ('P', 'O', 'R', 'A'):
                return False

        return True

    # ------------------------------------------------------------------
    # 妊娠数・損耗数カウント
    # ------------------------------------------------------------------

    def _count_preg_and_loss(self, cow_auto_id: int, cycle: Cycle,
                              all_events: List[Dict[str, Any]]) -> Tuple[int, int]:
        """
        妊娠数と損耗数を計算（牛単位: 0 or 1）

        AIイベントのoutcomeフィールドを使用:
          P = 妊娠確定   → preg=1
          A = 流産（確認妊娠後）→ preg=1, loss=1
          O/R = 不受胎・再種付 → preg=0, loss=0

        Returns:
            (preg, loss)
        """
        _, cycle_events = self._get_cycle_events(all_events, cycle)

        preg = 0
        loss = 0

        for e in cycle_events:
            if e.get('event_number') not in [self.EVENT_AI, self.EVENT_ET]:
                continue
            jd = self._parse_json_data(e.get('json_data'))
            outcome = jd.get('outcome') or 'N'

            if outcome in ('P', 'A'):
                preg = 1  # 牛単位なので最大1
            if outcome == 'A':
                loss = 1

        return preg, loss

    # ------------------------------------------------------------------
    # 授精数カウント
    # ------------------------------------------------------------------

    def count_bred(self, cow_auto_id: int, cycle: Cycle,
                   all_events: List[Dict[str, Any]]) -> int:
        """
        授精数（Bred）を計算

        outcome='R'（再種付）のイベントは除外。
        R以外のAI/ETイベントを1頭につき1件ずつカウント。
        """
        _, cycle_events = self._get_cycle_events(all_events, cycle)

        count = 0
        for event in cycle_events:
            if event.get('event_number') not in [self.EVENT_AI, self.EVENT_ET]:
                continue
            jd = self._parse_json_data(event.get('json_data'))
            outcome = jd.get('outcome')
            if outcome != 'R':
                count += 1

        return count

    # ------------------------------------------------------------------
    # 状態計算（簡易版）
    # ------------------------------------------------------------------

    def _calculate_state_at_date(self, cow_auto_id: int, target_date: datetime,
                                 events_before: List[Dict[str, Any]]) -> Dict[str, Any]:
        """指定日時点の状態を計算（簡易版）"""
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return {'rc': self.RC_OPEN}

        state = {
            'rc': cow.get('rc') or self.RC_OPEN,
            'clvd': cow.get('clvd'),
            'last_ai_date': None
        }

        sorted_events = sorted(events_before,
                               key=lambda e: (e.get('event_date', ''), e.get('id', 0)))

        for event in sorted_events:
            event_number = event.get('event_number')

            if event_number == self.EVENT_CALV:
                state['clvd'] = event.get('event_date')
                state['rc'] = self.RC_FRESH
            elif event_number in [self.EVENT_AI, self.EVENT_ET]:
                state['last_ai_date'] = event.get('event_date')
                if state['rc'] in [self.RC_FRESH, self.RC_OPEN]:
                    state['rc'] = self.RC_BRED
            elif event_number in [self.EVENT_PDP, self.EVENT_PDP2, self.EVENT_PAGP]:
                state['rc'] = self.RC_PREGNANT
            elif event_number in [self.EVENT_PDN, self.EVENT_PAGN]:
                state['rc'] = self.RC_OPEN
            elif event_number == self.EVENT_ABRT:
                state['rc'] = self.RC_OPEN
            elif event_number == self.EVENT_STOPR:
                state['rc'] = self.RC_STOPPED

        return state

    # ------------------------------------------------------------------
    # サイクル分析
    # ------------------------------------------------------------------

    def analyze_cycle(self, cycle: Cycle, cows: List[Dict[str, Any]],
                      today: datetime) -> ReproResult:
        """
        1サイクルの繁殖分析を実行

        Args:
            cycle: サイクル
            cows:  全牛リスト
            today: 今日の日付（28日ルール判定に使用）

        Returns:
            繁殖分析結果
        """
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')
        # サイクル終了から28日未満 → 妊娠確定不可のため妊娠対象=0
        recent_cycle = (today - cycle_end).days < self.PREG_CONFIRM_DAYS

        br_el = 0
        bred = 0
        preg_eligible = 0
        preg = 0
        loss = 0

        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue

            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)

            if not self.is_breeding_eligible(cow_auto_id, cycle, all_events):
                continue

            br_el += 1
            bred += self.count_bred(cow_auto_id, cycle, all_events)

            if not recent_cycle:
                if self._is_preg_eligible_cow(cow_auto_id, cycle, all_events):
                    preg_eligible += 1
                    p, l = self._count_preg_and_loss(cow_auto_id, cycle, all_events)
                    preg += p
                    loss += l

        hdr = round(bred / br_el * 100) if br_el > 0 else 0
        pr = round(preg / preg_eligible * 100) if preg_eligible > 0 else 0

        return ReproResult(
            cycle_number=cycle.cycle_number,
            start_date=cycle.start_date,
            end_date=cycle.end_date,
            br_el=br_el,
            bred=bred,
            preg_eligible=preg_eligible,
            preg=preg,
            loss=loss,
            hdr=float(hdr),
            pr=float(pr)
        )

    # ------------------------------------------------------------------
    # DIMレンジ生成
    # ------------------------------------------------------------------

    def generate_dim_ranges(self, n_ranges: int = 10) -> List[Tuple[int, int]]:
        """
        VWPを起点とした21日間隔のDIMレンジリストを生成

        Args:
            n_ranges: レンジ数（最後のレンジはオープンエンド）

        Returns:
            [(dim_start, dim_end), ...] のリスト（最後は dim_end=9999）
        """
        ranges = []
        for i in range(n_ranges - 1):
            dim_start = self.vwp + i * self.CYCLE_DAYS
            dim_end = dim_start + self.CYCLE_DAYS - 1
            ranges.append((dim_start, dim_end))
        last_start = self.vwp + (n_ranges - 1) * self.CYCLE_DAYS
        ranges.append((last_start, 9999))
        return ranges

    # ------------------------------------------------------------------
    # DIMレンジ別分析
    # ------------------------------------------------------------------

    def analyze_by_dim(self, period_start: Optional[str], period_end: Optional[str],
                       n_ranges: int = 10,
                       cow_filter: Optional[Callable[[Dict[str, Any]], bool]] = None) -> List[DimPrResult]:
        """
        DIMレンジ別妊娠率を分析（DC305 Bredsumer互換）

        各牛の現在の泌乳期（clvd基準）について、VWP以降の各21日DIMレンジで
        繁殖対象・授精・妊娠を集計する。

        Args:
            period_start: 分析期間開始（YYYY-MM-DD、Noneの場合は制限なし）
            period_end:   分析期間終了（YYYY-MM-DD、Noneの場合は今日）
            n_ranges:     DIMレンジ数（デフォルト10）
            cow_filter:   省略可。経産牛（lact>=1）に対し、True の牛だけを集計に含める。

        Returns:
            DimPrResultリスト（DIMレンジ昇順）
        """
        dim_ranges = self.generate_dim_ranges(n_ranges=n_ranges)
        # 経産牛のみ対象（未経産牛 lact=0 or None は除外）
        cows = [c for c in self.db.get_all_cows() if (c.get('lact') or 0) >= 1]
        if cow_filter:
            cows = [c for c in cows if cow_filter(c)]
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        ps_dt = self._parse_date(period_start) if period_start else None
        pe_dt = self._parse_date(period_end) if period_end else today

        # 結果格納用辞書
        counts: Dict[Tuple[int, int], Dict[str, int]] = {
            rng: {'br_el': 0, 'bred': 0, 'preg_eligible': 0, 'preg': 0, 'loss': 0}
            for rng in dim_ranges
        }

        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue

            clvd = cow.get('clvd')
            if not clvd:
                continue

            clvd_dt = self._parse_date(clvd)
            if not clvd_dt:
                continue

            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)

            for (dim_start, dim_end) in dim_ranges:
                # DIMレンジ開始がVWP未満なら除外
                if dim_start < self.vwp:
                    continue

                # このDIMレンジに対応する日付ウィンドウを計算
                window_start = clvd_dt + timedelta(days=dim_start)
                if dim_end == 9999:
                    # オープンエンド: 今日または期間終了日まで
                    window_end = min(today, pe_dt) if pe_dt else today
                else:
                    window_end = clvd_dt + timedelta(days=dim_end)

                # 登録日（entr）より前のウィンドウは繁殖指標の母数に含めない
                if not self._metrics_on_or_after_registration(cow, window_start):
                    continue

                # ウィンドウが今日より未来 → スキップ
                if window_start > today:
                    continue

                # ウィンドウが逆転している場合（オープンエンドで既に終了）→ スキップ
                if window_end < window_start:
                    continue

                # 期間フィルタ: 指定期間とウィンドウが重複しない場合はスキップ
                if ps_dt and window_end < ps_dt:
                    continue
                if pe_dt and window_start > pe_dt:
                    continue

                # ウィンドウ開始前のイベントを抽出
                before_window = [
                    e for e in all_events
                    if self._parse_date(e.get('event_date', '')) is not None
                    and self._parse_date(e.get('event_date')) < window_start
                ]

                # 除籍・繁殖停止チェック
                if any(e.get('event_number') in [self.EVENT_SOLD, self.EVENT_DEAD, self.EVENT_STOPR]
                       for e in before_window):
                    continue

                # ウィンドウ開始時点の繁殖コードを計算
                state = self._calculate_state_at_date(cow_auto_id, window_start, before_window)
                if state.get('rc') in [self.RC_STOPPED, self.RC_PREGNANT]:
                    continue

                counts[(dim_start, dim_end)]['br_el'] += 1

                # ウィンドウ内のAI/ETイベントを取得
                win_ai_events = [
                    e for e in all_events
                    if self._parse_date(e.get('event_date', '')) is not None
                    and window_start <= self._parse_date(e.get('event_date')) <= window_end
                    and e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]
                ]

                # 授精数（outcome='R' 再種付を除く）
                non_retry = [
                    e for e in win_ai_events
                    if self._parse_json_data(e.get('json_data')).get('outcome') != 'R'
                ]
                counts[(dim_start, dim_end)]['bred'] += len(non_retry)

                # 妊娠対象の判定
                # オープンエンドの場合は「今日」をウィンドウ終端として使う
                effective_end = window_end if dim_end != 9999 else today
                recent = (today - effective_end).days < self.PREG_CONFIRM_DAYS

                if not recent:
                    if not non_retry:
                        # 未授精 → 確定空胎
                        counts[(dim_start, dim_end)]['preg_eligible'] += 1
                    else:
                        # 全AIの結果が確定している場合のみ妊娠対象に含める
                        all_confirmed = all(
                            self._parse_json_data(e.get('json_data')).get('outcome')
                            in ('P', 'O', 'R', 'A')
                            for e in non_retry
                        )
                        if all_confirmed:
                            counts[(dim_start, dim_end)]['preg_eligible'] += 1
                            for e in non_retry:
                                outcome = self._parse_json_data(e.get('json_data')).get('outcome')
                                if outcome in ('P', 'A'):
                                    counts[(dim_start, dim_end)]['preg'] += 1
                                    if outcome == 'A':
                                        counts[(dim_start, dim_end)]['loss'] += 1
                                    break

        # DimPrResult リストに変換
        results: List[DimPrResult] = []
        for (dim_start, dim_end) in dim_ranges:
            c = counts[(dim_start, dim_end)]
            br_el = c['br_el']
            bred = c['bred']
            preg_eligible = c['preg_eligible']
            preg = c['preg']
            loss = c['loss']
            hdr = round(bred / br_el * 100) if br_el > 0 else 0
            pr = round(preg / preg_eligible * 100) if preg_eligible > 0 else 0
            results.append(DimPrResult(
                dim_start=dim_start,
                dim_end=dim_end,
                br_el=br_el,
                bred=bred,
                preg_eligible=preg_eligible,
                preg=preg,
                loss=loss,
                hdr=float(hdr),
                pr=float(pr)
            ))

        return results

    # ------------------------------------------------------------------
    # メイン分析エントリポイント
    # ------------------------------------------------------------------

    def get_cycle_cow_details(self, cycle: Cycle, cows: List[Dict[str, Any]],
                              today: datetime) -> List[Dict[str, Any]]:
        """
        サイクル内の繁殖対象個体の詳細リストを返す（ドリルダウン用）

        Returns:
            list of dict: auto_id, cow_id, lact, dim, category
            category: '未授精' | '授精済(未確定)' | '空胎' | '妊娠' | '流産'
        """
        cycle_start = datetime.strptime(cycle.start_date, '%Y-%m-%d')
        cycle_end = datetime.strptime(cycle.end_date, '%Y-%m-%d')
        recent_cycle = (today - cycle_end).days < self.PREG_CONFIRM_DAYS

        details = []
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue

            all_events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)

            if not self.is_breeding_eligible(cow_auto_id, cycle, all_events):
                continue

            # DIM計算（サイクル開始日基準）
            clvd = cow.get('clvd')
            dim = ''
            if clvd:
                clvd_dt = self._parse_date(clvd)
                if clvd_dt:
                    dim = (cycle_start - clvd_dt).days

            bred_count = self.count_bred(cow_auto_id, cycle, all_events)

            if recent_cycle or not self._is_preg_eligible_cow(cow_auto_id, cycle, all_events):
                category = '授精済(未確定)' if bred_count > 0 else '未授精'
            else:
                preg, loss = self._count_preg_and_loss(cow_auto_id, cycle, all_events)
                if loss == 1:
                    category = '流産'
                elif preg == 1:
                    category = '妊娠'
                elif bred_count > 0:
                    category = '空胎'
                else:
                    category = '未授精'

            # サイクル内の最後の非R授精イベントを取得
            _, cycle_events = self._get_cycle_events(all_events, cycle)
            cycle_ais = [
                e for e in cycle_events
                if e.get('event_number') in [self.EVENT_AI, self.EVENT_ET]
                and self._parse_json_data(e.get('json_data')).get('outcome') != 'R'
            ]
            last_ai = cycle_ais[-1] if cycle_ais else None
            ai_jd = self._parse_json_data(last_ai.get('json_data')) if last_ai else {}

            sire = ai_jd.get('sire', '') or ''
            technician_code = (ai_jd.get('technician') or
                               ai_jd.get('technician_code', '') or '')
            ai_type_code = (ai_jd.get('type') or
                            ai_jd.get('ai_type') or
                            ai_jd.get('insemination_type') or
                            ai_jd.get('inseminationType') or
                            ai_jd.get('insemination_type_code') or
                            ai_jd.get('inseminationTypeCode') or '')

            # 現泌乳期の累積授精回数（最終分娩以降・サイクル終了まで、非R）
            last_calv_dt = None
            for ev in sorted(all_events, key=lambda e: e.get('event_date', '')):
                if ev.get('event_number') == self.EVENT_CALV:
                    dt = self._parse_date(ev.get('event_date', ''))
                    if dt and dt <= cycle_end:
                        last_calv_dt = dt
            cum_count = 0
            for ev in all_events:
                if ev.get('event_number') not in [self.EVENT_AI, self.EVENT_ET]:
                    continue
                ev_dt = self._parse_date(ev.get('event_date', ''))
                if ev_dt is None or ev_dt > cycle_end:
                    continue
                if last_calv_dt and ev_dt < last_calv_dt:
                    continue
                if self._parse_json_data(ev.get('json_data')).get('outcome') != 'R':
                    cum_count += 1

            details.append({
                'auto_id': cow_auto_id,
                'cow_id': cow.get('cow_id', ''),
                'lact': cow.get('lact', ''),
                'dim': dim,
                'category': category,
                'sire': sire,
                'technician_code': str(technician_code).strip() if technician_code else '',
                'ai_type_code': str(ai_type_code).strip() if ai_type_code else '',
                'insemination_count': cum_count if bred_count > 0 else '',
            })

        _order = {'妊娠': 0, '流産': 1, '授精済(未確定)': 2, '空胎': 3, '未授精': 4}
        details.sort(key=lambda x: (_order.get(x['category'], 9), x.get('cow_id') or ''))
        return details

    def analyze(self, period_start: Optional[str], period_end: Optional[str],
                cow_filter: Optional[Callable[[Dict[str, Any]], bool]] = None) -> List[ReproResult]:
        """
        繁殖分析を実行（経産牛のみ: lact >= 1）

        Args:
            period_start: 期間開始（YYYY-MM-DD）
            period_end:   期間終了（YYYY-MM-DD）
            cow_filter:   省略可。経産牛（lact>=1）に対し、True の牛だけを集計に含める。

        Returns:
            繁殖分析結果リスト（各サイクル分、古い順）
        """
        cycles = self.generate_cycles(period_start, period_end)
        # 経産牛のみ対象（未経産牛 lact=0 or None は除外）
        cows = [c for c in self.db.get_all_cows() if (c.get('lact') or 0) >= 1]
        if cow_filter:
            cows = [c for c in cows if cow_filter(c)]
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        results = []
        for cycle in cycles:
            result = self.analyze_cycle(cycle, cows, today)
            results.append(result)

        return results
