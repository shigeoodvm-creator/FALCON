"""
FALCON2 - 受胎率分析モジュール
分析モード（受胎率分析）の実装
"""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List, Tuple
from datetime import date, datetime, timedelta
from pathlib import Path
import json
import logging

from db.db_handler import DBHandler

logger = logging.getLogger(__name__)


@dataclass
class InsemSequence:
    """授精シーケンス"""
    cow_auto_id: int
    parity: int
    seq_index: int  # 1,2,3...
    start_date: date
    end_date: date  # シーケンス内の最終AI/ET日
    events: List[Dict[str, Any]] = field(default_factory=list)  # 生イベント
    technician_code: Optional[str] = None  # 代表値：基本は start_event の値
    insemination_type_code: Optional[str] = None
    sire_code: Optional[str] = None
    conceived: bool = False  # 妊娠プラスが紐づいたら True
    conception_event_date: Optional[date] = None


@dataclass
class AnalysisRequest:
    """分析リクエスト"""
    group_key: str  # BY_MONTH, BY_TECHNICIAN, etc.
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    cycle_days: Optional[int] = None  # BY_DIM_CYCLE用


class EventDictionaryHelper:
    """event_dictionary.jsonのヘルパークラス"""
    
    def __init__(self, event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            event_dictionary_path: event_dictionary.jsonのパス
        """
        self.event_dictionary: Dict[str, Any] = {}
        self._load_event_dictionary(event_dictionary_path)
    
    def _load_event_dictionary(self, event_dictionary_path: Optional[Path]):
        """event_dictionary.jsonを読み込む"""
        if event_dictionary_path is None:
            from constants import CONFIG_DEFAULT_DIR
            event_dictionary_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
        
        if event_dictionary_path and event_dictionary_path.exists():
            try:
                with open(event_dictionary_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                logger.warning(f"event_dictionary.json読み込みエラー: {e}")
                self.event_dictionary = {}
    
    def is_insemination_event(self, event_number: int) -> bool:
        """
        授精イベント（AI/ET）かどうかを判定
        
        Args:
            event_number: イベント番号
        
        Returns:
            授精イベントの場合はTrue
        """
        event_key = str(event_number)
        if event_key not in self.event_dictionary:
            return False
        
        event_data = self.event_dictionary[event_key]
        alias = event_data.get("alias", "")
        category = event_data.get("category", "")
        
        # AI/ETはcategory="REPRODUCTION"かつalias="AI"または"ET"
        if category == "REPRODUCTION" and alias in ["AI", "ET"]:
            return True
        
        return False
    
    def is_preg_positive_event(self, event_number: int) -> bool:
        """
        妊娠プラスイベントかどうかを判定
        
        Args:
            event_number: イベント番号
        
        Returns:
            妊娠プラスイベントの場合はTrue
        """
        event_key = str(event_number)
        if event_key not in self.event_dictionary:
            return False
        
        event_data = self.event_dictionary[event_key]
        alias = event_data.get("alias", "")
        
        # 妊娠プラスイベント：PDP, PDP2, PAGP
        if alias in ["PDP", "PDP2", "PAGP"]:
            return True
        
        return False


class InseminationSequenceBuilder:
    """授精シーケンス生成クラス"""
    
    def __init__(self, event_helper: EventDictionaryHelper):
        """
        初期化
        
        Args:
            event_helper: EventDictionaryHelperインスタンス
        """
        self.event_helper = event_helper
    
    def build_sequences(
        self, cow_auto_id: int, events: List[Dict[str, Any]], 
        cow_parity_map: Optional[Dict[int, int]] = None
    ) -> List[InsemSequence]:
        """
        牛1頭の授精シーケンスを生成
        
        Args:
            cow_auto_id: 牛のauto_id
            events: イベント一覧（全期間）
            cow_parity_map: {cow_auto_id: parity} のマッピング（イベントにparityがない場合の補完用）
        
        Returns:
            授精シーケンスのリスト
        """
        # ① AI/ETイベントのみ抽出し、event_date昇順にソート
        insem_events = []
        for event in events:
            event_number = event.get('event_number')
            if event_number and self.event_helper.is_insemination_event(event_number):
                event_date_str = event.get('event_date')
                if event_date_str:
                    try:
                        event_date = datetime.strptime(event_date_str, '%Y-%m-%d').date()
                        insem_events.append({
                            'event': event,
                            'event_date': event_date,
                            'event_date_str': event_date_str
                        })
                    except (ValueError, TypeError):
                        continue
        
        if not insem_events:
            return []
        
        # event_date昇順にソート
        insem_events.sort(key=lambda x: x['event_date'])
        
        # ② parityごとに分割（同一parity内でのみ7日ルール適用）
        # イベントからparityを取得、なければcow_parity_mapから補完
        parity_events_map = {}  # {parity: [events]}
        
        for insem_event in insem_events:
            event = insem_event['event']
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # parityを取得
            parity = json_data.get('parity') or json_data.get('lactation')
            if parity is None and cow_parity_map:
                parity = cow_parity_map.get(cow_auto_id)
            
            if parity is None:
                # parityが取得できない場合はスキップ
                continue
            
            if parity not in parity_events_map:
                parity_events_map[parity] = []
            parity_events_map[parity].append(insem_event)
        
        # ③ 各parityごとにシーケンスを生成
        all_sequences = []
        
        for parity, parity_events in sorted(parity_events_map.items()):
            sequences = self._build_sequences_for_parity(
                cow_auto_id, parity, parity_events
            )
            all_sequences.extend(sequences)
        
        return all_sequences
    
    def _build_sequences_for_parity(
        self, cow_auto_id: int, parity: int, events: List[Dict[str, Any]]
    ) -> List[InsemSequence]:
        """
        同一parity内でシーケンスを生成
        
        Args:
            cow_auto_id: 牛のauto_id
            parity: 産次
            events: 同一parityの授精イベントリスト（日付昇順）
        
        Returns:
            授精シーケンスのリスト
        """
        sequences = []
        seq_index = 0
        last_date = None
        
        for event_data in events:
            event = event_data['event']
            event_date = event_data['event_date']
            event_date_str = event_data['event_date_str']
            
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            if last_date is None:
                # 新規シーケンス
                seq_index += 1
                seq = InsemSequence(
                    cow_auto_id=cow_auto_id,
                    parity=parity,
                    seq_index=seq_index,
                    start_date=event_date,
                    end_date=event_date,
                    events=[event]
                )
                # 代表値を設定
                self._set_representative_values(seq, json_data)
                sequences.append(seq)
                last_date = event_date
            else:
                days_diff = (event_date - last_date).days
                if days_diff <= 7:
                    # 同一シーケンスに追加
                    seq = sequences[-1]
                    seq.events.append(event)
                    seq.end_date = event_date
                    # 代表値がNoneの場合は更新
                    self._update_representative_values(seq, json_data)
                else:
                    # 新規シーケンス
                    seq_index += 1
                    seq = InsemSequence(
                        cow_auto_id=cow_auto_id,
                        parity=parity,
                        seq_index=seq_index,
                        start_date=event_date,
                        end_date=event_date,
                        events=[event]
                    )
                    self._set_representative_values(seq, json_data)
                    sequences.append(seq)
                last_date = event_date
        
        return sequences
    
    def _set_representative_values(self, seq: InsemSequence, json_data: Dict[str, Any]):
        """シーケンスの代表値を設定"""
        # technician_code
        technician = json_data.get('technician_code') or json_data.get('technician')
        if technician:
            seq.technician_code = technician
        
        # insemination_type_code
        insem_type = json_data.get('insemination_type_code') or json_data.get('type') or json_data.get('ai_type')
        if insem_type:
            seq.insemination_type_code = insem_type
        
        # sire_code
        sire = json_data.get('sire_code') or json_data.get('sire')
        if sire:
            seq.sire_code = sire
    
    def _update_representative_values(self, seq: InsemSequence, json_data: Dict[str, Any]):
        """シーケンスの代表値を更新（Noneの場合のみ）"""
        # technician_code
        if not seq.technician_code:
            technician = json_data.get('technician_code') or json_data.get('technician')
            if technician:
                seq.technician_code = technician
        
        # insemination_type_code
        if not seq.insemination_type_code:
            insem_type = json_data.get('insemination_type_code') or json_data.get('type') or json_data.get('ai_type')
            if insem_type:
                seq.insemination_type_code = insem_type
        
        # sire_code
        if not seq.sire_code:
            sire = json_data.get('sire_code') or json_data.get('sire')
            if sire:
                seq.sire_code = sire


class ConceptionLinker:
    """妊娠プラス紐づけクラス"""
    
    def __init__(self, event_helper: EventDictionaryHelper):
        """
        初期化
        
        Args:
            event_helper: EventDictionaryHelperインスタンス
        """
        self.event_helper = event_helper
    
    def link_conceptions(
        self, cow_auto_id: int, sequences: List[InsemSequence],
        events: List[Dict[str, Any]]
    ):
        """
        妊娠プラスイベントを授精シーケンスに紐づける
        
        Args:
            cow_auto_id: 牛のauto_id
            sequences: 授精シーケンスのリスト
            events: イベント一覧（全期間）
        """
        # 妊娠プラスイベントを抽出し、日付昇順にソート
        preg_events = []
        for event in events:
            event_number = event.get('event_number')
            if event_number and self.event_helper.is_preg_positive_event(event_number):
                event_date_str = event.get('event_date')
                if event_date_str:
                    try:
                        event_date = datetime.strptime(event_date_str, '%Y-%m-%d').date()
                        preg_events.append({
                            'event': event,
                            'event_date': event_date
                        })
                    except (ValueError, TypeError):
                        continue
        
        if not preg_events:
            return
        
        # 日付昇順にソート
        preg_events.sort(key=lambda x: x['event_date'])
        
        # 各妊娠プラスイベントについて、直近の授精シーケンスに紐づけ
        for preg_event in preg_events:
            preg_date = preg_event['event_date']
            
            # event_dateより前（<=）に開始した「直近の授精シーケンス」を探す
            candidate_seq = None
            for seq in sequences:
                if seq.start_date <= preg_date:
                    if candidate_seq is None or seq.start_date > candidate_seq.start_date:
                        candidate_seq = seq
            
            if candidate_seq and not candidate_seq.conceived:
                # 紐づけ（1シーケンス最大1受胎）
                candidate_seq.conceived = True
                candidate_seq.conception_event_date = preg_date


class ConceptionRateAnalyzer:
    """受胎率分析クラス"""
    
    def __init__(self, db_handler: DBHandler, event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            db_handler: DBHandlerインスタンス
            event_dictionary_path: event_dictionary.jsonのパス
        """
        self.db = db_handler
        self.event_helper = EventDictionaryHelper(event_dictionary_path)
        self.sequence_builder = InseminationSequenceBuilder(self.event_helper)
        self.conception_linker = ConceptionLinker(self.event_helper)
        
        # シーケンスキャッシュ（メモリ）
        self._sequence_cache: Dict[int, List[InsemSequence]] = {}  # {cow_auto_id: sequences}
    
    def analyze(self, request: AnalysisRequest) -> Dict[str, Any]:
        """
        受胎率分析を実行
        
        Args:
            request: 分析リクエスト
        
        Returns:
            分析結果辞書
            {
                "rows": [
                    {"group": "...", "rate": 42.9, "conceptions": 12, "inseminations": 28, "spc": 2.3},
                    ...
                ],
                "total": {"rate": 38.8, "conceptions": 50, "inseminations": 129, "spc": 2.5}
            }
        """
        # 全牛を取得
        cows = self.db.get_all_cows()
        
        # 各牛のparityマッピングを作成（イベントにparityがない場合の補完用）
        cow_parity_map = {}
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            lact = cow.get('lact', 0) or 0
            if cow_auto_id:
                cow_parity_map[cow_auto_id] = lact
        
        # 全牛についてシーケンスを生成（キャッシュ利用）
        all_sequences = []
        for cow in cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # キャッシュチェック
            if cow_auto_id in self._sequence_cache:
                sequences = self._sequence_cache[cow_auto_id]
            else:
                # イベント履歴を取得
                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                
                # シーケンス生成
                sequences = self.sequence_builder.build_sequences(
                    cow_auto_id, events, cow_parity_map
                )
                
                # 妊娠プラス紐づけ
                self.conception_linker.link_conceptions(
                    cow_auto_id, sequences, events
                )
                
                # キャッシュに保存
                self._sequence_cache[cow_auto_id] = sequences
            
            all_sequences.extend(sequences)
        
        # 期間フィルタ（seq.start_date基準）
        if request.period_start or request.period_end:
            filtered_sequences = []
            for seq in all_sequences:
                if request.period_start and seq.start_date < request.period_start:
                    continue
                if request.period_end and seq.start_date > request.period_end:
                    continue
                filtered_sequences.append(seq)
            all_sequences = filtered_sequences
        
        # グループごとに集計
        group_stats = {}  # {group_value: {inseminations: int, conceptions: int}}
        
        for seq in all_sequences:
            group_value = self._get_group_value(seq, request)
            if group_value is None:
                continue
            
            if group_value not in group_stats:
                group_stats[group_value] = {'inseminations': 0, 'conceptions': 0}
            
            group_stats[group_value]['inseminations'] += 1
            if seq.conceived:
                group_stats[group_value]['conceptions'] += 1
        
        # 結果をリスト形式に変換
        rows = []
        for group_value in sorted(group_stats.keys()):
            stats = group_stats[group_value]
            inseminations = stats['inseminations']
            conceptions = stats['conceptions']
            rate = (conceptions / inseminations * 100) if inseminations > 0 else 0.0
            spc = (inseminations / conceptions) if conceptions > 0 else None
            
            rows.append({
                'group': group_value,
                'rate': round(rate, 1),
                'conceptions': conceptions,
                'inseminations': inseminations,
                'spc': round(spc, 1) if spc is not None else None
            })
        
        # TOTAL行を計算
        total_inseminations = sum(s['inseminations'] for s in group_stats.values())
        total_conceptions = sum(s['conceptions'] for s in group_stats.values())
        total_rate = (total_conceptions / total_inseminations * 100) if total_inseminations > 0 else 0.0
        total_spc = (total_inseminations / total_conceptions) if total_conceptions > 0 else None
        
        return {
            'rows': rows,
            'total': {
                'rate': round(total_rate, 1),
                'conceptions': total_conceptions,
                'inseminations': total_inseminations,
                'spc': round(total_spc, 1) if total_spc is not None else None
            }
        }
    
    def _get_group_value(self, seq: InsemSequence, request: AnalysisRequest) -> Optional[str]:
        """
        シーケンスからグループ値を取得
        
        Args:
            seq: 授精シーケンス
            request: 分析リクエスト
        
        Returns:
            グループ値（文字列）、取得できない場合はNone
        """
        group_key = request.group_key
        
        if group_key == "BY_MONTH":
            # 月ごと（seq.start_date基準）
            return seq.start_date.strftime('%Y-%m')
        
        elif group_key == "BY_TECHNICIAN":
            # 授精師ごと
            return seq.technician_code if seq.technician_code else "Unknown"
        
        elif group_key == "BY_INSEMINATION_TYPE":
            # 授精種類別
            return seq.insemination_type_code if seq.insemination_type_code else "Unknown"
        
        elif group_key == "BY_SIRE":
            # SIREごと
            return seq.sire_code if seq.sire_code else "Unknown"
        
        elif group_key == "BY_PARITY":
            # 産次ごと
            return str(seq.parity)
        
        elif group_key == "BY_SERVICE_NUMBER":
            # 授精回数別
            if seq.seq_index == 1:
                return "1st Service"
            elif seq.seq_index == 2:
                return "2nd Service"
            elif seq.seq_index == 3:
                return "3rd Service"
            else:
                return "4th+ Service"
        
        elif group_key == "BY_DIM_CYCLE":
            # 分娩後サイクル別
            cycle_days = request.cycle_days or 10
            
            # DIMを計算（seq.start_date時点のDIM）
            dim = self._calculate_dim_at_date(seq.cow_auto_id, seq.start_date)
            if dim is None or dim < 0:
                return None
            
            cycle_start = (dim // cycle_days) * cycle_days
            cycle_end = cycle_start + cycle_days - 1
            
            if cycle_start >= 60:
                return "60+"
            else:
                return f"{cycle_start}–{cycle_end}"
        
        return None
    
    def _calculate_dim_at_date(self, cow_auto_id: int, target_date: date) -> Optional[int]:
        """
        指定日時点のDIMを計算
        
        Args:
            cow_auto_id: 牛のauto_id
            target_date: 対象日
        
        Returns:
            DIM（日数）、計算できない場合はNone
        """
        # 牛データを取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return None
        
        # 分娩日を取得（イベントから最新の分娩日を取得）
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 分娩イベント（CALV）を検索
        calving_dates = []
        for event in events:
            event_number = event.get('event_number')
            event_date_str = event.get('event_date')
            
            # 分娩イベントの判定（event_dictionaryから）
            event_key = str(event_number) if event_number else ""
            if event_key in self.event_helper.event_dictionary:
                event_data = self.event_helper.event_dictionary[event_key]
                if event_data.get("alias") == "CALV" and event_date_str:
                    try:
                        calv_date = datetime.strptime(event_date_str, '%Y-%m-%d').date()
                        if calv_date <= target_date:
                            calving_dates.append(calv_date)
                    except (ValueError, TypeError):
                        continue
        
        if not calving_dates:
            # イベントから取得できない場合はcowテーブルのclvdを使用
            clvd = cow.get('clvd')
            if clvd:
                try:
                    calv_date = datetime.strptime(clvd, '%Y-%m-%d').date()
                    if calv_date <= target_date:
                        calving_dates.append(calv_date)
                except (ValueError, TypeError):
                    pass
        
        if not calving_dates:
            return None
        
        # 最新の分娩日を使用
        latest_calv = max(calving_dates)
        dim = (target_date - latest_calv).days
        
        return dim if dim >= 0 else None


class ResultFormatter:
    """結果フォーマッター（TSV出力）"""
    
    @staticmethod
    def format_tsv(result: Dict[str, Any]) -> str:
        """
        TSV形式で結果をフォーマット
        
        Args:
            result: 分析結果辞書
        
        Returns:
            TSV形式の文字列
        """
        rows = result.get('rows', [])
        total = result.get('total', {})
        
        lines = []
        # ヘッダー
        lines.append("Group\tRate(%)\tn/N\tSPC")
        
        # データ行
        for row in rows:
            group = row.get('group', '')
            rate = row.get('rate', 0.0)
            conceptions = row.get('conceptions', 0)
            inseminations = row.get('inseminations', 0)
            spc = row.get('spc')
            
            n_n_str = f"{conceptions} / {inseminations}"
            spc_str = f"{spc:.1f}" if spc is not None else "-"
            
            lines.append(f"{group}\t{rate:.1f}\t{n_n_str}\t{spc_str}")
        
        # TOTAL行
        total_rate = total.get('rate', 0.0)
        total_conceptions = total.get('conceptions', 0)
        total_inseminations = total.get('inseminations', 0)
        total_spc = total.get('spc')
        
        total_n_n_str = f"{total_conceptions} / {total_inseminations}"
        total_spc_str = f"{total_spc:.1f}" if total_spc is not None else "-"
        
        lines.append(f"TOTAL\t{total_rate:.1f}\t{total_n_n_str}\t{total_spc_str}")
        
        return "\n".join(lines)





















