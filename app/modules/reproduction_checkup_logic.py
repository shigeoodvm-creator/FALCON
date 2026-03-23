"""
FALCON2 - 繁殖検診抽出ロジック
繁殖検診表の抽出条件を実装
"""

from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from pathlib import Path
import json

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


class ReproductionCheckupLogic:
    """繁殖検診抽出ロジック"""
    
    # 検診コード定義
    CHECKUP_FRESH = "ﾌﾚｯｼｭ"
    CHECKUP_REPRO1 = "繁殖1"
    CHECKUP_REPRO2 = "繁殖2"
    CHECKUP_PREG = "妊鑑"
    CHECKUP_REPREG = "再妊"
    CHECKUP_PREG2 = "妊鑑2"
    CHECKUP_DUE_OVER = "分娩？"
    CHECKUP_CHECK = "チェック"
    CHECKUP_HEIFER_REPRO1 = "育繁1"
    CHECKUP_HEIFER_REPRO2 = "育繁2"
    CHECKUP_HEIFER_PREG = "育妊"
    CHECKUP_HEIFER_REPREG = "育再妊"
    
    def __init__(self, db_handler: DBHandler, event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            event_dictionary_path: event_dictionary.json のパス
        """
        self.db = db_handler
        self.event_dict_path = event_dictionary_path
        self.event_dictionary: Dict[str, Dict[str, Any]] = {}
        self._load_event_dictionary()
    
    def _load_event_dictionary(self):
        """event_dictionary.json を読み込む"""
        if self.event_dict_path and self.event_dict_path.exists():
            try:
                with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
            except Exception as e:
                print(f"event_dictionary.json 読み込みエラー: {e}")
                self.event_dictionary = {}
    
    def _get_event_numbers_by_category(self, category: str) -> List[int]:
        """
        カテゴリに基づいてイベント番号を取得
        
        Args:
            category: イベントカテゴリ（例: "REPRODUCTION", "PREGNANCY"）
        
        Returns:
            イベント番号のリスト
        """
        event_numbers = []
        for event_str, event_data in self.event_dictionary.items():
            if event_data.get('category') == category:
                try:
                    event_numbers.append(int(event_str))
                except ValueError:
                    pass
        return event_numbers
    
    def _get_ai_event_numbers(self) -> List[int]:
        """AIイベント番号を取得"""
        # event_dictionaryから取得、なければデフォルト値
        ai_events = []
        for event_str, event_data in self.event_dictionary.items():
            if event_data.get('alias') == 'AI' or event_data.get('name_jp') == 'AI':
                try:
                    ai_events.append(int(event_str))
                except ValueError:
                    pass
        if not ai_events:
            ai_events = [RuleEngine.EVENT_AI]  # デフォルト: 200
        return ai_events
    
    def _get_et_event_numbers(self) -> List[int]:
        """ETイベント番号を取得"""
        et_events = []
        for event_str, event_data in self.event_dictionary.items():
            if event_data.get('alias') == 'ET' or event_data.get('name_jp') == 'ET':
                try:
                    et_events.append(int(event_str))
                except ValueError:
                    pass
        if not et_events:
            et_events = [RuleEngine.EVENT_ET]  # デフォルト: 201
        return et_events
    
    def _get_pregnancy_negative_event_numbers(self) -> List[int]:
        """妊娠鑑定マイナスイベント番号を取得"""
        pdn_events = []
        for event_str, event_data in self.event_dictionary.items():
            if (event_data.get('category') == 'PREGNANCY' and 
                event_data.get('outcome') == 'NEGATIVE'):
                try:
                    pdn_events.append(int(event_str))
                except ValueError:
                    pass
        if not pdn_events:
            pdn_events = [RuleEngine.EVENT_PDN, RuleEngine.EVENT_PAGN]  # デフォルト: 302, 306
        return pdn_events
    
    def _get_pregnancy_positive_event_numbers(self) -> List[int]:
        """妊娠鑑定プラスイベント番号を取得"""
        pdp_events = []
        for event_str, event_data in self.event_dictionary.items():
            if (event_data.get('category') == 'PREGNANCY' and 
                event_data.get('outcome') == 'POSITIVE'):
                try:
                    pdp_events.append(int(event_str))
                except ValueError:
                    pass
        if not pdp_events:
            pdp_events = [RuleEngine.EVENT_PDP, RuleEngine.EVENT_PDP2, RuleEngine.EVENT_PAGP]  # デフォルト: 303, 304, 307
        return pdp_events
    
    def _get_fresh_check_event_numbers(self) -> List[int]:
        """フレッシュチェックイベント番号を取得"""
        fchk_events = []
        for event_str, event_data in self.event_dictionary.items():
            if (event_data.get('alias') == 'FCHK' or 
                event_data.get('name_jp') == 'フレッシュチェック' or
                event_data.get('name_jp') == 'フレッシュ'):
                try:
                    fchk_events.append(int(event_str))
                except ValueError:
                    pass
        if not fchk_events:
            fchk_events = [300]  # デフォルト: 300
        return fchk_events
    
    def filter_events_current_lactation(self, events: List[Dict[str, Any]], cow: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        当該産次のイベントのみに絞る（産次更新＝分娩後は前産次の検診・AI等を参照しないため）。
        
        分娩イベントにより産次が更新された場合、前回検診日・前回検診結果・最終AI/ET日・
        授精後日数は当該産次のイベントのみから算出する。前産次のイベントは除外する。
        
        Args:
            events: イベント履歴
            cow: 牛データ（lact, clvd を使用）
        
        Returns:
            当該産次のイベントのみ（経産牛はevent_lact一致またはclvd以降、未経産は全件）
        """
        lact = cow.get('lact') or 0
        
        # 未経産牛（LACT == 0）：全イベントを対象
        if lact < 1:
            return list(events)
        
        # 経産牛（LACT >= 1）：当該産次のイベントのみ
        # 1) event_lact がある場合はそれでフィルタ（推奨・より正確）
        has_event_lact = any(e.get('event_lact') is not None for e in events)
        if has_event_lact:
            return [e for e in events if e.get('event_lact') == lact]
        
        # 2) event_lact がない場合は clvd 以降の日付でフィルタ（後方互換）
        clvd = cow.get('clvd')
        if not clvd:
            return list(events)
        return [e for e in events if e.get('event_date') and e.get('event_date') >= clvd]

    def extract_cows(self, checkup_date: str, settings: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        繁殖検診対象牛を抽出
        
        検診コードは当該産次のイベントのみで判定する（産次更新後は前回検診・最終AI等をリセット）。
        
        Args:
            checkup_date: 検診予定日（YYYY-MM-DD形式）
            settings: 繁殖検診設定（各検診区分の有効/無効、基準日数/月齢）
        
        Returns:
            抽出された牛のリスト（各牛にcheckup_codeが追加される）
        """
        # 全牛を取得
        all_cows = self.db.get_all_cows()
        
        results = []
        
        for cow in all_cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # 繁殖コードを取得して、繁殖停止の個体は除外
            rc = cow.get('rc')
            if rc == RuleEngine.RC_STOPPED:
                continue
            
            # イベント履歴を取得
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            # 当該産次のイベントのみで検診コードを判定（分娩後は前産次の検診・AIは使わない）
            events_this_lact = self.filter_events_current_lactation(events, cow)
            
            # 産次を取得
            lact = cow.get('lact') or 0
            
            # 経産牛（LACT >= 1）の場合
            if lact >= 1:
                checkup_code = self._check_parous_cow(cow, events_this_lact, checkup_date, settings)
            # 未経産牛（LACT == 0）の場合
            else:
                checkup_code = self._check_heifer_cow(cow, events_this_lact, checkup_date, settings)
            
            # 検診コードが設定された場合のみ結果に追加
            if checkup_code:
                cow_result = cow.copy()
                cow_result['checkup_code'] = checkup_code
                results.append(cow_result)
        
        return results
    
    def _check_parous_cow(self, cow: Dict[str, Any], events: List[Dict[str, Any]], 
                          checkup_date: str, settings: Dict[str, Any]) -> Optional[str]:
        """
        経産牛の検診コードを判定
        
        Args:
            cow: 牛データ
            events: イベント履歴
            checkup_date: 検診予定日
            settings: 設定
        
        Returns:
            検診コード（該当しない場合はNone）
        """
        lact = cow.get('lact') or 0
        if lact < 1:
            return None
        
        # フレッシュチェックイベントの有無を確認
        fchk_events = self._get_fresh_check_event_numbers()
        has_fresh_check = any(e.get('event_number') in fchk_events for e in events)
        
        # より具体的な条件を先にチェック（優先順位順）
        # 1) 繁殖検査２（AI/ET後、妊娠鑑定マイナス、繁殖コードOpen）
        if settings.get('repro2', {}).get('enabled', True):
            if self._check_repro2(events, cow):
                return self.CHECKUP_REPRO2
        
        # 2) 繁殖検査１（フレッシュチェック後、AI/ETなし）
        if settings.get('repro1', {}).get('enabled', True):
            if self._check_repro1(events):
                return self.CHECKUP_REPRO1
        
        # 3) 再妊娠鑑定（繁殖コードPregnant、授精後N日以上、妊娠イベント1回のみ）
        # より具体的な条件なので、妊娠鑑定より先にチェック
        if settings.get('repreg', {}).get('enabled', True):
            threshold = settings.get('repreg', {}).get('days') or 60
            if threshold is not None and self._check_re_pregnancy_check(events, cow, checkup_date, threshold):
                return self.CHECKUP_REPREG
        
        # 4) 妊娠鑑定（繁殖コードがPregnantでない、授精後N日以上）
        if settings.get('preg', {}).get('enabled', True):
            threshold = settings.get('preg', {}).get('days') or 30
            if threshold is not None and self._check_pregnancy_check(events, cow, checkup_date, threshold):
                return self.CHECKUP_PREG
        
        # 5) 任意妊娠鑑定（繁殖コードPregnant、授精後N日以上、妊娠イベント回数制限なし）
        if settings.get('preg2', {}).get('enabled', True):
            threshold = settings.get('preg2', {}).get('days')
            if threshold is not None:
                if self._check_optional_pregnancy_check(events, cow, checkup_date, threshold):
                    return self.CHECKUP_PREG2
        
        # 6) 分娩予定超過
        if settings.get('due_over', {}).get('enabled', True):
            threshold = settings.get('due_over', {}).get('days') or 14
            if threshold is not None and self._check_due_over(cow, events, checkup_date, threshold):
                return self.CHECKUP_DUE_OVER
        
        # 7) チェック
        if settings.get('check', {}).get('enabled', True):
            if self._check_check(events):
                return self.CHECKUP_CHECK
        
        # 8) フレッシュチェック（フレッシュチェックイベントがない場合のみ、分娩後N日以上）
        # ただし、妊娠中（繁殖コードPregnant）の個体は除外
        if settings.get('fresh', {}).get('enabled', True):
            # 繁殖コードがFresh（分娩後）であることを確認
            rc = cow.get('rc')
            # 妊娠中（Pregnant）の個体は除外
            if rc == RuleEngine.RC_FRESH:
                if not has_fresh_check:
                    dim = self._calculate_dim(cow, checkup_date)
                    if dim is not None:
                        threshold = settings.get('fresh', {}).get('days') or 30
                        if threshold is not None and dim >= threshold:
                            return self.CHECKUP_FRESH
        
        return None
    
    def _check_heifer_cow(self, cow: Dict[str, Any], events: List[Dict[str, Any]], 
                          checkup_date: str, settings: Dict[str, Any]) -> Optional[str]:
        """
        未経産牛の検診コードを判定
        
        Args:
            cow: 牛データ
            events: イベント履歴
            checkup_date: 検診予定日
            settings: 設定
        
        Returns:
            検診コード（該当しない場合はNone）
        """
        lact = cow.get('lact') or 0
        if lact != 0:
            return None
        
        # 1) 育成繁殖１
        if settings.get('heifer_repro1', {}).get('enabled', True):
            threshold = settings.get('heifer_repro1', {}).get('age_months') or 12.0
            if threshold is not None and self._check_heifer_repro1(cow, events, checkup_date, threshold):
                return self.CHECKUP_HEIFER_REPRO1
        
        # 2) 育成繁殖２
        if settings.get('heifer_repro2', {}).get('enabled', True):
            if self._check_heifer_repro2(events, cow):
                return self.CHECKUP_HEIFER_REPRO2
        
        # 3) 育成再妊娠鑑定（繁殖コードPregnant、授精後N日以上、妊娠イベント1回のみ）
        # より具体的な条件なので、育成妊娠鑑定より先にチェック
        if settings.get('heifer_repreg', {}).get('enabled', True):
            threshold = settings.get('heifer_repreg', {}).get('days') or 60
            if threshold is not None and self._check_re_pregnancy_check(events, cow, checkup_date, threshold):
                return self.CHECKUP_HEIFER_REPREG
        
        # 4) 育成妊娠鑑定（繁殖コードがPregnantでない、授精後N日以上）
        if settings.get('heifer_preg', {}).get('enabled', True):
            threshold = settings.get('heifer_preg', {}).get('days') or 30
            if threshold is not None and self._check_pregnancy_check(events, cow, checkup_date, threshold):
                return self.CHECKUP_HEIFER_PREG
        
        return None
    
    def _calculate_dim(self, cow: Dict[str, Any], checkup_date: str) -> Optional[int]:
        """分娩後日数を計算"""
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
            checkup_dt = datetime.strptime(checkup_date, '%Y-%m-%d')
            dim = (checkup_dt - clvd_dt).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_age_months(self, cow: Dict[str, Any], checkup_date: str) -> Optional[float]:
        """月齢を計算"""
        bthd = cow.get('bthd')
        if not bthd:
            return None
        
        try:
            birth_dt = datetime.strptime(bthd, '%Y-%m-%d')
            checkup_dt = datetime.strptime(checkup_date, '%Y-%m-%d')
            days_diff = (checkup_dt - birth_dt).days
            if days_diff < 0:
                return None
            # 月齢に変換（1ヶ月 = 30.44日）
            age_months = days_diff / 30.44
            return age_months
        except (ValueError, TypeError):
            return None
    
    def _get_last_ai_date(self, events: List[Dict[str, Any]], checkup_date: Optional[str] = None) -> Optional[str]:
        """
        最終授精日を取得（AI/ETイベント）
        
        基準日を起算日として、基準日時点での最新のAI/ETイベントを取得する。
        基準日が未来の場合も正しく動作する（基準日より後のイベントは除外される）。
        
        Args:
            events: イベントリスト
            checkup_date: 基準日（YYYY-MM-DD形式）。指定された場合、この日以前のイベントのみを対象とする。
                         基準日が未来の場合も正しく動作する。
        
        Returns:
            最終授精日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        ai_events = self._get_ai_event_numbers()
        et_events = self._get_et_event_numbers()
        all_insemination_events = ai_events + et_events
        
        dates = []
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if event_number in all_insemination_events and event_date:
                # 基準日が指定されている場合、基準日以前のイベントのみを対象とする
                # 基準日が未来の場合も正しく動作する（基準日より後のイベントは除外）
                if checkup_date:
                    if event_date > checkup_date:
                        continue
                dates.append(event_date)
        
        return max(dates) if dates else None
    
    def _get_insemination_date_for_dai(self, events: List[Dict[str, Any]], 
                                       last_ai_date: str) -> Optional[str]:
        """
        授精後日数計算用の授精日を取得（ETの場合は7日前）
        
        Args:
            events: イベント履歴
            last_ai_date: 最終授精日
        
        Returns:
            計算用授精日（ETの場合は7日前、AIの場合は当日）
        """
        if not last_ai_date:
            return None
        
        # 該当日付のイベントを探す
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if event_date == last_ai_date:
                # ETイベントの場合は7日前
                if event_number in self._get_et_event_numbers():
                    try:
                        event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                        insemination_dt = event_dt - timedelta(days=7)
                        return insemination_dt.strftime('%Y-%m-%d')
                    except (ValueError, TypeError):
                        pass
                # AIイベントの場合は当日
                elif event_number in self._get_ai_event_numbers():
                    return event_date
        
        return last_ai_date
    
    def _calculate_dai(self, events: List[Dict[str, Any]], checkup_date: str) -> Optional[int]:
        """
        授精後日数を計算（基準日からの起算）
        
        基準日を起算日として、基準日時点での最新のAI/ETイベントから授精後日数を計算する。
        基準日が未来の場合も正しく動作する。
        
        例：
        - 基準日: 2025-01-23（未来）
        - AI日: 2024-12-24
        - 授精後日数: (2025-01-23 - 2024-12-24).days = 30日
        
        Args:
            events: イベントリスト
            checkup_date: 基準日（YYYY-MM-DD形式）。起算日として使用される。
        
        Returns:
            授精後日数（日数）、計算できない場合はNone
        """
        last_ai_date = self._get_last_ai_date(events, checkup_date)
        if not last_ai_date:
            return None
        
        insemination_date = self._get_insemination_date_for_dai(events, last_ai_date)
        if not insemination_date:
            return None
        
        try:
            insemination_dt = datetime.strptime(insemination_date, '%Y-%m-%d')
            checkup_dt = datetime.strptime(checkup_date, '%Y-%m-%d')
            dai = (checkup_dt - insemination_dt).days
            return dai if dai >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _check_repro1(self, events: List[Dict[str, Any]]) -> bool:
        """
        繁殖検査１の条件をチェック
        - フレッシュチェックイベント後
        - AIイベントもETイベントも一度も存在しない個体
        """
        fchk_events = self._get_fresh_check_event_numbers()
        ai_events = self._get_ai_event_numbers()
        et_events = self._get_et_event_numbers()
        all_insemination_events = ai_events + et_events
        
        # フレッシュチェックイベントがあるか
        has_fchk = any(e.get('event_number') in fchk_events for e in events)
        if not has_fchk:
            return False
        
        # AI/ETイベントが存在しない
        has_insemination = any(e.get('event_number') in all_insemination_events for e in events)
        return not has_insemination
    
    def _check_repro2(self, events: List[Dict[str, Any]], cow: Dict[str, Any]) -> bool:
        """
        繁殖検査２の条件をチェック
        - AIまたはETイベント履歴あり
        - その後に妊娠鑑定マイナスまたはPAGマイナスイベントあり
        - 現在の繁殖コードが「Open」
        """
        ai_events = self._get_ai_event_numbers()
        et_events = self._get_et_event_numbers()
        all_insemination_events = ai_events + et_events
        pdn_events = self._get_pregnancy_negative_event_numbers()
        
        # AI/ETイベントがあるか
        has_insemination = any(e.get('event_number') in all_insemination_events for e in events)
        if not has_insemination:
            return False
        
        # 妊娠鑑定マイナスイベントがあるか（AI/ETの後）
        insemination_dates = []
        for event in events:
            if event.get('event_number') in all_insemination_events:
                event_date = event.get('event_date')
                if event_date:
                    insemination_dates.append(event_date)
        
        if not insemination_dates:
            return False
        
        last_insemination_date = max(insemination_dates)
        
        has_pdn_after = False
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if (event_number in pdn_events and event_date and 
                event_date > last_insemination_date):
                has_pdn_after = True
                break
        
        if not has_pdn_after:
            return False
        
        # 繁殖コードがOpen（4）
        rc = cow.get('rc')
        return rc == RuleEngine.RC_OPEN
    
    def _check_pregnancy_check(self, events: List[Dict[str, Any]], 
                               cow: Dict[str, Any], checkup_date: str, threshold: Optional[int]) -> bool:
        """
        妊娠鑑定の条件をチェック
        - 繁殖コードがPregnantでない（初回の妊娠鑑定）
        - 最終授精日DAIからN日以上
        """
        if threshold is None:
            return False
        
        # 繁殖コードがPregnantの場合は除外（再妊娠鑑定の対象）
        rc = cow.get('rc')
        if rc == RuleEngine.RC_PREGNANT:
            return False
        
        dai = self._calculate_dai(events, checkup_date)
        if dai is None:
            return False
        return dai >= threshold
    
    def _check_re_pregnancy_check(self, events: List[Dict[str, Any]], 
                                  cow: Dict[str, Any], checkup_date: str, 
                                  threshold: Optional[int]) -> bool:
        """
        再妊娠鑑定の条件をチェック
        - 繁殖コードが「Pregnant」
        - 最終授精日DAIからN日以上
        - 妊娠イベントが一回だけ（二回以上妊娠イベントがある個体は除外）
        """
        if threshold is None:
            return False
        rc = cow.get('rc')
        if rc != RuleEngine.RC_PREGNANT:
            return False
        
        # 妊娠プラスイベント（PDP, PDP2, PAGP）の数をカウント
        pdp_events = self._get_pregnancy_positive_event_numbers()
        pregnancy_positive_count = sum(1 for e in events 
                                       if e.get('event_number') in pdp_events)
        
        # 妊娠イベントが一回だけでなければ除外
        if pregnancy_positive_count != 1:
            return False
        
        dai = self._calculate_dai(events, checkup_date)
        if dai is None:
            return False
        return dai >= threshold
    
    def _check_optional_pregnancy_check(self, events: List[Dict[str, Any]], 
                                       cow: Dict[str, Any], checkup_date: str, 
                                       threshold: Optional[int]) -> bool:
        """
        任意妊娠鑑定の条件をチェック
        - 繁殖コードが「Pregnant」
        - 最終授精日DAIからN日以上
        - 妊娠イベントの回数制限なし（2回以上でも対象）
        """
        if threshold is None:
            return False
        rc = cow.get('rc')
        if rc != RuleEngine.RC_PREGNANT:
            return False
        
        dai = self._calculate_dai(events, checkup_date)
        if dai is None:
            return False
        return dai >= threshold
    
    def _check_due_over(self, cow: Dict[str, Any], events: List[Dict[str, Any]], 
                        checkup_date: str, threshold: Optional[int]) -> bool:
        """
        分娩予定超過の条件をチェック
        - 分娩予定日 + N日を超過
        """
        if threshold is None:
            return False
        # 分娩予定日を計算（受胎日 + 280日）
        last_ai_date = self._get_last_ai_date(events, checkup_date)
        if not last_ai_date:
            return False
        
        insemination_date = self._get_insemination_date_for_dai(events, last_ai_date)
        if not insemination_date:
            return False
        
        try:
            insemination_dt = datetime.strptime(insemination_date, '%Y-%m-%d')
            due_dt = insemination_dt + timedelta(days=280)
            checkup_dt = datetime.strptime(checkup_date, '%Y-%m-%d')
            
            # 分娩予定日 + threshold を超過しているか
            due_over_dt = due_dt + timedelta(days=threshold)
            return checkup_dt > due_over_dt
        except (ValueError, TypeError):
            return False
    
    def _check_check(self, events: List[Dict[str, Any]]) -> bool:
        """
        チェックの条件をチェック
        - 妊娠プラスイベント後
        - 妊娠マイナスイベントが存在しない
        - その後にAIまたはETイベントが記録されている個体
        """
        pdp_events = self._get_pregnancy_positive_event_numbers()
        pdn_events = self._get_pregnancy_negative_event_numbers()
        ai_events = self._get_ai_event_numbers()
        et_events = self._get_et_event_numbers()
        all_insemination_events = ai_events + et_events
        
        # 妊娠プラスイベントを探す
        preg_plus_dates = []
        for event in events:
            if event.get('event_number') in pdp_events:
                event_date = event.get('event_date')
                if event_date:
                    preg_plus_dates.append(event_date)
        
        if not preg_plus_dates:
            return False
        
        last_preg_plus_date = max(preg_plus_dates)
        
        # 妊娠マイナスイベントが存在しない（妊娠プラスより後）
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if (event_number in pdn_events and event_date and 
                event_date > last_preg_plus_date):
                return False
        
        # その後にAI/ETイベントが記録されている
        for event in events:
            event_number = event.get('event_number')
            event_date = event.get('event_date')
            if (event_number in all_insemination_events and event_date and 
                event_date > last_preg_plus_date):
                return True
        
        return False
    
    def _check_heifer_repro1(self, cow: Dict[str, Any], events: List[Dict[str, Any]], 
                              checkup_date: str, threshold: Optional[float]) -> bool:
        """
        育成繁殖１の条件をチェック
        - 月齢 >= N（月齢）
        - AI・ETイベント履歴なし
        """
        if threshold is None:
            return False
        age_months = self._calculate_age_months(cow, checkup_date)
        if age_months is None:
            return False
        
        if age_months < threshold:
            return False
        
        # AI/ETイベント履歴なし
        ai_events = self._get_ai_event_numbers()
        et_events = self._get_et_event_numbers()
        all_insemination_events = ai_events + et_events
        
        has_insemination = any(e.get('event_number') in all_insemination_events for e in events)
        return not has_insemination
    
    def _check_heifer_repro2(self, events: List[Dict[str, Any]], 
                              cow: Dict[str, Any]) -> bool:
        """
        育成繁殖２の条件をチェック
        - AIまたはETイベント履歴あり
        - 妊娠鑑定マイナスイベントあり
        - 繁殖コードが「Open」
        """
        # 繁殖検査２と同じロジック
        return self._check_repro2(events, cow)

