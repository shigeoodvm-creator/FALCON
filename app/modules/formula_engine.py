"""
FALCON2 - FormulaEngine
表示用・分析用の計算項目をオンデマンドで計算
設計書 第10章参照
"""

from typing import Dict, Any, Optional
from datetime import datetime, timedelta
from pathlib import Path
import json
import math
import logging

from db.db_handler import DBHandler
from modules.app_settings_manager import get_app_settings_manager
from modules.sire_list_opts import (
    sire_opts_to_type,
    SIRE_TYPE_F1,
    SIRE_TYPE_HOLSTEIN_FEMALE,
    SIRE_TYPE_HOLSTEIN_REGULAR,
    SIRE_TYPE_UNKNOWN_OTHER,
    SIRE_TYPE_WAGYU,
)


class FormulaEngine:
    """計算項目をオンデマンドで計算するエンジン"""
    
    def __init__(self, db_handler: DBHandler, item_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            item_dictionary_path: item_dictionary.json のパス（Noneの場合はデフォルトパスを使用）
        """
        self.db = db_handler
        self.item_dict_path = item_dictionary_path
        self.item_dictionary: Dict[str, Any] = {}
        self._load_item_dictionary()

    def reload_item_dictionary(self):
        """外部更新後に辞書を再読込する"""
        self._load_item_dictionary()
    
    def _load_item_dictionary(self):
        """item_dictionary.json を読み込む（本体 config_default を参照）"""
        if self.item_dict_path is None:
            app_root = Path(__file__).parent.parent.parent
            default_path = app_root / "config_default" / "item_dictionary.json"
            if default_path.exists():
                self.item_dict_path = default_path

        if self.item_dict_path and self.item_dict_path.exists():
            try:
                with open(self.item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
            except Exception as e:
                print(f"item_dictionary.json 読み込みエラー: {e}")
                self.item_dictionary = {}
        else:
            self.item_dictionary = {}
    
    def calculate(self, cow_auto_id: int, item_key: Optional[str] = None) -> Dict[str, Any]:
        """
        牛の計算項目を計算
        
        【重要】副作用なし。DB更新は行わない。
        
        Args:
            cow_auto_id: 牛の auto_id
            item_key: 計算対象の項目キー（指定時はこの項目のみ計算、Noneの場合は全項目計算）
        
        Returns:
            計算項目の辞書
        """
        # 常に最新の辞書を反映
        self._load_item_dictionary()

        # cow データを取得
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return {}
        
        # イベント履歴を取得
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        result = {}
        
        # item_key が指定されている場合、その項目のみを計算
        if item_key:
            item_def = self.item_dictionary.get(item_key, {})
            if not item_def:
                # item_dictionaryに存在しない場合、空の結果を返す
                return {}
            
            origin = item_def.get("origin") or item_def.get("type", "")
            
            # 1. source 定義による項目（event.field 形式）
            if origin == "source":
                source_items = self._calculate_source_items(events, target_item=item_key)
                if item_key in source_items:
                    result[item_key] = source_items[item_key]
            
            # 2. formula 定義による calc 項目
            elif origin == "calc":
                # 一時的にresultを作成（依存関係がある場合に備える）
                temp_result = {}
                formula_items = self._calculate_formula_items(cow, events, temp_result, cow_auto_id, target_item=item_key)
                if item_key in formula_items:
                    result[item_key] = formula_items[item_key]
            
            # 3. core項目（cowテーブルから直接取得）
            elif origin == "core":
                # formulaフィールドに定義された式を実行（例：cow.get('rc')）
                formula = item_def.get("formula")
                if formula:
                    try:
                        # cowオブジェクトをローカルスコープで使用可能にする
                        core_value = eval(formula, {"cow": cow, "__builtins__": {}})
                        if core_value is not None:
                            # データ型に応じて変換
                            data_type = item_def.get("data_type", "str")
                            try:
                                if data_type == "int":
                                    result[item_key] = int(core_value)
                                elif data_type == "float":
                                    result[item_key] = float(core_value)
                                else:
                                    result[item_key] = core_value
                            except (ValueError, TypeError):
                                result[item_key] = core_value
                    except Exception as e:
                        logging.debug(f"core項目の取得エラー: item_key={item_key}, formula={formula}, error={e}")
            
            # 4. custom項目（item_valueテーブルから取得）
            elif origin == "custom":
                custom_value = self.db.get_item_value(cow_auto_id, item_key)
                if custom_value is not None:
                    # データ型に応じて変換
                    data_type = item_def.get("data_type", "str")
                    try:
                        if data_type == "int":
                            result[item_key] = int(custom_value)
                        elif data_type == "float":
                            result[item_key] = float(custom_value)
                        else:
                            result[item_key] = custom_value
                    except (ValueError, TypeError):
                        result[item_key] = custom_value
            
            # 5. editableな計算項目の場合、item_valueテーブルから値を取得して上書き
            if item_key in result:
                if item_def.get("editable", False):
                    custom_value = self.db.get_item_value(cow_auto_id, item_key)
                    if custom_value is not None:
                        data_type = item_def.get("data_type", "str")
                        try:
                            if data_type == "int":
                                result[item_key] = int(custom_value)
                            elif data_type == "float":
                                result[item_key] = float(custom_value)
                            else:
                                result[item_key] = custom_value
                        except (ValueError, TypeError):
                            result[item_key] = custom_value
            
            return result
        
        # item_key が指定されていない場合、全項目を計算（従来通り）
        # 1. source 定義による項目（event.field 形式）
        source_items = self._calculate_source_items(events)
        result.update(source_items)
        
        # 2. formula 定義による calc 項目（DIM、DAI、DUE も含む）
        formula_items = self._calculate_formula_items(cow, events, result, cow_auto_id)
        result.update(formula_items)
        
        # 3. editableな計算項目の場合、item_valueテーブルから値を取得して上書き
        # （ユーザーが手動で編集した値があれば、それを優先）
        for item_key_iter in result.keys():
            item_def = self.item_dictionary.get(item_key_iter, {})
            if item_def.get("editable", False):
                # item_valueテーブルから値を取得
                custom_value = self.db.get_item_value(cow_auto_id, item_key_iter)
                if custom_value is not None:
                    # データ型に応じて変換
                    data_type = item_def.get("data_type", "str")
                    try:
                        if data_type == "int":
                            result[item_key_iter] = int(custom_value)
                        elif data_type == "float":
                            result[item_key_iter] = float(custom_value)
                        else:
                            result[item_key_iter] = custom_value
                    except (ValueError, TypeError):
                        # 変換エラーの場合は文字列のまま
                        result[item_key_iter] = custom_value
        
        # 4. custom項目（origin: "custom"でformula/sourceがない項目）をitem_valueテーブルから取得
        for item_key_iter, item_def in self.item_dictionary.items():
            origin = item_def.get("origin") or item_def.get("type", "")
            if origin == "custom" and item_key_iter not in result:
                # custom項目でまだresultに含まれていない場合
                custom_value = self.db.get_item_value(cow_auto_id, item_key_iter)
                if custom_value is not None:
                    # データ型に応じて変換
                    data_type = item_def.get("data_type", "str")
                    try:
                        if data_type == "int":
                            result[item_key_iter] = int(custom_value)
                        elif data_type == "float":
                            result[item_key_iter] = float(custom_value)
                        else:
                            result[item_key_iter] = custom_value
                    except (ValueError, TypeError):
                        # 変換エラーの場合は文字列のまま
                        result[item_key_iter] = custom_value
        
        return result
    
    def _calculate_age(self, bthd: Optional[str]) -> Optional[float]:
        """
        月齢（AGE）を計算
        
        生年月日から今日までの月齢（月単位）
        
        Args:
            bthd: 生年月日（YYYY-MM-DD形式）
        
        Returns:
            月齢（浮動小数点数）、bthdがNoneの場合はNone
        """
        if not bthd:
            return None
        
        try:
            birth_date = datetime.strptime(bthd, '%Y-%m-%d')
            today = datetime.now()
            days_diff = (today - birth_date).days
            if days_diff < 0:
                return None
            # 月齢に変換（1ヶ月 = 30.44日）
            age_months = days_diff / 30.44
            return round(age_months, 1)
        except (ValueError, TypeError):
            return None
    
    def _calculate_age_at_calving(self, bthd: Optional[str], clvd: Optional[str]) -> Optional[float]:
        """
        分娩時の月齢（AGE at Calving）を計算
        
        今産次の分娩日時点での月齢（月単位）
        
        Args:
            bthd: 生年月日（YYYY-MM-DD形式）
            clvd: 最終分娩日（今産次の分娩日、YYYY-MM-DD形式）
        
        Returns:
            分娩時の月齢（浮動小数点数）、計算できない場合はNone
        """
        if not bthd or not clvd:
            return None
        
        try:
            birth_date = datetime.strptime(bthd, '%Y-%m-%d')
            calving_date = datetime.strptime(clvd, '%Y-%m-%d')
            days_diff = (calving_date - birth_date).days
            if days_diff < 0:
                return None
            # 月齢に変換（1ヶ月 = 30.44日）
            age_months = days_diff / 30.44
            return round(age_months, 1)
        except (ValueError, TypeError):
            return None
    
    def _calculate_age_at_first_calving(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        初産分娩時の月齢（AGE at First Calving）を計算
        
        最初の分娩日時点での月齢（月単位）
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初産分娩時の月齢（浮動小数点数）、計算できない場合はNone
        """
        # 生年月日を取得
        bthd = cow.get('bthd')
        if not bthd:
            return None
        
        try:
            birth_date = datetime.strptime(bthd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩イベント（202）を取得
        calving_events = [
            e for e in events
            if e.get('event_number') == 202  # CALV
            and e.get('event_date')
        ]
        
        if not calving_events:
            return None
        
        # 日付順にソート（昇順：古い順）
        sorted_events = sorted(
            calving_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 最初の分娩イベントを取得（baseline_calvingフラグがないもの）
        for event in sorted_events:
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # baseline_calvingフラグがある場合はスキップ
            if json_data.get('baseline_calving', False):
                continue
            
            first_calving_date = event.get('event_date')
            if not first_calving_date:
                continue
            
            try:
                calving_date = datetime.strptime(first_calving_date, '%Y-%m-%d')
                days_diff = (calving_date - birth_date).days
                if days_diff < 0:
                    return None
                # 月齢に変換（1ヶ月 = 30.44日）
                age_months = days_diff / 30.44
                return round(age_months, 1)
            except (ValueError, TypeError):
                continue
        
            return None
    
    def _get_calf_info_by_lact(self, events: list, cow: Dict[str, Any], target_lact: int, cow_auto_id: int = None) -> Optional[str]:
        """
        指定産次の分娩イベントでの子牛情報を取得
        
        分娩イベントで乳用種メスの場合、導入イベント（600）が発生する可能性がある。
        導入イベントのdam_jpn10が母牛のJPN10と一致する場合、その導入イベント後に作成された
        新しい牛のJPN10を取得。なければ品種・性別・死産有無を文字列化
        
        Args:
            events: イベントリスト（現在の牛のイベント）
            cow: 牛データ（母牛）
            target_lact: 対象産次（1=今産次、2=前産次、3=前々産次）
            cow_auto_id: 牛のauto_id（母牛のauto_id）
        
        Returns:
            子牛情報（文字列）、存在しない場合はNone
        """
        # 分娩イベント（202）を取得
        calving_events = [
            e for e in events
            if e.get('event_number') == 202  # CALV
            and e.get('event_date')
        ]
        
        if not calving_events:
            return None
        
        # 日付順にソート（降順：新しい順）
        sorted_calving = sorted(
            calving_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # baseline_calvingフラグがない分娩イベントのみをカウント
        valid_calvings = []
        for event in sorted_calving:
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # baseline_calvingフラグがある場合はスキップ
            if not json_data.get('baseline_calving', False):
                valid_calvings.append(event)
        
        if len(valid_calvings) < target_lact:
            return None
        
        target_calving = valid_calvings[target_lact - 1]
        calving_date = target_calving.get('event_date')
        
        if not calving_date:
            return None
        
        # 分娩イベントのjson_dataを取得
        calving_json = target_calving.get('json_data') or {}
        if isinstance(calving_json, str):
            try:
                calving_json = json.loads(calving_json)
            except:
                calving_json = {}
        
        # 母牛のJPN10を取得
        dam_jpn10 = cow.get('jpn10')
        
        # 分娩日以降の導入イベント（600）を全牛から検索
        # 導入イベントのdam_jpn10が母牛のJPN10と一致するものを探す
        if cow_auto_id and dam_jpn10:
            try:
                # 全牛を取得
                all_cows = self.db.get_all_cows()
                
                # 全牛のイベントから導入イベントを検索
                for calf_cow in all_cows:
                    calf_events = self.db.get_events_by_cow(calf_cow['auto_id'], include_deleted=False)
                    import_events = [
                        e for e in calf_events
                        if e.get('event_number') == 600  # IN（導入）
                        and e.get('event_date')
                        and e.get('event_date', '') >= calving_date
                    ]
                    
                    for import_event in import_events:
                        import_json = import_event.get('json_data') or {}
                        if isinstance(import_json, str):
                            try:
                                import_json = json.loads(import_json)
                            except:
                                import_json = {}
                        
                        # 導入イベントのdam_jpn10が母牛のJPN10と一致するか確認
                        import_dam_jpn10 = import_json.get('dam_jpn10')
                        if import_dam_jpn10 and str(import_dam_jpn10).strip() == str(dam_jpn10).strip():
                            # この導入イベントの日付以降に作成された牛を検索
                            import_date = import_event.get('event_date')
                            
                            # 導入イベントの日付と一致するentrまたはbthdを持つ牛を探す
                            # または、導入イベントの日付以降に作成された牛を探す
                            for candidate_cow in all_cows:
                                candidate_entr = candidate_cow.get('entr')
                                candidate_bthd = candidate_cow.get('bthd')
                                
                                # 導入イベントの日付と一致するentrまたはbthdを持つ牛
                                if (candidate_entr and candidate_entr == import_date) or \
                                   (candidate_bthd and candidate_bthd == import_date):
                                    # この牛のJPN10を返す
                                    calf_jpn10 = candidate_cow.get('jpn10')
                                    if calf_jpn10:
                                        return str(calf_jpn10).strip()
                            
                            # 直接的な一致が見つからない場合、導入イベントの日付以降に作成された
                            # 最初の牛を探す（簡易的な方法）
                            for candidate_cow in all_cows:
                                candidate_entr = candidate_cow.get('entr')
                                if candidate_entr and candidate_entr >= import_date:
                                    # この牛が導入イベントの牛である可能性が高い
                                    calf_jpn10 = candidate_cow.get('jpn10')
                                    if calf_jpn10:
                                        return str(calf_jpn10).strip()
            except Exception as e:
                logging.warning(f"子牛情報取得エラー: {e}")
        
        # 導入イベントがない、またはJPN10が取得できない場合
        # 分娩イベントのjson_dataから品種・性別・死産有無を取得して文字列化
        calf_sex = calving_json.get('calf_sex', '')
        abnormal_flag = calving_json.get('abnormal_flag', '')
        calf_breed = calving_json.get('calf_breed', '')  # 品種（存在する場合）
        
        # 文字列を構築
        parts = []
        if calf_breed:
            parts.append(calf_breed)
        if calf_sex:
            parts.append(calf_sex)
        if abnormal_flag:
            parts.append(abnormal_flag)
        
        if parts:
            return '/'.join(parts)
        
        return None
    
    def _extract_month_from_date(self, date_str: Optional[str]) -> Optional[int]:
        """
        日付文字列から月を抽出
        
        Args:
            date_str: 日付文字列（YYYY-MM-DD形式）
        
        Returns:
            月（1-12）、日付が無効な場合はNone
        """
        if not date_str:
            return None
        
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return date_obj.month
        except (ValueError, TypeError):
            return None
    
    def _get_calving_month(self, cow: Dict[str, Any]) -> Optional[int]:
        """
        今産次の分娩月を取得
        
        Args:
            cow: 牛データ
        
        Returns:
            分娩月（1-12）、存在しない場合はNone
        """
        clvd = cow.get('clvd')
        return self._extract_month_from_date(clvd)
    
    def _calculate_calving_year_month(self, clvd: Optional[str]) -> Optional[str]:
        """
        分娩年月を取得（YYYY-MM形式）
        
        Args:
            clvd: 最終分娩日（YYYY-MM-DD形式）
        
        Returns:
            分娩年月（YYYY-MM形式）、存在しない場合はNone
        """
        if not clvd:
            return None
        
        try:
            date_obj = datetime.strptime(clvd, '%Y-%m-%d')
            return date_obj.strftime('%Y-%m')
        except (ValueError, TypeError):
            return None
    
    def _calculate_birth_year(self, bthd: Optional[str]) -> Optional[int]:
        """
        生まれた西暦年を取得
        
        Args:
            bthd: 生年月日（YYYY-MM-DD形式）
        
        Returns:
            西暦年（整数）、存在しない場合はNone
        """
        if not bthd:
            return None
        try:
            date_obj = datetime.strptime(bthd, '%Y-%m-%d')
            return date_obj.year
        except (ValueError, TypeError):
            return None
    
    def _calculate_birth_year_month(self, bthd: Optional[str]) -> Optional[str]:
        """
        生まれた年月を取得（YYYY-MM形式）
        
        Args:
            bthd: 生年月日（YYYY-MM-DD形式）
        
        Returns:
            年月（YYYY-MM形式）、存在しない場合はNone
        """
        if not bthd:
            return None
        try:
            date_obj = datetime.strptime(bthd, '%Y-%m-%d')
            return date_obj.strftime('%Y-%m')
        except (ValueError, TypeError):
            return None
    
    def _calculate_first_ai_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        その産次の「授精回数1回目」の授精日を取得。
        授精回数1回目のイベントが存在しない場合はNone。
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回授精日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        current_lact = cow.get('lact')
        if current_lact is None:
            return None
        
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩日より後のAI/ETで、授精回数1回目のもののみ
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]
            and e.get('event_date')
            and e.get('event_date', '') > clvd
            and self._event_insemination_count_is_one(e)
        ]
        if not ai_et_events:
            return None
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        return sorted_events[0].get('event_date')
    
    def _calculate_first_milk_test_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        その産次の初回乳検DIMを計算
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回乳検DIM（分娩後日数）、存在しない場合はNone
        """
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        first_test_date = self._calculate_first_milk_test_date(events, cow)
        if not first_test_date:
            return None
        
        return self._calculate_dim_from_date(clvd, first_test_date)
    
    def _calculate_first_milk_test_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        その産次の初回乳検日を取得
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回乳検日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        current_lact = cow.get('lact')
        if current_lact is None:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 現在の産次に対応する乳検イベントを探す
        # event_lactフィールドがある場合はそれを使用
        lact_events = []
        for event in milk_test_events:
            event_lact = event.get('event_lact')
            event_date = event.get('event_date')
            
            if event_lact is not None and event_lact == current_lact:
                try:
                    test_date = datetime.strptime(event_date, '%Y-%m-%d')
                    if test_date >= clvd_date:
                        lact_events.append(event)
                except (ValueError, TypeError):
                    continue
        
        # event_lactがない場合、clvd以降の乳検イベントを探す
        if not lact_events:
            for event in milk_test_events:
                event_date = event.get('event_date')
                if event_date:
                    try:
                        test_date = datetime.strptime(event_date, '%Y-%m-%d')
                        if test_date >= clvd_date:
                            lact_events.append(event)
                    except (ValueError, TypeError):
                        continue
        
        if not lact_events:
            return None
        
        # 日付順にソート（古い順）
        sorted_events = sorted(
            lact_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        return sorted_events[0].get('event_date')
    
    def _calculate_second_milk_test_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        その産次の二回乳検DIMを計算
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            二回乳検DIM（分娩後日数）、存在しない場合はNone
        """
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        second_test_date = self._calculate_second_milk_test_date(events, cow)
        if not second_test_date:
            return None
        
        return self._calculate_dim_from_date(clvd, second_test_date)
    
    def _calculate_second_milk_test_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        その産次の二回乳検日を取得
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            二回乳検日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        current_lact = cow.get('lact')
        if current_lact is None:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 現在の産次に対応する乳検イベントを探す
        lact_events = []
        for event in milk_test_events:
            event_lact = event.get('event_lact')
            event_date = event.get('event_date')
            
            if event_lact is not None and event_lact == current_lact:
                try:
                    test_date = datetime.strptime(event_date, '%Y-%m-%d')
                    if test_date >= clvd_date:
                        lact_events.append(event)
                except (ValueError, TypeError):
                    continue
        
        # event_lactがない場合、clvd以降の乳検イベントを探す
        if not lact_events:
            for event in milk_test_events:
                event_date = event.get('event_date')
                if event_date:
                    try:
                        test_date = datetime.strptime(event_date, '%Y-%m-%d')
                        if test_date >= clvd_date:
                            lact_events.append(event)
                    except (ValueError, TypeError):
                        continue
        
        if len(lact_events) < 2:
            return None
        
        # 日付順にソート（古い順）
        sorted_events = sorted(
            lact_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 2番目のイベントを返す
        return sorted_events[1].get('event_date')
    
    def _calculate_dim_from_date(self, clvd: str, target_date: str) -> Optional[int]:
        """
        分娩日から指定日までのDIMを計算
        
        Args:
            clvd: 最終分娩日（YYYY-MM-DD形式）
            target_date: 対象日（YYYY-MM-DD形式）
        
        Returns:
            DIM（分娩後日数）、計算できない場合はNone
        """
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
            target_datetime = datetime.strptime(target_date, '%Y-%m-%d')
            dim = (target_datetime - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_nth_milk_test_date(self, events: list, cow: Dict[str, Any], n: int) -> Optional[str]:
        """
        今産次のN回目の乳検イベントの日付を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        n=1が1回目（最も古い）、n=2が2回目、n=3が3回目
        
        Args:
            events: イベントリスト
            cow: 牛データ
            n: 何回目のイベントか（1-indexed、1が1回目）
        
        Returns:
            乳検日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if len(milk_test_events) < n:
            return None
        
        # 日付順にソート（昇順：古い順、1回目が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # N回目の乳検イベントの日付を取得
        nth_event = sorted_events[n - 1]
        return nth_event.get('event_date')
    
    def _get_previous_calving_month(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        前産次の分娩月を取得
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            前産分娩月（1-12）、存在しない場合はNone
        """
        previous_clvd = self._get_previous_calving_date(events, cow, 2)
        return self._extract_month_from_date(previous_clvd)
    
    def _get_previous_calving_date(self, events: list, cow: Dict[str, Any], target_lact: int) -> Optional[str]:
        """
        指定産次の分娩日を取得
        
        Args:
            events: イベントリスト（現在の牛のイベント）
            cow: 牛データ
            target_lact: 対象産次（1=今産次、2=前産次、3=前々産次）
        
        Returns:
            分娩日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩イベント（202）を取得
        calving_events = [
            e for e in events
            if e.get('event_number') == 202  # CALV
            and e.get('event_date')
        ]
        
        if not calving_events:
            return None
        
        # 日付順にソート（降順：新しい順）
        sorted_calving = sorted(
            calving_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # baseline_calvingフラグがない分娩イベントのみをカウント
        valid_calvings = []
        for event in sorted_calving:
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # baseline_calvingフラグがある場合はスキップ
            if not json_data.get('baseline_calving', False):
                valid_calvings.append(event)
        
        if len(valid_calvings) < target_lact:
            return None
        
        target_calving = valid_calvings[target_lact - 1]
        calving_date = target_calving.get('event_date')
        
        return calving_date if calving_date else None
    
    def _calculate_previous_lact_days_open(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        前産次の空胎日数を計算
        
        前産次の分娩日から前産次の最初の授精日（AI/ET）までの日数
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            前産次の空胎日数（整数）、計算できない場合はNone
        """
        # 今産次の分娩日を取得
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 前産次の分娩日を取得
        previous_clvd = self._get_previous_calving_date(events, cow, 2)
        if not previous_clvd:
            return None
        
        try:
            previous_clvd_date = datetime.strptime(previous_clvd, '%Y-%m-%d')
            current_clvd_date = datetime.strptime(current_clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 前産次の分娩日以降、今産次の分娩日より前の授精イベント（200: AI, 201: ET）を取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and previous_clvd <= e.get('event_date', '') < current_clvd
        ]
        
        if not ai_et_events:
            return None
        
        # 日付順にソート（昇順：古い順）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 前産次の最初の授精イベントの日付を取得
        first_ai_et = sorted_events[0]
        first_ai_et_date = first_ai_et.get('event_date')
        
        if not first_ai_et_date:
            return None
        
        try:
            first_ai_et_datetime = datetime.strptime(first_ai_et_date, '%Y-%m-%d')
            # 前産次の分娩日から最初の授精日までの日数
            days_open = (first_ai_et_datetime - previous_clvd_date).days
            return days_open if days_open >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_previous_lact_gestation_days(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        前産次の在胎日数を計算
        
        前産次の受胎日から前産次の分娩日までの日数
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            前産次の在胎日数（整数）、計算できない場合はNone
        """
        # 今産次の分娩日を取得（これが前産次の分娩日）
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 前々産次の分娩日を取得（前産次の開始点）
        previous_previous_clvd = self._get_previous_calving_date(events, cow, 3)
        
        try:
            current_clvd_date = datetime.strptime(current_clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 前産次のAI/ETイベントを取得（前々産次の分娩日より後、今産次の分娩日より前）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
        ]
        
        # 前産次の範囲でフィルタ
        if previous_previous_clvd:
            ai_et_events = [
                e for e in ai_et_events
                if previous_previous_clvd < e.get('event_date', '') < current_clvd
            ]
        else:
            ai_et_events = [
                e for e in ai_et_events
                if e.get('event_date', '') < current_clvd
            ]
        
        if not ai_et_events:
            return None
        
        # 前産次の妊娠鑑定プラスイベントを取得
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
            and e.get('event_date')
        ]
        
        # 前産次の範囲でフィルタ
        if previous_previous_clvd:
            preg_events = [
                e for e in preg_events
                if previous_previous_clvd < e.get('event_date', '') < current_clvd
            ]
        else:
            preg_events = [
                e for e in preg_events
                if e.get('event_date', '') < current_clvd
            ]
        
        if not preg_events:
            return None
        
        # 時系列順にソート（古い順）
        sorted_preg_events = sorted(
            preg_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 最初の妊娠イベントで受胎日を確定
        first_preg = sorted_preg_events[0]
        first_preg_event_number = first_preg.get('event_number')
        first_preg_date = first_preg.get('event_date')
        
        # 最初の妊娠イベントがPDP2で受胎日が指定されている場合
        conception_date = None
        if first_preg_event_number == 304:  # PDP2
            json_data = first_preg.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            ai_event_id = json_data.get('ai_event_id')
            if ai_event_id:
                # 指定されたAI/ETイベントを検索
                for event in ai_et_events:
                    if event.get('id') == ai_event_id:
                        ai_et_date = event.get('event_date')
                        event_number = event.get('event_number')
                        try:
                            ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
                            # ETの場合は7日前、AIの場合は当日が受胎日
                            if event_number == 201:  # ET
                                conception_dt = ai_et_dt - timedelta(days=7)
                            else:  # AI
                                conception_dt = ai_et_dt
                            conception_date = conception_dt.strftime('%Y-%m-%d')
                            break
                        except (ValueError, TypeError):
                            pass
        
        # 最初の妊娠イベントがPDP/PAGPの場合、直近のAI/ETイベントを使用（妊娠鑑定より前）
        if not conception_date:
            ai_et_before_first_preg = [
                e for e in ai_et_events
                if e.get('event_date', '') <= first_preg_date
            ]
            
            if not ai_et_before_first_preg:
                return None
            
            latest_ai_et = sorted(
                ai_et_before_first_preg,
                key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                reverse=True
            )[0]
            
            ai_et_date = latest_ai_et.get('event_date')
            event_number = latest_ai_et.get('event_number')
            
            try:
                ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
                # ETの場合は7日前、AIの場合は当日が受胎日
                if event_number == 201:  # ET
                    conception_dt = ai_et_dt - timedelta(days=7)
                else:  # AI
                    conception_dt = ai_et_dt
                conception_date = conception_dt.strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                return None
            
            # 二回目以降のPDP2イベントで受胎日が指定されている場合は優先
            if len(sorted_preg_events) > 1:
                for preg_event in sorted_preg_events[1:]:  # 二回目以降
                    if preg_event.get('event_number') == 304:  # PDP2
                        json_data = preg_event.get('json_data') or {}
                        if isinstance(json_data, str):
                            try:
                                json_data = json.loads(json_data)
                            except:
                                json_data = {}
                        
                        ai_event_id = json_data.get('ai_event_id')
                        if ai_event_id:
                            # 指定されたAI/ETイベントを検索
                            for event in ai_et_events:
                                if event.get('id') == ai_event_id:
                                    ai_et_date = event.get('event_date')
                                    event_number = event.get('event_number')
                                    try:
                                        ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
                                        # ETの場合は7日前、AIの場合は当日が受胎日
                                        if event_number == 201:  # ET
                                            conception_dt = ai_et_dt - timedelta(days=7)
                                        else:  # AI
                                            conception_dt = ai_et_dt
                                        conception_date = conception_dt.strftime('%Y-%m-%d')
                                        break
                                    except (ValueError, TypeError):
                                        pass
        
        if not conception_date:
            return None
        
        try:
            conception_dt = datetime.strptime(conception_date, '%Y-%m-%d')
            # 受胎日から分娩日までの日数
            gestation_days = (current_clvd_date - conception_dt).days
            return gestation_days if gestation_days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_calving_interval(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        分娩間隔（CINT）を計算
        
        前産次の分娩日と今産次の分娩日の差（単位：日）
        分娩イベントが2回以上存在する場合のみ計算
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            分娩間隔（日数）、計算できない場合はNone
        """
        # 分娩イベント（202）を取得（baseline_calvingフラグに関係なくすべて取得）
        calving_events = [
            e for e in events
            if e.get('event_number') == 202  # CALV
            and e.get('event_date')
        ]
        
        if len(calving_events) < 2:
            return None
        
        # 日付順にソート（降順：新しい順）
        sorted_calving = sorted(
            calving_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の分娩日（今産次）
        current_clvd = sorted_calving[0].get('event_date')
        if not current_clvd:
            return None
        
        # 前回の分娩日（前産次）
        previous_clvd = sorted_calving[1].get('event_date')
        if not previous_clvd:
            return None
        
        try:
            current_date = datetime.strptime(current_clvd, '%Y-%m-%d')
            previous_date = datetime.strptime(previous_clvd, '%Y-%m-%d')
            # 前産次の分娩日から今産次の分娩日までの日数
            interval = (current_date - previous_date).days
            return interval if interval >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_days_from_conception(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        受胎後日数（DCC）を計算
        
        受胎日から本日までの日数
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            受胎後日数（日数）、計算できない場合はNone
        """
        # 受胎日を取得
        conception_date = self._calculate_conception_date(events, cow)
        if not conception_date:
            return None
        
        try:
            conception_dt = datetime.strptime(conception_date, '%Y-%m-%d')
            today = datetime.now()
            # 受胎日から今日までの日数
            days = (today - conception_dt).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_dry_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        今産次の乾乳日を取得
        
        分娩日（clvd）以降の乾乳イベント（203）の日付を取得
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            乾乳日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乾乳イベント（203）を取得
        dry_events = [
            e for e in events
            if e.get('event_number') == 203  # DRY
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not dry_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            dry_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乾乳イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _get_stopr_date(self, events: list) -> Optional[str]:
        """
        繁殖停止イベントの日付を取得
        
        繁殖停止イベント（204）の日付を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            繁殖停止日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 繁殖停止イベント（204）を取得
        stopr_events = [
            e for e in events
            if e.get('event_number') == 204  # STOPR
            and e.get('event_date')
        ]
        
        if not stopr_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            stopr_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の繁殖停止イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _calculate_stopr_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        繁殖停止DIMを計算
        
        STOPRイベント時のDIM
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            DIM（日数）、計算できない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 繁殖停止イベント（204）を取得
        stopr_events = [
            e for e in events
            if e.get('event_number') == 204  # STOPR
            and e.get('event_date')
        ]
        
        if not stopr_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            stopr_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の繁殖停止イベントの日付を取得
        latest_event = sorted_events[0]
        stopr_date = latest_event.get('event_date')
        
        try:
            stopr_datetime = datetime.strptime(stopr_date, '%Y-%m-%d')
            # 分娩日から繁殖停止日までの日数
            dim = (stopr_datetime - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_milk_test_date_before_dry(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        乾乳前乳検日を取得
        
        乾乳イベントの直前の乳検イベントの日付
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            乾乳前乳検日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乾乳イベント（203）を取得
        dry_events = [
            e for e in events
            if e.get('event_number') == 203  # DRY
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not dry_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_dry_events = sorted(
            dry_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乾乳イベントの日付を取得
        latest_dry_event = sorted_dry_events[0]
        dry_date = latest_dry_event.get('event_date')
        
        # 乾乳日より前の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
            and e.get('event_date', '') < dry_date  # 乾乳日より前
        ]
        
        if not milk_test_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_milk_test_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 乾乳直前の乳検イベント（最新）の日付を取得
        last_milk_test_before_dry = sorted_milk_test_events[0]
        return last_milk_test_before_dry.get('event_date')
    
    def _get_linear_score_before_dry(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        乾乳前リニアスコアを取得
        
        乾乳イベントの直前の乳検イベントのリニアスコア（現産次）
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            リニアスコア（浮動小数点数）、存在しない場合はNone
        """
        # 乾乳前乳検日を取得
        milk_test_date = self._get_milk_test_date_before_dry(events, cow)
        if not milk_test_date:
            return None
        
        # その日付の乳検イベントからリニアスコアを取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date') == milk_test_date
        ]
        
        if not milk_test_events:
            return None
        
        # 最初のイベントから値を取得
        event = milk_test_events[0]
        json_data = event.get('json_data') or {}
        
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        ls_value = json_data.get('ls')
        if ls_value is not None:
            try:
                return float(ls_value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_previous_dry_linear_score(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        前産次の乾乳前リニアスコアを取得
        
        前産次の乾乳イベント直前の乳検イベントでのリニアスコア
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            リニアスコア（浮動小数点数）、存在しない場合はNone
        """
        # 今産次の分娩日を取得
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 前産次の分娩日を取得
        previous_clvd = self._get_previous_calving_date(events, cow, 2)
        if not previous_clvd:
            return None
        
        # 前産次の乾乳日を取得
        previous_dry_date = self._get_previous_dry_date(events, cow)
        if not previous_dry_date:
            return None
        
        # 前産次の分娩日以降、前産次の乾乳日より前の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= previous_clvd  # 前産次の分娩日以降
            and e.get('event_date', '') < previous_dry_date  # 前産次の乾乳日より前
        ]
        
        if not milk_test_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_milk_test_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 前産次の乾乳直前の乳検イベント（最新）からリニアスコアを取得
        last_milk_test_before_dry = sorted_milk_test_events[0]
        json_data = last_milk_test_before_dry.get('json_data') or {}
        
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        ls_value = json_data.get('ls')
        if ls_value is not None:
            try:
                return float(ls_value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_previous_dry_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        前産次の乾乳日を取得
        
        前産次の分娩日以降、今産次の分娩日より前の乾乳イベント（203）の日付を取得
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            前産次の乾乳日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 今産次の分娩日を取得
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 前産次の分娩日を取得
        previous_clvd = self._get_previous_calving_date(events, cow, 2)
        if not previous_clvd:
            return None
        
        # 前産次の分娩日以降、今産次の分娩日より前の乾乳イベント（203）を取得
        dry_events = [
            e for e in events
            if e.get('event_number') == 203  # DRY
            and e.get('event_date')
            and e.get('event_date', '') >= previous_clvd  # 前産次の分娩日以降
            and e.get('event_date', '') < current_clvd  # 今産次の分娩日より前
        ]
        
        if not dry_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            dry_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乾乳イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _get_last_reproduction_check_date(self, events: list) -> Optional[str]:
        """
        最終繁殖検診日を取得
        
        イベント300、301、302、303、304のうち最終日付
        
        Args:
            events: イベントリスト
        
        Returns:
            最終繁殖検診日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 繁殖検診イベント（300, 301, 302, 303, 304）を取得
        repro_check_events = [
            e for e in events
            if e.get('event_number') in [300, 301, 302, 303, 304]  # FCHK, REPRO, PDN, PDP, PDP2
            and e.get('event_date')
        ]
        
        if not repro_check_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            repro_check_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の繁殖検診イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _get_fresh_check_date(self, events: list) -> Optional[str]:
        """
        フレッシュチェック日を取得
        
        フレッシュチェックイベント（300）の最新日付
        
        Args:
            events: イベントリスト
        
        Returns:
            フレッシュチェック日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # フレッシュチェックイベント（300）を取得
        fresh_check_events = [
            e for e in events
            if e.get('event_number') == 300  # FCHK
            and e.get('event_date')
        ]
        
        if not fresh_check_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            fresh_check_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のフレッシュチェックイベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _get_last_reproduction_check_note(self, events: list) -> Optional[str]:
        """
        検診内容を取得
        
        最終検診日（イベント300～304、307）のNOTE内容
        - 妊娠鑑定イベント（303, 304, 307 = PDP, PDP2, PAGP）の場合：
          その時点での授精後日数を計算して「妊娠〇日」と表示
        - その他の検診イベントの場合：
          json_dataから動的に生成する（cow_card.pyの_display_eventsと同じロジック）
        
        Args:
            events: イベントリスト
        
        Returns:
            検診内容（NOTE文字列）、存在しない場合はNone
        """
        # 繁殖検診イベント（300, 301, 302, 303, 304, 307）を取得
        repro_check_events = [
            e for e in events
            if e.get('event_number') in [300, 301, 302, 303, 304, 307]  # FCHK, REPRO, PDN, PDP, PDP2, PAGP
            and e.get('event_date')
        ]
        
        if not repro_check_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            repro_check_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の繁殖検診イベント
        latest_event = sorted_events[0]
        event_number = latest_event.get('event_number')
        event_date = latest_event.get('event_date')
        json_data = latest_event.get('json_data') or {}
        
        # 妊娠鑑定イベント（303, 304, 307 = PDP, PDP2, PAGP）の場合
        if event_number in [303, 304, 307]:
            # その時点での授精後日数を計算
            dai = self._calculate_dai_at_event_date(events, event_date)
            if dai is not None:
                return f"妊娠{dai}日"
            # 授精後日数が計算できない場合は通常の処理にフォールバック
        
        # その他の検診イベント、または妊娠鑑定イベントで授精後日数が計算できない場合
        # 有効な値のみをリスト化（空文字、None、"-"は除外）
        valid_parts = []
        
        # 処置（新しいキー名と古いキー名の両方をサポート）
        treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
        if treatment and str(treatment).strip() and str(treatment).strip() != '-':
            valid_parts.append(str(treatment).strip())
        
        # 子宮所見（新しいキー名と古いキー名の両方をサポート）
        uterine = (json_data.get('uterine_findings') or 
                  json_data.get('uterus_findings') or 
                  json_data.get('uterus_finding') or 
                  json_data.get('uterus', ''))
        if uterine and str(uterine).strip() and str(uterine).strip() != '-':
            valid_parts.append(f"子宮{str(uterine).strip()}")
        
        # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
        left_ovary = (json_data.get('left_ovary_findings') or 
                     json_data.get('leftovary_findings') or 
                     json_data.get('leftovary_finding') or 
                     json_data.get('left_ovary', '') or
                     json_data.get('leftovary', ''))
        if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
            valid_parts.append(f"左{str(left_ovary).strip()}")
        
        # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
        right_ovary = (json_data.get('right_ovary_findings') or 
                      json_data.get('rightovary_findings') or 
                      json_data.get('rightovary_finding') or 
                      json_data.get('right_ovary', '') or
                      json_data.get('rightovary', ''))
        if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
            valid_parts.append(f"右{str(right_ovary).strip()}")
        
        # remark（新しいキー名と古いキー名の両方をサポート）
        # まず remark を取得、なければ other 系を確認
        remark = json_data.get('remark')
        if not remark:
            # remark が無い場合は other 系を確認
            remark = (json_data.get('other') or 
                     json_data.get('other_info') or 
                     json_data.get('other_findings', ''))
        if remark and str(remark).strip() and str(remark).strip() != '-':
            valid_parts.append(str(remark).strip())
        
        # 有効な項目のみを「  」で区切って返す
        if valid_parts:
            return "  ".join(valid_parts)
        
        # json_dataから生成できない場合は、既存のnoteフィールドを確認
        note = latest_event.get('note')
        return note.strip() if note else None
    
    def _get_fresh_check_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        フレッシュチェックイベントのDIMを取得
        
        フレッシュチェックイベント（300）のevent_dimを取得。
        DBのevent_dimがNULLの場合は、イベント日と分娩日からその場で計算する（フォールバック）。
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            フレッシュチェックイベントのDIM（整数）、存在しない場合はNone
        """
        # フレッシュチェックイベント（300）を取得
        fresh_check_events = [
            e for e in events
            if e.get('event_number') == 300  # FCHK
            and e.get('event_date')
        ]
        
        if not fresh_check_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            fresh_check_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のフレッシュチェックイベントのevent_dimを取得
        latest_event = sorted_events[0]
        event_dim = latest_event.get('event_dim')
        
        if event_dim is not None:
            try:
                return int(event_dim)
            except (ValueError, TypeError):
                pass  # 不正値の場合はフォールバックへ
        
        # フォールバック: event_dimがNULLのときはイベント日からその場で計算（EventDimLactCalculator と同じロジック）
        cow_auto_id = cow.get('auto_id')
        if cow_auto_id is not None:
            event_dim = self._compute_event_dim_for_event(
                cow_auto_id, latest_event.get('event_date', ''), 300
            )
            if event_dim is not None:
                return event_dim
        
        return None
    
    def _compute_event_dim_for_event(
        self, cow_auto_id: int, event_date: str, event_number: int
    ) -> Optional[int]:
        """
        event_dimがDBにない場合のフォールバック: イベント日と分娩日からDIMを計算する。
        EventDimLactCalculator と同じロジックを使用。
        """
        if not event_date or cow_auto_id is None:
            return None
        try:
            from modules.event_dim_lact_calculator import EventDimLactCalculator
            calculator = EventDimLactCalculator(self.db)
            event_dim, _ = calculator.calculate_event_dim_and_lact(
                cow_auto_id, event_date, event_number, exclude_event_id=None
            )
            return int(event_dim) if event_dim is not None else None
        except Exception as e:
            logging.debug(f"event_dim fallback calculation: {e}")
            return None
    
    def _get_fresh_check_note(self, events: list) -> Optional[str]:
        """
        フレッシュチェックイベントのNOTE内容を取得
        
        フレッシュチェックイベント（300）のNOTE内容
        - json_dataから動的に生成する（cow_card.pyの_display_eventsと同じロジック）
        - json_dataから生成できない場合は、既存のnoteフィールドを返す
        
        Args:
            events: イベントリスト
        
        Returns:
            NOTE内容（文字列）、存在しない場合はNone
        """
        # フレッシュチェックイベント（300）を取得
        fresh_check_events = [
            e for e in events
            if e.get('event_number') == 300  # FCHK
            and e.get('event_date')
        ]
        
        if not fresh_check_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            fresh_check_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のフレッシュチェックイベント
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        # 有効な値のみをリスト化（空文字、None、"-"は除外）
        valid_parts = []
        
        # 処置（新しいキー名と古いキー名の両方をサポート）
        treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
        if treatment and str(treatment).strip() and str(treatment).strip() != '-':
            valid_parts.append(str(treatment).strip())
        
        # 子宮所見（新しいキー名と古いキー名の両方をサポート）
        uterine = (json_data.get('uterine_findings') or 
                  json_data.get('uterus_findings') or 
                  json_data.get('uterus_finding') or 
                  json_data.get('uterus', ''))
        if uterine and str(uterine).strip() and str(uterine).strip() != '-':
            valid_parts.append(f"子宮{str(uterine).strip()}")
        
        # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
        left_ovary = (json_data.get('left_ovary_findings') or 
                     json_data.get('leftovary_findings') or 
                     json_data.get('leftovary_finding') or 
                     json_data.get('left_ovary', '') or
                     json_data.get('leftovary', ''))
        if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
            valid_parts.append(f"左{str(left_ovary).strip()}")
        
        # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
        right_ovary = (json_data.get('right_ovary_findings') or 
                      json_data.get('rightovary_findings') or 
                      json_data.get('rightovary_finding') or 
                      json_data.get('right_ovary', '') or
                      json_data.get('rightovary', ''))
        if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
            valid_parts.append(f"右{str(right_ovary).strip()}")
        
        # remark（新しいキー名と古いキー名の両方をサポート）
        # まず remark を取得、なければ other 系を確認
        remark = json_data.get('remark')
        if not remark:
            # remark が無い場合は other 系を確認
            remark = (json_data.get('other') or 
                     json_data.get('other_info') or 
                     json_data.get('other_findings', ''))
        if remark and str(remark).strip() and str(remark).strip() != '-':
            valid_parts.append(str(remark).strip())
        
        # 有効な項目のみを「  」で区切って返す
        if valid_parts:
            return "  ".join(valid_parts)
        
        # json_dataから生成できない場合は、既存のnoteフィールドを確認
        note = latest_event.get('note')
        return note.strip() if note else None
    
    def _calculate_dai_at_event_date(self, events: list, target_date: str) -> Optional[int]:
        """
        指定されたイベント日付時点での授精後日数（DAI）を計算
        
        Args:
            events: イベントリスト
            target_date: 対象日付（YYYY-MM-DD形式）
        
        Returns:
            授精後日数（日数）、計算できない場合はNone
        """
        if not target_date:
            return None
        
        try:
            target_dt = datetime.strptime(target_date, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 対象日付より前のAI/ETイベントを取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and e.get('event_date', '') <= target_date
        ]
        
        if not ai_et_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のAI/ETイベントを取得
        latest_event = sorted_events[0]
        event_date = latest_event.get('event_date')
        event_number = latest_event.get('event_number')
        
        if not event_date:
            return None
        
        # 授精日と検診日（target_date）の間に分娩イベント（202）があれば授精後日数は表示しない
        # （分娩を挟むと産次が変わり、その授精は当該検診には無関係となる。フレッシュは原則空欄）
        calving_event_number = 202  # 分娩
        for e in events:
            en = e.get('event_number')
            ed = e.get('event_date')
            if en != calving_event_number or not ed:
                continue
            if event_date < ed <= target_date:
                return None
        
        try:
            insemination_dt = datetime.strptime(event_date, '%Y-%m-%d')
            
            # ETイベントの場合は7日差し引く
            if event_number == 201:  # ET
                insemination_dt = insemination_dt - timedelta(days=7)
            
            # 対象日付から授精日までの日数
            dai = (target_dt - insemination_dt).days
            return dai if dai >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_days_from_previous_dry_to_calving(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        前産次の乾乳日から今産次の分娩日までの日数を計算
        
        前産次の乾乳月日（PDRYD）と今産次の分娩月日との差（日数）
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            日数、計算できない場合はNone
        """
        # 今産次の分娩日を取得
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 前産次の乾乳日を取得
        previous_dry_date = self._get_previous_dry_date(events, cow)
        if not previous_dry_date:
            return None
        
        try:
            current_date = datetime.strptime(current_clvd, '%Y-%m-%d')
            previous_dry_dt = datetime.strptime(previous_dry_date, '%Y-%m-%d')
            # 前産次の乾乳日から今産次の分娩日までの日数
            days = (current_date - previous_dry_dt).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_dam_jpn10(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        導入イベント時の母牛のJPN10を取得
        
        導入イベント（600）のjson_dataからdam_jpn10を取得
        
        Args:
            events: イベントリスト（現在の牛のイベント）
            cow: 牛データ
        
        Returns:
            母牛のJPN10（文字列）、存在しない場合はNone
        """
        # 導入イベント（600）を取得（日付順：古い順）
        import_events = [
            e for e in events
            if e.get('event_number') == 600  # IN（導入）
            and e.get('event_date')
        ]
        
        if not import_events:
            return None
        
        # 最初の導入イベントを取得（最も古い導入イベント）
        first_import = sorted(
            import_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )[0]
        
        import_json = first_import.get('json_data') or {}
        if isinstance(import_json, str):
            try:
                import_json = json.loads(import_json)
            except:
                import_json = {}
        
        # 導入イベントのjson_dataからdam_jpn10を取得
        dam_jpn10 = import_json.get('dam_jpn10')
        if dam_jpn10:
            return str(dam_jpn10).strip()
        
        return None
    
    def _calculate_dim(self, clvd: Optional[str]) -> Optional[int]:
        """
        DIM（分娩後日数）を計算
        
        Args:
            clvd: 最終分娩日（YYYY-MM-DD形式）
        
        Returns:
            DIM（日数）、clvdがNoneの場合はNone
        """
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
            today = datetime.now()
            dim = (today - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def calculate_dim_at_date(self, cow_auto_id: int, reference_date: str) -> Optional[int]:
        """
        指定日時点のDIM（分娩後日数）を計算。
        繁殖検診入力など、本日ではなく検診日等を基準にしたい場合に使用する。
        
        Args:
            cow_auto_id: 牛の auto_id
            reference_date: 基準日（YYYY-MM-DD形式）
        
        Returns:
            基準日時点のDIM（日数）、計算できない場合はNone
        """
        cow = self.db.get_cow_by_auto_id(cow_auto_id)
        if not cow:
            return None
        clvd = cow.get('clvd')
        if not clvd:
            return None
        return self._calculate_dim_from_date(clvd, reference_date)
    
    def _calculate_days_after_insemination(self, events: list) -> Optional[int]:
        """
        授精後日数を計算（当該産次のみ対象）。
        
        直近の分娩イベント(202)より後に発生したAI/ETのうち、
        もっとも新しい授精日から本日までの日数を返す。
        分娩後に授精がなければ None（空欄）を返す。
        分娩イベントが削除されると、前産次の授精が再度「直近分娩後」とみなされ復活する。
        
        AIイベントから本日までの日数（ETイベントであれば7日差し引いたうえで計算）
        
        Args:
            events: イベントリスト
        
        Returns:
            授精後日数（日数）、計算できない場合はNone
        """
        # 直近の分娩日を取得（202 = 分娩）
        calving_dates = [
            e.get('event_date') for e in events
            if e.get('event_number') == 202 and e.get('event_date')
        ]
        latest_calving_date = max(calving_dates) if calving_dates else None

        # 直近分娩より後のAI/ETのみ対象（その産次における授精後日数）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
        ]
        if latest_calving_date is not None:
            ai_et_events = [
                e for e in ai_et_events
                if e.get('event_date', '') > latest_calving_date
            ]
        
        if not ai_et_events:
            return None
        
        # event_date でソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のAI/ETイベントを取得
        latest_event = sorted_events[0]
        event_date = latest_event.get('event_date')
        event_number = latest_event.get('event_number')
        
        if not event_date:
            return None
        
        try:
            insemination_date = datetime.strptime(event_date, '%Y-%m-%d')
            
            # ETイベントの場合は7日差し引く
            if event_number == 201:  # ET
                insemination_date = insemination_date - timedelta(days=7)
            
            today = datetime.now()
            days = (today - insemination_date).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_last_items(self, cow_auto_id: int) -> Dict[str, Any]:
        """
        乳検（601）由来の LAST_ 系項目を計算
        
        item_dictionary.json の source 定義（例: "601.milk_yield"）から
        汎用的に解決する
        
        Args:
            cow_auto_id: 牛の auto_id
        
        Returns:
            LAST_ 系項目の辞書
        """
        # 互換性のために source 処理へ委譲
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        return self._calculate_source_items(events)
    
    def _get_last_ai_date(self, events: list) -> Optional[str]:
        """
        イベント履歴から最新のAI/ET日を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            最新のAI/ET日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
        ]
        
        if not ai_et_events:
            return None
        
        # event_date でソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        return sorted_events[0].get('event_date')
    
    def _calculate_due_date(self, events: list, clvd: Optional[str]) -> Optional[str]:
        """
        分娩予定日を計算
        
        Args:
            events: イベントリスト
            clvd: 最終分娩日（現在の産次での分娩日）
        
        Returns:
            分娩予定日（YYYY-MM-DD形式）、計算できない場合はNone
        """
        # 最新の妊娠鑑定プラスイベントを取得
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
        ]
        
        if not preg_events:
            return None
        
        # 最新の妊娠鑑定プラスイベント
        latest_preg = sorted(
            preg_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        
        preg_date = latest_preg.get('event_date')
        
        # 最新のAI/ETイベントを取得（妊娠鑑定より前）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date', '') <= preg_date
        ]
        
        if not ai_et_events:
            return None
        
        latest_ai_et = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        
        ai_et_date = latest_ai_et.get('event_date')
        event_number = latest_ai_et.get('event_number')
        
        try:
            # ETの場合は7日前、AIの場合は当日が受胎日
            if event_number == 201:  # ET
                conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d') - timedelta(days=7)
            else:  # AI
                conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d')
            
            # 分娩予定日 = 受胎日 + 280日
            due_date = conception_date + timedelta(days=280)
            return due_date.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None
    
    def _calculate_due_date_minus_days(self, events: list, clvd: Optional[str], days: int) -> Optional[str]:
        """
        分娩予定日から指定日数を引いた日付を計算
        
        Args:
            events: イベントリスト
            clvd: 最終分娩日（現在の産次での分娩日）
            days: 引く日数（60, 30, 21, 14, 7など）
        
        Returns:
            日付（YYYY-MM-DD形式）、計算できない場合はNone
        """
        due_date = self._calculate_due_date(events, clvd)
        if not due_date:
            return None
        
        try:
            due_dt = datetime.strptime(due_date, '%Y-%m-%d')
            target_date = due_dt - timedelta(days=days)
            return target_date.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None
    
    def _calculate_expected_calving_interval(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        予定分娩間隔（PCI）を計算
        
        今産次の分娩月日と予定分娩月日の間隔日数
        分娩予定日が決まっている（妊娠中または乾乳中）の牛のみ計算
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            予定分娩間隔（日数）、計算できない場合はNone
        """
        # 妊娠中（RC=5）または乾乳中（RC=6）の牛のみ計算
        rc = cow.get('rc')
        if rc not in [5, 6]:  # RC_PREGNANT = 5, RC_DRY = 6
            return None
        
        # 今産次の分娩日を取得
        current_clvd = cow.get('clvd')
        if not current_clvd:
            return None
        
        # 予定分娩日を取得
        due_date = self._calculate_due_date(events, current_clvd)
        if not due_date:
            return None
        
        try:
            clvd_date = datetime.strptime(current_clvd, '%Y-%m-%d')
            due_dt = datetime.strptime(due_date, '%Y-%m-%d')
            # 今産次の分娩日から予定分娩日までの日数
            interval = (due_dt - clvd_date).days
            return interval if interval >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_dopn(self, events: list, clvd: Optional[str], cow: Optional[Dict[str, Any]] = None) -> Optional[int]:
        """
        空胎日数（Days Open）を計算
        
        【重要】経産牛（lact > 0）かつ妊娠している牛（RC = 5）のみを対象とする。
        それ以外の牛はNoneを返す。
        
        分娩日から受胎までの日数
        - 受胎が確定している場合：分娩日から受胎のAIまでの日数（ETの場合はETイベント日から7日前）
        
        Args:
            events: イベントリスト
            clvd: 最終分娩日（YYYY-MM-DD形式）
            cow: 牛の情報（lact, rcを含む）
        
        Returns:
            空胎日数（整数）、計算できない場合または対象外の場合はNone
        """
        # 経産牛かつ妊娠している牛のみを対象とする
        if cow is None:
            return None
        
        lact = cow.get('lact')
        rc = cow.get('rc')
        
        # 経産牛（lact > 0）かつ妊娠している牛（RC = 5）のみを対象
        if not (lact and lact > 0 and rc == 5):
            return None
        
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 最新の妊娠鑑定プラスイベントを取得（受胎確定の判定）
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
        ]
        
        if preg_events:
            # 受胎が確定している場合
            # 最新の妊娠鑑定プラスイベント
            latest_preg = sorted(
                preg_events,
                key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                reverse=True
            )[0]
            
            preg_date = latest_preg.get('event_date')
            
            # 最新のAI/ETイベントを取得（妊娠鑑定より前）
            ai_et_events = [
                e for e in events
                if e.get('event_number') in [200, 201]  # AI, ET
                and e.get('event_date', '') <= preg_date
            ]
            
            if ai_et_events:
                latest_ai_et = sorted(
                    ai_et_events,
                    key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                    reverse=True
                )[0]
                
                ai_et_date = latest_ai_et.get('event_date')
                event_number = latest_ai_et.get('event_number')
                
                try:
                    # ETの場合は7日前、AIの場合は当日が受胎日
                    if event_number == 201:  # ET
                        conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d') - timedelta(days=7)
                    else:  # AI
                        conception_date = datetime.strptime(ai_et_date, '%Y-%m-%d')
                    
                    # 分娩日から受胎日までの日数
                    dopn = (conception_date - clvd_date).days
                    return dopn if dopn >= 0 else None
                except (ValueError, TypeError):
                    pass
        
        # 受胎が確定していない場合はNoneを返す（経産牛かつ妊娠している牛のみが対象のため）
        return None
    
    def _calculate_breeding_count(self, events: list) -> int:
        """
        授精回数を計算
        
        AIイベント（200）またはETイベント（201）により1加算される。
        ただし、前回のAI/ETイベントから7日以内の場合は1回のAI/ETとみなし、加算しない。
        
        Args:
            events: イベントリスト
        
        Returns:
            授精回数（整数）
        """
        # AI/ETイベントを取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')  # event_dateが存在するもののみ
        ]
        
        if not ai_et_events:
            return 0
        
        # 日付順にソート（古い順：昇順）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        count = 0
        last_date = None
        
        for event in sorted_events:
            event_date = event.get('event_date')
            if not event_date:
                continue
            
            # Rがついているイベントはカウント対象外
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            outcome = json_data.get('outcome')
            if outcome == 'R':
                # Rがついているイベントは授精回数のカウント対象にならない
                continue
            
            try:
                current_date = datetime.strptime(event_date, '%Y-%m-%d')
                
                # 最初のイベント、または前回から7日以上経過している場合はカウント
                if last_date is None:
                    count += 1
                    last_date = current_date
                else:
                    days_diff = (current_date - last_date).days
                    if days_diff >= 7:
                        count += 1
                        last_date = current_date
                    # 7日以内の場合はスキップ（last_dateは更新しない）
            except (ValueError, TypeError):
                # 日付パースエラーの場合はスキップ
                continue
        
        return count
    
    def _get_last_sire(self, events: list) -> Optional[str]:
        """
        最終使用SIREを取得
        
        直近のAIイベント（200）またはETイベント（201）のjson_dataからSIREを取得
        
        Args:
            events: イベントリスト
        
        Returns:
            SIRE（文字列）、存在しない場合はNone
        """
        # AI/ETイベントを取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')  # event_dateが存在するもののみ
        ]
        
        if not ai_et_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のAI/ETイベントからSIREを取得
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        sire = json_data.get('sire')
        if sire:
            return str(sire).strip()
        
        return None
    
    def _calculate_conception_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        受胎日（CONCDT）を計算
        
        ルール：
        1. AI/ETイベント後、初めての妊娠イベント（PDP/PDP2/PAGP）で受胎日が確定される
        2. その後、再度PDP/PAGPイベントが起こっても、初回の妊娠イベントでの受胎日は変更されない
        3. ただし、二回目の妊娠鑑定イベントがPDP2で、受胎日の指定が行われた場合はこれが優先される
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            受胎日（YYYY-MM-DD形式）、計算できない場合はNone
        """
        # 今産次の分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 今産次のAI/ETイベントを取得（分娩日より後）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and e.get('event_date', '') > clvd  # 分娩日より後
        ]
        
        if not ai_et_events:
            return None
        
        # 妊娠鑑定プラスイベントを取得（時系列順にソート）
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
            and e.get('event_date')
            and e.get('event_date', '') > clvd  # 分娩日より後
        ]
        
        if not preg_events:
            return None
        
        # 時系列順にソート（古い順）
        sorted_preg_events = sorted(
            preg_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 最初の妊娠イベントで受胎日を確定
        first_preg = sorted_preg_events[0]
        first_preg_event_number = first_preg.get('event_number')
        first_preg_date = first_preg.get('event_date')
        
        # 最初の妊娠イベントがPDP2で受胎日が指定されている場合
        if first_preg_event_number == 304:  # PDP2
            json_data = first_preg.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            ai_event_id = json_data.get('ai_event_id')
            if ai_event_id:
                # 指定されたAI/ETイベントを検索
                for event in ai_et_events:
                    if event.get('id') == ai_event_id:
                        ai_et_date = event.get('event_date')
                        event_number = event.get('event_number')
                        try:
                            ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
                            # ETの場合は7日前、AIの場合は当日が受胎日（設計書 10.3.6）
                            if event_number == 201:  # ET
                                conception_dt = ai_et_dt - timedelta(days=7)
                            else:  # AI
                                conception_dt = ai_et_dt
                            return conception_dt.strftime('%Y-%m-%d')
                        except (ValueError, TypeError):
                            return None
        
        # 最初の妊娠イベントがPDP/PAGPの場合、直近のAI/ETイベントを使用（妊娠鑑定より前）
        ai_et_before_first_preg = [
            e for e in ai_et_events
            if e.get('event_date', '') <= first_preg_date
        ]
        
        if not ai_et_before_first_preg:
            return None
        
        latest_ai_et = sorted(
            ai_et_before_first_preg,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        
        ai_et_date = latest_ai_et.get('event_date')
        event_number = latest_ai_et.get('event_number')
        
        try:
            ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
            # ETの場合は7日前、AIの場合は当日が受胎日（設計書 10.3.6）
            if event_number == 201:  # ET
                conception_dt = ai_et_dt - timedelta(days=7)
            else:  # AI
                conception_dt = ai_et_dt
            first_conception_date = conception_dt.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 二回目以降のPDP2イベントで受胎日が指定されている場合は優先
        if len(sorted_preg_events) > 1:
            for preg_event in sorted_preg_events[1:]:  # 二回目以降
                if preg_event.get('event_number') == 304:  # PDP2
                    json_data = preg_event.get('json_data') or {}
                    if isinstance(json_data, str):
                        try:
                            json_data = json.loads(json_data)
                        except:
                            json_data = {}
                    
                    ai_event_id = json_data.get('ai_event_id')
                    if ai_event_id:
                        # 指定されたAI/ETイベントを検索
                        for event in ai_et_events:
                            if event.get('id') == ai_event_id:
                                ai_et_date = event.get('event_date')
                                event_number = event.get('event_number')
                                try:
                                    ai_et_dt = datetime.strptime(ai_et_date, '%Y-%m-%d')
                                    # ETの場合は7日前、AIの場合は当日が受胎日（設計書 10.3.6）
                                    if event_number == 201:  # ET
                                        conception_dt = ai_et_dt - timedelta(days=7)
                                    else:  # AI
                                        conception_dt = ai_et_dt
                                    return conception_dt.strftime('%Y-%m-%d')
                                except (ValueError, TypeError):
                                    pass
        
        # 二回目以降のPDP2で指定がない場合は、最初の妊娠イベントでの受胎日を返す
        return first_conception_date
    
    def _get_conceiving_ai_et_event(self, events: list, cow: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        受胎したAIもしくはETイベントを1件取得する。
        
        優先順位:
        1. 今産次のAI/ETのうち json_data.outcome == 'P' のもの（複数あれば日付が最も新しい1件）
        2. 上記が無い場合: 妊娠鑑定プラス（PDP/PDP2/PAGP）から受胎イベントを推定
           - 最初の妊娠イベントがPDP2で ai_event_id 指定あり → そのAI/ET
           - それ以外 → 最初の妊娠鑑定日以前の直近のAI/ET
        
        Returns:
            受胎したAIもしくはETイベントの辞書、存在しない場合はNone
        """
        clvd = cow.get('clvd')
        if not clvd:
            return None
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]
            and e.get('event_date')
            and e.get('event_date', '') > clvd
        ]
        if not ai_et_events:
            return None

        def _outcome(e: Dict[str, Any]) -> Optional[str]:
            j = e.get('json_data') or {}
            if isinstance(j, str):
                try:
                    j = json.loads(j)
                except Exception:
                    j = {}
            return j.get('outcome')

        # 1. outcome == 'P' のAI/ETを優先
        events_p = [e for e in ai_et_events if _outcome(e) == 'P']
        if events_p:
            latest = sorted(
                events_p,
                key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
                reverse=True
            )[0]
            return latest

        # 2. フォールバック: 妊娠鑑定プラスから受胎イベントを推定（受胎日計算と同じロジック）
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
            and e.get('event_date')
            and e.get('event_date', '') > clvd
        ]
        if not preg_events:
            return None
        sorted_preg = sorted(preg_events, key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
        first_preg = sorted_preg[0]
        first_preg_number = first_preg.get('event_number')
        first_preg_date = first_preg.get('event_date', '')

        if first_preg_number == 304:  # PDP2
            j = first_preg.get('json_data') or {}
            if isinstance(j, str):
                try:
                    j = json.loads(j)
                except Exception:
                    j = {}
            ai_event_id = j.get('ai_event_id')
            if ai_event_id:
                for ev in ai_et_events:
                    if ev.get('id') == ai_event_id:
                        return ev

        ai_et_before = [e for e in ai_et_events if e.get('event_date', '') <= first_preg_date]
        if not ai_et_before:
            return None
        return sorted(
            ai_et_before,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
    
    def _get_conception_sire(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        受胎したAIもしくはETイベントのSIREを取得
        
        Returns:
            SIRE（文字列）、存在しない場合はNone
        """
        event = self._get_conceiving_ai_et_event(events, cow)
        if not event:
            return None
        json_data = event.get('json_data') or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except Exception:
                json_data = {}
        # 複数キーを試す（sire / sire_id）
        sire = json_data.get('sire') or json_data.get('sire_id')
        if sire:
            return str(sire).strip()
        return None
    
    def _get_conception_insemination_type(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        受胎したAIもしくはETイベントの授精種類を取得
        
        Returns:
            授精種類（コードまたは「コード：名称」形式の文字列）、存在しない場合はNone
        """
        event = self._get_conceiving_ai_et_event(events, cow)
        if not event:
            return None
        json_data = event.get('json_data') or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except Exception:
                json_data = {}
        # 複数キーを試す（insemination_type_code / insemination_type）
        raw = json_data.get('insemination_type_code') or json_data.get('insemination_type')
        if raw:
            s = str(raw).strip()
            # 「コード：名称」形式の場合はそのまま返す（表示用）
            return s if s else None
        return None
    
    def _calculate_twin_flag(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        双子フラグ（TWIN）を計算
        
        ルール：
        1. すべての妊娠プラスイベント（PDP, PDP2, PAGP）を確認し、どれか一つでもtwinがtrueの場合、「双子」を返す
        2. 妊娠していない個体は「未受胎」を返す
        3. 妊娠しているが双子とされていない場合は空文字列を返す
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            「双子」「未受胎」または空文字列
        """
        # すべての妊娠プラスイベント（PDP, PDP2, PAGP）を取得
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]  # PDP, PDP2, PAGP
            and e.get('event_date')
        ]
        
        if not preg_events:
            # 妊娠していない場合は「未受胎」
            return "未受胎"
        
        # すべての妊娠プラスイベントを確認し、どれか一つでもtwinがtrueの場合は「双子」を返す
        for preg_event in preg_events:
            json_data = preg_event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            # twinがtrueの場合は「双子」を返す
            if json_data.get('twin') is True:
                return "双子"
        
        # 妊娠しているが双子とされていない場合は空文字列
        return ""

    def _calculate_fetus_dairy_female_flag(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        胎子乳用メスフラグ（FDMF）を計算。
        目的は後継牛（乳用種♀）の見通しを予測すること。
        
        以下のいずれかを満たす場合に「○」を返す。
        1) 受胎のAI/ETのSIREが乳用種メス（sire_listで種別が乳用種♀）
        2) 受胎のAI/ETのSIREが乳用種（種別が乳用種レギュラー）かつ、
           妊娠鑑定プラスイベント（303, 304, 307）のいずれかで♀判定（female_judgment）にチェックあり
        
        ♀判定ありだけでは不十分。使用SIREが乳用種で、かつ♀判定がついている場合に○。
        
        Returns:
            「○」または空文字列
        """
        conceiving = self._get_conceiving_ai_et_event(events, cow)
        if not conceiving:
            return ""
        json_data = conceiving.get('json_data') or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except Exception:
                json_data = {}
        sire = (json_data.get('sire') or json_data.get('sire_id'))
        if not sire:
            return ""
        sire_name = str(sire).strip()
        farm_path = self.item_dict_path.parent if self.item_dict_path else None
        if not farm_path:
            return ""
        sire_list_path = farm_path / "sire_list.json"
        if not sire_list_path.exists():
            return ""
        try:
            with open(sire_list_path, 'r', encoding='utf-8') as f:
                sire_list = json.load(f)
        except Exception:
            return ""
        if not isinstance(sire_list, dict):
            return ""
        opts = sire_list.get(sire_name)
        if not isinstance(opts, dict):
            return ""
        st = sire_opts_to_type(opts)
        if st in (SIRE_TYPE_UNKNOWN_OTHER, SIRE_TYPE_F1, SIRE_TYPE_WAGYU):
            return ""
        sire_is_dairy = st in (SIRE_TYPE_HOLSTEIN_REGULAR, SIRE_TYPE_HOLSTEIN_FEMALE)
        sire_is_dairy_female = st == SIRE_TYPE_HOLSTEIN_FEMALE
        # 条件1: SIREが乳用種メス
        if sire_is_dairy_female:
            return "○"
        # 条件2: SIREが乳用種 かつ 妊娠鑑定プラスで♀判定あり
        if not sire_is_dairy:
            return ""
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]
            and e.get('event_date')
        ]
        for preg_event in preg_events:
            jd = preg_event.get('json_data') or {}
            if isinstance(jd, str):
                try:
                    jd = json.loads(jd)
                except Exception:
                    jd = {}
            if jd.get('female_judgment') is True:
                return "○"
        return ""

    def _calculate_fetus_sex_determination(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        胎子雌雄判別（FSD）を計算
        
        妊娠鑑定プラスイベント（303, 304, 307）のいずれかで
        ♀判定（female_judgment）にチェックあり → 「メス」
        ♂判定（male_judgment）にチェックあり → 「オス」
        いずれもなし → 空文字列（Null）
        ♀♂両方ある場合はメスを優先。
        
        Returns:
            「メス」「オス」または空文字列（Null）
        """
        preg_events = [
            e for e in events
            if e.get('event_number') in [303, 304, 307]
            and e.get('event_date')
        ]
        for preg_event in preg_events:
            json_data = preg_event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            if json_data.get('female_judgment') is True:
                return "メス"
        for preg_event in preg_events:
            json_data = preg_event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            if json_data.get('male_judgment') is True:
                return "オス"
        return ""

    def get_ai_conception_status(self, events: list, cow: Dict[str, Any]) -> Dict[int, str]:
        """
        AI/ETイベントごとの受胎ステータスを取得
        
        【重要】outcomeは計算せず、event.json_data["outcome"]から取得する
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            イベントIDをキーにした受胎ステータスの辞書 {event_id: status}
            outcomeが存在しない場合は空文字列を返す
        """
        result = {}
        
        # 今産次の分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return result
        
        # 今産次のAI/ETイベントを取得（分娩日より後、時系列順）
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and e.get('event_date', '') > clvd  # 分娩日より後
        ]
        
        if not ai_et_events:
            return result
        
        # 各AI/ETイベントのjson_dataからoutcomeを取得
        for event in ai_et_events:
            event_id = event.get('id')
            if not event_id:
                continue
            
            # json_dataからoutcomeを取得
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            outcome = json_data.get('outcome')
            # outcomeが存在する場合はそのまま返す、存在しない場合は空文字列
            result[event_id] = outcome if outcome else ''
        
        return result
    
    def _get_last_insemination_type(self, events: list) -> Optional[str]:
        """
        最終AIの授精種類を取得
        
        直近のAIイベント（200）のjson_dataから授精種類コードを取得
        
        Args:
            events: イベントリスト
        
        Returns:
            授精種類コード（文字列）、存在しない場合はNone
        """
        # AIイベント（200のみ、ETは除外）を取得
        ai_events = [
            e for e in events
            if e.get('event_number') == 200  # AIのみ
            and e.get('event_date')  # event_dateが存在するもののみ
        ]
        
        if not ai_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のAIイベントから授精種類コードを取得
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        insemination_type = json_data.get('insemination_type_code')
        if insemination_type:
            return str(insemination_type).strip()
        
        return None
    
    def _event_insemination_count_is_one(self, event: Dict[str, Any]) -> bool:
        """イベントの json_data で授精回数が1回目かどうかを返す。"""
        json_data = event.get('json_data') or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except Exception:
                json_data = {}
        c = json_data.get('insemination_count')
        if c is None:
            return False
        try:
            return int(c) == 1
        except (ValueError, TypeError):
            return False

    def _calculate_dimfai(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        初回授精日数（DIM First AI）を計算
        
        その産次の「授精回数1回目」のAI/ETイベントの分娩後日数。
        - 経産牛のみ（lact >= 1）
        - 授精回数1回目のイベントが存在しない場合（途中導入など）はNone
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回授精日数（整数）、計算できない場合はNone
        """
        # 経産牛のみ（lact >= 1）
        lact = cow.get('lact')
        if lact is None or lact < 1:
            return None
        
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩日より後のAI/ETイベントのうち、授精回数1回目のもののみ対象
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and e.get('event_date', '') > clvd  # 分娩日より後
            and self._event_insemination_count_is_one(e)
        ]
        
        if not ai_et_events:
            return None
        
        # 日付順にソート（古い順：昇順）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 授精回数1回目のイベントの日付
        first_ai_et = sorted_events[0]
        first_ai_date = first_ai_et.get('event_date')
        
        try:
            first_ai_datetime = datetime.strptime(first_ai_date, '%Y-%m-%d')
            # 分娩日から初回授精日までの日数
            dimfai = (first_ai_datetime - clvd_date).days
            return dimfai if dimfai >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_agefai(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        未経産初回授精月齢（Age First AI）を計算
        
        未経産牛のみ（lact == 0 または None）
        「授精回数1回目」の授精月齢（月単位）。授精回数1回目のイベントがない場合はNone。
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回授精月齢（浮動小数点数）、計算できない場合はNone
        """
        # 未経産牛のみ（lact == 0 または None）
        lact = cow.get('lact')
        if lact is not None and lact > 0:
            return None
        
        # 生年月日を取得
        bthd = cow.get('bthd')
        if not bthd:
            return None
        
        try:
            birth_date = datetime.strptime(bthd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 授精回数1回目のAI/ETイベントのみ対象
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
            and self._event_insemination_count_is_one(e)
        ]
        
        if not ai_et_events:
            return None
        
        # 日付順にソート（古い順：昇順）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        first_ai_et = sorted_events[0]
        first_ai_date = first_ai_et.get('event_date')
        
        try:
            first_ai_datetime = datetime.strptime(first_ai_date, '%Y-%m-%d')
            days_diff = (first_ai_datetime - birth_date).days
            if days_diff < 0:
                return None
            age_months = days_diff / 30.44
            return round(age_months, 1)
        except (ValueError, TypeError):
            return None
    
    def _get_last_bcs(self, events: list) -> Optional[float]:
        """
        直近のBCSイベントのBCS値を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            BCS値（浮動小数点数）、存在しない場合はNone
        """
        # BCSイベント（101）を取得
        bcs_events = [
            e for e in events
            if e.get('event_number') == 101  # BCS
            and e.get('event_date')
        ]
        
        if not bcs_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            bcs_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のBCSイベントから値を取得
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        bcs_value = json_data.get('bcs')
        if bcs_value is not None:
            try:
                return float(bcs_value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_previous_bcs(self, events: list) -> Optional[float]:
        """
        直近のBCSイベントのひとつ前のBCSイベントのBCS値を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            BCS値（浮動小数点数）、存在しない場合はNone
        """
        # BCSイベント（101）を取得
        bcs_events = [
            e for e in events
            if e.get('event_number') == 101  # BCS
            and e.get('event_date')
        ]
        
        if len(bcs_events) < 2:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            bcs_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 2番目（直近のひとつ前）のBCSイベントから値を取得
        previous_event = sorted_events[1]
        json_data = previous_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        bcs_value = json_data.get('bcs')
        if bcs_value is not None:
            try:
                return float(bcs_value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_last_mastitis_date(self, events: list) -> Optional[str]:
        """
        直近の乳房炎イベントの日付を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            日付（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 乳房炎イベント（400）を取得
        mastitis_events = [
            e for e in events
            if e.get('event_number') == 400  # MAST
            and e.get('event_date')
        ]
        
        if not mastitis_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            mastitis_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乳房炎イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _calculate_last_mastitis_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        直近の乳房炎イベントのDIMを計算
        
        直近の乳房炎イベントの日付から分娩日までの日数
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            DIM（日数）、計算できない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 直近の乳房炎イベントの日付を取得
        last_mastitis_date = self._get_last_mastitis_date(events)
        if not last_mastitis_date:
            return None
        
        try:
            mastitis_date = datetime.strptime(last_mastitis_date, '%Y-%m-%d')
            # 分娩日から直近の乳房炎日までの日数
            dim = (mastitis_date - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_last_lame_date(self, events: list) -> Optional[str]:
        """
        直近のLAMEイベントの日付を取得
        
        Args:
            events: イベントリスト
        
        Returns:
            日付（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # LAMEイベント（402）を取得
        lame_events = [
            e for e in events
            if e.get('event_number') == 402  # LAME
            and e.get('event_date')
        ]
        
        if not lame_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            lame_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のLAMEイベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _calculate_days_from_last_mastitis(self, events: list) -> Optional[int]:
        """
        前回乳房炎イベントからの日数を計算
        
        Args:
            events: イベントリスト
        
        Returns:
            日数、計算できない場合はNone
        """
        last_mastitis_date = self._get_last_mastitis_date(events)
        if not last_mastitis_date:
            return None
        
        try:
            mastitis_date = datetime.strptime(last_mastitis_date, '%Y-%m-%d')
            today = datetime.now()
            days = (today - mastitis_date).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_days_from_last_lame(self, events: list) -> Optional[int]:
        """
        前回LAMEイベントからの日数を計算
        
        Args:
            events: イベントリスト
        
        Returns:
            日数、計算できない場合はNone
        """
        last_lame_date = self._get_last_lame_date(events)
        if not last_lame_date:
            return None
        
        try:
            lame_date = datetime.strptime(last_lame_date, '%Y-%m-%d')
            today = datetime.now()
            days = (today - lame_date).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_last_lame_note(self, events: list) -> Optional[str]:
        """
        直近のLAMEイベントのNOTEを取得
        
        Args:
            events: イベントリスト
        
        Returns:
            NOTE文字列、存在しない場合はNone
        """
        # LAMEイベント（402）を取得
        lame_events = [
            e for e in events
            if e.get('event_number') == 402  # LAME
            and e.get('event_date')
        ]
        
        if not lame_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            lame_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のLAMEイベントのNOTEを取得
        latest_event = sorted_events[0]
        note = latest_event.get('note')
        return note.strip() if note else None
    
    def _calculate_days_from_last_reproduction_check(self, events: list) -> Optional[int]:
        """
        直近検診日からの日数を計算
        
        Args:
            events: イベントリスト
        
        Returns:
            日数、計算できない場合はNone
        """
        last_repro_date = self._get_last_reproduction_check_date(events)
        if not last_repro_date:
            return None
        
        try:
            repro_date = datetime.strptime(last_repro_date, '%Y-%m-%d')
            today = datetime.now()
            days = (today - repro_date).days
            return days if days >= 0 else None
        except (ValueError, TypeError):
            return None

    def _get_latest_event_number(self, events: list) -> Optional[int]:
        """
        直近イベントのイベント番号を取得

        Args:
            events: イベントリスト

        Returns:
            イベント番号（int）、存在しない場合はNone
        """
        events_with_date = [
            e for e in events
            if e.get('event_date') and e.get('event_number') is not None
        ]
        if not events_with_date:
            return None

        latest_event = sorted(
            events_with_date,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )[0]
        try:
            return int(latest_event.get('event_number'))
        except (ValueError, TypeError):
            return None
    
    def _get_last_milk_test_value(self, events: list, cow: Dict[str, Any], field_name: str) -> Optional[float]:
        """
        今産次の最新乳検イベントから指定フィールドの値を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
            field_name: 取得するフィールド名（例: 'milk_yield', 'fat'）
        
        Returns:
            値（浮動小数点数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not milk_test_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乳検イベントから値を取得
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        value = json_data.get(field_name)
        if value is not None:
            try:
                return float(value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _calculate_max_milk_yield(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        その産次の最高乳量を計算
        
        現在の産次＝分娩日（clvd）以降の乳検とし、その乳検の乳量の最大値を返す。
        event_lact が一致する乳検があればそれを優先するが、一致しなくても
        event_date >= clvd の乳検は今産次として扱う（event_lact の取り残し対策）。
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            最高乳量（浮動小数点数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        # 今産次＝分娩日（clvd）以降の乳検。event_lact のずれがあっても clvd 以降はすべて対象
        lact_events = [
            e for e in milk_test_events
            if e.get('event_date') and e.get('event_date') >= clvd
        ]
        
        if not lact_events:
            return None
        
        # 乳量を取得して最大値を返す
        milk_yields = []
        for event in lact_events:
            json_data = event.get('json_data') or {}
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except Exception:
                    json_data = {}
            milk_yield = json_data.get('milk_yield')
            if milk_yield is not None:
                try:
                    milk_yields.append(float(milk_yield))
                except (ValueError, TypeError):
                    continue
        
        if milk_yields:
            return max(milk_yields)
        return None
    
    def _calculate_max_milk_yield_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        その産次の最高乳量時のDIMを計算
        
        分娩日（clvd）以降の乳検イベント（601）から、最高乳量を記録したイベントの event_dim を取得。
        _calculate_max_milk_yield と同じく clvd 以降の乳検を今産次として扱う。
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            最高乳量時のDIM（整数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        # 今産次＝分娩日（clvd）以降の乳検
        lact_events = [
            e for e in milk_test_events
            if e.get('event_date') and e.get('event_date') >= clvd
        ]
        
        if not lact_events:
            return None
        
        # すべての乳検イベントから乳量とDIMを取得
        max_milk_yield = None
        max_milk_yield_dim = None
        
        for event in lact_events:
            json_data = event.get('json_data') or {}
            
            # json_dataが文字列の場合はパース
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            milk_yield = json_data.get('milk_yield')
            if milk_yield is not None:
                try:
                    milk_yield_float = float(milk_yield)
                    # 最高乳量を更新
                    if max_milk_yield is None or milk_yield_float > max_milk_yield:
                        max_milk_yield = milk_yield_float
                        # event_dimを取得（NULLの場合はその場で計算）
                        event_dim = event.get('event_dim')
                        if event_dim is not None:
                            try:
                                max_milk_yield_dim = int(event_dim)
                            except (ValueError, TypeError):
                                pass
                        else:
                            cow_auto_id = cow.get('auto_id')
                            if cow_auto_id is not None:
                                dim = self._compute_event_dim_for_event(
                                    cow_auto_id, event.get('event_date', ''), 601
                                )
                                if dim is not None:
                                    max_milk_yield_dim = dim
                except (ValueError, TypeError):
                    continue
        
        return max_milk_yield_dim
    
    def _calculate_max_linear_score(self, events: list, cow: Dict[str, Any]) -> Optional[float]:
        """
        その産次の最大リニアスコアを計算
        
        現在の産次（lact）に対応する乳検イベント（601）から、リニアスコア（ls）の最大値を取得
        event_lactフィールドがある場合はそれを使用し、ない場合は分娩日（clvd）以降の乳検を対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            最大リニアスコア（浮動小数点数）、存在しない場合はNone
        """
        # 現在の産次を取得
        current_lact = cow.get('lact')
        if current_lact is None:
            return None
        
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        # event_lactフィールドがある場合はそれを使用して産次でフィルタ
        lact_events = []
        for event in milk_test_events:
            event_lact = event.get('event_lact')
            event_date = event.get('event_date')
            
            # event_lactがある場合は、現在の産次と一致するもののみ
            if event_lact is not None:
                if event_lact == current_lact:
                    lact_events.append(event)
            else:
                # event_lactがない場合は、分娩日以降の乳検を対象とする（後方互換性のため）
                if event_date and event_date >= clvd:
                    lact_events.append(event)
        
        if not lact_events:
            return None
        
        # すべての乳検イベントからリニアスコアを取得
        linear_scores = []
        for event in lact_events:
            json_data = event.get('json_data') or {}
            
            # json_dataが文字列の場合はパース
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            ls_value = json_data.get('ls')
            if ls_value is not None:
                try:
                    linear_scores.append(float(ls_value))
                except (ValueError, TypeError):
                    continue
        
        # 最大値を返す
        if linear_scores:
            return max(linear_scores)
        
        return None
    
    def _calculate_max_linear_score_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        その産次の最大リニアスコア時のDIMを計算
        
        現在の産次（lact）に対応する乳検イベント（601）から、最大リニアスコアを記録したイベントのevent_dimを取得
        event_lactフィールドがある場合はそれを使用し、ない場合は分娩日（clvd）以降の乳検を対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            最大リニアスコア時のDIM（整数）、存在しない場合はNone
        """
        # 現在の産次を取得
        current_lact = cow.get('lact')
        if current_lact is None:
            return None
        
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
        ]
        
        if not milk_test_events:
            return None
        
        # event_lactフィールドがある場合はそれを使用して産次でフィルタ
        lact_events = []
        for event in milk_test_events:
            event_lact = event.get('event_lact')
            event_date = event.get('event_date')
            
            # event_lactがある場合は、現在の産次と一致するもののみ
            if event_lact is not None:
                if event_lact == current_lact:
                    lact_events.append(event)
            else:
                # event_lactがない場合は、分娩日以降の乳検を対象とする（後方互換性のため）
                if event_date and event_date >= clvd:
                    lact_events.append(event)
        
        if not lact_events:
            return None
        
        # 最大リニアスコアとそのDIMを追跡
        max_linear_score = None
        max_linear_score_dim = None
        
        for event in lact_events:
            json_data = event.get('json_data') or {}
            
            # json_dataが文字列の場合はパース
            if isinstance(json_data, str):
                try:
                    json_data = json.loads(json_data)
                except:
                    json_data = {}
            
            ls_value = json_data.get('ls')
            if ls_value is not None:
                try:
                    ls_float = float(ls_value)
                    # 最大リニアスコアを更新
                    if max_linear_score is None or ls_float > max_linear_score:
                        max_linear_score = ls_float
                        # event_dimを取得（NULLの場合はその場で計算）
                        event_dim = event.get('event_dim')
                        if event_dim is not None:
                            try:
                                max_linear_score_dim = int(event_dim)
                            except (ValueError, TypeError):
                                pass
                        else:
                            cow_auto_id = cow.get('auto_id')
                            if cow_auto_id is not None:
                                dim = self._compute_event_dim_for_event(
                                    cow_auto_id, event.get('event_date', ''), 601
                                )
                                if dim is not None:
                                    max_linear_score_dim = dim
                except (ValueError, TypeError):
                    continue
        
        return max_linear_score_dim
    
    def _get_previous_milk_test_value(self, events: list, cow: Dict[str, Any], field_name: str) -> Optional[float]:
        """
        今産次の2回目（直近のひとつ前）の乳検イベントから指定フィールドの値を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
            field_name: 取得するフィールド名（例: 'milk_yield'）
        
        Returns:
            値（浮動小数点数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if len(milk_test_events) < 2:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 2回目（直近のひとつ前）の乳検イベントから値を取得
        previous_event = sorted_events[1]
        json_data = previous_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        value = json_data.get(field_name)
        if value is not None:
            try:
                return float(value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_previous_previous_milk_test_value(self, events: list, cow: Dict[str, Any], field_name: str) -> Optional[float]:
        """
        今産次の3回目（前々回）の乳検イベントから指定フィールドの値を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
            field_name: 取得するフィールド名（例: 'scc'）
        
        Returns:
            値（浮動小数点数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if len(milk_test_events) < 3:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 3回目（前々回）の乳検イベントから値を取得
        previous_previous_event = sorted_events[2]
        json_data = previous_previous_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        value = json_data.get(field_name)
        if value is not None:
            try:
                return float(value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_nth_milk_test_value(self, events: list, cow: Dict[str, Any], field_name: str, n: int) -> Optional[float]:
        """
        今産次のN回目の乳検イベントから指定フィールドの値を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        n=1が1回目（最も古い）、n=2が2回目、n=3が3回目
        
        Args:
            events: イベントリスト
            cow: 牛データ
            field_name: 取得するフィールド名（例: 'bhb'）
            n: 何回目のイベントか（1-indexed、1が1回目）
        
        Returns:
            値（浮動小数点数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if len(milk_test_events) < n:
            return None
        
        # 日付順にソート（昇順：古い順、1回目が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # N回目の乳検イベントから値を取得（0-indexedなのでn-1）
        nth_event = sorted_events[n - 1]
        json_data = nth_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        value = json_data.get(field_name)
        if value is not None:
            try:
                return float(value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_milk_test_value_by_month(self, events: list, cow: Dict[str, Any], field_name: str, month: int, which: int = 1) -> Optional[float]:
        """
        指定月の直近の乳検イベントから指定フィールドの値を取得
        
        which=1: 指定月1の設定を使用
        which=2: 指定月2の設定を使用
        
        Args:
            events: イベントリスト
            cow: 牛データ
            field_name: 取得するフィールド名（例: 'milk_yield'）
            month: 指定月（1-12、Noneの場合は設定から取得）
            which: どちらの設定を使用するか（1または2）
        
        Returns:
            値（浮動小数点数）、存在しない場合はNone
        """
        # 月が指定されていない場合は設定から取得
        if month is None:
            settings = get_app_settings_manager()
            if which == 1:
                month = settings.get_selected_month_1()
            else:
                month = settings.get_selected_month_2()
        
        if month is None or month < 1 or month > 12:
            return None
        
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not milk_test_events:
            return None
        
        # 指定月のイベントをフィルタリング
        target_month_events = []
        for event in milk_test_events:
            event_date = event.get('event_date')
            if event_date:
                try:
                    event_dt = datetime.strptime(event_date, '%Y-%m-%d')
                    if event_dt.month == month:
                        target_month_events.append(event)
                except (ValueError, TypeError):
                    continue
        
        if not target_month_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            target_month_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のイベントから値を取得
        latest_event = sorted_events[0]
        json_data = latest_event.get('json_data') or {}
        
        # json_dataが文字列の場合はパース
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        value = json_data.get(field_name)
        if value is not None:
            try:
                return float(value)
            except (ValueError, TypeError):
                return None
        
        return None
    
    def _get_last_milk_test_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        今産次の最新乳検イベントの日付を取得
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            日付（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not milk_test_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乳検イベントの日付を取得
        latest_event = sorted_events[0]
        return latest_event.get('event_date')
    
    def _get_last_milk_test_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        直近の乳検日のDIMを計算
        
        今産次の最新乳検イベント時の分娩後日数（DIM）
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            DIM（日数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not milk_test_events:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新の乳検イベントの日付を取得
        latest_event = sorted_events[0]
        latest_event_date = latest_event.get('event_date')
        
        if not latest_event_date:
            return None
        
        try:
            latest_event_datetime = datetime.strptime(latest_event_date, '%Y-%m-%d')
            # 分娩日から最新の乳検日までの日数
            dim = (latest_event_datetime - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_nth_milk_test_dim(self, events: list, cow: Dict[str, Any], n: int) -> Optional[int]:
        """
        今産次のN回目の乳検イベント時の分娩後日数（DIM）を計算
        
        分娩日（clvd）以降の乳検イベント（601）のみを対象とする
        n=1が1回目（最も古い）、n=2が2回目、n=3が3回目
        
        Args:
            events: イベントリスト
            cow: 牛データ
            n: 何回目のイベントか（1-indexed、1が1回目）
        
        Returns:
            DIM（日数）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩日以降の乳検イベント（601）を取得
        milk_test_events = [
            e for e in events
            if e.get('event_number') == 601  # 乳検
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if len(milk_test_events) < n:
            return None
        
        # 日付順にソート（昇順：古い順、1回目が先頭）
        sorted_events = sorted(
            milk_test_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # N回目の乳検イベントの日付を取得
        nth_event = sorted_events[n - 1]
        nth_event_date = nth_event.get('event_date')
        
        if not nth_event_date:
            return None
        
        try:
            nth_event_datetime = datetime.strptime(nth_event_date, '%Y-%m-%d')
            # 分娩日からN回目の乳検日までの日数
            dim = (nth_event_datetime - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_ai_interval(self, events: list) -> Optional[int]:
        """
        授精間隔を計算
        
        直近AIイベントとその前のAIイベントの間隔（日数）
        ETイベントでは7日引いたうえで計算する
        授精回数が1回もしくは0回の場合はNone
        
        Args:
            events: イベントリスト
        
        Returns:
            授精間隔（日数）、計算できない場合はNone
        """
        # AI/ETイベント（200: AI, 201: ET）を取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
        ]
        
        # 授精回数が1回もしくは0回の場合はNone
        if len(ai_et_events) < 2:
            return None
        
        # 日付順にソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のイベントとその前のイベント
        latest_event = sorted_events[0]
        previous_event = sorted_events[1]
        
        latest_date = latest_event.get('event_date')
        previous_date = previous_event.get('event_date')
        latest_event_number = latest_event.get('event_number')
        previous_event_number = previous_event.get('event_number')
        
        if not latest_date or not previous_date:
            return None
        
        try:
            # 日付をdatetimeオブジェクトに変換
            latest_dt = datetime.strptime(latest_date, '%Y-%m-%d')
            previous_dt = datetime.strptime(previous_date, '%Y-%m-%d')
            
            # ETイベントの場合は7日引く
            if latest_event_number == 201:  # ET
                latest_dt = latest_dt - timedelta(days=7)
            
            if previous_event_number == 201:  # ET
                previous_dt = previous_dt - timedelta(days=7)
            
            # 間隔を計算（最新 - 前回）
            interval = (latest_dt - previous_dt).days
            
            return interval if interval >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _calculate_first_mastitis_dim(self, events: list, cow: Dict[str, Any]) -> Optional[int]:
        """
        初回乳房炎DIMを計算
        
        分娩後初回の乳房炎イベントのDIM
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回乳房炎DIM（日数）、計算できない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        try:
            clvd_date = datetime.strptime(clvd, '%Y-%m-%d')
        except (ValueError, TypeError):
            return None
        
        # 分娩日以降の乳房炎イベント（400）を取得
        mastitis_events = [
            e for e in events
            if e.get('event_number') == 400  # MAST
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not mastitis_events:
            return None
        
        # 日付順にソート（古い順：昇順）
        sorted_events = sorted(
            mastitis_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 最初の乳房炎イベント（初回）
        first_mastitis = sorted_events[0]
        first_mastitis_date = first_mastitis.get('event_date')
        
        try:
            first_mastitis_datetime = datetime.strptime(first_mastitis_date, '%Y-%m-%d')
            # 分娩日から初回乳房炎日までの日数
            dim = (first_mastitis_datetime - clvd_date).days
            return dim if dim >= 0 else None
        except (ValueError, TypeError):
            return None
    
    def _get_first_mastitis_date(self, events: list, cow: Dict[str, Any]) -> Optional[str]:
        """
        初回の乳房炎イベントの日付を取得
        
        分娩後初回の乳房炎イベントの日付
        
        Args:
            events: イベントリスト
            cow: 牛データ
        
        Returns:
            初回乳房炎日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 分娩日を取得
        clvd = cow.get('clvd')
        if not clvd:
            return None
        
        # 分娩日以降の乳房炎イベント（400）を取得
        mastitis_events = [
            e for e in events
            if e.get('event_number') == 400  # MAST
            and e.get('event_date')
            and e.get('event_date', '') >= clvd  # 分娩日以降
        ]
        
        if not mastitis_events:
            return None
        
        # 日付順にソート（古い順：昇順）
        sorted_events = sorted(
            mastitis_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0))
        )
        
        # 最初の乳房炎イベント（初回）の日付を取得
        first_mastitis = sorted_events[0]
        return first_mastitis.get('event_date')
    
    def _calculate_next_estrus_date(self, events: list, cow: Dict[str, Any] = None) -> Optional[str]:
        """
        直近の授精から21日後の発情予定日を計算
        
        - AIイベント（200）の場合：AI日 + 21日
        - ETイベント（201）の場合：ET日 - 7日 + 21日 = ET日 + 14日
        
        繁殖コードがBred（RC=3、授精後）の牛のみ表示、それ以外はNoneを返す。
        直近繁殖イベントが繁殖検査であっても、直近のAIイベントで受胎が確認されていない場合
        （つまり繁殖コードがBredの場合）は次回発情予定日を表示する。
        
        Args:
            events: イベントリスト
            cow: 牛データ（RCを取得するため）
        
        Returns:
            発情予定日（YYYY-MM-DD形式）、存在しない場合はNone
        """
        # 繁殖コードがBred（RC=3、授精後）の牛のみ計算、それ以外はNoneを返す
        if cow:
            rc = cow.get('rc')
            # RC_BRED = 3（授精後）の場合のみ計算
            if rc != 3:
                return None
        
        # AI/ETイベントを取得
        ai_et_events = [
            e for e in events
            if e.get('event_number') in [200, 201]  # AI, ET
            and e.get('event_date')
        ]
        
        if not ai_et_events:
            return None
        
        # event_date でソート（降順：最新が先頭）
        sorted_events = sorted(
            ai_et_events,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        
        # 最新のAI/ETイベントを取得
        latest_event = sorted_events[0]
        event_number = latest_event.get('event_number')
        event_date = latest_event.get('event_date')
        
        if not event_date:
            return None
        
        try:
            event_datetime = datetime.strptime(event_date, '%Y-%m-%d')
            
            # AIイベント（200）の場合：+21日
            if event_number == 200:
                estrus_date = event_datetime + timedelta(days=21)
            # ETイベント（201）の場合：-7日+21日 = +14日
            elif event_number == 201:
                estrus_date = event_datetime + timedelta(days=14)
            else:
                return None
            
            return estrus_date.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None
    
    def _calculate_source_items(self, events: list, target_item: Optional[str] = None) -> Dict[str, Any]:
        """
        item_dictionary の source 定義 (event_number.field) を汎用的に解決
        
        Args:
            events: イベントリスト
            target_item: 計算対象の項目キー（指定時はこの項目のみ計算、Noneの場合は全項目計算）
        """
        result: Dict[str, Any] = {}
        
        # target_item が指定されている場合、その項目のみを処理
        items_to_process = [target_item] if target_item else list(self.item_dictionary.keys())
        
        # event_date が存在するイベントのみをフィルタリング
        events_with_date = [e for e in events if e.get('event_date')]
        # イベントを event_number ごとに最新順へソート
        events_sorted = sorted(
            events_with_date,
            key=lambda e: (e.get('event_date', ''), e.get('id', 0)),
            reverse=True
        )
        latest_by_number: Dict[int, Dict[str, Any]] = {}
        for ev in events_sorted:
            num = ev.get('event_number')
            if num is None:
                continue
            if num not in latest_by_number:
                latest_by_number[num] = ev

        for item_key in items_to_process:
            if item_key not in self.item_dictionary:
                continue
            item_def = self.item_dictionary[item_key]
            source = item_def.get('source')
            if not source or '.' not in source:
                continue
            event_num_str, field_key = source.split('.', 1)
            try:
                event_num = int(event_num_str)
            except ValueError:
                continue
            latest_event = latest_by_number.get(event_num)
            if not latest_event:
                # target_item が指定されている場合のみログ出力
                if target_item:
                    logging.debug(f"[FormulaEngine] {item_key}: イベント{event_num}が見つかりません")
                result[item_key] = None
                continue
            json_data = latest_event.get('json_data') or {}
            # json_data になければトップレベルも参照
            value = json_data.get(field_key)
            if value is None:
                value = latest_event.get(field_key)
            # target_item が指定されている場合のみログ出力
            if target_item:
                logging.debug(f"[FormulaEngine] {item_key}: source={source}, value={value}, json_data={json_data}")
            result[item_key] = value
        return result

    def latest_event_value(self, event_number: int, key: str) -> Optional[Any]:
        """
        指定されたイベント番号の最新レコードから指定キーの値を取得
        
        Args:
            event_number: イベント番号
            key: json_data 内のキー名
        
        Returns:
            値（存在しない場合は None）
        """
        # 現在の cow_auto_id を取得する必要があるため、
        # このメソッドは _calculate_formula_items 内で呼び出されることを想定
        # cow_auto_id は self._current_cow_auto_id に保存されている
        if not hasattr(self, '_current_cow_auto_id'):
            return None
        
        cow_auto_id = self._current_cow_auto_id
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        
        # 指定されたイベント番号のイベントをフィルタリング
        target_events = [
            e for e in events
            if e.get("event_number") == event_number and e.get("event_date")
        ]
        
        if not target_events:
            return None
        
        # event_date でソートして最新を取得
        latest = max(
            target_events,
            key=lambda e: e["event_date"],
            default=None
        )
        
        if not latest:
            return None
        
        # json_data から値を取得
        json_data = latest.get("json_data") or {}
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except:
                json_data = {}
        
        return json_data.get(key)

    def _calculate_formula_items(
        self,
        cow: Dict[str, Any],
        events: list,
        current: Dict[str, Any],
        cow_auto_id: int,
        target_item: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        formula 定義を評価して値を算出
        
        Args:
            cow: 牛データ
            events: イベントリスト
            current: 現在の計算結果
            cow_auto_id: 牛のauto_id
            target_item: 計算対象の項目キー（指定時はこの項目のみ計算、Noneの場合は全項目計算）
        """
        # latest_event_value で使用するために cow_auto_id を保存
        self._current_cow_auto_id = cow_auto_id
        
        results: Dict[str, Any] = {}
        safe_builtins = {
            "min": min,
            "max": max,
            "sum": sum,
            "len": len,
            "abs": abs,
            "round": round,
            "math": math,
        }
        
        # target_item が指定されている場合、その項目のみを処理
        items_to_process = [target_item] if target_item else list(self.item_dictionary.keys())
        
        for item_key in items_to_process:
            if item_key not in self.item_dictionary:
                continue
            item_def = self.item_dictionary[item_key]
            formula = item_def.get("formula")
            if not formula:
                continue
            local_ctx = {
                "cow": cow,
                "events": events,
                "result": {**current, **results},  # 逐次計算された値を利用可能に
                "auto_id": cow_auto_id,
                "cow_id": cow.get("cow_id"),
                "latest_event_value": lambda en, k: self.latest_event_value(en, k),  # latest_event_value を利用可能に
                "age": lambda bthd: self._calculate_age(bthd),  # AGE 計算関数
                "age_at_calving": lambda bthd, clvd: self._calculate_age_at_calving(bthd, clvd),  # AGECLV 計算関数
                "age_at_first_calving": lambda evs, cw: self._calculate_age_at_first_calving(evs, cw),  # AGEFCLV 計算関数
                "calf_info_by_lact": lambda evs, cw, lact: self._get_calf_info_by_lact(evs, cw, lact, cow_auto_id),  # CALF1/CALF2/CALF3 計算関数
                "previous_calving_date": lambda evs, cw, lact: self._get_previous_calving_date(evs, cw, lact),  # 前産次の分娩日 計算関数
                "calving_interval": lambda evs, cw: self._calculate_calving_interval(evs, cw),  # CINT 計算関数
                "predicted_calving_interval": lambda evs, cw: self._calculate_expected_calving_interval(evs, cw),  # PCI 計算関数
                "days_from_conception": lambda evs, cw: self._calculate_days_from_conception(evs, cw),  # DCC 計算関数
                "dry_date": lambda evs, cw: self._get_dry_date(evs, cw),  # DRYD 計算関数
                "previous_dry_date": lambda evs, cw: self._get_previous_dry_date(evs, cw),  # PDRYD 計算関数
                "stopr_date": lambda evs: self._get_stopr_date(evs),  # STOPR 計算関数
                "stopr_dim": lambda evs, cw: self._calculate_stopr_dim(evs, cw),  # STOPRD 計算関数
                "last_reproduction_check_date": lambda evs: self._get_last_reproduction_check_date(evs),  # LREPRO 計算関数
                "last_reproduction_check_note": lambda evs: self._get_last_reproduction_check_note(evs),  # REPCT 計算関数
                "fresh_check_date": lambda evs: self._get_fresh_check_date(evs),  # FCHKDATE 計算関数
                "fresh_check_dim": lambda evs, cw: self._get_fresh_check_dim(evs, cw),  # FCHKDIM 計算関数
                "fresh_check_note": lambda evs: self._get_fresh_check_note(evs),  # FCHKNOTE 計算関数
                "days_from_previous_dry_to_calving": lambda evs, cw: self._calculate_days_from_previous_dry_to_calving(evs, cw),  # DPDC 計算関数
                "dam_jpn10": lambda evs, cw: self._get_dam_jpn10(evs, cw),  # DAM 計算関数
                "dim": lambda clvd: self._calculate_dim(clvd),  # DIM 計算関数
                "latest_event_number": lambda evs: self._get_latest_event_number(evs),  # LEVNT 計算関数
                "days_after_insemination": lambda evs: self._calculate_days_after_insemination(evs),  # DAI 計算関数
                "last_ai_date": lambda evs: self._get_last_ai_date(evs),  # LASTAI 計算関数
                "due_date": lambda evs, clvd: self._calculate_due_date(evs, clvd),  # DUE 計算関数
                "due_date_minus_60": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 60),  # DUEM60 計算関数
                "due_date_minus_40": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 40),  # DUEM40 計算関数
                "due_date_minus_30": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 30),  # DUEM30 計算関数
                "due_date_minus_21": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 21),  # DUEM21 計算関数
                "due_date_minus_14": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 14),  # DUEM14 計算関数
                "due_date_minus_7": lambda evs, clvd: self._calculate_due_date_minus_days(evs, clvd, 7),  # DUEM7 計算関数
                "breeding_count": lambda evs: self._calculate_breeding_count(evs),  # BRED 計算関数
                "last_sire": lambda evs: self._get_last_sire(evs),  # LSIR 計算関数
                "last_insemination_type": lambda evs: self._get_last_insemination_type(evs),  # LSIT 計算関数
                "dopn": lambda evs, clvd: self._calculate_dopn(evs, clvd, cow),  # DOPN 計算関数
                "previous_lact_days_open": lambda evs, cw: self._calculate_previous_lact_days_open(evs, cw),  # PDO 計算関数
                "previous_lact_gestation_days": lambda evs, cw: self._calculate_previous_lact_gestation_days(evs, cw),  # PGEST 計算関数
                "dimfai": lambda evs, cw: self._calculate_dimfai(evs, cw),  # DIMFAI 計算関数
                "agefai": lambda evs, cw: self._calculate_agefai(evs, cw),  # AGEFAI 計算関数
                "last_bcs": lambda evs: self._get_last_bcs(evs),  # BCS 計算関数
                "previous_bcs": lambda evs: self._get_previous_bcs(evs),  # PBCS 計算関数
                "last_mastitis_date": lambda evs: self._get_last_mastitis_date(evs),  # LMAST 計算関数
                "last_mastitis_dim": lambda evs, cw: self._calculate_last_mastitis_dim(evs, cw),  # LMASTD 計算関数
                "days_from_last_mastitis": lambda evs: self._calculate_days_from_last_mastitis(evs),  # DMAST 計算関数
                "last_lame_date": lambda evs: self._get_last_lame_date(evs),  # LLAME 計算関数
                "days_from_last_lame": lambda evs: self._calculate_days_from_last_lame(evs),  # DLAME 計算関数
                "last_lame_note": lambda evs: self._get_last_lame_note(evs),  # LLAMEN 計算関数
                "days_from_last_reproduction_check": lambda evs: self._calculate_days_from_last_reproduction_check(evs),  # DREPRO 計算関数
                "milk_test_value_by_month_1": lambda evs, cw, field: self._get_milk_test_value_by_month(evs, cw, field, None, 1),  # 任意月の乳検データ1 計算関数
                "milk_test_value_by_month_2": lambda evs, cw, field: self._get_milk_test_value_by_month(evs, cw, field, None, 2),  # 任意月の乳検データ2 計算関数
                "last_milk_test_value": lambda evs, cw, field: self._get_last_milk_test_value(evs, cw, field),  # MILK/MFAT/MPROT/MUN/SCC/LS/BHB 計算関数
                "previous_milk_test_value": lambda evs, cw, field: self._get_previous_milk_test_value(evs, cw, field),  # PMILK/PSCC/PLS 計算関数
                "previous_previous_milk_test_value": lambda evs, cw, field: self._get_previous_previous_milk_test_value(evs, cw, field),  # PPSCC 計算関数
                "linear_score_before_dry": lambda evs, cw: self._get_linear_score_before_dry(evs, cw),  # DRYLS 計算関数
                "previous_dry_linear_score": lambda evs, cw: self._get_previous_dry_linear_score(evs, cw),  # PDRYLS 計算関数
                "milk_test_date_before_dry": lambda evs, cw: self._get_milk_test_date_before_dry(evs, cw),  # DRYTD 計算関数
                "nth_milk_test_value": lambda evs, cw, field, n: self._get_nth_milk_test_value(evs, cw, field, n),  # 1STBHB/2NDBHB/3RDBHB/LS1ST/LS2ND/LS3RD 計算関数
                "last_milk_test_date": lambda evs, cw: self._get_last_milk_test_date(evs, cw),  # TDATE 計算関数
                "last_milk_test_dim": lambda evs, cw: self._get_last_milk_test_dim(evs, cw),  # TDIM 計算関数
                "nth_milk_test_dim": lambda evs, cw, n: self._get_nth_milk_test_dim(evs, cw, n),  # 1STTDIM 計算関数
                "next_estrus_date": lambda evs, cw: self._calculate_next_estrus_date(evs, cw),  # 次回発情予定日 計算関数
                "first_mastitis_dim": lambda evs, cw: self._calculate_first_mastitis_dim(evs, cw),  # FMASTD 計算関数
                "first_mastitis_date": lambda evs, cw: self._get_first_mastitis_date(evs, cw),  # FMAST 計算関数
                "conception_date": lambda evs, cw: self._calculate_conception_date(evs, cw),  # CONCDT 計算関数
                "conception_sire": lambda evs, cw: self._get_conception_sire(evs, cw),  # CSIR 受胎SIRE 計算関数
                "conception_insemination_type": lambda evs, cw: self._get_conception_insemination_type(evs, cw),  # CSIT 受胎授精種類 計算関数
                "twin_flag": lambda evs, cw: self._calculate_twin_flag(evs, cw),  # TWIN 計算関数
                "fetus_dairy_female_flag": lambda evs, cw: self._calculate_fetus_dairy_female_flag(evs, cw),  # FDMF 胎子乳用メスフラグ
                "fetus_sex_determination": lambda evs, cw: self._calculate_fetus_sex_determination(evs, cw),  # FSD 胎子雌雄判別
                "calving_month": lambda cw: self._get_calving_month(cw),  # 分娩月 計算関数
                "previous_calving_month": lambda evs, cw: self._get_previous_calving_month(evs, cw),  # 前産分娩月 計算関数
                "calving_year_month": lambda clvd: self._calculate_calving_year_month(clvd),  # 分娩年月 計算関数
                "birth_year": lambda bthd: self._calculate_birth_year(bthd),  # 生まれた年（西暦）計算関数
                "birth_year_month": lambda bthd: self._calculate_birth_year_month(bthd),  # 生まれた年月 計算関数
                "first_ai_date": lambda evs, cw: self._calculate_first_ai_date(evs, cw),  # 初回授精日 計算関数
                "nth_milk_test_date": lambda evs, cw, n: self._get_nth_milk_test_date(evs, cw, n),  # 1STTD/2NDTD 計算関数
                "ai_interval": lambda evs: self._calculate_ai_interval(evs),  # 授精間隔 計算関数
                "max_milk_yield": lambda evs, cw: self._calculate_max_milk_yield(evs, cw),  # 最高乳量 計算関数
                "max_milk_yield_dim": lambda evs, cw: self._calculate_max_milk_yield_dim(evs, cw),  # 最高乳量時のDIM 計算関数
                "max_linear_score": lambda evs, cw: self._calculate_max_linear_score(evs, cw),  # 最大リニアスコア 計算関数
                "max_linear_score_dim": lambda evs, cw: self._calculate_max_linear_score_dim(evs, cw),  # 最大リニアスコア時のDIM 計算関数
            }
            try:
                value = eval(formula, {"__builtins__": safe_builtins}, local_ctx)
                # デバッグログ（LAST_DENOVO_FA の場合のみ、またはtarget_itemが指定されている場合のみ）
                if item_key == "LAST_DENOVO_FA" or target_item:
                    if item_key == "LAST_DENOVO_FA":
                        logging.info(
                            f"[Formula] LAST_DENOVO_FA={value} type={type(value)}"
                        )
                results[item_key] = value
            except Exception as e:
                # 計算失敗時は None にして継続
                # target_item が指定されている場合、または DUEM*（分娩予定〇日前）の場合はログ出力
                if target_item or (item_key and item_key.startswith("DUEM")):
                    logging.warning(f"[Formula] 計算失敗: item_key={item_key}, formula={formula}, error={e}")
                results[item_key] = None
        
        # クリーンアップ
        if hasattr(self, '_current_cow_auto_id'):
            delattr(self, '_current_cow_auto_id')
        
        return results
    
    def get_item_value(self, cow_auto_id: int, item_key: str) -> Any:
        """
        特定の計算項目の値を取得
        
        Args:
            cow_auto_id: 牛の auto_id
            item_key: アイテムキー（例: "DIM", "LAST_MILK_YIELD"）
        
        Returns:
            計算値、存在しない場合はNone
        """
        calculated = self.calculate(cow_auto_id)
        return calculated.get(item_key)
    
    def set_item_dictionary_path(self, item_dictionary_path: Path):
        """
        item_dictionary.json のパスを設定（農場切替時など）
        
        Args:
            item_dictionary_path: item_dictionary.json のパス
        """
        self.item_dict_path = item_dictionary_path
        self._load_item_dictionary()

