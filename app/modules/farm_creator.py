"""
FALCON2 - 農場作成モジュール
新規農場作成時に乳検速報CSVを読み込んで初期データを構築する
"""

import shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple
import pandas as pd

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager


class FarmCreator:
    """新規農場作成クラス"""
    
    def __init__(self, farm_path: Path):
        """
        初期化
        
        Args:
            farm_path: 作成する農場のパス（例: C:/FARMS/FarmA）
        """
        self.farm_path = Path(farm_path)
        self.db_path = self.farm_path / "farm.db"
    
    def create_farm(
        self,
        farm_name: str,
        csv_path: Optional[Path] = None,
        template_farm_path: Optional[Path] = None
    ) -> Path:
        """
        新規農場を作成
        
        Args:
            farm_name: 農場名
            csv_path: 乳検速報CSVファイルのパス（オプション）
            template_farm_path: テンプレート農場のパス（オプション）
        
        Returns:
            作成された農場のパス
        """
        # 既存チェック（既存の場合は削除して再作成）
        if self.farm_path.exists() and self.db_path.exists():
            import shutil
            print(f"既存の農場を削除します: {self.farm_path}")
            shutil.rmtree(self.farm_path)
        
        # 農場フォルダを作成
        self.farm_path.mkdir(parents=True, exist_ok=True)
        
        # データベースを作成
        db = DBHandler(self.db_path)
        
        # 設定ファイルを作成
        settings = SettingsManager(self.farm_path)
        settings.set("farm_name", farm_name)
        
        # 項目辞書・イベント辞書は本体（config_default）を参照するため農場にはコピーしない
        
        # テンプレート農場から設定をコピー（オプション）
        if template_farm_path:
            self._copy_template_settings(template_farm_path)
        
        # CSVからデータを読み込む（オプション）
        if csv_path:
            self._import_from_csv(db, csv_path)
        
        db.close()
        
        return self.farm_path
    
    def _copy_template_settings(self, template_farm_path: Path):
        """テンプレート農場から設定をコピー"""
        template_settings = SettingsManager(template_farm_path)
        template_data = template_settings.load()
        
        target_settings = SettingsManager(self.farm_path)
        target_settings.load()
        
        # farm_settings.json 内の設定を引き継ぐ
        for key in ["pen_settings", "repro_checkup_settings", "inseminator_codes", "insemination_type_codes"]:
            if key in template_data:
                target_settings.set(key, template_data.get(key))
        
        # 授精設定ファイル（授精師・授精種類）をコピー
        insemination_settings_file = template_farm_path / "insemination_settings.json"
        if insemination_settings_file.exists():
            shutil.copy2(insemination_settings_file, self.farm_path / "insemination_settings.json")
        
        # 繁殖処置設定ファイルをコピー
        reproduction_treatment_settings_file = template_farm_path / "reproduction_treatment_settings.json"
        if reproduction_treatment_settings_file.exists():
            shutil.copy2(reproduction_treatment_settings_file, self.farm_path / "reproduction_treatment_settings.json")
    
    def _import_from_csv(self, db: DBHandler, csv_path: Path):
        """
        CSVファイルからデータをインポート
        
        処理フロー:
        1) CSV読み込み
        2) cow 作成（産次はCSV値）
        3) 分娩イベント（CALV, baseline_calving=true）作成
        4) 乳検イベント（601）作成
        5) RuleEngine.apply_events(cow_auto_id)
        """
        # CSVを読み込む
        test_date, data_rows = self._parse_csv(csv_path)
        
        # RuleEngineを初期化
        rule_engine = RuleEngine(db)
        
        # 各データ行を処理
        for row_data in data_rows:
            cow_id = row_data['cow_id']
            jpn10 = row_data['jpn10']
            clvd = row_data.get('clvd')  # 分娩月日
            lact = row_data.get('lact', 0)  # 産次（baseline）
            
            # 既存の牛をチェック（通常は存在しないが、念のため）
            existing_cow = db.get_cow_by_id(cow_id)
            
            if existing_cow:
                cow_auto_id = existing_cow['auto_id']
            else:
                # 新しい牛を作成（産次はCSVの値をそのまま使用）
                cow_data = {
                    'cow_id': cow_id,
                    'jpn10': jpn10,
                    'brd': row_data.get('brd', 'ホルスタイン'),
                    'bthd': row_data.get('bthd'),
                    'entr': None,
                    'lact': lact,  # CSVの産次をそのまま使用（baseline）
                    'clvd': clvd,  # 分娩月日（後でイベントで更新される）
                    'rc': None,  # RuleEngineで後から更新
                    'pen': row_data.get('pen', 'Lactating'),
                    'frm': self.farm_path.name
                }
                
                cow_auto_id = db.insert_cow(cow_data)
            
            # 分娩イベント（CALV）を作成（baseline_calvingフラグ付き）
            # 【重要】分娩月日が存在する場合は必ず分娩イベントを作成
            if clvd:
                # 既存の分娩イベントをチェック
                events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
                existing_calv = False
                for event in events:
                    if event.get('event_number') == RuleEngine.EVENT_CALV:
                        event_json = event.get('json_data', {})
                        # baseline_calvingフラグがある分娩イベントが既に存在するかチェック
                        if event_json.get('baseline_calving', False) and event.get('event_date') == clvd:
                            existing_calv = True
                            break
                
                if not existing_calv:
                    calv_event_data = {
                        'cow_auto_id': cow_auto_id,
                        'event_number': RuleEngine.EVENT_CALV,
                        'event_date': clvd,
                        'json_data': {'baseline_calving': True},  # 産次を増やさないフラグ
                        'note': 'CSVから作成（baseline）'
                    }
                    
                    db.insert_event(calv_event_data)
            
            # 乳検イベント（601）を作成（1頭につき1件のみ）
            events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
            existing_milk_test = False
            for event in events:
                if event.get('event_number') == RuleEngine.EVENT_MILK_TEST:
                    if event.get('event_date') == test_date:
                        existing_milk_test = True
                        break
            
            if not existing_milk_test:
                json_data = {}
                
                # 存在するデータのみ追加
                if row_data.get('milk_yield') is not None:
                    json_data['milk_yield'] = row_data['milk_yield']
                if row_data.get('fat') is not None:
                    json_data['fat'] = row_data['fat']
                if row_data.get('protein') is not None:
                    json_data['protein'] = row_data['protein']
                if row_data.get('snf') is not None:
                    json_data['snf'] = row_data['snf']
                if row_data.get('scc') is not None:
                    json_data['scc'] = row_data['scc']
                if row_data.get('ls') is not None:
                    json_data['ls'] = row_data['ls']
                if row_data.get('mun') is not None:
                    json_data['mun'] = row_data['mun']
                if row_data.get('bhb') is not None:
                    json_data['bhb'] = row_data['bhb']
                if row_data.get('denovo_fa') is not None:
                    json_data['denovo_fa'] = row_data['denovo_fa']
                if row_data.get('preformed_fa') is not None:
                    json_data['preformed_fa'] = row_data['preformed_fa']
                if row_data.get('mixed_fa') is not None:
                    json_data['mixed_fa'] = row_data['mixed_fa']
                if row_data.get('denovo_milk') is not None:
                    json_data['denovo_milk'] = row_data['denovo_milk']
                
                milk_event_data = {
                    'cow_auto_id': cow_auto_id,
                    'event_number': RuleEngine.EVENT_MILK_TEST,
                    'event_date': test_date,
                    'json_data': json_data if json_data else None,
                    'note': None
                }
                
                db.insert_event(milk_event_data)
            
            # RuleEngine.apply_events() を呼ぶ（状態を再計算）
            rule_engine.apply_events(cow_auto_id)
            # 状態を更新
            rule_engine._recalculate_and_update_cow(cow_auto_id)
    
    def _parse_csv(self, csv_path: Path) -> Tuple[str, List[Dict[str, Any]]]:
        """
        CSVファイルを読み込んで、検定日とデータ行を返す
        
        Args:
            csv_path: CSVファイルのパス
        
        Returns:
            (検定日, データ行のリスト)
        """
        # Shift-JISエンコーディングで読み込む
        lines = []
        with open(csv_path, 'r', encoding='shift_jis') as f:
            for line in f:
                lines.append(line.strip().split(','))
        
        # DataFrameに変換
        max_cols = max(len(row) for row in lines) if lines else 0
        for row in lines:
            while len(row) < max_cols:
                row.append('')
        
        df = pd.DataFrame(lines, dtype=str)
        
        # 検定日を取得（3行目、0列目）
        test_date = None
        if len(df) > 3:
            date_str = df.iloc[3, 0] if len(df.columns) > 0 else None
            if pd.notna(date_str):
                date_str = str(date_str).strip()
                try:
                    date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                    test_date = date_obj.strftime('%Y-%m-%d')
                except:
                    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d']:
                        try:
                            date_obj = datetime.strptime(date_str, fmt)
                            test_date = date_obj.strftime('%Y-%m-%d')
                            break
                        except:
                            pass
        
        if not test_date:
            # 全行を検索
            for idx in range(len(df)):
                for col_idx in range(min(3, len(df.columns))):
                    val = df.iloc[idx, col_idx] if len(df.columns) > col_idx else None
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if '/' in val_str and len(val_str) >= 8:
                            try:
                                date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                                test_date = date_obj.strftime('%Y-%m-%d')
                                break
                            except:
                                pass
                if test_date:
                    break
        
        if not test_date:
            raise ValueError("検定日が見つかりません")
        
        # ヘッダー行を取得（7行目、0列目から）
        header_row_idx = 7
        if len(df) <= header_row_idx:
            raise ValueError("ヘッダー行が見つかりません")
        
        headers = df.iloc[header_row_idx].tolist()
        
        # カラムインデックスを検索
        col_jpn10 = 0  # 個体識別番号
        col_cow_id = 1  # 拡大管理番号
        col_brd = 2  # 品種
        
        # 列位置を直接指定（Excel列名から数値インデックスに変換）
        # A=0, B=1, ..., Z=25, AA=26, AB=27, ..., AU=46, AX=49, AZ=51, BB=53
        def excel_col_to_index(col_name: str) -> int:
            """Excel列名（A, B, ..., Z, AA, AB, ...）を数値インデックスに変換"""
            result = 0
            for char in col_name:
                result = result * 26 + (ord(char.upper()) - ord('A') + 1)
            return result - 1  # 0-indexedに変換
        
        # 指定された列位置を使用
        col_milk_yield = excel_col_to_index('D')  # D列 = 3
        col_fat = excel_col_to_index('G')  # G列 = 6（乳脂率）
        col_snf = excel_col_to_index('J')  # J列 = 9（無脂固形分）
        col_protein = excel_col_to_index('L')  # L列 = 11（蛋白率）
        col_scc = excel_col_to_index('Q')  # Q列 = 16（体細胞）
        col_mun = excel_col_to_index('T')  # T列 = 19（MUN）
        col_ls = excel_col_to_index('X')  # X列 = 23（体細胞スコア/リニアスコア）
        col_bhb = excel_col_to_index('AA')  # AA列 = 26（BHB）
        col_denovo_fa = excel_col_to_index('AU')  # AU列 = 46（デノボFA）
        col_preformed_fa = excel_col_to_index('AX')  # AX列 = 49（プレフォームFA）
        col_mixed_fa = excel_col_to_index('AZ')  # AZ列 = 51（ミックスFA）
        col_denovo_milk = excel_col_to_index('BB')  # BB列 = 53（デノボMilk）
        
        # その他の列（ヘッダーから検索）
        col_bthd = self._find_column_index(headers, ['生年月日', '生年']) or 32
        col_brd_data = self._find_column_index(headers, ['品種']) or 33
        
        # 分娩月日: 複数のパターンで検索
        # 注意: 「分娩月日」と「最終分娩日」は別の列
        # - 「分娩月日」（31列目）: 実際の分娩日
        # - 「最終分娩日」（33列目）: DIM（分娩後日数）の可能性がある
        col_clvd = self._find_column_index(headers, ['分娩月日'])
        if col_clvd is None:
            col_clvd = self._find_column_index(headers, ['分娩', '月日'])
        if col_clvd is None:
            # フォールバック: デバッグ出力から31列目が「分娩月日」
            col_clvd = 31
        
        # 産次: 複数のパターンで検索
        col_lact = self._find_column_index(headers, ['産次', '産'])
        if col_lact is None:
            col_lact = self._find_column_index(headers, ['産次'])
        if col_lact is None:
            col_lact = 35  # フォールバック
        
        # 群情報（pen）を検索
        col_pen = self._find_column_index(headers, ['群', '管理'])
        
        # データ行を取得（8行目以降）
        data_rows = []
        for idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[idx].tolist()
            
            # 空行をスキップ
            if not any(pd.notna(val) and str(val).strip() for val in row[:5]):
                continue
            
            # 個体識別番号と拡大4桁を取得
            jpn10 = str(row[col_jpn10]).strip() if len(row) > col_jpn10 and pd.notna(row[col_jpn10]) else None
            cow_id_4digit = str(row[col_cow_id]).strip() if len(row) > col_cow_id and pd.notna(row[col_cow_id]) else None
            
            # 個体識別番号が10桁の数字でない場合はスキップ（参考値・集計行などを除外）
            if not jpn10:
                continue
            
            # JPN10が10桁の数字であることをチェック
            if len(jpn10) != 10 or not jpn10.isdigit():
                continue
            
            # cow_idを決定（拡大4桁があればそれを使用、なければJPN10の下4桁）
            if cow_id_4digit and len(cow_id_4digit) >= 4:
                cow_id = cow_id_4digit[-4:]  # 最後の4桁
            else:
                cow_id = jpn10[-4:]  # JPN10の下4桁
            
            # 品種
            brd = "ホルスタイン"  # デフォルト
            if col_brd_data and len(row) > col_brd_data and pd.notna(row[col_brd_data]):
                brd_val = str(row[col_brd_data]).strip()
                if brd_val:
                    brd = brd_val
            elif len(row) > col_brd and pd.notna(row[col_brd]):
                brd_val = str(row[col_brd]).strip()
                if brd_val:
                    brd = brd_val
            
            # 生年月日
            bthd = None
            if col_bthd and len(row) > col_bthd and pd.notna(row[col_bthd]):
                val_str = str(row[col_bthd]).strip()
                if '/' in val_str and len(val_str) >= 8:
                    try:
                        date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                        bthd = date_obj.strftime('%Y-%m-%d')
                    except:
                        pass
            
            # 分娩月日（clvd）: CSVから読み込む
            clvd = None
            if col_clvd is not None and len(row) > col_clvd:
                val = row[col_clvd]
                if pd.notna(val):
                    val_str = str(val).strip()
                    # 日付形式の文字列を検出（YYYY/MM/DD または YYYY-MM-DD）
                    if val_str and (('/' in val_str and len(val_str) >= 8) or ('-' in val_str and len(val_str) >= 8)):
                        try:
                            # YYYY/MM/DD形式を試す
                            date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                            clvd = date_obj.strftime('%Y-%m-%d')
                        except:
                            try:
                                # YYYY-MM-DD形式を試す
                                date_obj = datetime.strptime(val_str, '%Y-%m-%d')
                                clvd = date_obj.strftime('%Y-%m-%d')
                            except:
                                pass
            
            # 産次（CSVの値をそのまま使用）
            lact = 0
            if col_lact and len(row) > col_lact and pd.notna(row[col_lact]):
                try:
                    lact = int(float(str(row[col_lact]).strip()))
                except:
                    pass
            
            # 群情報
            pen = "Lactating"  # デフォルト
            if col_pen and len(row) > col_pen and pd.notna(row[col_pen]):
                pen_val = str(row[col_pen]).strip()
                if pen_val:
                    pen = pen_val
            
            # 乳検データを取得（当日値のみ）
            def get_float_value(col_idx: Optional[int]) -> Optional[float]:
                if col_idx is None or len(row) <= col_idx:
                    return None
                val = row[col_idx]
                if pd.notna(val):
                    try:
                        return float(str(val).strip())
                    except:
                        pass
                return None
            
            def get_int_value(col_idx: Optional[int]) -> Optional[int]:
                if col_idx is None or len(row) <= col_idx:
                    return None
                val = row[col_idx]
                if pd.notna(val):
                    try:
                        val_str = str(val).strip()
                        val_str = val_str.replace('*', '').replace('**', '').replace('***', '')
                        if val_str:
                            return int(float(val_str))
                    except:
                        pass
                return None
            
            # 乳検データを取得（当日値のみ）
            milk_yield = get_float_value(col_milk_yield)
            fat = get_float_value(col_fat)
            protein = get_float_value(col_protein)
            snf = get_float_value(col_snf)
            scc = get_int_value(col_scc)
            
            # 体細胞スコア（LS）: 整数値として取得
            ls = get_int_value(col_ls)
            if ls is None:
                # 整数として取得できない場合は浮動小数点数として取得してから整数に変換
                ls_float = get_float_value(col_ls)
                if ls_float is not None:
                    ls = int(ls_float)
            
            # MUN: 浮動小数点数として取得
            mun = get_float_value(col_mun)
            
            bhb = get_float_value(col_bhb)
            denovo_fa = get_float_value(col_denovo_fa)
            preformed_fa = get_float_value(col_preformed_fa)
            mixed_fa = get_float_value(col_mixed_fa)
            denovo_milk = get_float_value(col_denovo_milk)
            
            # デバッグ: 分娩月日が読み込まれているか確認（最初の数頭のみ）
            if not clvd and len([r for r in data_rows if r.get('cow_id') == cow_id]) == 0 and len(data_rows) < 3:
                print(f"デバッグ: cow_id={cow_id}, col_clvd={col_clvd}, row[col_clvd]={row[col_clvd] if col_clvd is not None and len(row) > col_clvd else 'N/A'}")
            
            data_rows.append({
                'cow_id': cow_id,
                'jpn10': jpn10,
                'brd': brd,
                'bthd': bthd,
                'clvd': clvd,  # 分娩月日
                'lact': lact,  # 産次（baseline）
                'pen': pen,
                'milk_yield': milk_yield,
                'fat': fat,
                'protein': protein,
                'snf': snf,
                'scc': scc,
                'ls': ls,
                'mun': mun,
                'bhb': bhb,
                'denovo_fa': denovo_fa,
                'preformed_fa': preformed_fa,
                'mixed_fa': mixed_fa,
                'denovo_milk': denovo_milk
            })
        
        return test_date, data_rows
    
    def _find_column_index(self, headers: List[str], keywords: List[str]) -> Optional[int]:
        """
        ヘッダー行からキーワードを含む列のインデックスを検索
        
        Args:
            headers: ヘッダー行のリスト
            keywords: 検索キーワードのリスト（すべてのキーワードが含まれる列を検索）
        
        Returns:
            見つかった列のインデックス、見つからない場合はNone
        """
        for idx, header in enumerate(headers):
            if pd.notna(header):
                header_str = str(header).strip().lower()
                # すべてのキーワードが含まれる列を検索
                if all(keyword.lower() in header_str for keyword in keywords):
                    return idx
        return None

