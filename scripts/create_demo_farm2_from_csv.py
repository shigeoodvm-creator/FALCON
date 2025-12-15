"""
FALCON2 - 乳検CSVからデモ農場2を作成
CSVファイルを読み込んで、DemoFarm2を作成し、牛・分娩イベント・乳検イベントを登録する

【仕様】
- 産次はCSVの値をそのまま使用（baseline）
- 分娩イベントはbaseline_calvingフラグ付きで作成（産次を増やさない）
"""

import sys
import json
import shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, List
import pandas as pd

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager


def find_column_index(headers: List[str], keywords: List[str]) -> Optional[int]:
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


def parse_csv(csv_path: Path) -> tuple[str, List[Dict[str, Any]]]:
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
    col_bthd = find_column_index(headers, ['生年月日', '生年']) or 32
    col_brd_data = find_column_index(headers, ['品種']) or 33
    col_clvd = find_column_index(headers, ['分娩', '最終']) or 34  # 最終分娩日（分娩月日）
    col_lact = find_column_index(headers, ['産次', '産']) or 35
    
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
        
        # 分娩月日（clvd）
        clvd = None
        if col_clvd and len(row) > col_clvd and pd.notna(row[col_clvd]):
            val_str = str(row[col_clvd]).strip()
            if '/' in val_str and len(val_str) >= 8:
                try:
                    date_obj = datetime.strptime(val_str, '%Y/%m/%d')
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
        
        data_rows.append({
            'cow_id': cow_id,
            'jpn10': jpn10,
            'brd': brd,
            'bthd': bthd,
            'clvd': clvd,  # 分娩月日
            'lact': lact,  # 産次（baseline）
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


def create_demo_farm2(csv_path: Path) -> Path:
    """
    デモ農場2を作成
    
    Args:
        csv_path: CSVファイルのパス
        
    Returns:
        作成された農場のパス
    """
    farm_path = Path("C:/FARMS/DemoFarm2")
    
    # 既存チェック（上書き禁止）
    if farm_path.exists() and (farm_path / "farm.db").exists():
        print(f"デモ農場2は既に存在します: {farm_path}")
        print("既存の農場を上書きする場合は、手動で削除してください。")
        return farm_path
    
    # 農場フォルダを作成
    farm_path.mkdir(parents=True, exist_ok=True)
    print(f"デモ農場2フォルダを作成: {farm_path}")
    
    # データベースを作成
    db_path = farm_path / "farm.db"
    db = DBHandler(db_path)
    print(f"データベースを作成: {db_path}")
    
    # 設定ファイルを作成
    settings = SettingsManager(farm_path)
    settings.set("farm_name", "DemoFarm2")
    print(f"設定ファイルを作成: {farm_path / 'farm_settings.json'}")
    
    # event_dictionary.json をコピー
    source_dict = APP_ROOT / "docs" / "event_dictionary.json"
    if source_dict.exists():
        shutil.copy2(source_dict, farm_path / "event_dictionary.json")
        print(f"event_dictionary.json をコピー")
    else:
        print(f"警告: {source_dict} が見つかりません")
    
    # item_dictionary.json をコピー
    source_item_dict = APP_ROOT / "docs" / "item_dictionary.json"
    if source_item_dict.exists():
        shutil.copy2(source_item_dict, farm_path / "item_dictionary.json")
        print(f"item_dictionary.json をコピー")
    else:
        print(f"警告: {source_item_dict} が見つかりません")
    
    # CSVを読み込む
    print(f"\nCSVファイルを読み込み中: {csv_path}")
    test_date, data_rows = parse_csv(csv_path)
    print(f"検定日: {test_date}")
    print(f"データ行数: {len(data_rows)}")
    
    # RuleEngineを初期化
    rule_engine = RuleEngine(db)
    
    # 各データ行を処理
    created_cows = 0
    created_calv_events = 0
    created_milk_events = 0
    
    for row_data in data_rows:
        cow_id = row_data['cow_id']
        jpn10 = row_data['jpn10']
        clvd = row_data.get('clvd')  # 分娩月日
        lact = row_data.get('lact', 0)  # 産次（baseline）
        
        # 既存の牛をチェック
        existing_cow = db.get_cow_by_id(cow_id)
        
        if existing_cow:
            cow_auto_id = existing_cow['auto_id']
            print(f"既存の牛を使用: cow_id={cow_id}, auto_id={cow_auto_id}")
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
                'pen': 'Lactating',
                'frm': 'DemoFarm2'
            }
            
            try:
                cow_auto_id = db.insert_cow(cow_data)
                created_cows += 1
                print(f"牛を作成: cow_id={cow_id}, jpn10={jpn10}, lact={lact}, auto_id={cow_auto_id}")
            except Exception as e:
                print(f"エラー: 牛の作成に失敗 (cow_id={cow_id}): {e}")
                continue
        
        # 分娩イベント（CALV）を作成（baseline_calvingフラグ付き）
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
                
                try:
                    calv_event_id = db.insert_event(calv_event_data)
                    created_calv_events += 1
                    print(f"分娩イベントを作成: cow_id={cow_id}, event_id={calv_event_id}, date={clvd}")
                except Exception as e:
                    print(f"エラー: 分娩イベントの作成に失敗 (cow_id={cow_id}): {e}")
        
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
            
            try:
                milk_event_id = db.insert_event(milk_event_data)
                created_milk_events += 1
                print(f"乳検イベントを作成: cow_id={cow_id}, event_id={milk_event_id}, date={test_date}")
            except Exception as e:
                print(f"エラー: 乳検イベントの作成に失敗 (cow_id={cow_id}): {e}")
        
        # RuleEngine.apply_events() を呼ぶ（状態を再計算）
        try:
            rule_engine.apply_events(cow_auto_id)
            # 状態を更新
            rule_engine._recalculate_and_update_cow(cow_auto_id)
        except Exception as e:
            print(f"警告: 状態の再計算に失敗 (cow_id={cow_id}): {e}")
    
    db.close()
    
    print(f"\n完了:")
    print(f"  作成された牛: {created_cows}頭")
    print(f"  作成された分娩イベント: {created_calv_events}件")
    print(f"  作成された乳検イベント: {created_milk_events}件")
    print(f"  農場パス: {farm_path}")
    
    return farm_path


def main():
    """メイン処理"""
    import argparse
    
    parser = argparse.ArgumentParser(description='乳検CSVからデモ農場2を作成')
    parser.add_argument('csv_path', type=str, help='CSVファイルのパス')
    
    args = parser.parse_args()
    
    csv_path = Path(args.csv_path)
    
    if not csv_path.exists():
        print(f"エラー: CSVファイルが見つかりません: {csv_path}")
        return 1
    
    try:
        create_demo_farm2(csv_path)
        return 0
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

