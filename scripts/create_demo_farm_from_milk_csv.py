"""
FALCON2 - 乳検CSVからデモ農場を作成
CSVファイルを読み込んで、DemoFarmを作成し、牛と乳検イベントを登録する
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
        keywords: 検索キーワードのリスト
        
    Returns:
        見つかった列のインデックス、見つからない場合はNone
    """
    for idx, header in enumerate(headers):
        if pd.notna(header):
            header_str = str(header).strip()
            for keyword in keywords:
                if keyword in header_str:
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
    # Shift-JISエンコーディングで手動で読み込む
    # CSVファイルの構造が複雑なため、行ごとに読み込む
    lines = []
    with open(csv_path, 'r', encoding='shift_jis') as f:
        for line in f:
            lines.append(line.strip().split(','))
    
    # DataFrameに変換
    # 最大列数を取得
    max_cols = max(len(row) for row in lines) if lines else 0
    # すべての行を同じ列数に揃える
    for row in lines:
        while len(row) < max_cols:
            row.append('')
    
    df = pd.DataFrame(lines, dtype=str)
    
    # 検定日を取得（3行目、0列目）
    test_date = None
    print(f"DataFrame shape: {df.shape}")
    print(f"最初の5行:")
    for i in range(min(5, len(df))):
        print(f"  行{i}: {df.iloc[i, 0] if len(df.columns) > 0 else 'N/A'}")
    
    if len(df) > 3:
        date_str = df.iloc[3, 0] if len(df.columns) > 0 else None
        if pd.notna(date_str):
            date_str = str(date_str).strip()
            print(f"検定日候補 (行3, 列0): '{date_str}'")
            # 2025/11/05 形式を 2025-11-05 に変換
            try:
                date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                test_date = date_obj.strftime('%Y-%m-%d')
                print(f"検定日を取得: {test_date}")
            except Exception as e:
                print(f"日付パースエラー: {e}")
                # 他の形式も試す
                for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y%m%d']:
                    try:
                        date_obj = datetime.strptime(date_str, fmt)
                        test_date = date_obj.strftime('%Y-%m-%d')
                        print(f"検定日を取得 (形式 {fmt}): {test_date}")
                        break
                    except:
                        pass
    
    if not test_date:
        # 全行を検索
        print("検定日を全行から検索中...")
        for idx in range(len(df)):
            for col_idx in range(min(3, len(df.columns))):
                val = df.iloc[idx, col_idx] if len(df.columns) > col_idx else None
                if pd.notna(val):
                    val_str = str(val).strip()
                    if '/' in val_str and len(val_str) >= 8:
                        try:
                            date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                            test_date = date_obj.strftime('%Y-%m-%d')
                            print(f"検定日を発見 (行{idx}, 列{col_idx}): {test_date}")
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
    
    # カラムインデックスを検索（当日値のみ使用）
    col_jpn10 = 0  # 個体識別番号
    col_cow_id = 1  # 拡大管理番号
    col_brd = 2  # 品種（データ行では空の場合が多い）
    
    # ヘッダーからカラム位置を検索（当日値の列）
    # CSV構造: 当日/前日が交互に並ぶため、当日値は奇数インデックス（0始まり）
    col_milk_yield = find_column_index(headers, ['乳量', '当日']) or 3
    col_fat = find_column_index(headers, ['脂肪', '当日']) or 6
    col_protein = find_column_index(headers, ['蛋白', '当日']) or 9
    col_snf = find_column_index(headers, ['無脂', '当日']) or 12
    col_scc = find_column_index(headers, ['体細胞', '当日']) or 15
    col_ls = find_column_index(headers, ['スコア', '当日']) or 19  # 体細胞スコア
    col_mun = find_column_index(headers, ['MUN', '当日']) or 22  # MUN
    col_bhb = find_column_index(headers, ['BHB', '当日']) or 24  # BHB
    col_bthd = find_column_index(headers, ['生年月日', '生年']) or 32  # 生年月日
    col_brd_data = find_column_index(headers, ['品種']) or 33  # 品種（データ行）
    col_clvd = find_column_index(headers, ['分娩', '最終']) or 34  # 最終分娩日
    col_lact = find_column_index(headers, ['産次', '産']) or 35  # 産次
    
    # FA関連（後半の列にある可能性が高い）
    col_denovo_fa = find_column_index(headers, ['denovo', 'FA', '当日'])
    col_preformed_fa = find_column_index(headers, ['preformed', 'FA', '当日'])
    col_mixed_fa = find_column_index(headers, ['mixed', 'FA', '当日'])
    col_denovo_milk = find_column_index(headers, ['denovo', 'Milk', '当日'])
    
    # データ行を取得（8行目以降）
    data_rows = []
    for idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[idx].tolist()
        
        # 空行をスキップ
        if not any(pd.notna(val) and str(val).strip() for val in row[:5]):
            continue
        
        # 個体識別番号（0列目）と拡大4桁（1列目）を取得
        jpn10 = str(row[col_jpn10]).strip() if len(row) > col_jpn10 and pd.notna(row[col_jpn10]) else None
        cow_id_4digit = str(row[col_cow_id]).strip() if len(row) > col_cow_id and pd.notna(row[col_cow_id]) else None
        
        if not jpn10 or len(jpn10) < 4:
            continue
        
        # cow_idを決定（拡大4桁があればそれを使用、なければJPN10の下4桁）
        if cow_id_4digit and len(cow_id_4digit) >= 4:
            cow_id = cow_id_4digit[-4:]  # 最後の4桁
        else:
            cow_id = jpn10[-4:]  # JPN10の下4桁
        
        # 品種（2列目またはデータ行の品種列）
        brd = "ホルスタイン"  # デフォルト
        # まずデータ行の品種列を確認
        if col_brd_data and len(row) > col_brd_data and pd.notna(row[col_brd_data]):
            brd_val = str(row[col_brd_data]).strip()
            if brd_val:
                brd = brd_val
        # なければ2列目を確認
        elif len(row) > col_brd and pd.notna(row[col_brd]):
            brd_val = str(row[col_brd]).strip()
            if brd_val:
                brd = brd_val
        
        # 生年月日
        bthd = None
        if col_bthd and len(row) > col_bthd and pd.notna(row[col_bthd]):
            val_str = str(row[col_bthd]).strip()
            # 日付形式（YYYY/MM/DD）をチェック
            if '/' in val_str and len(val_str) >= 8:
                try:
                    date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                    bthd = date_obj.strftime('%Y-%m-%d')
                except:
                    pass
        
        # 最終分娩日（clvd）も取得（RuleEngineで使用される可能性がある）
        clvd = None
        if col_clvd and len(row) > col_clvd and pd.notna(row[col_clvd]):
            val_str = str(row[col_clvd]).strip()
            # 日付形式（YYYY/MM/DD）をチェック
            if '/' in val_str and len(val_str) >= 8:
                try:
                    date_obj = datetime.strptime(val_str, '%Y/%m/%d')
                    clvd = date_obj.strftime('%Y-%m-%d')
                except:
                    pass
        
        # 産次
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
                    # *や**などのマーカーを除去
                    val_str = val_str.replace('*', '').replace('**', '').replace('***', '')
                    if val_str:
                        return int(float(val_str))
                except:
                    pass
            return None
        
        milk_yield = get_float_value(col_milk_yield)
        fat = get_float_value(col_fat)
        protein = get_float_value(col_protein)
        snf = get_float_value(col_snf)
        scc = get_int_value(col_scc)
        ls = get_float_value(col_ls)
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
            'clvd': clvd,
            'lact': lact,
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


def create_demo_farm(csv_path: Path, overwrite: bool = False) -> Path:
    """
    デモ農場を作成
    
    Args:
        csv_path: CSVファイルのパス
        overwrite: 既存の農場を上書きするか
        
    Returns:
        作成された農場のパス
    """
    farm_path = Path("C:/FARMS/DemoFarm")
    
    # 既存チェック
    if farm_path.exists() and (farm_path / "farm.db").exists():
        if not overwrite:
            print(f"デモ農場は既に存在します: {farm_path}")
            return farm_path
        else:
            print(f"既存のデモ農場を上書きします: {farm_path}")
            shutil.rmtree(farm_path)
    
    # 農場フォルダを作成
    farm_path.mkdir(parents=True, exist_ok=True)
    print(f"デモ農場フォルダを作成: {farm_path}")
    
    # データベースを作成
    db_path = farm_path / "farm.db"
    db = DBHandler(db_path)
    print(f"データベースを作成: {db_path}")
    
    # 設定ファイルを作成
    settings = SettingsManager(farm_path)
    settings.set("farm_name", "DemoFarm")
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
    created_events = 0
    
    for row_data in data_rows:
        cow_id = row_data['cow_id']
        jpn10 = row_data['jpn10']
        
        # 既存の牛をチェック
        existing_cow = db.get_cow_by_id(cow_id)
        
        if existing_cow:
            cow_auto_id = existing_cow['auto_id']
            print(f"既存の牛を使用: cow_id={cow_id}, auto_id={cow_auto_id}")
        else:
            # 新しい牛を作成
            cow_data = {
                'cow_id': cow_id,
                'jpn10': jpn10,
                'brd': row_data.get('brd', 'ホルスタイン'),
                'bthd': row_data.get('bthd'),
                'entr': None,
                'lact': row_data.get('lact', 0) or 0,  # lact_baseline
                'clvd': row_data.get('clvd'),  # CSVに含まれる場合は使用
                'rc': None,  # RuleEngineで後から更新
                'pen': 'Lactating',
                'frm': 'DemoFarm'
            }
            
            try:
                cow_auto_id = db.insert_cow(cow_data)
                created_cows += 1
                print(f"牛を作成: cow_id={cow_id}, jpn10={jpn10}, auto_id={cow_auto_id}")
            except Exception as e:
                print(f"エラー: 牛の作成に失敗 (cow_id={cow_id}): {e}")
                continue
        
        # このCSV由来の乳検イベントが既に存在するかチェック
        events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
        existing_milk_test = False
        for event in events:
            if event.get('event_number') == RuleEngine.EVENT_MILK_TEST:
                event_json = event.get('json_data', {})
                # 同じ検定日のイベントが既に存在するかチェック
                if event.get('event_date') == test_date:
                    existing_milk_test = True
                    break
        
        if existing_milk_test:
            print(f"既存の乳検イベントをスキップ: cow_id={cow_id}, date={test_date}")
            continue
        
        # 乳検イベント（601）を作成
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
        
        event_data = {
            'cow_auto_id': cow_auto_id,
            'event_number': RuleEngine.EVENT_MILK_TEST,
            'event_date': test_date,
            'json_data': json_data if json_data else None,
            'note': None
        }
        
        try:
            event_id = db.insert_event(event_data)
            created_events += 1
            print(f"乳検イベントを作成: cow_id={cow_id}, event_id={event_id}, date={test_date}")
            
            # RuleEngine.apply_events() を呼ぶ
            rule_engine.apply_events(cow_auto_id)
            # 状態を更新（on_event_addedを呼ぶ）
            rule_engine.on_event_added(event_id)
            
        except Exception as e:
            print(f"エラー: イベントの作成に失敗 (cow_id={cow_id}): {e}")
            continue
    
    db.close()
    
    print(f"\n完了:")
    print(f"  作成された牛: {created_cows}頭")
    print(f"  作成されたイベント: {created_events}件")
    print(f"  農場パス: {farm_path}")
    
    return farm_path


def main():
    """メイン処理"""
    import argparse
    
    parser = argparse.ArgumentParser(description='乳検CSVからデモ農場を作成')
    parser.add_argument('csv_path', type=str, help='CSVファイルのパス')
    parser.add_argument('--overwrite', action='store_true', help='既存の農場を上書きする')
    
    args = parser.parse_args()
    
    csv_path = Path(args.csv_path)
    print(f"CSVファイルパス: {csv_path}")
    print(f"絶対パス: {csv_path.resolve()}")
    print(f"存在確認: {csv_path.exists()}")
    
    if not csv_path.exists():
        print(f"エラー: CSVファイルが見つかりません: {csv_path}")
        print(f"絶対パス: {csv_path.resolve()}")
        # 親ディレクトリの内容を確認
        parent = csv_path.parent
        if parent.exists():
            print(f"親ディレクトリの内容:")
            for item in list(parent.iterdir())[:10]:
                print(f"  {item.name}")
        return 1
    
    try:
        create_demo_farm(csv_path, overwrite=args.overwrite)
        return 0
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

