"""
AnalysisUI基本動作テスト（エラー検出用）
"""

import sys
import tkinter as tk
from pathlib import Path

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.query_router_v2 import QueryRouterV2
from modules.executor_v2 import ExecutorV2
from modules.execution_plan import AnalysisType

def test_basic_operations():
    """基本的な操作をテスト"""
    print("=" * 80)
    print("AnalysisUI 基本動作テスト")
    print("=" * 80)
    
    # デモファームのパス
    farm_path = Path("C:/FARMS/デモファーム")
    if not farm_path.exists():
        farm_path = Path("C:/FARMS/DemoFarm")
    
    if not farm_path.exists():
        print(f"エラー: デモファームが見つかりません: {farm_path}")
        return
    
    print(f"農場パス: {farm_path}")
    
    # DBHandlerを初期化
    db_path = farm_path / "farm.db"
    if not db_path.exists():
        print(f"エラー: farm.dbが見つかりません: {db_path}")
        return
    
    db = DBHandler(db_path)
    print("[OK] DBHandler初期化完了")
    
    # FormulaEngineを初期化
    item_dict_path = farm_path / "item_dictionary.json"
    formula_engine = FormulaEngine(db, item_dict_path)
    print("[OK] FormulaEngine初期化完了")
    
    # RuleEngineを初期化
    rule_engine = RuleEngine(db)
    print("[OK] RuleEngine初期化完了")
    
    # QueryRouterV2を初期化
    event_dict_path = farm_path / "event_dictionary.json"
    query_router_v2 = QueryRouterV2(
        item_dictionary_path=item_dict_path,
        event_dictionary_path=event_dict_path
    )
    print("[OK] QueryRouterV2初期化完了")
    
    # ExecutorV2を初期化
    executor_v2 = ExecutorV2(
        db_handler=db,
        formula_engine=formula_engine,
        rule_engine=rule_engine,
        item_dictionary_path=item_dict_path,
        event_dictionary_path=event_dict_path
    )
    print("[OK] ExecutorV2初期化完了")
    
    # テスト1: QueryRouterV2.parse()のテスト
    print("\n[テスト1] QueryRouterV2.parse()のテスト")
    try:
        plan = query_router_v2.parse("list", None, None, "ID 分娩後日数")
        print(f"  analysis_type: {plan.analysis_type}")
        print(f"  columns: {plan.columns}")
        print(f"  errors: {plan.errors}")
        print(f"  warnings: {plan.warnings}")
        print("[OK] QueryRouterV2.parse()成功")
    except Exception as e:
        print(f"[エラー] QueryRouterV2.parse()失敗: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # テスト2: ExecutorV2.execute()のテスト
    print("\n[テスト2] ExecutorV2.execute()のテスト")
    try:
        result = executor_v2.execute(plan)
        print(f"  success: {result.get('success')}")
        print(f"  errors: {result.get('errors')}")
        print(f"  warnings: {result.get('warnings')}")
        if result.get('data'):
            data = result.get('data')
            print(f"  data type: {data.get('type')}")
            if data.get('type') == 'table':
                print(f"  columns: {data.get('columns')}")
                print(f"  rows count: {len(data.get('rows', []))}")
        print("[OK] ExecutorV2.execute()成功")
    except Exception as e:
        print(f"[エラー] ExecutorV2.execute()失敗: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # テスト3: 集計のテスト
    print("\n[テスト3] 集計のテスト")
    try:
        plan = query_router_v2.parse("agg", None, None, "授精回数\n産次")
        print(f"  analysis_type: {plan.analysis_type}")
        print(f"  aggregate_items: {plan.aggregate_items}")
        print(f"  group_by: {plan.group_by}")
        result = executor_v2.execute(plan)
        print(f"  success: {result.get('success')}")
        print(f"  errors: {result.get('errors')}")
        if result.get('data'):
            data = result.get('data')
            print(f"  data type: {data.get('type')}")
            if data.get('type') == 'table':
                print(f"  columns: {data.get('columns')}")
                print(f"  rows count: {len(data.get('rows', []))}")
        print("[OK] 集計テスト成功")
    except Exception as e:
        print(f"[エラー] 集計テスト失敗: {e}")
        import traceback
        traceback.print_exc()
    
    # テスト4: イベント集計のテスト
    print("\n[テスト4] イベント集計のテスト")
    try:
        plan = query_router_v2.parse("eventcount", None, None, "授精\n月")
        print(f"  analysis_type: {plan.analysis_type}")
        print(f"  event_keys: {plan.event_keys}")
        print(f"  group_by: {plan.group_by}")
        result = executor_v2.execute(plan)
        print(f"  success: {result.get('success')}")
        print(f"  errors: {result.get('errors')}")
        if result.get('data'):
            data = result.get('data')
            print(f"  data type: {data.get('type')}")
            if data.get('type') == 'table':
                print(f"  columns: {data.get('columns')}")
                print(f"  rows count: {len(data.get('rows', []))}")
        print("[OK] イベント集計テスト成功")
    except Exception as e:
        print(f"[エラー] イベント集計テスト失敗: {e}")
        import traceback
        traceback.print_exc()
    
    db.close()
    print("\nテスト完了")

if __name__ == "__main__":
    test_basic_operations()

















