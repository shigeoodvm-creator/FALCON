"""
AnalysisUI動作確認スクリプト
基本的な動作を確認するためのテスト
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
from ui.analysis_ui import AnalysisUI

def test_analysis_ui():
    """AnalysisUIの基本動作をテスト"""
    print("=" * 80)
    print("AnalysisUI 動作確認テスト")
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
    
    # テストウィンドウを作成
    root = tk.Tk()
    root.title("AnalysisUI テスト")
    root.geometry("1200x800")
    
    def on_execute(result):
        """実行結果のコールバック"""
        print("\n" + "=" * 80)
        print("実行結果:")
        print(f"  success: {result.get('success')}")
        if result.get('errors'):
            print(f"  errors: {result.get('errors')}")
        if result.get('warnings'):
            print(f"  warnings: {result.get('warnings')}")
        if result.get('data'):
            data = result.get('data')
            print(f"  data type: {data.get('type')}")
            if data.get('type') == 'table':
                print(f"  columns: {data.get('columns')}")
                print(f"  rows count: {len(data.get('rows', []))}")
        print("=" * 80)
    
    # AnalysisUIを作成
    try:
        analysis_ui = AnalysisUI(
            parent=root,
            farm_path=farm_path,
            query_router=query_router_v2,
            executor=executor_v2,
            on_execute=on_execute
        )
        analysis_ui.frame.pack(fill=tk.BOTH, expand=True)
        print("[OK] AnalysisUI作成完了")
        print("\nテストウィンドウを表示します。")
        print("手動で以下を確認してください:")
        print("  1. 分析種別選択が表示される")
        print("  2. コマンド入力欄が表示される")
        print("  3. 右パネル（参照・挿入）が表示される")
        print("  4. 結果表示エリアが表示される")
        print("\nウィンドウを閉じるとテスト終了です。")
        
        root.mainloop()
        
    except Exception as e:
        print(f"エラー: AnalysisUI作成に失敗: {e}")
        import traceback
        traceback.print_exc()
    
    db.close()
    print("\nテスト完了")

if __name__ == "__main__":
    test_analysis_ui()

