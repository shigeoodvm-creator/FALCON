"""
デモファームの確認スクリプト
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "app"))

from app.db.db_handler import DBHandler
from app.modules.rule_engine import RuleEngine

farm_path = Path("C:/FARMS/デモファーム")
db_path = farm_path / "farm.db"

if not db_path.exists():
    print(f"エラー: データベースが見つかりません: {db_path}")
    sys.exit(1)

db = DBHandler(db_path)
cows = db.get_all_cows()

print(f"個体数: {len(cows)}")

if cows:
    # 最初の3頭を確認
    for i, cow in enumerate(cows[:3]):
        cow_id = cow.get('cow_id')
        cow_auto_id = cow.get('auto_id')
        events = db.get_events_by_cow(cow_auto_id)
        
        calv_events = [e for e in events if e.get('event_number') == RuleEngine.EVENT_CALV]
        milk_events = [e for e in events if e.get('event_number') == RuleEngine.EVENT_MILK_TEST]
        
        print(f"\n個体 {i+1} (cow_id={cow_id}):")
        print(f"  産次: {cow.get('lact')}")
        print(f"  分娩イベント数: {len(calv_events)}")
        if calv_events:
            for calv in calv_events:
                json_data = calv.get('json_data', {})
                baseline = json_data.get('baseline_calving', False)
                print(f"    分娩日: {calv.get('event_date')}, baseline_calving={baseline}")
        print(f"  乳検イベント数: {len(milk_events)}")
        if milk_events:
            for milk in milk_events[:1]:
                json_data = milk.get('json_data', {})
                print(f"    検定日: {milk.get('event_date')}")
                print(f"    乳量: {json_data.get('milk_yield')}kg")

db.close()









































