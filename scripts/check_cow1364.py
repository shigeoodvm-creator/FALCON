"""個体1364のデータを確認するスクリプト"""
import sys
from pathlib import Path

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from datetime import datetime

# すべての農場をチェック
farms_root = Path("C:/FARMS")
found_cow = None
found_farm = None

for farm_dir in farms_root.iterdir():
    if not farm_dir.is_dir():
        continue
    db_path = farm_dir / "farm.db"
    if not db_path.exists():
        continue
    
    print(f"チェック中: {farm_dir.name}")
    db = DBHandler(db_path)
    cow = db.get_cow_by_id("1364")
    if cow:
        found_cow = cow
        found_farm = farm_dir.name
        print(f"  個体1364を発見: {farm_dir.name}")
        break

if not found_cow:
    print("エラー: 個体1364が見つかりません")
    sys.exit(1)

print(f"\n農場: {found_farm}")
db = DBHandler(Path(farms_root / found_farm / "farm.db"))
cow = found_cow
auto_id = cow.get('auto_id')

print("=" * 60)
print("個体1364のデータ:")
print("=" * 60)
print(f"auto_id: {cow.get('auto_id')}")
print(f"cow_id: {cow.get('cow_id')}")
print(f"jpn10: {cow.get('jpn10')}")
print(f"clvd (分娩日): {cow.get('clvd')}")
print(f"lact (産次): {cow.get('lact')}")
print(f"bthd (生年月日): {cow.get('bthd')}")
print(f"rc (繁殖コード): {cow.get('rc')}")

# DIMを計算
if cow.get('clvd'):
    clvd_date = datetime.strptime(cow.get('clvd'), '%Y-%m-%d')
    today = datetime.now()
    dim = (today - clvd_date).days
    print(f"\nDIM計算:")
    print(f"  分娩日: {cow.get('clvd')}")
    print(f"  現在日: {today.strftime('%Y-%m-%d')}")
    print(f"  DIM: {dim} 日")

# イベント履歴を取得
events = db.get_events_by_cow(auto_id, include_deleted=False)
print(f"\nイベント履歴 (削除済みを除く): {len(events)}件")
print("-" * 60)
for i, e in enumerate(events[:20], 1):
    event_date = e.get('event_date', '')
    event_number = e.get('event_number', '')
    note = e.get('note', '')
    deleted = e.get('deleted', 0)
    print(f"{i:2d}. {event_date} - イベント番号:{event_number} - {note} (deleted={deleted})")

if len(events) > 20:
    print(f"... 他 {len(events) - 20} 件")

# 分娩イベントを確認
rule_engine = RuleEngine(db)
calv_events = [e for e in events if e.get('event_number') == rule_engine.EVENT_CALV]
print(f"\n分娩イベント (event_number={rule_engine.EVENT_CALV}): {len(calv_events)}件")
for e in calv_events:
    print(f"  {e.get('event_date')} - {e.get('note', '')}")

# RuleEngineで状態を再計算
print(f"\nRuleEngineで状態を再計算:")
state = rule_engine.apply_events(auto_id)
print(f"  clvd: {state.get('clvd')}")
print(f"  lact: {state.get('lact')}")
print(f"  rc: {state.get('rc')}")
