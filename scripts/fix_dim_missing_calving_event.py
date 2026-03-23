"""
cow.clvd と分娩イベントの不一致を修正するスクリプト。

cow の分娩月日（clvd）はあるが、その日付の分娩イベント(202)が無い場合、
該当日付で分娩イベントを追加し、RuleEngine で再計算する。

使い方:
  python scripts/fix_dim_missing_calving_event.py --cow-id 1141
  python scripts/fix_dim_missing_calving_event.py --cow-id 1141 --farm デモファーム
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "app"))

from app.constants import FARMS_ROOT
from app.db.db_handler import DBHandler
from app.modules.rule_engine import RuleEngine


def fix_cow_if_needed(db: DBHandler, rule_engine: RuleEngine, cow: dict) -> bool:
    """牛の clvd に対応する分娩イベントが無ければ追加する。修正したら True."""
    cow_auto_id = cow.get("auto_id")
    cow_id = cow.get("cow_id", "")
    clvd = cow.get("clvd") or ""
    if not clvd or not cow_auto_id:
        return False
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    has_calving_on_clvd = any(
        e.get("event_number") == RuleEngine.EVENT_CALV and e.get("event_date") == clvd
        for e in events
    )
    if has_calving_on_clvd:
        return False
    event_data = {
        "cow_auto_id": cow_auto_id,
        "event_number": RuleEngine.EVENT_CALV,
        "event_date": clvd,
        "json_data": {},
        "note": "clvd整合のため追加",
    }
    event_id = db.insert_event(event_data)
    rule_engine.on_event_added(event_id)
    print(f"  追加: cow_id={cow_id}, 分娩日={clvd}, event_id={event_id}")
    return True


def main():
    parser = argparse.ArgumentParser(description="cow.clvd と分娩イベントの不一致を修正")
    parser.add_argument("--cow-id", required=True, help="牛ID（例: 1141）")
    parser.add_argument("--farm", default=None, help="農場名（省略時は全農場を対象）")
    args = parser.parse_args()

    cow_id_arg = args.cow_id.strip()
    if args.farm:
        farms = [FARMS_ROOT / args.farm]
        if not (farms[0] / "farm.db").exists():
            print(f"エラー: 農場が見つかりません: {farms[0]}")
            sys.exit(1)
    else:
        farms = [p for p in FARMS_ROOT.iterdir() if p.is_dir() and (p / "farm.db").exists()]

    if not farms:
        print(f"エラー: 農場がありません: {FARMS_ROOT}")
        sys.exit(1)

    fixed_any = False
    for farm_path in farms:
        db_path = farm_path / "farm.db"
        db = DBHandler(db_path)
        rule_engine = RuleEngine(db)
        try:
            for cid in (cow_id_arg, cow_id_arg.zfill(4)):
                cows = db.get_cows_by_id(cid)
                for cow in cows:
                    if fix_cow_if_needed(db, rule_engine, cow):
                        fixed_any = True
        finally:
            db.close()

    if fixed_any:
        print("修正を反映しました。")
    else:
        print("修正対象の牛はいませんでした（既に分娩イベントがあるか、clvd が未設定です）。")


if __name__ == "__main__":
    main()
