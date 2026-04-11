"""
導入登録の共通処理（手動の導入ウィンドウと CSV 一括導入で共有）
"""

from typing import Any, Dict

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


def register_introduction_cow(
    db: DBHandler,
    rule_engine: RuleEngine,
    cow_data: Dict[str, Any],
    introduction_date: str,
    registration_date: str,
    intro_note: str = "外部からの導入",
) -> None:
    """
    1頭分の導入を DB に保存する（IntroductionWindow._do_save と同一ロジック）

    Args:
        intro_note: 導入イベントの note（一括取り込み時は「CSV一括導入」など）
    """
    existing_cow = db.get_cow_by_id(cow_data["cow_id"])
    if existing_cow and existing_cow.get("jpn10") == cow_data["jpn10"]:
        cow_auto_id = existing_cow.get("auto_id")
    else:
        new_cow_data = {
            "cow_id": cow_data["cow_id"],
            "jpn10": cow_data["jpn10"],
            "brd": cow_data.get("breed"),
            "bthd": cow_data.get("birth_date"),
            "entr": registration_date,
            "lact": cow_data.get("lact", 0),
            "clvd": cow_data.get("clvd"),
            "rc": cow_data.get("rc", RuleEngine.RC_OPEN),
            "pen": cow_data.get("pen"),
            "frm": None,
        }
        cow_auto_id = db.insert_cow(new_cow_data)

    intro_json: Dict[str, Any] = {
        "birth_date": cow_data.get("birth_date"),
        "lactation": cow_data.get("lact", 0),
        "calving_date": cow_data.get("clvd"),
        "last_ai_date": cow_data.get("last_ai_date"),
        "reproduction_code": cow_data.get("rc"),
        "pen": cow_data.get("pen"),
        "source": "manual_external",
    }
    if cow_data.get("tag"):
        intro_json["tag"] = cow_data["tag"]
    intro_event = {
        "cow_auto_id": cow_auto_id,
        "event_number": RuleEngine.EVENT_IN,
        "event_date": introduction_date,
        "json_data": intro_json,
        "note": intro_note,
    }
    event_id = db.insert_event(intro_event)
    rule_engine.on_event_added(event_id)

    clvd = cow_data.get("clvd")
    if clvd:
        calv_event_id = db.insert_event(
            {
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_CALV,
                "event_date": clvd,
                "json_data": {"baseline_calving": True},
                "note": "導入時の分娩（baseline）",
            }
        )
        rule_engine.on_event_added(calv_event_id)

    last_ai_date = cow_data.get("last_ai_date")
    if last_ai_date:
        ai_json: Dict[str, Any] = {"ai_count": cow_data.get("ai_count", 0)}
        if cow_data.get("last_ai_sire"):
            ai_json["sire"] = cow_data["last_ai_sire"]
        ai_event_id = db.insert_event(
            {
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_AI,
                "event_date": last_ai_date,
                "json_data": ai_json,
                "note": "導入時の最終授精",
            }
        )
        rule_engine.on_event_added(ai_event_id)

    pregnant_status = cow_data.get("pregnant_status", "")

    if pregnant_status in ("pregnant", "dry"):
        preg_event_id = db.insert_event(
            {
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_PDP2,
                "event_date": introduction_date,
                "json_data": {"twin": False},
                "note": "導入時の妊娠確認",
            }
        )
        rule_engine.on_event_added(preg_event_id)

    if pregnant_status == "dry":
        dry_event_id = db.insert_event(
            {
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_DRY,
                "event_date": introduction_date,
                "json_data": {},
                "note": "導入時の乾乳",
            }
        )
        rule_engine.on_event_added(dry_event_id)

    ai_count = cow_data.get("ai_count", 0)
    if ai_count > 0 and cow_auto_id is not None:
        db.set_item_value(int(cow_auto_id), "BRED", ai_count)

    if cow_auto_id is not None:
        db.set_item_value(int(cow_auto_id), "TAG", (cow_data.get("tag") or "").strip())
