"""
ローカル操作ログ（イベント登録・更新・削除の履歴）
分析UI「操作ログ」タブおよびトラブル調査用。config/activity_log.json に保存。
"""

from __future__ import annotations

import json
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from constants import APP_CONFIG_DIR

logger = logging.getLogger(__name__)

_MAX_ENTRIES = 500
_LOG_FILE = APP_CONFIG_DIR / "activity_log.json"
_lock = threading.Lock()


def _farm_label_from_db(db: Any) -> str:
    """DB パスから農場フォルダ名を取得（表示用）"""
    try:
        p = getattr(db, "db_path", None)
        if p:
            return Path(p).parent.name
    except Exception:
        pass
    return "不明"


def _event_label(event_number: Optional[int], event_dictionary: Optional[Dict[str, Any]]) -> str:
    if event_number is None:
        return "—"
    s = str(event_number)
    if event_dictionary and s in event_dictionary:
        name = event_dictionary[s].get("name_jp") or event_dictionary[s].get("label")
        if name:
            return f"{name}({s})"
    return s


def append_entry(
    *,
    farm: str,
    cow_id: str,
    cow_auto_id: Optional[int],
    action: str,
    event_label: str,
    event_date: str = "",
    event_id: Optional[int] = None,
) -> None:
    """
    1件追加（失敗しても例外は握りつぶす）

    action: 登録 / 更新 / 削除 など
    """
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row: Dict[str, Any] = {
            "ts": ts,
            "farm": farm or "不明",
            "cow_id": cow_id or "",
            "cow_auto_id": cow_auto_id,
            "action": action,
            "event": event_label or "",
            "event_date": (event_date or "")[:10],
        }
        if event_id is not None:
            row["event_id"] = event_id

        with _lock:
            APP_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            entries: List[Dict[str, Any]] = []
            if _LOG_FILE.exists():
                try:
                    with open(_LOG_FILE, "r", encoding="utf-8") as f:
                        entries = json.load(f)
                    if not isinstance(entries, list):
                        entries = []
                except Exception as e:
                    logger.warning("activity_log: 読み込み失敗、新規にします: %s", e)
                    entries = []
            entries.insert(0, row)
            entries = entries[:_MAX_ENTRIES]
            tmp = _LOG_FILE.with_suffix(".json.tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(entries, f, ensure_ascii=False, indent=2)
            tmp.replace(_LOG_FILE)
    except Exception as e:
        logger.warning("activity_log: 追記失敗: %s", e)


def load_entries(limit: int = 200) -> List[Dict[str, Any]]:
    """新しい順で最大 limit 件"""
    try:
        if not _LOG_FILE.exists():
            return []
        with _lock:
            with open(_LOG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        if not isinstance(data, list):
            return []
        return data[:limit]
    except Exception as e:
        logger.warning("activity_log: 読み込み失敗: %s", e)
        return []


def record_from_event(
    db: Any,
    *,
    cow_auto_id: Optional[int],
    action: str,
    event_number: Optional[int],
    event_id: Optional[int],
    event_date: str = "",
    event_dictionary: Optional[Dict[str, Any]] = None,
) -> None:
    """DB から個体IDを解決して append_entry する共通ヘルパー"""
    cow_id = ""
    if cow_auto_id is not None:
        try:
            cow = db.get_cow_by_auto_id(cow_auto_id)
            if cow:
                cow_id = str(cow.get("cow_id") or "")
        except Exception:
            pass
    farm = _farm_label_from_db(db)
    label = _event_label(event_number, event_dictionary)
    append_entry(
        farm=farm,
        cow_id=cow_id,
        cow_auto_id=cow_auto_id,
        action=action,
        event_label=label,
        event_date=event_date or "",
        event_id=event_id,
    )
