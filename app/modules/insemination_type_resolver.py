"""
FALCON2 - 授精種類の自動判定（吸い込み・イベント入力で共通利用）
当該AIイベントの直近の繁殖処置イベントと「〇日後AI」設定から授精種類コードを算出する。
"""

import json
import logging
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List, Tuple

logger = logging.getLogger(__name__)

# 授精種類判定で「直近の繁殖関連」に使うイベント（AI, ET, フレッシュ, 繁殖検査, 妊鑑-）
_LATEST_REPRO_FOR_INSEM_TYPE = (200, 201, 300, 301, 302)


def _load_insemination_types(farm_path: Path) -> Dict[str, str]:
    """insemination_settings.json から授精種類を読み込む"""
    path = farm_path / "insemination_settings.json"
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("insemination_types", {})
    except Exception as e:
        logger.warning(f"insemination_settings 読み込みエラー: {e}")
        return {}


def _load_treatments(farm_path: Path) -> Dict[str, Any]:
    """reproduction_treatment_settings.json から処置設定を読み込む"""
    path = farm_path / "reproduction_treatment_settings.json"
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        raw = data.get("treatments", {})
        return {str(k): v for k, v in raw.items()}
    except Exception as e:
        logger.warning(f"reproduction_treatment_settings 読み込みエラー: {e}")
        return {}


def _get_treatment_ai_days_map(treatments: Dict[str, Any]) -> Dict[str, int]:
    """処置名 → AIまでの日数 のマップを構築"""
    result: Dict[str, int] = {}
    for code, treatment in treatments.items():
        name = (treatment.get("name") or "").strip()
        if not name:
            continue
        protocols = treatment.get("protocols_by_position") or treatment.get("protocols") or []
        if not isinstance(protocols, list):
            continue
        for p in protocols:
            if not p or not isinstance(p, dict):
                continue
            if str(p.get("instruction", "")).strip().upper() == "AI":
                days = p.get("days")
                if days is not None and str(days).strip() != "":
                    try:
                        result[name.upper()] = int(days)
                    except (ValueError, TypeError):
                        pass
                break
    return result


def _get_most_recent_reproduction_or_ai_et_before_date(
    events: List[Dict[str, Any]], before_date: str
) -> Optional[Dict[str, Any]]:
    """AI日付より前の直近の繁殖関連イベントを取得（AI, ET, フレッシュ, 繁殖検査, 妊鑑- のいずれか。event_date DESC 順を想定）"""
    try:
        before_dt = datetime.strptime(before_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        return None
    for ev in events:
        ev_date = ev.get("event_date", "")
        if not ev_date:
            continue
        try:
            ev_dt = datetime.strptime(ev_date, "%Y-%m-%d")
        except (ValueError, TypeError):
            continue
        if ev_dt >= before_dt:
            continue
        if ev.get("event_number") in _LATEST_REPRO_FOR_INSEM_TYPE:
            return ev
    return None


def _get_treatment_name_from_event(ev: Dict[str, Any], treatments: Dict[str, Any]) -> Optional[str]:
    """イベントの json_data から処置名を取得（コードの場合は設定から表示名に変換）"""
    json_data = ev.get("json_data") or {}
    treatment = json_data.get("treatment") or json_data.get("treatment_code", "")
    treatment = str(treatment).strip() if treatment else ""
    if not treatment or treatment == "-":
        return None
    if treatment.isdigit():
        t = treatments.get(str(treatment), {})
        if isinstance(t, dict):
            return (t.get("name") or treatment).strip() or None
    return treatment.strip() or None


def _find_natural_estrus_code(insemination_types: Dict[str, str]) -> Optional[str]:
    """授精種類設定から自然発情に該当するコードを返す（A を優先、なければ「自然発情」を含むもの）"""
    fallback = None
    for code, name in insemination_types.items():
        code_str = str(code).strip().upper()
        name_str = str(name).strip()
        if "自然発情" in name_str:
            if fallback is None:
                fallback = str(code)
            if code_str == "A":
                return str(code)
    return fallback


def compute_suggested_insemination_type_code(
    db,  # DBHandler
    farm_path: Path,
    cow_auto_id: int,
    ai_date: str,
) -> Optional[str]:
    """
    AI/ET日付入力時点で、直近の繁殖関連イベントから授精種類コードを推定する。
    1) 直近の繁殖関連（AI, ET, フレッシュ, 繁殖検査, 妊鑑-）を1件確定
    2) 直近がAI/ET → 自然発情
    3) 直近が繁殖検査/妊鑑-で処置が授精種類とリンクしていない（例: GN）→ 自然発情
    4) 直近が繁殖検査/妊鑑-で処置が授精種類とリンクし、〇日後AIと日付が±1日で一致 → その授精種類
    5) 上記以外（〇日後AIと一致しない等）→ 空欄（None）
    """
    insemination_types = _load_insemination_types(farm_path)
    treatments = _load_treatments(farm_path)
    natural_code = _find_natural_estrus_code(insemination_types)
    ai_days_map = _get_treatment_ai_days_map(treatments)

    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    latest_repro = _get_most_recent_reproduction_or_ai_et_before_date(events, ai_date)
    if not latest_repro:
        return natural_code

    # 3) 直近がAI/ET → 自然発情
    if latest_repro.get("event_number") in (200, 201):
        return natural_code

    # 直近がフレッシュ(300)/繁殖検査(301)/妊鑑-(302)の場合
    if latest_repro.get("event_number") not in (300, 301, 302):
        return natural_code

    treatment_name = _get_treatment_name_from_event(latest_repro, treatments)
    treatment_date = latest_repro.get("event_date", "")
    if not treatment_date:
        return natural_code

    # 処置なし → 自然発情（処置が授精種類とリンクしていないとみなす）
    if not treatment_name:
        return natural_code

    # 処置が授精種類とリンクしているか（処置名が授精種類名と一致するか）
    treatment_upper = treatment_name.strip().upper()
    matched_code = None
    for code, name in insemination_types.items():
        name_upper = (str(name).strip()).upper()
        if treatment_upper == name_upper or treatment_upper in name_upper:
            matched_code = str(code)
            break
    if matched_code is None:
        # 4) リンクしていない（例: GN）→ 自然発情
        return natural_code

    # 〇日後AIが設定されているか
    ai_days = ai_days_map.get(treatment_upper)
    if ai_days is None:
        # 6) 〇日後AIなし → 空欄
        return None

    # 5) 入力日付が 処置日 + 〇日 ±1 の範囲か
    try:
        td = datetime.strptime(treatment_date, "%Y-%m-%d")
        day_low = (td + timedelta(days=ai_days - 1)).strftime("%Y-%m-%d")
        day_high = (td + timedelta(days=ai_days + 1)).strftime("%Y-%m-%d")
        if day_low <= ai_date <= day_high:
            return matched_code
    except (ValueError, TypeError):
        pass
    # 6) 日付が一致しない → 空欄
    return None
