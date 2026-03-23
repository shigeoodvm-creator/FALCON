"""
牛群動態予測レポート：分娩予定月×予定産子種類の集計
分娩予定が決まっている個体について、受胎SIREから産子種類（乳用種メス/乳用種/F1/黒毛和種）を判定し、
分娩予定月別・産子種類別の頭数を集計する。
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from collections import defaultdict
from typing import Dict, Any, List, Optional

from db.db_handler import DBHandler
from modules.sire_list_opts import (
    sire_opts_to_type,
    SIRE_TYPE_F1,
    SIRE_TYPE_HOLSTEIN_FEMALE,
    SIRE_TYPE_HOLSTEIN_REGULAR,
    SIRE_TYPE_UNKNOWN_OTHER,
    SIRE_TYPE_WAGYU,
)

logger = logging.getLogger(__name__)

# 産子種類の表示順序
CALF_TYPE_ORDER = ["乳用種メス", "乳用種", "F1", "黒毛和種", "不明"]


def _classify_calf_type(sire_name: Optional[str], sire_list: Dict[str, Any]) -> str:
    """
    SIRE名とsire_listから産子種類を判定する（sire_list の種別キー／従来の f1・kurowa・female に対応）。
    """
    if not sire_name or not sire_name.strip():
        return "不明"
    sire_name = sire_name.strip()
    opts = sire_list.get(sire_name)
    if not isinstance(opts, dict):
        return "不明"
    st = sire_opts_to_type(opts)
    if st == SIRE_TYPE_UNKNOWN_OTHER:
        return "不明"
    if st == SIRE_TYPE_WAGYU:
        return "黒毛和種"
    if st == SIRE_TYPE_F1:
        return "F1"
    if st == SIRE_TYPE_HOLSTEIN_FEMALE:
        return "乳用種メス"
    if st == SIRE_TYPE_HOLSTEIN_REGULAR:
        return "乳用種"
    return "不明"


def build_herd_dynamics_data(
    db: DBHandler,
    formula_engine: Any,
    farm_path: Path,
) -> Dict[str, Any]:
    """
    分娩予定月×予定産子種類の集計データを構築する。

    Returns:
        {
            "table_rows": [ {"due_ym": "2026-03", "乳用種メス": n, ... }, ... ],
            "months": ["2026-01", ...],
            "calf_types": CALF_TYPE_ORDER,
            "total_by_month": {"2026-01": n, ...},
            "total_by_type": {"乳用種メス": n, ...},
        }
    """
    sire_list_path = farm_path / "sire_list.json"
    sire_list: Dict[str, Dict[str, bool]] = {}
    if sire_list_path.exists():
        try:
            with open(sire_list_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                sire_list = raw
        except Exception as e:
            logger.warning(f"牛群動態: sire_list読み込みエラー: {e}")

    cows = db.get_all_cows()
    if not cows:
        return {
            "table_rows": [],
            "months": [],
            "calf_types": CALF_TYPE_ORDER,
            "total_by_month": {},
            "total_by_type": {},
            "unknown_details": [],
        }

    count_by_ym_type: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    unknown_details: List[Dict[str, Any]] = []  # 不明の内訳（個体ID・SIRE・分娩予定月）

    for cow in cows:
        cow_auto_id = cow.get("auto_id")
        if not cow_auto_id:
            continue
        try:
            events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
            clvd = cow.get("clvd")
            due_date = formula_engine._calculate_due_date(events, clvd)
            if not due_date:
                continue
            due_ym = due_date[:7]
            sire = formula_engine._get_conception_sire(events, cow)
            calf_type = _classify_calf_type(sire, sire_list)
            count_by_ym_type[due_ym][calf_type] += 1
            if calf_type == "不明":
                unknown_details.append({
                    "cow_id": cow.get("cow_id") or cow.get("jpn10") or str(cow_auto_id),
                    "sire": sire if sire and sire.strip() else "（未登録）",
                    "due_ym": due_ym,
                })
        except Exception as e:
            logger.debug(f"牛群動態: cow_auto_id={cow_auto_id} 集計スキップ: {e}")
            continue

    if not count_by_ym_type:
        return {
            "table_rows": [],
            "months": [],
            "calf_types": CALF_TYPE_ORDER,
            "total_by_month": {},
            "total_by_type": {},
            "unknown_details": [],
        }

    months = sorted(count_by_ym_type.keys())
    total_by_type: Dict[str, int] = defaultdict(int)
    total_by_month: Dict[str, int] = {}
    table_rows: List[Dict[str, Any]] = []

    for due_ym in months:
        row: Dict[str, Any] = {"due_ym": due_ym}
        row_sum = 0
        for ct in CALF_TYPE_ORDER:
            cnt = count_by_ym_type[due_ym].get(ct, 0)
            row[ct] = cnt
            row_sum += cnt
            total_by_type[ct] += cnt
        row["合計"] = row_sum
        total_by_month[due_ym] = row_sum
        table_rows.append(row)

    return {
        "table_rows": table_rows,
        "months": months,
        "calf_types": CALF_TYPE_ORDER,
        "total_by_month": dict(total_by_month),
        "total_by_type": dict(total_by_type),
        "unknown_details": unknown_details,
    }
