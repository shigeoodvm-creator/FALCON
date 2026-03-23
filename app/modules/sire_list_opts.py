"""
sire_list.json の SIRE 種別定義（UI・胎子計算・牛群動態で共通）
"""

from __future__ import annotations

from typing import Any, Dict, List

SIRE_TYPE_HOLSTEIN_REGULAR = "holstein_regular"
SIRE_TYPE_HOLSTEIN_FEMALE = "holstein_female"
SIRE_TYPE_F1 = "f1"
SIRE_TYPE_WAGYU = "wagyu"
SIRE_TYPE_UNKNOWN_OTHER = "unknown_other"

SIRE_TYPE_ORDER: List[str] = [
    SIRE_TYPE_HOLSTEIN_REGULAR,
    SIRE_TYPE_HOLSTEIN_FEMALE,
    SIRE_TYPE_F1,
    SIRE_TYPE_WAGYU,
    SIRE_TYPE_UNKNOWN_OTHER,
]

SIRE_TYPE_LABELS_JA: Dict[str, str] = {
    SIRE_TYPE_HOLSTEIN_REGULAR: "乳用種レギュラー",
    SIRE_TYPE_HOLSTEIN_FEMALE: "乳用種♀",
    SIRE_TYPE_F1: "F1",
    SIRE_TYPE_WAGYU: "黒毛和種",
    SIRE_TYPE_UNKNOWN_OTHER: "不明、その他",
}


def sire_opts_to_type(opts: Any) -> str:
    """JSON の1エントリから種別キーを得る（sire_type 優先、無ければ従来の f1/kurowa/female から推定）。"""
    if not isinstance(opts, dict):
        return SIRE_TYPE_HOLSTEIN_REGULAR
    st = opts.get("sire_type")
    if isinstance(st, str) and st in SIRE_TYPE_ORDER:
        return st
    f1 = bool(opts.get("f1", False))
    kurowa = bool(opts.get("kurowa", False))
    female = bool(opts.get("female", False))
    if kurowa:
        return SIRE_TYPE_WAGYU
    if f1:
        return SIRE_TYPE_F1
    if female:
        return SIRE_TYPE_HOLSTEIN_FEMALE
    return SIRE_TYPE_HOLSTEIN_REGULAR


# 旧UIラベル（ホルスタイン表記）→ 現行キー（保存データはキー基準のため通常は不要）
_LEGACY_JA_TO_TYPE: Dict[str, str] = {
    "ホルスタインレギュラー": SIRE_TYPE_HOLSTEIN_REGULAR,
    "ホルスタイン♀": SIRE_TYPE_HOLSTEIN_FEMALE,
}


def ja_label_to_sire_type(label: str) -> str:
    """Combobox 等の日本語表示から種別キーへ。"""
    legacy = _LEGACY_JA_TO_TYPE.get(label)
    if legacy is not None:
        return legacy
    for key, ja in SIRE_TYPE_LABELS_JA.items():
        if ja == label:
            return key
    return SIRE_TYPE_HOLSTEIN_REGULAR


def sire_type_to_stored_dict(sire_type: str) -> Dict[str, Any]:
    """保存用 dict。sire_type と従来の f1/kurowa/female を併記（後方互換・参照用）。"""
    if sire_type not in SIRE_TYPE_ORDER:
        sire_type = SIRE_TYPE_HOLSTEIN_REGULAR
    f1 = sire_type == SIRE_TYPE_F1
    kurowa = sire_type == SIRE_TYPE_WAGYU
    female = sire_type == SIRE_TYPE_HOLSTEIN_FEMALE
    return {
        "sire_type": sire_type,
        "f1": f1,
        "kurowa": kurowa,
        "female": female,
    }
