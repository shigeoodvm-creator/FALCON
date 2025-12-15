"""
FALCON2 - イベント表示ユーティリティ
AI/ETイベントの表示文字列を生成する共通関数
"""

from typing import Dict, Any, Optional, Union
import logging


def _resolve_display_name(val: Any) -> Optional[str]:
    """
    表示用の名称を解決（優先順位: name, short, label, display, str(value)）
    
    Args:
        val: 辞書の値（str, dict, その他）
    
    Returns:
        表示用名称（文字列）、見つからない場合は None
    """
    if val is None:
        return None
    
    # 値が文字列の場合
    if isinstance(val, str):
        return val
    
    # 値が dict の場合
    if isinstance(val, dict):
        # 優先順位: name, short, label, display
        for key in ("name", "short", "label", "display"):
            if key in val and val[key]:
                return val[key]
    
    # その他の型の場合は文字列に変換
    return str(val) if val is not None else None


def _get_name(mapping: Dict[Union[str, int], Any], code: Any) -> Optional[str]:
    """
    辞書から名称を取得（辞書構造に依存しない実装）
    
    - code は str に正規化
    - キーが int / str の両方に対応
    - 値は _resolve_display_name で解決
    
    Args:
        mapping: コード辞書（キーは str または int、値は str または dict）
        code: コード値（任意の型）
    
    Returns:
        名称（文字列）、見つからない場合は None
    """
    if code is None:
        return None
    
    code_str = str(code)
    
    # 文字列キーで検索、なければ数値キーで検索
    val = mapping.get(code_str) or mapping.get(code)
    
    # デバッグログ
    if val is None:
        logging.debug(f"_get_name: code={code} (str={code_str}) not found in mapping keys={list(mapping.keys())[:10]}")
    else:
        logging.debug(f"_get_name: code={code} (str={code_str}) found val={val} (type={type(val).__name__})")
    
    return _resolve_display_name(val)


def format_insemination_event(
    json_data: Optional[Dict[str, Any]],
    technicians: Dict[Union[str, int], Any],
    insemination_types: Dict[Union[str, int], Any]
) -> Optional[str]:
    """
    AI/ETイベントの表示文字列を生成
    
    表示形式: SIRE <2スペース> 授精師名 <2スペース> 授精種類名（固定幅フォーマット）
    
    Args:
        json_data: イベントの json_data (dict)
        technicians: technician_code -> technician_name の辞書
                     （例: {"1": "Sonoda"} または {"1": {"name": "Sonoda"}}）
        insemination_types: insemination_type_code -> type_name の辞書
                           （例: {"1": "自然発情"} または {"1": {"name": "自然発情"}}）
    
    Returns:
        表示文字列（例: "6h5555  Sonoda      自然発情"）、
        データがない場合は None
    """
    if not json_data:
        return None
    
    # json_dataからキーを取得（複数のキー名に対応）
    sire = json_data.get("sire") or ""
    tech_code = json_data.get("technician_code") or json_data.get("technician")
    type_code = json_data.get("insemination_type_code") or json_data.get("ai_type") or json_data.get("type")
    
    # 名称を取得（辞書構造に依存しない実装）
    tech_name = _get_name(technicians, tech_code) or ""
    type_name = _get_name(insemination_types, type_code) or ""
    
    # 確認ログ（残す）
    logging.info(
        "AI display debug: sire=%s tech=%s type=%s tech_name=%s type_name=%s",
        sire,
        tech_code,
        type_code,
        tech_name,
        type_name
    )
    
    # 固定幅フォーマットで表示
    # SIRE <2スペース> 授精師 <2スペース> 授精種類
    parts = []
    
    if sire:
        # SIREは最大8文字分の幅を確保（右詰め）
        parts.append(str(sire).ljust(8))
    
    if tech_name:
        # 授精師名は最大12文字分の幅を確保（右詰め）
        parts.append(str(tech_name).ljust(12))
    
    if type_name:
        # 授精種類名は最後なので幅調整不要
        parts.append(str(type_name))
    
    # 半角スペース2つで区切る
    result = "  ".join(parts).strip()
    
    # NOTEの残り情報がある場合は後ろに続ける
    # 既存のnoteやmemoがあれば追加（ただし、sire/technician/type以外の情報）
    other_info = []
    for key in json_data.keys():
        if key not in ["sire", "technician_code", "technician", "insemination_type_code", "ai_type", "type"]:
            value = json_data.get(key)
            if value and str(value).strip() and str(value).strip() != "-":
                # 既に表示済みの情報は除外
                if str(value) not in [sire, tech_name, type_name, str(tech_code), str(type_code)]:
                    other_info.append(f"{key}:{value}")
    
    if other_info:
        result = f"{result}  {'  '.join(other_info)}"
    
    return result if result else None


def format_calving_event(
    json_data: Optional[Dict[str, Any]],
    difficulty_labels: Optional[Dict[str, str]] = None
) -> Optional[str]:
    """
    分娩イベントの表示文字列を生成
    
    表示形式例:
        難易度:2 / 双子 / ホルスタイン♀・ホルスタイン♂
    
    Args:
        json_data: 分娩イベントの json_data (dict)
        difficulty_labels: 難易度コード -> 名称 の辞書（任意）
    
    Returns:
        表示文字列、データがない場合は None
    """
    if not json_data:
        return None
    
    calves = json_data.get("calves") or []
    calv_count = len(calves)
    
    # 頭数ラベル
    count_label = None
    if calv_count == 1:
        count_label = "単子"
    elif calv_count == 2:
        count_label = "双子"
    elif calv_count == 3:
        count_label = "三つ子"
    elif calv_count > 0:
        count_label = f"{calv_count}子"
    
    # 難易度
    diff = json_data.get("calving_difficulty")
    if diff is not None:
        diff_str = str(diff)
        # 数値で渡された場合に備えて辞書を str キーで参照
        label = difficulty_labels.get(diff_str) if difficulty_labels else None
        diff_part = f"難易度:{diff}"
        if label and str(diff) != label:
            diff_part = f"{diff_part}({label})"
    else:
        diff_part = None
    
    # 子牛詳細
    calf_parts = []
    for calf in calves:
        breed = calf.get("breed") or ""
        sex = calf.get("sex")
        sex_mark = "♀" if str(sex).upper() == "F" else ("♂" if sex else "")
        stillborn = calf.get("stillborn", False)
        part = f"{breed}{sex_mark}".strip()
        if stillborn:
            part = f"{part}（死産）"
        if part:
            calf_parts.append(part)
    
    parts = []
    if diff_part:
        parts.append(diff_part)
    if count_label:
        parts.append(count_label)
    if calf_parts:
        parts.append("・".join(calf_parts))
    
    return " / ".join(parts) if parts else None

