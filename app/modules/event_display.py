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


def _get_display_width(text: str) -> int:
    """
    表示幅を計算（全角文字=2、半角文字=1としてカウント）
    
    Args:
        text: テキスト
    
    Returns:
        表示幅（全角換算）
    """
    width = 0
    for char in text:
        # 全角文字（日本語、全角記号など）は2、半角文字は1
        if ord(char) > 0x7F or char in '　':
            width += 2
        else:
            width += 1
    return width


def _pad_right(text: str, width: int) -> str:
    """
    テキストを指定幅まで右側に全角スペースでパディング
    
    Args:
        text: テキスト
        width: 目標幅（全角換算）
    
    Returns:
        パディングされたテキスト
    """
    current_width = _get_display_width(text)
    if current_width >= width:
        return text
    # 不足分を全角スペースで埋める
    padding = width - current_width
    return text + "　" * (padding // 2) + " " * (padding % 2)


def _pad_left(text: str, width: int) -> str:
    """
    テキストを指定幅まで左側に全角スペースでパディング
    
    Args:
        text: テキスト
        width: 目標幅（全角換算）
    
    Returns:
        パディングされたテキスト
    """
    current_width = _get_display_width(text)
    if current_width >= width:
        return text
    # 不足分を全角スペースで埋める
    padding = width - current_width
    return "　" * (padding // 2) + " " * (padding % 2) + text


def format_insemination_event(
    json_data: Optional[Dict[str, Any]],
    technicians: Dict[Union[str, int], Any],
    insemination_types: Dict[Union[str, int], Any],
    conception_status: Optional[str] = None,
    cow_id: Optional[str] = None,
    insemination_count: Optional[int] = None
) -> Optional[str]:
    """
    AI/ETイベントの表示文字列を生成（整列表示）
    
    表示形式: 各フィールドを固定幅でパディングして整列
    例: 3H61714　　　Yasuda　　　　3　CIDR P　　　　　P
    
    フィールド幅:
    - SIRE: 12文字幅（左揃え）
    - 授精師: 10文字幅（左揃え）
    - カウント: 3文字幅（右揃え）
    - 種類: 12文字幅（左揃え）
    - 結果: 2文字幅（左揃え）
    
    Args:
        json_data: イベントの json_data (dict)
        technicians: technician_code -> technician_name の辞書（授精師の表示に使用）
        insemination_types: insemination_type_code -> type_name の辞書（授精種類の表示に使用）
        conception_status: 受胎ステータス（P/R/O/R/N）、Noneの場合はjson_dataから取得
        cow_id: 個体ID（使用しない、後方互換性のため残す）
        insemination_count: 授精カウント（例: 1）
    
    Returns:
        表示文字列（整列済み）、データがない場合は None
    """
    if not json_data:
        return None
    
    # 種雄牛名（SIRE）を取得
    sire = json_data.get("sire") or ""
    
    # 授精師を取得
    tech_code = json_data.get("technician_code") or json_data.get("technician")
    tech_name = _get_name(technicians, tech_code) if tech_code else None
    
    # 授精カウントを取得（引数またはjson_dataから）
    count = insemination_count
    if count is None:
        count = json_data.get("insemination_count")
        if count is not None:
            try:
                count = int(count)
            except (ValueError, TypeError):
                count = None
    
    # 授精種類を取得
    type_code = json_data.get("insemination_type_code") or json_data.get("ai_type") or json_data.get("type")
    type_name = _get_name(insemination_types, type_code) if type_code else None
    
    # 受胎ステータスを取得（引数またはjson_dataから）
    outcome = conception_status
    if outcome is None:
        outcome = json_data.get("outcome")
    
    # 各フィールドを固定幅でパディング
    parts = []
    
    # SIRE（種雄牛名）- 12文字幅、左揃え
    if sire:
        parts.append(_pad_right(str(sire), 12))
    else:
        parts.append("　" * 6)  # 12文字分の全角スペース
    
    # 授精師 - 10文字幅、左揃え
    if tech_name:
        parts.append(_pad_right(str(tech_name), 10))
    else:
        parts.append("　" * 5)  # 10文字分の全角スペース
    
    # 授精カウント - 3文字幅、右揃え
    if count is not None:
        parts.append(_pad_left(str(count), 3))
    else:
        parts.append("　" * 1 + " ")  # 3文字分のスペース
    
    # 授精種類 - 12文字幅、左揃え
    if type_name:
        parts.append(_pad_right(str(type_name), 12))
    else:
        parts.append("　" * 6)  # 12文字分の全角スペース
    
    # 結果（P/O/R/N）- 2文字幅、左揃え
    if outcome:
        parts.append(_pad_right(str(outcome), 2))
    else:
        parts.append("　")  # 2文字分の全角スペース
    
    # 全角スペース1つで区切る（固定幅パディングにより整列される）
    if parts:
        result = "　".join(parts)
        return result
    
    return None


def format_reproduction_check_event(
    json_data: Optional[Dict[str, Any]]
) -> Optional[str]:
    """
    繁殖検査/フレッシュチェック/妊娠鑑定マイナスイベントの表示文字列を生成
    
    処置＝農場設定＞繁殖処置設定で設定する項目（WPG, CIDR, GN, PG, E2 等）。
    所見＝子宮・右・左・その他（子宮OK, NS 等は所見であり処置ではない）。
    表示形式: 処置を固定幅6文字で左揃え、以降は所見を詰めて表示。
    処置なしの場合は6文字分を空欄で開けて、所見のみ表示。
    例: WPG　　右CL　左F
    例: CIDR　　右CL
    例: 　　　　 左CL（処置なし・所見のみ）
    
    Args:
        json_data: イベントの json_data (dict)
    
    Returns:
        表示文字列、データがない場合は None
    """
    if not json_data:
        return None
    
    parts = []
    
    # 処置（新しいキー名と古いキー名の両方をサポート）- 固定幅6文字
    treatment = json_data.get('treatment') or json_data.get('treatment_code', '')
    if treatment and str(treatment).strip() and str(treatment).strip() != '-':
        parts.append(_pad_right(str(treatment).strip(), 6))
    else:
        # 処置なしの場合は6文字分の全角スペースを追加（所見を表示するため）
        parts.append("　" * 3)  # 6文字分の全角スペース
    
    # 以降は詰めて表示（空のフィールドは追加しない）
    # 子宮所見（新しいキー名と古いキー名の両方をサポート）
    uterine = (json_data.get('uterine_findings') or 
              json_data.get('uterus_findings') or 
              json_data.get('uterus_finding') or 
              json_data.get('uterus', ''))
    if uterine and str(uterine).strip() and str(uterine).strip() != '-':
        parts.append(f"子宮{str(uterine).strip()}")
    
    # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
    left_ovary = (json_data.get('left_ovary_findings') or 
                 json_data.get('leftovary_findings') or 
                 json_data.get('leftovary_finding') or 
                 json_data.get('left_ovary', '') or
                 json_data.get('leftovary', ''))
    if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
        parts.append(f"左{str(left_ovary).strip()}")
    
    # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
    right_ovary = (json_data.get('right_ovary_findings') or 
                  json_data.get('rightovary_findings') or 
                  json_data.get('rightovary_finding') or 
                  json_data.get('right_ovary', '') or
                  json_data.get('rightovary', ''))
    if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
        parts.append(f"右{str(right_ovary).strip()}")
    
    # remark（新しいキー名と古いキー名の両方をサポート）
    remark = json_data.get('remark')
    if not remark:
        remark = (json_data.get('other') or 
                 json_data.get('other_info') or 
                 json_data.get('other_findings', ''))
    if remark and str(remark).strip() and str(remark).strip() != '-':
        parts.append(str(remark).strip())
    
    # 全角スペース1つで区切る（処置は固定幅、以降は詰めて表示）
    if parts:
        result = "　".join(parts)
        return result
    
    return None


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
        su = str(sex).upper() if sex not in (None, "") else ""
        if su == "F":
            sex_mark = "♀"
        elif su == "M":
            sex_mark = "♂"
        elif su in ("U", "?"):
            sex_mark = "不明"
        else:
            sex_mark = ""
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


def build_ai_et_event_note(
    event: Dict[str, Any],
    technicians_dict: Dict[Union[str, int], Any],
    insemination_types_dict: Dict[Union[str, int], Any],
    formula_engine: Optional[Any] = None,
    db_handler: Optional[Any] = None
) -> str:
    """
    AI/ETイベントのNOTE文字列を生成（イベント詳細ウィンドウと同じロジック）
    
    この関数は、イベント詳細ウィンドウとCowCardイベント履歴の両方で使用される
    共通のNOTE生成ロジックです。
    
    Args:
        event: イベントデータ（event_number, json_data, id, cow_auto_id を含む）
        technicians_dict: technician_code -> technician_name の辞書
        insemination_types_dict: insemination_type_code -> type_name の辞書
        formula_engine: FormulaEngine インスタンス（受胎ステータス計算用、オプション）
        db_handler: DBHandler インスタンス（受胎ステータス計算用、オプション）
    
    Returns:
        NOTE表示文字列（整列済み）、データがない場合は空文字列
    """
    from modules.rule_engine import RuleEngine
    
    event_number = event.get('event_number')
    json_data = event.get('json_data', {})
    
    # AI/ETイベント以外の場合は空文字列を返す
    if event_number not in [RuleEngine.EVENT_AI, RuleEngine.EVENT_ET]:
        return ''
    
    # json_dataが空の場合は空文字列を返す
    if not json_data:
        return ''
    
    # 受胎ステータスを計算（formula_engineとdb_handlerが提供されている場合）
    conception_status = None
    if formula_engine and db_handler:
        try:
            event_id = event.get('id')
            cow_auto_id = event.get('cow_auto_id')
            if event_id and cow_auto_id:
                cow = db_handler.get_cow_by_auto_id(cow_auto_id)
                if cow:
                    all_events = db_handler.get_events_by_cow(cow_auto_id, include_deleted=False)
                    status_dict = formula_engine.get_ai_conception_status(all_events, cow)
                    conception_status = status_dict.get(event_id)
        except Exception as e:
            logging.debug(f"build_ai_et_event_note: 受胎ステータス計算エラー: {e}")
    
    # json_dataからoutcomeを取得（P/O/R/N）
    outcome = json_data.get('outcome')
    # 既存のconception_statusがある場合は優先、なければoutcomeを使用
    display_status = conception_status or outcome
    
    # 授精カウントを取得
    insemination_count = json_data.get('insemination_count')
    if insemination_count is not None:
        try:
            insemination_count = int(insemination_count)
        except (ValueError, TypeError):
            insemination_count = None
    
    # format_insemination_event関数を使用して表示文字列を生成
    note = format_insemination_event(
        json_data,
        technicians_dict,
        insemination_types_dict,
        display_status,
        cow_id=None,
        insemination_count=insemination_count
    )
    
    # format_insemination_eventがNoneを返した場合は空文字列
    if note is None:
        return ''
    
    # 先頭の全角スペースのみ削除して左揃え（内容は維持）
    # format_insemination_eventは常に何らかの文字列を返す（全角スペースだけでも）
    if note:
        # 先頭の全角スペースだけを削除（内容が全角スペースだけの場合は削除しない）
        original_note = note
        # 全角スペース以外の文字が含まれているか確認
        stripped = note.replace('　', '').replace(' ', '').strip()
        has_content = bool(stripped)
        
        if has_content:
            # 内容がある場合は先頭の全角スペースを削除して左揃え
            while note.startswith('　'):
                note = note[1:]
            note = note.strip()
        else:
            # 全角スペースだけの場合は先頭の全角スペースを削除して左揃え
            note = original_note.lstrip('　').strip()
            # すべて削除された場合は元に戻す
            if not note:
                note = original_note
    
    return note

