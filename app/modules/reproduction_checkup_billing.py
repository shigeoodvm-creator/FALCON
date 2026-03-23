"""
FALCON2 - 繁殖検診請求HTML生成
指定日の繁殖イベント（フレッシュチェック・繁殖検査・妊娠鑑定等）を集計し、請求HTMLを出力する。
"""

import html
import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional
from datetime import datetime

logger = logging.getLogger(__name__)

# 対象イベント番号: フレッシュチェック(300), 繁殖検査(301), 妊娠鑑定マイナス(302), 妊娠鑑定プラス(303), 妊娠鑑定プラス直近以外(304)
REPRO_CHECKUP_EVENT_NUMBERS = (300, 301, 302, 303, 304)

# 分娩・AI・ET（DIM・受胎日数計算用）
EVENT_CALV = 202
EVENT_AI = 200
EVENT_ET = 201
PREGNANCY_PLUS_EVENT_NUMBERS = (303, 304)  # 妊娠鑑定プラス（受胎日数表示対象）

# イベント番号 → 表示名（通常）
REPRO_EVENT_NAME: Dict[int, str] = {
    300: "フレッシュチェック",
    301: "繁殖検査",
    302: "妊娠鑑定マイナス",
    303: "妊娠鑑定プラス",
    304: "妊娠鑑定プラス（直近以外）",
}

# イベント番号 → 縦印刷用短縮表示（2段にならないようコンパクトに）
REPRO_EVENT_NAME_SHORT: Dict[int, str] = {
    300: "ﾌﾚﾁｪｯｸ",           # フレッシュチェック
    301: "繁検",               # 繁殖検査
    302: "妊鑑－",             # 妊娠鑑定マイナス
    303: "妊鑑＋",             # 妊娠鑑定プラス
    304: "妊鑑＋(他)",         # 妊娠鑑定プラス（直近以外）
}


def _generate_note_from_json_data(json_data: Dict[str, Any]) -> str:
    """
    json_dataからNOTE文字列を生成（cow_card.pyの_display_eventsと同じロジック）
    
    Args:
        json_data: イベントのjson_data
    
    Returns:
        生成されたNOTE文字列
    """
    if not json_data:
        return ""
    
    if isinstance(json_data, str):
        try:
            import json as json_module
            json_data = json_module.loads(json_data)
        except:
            return ""
    
    # 有効な値のみをリスト化（空文字、None、"-"は除外）
    valid_parts = []
    
    # 子宮所見（新しいキー名と古いキー名の両方をサポート）
    uterine = (json_data.get('uterine_findings') or 
              json_data.get('uterus_findings') or 
              json_data.get('uterus_finding') or 
              json_data.get('uterus', ''))
    if uterine and str(uterine).strip() and str(uterine).strip() != '-':
        valid_parts.append(f"子宮{str(uterine).strip()}")
    
    # 左卵巣所見（新しいキー名と古いキー名の両方をサポート）
    left_ovary = (json_data.get('left_ovary_findings') or 
                 json_data.get('leftovary_findings') or 
                 json_data.get('leftovary_finding') or 
                 json_data.get('left_ovary', '') or
                 json_data.get('leftovary', ''))
    if left_ovary and str(left_ovary).strip() and str(left_ovary).strip() != '-':
        valid_parts.append(f"左{str(left_ovary).strip()}")
    
    # 右卵巣所見（新しいキー名と古いキー名の両方をサポート）
    right_ovary = (json_data.get('right_ovary_findings') or 
                  json_data.get('rightovary_findings') or 
                  json_data.get('rightovary_finding') or 
                  json_data.get('right_ovary', '') or
                  json_data.get('rightovary', ''))
    if right_ovary and str(right_ovary).strip() and str(right_ovary).strip() != '-':
        valid_parts.append(f"右{str(right_ovary).strip()}")
    
    # remark（新しいキー名と古いキー名の両方をサポート）
    remark = json_data.get('remark')
    if not remark:
        remark = (json_data.get('other') or
                 json_data.get('other_note') or
                 json_data.get('remarks', ''))
    if remark and str(remark).strip() and str(remark).strip() != '-':
        valid_parts.append(str(remark).strip())
    
    if valid_parts:
        return "  ".join(valid_parts)
    
    return ""


def _parse_date(s: str) -> Optional[datetime]:
    """YYYY-MM-DD を datetime に変換。失敗時は None。"""
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def _days_between(start_date_str: str, end_date_str: str) -> Optional[int]:
    """start_date から end_date までの日数（end - start）。基準は end_date（検診日）。"""
    start = _parse_date(start_date_str)
    end = _parse_date(end_date_str)
    if start is None or end is None:
        return None
    delta = end - start
    return delta.days


def _days_since_conception(checkup_date: str, ai_et_date: str, event_number: Optional[int]) -> Optional[int]:
    """
    妊娠鑑定プラス時点での受胎日数を算出。
    - AI(200): 受胎日＝AI日（当日） → 受胎日数 = 検診日 - AI日
    - ET(201): 受胎日＝ET日-7日（胚齢7日、設計書10.3.6） → 受胎日数 = 検診日 - (ET日-7) = (検診日 - ET日) + 7
    """
    if not ai_et_date or not checkup_date:
        return None
    base = _days_between(ai_et_date, checkup_date)
    if base is None:
        return None
    if event_number == EVENT_ET:
        return base + 7
    # AI または未判定時は当日を受胎日とする
    return base


def _last_calving_and_ai_et(events: List[Dict[str, Any]], on_or_before: str) -> Tuple[Optional[str], Optional[str], Optional[int]]:
    """
    イベントリストから「on_or_before 以前の直近の分娩日」と「on_or_before 以前の直近のAI/ET」を返す。
    events は event_date DESC 順を想定（get_events_by_cow の戻り値）。
    Returns:
        (last_calving_date, last_ai_et_date, last_ai_et_event_number)
        event_number は受胎日算出に必要（AI=当日、ET=移植日-7日）。
    """
    last_calving: Optional[str] = None
    last_ai_et_date: Optional[str] = None
    last_ai_et_num: Optional[int] = None
    for ev in events:
        d = (ev.get("event_date") or "").strip()[:10]
        if not d or d > on_or_before:
            continue
        num = ev.get("event_number")
        if num == EVENT_CALV and last_calving is None:
            last_calving = d
        if num in (EVENT_AI, EVENT_ET) and last_ai_et_date is None:
            last_ai_et_date = d
            last_ai_et_num = num
        if last_calving is not None and last_ai_et_date is not None:
            break
    return (last_calving, last_ai_et_date, last_ai_et_num)


def _load_reproduction_treatment_settings(farm_path: Path) -> Dict[str, Dict[str, Any]]:
    """reproduction_treatment_settings.json をロード"""
    path = farm_path / "reproduction_treatment_settings.json"
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        treatments = data.get("treatments", {})
        return {str(k): v for k, v in treatments.items()}
    except Exception as e:
        logger.warning(f"繁殖処置設定の読み込みに失敗: {e}")
        return {}


def _load_drug_settings(farm_path: Path) -> Dict[str, Dict[str, Any]]:
    """drug_settings.json の処置料をロード"""
    path = farm_path / "drug_settings.json"
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("treatment_fees", {})
    except Exception as e:
        logger.warning(f"薬品設定の読み込みに失敗: {e}")
        return {}


def _unit_price_for_treatment_code(
    code: str,
    treatments_config: Dict[str, Dict[str, Any]],
    treatment_fees: Dict[str, Dict[str, Any]],
) -> int:
    """処置コードに対する単価を計算（繁殖処置設定のチェック済み処置料の合計）"""
    treatment = treatments_config.get(code, {})
    fee_keys = treatment.get("treatments", [])
    total = 0
    for key in fee_keys:
        fee_data = treatment_fees.get(key, {})
        fee = fee_data.get("fee")
        if fee is not None:
            try:
                total += int(fee)
            except (TypeError, ValueError):
                pass
    return total


def _ultrasound_only_fee(treatment_fees: Dict[str, Dict[str, Any]]) -> int:
    """検査のみ（処置なし）の単価＝超音波検査のみ"""
    data = treatment_fees.get("ultrasound", {})
    fee = data.get("fee")
    if fee is not None:
        try:
            return int(fee)
        except (TypeError, ValueError):
            pass
    return 0


def build_billing_html(
    db,
    farm_path: Path,
    checkup_date: str,
    farm_name: str,
    output_date: Optional[str] = None,
) -> str:
    """
    請求HTMLを組み立てる。

    Args:
        db: DBHandler インスタンス
        farm_path: 農場フォルダパス
        checkup_date: 繁殖検診実施日（YYYY-MM-DD）
        farm_name: 農場名
        output_date: 出力日（省略時は今日）

    Returns:
        HTML文字列
    """
    if output_date is None:
        output_date = datetime.now().strftime("%Y-%m-%d")

    # 処置設定を先にロード（行のソート・処置表示に必要）
    treatments_config = _load_reproduction_treatment_settings(farm_path)
    treatment_fees = _load_drug_settings(farm_path)

    # 処置文字列 → 設定上のコード（code or name で一致）
    def resolve_treatment_code(raw: str) -> str:
        if not raw:
            return "__inspection_only__"
        raw = raw.strip()
        for code, t in treatments_config.items():
            if code == raw:
                return code
            if (t.get("name") or "").strip() == raw:
                return code
        return raw  # 未登録はそのまま（その他として表示）

    # 指定日の対象イベントを取得
    all_events: List[Dict[str, Any]] = []
    for ev_num in REPRO_CHECKUP_EVENT_NUMBERS:
        events = db.get_events_by_number_and_period(
            ev_num, checkup_date, checkup_date, include_deleted=False
        )
        all_events.extend(events)

    # id でソート（元の登録順に近い）
    all_events.sort(key=lambda e: (e.get("id") or 0))

    # 行データ: (cow_id_4, jpn10, lact, dim, event_name, note, treatment, event_num, treatment_code_resolved, days_since_conception)
    def _cow_id_4(cow: Optional[Dict]) -> str:
        if not cow:
            return ""
        cid = cow.get("cow_id") or ""
        s = str(cid).strip()
        if len(s) >= 4:
            return s[-4:]  # 下4桁（例: 1589111521 → 1152）
        return s

    # 牛ごとのイベントキャッシュ（DIM・受胎日数計算用）
    cow_events_cache: Dict[int, List[Dict[str, Any]]] = {}

    rows: List[Tuple] = []
    # 検診対象の牛ごとの産次（LACT）。同一牛は1回だけカウントする
    cow_lact: Dict[int, int] = {}
    for ev in all_events:
        cow_auto_id = ev.get("cow_auto_id") or 0
        cow = db.get_cow_by_auto_id(cow_auto_id) if cow_auto_id else None
        if cow_auto_id and cow_auto_id not in cow_lact:
            lact = cow.get("lact") if cow else None
            try:
                lact = int(lact) if lact is not None else 0
            except (TypeError, ValueError):
                lact = 0
            cow_lact[cow_auto_id] = lact

        # DIM・受胎日数用にその牛のイベントを取得（検診日基準）
        if cow_auto_id not in cow_events_cache:
            cow_events_cache[cow_auto_id] = db.get_events_by_cow(cow_auto_id, include_deleted=False)
        last_calving, last_ai_et_date, last_ai_et_num = _last_calving_and_ai_et(cow_events_cache[cow_auto_id], checkup_date)
        dim_val: Optional[int] = _days_between(last_calving or "", checkup_date) if last_calving else None
        dim_str = str(dim_val) if dim_val is not None else ""
        lact_val = cow_lact.get(cow_auto_id, 0)
        lact_str = str(lact_val) if lact_val is not None else ""

        # 妊娠鑑定プラスの場合、その時点での受胎日数（AI=検診日-AI日、ET=検診日-(ET日-7)）
        days_since_conception: Optional[int] = None
        event_num = ev.get("event_number")
        if event_num in PREGNANCY_PLUS_EVENT_NUMBERS and last_ai_et_date:
            days_since_conception = _days_since_conception(checkup_date, last_ai_et_date, last_ai_et_num)

        cow_id_4 = _cow_id_4(cow)
        jpn10 = (cow.get("jpn10") or "").strip() if cow else ""
        event_name = REPRO_EVENT_NAME_SHORT.get(event_num) or REPRO_EVENT_NAME.get(event_num, str(event_num)) if event_num is not None else ""
        jdata = ev.get("json_data") or {}
        # NOTEはjson_dataから動的に生成（cow_card.pyと同じロジック）
        note = _generate_note_from_json_data(jdata)
        # json_dataから生成できない場合は既存のnoteフィールドを使用
        if not note:
            note = (ev.get("note") or "").strip()
        treatment = (jdata.get("treatment") or jdata.get("treatment_code") or "").strip()
        if isinstance(treatment, (int, float)):
            treatment = str(treatment).strip()
        treatment_code_resolved = resolve_treatment_code(treatment)
        rows.append((cow_id_4, jpn10, lact_str, dim_str, event_name, note, treatment, event_num, treatment_code_resolved, days_since_conception))

    # 総頭数・経産牛頭数（LACT>=1）・育成牛頭数（LACT=0）
    total_count = len(cow_lact)
    parous_count = sum(1 for l in cow_lact.values() if l >= 1)
    heifer_count = sum(1 for l in cow_lact.values() if l == 0)

    # 処置ごとの集計（行は (cow_id_4, jpn10, lact_str, dim_str, event_name, note, treatment, event_num, treatment_code_resolved, days_since_conception)）

    # 処置コード → 単価
    code_to_unit_price: Dict[str, int] = {}
    for code in treatments_config:
        code_to_unit_price[code] = _unit_price_for_treatment_code(
            code, treatments_config, treatment_fees
        )

    # 処置ごとの頭数（設定コードまたは __inspection_only__ またはその他文字列）
    count_by_code: Dict[str, int] = {}
    for row in rows:
        code = row[8]  # treatment_code_resolved
        count_by_code[code] = count_by_code.get(code, 0) + 1

    # 処置の表示順: 妊娠鑑定プラスを最初にまとめ、以後は繁殖処置設定の順、最後に無処置
    _treatment_order_list = list(treatments_config.keys()) + ["__other__"] + ["__inspection_only__"]

    def _treatment_sort_index(code: str) -> int:
        if code in _treatment_order_list:
            return _treatment_order_list.index(code)
        return _treatment_order_list.index("__other__")

    # 行のソート: (1) 妊娠鑑定プラス(303,304)を先頭、(2) 処置の設定順、(3) ID
    def _row_sort_key(row: Tuple) -> Tuple:
        event_num = row[7]
        is_pregnancy_plus = 0 if event_num in PREGNANCY_PLUS_EVENT_NUMBERS else 1
        code = row[8]
        return (is_pregnancy_plus, _treatment_sort_index(code), row[0])

    rows = sorted(rows, key=_row_sort_key)

    ultrasound_fee = _ultrasound_only_fee(treatment_fees)
    total_amount = 0

    # 処置コード→表示名（設定の name、なければ code）
    def label_for_code(c: str) -> str:
        if c == "__inspection_only__":
            return "検査のみ"
        t = treatments_config.get(c, {})
        return (t.get("name") or c).strip() or c

    # 処置別サマリ行（表示用）
    summary_rows: List[Tuple[str, str, int, int, int]] = []  # (code_key, label, count, unit, subtotal)
    # 登録されている処置コード順
    for code in sorted(treatments_config.keys(), key=lambda c: (c == "", c)):
        cnt = count_by_code.get(code, 0)
        if cnt == 0:
            continue
        unit = code_to_unit_price.get(code, 0)
        subtotal = cnt * unit
        total_amount += subtotal
        summary_rows.append((code, label_for_code(code), cnt, unit, subtotal))

    # 設定にない処置文字列（その他）
    for raw_code, cnt in count_by_code.items():
        if raw_code == "__inspection_only__":
            continue
        if raw_code in treatments_config:
            continue
        unit = 0
        subtotal = 0
        total_amount += subtotal
        summary_rows.append((raw_code, raw_code or "（未設定）", cnt, unit, subtotal))

    # 検査のみ（処置空欄）
    inspection_only_count = count_by_code.get("__inspection_only__", 0)
    if inspection_only_count > 0:
        summary_rows.append(
            ("__inspection_only__", "検査のみ", inspection_only_count, ultrasound_fee, inspection_only_count * ultrasound_fee)
        )
        total_amount += inspection_only_count * ultrasound_fee

    # HTML生成
    title = "繁殖検診請求"
    css = """
    body { font-family: "Meiryo UI", sans-serif; background: #fff; margin: 24px; color: #263238; }
    .container { max-width: 900px; margin: 0 auto; background: #fff; padding: 32px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    h1 { font-size: 22px; margin: 0 0 20px 0; color: #3949ab; border-bottom: 2px solid #3949ab; padding-bottom: 8px; }
    .meta { margin-bottom: 24px; color: #607d8b; font-size: 14px; }
    .meta p { margin: 6px 0; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 24px; font-size: 14px; }
    th, td { border: 1px solid #b0bec5; padding: 10px 12px; text-align: left; }
    .event-list-table td.note-cell { min-width: 16em; }
    .event-list-table td.treatment-cell { min-width: 8em; }
    .event-list-table td.event-cell { white-space: nowrap; width: 6em; }
    th { background: #eceff1; font-weight: bold; color: #37474f; }
    tr:nth-child(even) { background: #fafafa; }
    .summary-table { margin-top: 24px; }
    .summary-table td { border: 1px solid #b0bec5; padding: 8px 12px; }
    .total-row { font-weight: bold; background: #e8eaf6 !important; }
    .inspection-only { color: #546e7a; }
    """

    parts = [
        "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>",
        html.escape(title),
        "</title><style>",
        css,
        "</style></head><body><div class='container'>",
        "<h1>",
        html.escape(title),
        "</h1>",
        "<div class='meta'>",
        "<p><strong>繁殖検診実施日</strong>：",
        html.escape(checkup_date),
        "</p>",
        "<p><strong>出力日</strong>：",
        html.escape(output_date),
        "</p>",
        "<p><strong>農場名</strong>：",
        html.escape(farm_name),
        "</p>",
        "<p><strong>総頭数</strong>：",
        str(total_count),
        "頭（経産牛：",
        str(parous_count),
        "頭、育成牛：",
        str(heifer_count),
        "頭）</p>",
        "</div>",
    ]

    # 処置コード→表示名（設定の name、なければ code）。イベント一覧の処置列表示用
    def treatment_display(raw: str) -> str:
        if not raw or not str(raw).strip():
            return "検査のみ"
        code = resolve_treatment_code(raw)
        return label_for_code(code)

    # 妊娠鑑定プラス時のNOTE表示（受胎日数をNOTE内に記載）
    def note_display(row: Tuple) -> str:
        note = row[5]
        event_num = row[7]
        days = row[9]  # days_since_conception
        if event_num in PREGNANCY_PLUS_EVENT_NUMBERS and days is not None:
            days_part = f"{days}日"
            return f"{days_part}  {note}" if note else days_part
        return note or ""

    # イベント一覧表（ID・JPN10・産次・DIM・イベント名・処置・NOTE）。処置ごとにまとめて表示
    parts.append(
        "<table class='event-list-table'><colgroup><col style='width:4em'><col style='width:8em'>"
        "<col style='width:3em'><col style='width:4em'><col style='width:6em'><col style='width:10em'><col></colgroup>"
        "<thead><tr><th>ID</th><th>JPN10</th><th>産次</th><th>DIM</th><th>ｲﾍﾞﾝﾄ</th><th>処置</th><th>NOTE</th></tr></thead><tbody>"
    )
    for row in rows:
        cow_id_4, jpn10, lact_str, dim_str, event_name, note, treatment = row[0], row[1], row[2], row[3], row[4], row[5], row[6]
        treatment_label = treatment_display(treatment)
        display_note = note_display(row)
        parts.append("<tr>")
        parts.append("<td>")
        parts.append(html.escape(str(cow_id_4)))
        parts.append("</td><td>")
        parts.append(html.escape(jpn10))
        parts.append("</td><td>")
        parts.append(html.escape(lact_str))
        parts.append("</td><td>")
        parts.append(html.escape(dim_str))
        parts.append("</td><td class='event-cell'>")
        parts.append(html.escape(event_name))
        parts.append("</td><td class='treatment-cell'>")
        parts.append(html.escape(treatment_label))
        parts.append("</td><td class='note-cell'>")
        parts.append(html.escape(display_note))
        parts.append("</td></tr>")
    parts.append("</tbody></table>")

    # 処置ごと集計表
    parts.append("<table class='summary-table'><thead><tr><th>処置</th><th>頭数</th><th>単価（円）</th><th>本日合計（円）</th></tr></thead><tbody>")
    for code_key, label, cnt, unit, subtotal in summary_rows:
        row_class = " class='inspection-only'" if code_key == "__inspection_only__" else ""
        parts.append(f"<tr{row_class}>")
        parts.append("<td>")
        parts.append(html.escape(label))
        parts.append("</td><td>")
        parts.append(str(cnt))
        parts.append("頭</td><td>")
        parts.append(f"{unit:,}")
        parts.append("</td><td>")
        parts.append(f"{subtotal:,}")
        parts.append("</td></tr>")
    parts.append("<tr class='total-row'><td colspan='3'>本日総額</td><td>")
    parts.append(f"{total_amount:,}")
    parts.append("円</td></tr>")
    parts.append("</tbody></table>")

    parts.append("</div></body></html>")
    return "".join(parts)
