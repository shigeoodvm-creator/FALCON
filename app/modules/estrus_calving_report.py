"""
FALCON2 - 発情周期表・発情カレンダー・分娩予定表・分娩予定カレンダー HTML生成モジュール

対象牛: 経産牛（lact >= 1）・在籍中のみ
"""

import json
import html as html_mod
import calendar
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.sire_list_opts import SIRE_TYPE_HOLSTEIN_FEMALE, sire_opts_to_type


# ────────────────────────────────────────────────────────────────────────────
# 共通ユーティリティ
# ────────────────────────────────────────────────────────────────────────────

def _esc(s: Any) -> str:
    return html_mod.escape(str(s)) if s is not None else ""


def _parse_date(s: Optional[str]) -> Optional[date]:
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _fmt_date(d: Optional[date]) -> str:
    return d.strftime("%Y/%m/%d") if d else "—"


def _is_cow_disposed(db: DBHandler, rule_engine: RuleEngine, cow_auto_id: Any) -> bool:
    if cow_auto_id is None:
        return True
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    for ev in events:
        if ev.get("event_number") in [rule_engine.EVENT_SOLD, rule_engine.EVENT_DEAD]:
            return True
    return False


def _get_json(ev: Dict) -> Dict:
    j = ev.get("json_data") or {}
    if isinstance(j, str):
        try:
            j = json.loads(j)
        except Exception:
            j = {}
    return j


def _get_active_parous_cows(
    db: DBHandler, rule_engine: RuleEngine
) -> List[Dict[str, Any]]:
    """在籍中の経産牛と計算済み状態を返す。"""
    all_cows = db.get_all_cows()
    result: List[Dict[str, Any]] = []
    for cow in all_cows:
        lact = cow.get("lact")
        try:
            lact_int = int(lact) if lact is not None else 0
        except (TypeError, ValueError):
            lact_int = 0
        if lact_int < 1:
            continue
        auto_id = cow.get("auto_id")
        if _is_cow_disposed(db, rule_engine, auto_id):
            continue
        try:
            state = rule_engine.apply_events(auto_id)
        except Exception:
            state = {
                "lact": lact_int,
                "clvd": cow.get("clvd"),
                "rc": cow.get("rc") or rule_engine.RC_OPEN,
                "last_ai_date": None,
                "conception_date": None,
                "due_date": None,
            }
        result.append({
            "auto_id": auto_id,
            "cow_id": cow.get("cow_id") or "",
            "jpn10": cow.get("jpn10") or "",
            "lact": state.get("lact", lact_int),
            "clvd": state.get("clvd") or cow.get("clvd"),
            "rc": state.get("rc") or rule_engine.RC_OPEN,
            "last_ai_date": state.get("last_ai_date"),
            "conception_date": state.get("conception_date"),
            "due_date": state.get("due_date"),
            "pen": state.get("pen") or cow.get("pen") or "",
        })
    return result


def _get_cycle_day_info(db: DBHandler, rule_engine: RuleEngine, cow_auto_id: int) -> Optional[Tuple[date, int]]:
    """
    最新の cycle_day が記録された検査イベント（300/301/302）を探し、
    (検査日, cycle_day) を返す。なければ None。
    """
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    candidate: Optional[Tuple[date, int]] = None
    for ev in events:
        if ev.get("event_number") not in (
            rule_engine.EVENT_FCHK,
            rule_engine.EVENT_REPRO,
            rule_engine.EVENT_PDN,
        ):
            continue
        j = _get_json(ev)
        raw = j.get("cycle_day")
        if raw is None:
            continue
        try:
            cd = int(raw)
        except (TypeError, ValueError):
            continue
        if cd < 1 or cd > 21:
            continue
        ev_date = _parse_date(ev.get("event_date"))
        if ev_date is None:
            continue
        if candidate is None or ev_date > candidate[0]:
            candidate = (ev_date, cd)
    return candidate


def _get_conception_ai_event(db: DBHandler, rule_engine: RuleEngine, cow_auto_id: int, conception_date_str: Optional[str]) -> Optional[Dict]:
    """
    受胎AI/ETイベントを返す（conception_date に最も近いもの）。
    """
    if not conception_date_str:
        return None
    conception_date = _parse_date(conception_date_str)
    if conception_date is None:
        return None
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    # AIの場合: AI日 = conception_date、ETの場合: AI日 = conception_date + 7日
    best: Optional[Dict] = None
    best_diff = 999
    for ev in events:
        if ev.get("event_number") not in (rule_engine.EVENT_AI, rule_engine.EVENT_ET):
            continue
        ev_date = _parse_date(ev.get("event_date"))
        if ev_date is None:
            continue
        diff = abs((ev_date - conception_date).days)
        if ev.get("event_number") == rule_engine.EVENT_ET:
            diff = abs((ev_date - timedelta(days=7) - conception_date).days)
        if diff < best_diff:
            best_diff = diff
            best = ev
    return best if best_diff <= 7 else None


def _get_pregnancy_event(db: DBHandler, rule_engine: RuleEngine, cow_auto_id: int) -> Optional[Dict]:
    """最新の妊娠確定イベント（303/307）を返す。"""
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    candidates = [
        ev for ev in events
        if ev.get("event_number") in (rule_engine.EVENT_PDP, rule_engine.EVENT_PAGP,
                                       getattr(rule_engine, "EVENT_PDP2", 304))
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda e: (e.get("event_date") or "", e.get("id") or 0))


def _common_html_head(title: str, extra_style: str = "") -> str:
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>{_esc(title)}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: "Meiryo UI", "Yu Gothic", "MS Gothic", sans-serif;
    background: #f8f9fa;
    color: #263238;
    font-size: 13px;
  }}
  .page-header {{
    background: #3949ab;
    color: #fff;
    padding: 14px 24px 10px;
    display: flex;
    align-items: flex-end;
    gap: 16px;
  }}
  .page-header h1 {{ font-size: 1.3em; font-weight: bold; }}
  .page-header .meta {{ font-size: 0.85em; opacity: 0.85; margin-left: auto; }}
  .container {{ padding: 16px 20px; }}
  table.report-table {{
    border-collapse: collapse;
    width: 100%;
    background: #fff;
    border-radius: 6px;
    overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }}
  table.report-table th {{
    background: #3949ab;
    color: #fff;
    padding: 8px 10px;
    text-align: center;
    font-size: 0.85em;
    white-space: nowrap;
  }}
  table.report-table td {{
    padding: 7px 10px;
    border-bottom: 1px solid #e0e7ef;
    text-align: center;
    vertical-align: middle;
  }}
  table.report-table tr:hover td {{ background: #e8eaf6; }}
  table.report-table tr:last-child td {{ border-bottom: none; }}
  .badge {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 10px;
    font-size: 0.8em;
    font-weight: bold;
    white-space: nowrap;
  }}
  .badge-1st {{ background: #1976d2; color: #fff; }}
  .badge-2nd {{ background: #7b1fa2; color: #fff; }}
  .badge-alert {{ background: #f57c00; color: #fff; }}
  .badge-past {{ background: #b0bec5; color: #fff; font-weight: normal; }}
  .badge-today {{ background: #2e7d32; color: #fff; }}
  .section-title {{
    font-size: 1.05em;
    font-weight: bold;
    color: #3949ab;
    border-left: 4px solid #3949ab;
    padding-left: 10px;
    margin: 20px 0 10px;
  }}
  {extra_style}

  /* ── 印刷共通 ─────────────────────────────────────────── */
  @media print {{
    @page {{ margin: 8mm 6mm; }}
    body {{
      background: #fff;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
      color-adjust: exact;
    }}
    .page-header {{
      background: #3949ab !important;
      color: #fff !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }}
    table.report-table th {{
      background: #3949ab !important;
      color: #fff !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }}
    /* テーブル行がページをまたがないように */
    table.report-table tr {{
      page-break-inside: avoid;
      break-inside: avoid;
    }}
  }}
</style>
</head>
<body>
"""


# ────────────────────────────────────────────────────────────────────────────
# 1. 発情周期表
# ────────────────────────────────────────────────────────────────────────────

def build_estrus_table_html(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
    farm_name: str,
    id_label: str = "ID",
) -> str:
    today = _parse_date(checkup_date) or date.today()

    cows = _get_active_parous_cows(db, rule_engine)

    rows: List[Dict[str, Any]] = []
    for cow in cows:
        rc = cow["rc"]
        # 妊娠中・乾乳中は除外
        if rc in (rule_engine.RC_PREGNANT, rule_engine.RC_DRY):
            continue
        last_ai = _parse_date(cow["last_ai_date"])
        clvd = _parse_date(cow["clvd"])
        dim = (today - clvd).days if clvd else None
        day21 = (last_ai + timedelta(days=21)) if last_ai else None
        day42 = (last_ai + timedelta(days=42)) if last_ai else None

        # cycle_day 由来の発情注意日を計算
        alert_date: Optional[date] = None
        cd_info = _get_cycle_day_info(db, rule_engine, cow["auto_id"])
        if cd_info:
            exam_date, cd = cd_info
            days_to_estrus = 21 - cd
            alert_date = exam_date + timedelta(days=days_to_estrus)

        rows.append({
            "cow_id": cow["cow_id"],
            "jpn10": cow["jpn10"],
            "lact": cow["lact"],
            "dim": dim,
            "last_ai": last_ai,
            "day21": day21,
            "day42": day42,
            "alert_date": alert_date,
            "sort_key": day21 or date(9999, 12, 31),
        })

    rows.sort(key=lambda r: (r["sort_key"], r["cow_id"]))

    body = _common_html_head(
        f"発情周期表 | {farm_name}",
        extra_style="""
        .dim-cell { font-size: 0.9em; color: #546e7a; }
        .date-past { color: #90a4ae; }
        .alert-row td { background: #fff8e1 !important; }
        """,
    )
    body += f"""
<div class="page-header">
  <div>
    <h1>📋 発情周期表</h1>
    <div style="font-size:0.85em;opacity:0.85">{_esc(farm_name)}</div>
  </div>
  <div class="meta">基準日: {_esc(checkup_date)}&nbsp;&nbsp;出力: {date.today().strftime('%Y/%m/%d')}&nbsp;&nbsp;対象: 経産牛（オープン）</div>
</div>
<div class="container">
<div class="section-title">発情周期表（最終AI日 ± 21・42日）</div>
<table class="report-table">
<thead>
<tr>
  <th>{_esc(id_label)}</th>
  <th>JPN10</th>
  <th>産次</th>
  <th>DIM</th>
  <th>最終AI日</th>
  <th>21日後<br><span style="font-size:0.8em;font-weight:normal">（1周期）</span></th>
  <th>42日後<br><span style="font-size:0.8em;font-weight:normal">（2周期）</span></th>
  <th>発情注意日<br><span style="font-size:0.8em;font-weight:normal">（周期Day記録）</span></th>
</tr>
</thead>
<tbody>
"""

    if not rows:
        body += '<tr><td colspan="8" style="text-align:center;color:#90a4ae;padding:20px">対象牛なし</td></tr>'
    for r in rows:
        def _date_badge(d: Optional[date], cycle_num: int) -> str:
            if d is None:
                return "—"
            is_past = d < today
            badge_cls = "badge-past" if is_past else (f"badge-{cycle_num}{'st' if cycle_num == 1 else 'nd'}")
            return f'<span class="badge {badge_cls}">{_fmt_date(d)}</span>'

        alert_str = "—"
        if r["alert_date"]:
            is_past = r["alert_date"] < today
            cls = "badge-past" if is_past else "badge-alert"
            alert_str = f'<span class="badge {cls}">{_fmt_date(r["alert_date"])}</span>'

        row_cls = ""
        if r["alert_date"] and not r["alert_date"] < today and abs((r["alert_date"] - today).days) <= 2:
            row_cls = ' class="alert-row"'

        body += f"""<tr{row_cls}>
  <td><b>{_esc(r["cow_id"])}</b></td>
  <td style="font-size:0.85em">{_esc(r["jpn10"])}</td>
  <td>{_esc(r["lact"])}</td>
  <td class="dim-cell">{r["dim"] if r["dim"] is not None else "—"}</td>
  <td>{_fmt_date(r["last_ai"])}</td>
  <td>{_date_badge(r["day21"], 1)}</td>
  <td>{_date_badge(r["day42"], 2)}</td>
  <td>{alert_str}</td>
</tr>
"""

    legend = """
<div style="margin-top:14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap">
  <span style="font-size:0.85em;color:#546e7a">凡例：</span>
  <span class="badge badge-1st">1周期（21日後）</span>
  <span class="badge badge-2nd">2周期（42日後）</span>
  <span class="badge badge-alert">発情注意（周期Day記録あり）</span>
  <span class="badge badge-past">過去日</span>
</div>
"""
    body += f"</tbody></table>{legend}</div></body></html>"
    return body


# ────────────────────────────────────────────────────────────────────────────
# 週始まり設定ヘルパー
# ────────────────────────────────────────────────────────────────────────────

def _week_config(week_start: str):
    """
    week_start: "sunday" or "monday"
    Returns:
      headers: [(label, css_class), ...] × 7
      col_of(python_weekday) -> column index (0-6)
      is_sun_col(col) -> bool
      is_sat_col(col) -> bool
    """
    if week_start == "sunday":
        # 列0=日(Sun), 1=月, 2=火, ..., 6=土(Sat)
        # Python weekday: Mon=0, Tue=1, ..., Sat=5, Sun=6
        # col = (weekday + 1) % 7  →  Mon→1, ..., Sat→6, Sun→0
        headers = [("日","sun"),("月",""),("火",""),("水",""),("木",""),("金",""),("土","sat")]
        col_of  = lambda wd: (wd + 1) % 7
        is_sun_col = lambda col: col == 0
        is_sat_col = lambda col: col == 6
    else:  # monday
        # 列0=月(Mon), ..., 5=土(Sat), 6=日(Sun)
        headers = [("月",""),("火",""),("水",""),("木",""),("金",""),("土","sat"),("日","sun")]
        col_of  = lambda wd: wd          # Mon=0→0, ..., Sun=6→6
        is_sun_col = lambda col: col == 6
        is_sat_col = lambda col: col == 5
    return headers, col_of, is_sun_col, is_sat_col


def _week_start_of(d: date, col_of) -> date:
    """col_of を使って d の属する週の最初の日を返す。"""
    # col_of(d.weekday()) = この日が何列目か → その分だけ引く
    return d - timedelta(days=col_of(d.weekday()))


# ────────────────────────────────────────────────────────────────────────────
# カレンダー共通: ICS / Google Calendar URL ヘルパー
# ────────────────────────────────────────────────────────────────────────────

def _ics_text(s: Any) -> str:
    """ICS テキストエスケープ（RFC 5545）"""
    return str(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _google_cal_url(summary: str, d: date, description: str = "") -> str:
    """Google Calendar へのイベント追加 URL（終日）を生成"""
    from urllib.parse import quote
    d_end = (d + timedelta(days=1)).strftime("%Y%m%d")
    url = (
        "https://calendar.google.com/calendar/render?action=TEMPLATE"
        f"&text={quote(summary, safe='')}"
        f"&dates={d.strftime('%Y%m%d')}/{d_end}"
    )
    if description:
        url += f"&details={quote(description, safe='')}"
    return url


def _build_ics(events: List[Dict], cal_name: str = "FALCON") -> str:
    """イベントリストから ICS 文字列を生成（RFC 5545 準拠）"""
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//FALCON//FALCON Dairy Management//JA",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{_ics_text(cal_name)}",
        "X-WR-TIMEZONE:Asia/Tokyo",
    ]
    for ev in events:
        d: date = ev["date"]
        uid = (
            f"falcon-{ev.get('uid_key', 'ev')}-{d.strftime('%Y%m%d')}"
            f"-{abs(hash(ev.get('summary', ''))) % 99999}@falcon"
        )
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTART;VALUE=DATE:{d.strftime('%Y%m%d')}",
            f"DTEND;VALUE=DATE:{(d + timedelta(days=1)).strftime('%Y%m%d')}",
            f"SUMMARY:{_ics_text(ev.get('summary', ''))}",
        ]
        if ev.get("description"):
            lines.append(f"DESCRIPTION:{_ics_text(ev['description'])}")
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


def _ics_embed_script(ics_content: str, fn_name: str, filename: str) -> str:
    """ICS を Base64 埋め込みした JS ダウンロード関数を返す（ダウンロード後にガイドモーダルを表示）"""
    import base64
    b64 = base64.b64encode(ics_content.encode("utf-8")).decode("ascii")
    return (
        f"<script>\n"
        f"function {fn_name}(){{\n"
        f"  var b=atob('{b64}'),n=new Uint8Array(b.length);\n"
        f"  for(var i=0;i<b.length;i++)n[i]=b.charCodeAt(i);\n"
        f"  var u=URL.createObjectURL(new Blob([n],{{type:'text/calendar;charset=utf-8'}}));\n"
        f"  var a=document.createElement('a');a.href=u;a.download='{filename}';a.click();\n"
        f"  URL.revokeObjectURL(u);\n"
        f"  document.getElementById('ics-guide-modal').style.display='flex';\n"
        f"}}\n"
        f"</script>\n"
    )


def _ics_guide_modal_html() -> str:
    """iCal ダウンロード後に表示するガイドモーダルの HTML + CSS を返す"""
    return """
<style>
#ics-guide-modal {
  display: none; position: fixed; inset: 0;
  background: rgba(0,0,0,0.45); z-index: 9999;
  align-items: center; justify-content: center;
}
.ics-modal {
  background: #fff; border-radius: 12px; padding: 28px 30px 22px;
  max-width: 480px; width: 92%;
  box-shadow: 0 8px 30px rgba(0,0,0,0.20);
  font-family: "Meiryo UI","Yu Gothic","MS Gothic",sans-serif;
}
.ics-modal-head {
  display: flex; align-items: center; gap: 10px; margin-bottom: 18px;
}
.ics-modal-head h3 {
  font-size: 1.05em; font-weight: 500; color: #3c4043; margin: 0;
}
.ics-steps { list-style: none; padding: 0; margin: 0 0 16px; }
.ics-steps li {
  display: flex; gap: 12px; align-items: flex-start;
  padding: 9px 0; border-bottom: 1px solid #f1f3f4;
  font-size: 0.87em; color: #3c4043; line-height: 1.5;
}
.ics-steps li:last-child { border-bottom: none; }
.ics-num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 22px; height: 22px; min-width: 22px; border-radius: 50%;
  background: #1a73e8; color: #fff; font-size: 0.75em; font-weight: 600;
  margin-top: 1px;
}
.ics-step-body strong { color: #1a73e8; }
.ics-step-body code {
  background: #f1f3f4; border-radius: 3px; padding: 1px 5px;
  font-size: 0.88em; color: #3c4043;
}
.ics-tip {
  background: #e8f0fe; border-radius: 8px; padding: 9px 14px;
  font-size: 0.82em; color: #1967d2; margin-bottom: 18px; line-height: 1.5;
}
.ics-modal-footer {
  display: flex; gap: 8px; justify-content: flex-end; align-items: center;
}
.ics-btn-gcal {
  background: #fff; color: #1a73e8; border: 1px solid #dadce0;
  border-radius: 20px; padding: 6px 16px; font-size: 0.83em;
  font-weight: 500; cursor: pointer; text-decoration: none;
  display: inline-flex; align-items: center; gap: 4px;
}
.ics-btn-gcal:hover { background: #e8f0fe; }
.ics-btn-close {
  background: #1a73e8; color: #fff; border: none;
  border-radius: 20px; padding: 6px 20px; font-size: 0.83em;
  font-weight: 500; cursor: pointer;
}
.ics-btn-close:hover { background: #1557b0; }
@media print { #ics-guide-modal { display: none !important; } }
</style>
<div id="ics-guide-modal">
  <div class="ics-modal">
    <div class="ics-modal-head">
      <span style="font-size:1.4em">📅</span>
      <h3>Googleカレンダーへの取り込み方法</h3>
    </div>
    <ol class="ics-steps">
      <li>
        <span class="ics-num">1</span>
        <span class="ics-step-body">
          <strong>.ics ファイルがダウンロードされました</strong><br>
          ブラウザ下部のバー、または「ダウンロード」フォルダを確認してください
        </span>
      </li>
      <li>
        <span class="ics-num">2</span>
        <span class="ics-step-body">
          下の「Googleカレンダーを開く」ボタンをクリック<br>
          （または <a href="https://calendar.google.com" target="_blank" style="color:#1a73e8">calendar.google.com</a> を開く）
        </span>
      </li>
      <li>
        <span class="ics-num">3</span>
        <span class="ics-step-body">
          右上の <strong>⚙ アイコン</strong> をクリック → <strong>「設定」</strong> を選ぶ
        </span>
      </li>
      <li>
        <span class="ics-num">4</span>
        <span class="ics-step-body">
          左メニューの <strong>「インポート / エクスポート」</strong> をクリック
        </span>
      </li>
      <li>
        <span class="ics-num">5</span>
        <span class="ics-step-body">
          「コンピュータからファイルを選択」で、ダウンロードした
          <code>.ics</code> ファイルを選ぶ
        </span>
      </li>
      <li>
        <span class="ics-num">6</span>
        <span class="ics-step-body">
          <strong>「インポート」</strong> ボタンを押して完了！
        </span>
      </li>
    </ol>
    <div class="ics-tip">
      💡 各チップをクリックすると、1件ずつ直接 Googleカレンダーに追加することもできます
    </div>
    <div class="ics-modal-footer">
      <a class="ics-btn-gcal" href="https://calendar.google.com" target="_blank">
        Googleカレンダーを開く ↗
      </a>
      <button class="ics-btn-close"
        onclick="document.getElementById('ics-guide-modal').style.display='none'">
        閉じる
      </button>
    </div>
  </div>
</div>
"""


# ────────────────────────────────────────────────────────────────────────────
# 2. 発情カレンダー
# ────────────────────────────────────────────────────────────────────────────

def build_estrus_calendar_html(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
    farm_name: str,
    id_label: str = "ID",
    week_start: str = "sunday",
) -> str:
    today = _parse_date(checkup_date) or date.today()

    headers, col_of, is_sun_col, is_sat_col = _week_config(week_start)

    # カレンダー範囲: 今日の属する週の2週前〜12週後
    week_of_today = _week_start_of(today, col_of)
    cal_start = week_of_today - timedelta(weeks=2)
    cal_end   = cal_start + timedelta(weeks=12)

    cows = _get_active_parous_cows(db, rule_engine)

    # 日付 → イベントリスト
    calendar_data: Dict[date, List[Dict]] = {}

    def _add(d: date, entry: Dict):
        calendar_data.setdefault(d, []).append(entry)

    for cow in cows:
        rc = cow["rc"]
        if rc in (rule_engine.RC_PREGNANT, rule_engine.RC_DRY):
            continue
        last_ai = _parse_date(cow["last_ai_date"])
        if last_ai:
            d21 = last_ai + timedelta(days=21)
            d42 = last_ai + timedelta(days=42)
            if cal_start <= d21 <= cal_end:
                pen_str = f" [{cow['pen']}]" if cow.get("pen") else ""
                _add(d21, {"cow_id": cow["cow_id"], "pen": cow.get("pen", ""), "type": "cycle1",
                            "label": f"① {cow['cow_id']}{pen_str}"})
            if cal_start <= d42 <= cal_end:
                pen_str = f" [{cow['pen']}]" if cow.get("pen") else ""
                _add(d42, {"cow_id": cow["cow_id"], "pen": cow.get("pen", ""), "type": "cycle2",
                            "label": f"② {cow['cow_id']}{pen_str}"})

        # cycle_day 発情注意
        cd_info = _get_cycle_day_info(db, rule_engine, cow["auto_id"])
        if cd_info:
            exam_date, cd = cd_info
            alert_d = exam_date + timedelta(days=21 - cd)
            pen_str = f" [{cow['pen']}]" if cow.get("pen") else ""
            for offset, suffix in [(-1, "前日"), (0, ""), (1, "翌日")]:
                ad = alert_d + timedelta(days=offset)
                if cal_start <= ad <= cal_end:
                    extra = f"（{suffix}）" if suffix else ""
                    _add(ad, {"cow_id": cow["cow_id"], "pen": cow.get("pen", ""), "type": "alert",
                               "label": f"⚡ {cow['cow_id']}{pen_str} 発情注意{extra}"})

    # ── ICS 生成（calendar_data から全イベントを収集）──
    _type_labels = {"cycle1": "AI+21日", "cycle2": "AI+42日", "alert": "発情注意"}
    ics_events: List[Dict] = []
    for _d, _evs in sorted(calendar_data.items()):
        for _ev in _evs:
            _label = _type_labels.get(_ev["type"], _ev["type"])
            _summary = f"[発情] {_ev['cow_id']} {_label}"
            if _ev.get("pen"):
                _summary += f" [{_ev['pen']}]"
            ics_events.append({
                "date": _d,
                "uid_key": f"estrus-{_ev['cow_id']}-{_ev['type']}",
                "summary": _summary,
                "description": f"農場: {farm_name}\\n基準日: {checkup_date}",
            })
    _ics_str = _build_ics(ics_events, cal_name=f"発情カレンダー - {farm_name}")
    _ics_script = _ics_embed_script(_ics_str, "downloadEstrusICS", "estrus_calendar.ics")

    # CSS (Google Calendar inspired)
    css_extra = """
    /* ── ページヘッダー上書き ── */
    .page-header {
      background: #fff; color: #3c4043; border-bottom: 1px solid #e0e0e0;
      padding: 12px 20px;
    }
    .page-header h1 { font-size: 1.25em; font-weight: 400; color: #3c4043; }
    .page-header .meta { color: #70757a; }

    /* ── カレンダー本体 ── */
    .cal-wrapper { overflow-x: auto; background: #fff; border-radius: 8px;
                   box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.08); }
    .cal-table { border-collapse: collapse; min-width: 700px; width: 100%; }

    /* 曜日ヘッダー */
    .cal-table th {
      background: #fff; color: #70757a; text-align: center; padding: 10px 4px 8px;
      font-size: 0.72em; font-weight: 500; letter-spacing: 0.04em;
      text-transform: uppercase; border-bottom: 1px solid #e0e0e0;
    }
    .cal-table th.sun { color: #c62828; }
    .cal-table th.sat { color: #1565c0; }

    /* 日付セル */
    .cal-table td {
      border: 1px solid #e0e0e0; vertical-align: top; width: 14.28%;
      min-height: 96px; padding: 6px 6px 4px; font-size: 0.77em;
      background: #fff;
    }
    .cal-table td.sun-cell { background: #fafafa; }
    .cal-table td.sat-cell { background: #fafafa; }

    /* 日付番号 */
    .day-num {
      display: inline-flex; align-items: center; justify-content: center;
      width: 26px; height: 26px; border-radius: 50%;
      font-size: 0.88em; font-weight: 400; color: #3c4043;
      margin-bottom: 4px;
    }
    .day-num.today-num { background: #1a73e8; color: #fff; font-weight: 500; }
    .day-num.sun-num   { color: #c62828; }
    .day-num.sat-num   { color: #1565c0; }

    /* イベントチップ（淡背景＋左ボーダーアクセント） */
    .ev-chip {
      display: block; border-radius: 4px; padding: 2px 6px; margin-bottom: 3px;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%;
      font-size: 0.92em; line-height: 1.5;
    }
    .ev-chip.cycle1 { background: #e8f0fe; color: #1967d2; border-left: 3px solid #1a73e8; }
    .ev-chip.cycle2 { background: #f3e8fd; color: #7627bb; border-left: 3px solid #9334e6; }
    .ev-chip.alert  { background: #fce8e6; color: #c5221f; border-left: 3px solid #ea4335;
                      font-weight: 500; }

    /* 凡例 */
    .legend-wrap { display: flex; gap: 10px; align-items: center; flex-wrap: wrap;
                   margin-bottom: 12px; }

    /* iCal ダウンロードボタン */
    .ics-btn {
      display: inline-flex; align-items: center; gap: 5px;
      background: #1a73e8; color: #fff; border: none; border-radius: 20px;
      padding: 5px 14px; font-size: 0.8em; font-weight: 500; cursor: pointer;
      margin-left: auto; text-decoration: none; white-space: nowrap;
    }
    .ics-btn:hover { background: #1557b0; }

    /* チップリンク */
    .chip-link { text-decoration: none; display: block; }
    .chip-link .ev-chip:hover { opacity: 0.82; }

    /* ── 印刷設定 ── */
    @media print {
      @page { size: A4 landscape; margin: 8mm 6mm; }
      body { background: #fff !important; font-size: 10px; }
      .page-header { background: #fff !important; border-bottom: 1px solid #e0e0e0 !important; }
      .cal-wrapper { box-shadow: none !important; }
      .ics-btn { display: none !important; }
      .cal-table tr { page-break-inside: avoid; break-inside: avoid; }
      .ev-chip, .day-num.today-num, .cal-table td.sun-cell, .cal-table td.sat-cell {
        -webkit-print-color-adjust: exact; print-color-adjust: exact; color-adjust: exact;
      }
      .ev-chip.cycle1 { background: #e8f0fe !important; color: #1967d2 !important; }
      .ev-chip.cycle2 { background: #f3e8fd !important; color: #7627bb !important; }
      .ev-chip.alert  { background: #fce8e6 !important; color: #c5221f !important; }
      .day-num.today-num { background: #1a73e8 !important; color: #fff !important; }
    }
    """

    week_label = "日曜始まり" if week_start == "sunday" else "月曜始まり"
    body = _common_html_head(f"発情カレンダー | {farm_name}", extra_style=css_extra)
    body += _ics_script
    body += _ics_guide_modal_html()
    body += f"""
<div class="page-header">
  <div>
    <h1>発情カレンダー</h1>
    <div style="font-size:0.82em;color:#70757a;margin-top:2px">{_esc(farm_name)}&nbsp;·&nbsp;{week_label}</div>
  </div>
  <div class="meta" style="display:flex;align-items:center;gap:12px">
    <span>基準日: {_esc(checkup_date)}&nbsp;&nbsp;出力: {date.today().strftime('%Y/%m/%d')}</span>
    <button class="ics-btn" onclick="downloadEstrusICS()">📥 iCalダウンロード（全{len(ics_events)}件）</button>
  </div>
</div>
<div class="container">
<div class="legend-wrap">
  <span style="font-size:0.82em;color:#5f6368;font-weight:500">凡例</span>
  <span class="ev-chip cycle1" style="display:inline-block;max-width:none">① AI+21日</span>
  <span class="ev-chip cycle2" style="display:inline-block;max-width:none">② AI+42日</span>
  <span class="ev-chip alert"  style="display:inline-block;max-width:none">⚡ 発情注意</span>
  <span style="font-size:0.78em;color:#9e9e9e">※ チップクリックで個別にGoogleカレンダーへ追加</span>
</div>
<div class="cal-wrapper">
<table class="cal-table">
<thead>
<tr>
"""
    for label, cls in headers:
        body += f'  <th class="{cls}">{label}</th>\n'
    body += "</tr></thead>\n<tbody>\n"

    # 週ごとに描画
    cur = cal_start
    while cur < cal_end:
        body += "<tr>"
        for col in range(7):
            d = cur + timedelta(days=col)
            is_today = (d == today)
            is_sun = is_sun_col(col)
            is_sat = is_sat_col(col)
            cell_cls = "today-cell" if is_today else ("sun-cell" if is_sun else ("sat-cell" if is_sat else ""))
            num_cls  = "today-num"  if is_today else ("sun-num"  if is_sun else ("sat-num"  if is_sat else ""))

            body += f'<td class="{cell_cls}">'
            body += f'<div class="day-num {num_cls}">'
            body += str(d.day)
            body += '</div>'
            if d.day == 1:
                body += f'<div style="font-size:0.78em;color:#70757a;margin-bottom:2px">{d.month}月</div>'

            events_on_day = sorted(
                calendar_data.get(d, []),
                key=lambda e: ({"cycle1": 0, "cycle2": 1, "alert": 2}.get(e["type"], 9),
                               e.get("pen", ""), e["cow_id"]),
            )
            for ev in events_on_day:
                _type_label = _type_labels.get(ev["type"], ev["type"])
                _gcal_summary = f"[発情] {ev['cow_id']} {_type_label}"
                if ev.get("pen"):
                    _gcal_summary += f" [{ev['pen']}]"
                _gcal_url = _google_cal_url(
                    _gcal_summary, d,
                    f"農場: {farm_name} / 基準日: {checkup_date}"
                )
                body += (
                    f'<a class="chip-link" href="{_esc(_gcal_url)}" target="_blank"'
                    f' title="Googleカレンダーに追加: {_esc(_gcal_summary)}">'
                    f'<span class="ev-chip {_esc(ev["type"])}">{_esc(ev["label"])}</span>'
                    f'</a>'
                )

            body += '</td>'
        body += "</tr>"
        cur += timedelta(weeks=1)

    body += "</tbody></table></div></div></body></html>"
    return body


# ────────────────────────────────────────────────────────────────────────────
# 3. 分娩予定表
# ────────────────────────────────────────────────────────────────────────────

def build_calving_plan_table_html(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
    farm_name: str,
    id_label: str = "ID",
) -> str:
    today = _parse_date(checkup_date) or date.today()

    cows = _get_active_parous_cows(db, rule_engine)

    rows: List[Dict[str, Any]] = []
    for cow in cows:
        due = _parse_date(cow["due_date"])
        if due is None:
            continue  # 分娩予定日のない牛は除外

        conception = _parse_date(cow["conception_date"])
        clvd = _parse_date(cow["clvd"])

        # 受胎AIイベントから SIRE を取得
        sire = ""
        ai_ev = _get_conception_ai_event(db, rule_engine, cow["auto_id"], cow["conception_date"])
        if ai_ev:
            j = _get_json(ai_ev)
            sire = j.get("sire") or ""

        # 妊娠確定イベントから twin / sex を取得
        twin = False
        sex = ""
        preg_ev = _get_pregnancy_event(db, rule_engine, cow["auto_id"])
        if preg_ev:
            j = _get_json(preg_ev)
            twin = bool(j.get("twin") or False)
            sex = j.get("sex") or ""

        # 予定分娩間隔: 現在の分娩日 → 分娩予定日
        pci = (due - clvd).days if clvd else None

        # 分娩予定日 -60日 / -21日
        due_m60 = due - timedelta(days=60)
        due_m21 = due - timedelta(days=21)

        # 状態フラグ
        status = ""
        if due < today:
            status = "overdue"
        elif (due - today).days <= 21:
            status = "near"
        elif (due - today).days <= 60:
            status = "soon"

        rows.append({
            "cow_id": cow["cow_id"],
            "jpn10": cow["jpn10"],
            "lact": cow["lact"],
            "conception": conception,
            "due_m60": due_m60,
            "due_m21": due_m21,
            "due": due,
            "pci": pci,
            "sire": sire,
            "sex": sex,
            "twin": twin,
            "status": status,
        })

    rows.sort(key=lambda r: (r["due"], r["cow_id"]))

    css_extra = """
    .row-overdue td { background: #ffebee !important; }
    .row-near td { background: #fff8e1 !important; }
    .row-soon td { background: #f3e5f5 !important; }
    .twin-yes { color: #c62828; font-weight: bold; }
    """

    body = _common_html_head(f"分娩予定表 | {farm_name}", extra_style=css_extra)
    body += f"""
<div class="page-header">
  <div>
    <h1>🐄 分娩予定表</h1>
    <div style="font-size:0.85em;opacity:0.85">{_esc(farm_name)}</div>
  </div>
  <div class="meta">基準日: {_esc(checkup_date)}&nbsp;&nbsp;出力: {date.today().strftime('%Y/%m/%d')}</div>
</div>
<div class="container">
<div class="section-title">分娩予定表（妊娠確定牛・分娩予定日順）</div>
<div style="margin-bottom:10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:0.85em">
  <span>行の色：</span>
  <span style="background:#ffebee;padding:2px 8px;border-radius:3px">予定日超過</span>
  <span style="background:#fff8e1;padding:2px 8px;border-radius:3px">21日以内</span>
  <span style="background:#f3e5f5;padding:2px 8px;border-radius:3px">60日以内</span>
</div>
<table class="report-table">
<thead>
<tr>
  <th>{_esc(id_label)}</th>
  <th>JPN10</th>
  <th>産次</th>
  <th>受胎AI日</th>
  <th>予定-60日</th>
  <th>予定-21日</th>
  <th>分娩予定日</th>
  <th>予定分娩間隔</th>
  <th>受胎SIRE</th>
  <th>性別</th>
  <th>双子</th>
</tr>
</thead>
<tbody>
"""

    if not rows:
        body += '<tr><td colspan="11" style="text-align:center;color:#90a4ae;padding:20px">対象牛なし（妊娠確定牛なし）</td></tr>'

    for r in rows:
        row_cls = f' class="row-{r["status"]}"' if r["status"] else ""
        twin_str = '<span class="twin-yes">双子</span>' if r["twin"] else "—"
        sex_disp = _esc(r["sex"]) if r["sex"] else "—"
        pci_str = f'{r["pci"]}日' if r["pci"] else "—"

        def _due_cell(d: date, ref: date) -> str:
            diff = (d - ref).days
            if diff < 0:
                return f'<span style="color:#90a4ae">{_fmt_date(d)}</span>'
            elif diff <= 7:
                return f'<b style="color:#c62828">{_fmt_date(d)}</b>'
            return _fmt_date(d)

        body += f"""<tr{row_cls}>
  <td><b>{_esc(r["cow_id"])}</b></td>
  <td style="font-size:0.85em">{_esc(r["jpn10"])}</td>
  <td>{_esc(r["lact"])}</td>
  <td>{_fmt_date(r["conception"])}</td>
  <td>{_due_cell(r["due_m60"], today)}</td>
  <td>{_due_cell(r["due_m21"], today)}</td>
  <td><b>{_due_cell(r["due"], today)}</b></td>
  <td>{pci_str}</td>
  <td>{_esc(r["sire"]) if r["sire"] else "—"}</td>
  <td>{sex_disp}</td>
  <td>{twin_str}</td>
</tr>
"""

    body += "</tbody></table></div></body></html>"
    return body


# ────────────────────────────────────────────────────────────────────────────
# 分娩予定カレンダー用: イベントから SIRE / SEX / TWIN を一括抽出
# ────────────────────────────────────────────────────────────────────────────

def _extract_sire_sex_twin(events: List[Dict], rule_engine: RuleEngine) -> Tuple[str, str, bool]:
    """
    牛の全イベントリストから (sire, sex, twin) を抽出する。
    - sire : 最後の AI/ET イベントの json_data['sire']
    - sex  : 最新の妊娠確定イベント (PDP/PAGP) の json_data['sex']
    - twin : 同上の json_data['twin']
    """
    sire = ""
    sex = ""
    twin = False

    preg_nums = (rule_engine.EVENT_PDP, rule_engine.EVENT_PAGP,
                 getattr(rule_engine, "EVENT_PDP2", 304))
    preg_evs = sorted(
        [e for e in events if e.get("event_number") in preg_nums],
        key=lambda e: (e.get("event_date") or "", e.get("id") or 0),
        reverse=True,
    )
    if preg_evs:
        j = _get_json(preg_evs[0])
        sex  = str(j.get("sex") or "").strip()
        twin = bool(j.get("twin") or False)

    ai_evs = sorted(
        [e for e in events if e.get("event_number") in (rule_engine.EVENT_AI, rule_engine.EVENT_ET)],
        key=lambda e: (e.get("event_date") or "", e.get("id") or 0),
        reverse=True,
    )
    if ai_evs:
        j = _get_json(ai_evs[0])
        sire = str(j.get("sire") or "").strip()

    return sire, sex, twin


def _load_sire_list(farm_path: Path) -> Dict[str, Any]:
    p = farm_path / "sire_list.json"
    if not p.exists():
        return {}
    try:
        with open(p, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return raw if isinstance(raw, dict) else {}
    except Exception:
        return {}


def _is_holstein_female_sire(sire_name: str, sire_list: Dict[str, Any]) -> bool:
    """sire_list で種別が乳用種♀（メスSIRE）か。"""
    if not sire_name or not str(sire_name).strip():
        return False
    sn = str(sire_name).strip()
    for key in (sn, sn.upper(), sn.lower()):
        opts = sire_list.get(key)
        if isinstance(opts, dict) and sire_opts_to_type(opts) == SIRE_TYPE_HOLSTEIN_FEMALE:
            return True
    return False


def _calving_calendar_sex_suffix(sex: str, sire: str, sire_list: Dict[str, Any]) -> str:
    """
    カレンダー1行目に付ける性別まわりの文言。
    胎子性別が未登録のとき、メスSIRE（乳用種♀）のみ「性別不明」を付ける。
    それ以外は空（2行目の種雄牛表示に任せる）。
    """
    raw = (sex or "").strip()
    if raw:
        return raw
    if _is_holstein_female_sire(sire, sire_list):
        return "性別不明"
    return ""


# ────────────────────────────────────────────────────────────────────────────
# 4. 分娩予定カレンダー
# ────────────────────────────────────────────────────────────────────────────

def build_calving_plan_calendar_html(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
    farm_name: str,
    id_label: str = "ID",
    week_start: str = "sunday",
) -> str:
    today = _parse_date(checkup_date) or date.today()

    headers, col_of, is_sun_col, is_sat_col = _week_config(week_start)

    # カレンダー範囲: 今月から7ヶ月分
    cal_start = date(today.year, today.month, 1)
    end_month = today.month + 6
    end_year = today.year + (end_month - 1) // 12
    end_month = ((end_month - 1) % 12) + 1
    cal_end = date(end_year, end_month, calendar.monthrange(end_year, end_month)[1])

    cows = _get_active_parous_cows(db, rule_engine)

    # 各牛のイベントを1回だけ取得し、SIRE/SEX/TWIN を抽出
    calving_data: Dict[date, List[Dict]] = {}
    for cow in cows:
        due = _parse_date(cow["due_date"])
        if due is None or not (cal_start <= due <= cal_end):
            continue
        events = db.get_events_by_cow(cow["auto_id"], include_deleted=False)
        sire, sex, twin = _extract_sire_sex_twin(events, rule_engine)
        calving_data.setdefault(due, []).append({
            "cow_id": cow["cow_id"],
            "lact": cow["lact"],
            "sire": sire,
            "sex": sex,
            "twin": twin,
        })

    sire_list = _load_sire_list(farm_path)

    # ── ICS 生成（calving_data から全イベントを収集）──
    _calving_ics_events: List[Dict] = []
    for _due, _cows_list in sorted(calving_data.items()):
        for _c in _cows_list:
            _parts = []
            if _c["lact"]: _parts.append(f"{_c['lact']}産")
            _sex_suf = _calving_calendar_sex_suffix(_c["sex"], _c["sire"], sire_list)
            if _sex_suf:
                _parts.append(_sex_suf)
            if _c["twin"]: _parts.append("双子")
            _summary = f"[分娩予定] {_c['cow_id']}"
            _desc_parts = [f"農場: {farm_name}", f"基準日: {checkup_date}"]
            if _parts: _desc_parts.append(" / ".join(_parts))
            if _c["sire"]: _desc_parts.append(f"SIRE: {_c['sire']}")
            _calving_ics_events.append({
                "date": _due,
                "uid_key": f"calving-{_c['cow_id']}",
                "summary": _summary,
                "description": "\\n".join(_desc_parts),
            })
    _calving_ics_str = _build_ics(_calving_ics_events, cal_name=f"分娩予定カレンダー - {farm_name}")
    _calving_ics_script = _ics_embed_script(_calving_ics_str, "downloadCalvingICS", "calving_calendar.ics")

    css_extra = """
    /* ── ページヘッダー上書き ── */
    .page-header {
      background: #fff; color: #3c4043; border-bottom: 1px solid #e0e0e0; padding: 12px 20px;
    }
    .page-header h1 { font-size: 1.25em; font-weight: 400; color: #3c4043; }
    .page-header .meta { color: #70757a; }

    /* ── 月ブロック ── */
    .month-block { margin-bottom: 40px; }
    .month-title {
      font-size: 1.15em; font-weight: 400; color: #3c4043;
      padding: 8px 0 10px; margin-bottom: 0; letter-spacing: -0.01em;
    }

    /* ── 月次サマリー ── */
    .month-summary {
      display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
      background: #f8f9fa; border-radius: 8px; padding: 8px 12px;
      margin-bottom: 10px; font-size: 0.85em; border: 1px solid #e0e0e0;
    }
    .sum-item {
      display: flex; flex-direction: column; align-items: center;
      background: #fff; border-radius: 6px; padding: 4px 14px;
      border: 1px solid #e0e0e0; min-width: 68px;
    }
    .sum-label { font-size: 0.76em; color: #70757a; margin-bottom: 1px; }
    .sum-val   { font-size: 1.2em; font-weight: 500; color: #3c4043; }
    .sum-item.female .sum-val  { color: #d81b60; }
    .sum-item.male   .sum-val  { color: #1565c0; }
    .sum-item.twin   .sum-val  { color: #e65100; }
    .sum-item.unknown .sum-val { color: #9e9e9e; }
    .twin-warn {
      background: #e65100; color: #fff; border-radius: 4px;
      padding: 1px 6px; font-size: 0.76em; font-weight: 500; display: inline-block; margin-left: 4px;
    }

    /* ── カレンダーグリッド ── */
    .cal-grid {
      display: grid; grid-template-columns: repeat(7, 1fr);
      border-top: 1px solid #e0e0e0; border-left: 1px solid #e0e0e0;
    }
    .cal-dow {
      text-align: center; font-size: 0.72em; font-weight: 500; padding: 8px 4px;
      background: #fff; color: #70757a; letter-spacing: 0.04em; text-transform: uppercase;
      border-right: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0;
    }
    .cal-dow.sat { color: #1565c0; }
    .cal-dow.sun { color: #c62828; }

    /* ── カレンダーセル ── */
    .cal-cell {
      border-right: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0;
      background: #fff; min-height: 90px; padding: 6px 6px 4px; font-size: 0.75em;
    }
    .cal-cell.empty { background: #fafafa; }
    .cal-cell.sun-cell { background: #fafafa; }
    .cal-cell.sat-cell { background: #fafafa; }
    .cal-cell.has-calving { background: #fff; }

    /* 日付番号（Google Calendar 式丸バッジ） */
    .cal-day {
      display: inline-flex; align-items: center; justify-content: center;
      width: 24px; height: 24px; border-radius: 50%;
      font-size: 0.85em; font-weight: 400; color: #3c4043; margin-bottom: 4px;
    }
    .cal-day.today-day { background: #1a73e8; color: #fff; font-weight: 500; }
    .cal-day.sun-day   { color: #c62828; }
    .cal-day.sat-day   { color: #1565c0; }

    /* ── 牛チップ ── */
    .cow-chip {
      border-radius: 4px; padding: 3px 6px; margin-bottom: 3px; display: block;
      line-height: 1.4; border-left: 3px solid #1a73e8; background: #f8f9fa;
    }
    .cow-chip .chip-id   { font-weight: 500; color: #3c4043; font-size: 0.95em; }
    .cow-chip .chip-sub  { color: #70757a; font-size: 0.85em; }
    .cow-chip .chip-sire {
      color: #5f6368; font-size: 0.82em; white-space: nowrap;
      overflow: hidden; text-overflow: ellipsis; display: block;
    }
    .cow-chip.has-twin { border-left-color: #e65100; background: #fff3e0; }
    .cow-count  { font-size: 0.78em; color: #9e9e9e; text-align: right; margin-top: 2px; }
    .more-badge {
      display: inline-block; background: #e0e0e0; color: #5f6368;
      border-radius: 10px; padding: 1px 7px; font-size: 0.78em; margin-top: 2px;
    }
    .print-note { font-size: 0.82em; color: #9e9e9e; margin-bottom: 12px; }

    /* iCal ダウンロードボタン */
    .ics-btn {
      display: inline-flex; align-items: center; gap: 5px;
      background: #1a73e8; color: #fff; border: none; border-radius: 20px;
      padding: 5px 14px; font-size: 0.8em; font-weight: 500; cursor: pointer;
      white-space: nowrap; text-decoration: none;
    }
    .ics-btn:hover { background: #1557b0; }

    /* チップリンク */
    .chip-link { text-decoration: none; display: block; }
    .chip-link .cow-chip:hover { opacity: 0.82; cursor: pointer; }

    /* ── 印刷設定 ── */
    @media print {
      @page { size: A4 landscape; margin: 8mm 6mm; }
      body { background: #fff !important; font-size: 11px; }
      .print-note { display: none; }
      .ics-btn { display: none !important; }
      .page-header { background: #fff !important; border-bottom: 1px solid #e0e0e0 !important; }
      .month-block {
        page-break-before: always; break-before: page;
        page-break-inside: avoid; break-inside: avoid; margin-bottom: 0;
      }
      .month-block.first-month { page-break-before: auto; break-before: auto; }
      .cal-grid { page-break-inside: avoid; break-inside: avoid; }
      .cal-cell { page-break-inside: avoid; break-inside: avoid; min-height: 60px; }
      .cal-day.today-day, .cow-chip, .cow-chip.has-twin,
      .sum-item, .month-summary, .twin-warn {
        -webkit-print-color-adjust: exact; print-color-adjust: exact; color-adjust: exact;
      }
      .cal-day.today-day { background: #1a73e8 !important; color: #fff !important; }
      .cow-chip           { background: #f8f9fa !important; }
      .cow-chip.has-twin  { background: #fff3e0 !important; }
      .month-summary      { background: #f8f9fa !important; }
      .twin-warn          { background: #e65100 !important; color: #fff !important; }
    }
    """

    week_label = "日曜始まり" if week_start == "sunday" else "月曜始まり"
    body = _common_html_head(f"分娩予定カレンダー | {farm_name}", extra_style=css_extra)
    body += _calving_ics_script
    body += _ics_guide_modal_html()
    body += f"""
<div class="page-header">
  <div>
    <h1>分娩予定カレンダー</h1>
    <div style="font-size:0.82em;color:#70757a;margin-top:2px">{_esc(farm_name)}&nbsp;·&nbsp;{week_label}</div>
  </div>
  <div class="meta" style="display:flex;align-items:center;gap:12px">
    <span>基準日: {_esc(checkup_date)}&nbsp;&nbsp;出力: {date.today().strftime('%Y/%m/%d')}</span>
    <button class="ics-btn" onclick="downloadCalvingICS()">📥 iCalダウンロード（全{len(_calving_ics_events)}件）</button>
  </div>
</div>
<div class="container">
<div class="print-note">
  ※ 乳用種♀ = 性別判定済み（超音波検査記録）の♀。SIRE は受胎AI記録より。
  双子⚠️ は事故リスク高・哺育準備を早めに。
</div>
"""

    year_m = today.year
    month_m = today.month
    DISPLAY_MAX = 6  # 1日に最大表示頭数
    is_first_month = True

    for _ in range(7):
        _, last_day = calendar.monthrange(year_m, month_m)
        month_label = f"{year_m}年{month_m}月"

        # ── 月次サマリー集計 ──
        month_cows: List[Dict] = []
        for day_num in range(1, last_day + 1):
            d = date(year_m, month_m, day_num)
            month_cows.extend(calving_data.get(d, []))

        total   = len(month_cows)
        female  = sum(1 for c in month_cows if c["sex"] in ("♀", "F", "f", "めす", "雌"))
        male    = sum(1 for c in month_cows if c["sex"] in ("♂", "M", "m", "おす", "雄"))
        unknown = total - female - male
        twins   = sum(1 for c in month_cows if c["twin"])

        block_cls = "month-block first-month" if is_first_month else "month-block"
        is_first_month = False
        body += f'<div class="{block_cls}">'
        body += f'<div class="month-title">{month_label}</div>'

        # サマリーボックス
        if total > 0:
            twin_warn = f'<span class="twin-warn">⚠️双子{twins}件</span>' if twins > 0 else ""
            body += f"""
<div class="month-summary">
  <div class="sum-item">
    <span class="sum-label">合計</span>
    <span class="sum-val">{total}頭</span>
  </div>
  <div class="sum-item female">
    <span class="sum-label">乳用種♀</span>
    <span class="sum-val">{female}頭</span>
  </div>
  <div class="sum-item male">
    <span class="sum-label">乳用種♂</span>
    <span class="sum-val">{male}頭</span>
  </div>
  <div class="sum-item unknown">
    <span class="sum-label">性別不明</span>
    <span class="sum-val">{unknown}頭</span>
  </div>
  {f'<div class="sum-item twin"><span class="sum-label">双子⚠️</span><span class="sum-val">{twins}件</span></div>' if twins > 0 else ""}
</div>
"""
        else:
            body += '<div class="month-summary" style="color:#9e9e9e;font-size:0.85em">分娩予定なし</div>'

        # カレンダーグリッド
        body += '<div class="cal-grid">'
        for wd_label, wd_cls in headers:
            body += f'<div class="cal-dow {wd_cls}">{wd_label}</div>'

        first_wd = date(year_m, month_m, 1).weekday()
        blank_before = col_of(first_wd)
        for _ in range(blank_before):
            body += '<div class="cal-cell empty"></div>'

        for day_num in range(1, last_day + 1):
            d = date(year_m, month_m, day_num)
            col = col_of(d.weekday())
            is_today = (d == today)
            is_sun = is_sun_col(col)
            is_sat = is_sat_col(col)
            cows_on_day = sorted(calving_data.get(d, []), key=lambda c: c["cow_id"])

            cell_cls = "today-cell" if is_today else ("sun-cell" if is_sun else ("sat-cell" if is_sat else ""))
            if cows_on_day:
                cell_cls = (cell_cls + " has-calving").strip()
            day_cls = "today-day" if is_today else ("sun-day" if is_sun else ("sat-day" if is_sat else ""))

            body += f'<div class="cal-cell {cell_cls}">'
            body += f'<div class="cal-day {day_cls}">{day_num}</div>'

            for c in cows_on_day[:DISPLAY_MAX]:
                chip_cls = "cow-chip has-twin" if c["twin"] else "cow-chip"
                lact_str = f'{c["lact"]}産' if c["lact"] else ""
                sex_str = _calving_calendar_sex_suffix(c["sex"], c["sire"], sire_list)
                sub_parts = [s for s in [lact_str, sex_str] if s]
                sub_str  = " / ".join(sub_parts)
                sire_str = c["sire"] if c["sire"] else ""
                twin_badge = '<span class="twin-warn">⚠️双子</span>' if c["twin"] else ""

                # Google Calendar リンク生成
                _gcal_summary = f"[分娩予定] {c['cow_id']}"
                _gcal_desc_parts = [f"農場: {farm_name}"]
                if sub_parts: _gcal_desc_parts.append(" / ".join(sub_parts))
                if c["twin"]:  _gcal_desc_parts.append("双子")
                if sire_str:   _gcal_desc_parts.append(f"SIRE: {sire_str}")
                _gcal_url = _google_cal_url(_gcal_summary, d, " / ".join(_gcal_desc_parts))

                body += (
                    f'<a class="chip-link" href="{_esc(_gcal_url)}" target="_blank"'
                    f' title="Googleカレンダーに追加: {_esc(_gcal_summary)}">'
                )
                body += f'<div class="{chip_cls}">'
                body += f'<span class="chip-id">{_esc(c["cow_id"])}</span>{twin_badge} '
                body += f'<span class="chip-sub">{_esc(sub_str)}</span>'
                if sire_str:
                    body += f'<span class="chip-sire">{_esc(sire_str)}</span>'
                body += '</div></a>'

            if len(cows_on_day) > DISPLAY_MAX:
                body += f'<span class="more-badge">+{len(cows_on_day) - DISPLAY_MAX}頭</span>'
            if cows_on_day:
                body += f'<div class="cow-count">計{len(cows_on_day)}頭</div>'

            body += '</div>'  # cal-cell

        last_col = col_of(date(year_m, month_m, last_day).weekday())
        for _ in range(6 - last_col):
            body += '<div class="cal-cell empty"></div>'

        body += '</div></div>'  # cal-grid / month-block

        month_m += 1
        if month_m > 12:
            month_m = 1
            year_m += 1

    body += "</div></body></html>"
    return body
