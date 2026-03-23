"""
乳検レポート用：月次トレンド（直近5か月）とコメントモード（DairyNotebookMT index.html 準拠）。
"""
from __future__ import annotations

import html as html_module
import re
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple


def _parse_event_json(raw: Any) -> Dict[str, Any]:
    if raw is None:
        return {}
    if isinstance(raw, dict):
        return raw
    if isinstance(raw, str):
        try:
            import json as _json

            return _json.loads(raw)
        except Exception:
            return {}
    return {}


def _months_ending_at(year: int, month: int, count: int = 5) -> List[Tuple[int, int]]:
    cy, cm = year, month
    rev: List[Tuple[int, int]] = []
    for _ in range(count):
        rev.append((cy, cm))
        if cm == 1:
            cy -= 1
            cm = 12
        else:
            cm -= 1
    rev.reverse()
    return rev


def compute_monthly_milk_trend(
    events: List[Dict[str, Any]], selected_date: str, num_months: int = 5
) -> List[Dict[str, Any]]:
    """
    各暦月について「その月内で最も新しい検定日」のイベントのみ集計（既存の前月集計と同様）。
    戻り値は時系列順。欠月は数値が None。
    """
    try:
        end_dt = datetime.strptime(selected_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        return []

    rows: List[Dict[str, Any]] = []
    for y, m in _months_ending_at(end_dt.year, end_dt.month, num_months):
        prefix = f"{y:04d}-{m:02d}"
        dates_in_month = sorted(
            {
                e.get("event_date")
                for e in events
                if e.get("event_date") and str(e.get("event_date", "")).startswith(prefix)
            },
            reverse=True,
        )
        if not dates_in_month:
            rows.append(
                {
                    "label": f"{y}/{m:02d}",
                    "ym": prefix,
                    "test_date": None,
                    "avg_milk": None,
                    "avg_scc_man": None,
                    "avg_ls": None,
                    "avg_mun": None,
                    "avg_denovo_fa": None,
                    "head_count": None,
                }
            )
            continue
        latest = dates_in_month[0]
        month_events = [e for e in events if e.get("event_date") == latest]
        milk_vals: List[float] = []
        scc_vals: List[float] = []
        ls_vals: List[float] = []
        mun_vals: List[float] = []
        denovo_vals: List[float] = []
        for e in month_events:
            j = _parse_event_json(e.get("json_data"))
            mv = j.get("milk_yield")
            if mv not in ("", None):
                try:
                    milk_vals.append(float(mv))
                except (ValueError, TypeError):
                    pass
            sv = j.get("scc")
            if sv not in ("", None):
                try:
                    sn = float(sv)
                    if sn > 0:
                        scc_vals.append(sn)
                except (ValueError, TypeError):
                    pass
            lv = j.get("ls")
            if lv not in ("", None):
                try:
                    ls_vals.append(float(lv))
                except (ValueError, TypeError):
                    pass
            mun_v = j.get("mun")
            if mun_v not in ("", None):
                try:
                    mun_vals.append(float(mun_v))
                except (ValueError, TypeError):
                    pass
            dv = j.get("denovo_fa")
            if dv not in ("", None):
                try:
                    denovo_vals.append(float(dv))
                except (ValueError, TypeError):
                    pass
        rows.append(
            {
                "label": f"{y}/{m:02d}",
                "ym": prefix,
                "test_date": latest,
                "avg_milk": round(sum(milk_vals) / len(milk_vals), 1) if milk_vals else None,
                "avg_scc_man": round((sum(scc_vals) / len(scc_vals)) / 10.0, 1) if scc_vals else None,
                "avg_ls": round(sum(ls_vals) / len(ls_vals), 1) if ls_vals else None,
                "avg_mun": round(sum(mun_vals) / len(mun_vals), 1) if mun_vals else None,
                "avg_denovo_fa": round(sum(denovo_vals) / len(denovo_vals), 1) if denovo_vals else None,
                "head_count": len(month_events),
            }
        )
    return rows


def _scale_y(values: List[Optional[float]], pad_ratio: float = 0.08) -> Tuple[float, float]:
    nums = [v for v in values if v is not None]
    if not nums:
        return 0.0, 1.0
    lo, hi = min(nums), max(nums)
    if lo == hi:
        lo -= 0.5
        hi += 0.5
    span = hi - lo
    pad = span * pad_ratio
    return lo - pad, hi + pad


def _svg_line_chart(
    title: str,
    labels: List[str],
    values: List[Optional[float]],
    stroke: str,
    fmt_label: Callable[[float], str],
) -> str:
    w, h = 280, 188
    pad_l, pad_r, pad_t, pad_b = 38, 10, 28, 34
    iw = w - pad_l - pad_r
    ih = h - pad_t - pad_b
    n = max(len(labels), 1)
    y0, y1 = _scale_y(values)

    def x_at(i: int) -> float:
        if n <= 1:
            return pad_l + iw / 2
        return pad_l + (i / (n - 1)) * iw

    def y_at(v: float) -> float:
        if y1 == y0:
            return pad_t + ih / 2
        return pad_t + ih * (1 - (v - y0) / (y1 - y0))

    parts: List[str] = [
        f'<div class="milk-trend-chart-card">',
        f'<div class="milk-trend-chart-title">{html_module.escape(title)}</div>',
        f'<svg class="milk-trend-svg" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg" '
        f'aria-label="{html_module.escape(title)}">',
        f'<rect x="0" y="0" width="{w}" height="{h}" fill="transparent"/>',
    ]

    # 軸ラベル（月）
    for i, lb in enumerate(labels):
        cx = x_at(i)
        parts.append(
            f'<text x="{cx:.1f}" y="{h - 8}" text-anchor="middle" class="milk-trend-axis-label">'
            f"{html_module.escape(lb)}</text>"
        )

    # 折れ線セグメント
    segments: List[List[Tuple[float, float, float]]] = []
    cur: List[Tuple[float, float, float]] = []
    for i, v in enumerate(values):
        if v is None:
            if cur:
                segments.append(cur)
                cur = []
            continue
        cur.append((float(i), x_at(i), float(v)))
    if cur:
        segments.append(cur)

    for seg in segments:
        if len(seg) == 1:
            i, xv, yv = seg[0]
            yp = y_at(yv)
            parts.append(
                f'<circle cx="{xv:.1f}" cy="{yp:.1f}" r="4" fill="{stroke}" '
                f'stroke="#fff" stroke-width="1.5"/>'
            )
            txt = fmt_label(yv)
            parts.append(
                f'<text x="{xv:.1f}" y="{yp - 10:.1f}" text-anchor="middle" '
                f'class="milk-trend-value-label" fill="{stroke}">{html_module.escape(txt)}</text>'
            )
            continue
        d_parts: List[str] = []
        for idx, (i, xv, yv) in enumerate(seg):
            yp = y_at(yv)
            cmd = "M" if idx == 0 else "L"
            d_parts.append(f"{cmd}{xv:.1f},{yp:.1f}")
        parts.append(
            f'<path d="{" ".join(d_parts)}" fill="none" stroke="{stroke}" stroke-width="2.2" '
            f'stroke-linecap="round" stroke-linejoin="round"/>'
        )
        for i, xv, yv in seg:
            yp = y_at(yv)
            parts.append(
                f'<circle cx="{xv:.1f}" cy="{yp:.1f}" r="4" fill="{stroke}" '
                f'stroke="#fff" stroke-width="1.5"/>'
            )
            txt = fmt_label(yv)
            parts.append(
                f'<text x="{xv:.1f}" y="{yp - 10:.1f}" text-anchor="middle" '
                f'class="milk-trend-value-label" fill="{stroke}">{html_module.escape(txt)}</text>'
            )

    parts.append("</svg></div>")
    return "".join(parts)


def build_monthly_trend_section_html(trend_rows: List[Dict[str, Any]]) -> str:
    if not trend_rows:
        return ""
    labels = [r["label"] for r in trend_rows]
    milk = [r.get("avg_milk") for r in trend_rows]
    scc = [r.get("avg_scc_man") for r in trend_rows]
    ls = [r.get("avg_ls") for r in trend_rows]
    heads = [float(r["head_count"]) if r.get("head_count") is not None else None for r in trend_rows]
    mun = [r.get("avg_mun") for r in trend_rows]
    denovo = [r.get("avg_denovo_fa") for r in trend_rows]

    c_milk = "#1976d2"
    c_scc = "#e53935"
    c_ls = "#fb8c00"
    c_head = "#43a047"
    c_mun = "#8e24aa"
    c_denovo = "#00838f"

    grid = (
        '<div class="milk-trend-grid">'
        + _svg_line_chart("平均乳量 (kg)", labels, milk, c_milk, lambda v: f"{v:.1f}kg")
        + _svg_line_chart("平均体細胞 (万/mL)", labels, scc, c_scc, lambda v: f"{v:.1f}万")
        + _svg_line_chart("リニアスコア平均", labels, ls, c_ls, lambda v: f"{v:.1f}")
        + _svg_line_chart("検定頭数", labels, heads, c_head, lambda v: f"{v:.0f}頭")
        + _svg_line_chart("平均MUN (mg/dL)", labels, mun, c_mun, lambda v: f"{v:.1f}")
        + _svg_line_chart("平均デノボFA (%)", labels, denovo, c_denovo, lambda v: f"{v:.1f}%")
        + "</div>"
    )

    note = (
        '<p class="milk-trend-note subnote">各月は「その月のうち最も新しい検定日」の集計です。'
        "体細胞は千単位のデータを10で割り万/mL相当で表示しています（概要表の千表示と整合）。"
        "MUN・デノボFAは乳検イベントに値がある頭のみで平均しています（欠測月は線が途切れます）。</p>"
    )

    return (
        '<div id="s-trend" class="report-layout page-break-before page-break-after milk-trend-section">'
        '<div style="width:100%">'
        '<div class="milk-trend-heading">月次トレンド</div>'
        '<div class="rsect-sub">月次トレンド（5ヶ月分）</div>'
        f"{grid}{note}"
        "</div></div>"
    )


MILK_REPORT_COMMENT_CSS = """
        .milk-report-toolbar { background:#fff; border:1px solid #e0e0e0; border-radius:10px; padding:12px 16px;
          margin-bottom:16px; display:flex; align-items:center; gap:10px; flex-wrap:wrap;
          position:sticky; top:0; z-index:40; box-shadow:0 2px 8px rgba(0,0,0,.06); }
        .milk-report-toolbar .btn-ghost { padding:7px 14px; border:1px solid #dadce0; border-radius:4px;
          background:#fff; cursor:pointer; font-size:13px; color:#3c4043; }
        .milk-report-toolbar .btn-ghost:hover { background:#f8f9fa; }
        .btn-comment-mode.active { background:#fff0f0; color:#cc0000; border-color:#cc0000; font-weight:600; }
        .comment-zone { margin:.1rem 0; min-height:0; }
        .comment-zone.has-comment { margin:.25rem 0; }
        .comment-add-btn { display:none; width:100%; padding:.45rem 1rem; border:2px dashed #e0c0c0; border-radius:6px;
          background:rgba(255,240,240,.4); color:#c08080; font-size:.85rem; cursor:pointer; text-align:left; }
        .comment-add-btn:hover { border-color:#cc0000; color:#cc0000; background:rgba(255,230,230,.6); }
        .comment-mode-active .comment-add-btn { display:block; }
        .comment-text-block { padding:.3rem .75rem; position:relative; }
        .comment-text { font-family:'Yomogi','Hachi Maru Pop','Comic Sans MS',cursive; color:#cc0000; font-size:.95rem;
          line-height:1.75; white-space:pre-wrap; display:inline-block; transform:rotate(-.4deg); margin:0; padding:0;
          filter:drop-shadow(.5px .5px 0 rgba(160,0,0,.12)); letter-spacing:.01em; }
        .comment-actions { display:none; gap:.4rem; margin-top:.3rem; align-items:center; }
        .comment-mode-active .comment-actions { display:flex; }
        .btn-comment-edit,.btn-comment-delete { font-size:.75rem; padding:.15rem .55rem; border-radius:3px; cursor:pointer;
          background:#fff; line-height:1.4; }
        .btn-comment-edit { color:#2563eb; border:1px solid #93c5fd; }
        .btn-comment-edit:hover { background:#eff6ff; }
        .btn-comment-delete { color:#dc2626; border:1px solid #fca5a5; }
        .btn-comment-delete:hover { background:#fef2f2; }
        #commentEditorDiv { width:100%; min-height:100px; padding:.6rem .75rem; border:1.5px solid #ddd; border-radius:6px;
          font-size:.95rem; line-height:1.65; outline:none; background:#fff; box-sizing:border-box;
          font-family:'Yomogi',cursive; color:#cc0000; }
        #commentEditorDiv:focus { border-color:#cc0000; box-shadow:0 0 0 2px rgba(204,0,0,.1); }
        #commentEditorDiv:empty::before { content:attr(data-placeholder); color:#aaa; pointer-events:none; font-style:italic; }
        #commentEditorDiv span[style*="background-color"] { border-radius:2px; padding:0 1px; }
        .milk-trend-heading { font-size:16px; font-weight:700; color:#1a237e; margin:0 0 6px; padding-bottom:8px;
          border-bottom:2px solid #e8eaf0; }
        .rsect-sub { font-size:12px; font-weight:600; color:#5f6368; margin:0 0 8px; }
        .milk-trend-section { background:#fafbfc; }
        .milk-trend-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:4px; }
        .milk-trend-chart-card { background:#fff; border-radius:10px; border:1px solid #e0e0e0; padding:14px 12px 8px; }
        .milk-trend-chart-title { font-size:12px; font-weight:600; color:#5f6368; margin-bottom:6px; }
        .milk-trend-svg { width:100%; height:auto; display:block; max-height:220px; }
        .milk-trend-axis-label { fill:#9aa0a6; font-size:10px; font-family:inherit; }
        .milk-trend-value-label { font-size:10px; font-weight:600; font-family:inherit; }
        .milk-trend-note { margin-top:8px; margin-bottom:0; font-size:10px; line-height:1.45; }
        @media (max-width:700px) { .milk-trend-grid { grid-template-columns:1fr; } }
        @media print {
          .milk-report-toolbar { display:none !important; }
          .comment-add-btn,.comment-actions { display:none !important; }
          .comment-text-block { background:transparent; padding:.2rem .5rem; }
          .comment-text { font-size:.95rem !important; }
          .comment-text span[style*="background-color"] { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
          .milk-trend-section { background:transparent !important; }
          .milk-trend-heading { margin-bottom:4px; padding-bottom:4px; font-size:14px; }
          .milk-trend-grid { gap:8px !important; }
          .milk-trend-chart-card { padding:6px 6px 2px !important; }
          .milk-trend-svg { max-height:180px !important; }
          .milk-trend-note { margin-top:6px !important; }
        }
"""


def build_comment_modal_html() -> str:
    return """
<div id="commentEditorModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;align-items:center;justify-content:center">
  <div style="background:#fff;border-radius:10px;padding:0;max-width:500px;width:92%;box-shadow:0 8px 32px rgba(0,0,0,.2);overflow:hidden" onclick="event.stopPropagation()">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid #f1f3f4">
      <h3 style="font-size:15px;font-weight:600;color:#3c4043;margin:0">コメントを編集</h3>
      <button type="button" onclick="closeCommentEditor()" style="background:none;border:none;font-size:20px;cursor:pointer;color:#5f6368;line-height:1;padding:0 4px">×</button>
    </div>
    <div style="padding:18px">
      <p style="font-size:.85rem;color:#555;margin-bottom:.75rem">入力したコメントがレポートの該当箇所に赤ペン風で表示されます。空欄で保存するとコメントを削除します。</p>
      <div style="margin-bottom:.75rem;display:flex;align-items:center;gap:.5rem;flex-wrap:wrap">
        <label for="commentFontSelect" style="font-size:.85rem;color:#555;white-space:nowrap">フォント：</label>
        <select id="commentFontSelect" style="font-size:.9rem;padding:.2rem .4rem;border-radius:4px;border:1px solid #ccc;flex:1;min-width:140px">
          <option value="Yomogi">よもぎ (Yomogi)</option>
          <option value="Hachi Maru Pop">はちまるポップ</option>
          <option value="Klee One">クレーワン</option>
          <option value="Zen Kurenaido">禅紅梅</option>
          <option value="Yuji Syuku">勇字塾</option>
        </select>
        <label for="commentFontSizeSelect" style="font-size:.85rem;color:#555;white-space:nowrap">サイズ：</label>
        <select id="commentFontSizeSelect" style="font-size:.9rem;padding:.2rem .4rem;border-radius:4px;border:1px solid #ccc">
          <option value="0.75rem">極小</option>
          <option value="0.85rem">小</option>
          <option value="0.95rem" selected>中</option>
          <option value="1.1rem">大</option>
          <option value="1.25rem">極大</option>
        </select>
      </div>
      <div id="commentHighlightToolbar" style="margin-bottom:.5rem;display:flex;align-items:center;gap:.4rem;flex-wrap:wrap">
        <span style="font-size:.8rem;color:#555;white-space:nowrap">蛍光ペン：</span>
        <button type="button" class="btn-highlight" data-color="#FFFF66" title="黄色" style="background:#FFFF66;border:1px solid #bbb;border-radius:3px;width:24px;height:24px;cursor:pointer;padding:0"></button>
        <button type="button" class="btn-highlight" data-color="#66FF99" title="緑" style="background:#66FF99;border:1px solid #bbb;border-radius:3px;width:24px;height:24px;cursor:pointer;padding:0"></button>
        <button type="button" class="btn-highlight" data-color="#FF99CC" title="ピンク" style="background:#FF99CC;border:1px solid #bbb;border-radius:3px;width:24px;height:24px;cursor:pointer;padding:0"></button>
        <button type="button" class="btn-highlight" data-color="#99CCFF" title="水色" style="background:#99CCFF;border:1px solid #bbb;border-radius:3px;width:24px;height:24px;cursor:pointer;padding:0"></button>
        <button type="button" class="btn-highlight" data-color="#FFCC66" title="オレンジ" style="background:#FFCC66;border:1px solid #bbb;border-radius:3px;width:24px;height:24px;cursor:pointer;padding:0"></button>
        <button type="button" id="btnRemoveHighlight" style="background:#fff;border:1px solid #ccc;border-radius:3px;padding:0 7px;height:24px;cursor:pointer;font-size:.75rem;color:#555">消去</button>
      </div>
      <div id="commentEditorDiv" contenteditable="true" role="textbox" aria-multiline="true" aria-label="コメント入力" data-placeholder="コメントを入力..."></div>
      <div style="display:flex;gap:8px;margin-top:1rem">
        <button type="button" onclick="saveCommentFromEditor()" class="btn-primary" style="padding:8px 16px;border:none;border-radius:6px;background:#1976d2;color:#fff;cursor:pointer;font-size:13px;font-weight:500">保存</button>
        <button type="button" onclick="closeCommentEditor()" style="padding:8px 16px;border:1px solid #1976d2;border-radius:6px;background:#fff;color:#1976d2;cursor:pointer;font-size:13px;font-weight:500">キャンセル</button>
      </div>
    </div>
  </div>
</div>
"""


def build_comment_script(storage_key_js: str) -> str:
    """storage_key_js は json.dumps でエスケープ済みの JS 文字列リテラル。"""
    # pylint: disable=anomalous-backslash-in-string
    return f"""
<script>
(function() {{
'use strict';
var COMMENT_STORAGE_KEY = {storage_key_js};
var reportComments = {{}};
var commentFont = 'Yomogi';
var commentFontSize = '0.95rem';
var commentModeActive = false;
var commentEditorTarget = null;
var HIGHLIGHT_COLORS = ['#FFFF66','#66FF99','#FF99CC','#99CCFF','#FFCC66'];

function rgbToHex(rgb) {{
  if (!rgb) return null;
  var s = rgb.trim();
  if (/^#[0-9a-fA-F]{{3,6}}$/.test(s)) return s.toUpperCase();
  var m = s.match(/^rgb\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\s*\\)$/);
  if (!m) return null;
  return '#' + [m[1],m[2],m[3]].map(function(n) {{
    return ('0' + parseInt(n, 10).toString(16)).slice(-2);
  }}).join('');
}}
function escapeHtmlComment(s) {{
  var d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}}
function sanitizeCommentHtml(html) {{
  var tmp = document.createElement('div');
  tmp.innerHTML = html;
  function processNode(node) {{
    if (node.nodeType === Node.TEXT_NODE) return node.cloneNode(false);
    if (node.nodeType !== Node.ELEMENT_NODE) return null;
    var tag = node.tagName.toUpperCase();
    if (tag === 'BR') return document.createElement('br');
    if (tag === 'SPAN') {{
      var hex = rgbToHex(node.style.backgroundColor);
      if (hex && HIGHLIGHT_COLORS.indexOf(hex) !== -1) {{
        var span = document.createElement('span');
        span.style.backgroundColor = hex;
        for (var i = 0; i < node.childNodes.length; i++) {{
          var c = processNode(node.childNodes[i]);
          if (c) span.appendChild(c);
        }}
        return span;
      }}
      var frag = document.createDocumentFragment();
      for (var i = 0; i < node.childNodes.length; i++) {{
        var c = processNode(node.childNodes[i]);
        if (c) frag.appendChild(c);
      }}
      return frag;
    }}
    var BLOCK = ['DIV','P','LI','H1','H2','H3','H4','H5','H6','BLOCKQUOTE'];
    var frag = document.createDocumentFragment();
    for (var i = 0; i < node.childNodes.length; i++) {{
      var c = processNode(node.childNodes[i]);
      if (c) frag.appendChild(c);
    }}
    if (BLOCK.indexOf(tag) !== -1) frag.appendChild(document.createElement('br'));
    return frag;
  }}
  var out = document.createElement('div');
  for (var i = 0; i < tmp.childNodes.length; i++) {{
    var r = processNode(tmp.childNodes[i]);
    if (r) out.appendChild(r);
  }}
  while (out.lastChild && out.lastChild.nodeName === 'BR') out.removeChild(out.lastChild);
  return out.innerHTML;
}}
function loadCommentIntoEditor(divEl, stored) {{
  if (!stored) {{ divEl.innerHTML = ''; return; }}
  if (stored.indexOf('<') !== -1) {{
    divEl.innerHTML = sanitizeCommentHtml(stored);
  }} else {{
    divEl.innerHTML = escapeHtmlComment(stored).replace(/\\n/g, '<br>');
  }}
}}
function applyHighlight(color) {{
  var editorDiv = document.getElementById('commentEditorDiv');
  if (!editorDiv) return;
  editorDiv.focus();
  var sel = window.getSelection();
  if (editorDiv._savedRange) {{
    sel.removeAllRanges();
    sel.addRange(editorDiv._savedRange);
    editorDiv._savedRange = null;
  }}
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
  var range = sel.getRangeAt(0);
  if (!editorDiv.contains(range.commonAncestorContainer)) return;
  try {{
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('backColor', false, color);
  }} catch (e) {{}}
}}
function removeHighlight() {{
  var editorDiv = document.getElementById('commentEditorDiv');
  if (!editorDiv) return;
  editorDiv.focus();
  var sel = window.getSelection();
  if (editorDiv._savedRange) {{
    sel.removeAllRanges();
    sel.addRange(editorDiv._savedRange);
    editorDiv._savedRange = null;
  }}
  var spans = editorDiv.querySelectorAll('span[style*="background-color"]');
  var range = (sel && sel.rangeCount > 0) ? sel.getRangeAt(0) : null;
  spans.forEach(function(span) {{
    try {{
      if (!range || range.isCollapsed || range.intersectsNode(span)) {{
        var parent = span.parentNode;
        if (!parent) return;
        while (span.firstChild) parent.insertBefore(span.firstChild, span);
        parent.removeChild(span);
        parent.normalize();
      }}
    }} catch(e) {{}}
  }});
}}

var COMMENT_ZONES = [
  {{ after: 's-overview', zone: 'overview', label: '概要後' }},
  {{ after: 's-milkquality', zone: 'milkquality', label: '乳質後' }},
  {{ after: 's-firstparity', zone: 'firstparity', label: '初産成績後' }},
  {{ after: 's-multiparity', zone: 'multiparity', label: '2産以上後' }},
  {{ after: 's-anomaly', zone: 'anomaly', label: '異常牛検知後' }},
  {{ after: 's-trend', zone: 'trend', label: 'トレンド後' }},
  {{ after: 's-cowlist', zone: 'cowlist', label: '一覧アンカー後' }}
];

function persistComments() {{
  try {{
    localStorage.setItem(COMMENT_STORAGE_KEY, JSON.stringify(reportComments));
  }} catch (e) {{}}
}}
function loadStoredComments() {{
  try {{
    var raw = localStorage.getItem(COMMENT_STORAGE_KEY);
    if (raw) {{
      var o = JSON.parse(raw);
      if (o && typeof o === 'object') reportComments = o;
    }}
  }} catch (e) {{}}
}}

function insertCommentZones() {{
  var existing = document.querySelectorAll('#main-content .comment-zone');
  for (var i = 0; i < existing.length; i++) {{
    if (existing[i].parentNode) existing[i].parentNode.removeChild(existing[i]);
  }}
  COMMENT_ZONES.forEach(function(s) {{
    var sectionEl = document.getElementById(s.after);
    if (!sectionEl || !sectionEl.parentNode) return;
    var zone = document.createElement('div');
    zone.className = 'comment-zone';
    zone.setAttribute('data-zone', s.zone);
    sectionEl.parentNode.insertBefore(zone, sectionEl.nextSibling);
  }});
  var mc = document.getElementById('main-content');
  if (mc) {{
    if (commentModeActive) mc.classList.add('comment-mode-active');
    else mc.classList.remove('comment-mode-active');
  }}
  renderCommentZones();
}}

function renderCommentZones() {{
  var zones = document.querySelectorAll('#main-content .comment-zone');
  zones.forEach(function(zone) {{
    var zoneId = zone.getAttribute('data-zone');
    var comment = reportComments[zoneId] || '';
    zone.innerHTML = '';
    if (comment) {{
      zone.classList.add('has-comment');
      var textBlock = document.createElement('div');
      textBlock.className = 'comment-text-block';
      var textEl = document.createElement('p');
      textEl.className = 'comment-text';
      textEl.style.fontFamily = "'" + commentFont + "', cursive";
      textEl.style.fontSize = commentFontSize;
      if (comment.indexOf('<') !== -1) {{
        textEl.innerHTML = sanitizeCommentHtml(comment);
      }} else {{
        textEl.innerHTML = escapeHtmlComment(comment).replace(/\\n/g, '<br>');
      }}
      textBlock.appendChild(textEl);
      var actions = document.createElement('div');
      actions.className = 'comment-actions';
      var editBtn = document.createElement('button');
      editBtn.type = 'button';
      editBtn.className = 'btn-comment-edit';
      editBtn.textContent = '✏ 編集';
      (function(id) {{ editBtn.addEventListener('click', function() {{ openCommentEditor(id); }}); }})(zoneId);
      var delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn-comment-delete';
      delBtn.textContent = '✕ 削除';
      (function(id) {{ delBtn.addEventListener('click', function() {{ deleteComment(id); }}); }})(zoneId);
      actions.appendChild(editBtn);
      actions.appendChild(delBtn);
      textBlock.appendChild(actions);
      zone.appendChild(textBlock);
    }} else {{
      zone.classList.remove('has-comment');
      var addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'comment-add-btn';
      addBtn.textContent = '✏ ここにコメントを追加';
      (function(id) {{ addBtn.addEventListener('click', function() {{ openCommentEditor(id); }}); }})(zoneId);
      zone.appendChild(addBtn);
    }}
  }});
}}

function toggleCommentMode() {{
  commentModeActive = !commentModeActive;
  var mc = document.getElementById('main-content');
  if (mc) {{
    if (commentModeActive) mc.classList.add('comment-mode-active');
    else mc.classList.remove('comment-mode-active');
  }}
  var btn = document.querySelector('.btn-comment-mode');
  if (btn) {{
    if (commentModeActive) {{
      btn.classList.add('active');
      btn.textContent = '✏ コメントモード ON';
    }} else {{
      btn.classList.remove('active');
      btn.textContent = '✏ コメントモード';
    }}
  }}
}}

function openCommentEditor(zoneId) {{
  commentEditorTarget = zoneId;
  var editorDiv = document.getElementById('commentEditorDiv');
  if (editorDiv) {{
    loadCommentIntoEditor(editorDiv, reportComments[zoneId] || '');
    editorDiv.style.fontFamily = "'" + commentFont + "', cursive";
    editorDiv.style.fontSize = commentFontSize;
    editorDiv._savedRange = null;
  }}
  var fontSel = document.getElementById('commentFontSelect');
  var sizeSel = document.getElementById('commentFontSizeSelect');
  if (fontSel) fontSel.value = commentFont;
  if (sizeSel) sizeSel.value = commentFontSize;
  var modal = document.getElementById('commentEditorModal');
  if (modal) {{ modal.style.display = 'flex'; }}
  if (editorDiv) editorDiv.focus();
}}

function closeCommentEditor() {{
  commentEditorTarget = null;
  var modal = document.getElementById('commentEditorModal');
  if (modal) modal.style.display = 'none';
}}

function saveCommentFromEditor() {{
  if (!commentEditorTarget) return;
  var editorDiv = document.getElementById('commentEditorDiv');
  var html = editorDiv ? sanitizeCommentHtml(editorDiv.innerHTML) : '';
  var textOnly = editorDiv ? editorDiv.textContent.trim() : '';
  if (textOnly) {{
    reportComments[commentEditorTarget] = html;
  }} else {{
    delete reportComments[commentEditorTarget];
  }}
  closeCommentEditor();
  renderCommentZones();
  persistComments();
}}

function deleteComment(zoneId) {{
  delete reportComments[zoneId];
  renderCommentZones();
  persistComments();
}}

function initCommentEditorEvents() {{
  var modal = document.getElementById('commentEditorModal');
  if (modal) modal.addEventListener('click', function(e) {{ if (e.target === modal) closeCommentEditor(); }});

  var toolbar = document.getElementById('commentHighlightToolbar');
  if (toolbar) {{
    toolbar.addEventListener('mousedown', function(e) {{
      var btn = e.target.closest('.btn-highlight');
      if (btn) e.preventDefault();
    }});
    toolbar.addEventListener('click', function(e) {{
      var btn = e.target.closest('.btn-highlight');
      if (btn) applyHighlight(btn.getAttribute('data-color'));
    }});
  }}

  var removeBtn = document.getElementById('btnRemoveHighlight');
  if (removeBtn) {{
    removeBtn.addEventListener('mousedown', function(e) {{ e.preventDefault(); }});
    removeBtn.addEventListener('click', function() {{ removeHighlight(); }});
  }}

  var fontSel = document.getElementById('commentFontSelect');
  if (fontSel) fontSel.addEventListener('change', function() {{
    commentFont = fontSel.value;
    var editorDiv = document.getElementById('commentEditorDiv');
    if (editorDiv) editorDiv.style.fontFamily = "'" + commentFont + "', cursive";
    document.querySelectorAll('#main-content .comment-text').forEach(function(t) {{
      t.style.fontFamily = "'" + commentFont + "', cursive";
    }});
  }});

  var sizeSel = document.getElementById('commentFontSizeSelect');
  if (sizeSel) sizeSel.addEventListener('change', function() {{
    commentFontSize = sizeSel.value;
    var editorDiv = document.getElementById('commentEditorDiv');
    if (editorDiv) editorDiv.style.fontSize = commentFontSize;
    document.querySelectorAll('#main-content .comment-text').forEach(function(t) {{
      t.style.fontSize = commentFontSize;
    }});
  }});

  var editorDiv = document.getElementById('commentEditorDiv');
  if (editorDiv) {{
    editorDiv.addEventListener('keydown', function(e) {{
      if (e.key === 'Enter' && !e.shiftKey) {{
        e.preventDefault();
        document.execCommand('insertLineBreak');
      }}
    }});
    editorDiv.addEventListener('paste', function(e) {{
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData('text/plain');
      document.execCommand('insertText', false, text);
    }});
    editorDiv.addEventListener('blur', function() {{
      if (editorDiv.textContent.trim() === '') editorDiv.innerHTML = '';
    }});
    editorDiv.addEventListener('mouseup', function() {{
      var sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {{
        editorDiv._savedRange = sel.getRangeAt(0).cloneRange();
      }}
    }});
    editorDiv.addEventListener('keyup', function() {{
      var sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {{
        editorDiv._savedRange = sel.getRangeAt(0).cloneRange();
      }} else {{
        editorDiv._savedRange = null;
      }}
    }});
  }}
}}

function bootComments() {{
  loadStoredComments();
  initCommentEditorEvents();
  insertCommentZones();
}}
if (document.readyState === 'loading') {{
  document.addEventListener('DOMContentLoaded', bootComments);
}} else {{
  bootComments();
}}
window.toggleCommentMode = toggleCommentMode;
window.closeCommentEditor = closeCommentEditor;
window.saveCommentFromEditor = saveCommentFromEditor;
}})();
</script>
"""


def make_milk_report_comment_storage_key(farm_name: str, selected_date: str) -> str:
    safe_farm = re.sub(r"[^a-zA-Z0-9_.-]+", "_", str(farm_name))[:80]
    return f"falcon_milk_report_comments|{safe_farm}|{selected_date}"
