"""
FALCON2 - ゲノムレポート HTML 生成
PDF風レイアウト・セル単位の色分け（上位25％・下位25％）・散布図埋め込み。
散布図は Plotly.js でインタラクティブ表示（ホバーでID表示・複数グラフで同一IDハイライト）。
"""

import base64
import json
import io
import logging
from typing import Dict, Any, List, Optional
from datetime import datetime
import html as html_module


def _configure_matplotlib_japanese() -> None:
    """matplotlib で日本語表示用フォントを設定し、文字化け（□）を防ぐ。"""
    try:
        import matplotlib
        from matplotlib import font_manager
        candidates = ["MS Gothic", "MS PGothic", "Meiryo", "Yu Gothic", "MS UI Gothic", "Yu Gothic UI", "IPAexGothic", "IPAPGothic"]
        available = [f.name for f in font_manager.fontManager.ttflist]
        for name in candidates:
            if name in available:
                matplotlib.rcParams["font.family"] = name
                matplotlib.rcParams["axes.unicode_minus"] = False
                return
        matplotlib.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass


def _parse_date(s: Optional[str]) -> Optional[datetime]:
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    try:
        return datetime.strptime(s[:10].replace("/", "-"), "%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def _direction(key: str) -> str:
    from modules.genome_trait_descriptions import get_trait_description
    return get_trait_description(key).get("direction", "higher")


def _cell_class(val: Optional[float], p25: Optional[float], p75: Optional[float], key: str) -> str:
    if val is None or p25 is None or p75 is None:
        return ""
    d = _direction(key)
    if d == "neutral":
        return ""
    if d == "lower":
        if val <= p25:
            return "pct-good"
        if val >= p75:
            return "pct-warn"
    else:
        if val >= p75:
            return "pct-good"
        if val <= p25:
            return "pct-warn"
    return ""


# 共通で使う定数（genome_report_window と揃える）
COMPOSITE_INDEX_KEYS = ["GDWP$", "GNM$", "GTPI"]
# 表の「他のゲノム項目」列数。1～5つ選択可だが、表幅は常に5列で統一する
MAX_ADDITIONAL_DISPLAY = 5

# レポート共通スタイル（ゲノム・乳検で統一。フォント・配色・コンテナ）
REPORT_BASE_CSS = """
    * { box-sizing: border-box; }
    body { font-family: 'Meiryo', 'Yu Gothic', sans-serif; font-size: 12px; margin: 0; padding: 24px; background: #f8f9fa; color: #212529;
      -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .report-container { max-width: 1200px; margin: 0 auto; background: #fff; padding: 32px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    .report-header { border-bottom: 2px solid #0d6efd; padding-bottom: 16px; margin-bottom: 24px; }
    .report-header dl { display: grid; grid-template-columns: auto 1fr; gap: 8px 24px; margin: 0; }
    .report-header dt { font-weight: bold; color: #495057; }
    .report-header dd { margin: 0; }
    .report-brand { font-size: 0.85em; color: #6c757d; margin-top: 8px; }
    h2 { font-size: 1.1rem; margin: 28px 0 12px; color: #0d6efd; }
    h3 { font-size: 1rem; margin: 20px 0 10px; color: #212529; }
    .section-title { font-size: 1.1rem; margin: 28px 0 12px; color: #0d6efd; padding-bottom: 4px; border-bottom: 2px solid #0d6efd; }
    .subheader { font-size: 12px; color: #495057; margin-bottom: 12px; }
    .subnote { font-size: 11px; color: #6c757d; font-weight: normal; }
    .dashboard-table, .summary-table { width: 100%; border-collapse: collapse; margin-bottom: 24px; font-size: 11px; }
    .dashboard-table th, .dashboard-table td, .summary-table th, .summary-table td { border: 1px solid #dee2e6; padding: 8px 10px; text-align: left; }
    .dashboard-table th, .summary-table th { background: #e9ecef; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .summary-table .num { text-align: right; font-variant-numeric: tabular-nums; }
    .summary-table .data-row:nth-child(even) { background: #f8f9fa; }
    .summary-table td.id-pregnant { background: #d4edda !important; }
    .bhb-alert, .scc-alert { background: #fff3cd !important; }
    .scatter-box, .scatter-card { border: 1px solid #dee2e6; border-radius: 6px; padding: 12px; background: #fff; max-width: 100%; page-break-inside: avoid; }
    .scatter-box h4, .chart-title { margin: 0 0 8px; font-size: 0.95rem; font-weight: 600; color: #212529; }
    .plotly-scatter-div { min-height: 260px; width: 100%; }
    .chart-card, .donut-card { border: 1px solid #dee2e6; border-radius: 6px; padding: 12px; background: #fff; page-break-inside: avoid; }
    .print-section { page-break-inside: avoid; }
    @media print {
      body { background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .report-container { box-shadow: none; padding: 16px; }
      .dashboard-table th, .dashboard-table td, .summary-table th, .summary-table td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .summary-table .data-row:nth-child(even) { background: #f8f9fa !important; }
      .summary-table td.id-pregnant { background: #d4edda !important; }
      .scatter-box, .scatter-card, .chart-card, .donut-card { page-break-inside: avoid; }
      @page { size: A4 portrait; margin: 1cm; }
      body { font-size: 11px; }
    }
"""


def _get_genome_report_css() -> str:
    """ゲノムレポート専用の追加CSS（REPORT_BASE_CSS に上乗せ）。"""
    return REPORT_BASE_CSS + """
    .dashboard-table .metric-name { font-weight: 500; }
    .table-legend { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 6px; padding: 12px 16px; margin-bottom: 16px; font-size: 11px;
      -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .table-legend h4 { margin: 0 0 8px; font-size: 1em; }
    .table-legend ul { margin: 0; padding-left: 20px; }
    .table-legend .swatch { display: inline-block; width: 18px; height: 18px; vertical-align: middle; margin-right: 6px; border-radius: 3px;
      -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .swatch-good { background: #d4edda; border: 1px solid #28a745; }
    .swatch-warn { background: #fff3cd; border: 1px solid #ffc107; }
    .report-table { width: 100%; table-layout: fixed; border-collapse: collapse; margin-bottom: 32px; font-size: 10px; }
    .report-table th, .report-table td { border: 1px solid #dee2e6; padding: 4px 6px; text-align: center; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .report-table th.cell-full, .report-table td.cell-full { white-space: normal; overflow: visible; text-overflow: clip; word-break: break-all; }
    .report-table th { background: #e9ecef; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .report-table td.text-left { text-align: left; }
    .report-table .sticky-col { position: sticky; left: 0; background: #fff; z-index: 1; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .report-table th.sticky-col { background: #e9ecef; }
    .report-table tr:nth-child(even) td.sticky-col { background: #f8f9fa; }
    .report-table .pct-good { background: #d4edda !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .report-table .pct-warn { background: #fff3cd !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .report-table .col-1 { font-size: 9px; }
    .report-table .col-2, .report-table .col-3, .report-table .col-4, .report-table .col-5 { font-size: 8px; }
    .report-table .col-6, .report-table .col-7, .report-table .col-8 { font-size: 7px; }
    .report-table .col-9, .report-table .col-10, .report-table .col-11, .report-table .col-12, .report-table .col-13 { font-size: 8px; }
    .scatter-section { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-top: 16px; }
    .scatter-box img { max-width: 100%; height: auto; display: block; }
    .section-scatter { page-break-inside: avoid; }
    .report-table { page-break-inside: avoid; }
    h2 { page-break-after: avoid; }
    @media print {
      .report-table th, .report-table td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .report-table .pct-good, .report-table .pct-warn { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .report-table .sticky-col, .report-table th.sticky-col, .report-table tr:nth-child(even) td.sticky-col { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .report-table { font-size: 8px; }
      .report-table .col-1 { font-size: 8px; }
      .report-table .col-2, .report-table .col-3, .report-table .col-4, .report-table .col-5 { font-size: 7px; }
      .report-table .col-6, .report-table .col-7, .report-table .col-8 { font-size: 6px; }
      .report-table .col-9, .report-table .col-10, .report-table .col-11, .report-table .col-12, .report-table .col-13 { font-size: 7px; }
      .report-table th, .report-table td { padding: 2px 4px; }
    }
    """


def _scatter_by_year_png_base64(
    rows: List[Dict],
    metric_key: str,
    show_trend: bool,
    show_avg: bool,
    trait_label: str,
) -> str:
    """生年別散布図。X軸=生年（小数で実質生年月日）、Y軸=指標。近似曲線の式と R² を表示。"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import numpy as np
        from matplotlib.figure import Figure
    except ImportError:
        return ""

    _configure_matplotlib_japanese()

    x_years = []
    ys = []
    for r in rows:
        bthd = r.get("bthd")
        dt = _parse_date(bthd) if bthd else None
        y = r.get(metric_key)
        if dt is not None and y is not None:
            year_frac = dt.year + (dt - datetime(dt.year, 1, 1)).days / 365.25
            x_years.append(year_frac)
            ys.append(y)
    if len(x_years) < 2:
        return ""
    x_years = np.array(x_years)
    ys = np.array(ys, dtype=float)

    fig = Figure(figsize=(6, 3), dpi=100)
    ax = fig.add_subplot(111)
    ax.scatter(x_years, ys, alpha=0.7, s=24)
    ax.set_ylabel(trait_label, fontsize=10)
    ax.set_xlabel("生年", fontsize=10)
    ax.set_title(trait_label, fontsize=11)

    eq_text = ""
    if show_trend:
        z = np.polyfit(x_years, ys, 1)
        slope, intercept = z[0], z[1]
        p = np.poly1d(z)
        y_pred = p(x_years)
        ss_tot = np.sum((ys - np.mean(ys)) ** 2)
        ss_res = np.sum((ys - y_pred) ** 2)
        r2 = (1 - ss_res / ss_tot) if ss_tot > 0 else 0.0
        x_line = np.linspace(min(x_years), max(x_years), 100)
        ax.plot(x_line, p(x_line), "r-", alpha=0.8, label="近似曲線")
        if intercept >= 0:
            eq_str = f"y = {slope:.4f}x + {intercept:.4f}"
        else:
            eq_str = f"y = {slope:.4f}x - {abs(intercept):.4f}"
        eq_text = f"{eq_str}\nR² = {r2:.4f}"

    if show_avg and len(ys) > 0:
        mean_y = float(np.mean(ys))
        ax.axhline(y=mean_y, color="gray", linestyle="--", alpha=0.8, label="平均")

    ax.legend(loc="best", fontsize=8)
    if eq_text:
        ax.text(0.02, 0.98, eq_text, transform=ax.transAxes, fontsize=8, verticalalignment="top",
                bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.8), family="monospace")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _scatter_by_year_only_png_base64(
    rows: List[Dict],
    metric_key: str,
    show_trend: bool,
    show_avg: bool,
    trait_label: str,
) -> str:
    """年のみの散布図。X軸=生年（2024, 2025, 2026 の目盛）、Y軸=指標。"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import numpy as np
        from matplotlib.figure import Figure
    except ImportError:
        return ""

    _configure_matplotlib_japanese()

    x_years = []
    ys = []
    for r in rows:
        bthd = r.get("bthd")
        dt = _parse_date(bthd) if bthd else None
        y = r.get(metric_key)
        if dt is not None and y is not None:
            x_years.append(dt.year)
            ys.append(y)
    if len(x_years) < 2:
        return ""
    x_years = np.array(x_years, dtype=int)
    ys = np.array(ys, dtype=float)
    years_unique = np.unique(x_years)
    year_min, year_max = int(years_unique.min()), int(years_unique.max())

    fig = Figure(figsize=(6, 3), dpi=100)
    ax = fig.add_subplot(111)
    ax.scatter(x_years, ys, alpha=0.7, s=24)
    ax.set_ylabel(trait_label, fontsize=10)
    ax.set_xlabel("生年", fontsize=10)
    ax.set_title(trait_label, fontsize=11)
    ax.set_xticks(np.arange(year_min, year_max + 1))
    ax.set_xlim(year_min - 0.5, year_max + 0.5)

    eq_text = ""
    if show_trend and len(years_unique) >= 2:
        x_cont = x_years.astype(float)
        z = np.polyfit(x_cont, ys, 1)
        slope, intercept = z[0], z[1]
        p = np.poly1d(z)
        y_pred = p(x_cont)
        ss_tot = np.sum((ys - np.mean(ys)) ** 2)
        ss_res = np.sum((ys - y_pred) ** 2)
        r2 = (1 - ss_res / ss_tot) if ss_tot > 0 else 0.0
        x_line = np.linspace(year_min, year_max, 100)
        ax.plot(x_line, p(x_line), "r-", alpha=0.8, label="近似曲線")
        if intercept >= 0:
            eq_str = f"y = {slope:.4f}x + {intercept:.4f}"
        else:
            eq_str = f"y = {slope:.4f}x - {abs(intercept):.4f}"
        eq_text = f"{eq_str}\nR² = {r2:.4f}"

    if show_avg and len(ys) > 0:
        mean_y = float(np.mean(ys))
        ax.axhline(y=mean_y, color="gray", linestyle="--", alpha=0.8, label="平均")

    # 年ごとの平均をつなぐ線（複数年で近似曲線・全体平均と区別できる）
    if len(years_unique) >= 1:
        year_means = [float(np.mean(ys[x_years == y])) for y in years_unique]
        ax.plot(
            years_unique,
            year_means,
            color="green",
            linestyle="-",
            linewidth=2,
            marker="o",
            markersize=6,
            alpha=0.9,
            label="年の平均",
        )

    ax.legend(loc="best", fontsize=8)
    if eq_text:
        ax.text(0.02, 0.98, eq_text, transform=ax.transAxes, fontsize=8, verticalalignment="top",
                bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.8), family="monospace")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _scatter_by_year_plotly_data(
    rows: List[Dict],
    metric_key: str,
    show_trend: bool,
    show_avg: bool,
    trait_label: str,
) -> Optional[Dict[str, Any]]:
    """生年別散布図用の Plotly データ。x, y, cow_ids, jpn10s, 近似曲線・平均・式を含む。"""
    x_list: List[float] = []
    y_list: List[float] = []
    cow_ids: List[str] = []
    jpn10s: List[str] = []
    for r in rows:
        bthd = r.get("bthd")
        dt = _parse_date(bthd) if bthd else None
        y = r.get(metric_key)
        if dt is not None and y is not None:
            year_frac = dt.year + (dt - datetime(dt.year, 1, 1)).days / 365.25
            x_list.append(year_frac)
            y_list.append(float(y))
            cow_ids.append(str(r.get("cow_id", "")))
            jpn10s.append(str(r.get("jpn10", "")))
    if len(x_list) < 2:
        return None
    try:
        import numpy as np
    except ImportError:
        return None
    x_arr = np.array(x_list)
    y_arr = np.array(y_list, dtype=float)
    result: Dict[str, Any] = {
        "title": trait_label,
        "xlabel": "生年",
        "ylabel": trait_label,
        "x": x_list,
        "y": y_list,
        "cow_ids": cow_ids,
        "jpn10s": jpn10s,
    }
    if show_avg and len(y_arr) > 0:
        result["avg_y"] = float(np.mean(y_arr))
    if show_trend:
        z = np.polyfit(x_arr, y_arr, 1)
        slope, intercept = float(z[0]), float(z[1])
        p = np.poly1d(z)
        y_pred = p(x_arr)
        ss_tot = np.sum((y_arr - np.mean(y_arr)) ** 2)
        ss_res = np.sum((y_arr - y_pred) ** 2)
        r2_val = (1 - ss_res / ss_tot) if ss_tot > 0 else 0.0
        x_line = np.linspace(min(x_arr), max(x_arr), 100)
        result["trend_x"] = x_line.tolist()
        result["trend_y"] = p(x_line).tolist()
        result["eq_str"] = f"y = {slope:.4f}x + {intercept:.4f}" if intercept >= 0 else f"y = {slope:.4f}x - {abs(intercept):.4f}"
        result["r2"] = r2_val
    return result


def _scatter_by_year_only_plotly_data(
    rows: List[Dict],
    metric_key: str,
    show_trend: bool,
    show_avg: bool,
    trait_label: str,
) -> Optional[Dict[str, Any]]:
    """年のみ散布図用の Plotly データ。"""
    x_list: List[int] = []
    y_list: List[float] = []
    cow_ids_list: List[str] = []
    jpn10s_list: List[str] = []
    for r in rows:
        bthd = r.get("bthd")
        dt = _parse_date(bthd) if bthd else None
        y = r.get(metric_key)
        if dt is not None and y is not None:
            x_list.append(dt.year)
            y_list.append(float(y))
            cow_ids_list.append(str(r.get("cow_id", "")))
            jpn10s_list.append(str(r.get("jpn10", "")))
    if len(x_list) < 2:
        return None
    try:
        import numpy as np
    except ImportError:
        return None
    x_arr = np.array(x_list, dtype=int)
    y_arr = np.array(y_list, dtype=float)
    years_unique = np.unique(x_arr)
    year_min, year_max = int(years_unique.min()), int(years_unique.max())
    result: Dict[str, Any] = {
        "title": trait_label,
        "xlabel": "生年",
        "ylabel": trait_label,
        "x": [int(a) for a in x_list],
        "y": y_list,
        "cow_ids": cow_ids_list,
        "jpn10s": jpn10s_list,
        "year_min": year_min,
        "year_max": year_max,
    }
    if show_avg and len(y_arr) > 0:
        result["avg_y"] = float(np.mean(y_arr))
    if show_trend and len(years_unique) >= 2:
        x_cont = x_arr.astype(float)
        z = np.polyfit(x_cont, y_arr, 1)
        slope, intercept = float(z[0]), float(z[1])
        p = np.poly1d(z)
        y_pred = p(x_cont)
        ss_tot = np.sum((y_arr - np.mean(y_arr)) ** 2)
        ss_res = np.sum((y_arr - y_pred) ** 2)
        r2_val = (1 - ss_res / ss_tot) if ss_tot > 0 else 0.0
        x_line = np.linspace(year_min, year_max, 100)
        result["trend_x"] = x_line.tolist()
        result["trend_y"] = p(x_line).tolist()
        result["eq_str"] = f"y = {slope:.4f}x + {intercept:.4f}" if intercept >= 0 else f"y = {slope:.4f}x - {abs(intercept):.4f}"
        result["r2"] = r2_val
    if len(years_unique) >= 1:
        result["year_means_x"] = [int(y) for y in years_unique]
        result["year_means_y"] = [float(np.mean(y_arr[x_arr == y])) for y in years_unique]
    return result


def _build_plotly_scatter_html(chart_configs: List[Dict[str, Any]]) -> str:
    """Plotly.js で散布図を描画。ホバーで動物ID表示・同一IDを全グラフでハイライト。"""
    if not chart_configs:
        return ""
    config_json = json.dumps(chart_configs, ensure_ascii=False).replace("</", "<\\/")
    divs_html = "\n".join(
        '<div class="scatter-box"><h4>' + html_module.escape(c["data"]["title"]) + '</h4><div id="' + html_module.escape(c["id"]) + '" class="plotly-scatter-div"></div></div>'
        for c in chart_configs
    )
    script = _get_plotly_scatter_script(config_json)
    return divs_html + "\n" + script


def _build_plotly_scatter_divs_only(chart_configs: List[Dict[str, Any]]) -> str:
    """散布図の div のみ返す（スクリプトなし）。"""
    if not chart_configs:
        return ""
    return "\n".join(
        '<div class="scatter-box"><h4>' + html_module.escape(c["data"]["title"]) + '</h4><div id="' + html_module.escape(c["id"]) + '" class="plotly-scatter-div"></div></div>'
        for c in chart_configs
    )


def _get_plotly_scatter_script(config_json: str) -> str:
    """散布図用の Plotly 初期化・ホバー連動の JS。config_json はエスケープ済みの JSON 文字列。"""
    return (
        r'<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>'
        "\n<script>\n(function() {\n  var configs = "
        + config_json
        + ";\n  if (!configs || !configs.length) return;\n"
        r"""  var plots = {};
  var DEFAULT_SIZE = 7, HIGHLIGHT_SIZE = 14, DEFAULT_OPACITY = 0.7, HIGHLIGHT_OPACITY = 1, DIM_OPACITY = 0.25;

  function buildTrace(d) {
    var customdata = (d.cow_ids || []).map(function(_, i) { return [d.cow_ids[i] || '', (d.jpn10s || [])[i] || '', d.y[i]]; });
    var trace = { x: d.x, y: d.y, mode: 'markers', type: 'scatter', name: 'データ', showlegend: false,
      marker: { size: DEFAULT_SIZE, color: '#1f77b4', opacity: DEFAULT_OPACITY },
      customdata: customdata,
      hovertemplate: (d.hovertemplate != null) ? d.hovertemplate : '<b>動物ID</b>: %{customdata[0]}<br><b>Official ID</b>: %{customdata[1]}<br><b>値</b>: %{customdata[2]:.2f}<extra></extra>' };
    if (d.parity && d.parity.length) {
      var parityColors = { 1: '#E91E63', 2: '#00ACC1', 3: '#26A69A' };
      trace.marker.color = d.parity.map(function(p) { return parityColors[p] || '#1f77b4'; });
    }
    return trace;
  }
  function buildLayout(d, isYearOnly) {
    var layout = { title: { text: d.title, font: { size: 14 } }, xaxis: { title: d.xlabel }, yaxis: { title: d.ylabel },
      margin: { t: 36, b: 40, l: 50, r: 24 }, showlegend: true, font: { family: 'Meiryo, Yu Gothic, sans-serif', size: 11 } };
    if (d.x_min != null && d.x_max != null) {
      layout.xaxis.range = [d.x_min, d.x_max];
      layout.xaxis.tick0 = d.x_min;
      layout.xaxis.dtick = (d.x_max - d.x_min) <= 100 ? 20 : 100;
    }
    else if (isYearOnly && d.year_min != null) { layout.xaxis.range = [d.year_min - 0.5, d.year_max + 0.5]; if (d.year_max - d.year_min <= 15) layout.xaxis.dtick = 1; }
    if (d.y_start_zero && d.y && d.y.length) { var maxY = Math.max.apply(null, d.y); layout.yaxis.range = [0, maxY * 1.05]; }
    return layout;
  }
  configs.forEach(function(cfg) {
    var d = cfg.data, id = cfg.id;
    var traces = [buildTrace(d)];
    if (d.trend_x && d.trend_y) traces.push({ x: d.trend_x, y: d.trend_y, mode: 'lines', type: 'scatter', name: '近似曲線', line: { color: '#d62728', width: 1.5 }, hoverinfo: 'skip' });
    if (d.avg_y != null) {
      var xr = d.x.length ? [Math.min.apply(null, d.x), Math.max.apply(null, d.x)] : [0, 1];
      if (d.year_min != null) xr = [d.year_min - 0.5, d.year_max + 0.5];
      traces.push({ x: xr, y: [d.avg_y, d.avg_y], mode: 'lines', type: 'scatter', name: '平均', line: { color: 'gray', dash: 'dash' }, hoverinfo: 'skip' });
    }
    if (d.year_means_x && d.year_means_y) traces.push({ x: d.year_means_x, y: d.year_means_y, mode: 'lines+markers', type: 'scatter', name: '年の平均', line: { color: 'green', width: 2 }, marker: { size: 8 }, hoverinfo: 'skip' });
    var layout = buildLayout(d, !!d.year_min);
    if (d.eq_str != null && d.r2 != null) layout.annotations = [{ x: 0.02, y: 0.98, xref: 'paper', yref: 'paper', text: d.eq_str + '\nR2 = ' + d.r2.toFixed(4), showarrow: false, font: { size: 9 }, bgcolor: 'rgba(245,222,179,0.9)' }];
    var gd = document.getElementById(id);
    if (!gd) return;
    Plotly.newPlot(id, traces, layout, { responsive: true, displayModeBar: true });
    plots[id] = { cowIds: d.cow_ids || [], element: gd };
  });
  var allPlotIds = configs.map(function(c) { return c.id; });
  function highlightCowId(cowId) {
    allPlotIds.forEach(function(plotId) {
      var p = plots[plotId]; if (!p || !p.cowIds.length) return;
      var sizes = p.cowIds.map(function(cid) { return cid === cowId ? HIGHLIGHT_SIZE : DEFAULT_SIZE; });
      var opacities = p.cowIds.map(function(cid) { return cid === cowId ? HIGHLIGHT_OPACITY : DIM_OPACITY; });
      Plotly.restyle(plotId, { 'marker.size': [sizes], 'marker.opacity': [opacities] }, [0]);
    });
  }
  function resetHighlight() {
    allPlotIds.forEach(function(plotId) {
      var p = plots[plotId]; if (!p || !p.cowIds.length) return;
      var n = p.cowIds.length;
      Plotly.restyle(plotId, { 'marker.size': [Array(n).fill(DEFAULT_SIZE)], 'marker.opacity': [Array(n).fill(DEFAULT_OPACITY)] }, [0]);
    });
  }
  allPlotIds.forEach(function(plotId) {
    document.getElementById(plotId).on('plotly_hover', function(event) {
      if (event.points && event.points[0] && event.points[0].data.customdata) {
        var pt = event.points[0]; var cowId = pt.data.customdata[pt.pointIndex] ? pt.data.customdata[pt.pointIndex][0] : null;
        if (cowId != null) highlightCowId(cowId);
      }
    });
    document.getElementById(plotId).on('plotly_unhover', function() { resetHighlight(); });
    if (typeof FALCON_OPEN_COW_PORT !== 'undefined') {
      document.getElementById(plotId).on('plotly_click', function(event) {
        if (!event.points || !event.points[0] || !event.points[0].data.customdata) return;
        var pt = event.points[0];
        var cowId = pt.data.customdata[pt.pointIndex] ? pt.data.customdata[pt.pointIndex][0] : null;
        if (!cowId) return;
        var url = 'http://127.0.0.1:' + FALCON_OPEN_COW_PORT + '/open_cow?cow_id=' + encodeURIComponent(cowId);
        fetch(url).catch(function() {});
      });
    }
  });
  if (typeof FALCON_OPEN_COW_PORT !== 'undefined') {
    function attachGenomeReportCowLinks() {
      var links = document.querySelectorAll('.report-cow-link');
      links.forEach(function(link) {
        link.addEventListener('click', function(ev) {
          ev.preventDefault();
          var cowId = link.getAttribute('data-cow-id');
          if (!cowId) return;
          var url = 'http://127.0.0.1:' + FALCON_OPEN_COW_PORT + '/open_cow?cow_id=' + encodeURIComponent(cowId);
          fetch(url).catch(function() {});
        });
      });
    }
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', attachGenomeReportCowLinks);
    } else {
      attachGenomeReportCowLinks();
    }
  }
})();
</script>"""
    )


def build_genome_report_html(
    farm_name: str,
    date_from: Optional[str],
    date_to: Optional[str],
    rows: List[Dict[str, Any]],
    stats: Dict[str, Dict[str, float]],
    composite_key: str,
    additional_keys: List[str],
    show_trend: bool = True,
    show_avg_line: bool = True,
    trait_display_names: Optional[Dict[str, str]] = None,
) -> str:
    """
    ゲノムレポートの HTML 文字列を生成する。
    セル単位で上位25％＝緑、下位25％＝黄の背景色を付与。
    表・散布図の項目名は略称（キー）のみ使用。
    """
    trait_display_names = trait_display_names or {}
    # 表ヘッダー・散布図タイトルは略称（キー）のみ
    def short_label(k: str) -> str:
        return k

    out_date = datetime.now().strftime("%Y年%m月%d日")
    period_text = "—"
    if rows:
        bthds = [_parse_date(r.get("bthd")) for r in rows if r.get("bthd")]
        bthds = [x for x in bthds if x is not None]
        if bthds:
            min_d = min(bthds).strftime("%Y/%m/%d")
            max_d = max(bthds).strftime("%Y/%m/%d")
            period_text = f"{min_d} ～ {max_d}"

    metric_keys = [composite_key] + list(additional_keys)

    # 表の幅を5列で統一するため、不足分を None でパディング
    additional_keys_padded: List[Optional[str]] = list(additional_keys) + [None] * (MAX_ADDITIONAL_DISPLAY - len(additional_keys))

    # 散布図（Plotly・ホバーでID表示・全グラフで同一IDハイライト連動）
    scatter_year_configs: List[Dict[str, Any]] = []
    for mk in metric_keys:
        data = _scatter_by_year_plotly_data(rows, mk, show_trend, show_avg_line, short_label(mk))
        if data:
            safe_id = mk.replace("$", "_").replace(" ", "-")
            scatter_year_configs.append({"id": "scatter-year-" + safe_id, "data": data})
    scatter_year_only_configs: List[Dict[str, Any]] = []
    for mk in metric_keys:
        data = _scatter_by_year_only_plotly_data(rows, mk, show_trend, show_avg_line, short_label(mk))
        if data:
            safe_id = mk.replace("$", "_").replace(" ", "-")
            scatter_year_only_configs.append({"id": "scatter-yearonly-" + safe_id, "data": data})
    all_scatter_configs = scatter_year_configs + scatter_year_only_configs
    scatter_year_divs = _build_plotly_scatter_divs_only(scatter_year_configs)
    scatter_year_only_divs = _build_plotly_scatter_divs_only(scatter_year_only_configs)
    scatter_script = _get_plotly_scatter_script(json.dumps(all_scatter_configs, ensure_ascii=False).replace("</", "<\\/")) if all_scatter_configs else ""


    # サマリ（略称で表示）
    summary_lines = []
    for k in metric_keys:
        s = stats.get(k) or {}
        mean = s.get("mean")
        mn = s.get("min")
        mx = s.get("max")
        mean_s = f"{mean:.2f}" if mean is not None else "—"
        min_s = f"{mn:.2f}" if mn is not None else "—"
        max_s = f"{mx:.2f}" if mx is not None else "—"
        summary_lines.append(f"<tr><td class=\"metric-name\">{html_module.escape(short_label(k))}</td><td>{mean_s}</td><td>{min_s}</td><td>{max_s}</td></tr>")
    summary_body = "\n".join(summary_lines)

    # 総合指標順位表（順位＝行番号＝総合指標の順位）
    sorted_by_index = sorted(rows, key=lambda r: (r.get(composite_key) is None, -(r.get(composite_key) or 0)))
    for idx, row in enumerate(sorted_by_index):
        row["_composite_rank"] = idx + 1
    table1_rows = _build_table_rows(sorted_by_index, stats, metric_keys, composite_key, additional_keys_padded, short_label, show_rank=True, rank_value=None)

    # 生年月日順表（順位＝総合指標の順位。行は生年月日新しい順なので順位はバラバラになる）
    sorted_by_bthd = sorted(rows, key=lambda r: (r.get("bthd") or ""), reverse=True)
    table2_rows = _build_table_rows(sorted_by_bthd, stats, metric_keys, composite_key, additional_keys_padded, short_label, show_rank=True, rank_value="_composite_rank")

    css = _get_genome_report_css()

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ゲノムレポート - {html_module.escape(farm_name or "FALCON")}</title>
<style>{css}</style>
</head>
<body>
<div class="report-container">
<header class="report-header">
  <dl>
    <dt>農場名</dt><dd>{html_module.escape(farm_name or "")}</dd>
    <dt>出力年月日</dt><dd>{out_date}</dd>
    <dt>生年月日対象期間</dt><dd>{html_module.escape(period_text)}</dd>
  </dl>
  <p class="report-brand">FALCON ゲノムレポート</p>
</header>

<div class="print-section">
<h2>牛群ダッシュボード</h2>
<p>対象頭数: {len(rows)} 頭</p>
<table class="dashboard-table">
<thead><tr><th>指標</th><th>平均</th><th>最小</th><th>最大</th></tr></thead>
<tbody>
{summary_body}
</tbody>
</table>

<div class="table-legend">
  <h4>表の色分けについて</h4>
  <p>各指標の<strong>セル</strong>は、この群れの中での相対的な位置で色分けしています。</p>
  <ul>
    <li><span class="swatch swatch-good"></span><strong>緑</strong>：上位25％</li>
    <li><span class="swatch swatch-warn"></span><strong>黄</strong>：下位25％</li>
  </ul>
</div>
</div>

<div class="print-section">
<h2>総合指標順位表</h2>
<p>総合指標（{html_module.escape(short_label(composite_key))}）の高い順</p>
<table class="report-table">
<colgroup>
<col style="width:4%"><col style="width:9%"><col style="width:12%"><col style="width:8%"><col style="width:11%"><col style="width:6%"><col style="width:6%"><col style="width:6%"><col style="width:8%"><col style="width:8%"><col style="width:8%"><col style="width:8%"><col style="width:8%">
</colgroup>
<thead>
<tr>
<th class="sticky-col col-1">順位</th><th class="sticky-col text-left cell-full col-2">動物ID</th><th class="cell-full col-3">Official ID</th><th class="cell-full col-4">父</th><th class="cell-full col-5">生年月日</th>
"""
    for i, k in enumerate(COMPOSITE_INDEX_KEYS):
        c = 6 + i
        html += f'<th class="col-{c}">{html_module.escape(short_label(k))}</th>'
    for i, k in enumerate(additional_keys_padded):
        c = 9 + i
        if k is None:
            html += f'<th class="col-{c}">—</th>'
        else:
            html += f'<th class="col-{c}">{html_module.escape(short_label(k))}</th>'
    html += """
</tr>
</thead>
<tbody>
"""
    html += table1_rows
    html += """
</tbody>
</table>
</div>

<div class="print-section">
<h2>生年月日順表</h2>
<p>生年月日の新しい順（※順位は総合指標の順位です）</p>
<table class="report-table">
<colgroup>
<col style="width:4%"><col style="width:9%"><col style="width:12%"><col style="width:8%"><col style="width:11%"><col style="width:6%"><col style="width:6%"><col style="width:6%"><col style="width:8%"><col style="width:8%"><col style="width:8%"><col style="width:8%"><col style="width:8%">
</colgroup>
<thead>
<tr>
<th class="sticky-col col-1">順位</th><th class="sticky-col text-left cell-full col-2">動物ID</th><th class="cell-full col-3">Official ID</th><th class="cell-full col-4">父</th><th class="cell-full col-5">生年月日</th>
"""
    for i, k in enumerate(COMPOSITE_INDEX_KEYS):
        c = 6 + i
        html += f'<th class="col-{c}">{html_module.escape(short_label(k))}</th>'
    for i, k in enumerate(additional_keys_padded):
        c = 9 + i
        if k is None:
            html += f'<th class="col-{c}">—</th>'
        else:
            html += f'<th class="col-{c}">{html_module.escape(short_label(k))}</th>'
    html += """
</tr>
</thead>
<tbody>
"""
    html += table2_rows
    html += """
</tbody>
</table>
</div>

<div class="print-section">
<h2>生年別 散布図</h2>
<p>X軸：生年（実質生年月日） / Y軸：各指標の値（近似曲線の式と R² を図内に表示）</p>
<div class="scatter-section section-scatter">
"""
    html += scatter_year_divs
    html += """
</div>
</div>

<div class="print-section">
<h2>年のみ 散布図</h2>
<p>X軸：生年（2024, 2025, 2026…の目盛） / Y軸：各指標の値</p>
<div class="scatter-section section-scatter">
"""
    html += scatter_year_only_divs
    if scatter_script:
        html += "\n" + scatter_script
    html += """
</div>
</div>
</div>
</body>
</html>
"""
    return html


def _build_table_rows(
    rows: List[Dict],
    stats: Dict[str, Dict[str, float]],
    metric_keys: List[str],
    composite_key: str,
    additional_keys: List[Optional[str]],
    label: Any,
    show_rank: bool,
    rank_value: Optional[str] = None,
) -> str:
    """
    rank_value: None のときは行番号（idx+1）を順位に。文字列（例 "_composite_rank"）のときは row のそのキーを順位に（生年月日順表で総合指標順位を表示するため）。
    additional_keys に None が含まれる場合（表幅統一用の空列）は空セルを出力する。
    """
    out = []
    for idx, row in enumerate(rows):
        if show_rank:
            rank = str(row.get(rank_value, idx + 1)) if rank_value else str(idx + 1)
        else:
            rank = ""
        cow_id = html_module.escape(str(row.get("cow_id", "")))
        jpn10 = html_module.escape(str(row.get("jpn10", "")))
        sire = html_module.escape(str(row.get("sire", "")))
        bthd = html_module.escape(str(row.get("bthd", "")))
        cells = [
            f'<td class="sticky-col col-1">{rank}</td>',
            f'<td class="sticky-col text-left cell-full col-2"><a href="#" class="report-cow-link" data-cow-id="{cow_id}">{cow_id}</a></td>',
            f'<td class="text-left cell-full col-3">{jpn10}</td>',
            f'<td class="text-left cell-full col-4">{sire}</td>',
            f'<td class="cell-full col-5">{bthd}</td>',
        ]
        for col_idx, k in enumerate(COMPOSITE_INDEX_KEYS):
            c = 6 + col_idx
            v = row.get(k)
            st = stats.get(k) or {}
            cls = _cell_class(v, st.get("p25"), st.get("p75"), k)
            val_s = f"{v:.2f}" if v is not None else ""
            cells.append(f'<td class="{cls} col-{c}">{val_s}</td>')
        for col_idx, k in enumerate(additional_keys):
            c = 9 + col_idx
            if k is None:
                cells.append(f'<td class="col-{c}"></td>')
                continue
            v = row.get(k)
            st = stats.get(k) or {}
            cls = _cell_class(v, st.get("p25"), st.get("p75"), k)
            val_s = f"{v:.2f}" if v is not None else ""
            cells.append(f'<td class="{cls} col-{c}">{val_s}</td>')
        out.append("<tr>" + "".join(cells) + "</tr>")
    return "\n".join(out)
