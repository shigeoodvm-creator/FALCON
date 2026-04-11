"""
FALCON2 - ダッシュボード計算・HTML生成 Mixin
MainWindow から分離した Mixin クラス。
MainWindow のみが継承することを前提とし、self.* 属性は MainWindow.__init__ で初期化される。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import TYPE_CHECKING, Optional, Literal, List, Dict, Any, Tuple, Union
from pathlib import Path
from datetime import datetime, timedelta, date
import json
import logging
import re
import io
import threading
import webbrowser
import tempfile
import html

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False

if TYPE_CHECKING:
    from db.db_handler import DBHandler
    from modules.formula_engine import FormulaEngine
    from modules.rule_engine import RuleEngine

from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager

logger = logging.getLogger(__name__)

# 散布図の Plotly 表示用（乳検・ゲノム共通）
try:
    from modules.genome_report_html import _get_plotly_scatter_script, _build_plotly_scatter_divs_only
except ImportError:
    _get_plotly_scatter_script = None
    _build_plotly_scatter_divs_only = None


def _build_inline_scatter(points, chart_id, title, x_label, y_label, x_key, y_key,
                          color_fn=None, legend_html="", y_tick_labels=None, hline_default=None):
    """
    インラインSVG散布図を生成。クリックで牛カードへ。基準線（十字）機能付き。
    points: [{"cow_id":..., "x_key_value":..., "y_key_value":..., ...}, ...]
    """
    if not points:
        return f'<div class="section-title">{title}</div><div class="subnote" style="padding:20px;text-align:center;color:#aaa">データがありません</div>'

    W, H = 560, 320
    ML, MR, MT, MB = 52, 16, 28, 44

    xs_raw = [p.get(x_key) for p in points]
    ys_raw = [p.get(y_key) for p in points]

    # X軸: 文字列の場合はカテゴリ（index）扱い
    is_x_cat = isinstance(xs_raw[0], str) if xs_raw else False
    if is_x_cat:
        x_cats = list(dict.fromkeys(xs_raw))  # 順序保持ユニーク
        x_nums = [x_cats.index(x) for x in xs_raw]
        x_min, x_max = 0, max(len(x_cats)-1, 1)
    else:
        x_nums = [float(v) for v in xs_raw if v is not None]
        x_min = min(x_nums) if x_nums else 0
        x_max = max(x_nums) if x_nums else 1
        if x_max == x_min: x_max = x_min + 1

    ys_num = [float(v) for v in ys_raw if v is not None]
    y_min = min(ys_num) if ys_num else 0
    y_max = max(ys_num) if ys_num else 1
    if y_max == y_min: y_max = y_min + 1

    PW = W - ML - MR
    PH = H - MT - MB

    def px(xi):
        return ML + (xi - x_min) / (x_max - x_min) * PW

    def py(yi):
        return MT + (1.0 - (yi - y_min) / (y_max - y_min)) * PH

    # グリッド・軸ラベル
    svg_parts = [f'<svg id="{chart_id}_svg" viewBox="0 0 {W} {H}" style="width:100%;height:auto;display:block;cursor:crosshair" xmlns="http://www.w3.org/2000/svg">']

    # Y軸グリッド
    if y_tick_labels:
        y_ticks = list(y_tick_labels.keys())
    else:
        raw_range = y_max - y_min
        step = max(1, round(raw_range / 5))
        y_ticks = list(range(int(y_min), int(y_max)+1, step))

    for yt in y_ticks:
        yy = py(yt)
        if MT <= yy <= MT + PH:
            svg_parts.append(f'<line x1="{ML}" y1="{yy:.1f}" x2="{ML+PW}" y2="{yy:.1f}" stroke="#f0f0f0" stroke-width="1"/>')
            label = y_tick_labels.get(yt, str(yt)) if y_tick_labels else str(yt)
            svg_parts.append(f'<text x="{ML-4}" y="{yy+3:.1f}" text-anchor="end" font-size="8" fill="#888">{label}</text>')

    # X軸グリッド・ラベル
    if is_x_cat:
        step_x = max(1, len(x_cats)//8)
        for i, cat in enumerate(x_cats):
            if i % step_x == 0:
                xx = px(i)
                svg_parts.append(f'<text x="{xx:.1f}" y="{MT+PH+14}" text-anchor="middle" font-size="7" fill="#888">{cat}</text>')
    else:
        raw_xrange = x_max - x_min
        x_step = max(1, round(raw_xrange / 6))
        xt = int(x_min)
        while xt <= x_max:
            xx = px(xt)
            svg_parts.append(f'<line x1="{xx:.1f}" y1="{MT}" x2="{xx:.1f}" y2="{MT+PH}" stroke="#f0f0f0" stroke-width="1"/>')
            svg_parts.append(f'<text x="{xx:.1f}" y="{MT+PH+14}" text-anchor="middle" font-size="8" fill="#888">{xt}</text>')
            xt += x_step

    # 軸線
    svg_parts.append(f'<line x1="{ML}" y1="{MT}" x2="{ML}" y2="{MT+PH}" stroke="#adb5bd" stroke-width="1.5"/>')
    svg_parts.append(f'<line x1="{ML}" y1="{MT+PH:.1f}" x2="{ML+PW}" y2="{MT+PH:.1f}" stroke="#adb5bd" stroke-width="1.5"/>')

    # 軸タイトル
    svg_parts.append(f'<text x="{ML+PW//2}" y="{H-4}" text-anchor="middle" font-size="9" fill="#666">{x_label}</text>')
    svg_parts.append(f'<text x="10" y="{MT+PH//2}" text-anchor="middle" font-size="9" fill="#666" transform="rotate(-90 10 {MT+PH//2})">{y_label}</text>')

    # 基準線（固定）プレースホルダー
    svg_parts.append(f'<line id="{chart_id}_hline" x1="{ML}" y1="-100" x2="{ML+PW}" y2="-100" stroke="#e53935" stroke-width="1.2" stroke-dasharray="5 3" opacity="0.8"/>')
    svg_parts.append(f'<line id="{chart_id}_vline" x1="-100" y1="{MT}" x2="-100" y2="{MT+PH}" stroke="#e53935" stroke-width="1.2" stroke-dasharray="5 3" opacity="0.8"/>')
    svg_parts.append(f'<text id="{chart_id}_hlabel" x="{ML+PW-2}" y="-100" text-anchor="end" font-size="8" fill="#e53935"></text>')
    svg_parts.append(f'<text id="{chart_id}_vlabel" x="-100" y="{MT+12}" text-anchor="middle" font-size="8" fill="#e53935"></text>')

    # カーソル基準線（移動中）
    svg_parts.append(f'<line id="{chart_id}_chline" x1="{ML}" y1="-100" x2="{ML+PW}" y2="-100" stroke="#e53935" stroke-width="1" stroke-dasharray="3 2" opacity="0.5"/>')
    svg_parts.append(f'<line id="{chart_id}_cvline" x1="-100" y1="{MT}" x2="-100" y2="{MT+PH}" stroke="#e53935" stroke-width="1" stroke-dasharray="3 2" opacity="0.5"/>')

    # プロット点（cow_id をdata属性に）
    for i, p in enumerate(points):
        xi = x_cats.index(p.get(x_key)) if is_x_cat else float(p.get(x_key, 0))
        yi = float(p.get(y_key, 0))
        cx = px(xi)
        cy = py(yi)
        fill = color_fn(p) if color_fn else "#1565C0"
        cow_id = p.get("cow_id", "")
        y_display = p.get("y_label", p.get(y_key, ""))
        svg_parts.append(
            f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="4" fill="{fill}" opacity="0.72" '
            f'stroke="white" stroke-width="0.8" '
            f'data-cow="{cow_id}" data-xv="{p.get(x_key,"")}" data-yv="{y_display}" '
            f'style="cursor:pointer" class="{chart_id}_dot"/>'
        )

    svg_parts.append('</svg>')

    svg_str = "\n".join(svg_parts)

    # JavaScript（クリック→牛カード、基準線）
    js = f"""
<script>
(function(){{
  var chartId = '{chart_id}';
  var svgEl = document.getElementById(chartId + '_svg');
  if (!svgEl) return;
  var crosshairActive = false;
  // SVG座標系パラメータ
  var ML={ML}, MR={MR}, MT={MT}, MB={MB}, W={W}, H={H};
  var xMin={x_min if not is_x_cat else 0}, xMax={x_max if not is_x_cat else max(len(xs_raw)-1, 1)};
  var yMin={y_min}, yMax={y_max};
  var PW={PW}, PH={PH};

  function svgPoint(e) {{
    var rect = svgEl.getBoundingClientRect();
    var scaleX = W / rect.width;
    var scaleY = H / rect.height;
    return {{
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY
    }};
  }}

  function toDataX(svgX) {{ return xMin + (svgX - ML) / PW * (xMax - xMin); }}
  function toDataY(svgY) {{ return yMax - (svgY - MT) / PH * (yMax - yMin); }}

  // ボタン
  var btn = document.getElementById(chartId + '_btn');
  if (btn) {{
    btn.addEventListener('click', function() {{
      crosshairActive = !crosshairActive;
      btn.style.background = crosshairActive ? '#e53935' : '';
      btn.style.color = crosshairActive ? '#fff' : '';
      svgEl.style.cursor = crosshairActive ? 'crosshair' : 'default';
      if (!crosshairActive) {{
        document.getElementById(chartId+'_chline').setAttribute('y1','-100');
        document.getElementById(chartId+'_chline').setAttribute('y2','-100');
        document.getElementById(chartId+'_cvline').setAttribute('x1','-100');
        document.getElementById(chartId+'_cvline').setAttribute('x2','-100');
      }}
    }});
  }}

  // マウス移動：カーソル基準線
  svgEl.addEventListener('mousemove', function(e) {{
    if (!crosshairActive) return;
    var pt = svgPoint(e);
    var svgX = Math.max(ML, Math.min(ML+PW, pt.x));
    var svgY = Math.max(MT, Math.min(MT+PH, pt.y));
    var chline = document.getElementById(chartId+'_chline');
    var cvline = document.getElementById(chartId+'_cvline');
    chline.setAttribute('y1', svgY.toFixed(1));
    chline.setAttribute('y2', svgY.toFixed(1));
    cvline.setAttribute('x1', svgX.toFixed(1));
    cvline.setAttribute('x2', svgX.toFixed(1));
  }});

  // マウスが外れたらカーソル線を消す（固定線は維持）
  svgEl.addEventListener('mouseleave', function() {{
    if (!crosshairActive) return;
    document.getElementById(chartId+'_chline').setAttribute('y1','-1000');
    document.getElementById(chartId+'_chline').setAttribute('y2','-1000');
    document.getElementById(chartId+'_cvline').setAttribute('x1','-1000');
    document.getElementById(chartId+'_cvline').setAttribute('x2','-1000');
  }});

  // クリック：基準線モード中は十字を固定 / 通常モードはドットで牛カードへ
  svgEl.addEventListener('click', function(e) {{
    var target = e.target;
    if (crosshairActive) {{
      // 基準線モード：どこをクリックしても固定
      var pt = svgPoint(e);
      var svgX = Math.max(ML, Math.min(ML+PW, pt.x));
      var svgY = Math.max(MT, Math.min(MT+PH, pt.y));
      var hline = document.getElementById(chartId+'_hline');
      var vline = document.getElementById(chartId+'_vline');
      var hlabel = document.getElementById(chartId+'_hlabel');
      var vlabel = document.getElementById(chartId+'_vlabel');
      hline.setAttribute('y1', svgY.toFixed(1));
      hline.setAttribute('y2', svgY.toFixed(1));
      vline.setAttribute('x1', svgX.toFixed(1));
      vline.setAttribute('x2', svgX.toFixed(1));
      var dy = toDataY(svgY);
      var dx = toDataX(svgX);
      hlabel.setAttribute('y', (svgY - 3).toFixed(1));
      hlabel.textContent = dy.toFixed(1);
      vlabel.setAttribute('x', svgX.toFixed(1));
      vlabel.setAttribute('y', (MT+11).toFixed(1));
      vlabel.textContent = dx.toFixed(1);
    }} else {{
      // 通常モード：ドットクリックで牛カードへ
      if (target.classList.contains(chartId+'_dot')) {{
        var cowId = target.getAttribute('data-cow');
        if (cowId && typeof FALCON_OPEN_COW_PORT !== 'undefined') {{
          fetch('http://127.0.0.1:' + FALCON_OPEN_COW_PORT + '/open_cow?cow_id=' + encodeURIComponent(cowId)).catch(function(){{}});
        }}
      }}
    }}
  }});

  // ダブルクリック：基準線リセット（固定線を消す）
  svgEl.addEventListener('dblclick', function(e) {{
    document.getElementById(chartId+'_hline').setAttribute('y1','-1000');
    document.getElementById(chartId+'_hline').setAttribute('y2','-1000');
    document.getElementById(chartId+'_vline').setAttribute('x1','-1000');
    document.getElementById(chartId+'_vline').setAttribute('x2','-1000');
    document.getElementById(chartId+'_hlabel').textContent='';
    document.getElementById(chartId+'_vlabel').textContent='';
  }});
}})();
</script>
"""

    btn_html = (
        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">'
        f'<div class="section-title" style="margin:0">{title}</div>'
        f'<button id="{chart_id}_btn" style="font-size:10px;padding:3px 10px;border:1px solid #e53935;'
        f'border-radius:4px;background:#fff;color:#e53935;cursor:pointer;white-space:nowrap">＋ 基準線</button>'
        f'</div>'
    )
    legend_div = f'<div style="font-size:9px;margin-bottom:4px;display:flex;gap:10px;flex-wrap:wrap">{legend_html}</div>' if legend_html else ''
    hint = '<div style="font-size:8px;color:#bbb;margin-top:2px">基準線モード：クリックで固定 / ダブルクリックでリセット</div>'

    return btn_html + legend_div + svg_str + hint + js


class DashboardMixin:
    """Mixin: FALCON2 - ダッシュボード計算・HTML生成 Mixin"""
    def _is_cow_disposed_for_dashboard(self, cow_auto_id: int) -> bool:
        """
        個体が売却または死亡廃用されているかをチェック（ダッシュボード用）
        
        Args:
            cow_auto_id: 牛の auto_id
            
        Returns:
            売却または死亡廃用されている場合True
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for event in events:
            event_number = event.get('event_number')
            if event_number in [RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD]:
                return True
        return False

    def _on_dashboard(self):
        """ダッシュボードメニューをクリック（HTML方式）"""
        try:
            # 農場名を取得
            settings_manager = SettingsManager(self.farm_path)
            farm_name = settings_manager.get("farm_name", self.farm_path.name)
            
            # 全個体を取得
            all_cows = self.db.get_all_cows()
            
            # 現存牛のみを取得（売却・死亡廃用を除く）
            existing_cows = [
                cow for cow in all_cows
                if not self._is_cow_disposed_for_dashboard(cow['auto_id'])
            ]
            
            # 経産牛（LACT>=1）を取得（現存牛のみ）
            lactated_cows = [cow for cow in existing_cows if cow.get('lact') is not None and cow.get('lact', 0) >= 1]
            
            # 産次ごとの頭数を集計
            parity_counts = {}
            for cow in lactated_cows:
                lact = cow.get('lact', 0)
                # LACTを整数に変換して確認
                try:
                    lact_int = int(lact) if lact is not None else 0
                except (ValueError, TypeError):
                    lact_int = 0
                
                if lact_int == 1:
                    parity = "1産"
                elif lact_int == 2:
                    parity = "2産"
                elif lact_int >= 3:
                    parity = "3産以上"
                else:
                    # LACTが1未満の場合はスキップ（本来はフィルタリングされているはず）
                    continue
                parity_counts[parity] = parity_counts.get(parity, 0) + 1
            
            total = sum(parity_counts.values())
            total_lact = sum(cow.get('lact', 0) for cow in lactated_cows)
            avg_parity = total_lact / total if total > 0 else 0
            
            # 産次ごとのドーナツチャート用データを準備
            parity_segments = []
            # モダンなカラーパレット（洗練された色）
            colors = ['#FF6B9D', '#4ECDC4', '#95E1D3']
            radius = 50
            circumference = 2 * 3.141592653589793 * radius
            offset = 0
            
            # 固定順序でソート（1産、2産、3産以上）
            parity_order = ["1産", "2産", "3産以上"]
            sorted_parities = []
            for parity in parity_order:
                if parity in parity_counts and parity_counts[parity] > 0:
                    sorted_parities.append((parity, parity_counts[parity]))
            
            for idx, (parity, count) in enumerate(sorted_parities):
                if count <= 0:
                    continue
                percent = (count / total * 100) if total > 0 else 0
                length = (percent / 100) * circumference
                parity_segments.append({
                    "label": parity,
                    "count": count,
                    "percent": round(percent, 1),
                    "color": colors[idx % len(colors)],
                    "length": length,
                    "offset": -offset  # 負の値として保存（stroke-dashoffsetで使用）
                })
                offset += length
            
            # 平均乳量を計算
            milk_stats = self._calculate_dashboard_milk_stats(lactated_cows)
            
            # 平均分娩後日数と妊娠牛の割合を計算
            herd_summary = self._calculate_herd_summary(lactated_cows)
            
            # 体細胞要約を計算
            scc_summary = self._calculate_dashboard_scc_stats(lactated_cows)
            
            # 散布図データを計算（乳検レポートと同じ：直近乳検日の DIM vs 乳量/リニアスコア）
            scatter_data = self._calculate_dashboard_scatter_data(lactated_cows)
            
            # 受胎率を計算（産次別）
            fertility_stats = self._calculate_dashboard_fertility_stats()
            
            # 月ごとの受胎率を計算
            monthly_fertility_stats = self._calculate_dashboard_monthly_fertility_stats()
            
            # 授精回数ごとの受胎率を計算
            insemination_count_fertility_stats = self._calculate_dashboard_insemination_count_fertility_stats()

            # 未経産牛の月ごと受胎率を計算
            heifer_monthly_fertility_stats = self._calculate_dashboard_heifer_monthly_fertility_stats()

            # DIMカテゴリ × 産次 分布
            dim_parity_breakdown = self._calculate_dashboard_dim_parity_breakdown(lactated_cows)

            # 授精種類別受胎率
            ai_et_conception_stats = self._calculate_ai_et_conception_stats(lactated_cows)
            # 初回授精DIM×分娩月日 散布図
            first_ai_scatter = self._calculate_first_ai_dim_scatter(lactated_cows)
            # RC×DIM 散布図
            rc_dim_scatter = self._calculate_rc_dim_scatter(lactated_cows)

            # 21日妊娠率・授精実施率を計算
            pr_hdr_data = self._calculate_dashboard_pr_hdr()
            repro_detail = self._calculate_dashboard_repro_detail()

            # 累積妊娠頭数グラフ
            cumulative_pregnancy_data = self._calculate_dashboard_cumulative_pregnancy()

            # 牛群動態（分娩予定月×産子種類）を計算
            herd_dynamics_data = None
            if self.farm_path:
                try:
                    from modules.herd_dynamics_report import build_herd_dynamics_data
                    herd_dynamics_data = build_herd_dynamics_data(self.db, self.formula_engine, self.farm_path)
                except Exception as e:
                    logging.debug(f"ダッシュボード: 牛群動態データ取得スキップ: {e}")

            # BHB高値牛・乳量低下牛アラート
            milk_alerts = self._calculate_dashboard_milk_alerts()

            # 乳検月次トレンドHTML（12ヶ月）
            milk_trend_html = self._calculate_dashboard_milk_trend()

            # HTML生成
            html_content = self._build_dashboard_html(
                farm_name,
                parity_segments,
                avg_parity,
                total,
                milk_stats,
                fertility_stats,
                monthly_fertility_stats,
                herd_summary,
                scc_summary,
                scatter_data,
                insemination_count_fertility_stats,
                herd_dynamics_data,
                pr_hdr_data=pr_hdr_data,
                repro_detail=repro_detail,
                heifer_monthly_fertility_stats=heifer_monthly_fertility_stats,
                cumulative_pregnancy_data=cumulative_pregnancy_data,
                dim_parity_breakdown=dim_parity_breakdown,
                ai_et_conception_stats=ai_et_conception_stats,
                first_ai_scatter=first_ai_scatter,
                rc_dim_scatter=rc_dim_scatter,
                milk_alerts=milk_alerts,
                milk_trend_html=milk_trend_html,
            )
            
            # 一時HTMLファイルを作成
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.html', delete=False) as f:
                f.write(html_content)
                temp_file_path = f.name
            
            # ブラウザで開く
            webbrowser.open(f'file://{temp_file_path}')
        except Exception as e:
            logging.error(f"ダッシュボード表示エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ダッシュボードの準備中にエラーが発生しました: {e}")
    
    def _calculate_dashboard_milk_stats(self, lactated_cows):
        """平均乳量統計を計算（群全体の直近の乳検日の平均）"""
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "avg_milk": None,
                "max_milk": None,
                "min_milk": None,
                "median_milk": None,
                "avg_first_parity": None,
                "avg_multiparous": None,
                "latest_date": None,
                "milk_bins": {},
                "milk_yields": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "avg_milk": None,
                "max_milk": None,
                "min_milk": None,
                "median_milk": None,
                "avg_first_parity": None,
                "avg_multiparous": None,
                "latest_date": None,
                "milk_bins": {},
                "milk_yields": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        # 乳量データを取得
        milk_yields = []
        first_parity_milk_yields = []
        multiparous_milk_yields = []
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得（産次を確認するため）
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 経産牛（LACT>=1）のみを対象
            lact = cow.get('lact', 0)
            if lact < 1:
                continue
            
            # json_dataから乳量を取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                milk_yield = json_data.get('milk_yield') or json_data.get('milk_kg')
                if milk_yield is not None:
                    try:
                        milk_value = float(milk_yield)
                        if milk_value > 0:
                            milk_yields.append(milk_value)
                            
                            # 産次で分類
                            if lact == 1:
                                first_parity_milk_yields.append(milk_value)
                            elif lact >= 2:
                                multiparous_milk_yields.append(milk_value)
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        # 平均乳量を計算
        avg_milk = sum(milk_yields) / len(milk_yields) if milk_yields else None
        avg_first_parity = sum(first_parity_milk_yields) / len(first_parity_milk_yields) if first_parity_milk_yields else None
        avg_multiparous = sum(multiparous_milk_yields) / len(multiparous_milk_yields) if multiparous_milk_yields else None
        max_milk = max(milk_yields) if milk_yields else None
        min_milk = min(milk_yields) if milk_yields else None
        _sy = sorted(milk_yields)
        _n = len(_sy)
        median_milk = (_sy[_n // 2] if _n % 2 == 1 else (_sy[_n // 2 - 1] + _sy[_n // 2]) / 2) if _n else None
        
        # 乳量階層を計算（~20kg, 20kg台, 30kg台, 40kg台, ~50kg）
        milk_bins = {
            "~20kg": 0,
            "20kg台": 0,
            "30kg台": 0,
            "40kg台": 0,
            "~50kg": 0
        }
        
        for milk_value in milk_yields:
            if milk_value < 20:
                milk_bins["~20kg"] += 1
            elif 20 <= milk_value < 30:
                milk_bins["20kg台"] += 1
            elif 30 <= milk_value < 40:
                milk_bins["30kg台"] += 1
            elif 40 <= milk_value < 50:
                milk_bins["40kg台"] += 1
            else:  # >= 50
                milk_bins["~50kg"] += 1
        
        return {
            "avg_milk": avg_milk,
            "max_milk": max_milk,
            "min_milk": min_milk,
            "median_milk": median_milk,
            "avg_first_parity": avg_first_parity,
            "avg_multiparous": avg_multiparous,
            "latest_date": latest_milk_test_date,
            "milk_bins": milk_bins,
            "milk_yields": milk_yields
        }

    def _calculate_dashboard_milk_alerts(self):
        """BHB高値牛（>=0.13）と乳量15%以上低下牛を計算（最新2乳検日ベース）"""
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        if not all_events:
            return {"bhb_rows": [], "milk_drop_rows": [], "latest_date": None}
        dates = sorted({e.get("event_date") for e in all_events if e.get("event_date")}, reverse=True)
        if not dates:
            return {"bhb_rows": [], "milk_drop_rows": [], "latest_date": None}
        latest_date = dates[0]
        prev_date = dates[1] if len(dates) > 1 else None

        # 前月乳量を牛ごとに収集
        prev_milk_by_cow: Dict[int, float] = {}
        if prev_date:
            for e in all_events:
                if e.get("event_date") != prev_date:
                    continue
                cid = e.get("cow_auto_id")
                if not cid:
                    continue
                j = e.get("json_data") or {}
                if isinstance(j, str):
                    try:
                        j = json.loads(j)
                    except Exception:
                        j = {}
                mv = j.get("milk_yield") or j.get("milk_kg")
                try:
                    mv = float(mv) if mv not in (None, "") else None
                except (ValueError, TypeError):
                    mv = None
                if mv:
                    prev_milk_by_cow[cid] = mv

        # 最新乳検日の全分娩イベントを一括取得してDIM計算用にキャッシュ
        calv_events = self.db.get_events_by_number(RuleEngine.EVENT_CALV, include_deleted=False)
        last_calv_by_cow: Dict[int, str] = {}
        for ce in calv_events:
            cid = ce.get("cow_auto_id")
            cd = ce.get("event_date")
            if not cid or not cd:
                continue
            if cd <= latest_date:
                if cid not in last_calv_by_cow or cd > last_calv_by_cow[cid]:
                    last_calv_by_cow[cid] = cd

        bhb_rows: List[Dict] = []
        milk_drop_rows: List[Dict] = []

        for event in all_events:
            if event.get("event_date") != latest_date:
                continue
            cow_auto_id = event.get("cow_auto_id")
            if not cow_auto_id:
                continue
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            lact = cow.get("lact")
            try:
                lact = int(lact) if lact is not None else None
            except (ValueError, TypeError):
                lact = None
            if lact is None or lact < 1:
                continue

            j = event.get("json_data") or {}
            if isinstance(j, str):
                try:
                    j = json.loads(j)
                except Exception:
                    j = {}

            def _fv(v):
                return str(v) if v not in (None, "") else "-"

            milk_v = j.get("milk_yield") or j.get("milk_kg")
            try:
                milk_num = float(milk_v) if milk_v not in (None, "") else None
            except (ValueError, TypeError):
                milk_num = None

            bhb_v = j.get("bhb")
            try:
                bhb_num = float(bhb_v) if bhb_v not in (None, "") else None
            except (ValueError, TypeError):
                bhb_num = None

            scc_v = j.get("scc")
            try:
                scc_num = float(scc_v) if scc_v not in (None, "") else None
            except (ValueError, TypeError):
                scc_num = None

            last_calv = last_calv_by_cow.get(cow_auto_id)
            dim = None
            if last_calv:
                try:
                    cd_dt = datetime.strptime(last_calv, "%Y-%m-%d")
                    ld_dt = datetime.strptime(latest_date, "%Y-%m-%d")
                    dim = (ld_dt - cd_dt).days
                except Exception:
                    pass

            prev_milk = prev_milk_by_cow.get(cow_auto_id)
            row = {
                "cow_id": _fv(cow.get("cow_id", "")),
                "jpn10": _fv(cow.get("jpn10", "")),
                "lact": _fv(lact),
                "dim": _fv(dim),
                "milk": f"{milk_num:.1f}" if milk_num is not None else "-",
                "prev_milk": f"{prev_milk:.1f}" if prev_milk is not None else "-",
                "scc": f"{scc_num:.0f}" if scc_num is not None else "-",
                "bhb": f"{bhb_num:.3f}" if bhb_num is not None else "-",
                "bhb_class": " bhb-alert" if bhb_num is not None and bhb_num >= 0.13 else "",
                "scc_class": " scc-alert" if scc_num is not None and scc_num >= 200 else "",
            }
            if bhb_num is not None and bhb_num >= 0.13:
                bhb_rows.append(row)
            if milk_num is not None and prev_milk is not None and prev_milk > 0:
                if milk_num <= prev_milk * 0.85:
                    milk_drop_rows.append(dict(row))

        bhb_rows.sort(key=lambda r: float(r["bhb"]) if r["bhb"] != "-" else 0, reverse=True)
        return {"bhb_rows": bhb_rows, "milk_drop_rows": milk_drop_rows, "latest_date": latest_date}

    def _calculate_dashboard_milk_trend(self):
        """直近12ヶ月の乳検月次トレンドHTMLを生成"""
        try:
            from modules.milk_report_extras import compute_monthly_milk_trend, build_12month_trend_section_html
            today_str = datetime.now().strftime("%Y-%m-%d")
            all_events = self.db.get_events_by_number(601, include_deleted=False)
            if not all_events:
                return ""
            trend_rows = compute_monthly_milk_trend(all_events, today_str, 12)
            if not trend_rows:
                return ""
            return build_12month_trend_section_html(trend_rows)
        except Exception as e:
            logging.debug(f"乳検トレンドHTML生成スキップ: {e}")
            return ""

    def _filter_existing_cows(self, cow_rows: List) -> List:
        """
        現存牛のみをフィルタリング（集計・リスト・グラフコマンド用）
        
        Args:
            cow_rows: 牛のデータ行のリスト（auto_idを含む辞書またはタプル）
            
        Returns:
            現存牛のみのリスト
        """
        existing_cows = []
        for cow_row in cow_rows:
            # cow_rowが辞書の場合
            if isinstance(cow_row, dict):
                cow_auto_id = cow_row.get('auto_id')
            # cow_rowがタプルやRowオブジェクトの場合
            else:
                # タプルの場合、最初の要素がauto_idと仮定
                try:
                    cow_auto_id = cow_row[0] if hasattr(cow_row, '__getitem__') else None
                except (IndexError, TypeError):
                    continue
            
            if cow_auto_id and not self._is_cow_disposed_for_dashboard(cow_auto_id):
                existing_cows.append(cow_row)
        
        return existing_cows
    
    def _calculate_dashboard_fertility_stats(self):
        """産次ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（産次別）
            result = self._calculate_conception_rate(
                self.db,
                "産次",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_dashboard_monthly_fertility_stats(self):
        """月ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（月別）
            result = self._calculate_conception_rate(
                self.db,
                "月",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"月ごと受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_dashboard_heifer_monthly_fertility_stats(self):
        """未経産牛の月ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        try:
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            result = self._calculate_conception_rate(
                self.db, "月", start_date, end_date, "", cow_type="未経産"
            )
            if not result:
                return None
            stats = result.get('stats', {})
            if not stats:
                return None
            return {"stats": stats, "start_date": start_date, "end_date": end_date}
        except Exception as e:
            logging.error(f"未経産牛月ごと受胎率統計計算エラー: {e}", exc_info=True)
            return None

    def _calculate_dashboard_dim_parity_breakdown(self, lactated_cows):
        """DIMカテゴリ × 産次 の分布を計算（経産牛全体）。妊娠/未受胎/繁殖中止ステータスも集計。"""
        from datetime import datetime
        # 繁殖検診レポートの open_days histogram と同じ区切り
        bins   = [0, 50, 85, 115, 150, 200, 300, 99999]
        b_labels = ["~50", "51~85", "86~115", "116~150", "151~200", "201~300", "300超"]
        p_labels = ["1産", "2産", "3産", "4産以上"]
        n_bins = len(b_labels)
        n_par  = len(p_labels)
        data = [[0] * n_par for _ in range(n_bins)]
        # 妊娠ステータス別集計（RC: 5/6=妊娠, 1=繁殖中止, 2/3/4=未受胎）
        status_data = [{"pregnant": 0, "open": 0, "dnb": 0} for _ in range(n_bins)]
        today = datetime.now()
        for cow in lactated_cows:
            lact = cow.get("lact")
            clvd = cow.get("clvd")
            if lact is None or not clvd:
                continue
            try:
                dim = (today - datetime.strptime(clvd, "%Y-%m-%d")).days
                lact_int = int(lact)
            except (ValueError, TypeError):
                continue
            if dim < 0 or lact_int <= 0:
                continue
            # DIM bin
            b_idx = None
            for i in range(len(bins) - 1):
                if bins[i] <= dim < bins[i + 1]:
                    b_idx = i; break
            if b_idx is None:
                continue
            # 産次グループ
            p_idx = 0 if lact_int == 1 else 1 if lact_int == 2 else 2 if lact_int == 3 else 3
            data[b_idx][p_idx] += 1
            # 妊娠ステータス
            rc = cow.get("rc", 4)
            if rc in [5, 6]:
                status_data[b_idx]["pregnant"] += 1
            elif rc == 1:
                status_data[b_idx]["dnb"] += 1
            else:
                status_data[b_idx]["open"] += 1
        row_totals = [sum(row) for row in data]
        col_totals = [sum(data[r][c] for r in range(n_bins)) for c in range(n_par)]
        return {
            "data": data,
            "bin_labels": b_labels,
            "parity_labels": p_labels,
            "row_totals": row_totals,
            "col_totals": col_totals,
            "grand_total": sum(row_totals),
            "status_data": status_data,
        }

    def _calculate_dashboard_cumulative_pregnancy(self):
        """累積妊娠頭数グラフ用データを計算（reproduction_checkup_reportの関数を再利用）"""
        try:
            from datetime import datetime
            from modules.reproduction_checkup_report import (
                annual_required_pregnancies,
                get_current_parous_count,
                get_goal_values,
                get_monthly_pregnancy_counts,
                _cumulative_from_monthly,
                get_latest_pregnancy_confirmed_ai_et_date,
            )
            from settings_manager import SettingsManager

            if not self.farm_path:
                return None

            settings_manager = SettingsManager(self.farm_path)
            parous = get_current_parous_count(self.db, self.rule_engine)
            interval, replacement, abortion = get_goal_values(settings_manager)
            annual_required = annual_required_pregnancies(
                parous, float(interval), float(replacement), float(abortion)
            )

            latest_ai_et_date = get_latest_pregnancy_confirmed_ai_et_date(self.db)
            if latest_ai_et_date:
                try:
                    current_year = int(latest_ai_et_date[:4])
                except ValueError:
                    current_year = datetime.now().year
            else:
                current_year = datetime.now().year

            start_cur  = f"{current_year}-01-01"
            end_cur    = f"{current_year}-12-31"
            start_prev = f"{current_year - 1}-01-01"
            end_prev   = f"{current_year - 1}-12-31"

            monthly_current = get_monthly_pregnancy_counts(self.db, start_cur, end_cur)
            monthly_prev    = get_monthly_pregnancy_counts(self.db, start_prev, end_prev)

            cum_cur  = _cumulative_from_monthly(monthly_current, current_year)
            cum_prev = _cumulative_from_monthly(monthly_prev, current_year - 1)

            # 確定実績頭数（直近妊娠確定AI/ET日時点）
            current_pregnancies = None
            latest_dt = None
            if latest_ai_et_date:
                try:
                    latest_dt = datetime.strptime(latest_ai_et_date[:10], "%Y-%m-%d")
                    cur_dates = [datetime.strptime(d, "%Y-%m-%d") for d, _ in cum_cur]
                    cur_vals  = [v for _, v in cum_cur]
                    solid_end = len(cur_dates)
                    for i, d in enumerate(cur_dates):
                        if d > latest_dt:
                            solid_end = i
                            break
                    current_pregnancies = cur_vals[solid_end - 1] if solid_end > 0 else 0.0
                except (ValueError, TypeError):
                    pass

            month_per = round(annual_required / 12, 1) if annual_required > 0 else 0.0

            return {
                "annual_required":      annual_required,
                "month_per":            month_per,
                "current_year":         current_year,
                "cum_cur":              cum_cur,
                "cum_prev":             cum_prev,
                "latest_ai_et_date":    latest_ai_et_date,
                "latest_dt":            latest_dt,
                "current_pregnancies":  current_pregnancies,
                "parous":               parous,
                "interval":             interval,
            }
        except Exception as e:
            logging.error(f"累積妊娠頭数計算エラー: {e}", exc_info=True)
            return None

    def _calculate_dashboard_insemination_count_fertility_stats(self):
        """授精回数ごとの受胎率統計を計算"""
        from datetime import datetime, timedelta
        
        try:
            # 期間を設定（現在の日付から30日前からの1年間）
            today = datetime.now()
            end_date_dt = today - timedelta(days=30)
            end_date = end_date_dt.strftime('%Y-%m-%d')
            start_date_dt = end_date_dt - timedelta(days=365)
            start_date = start_date_dt.strftime('%Y-%m-%d')
            
            # 受胎率を計算（授精回数別）
            result = self._calculate_conception_rate(
                self.db,
                "授精回数",
                start_date,
                end_date,
                ""  # 条件なし
            )
            
            if not result:
                return None
            
            stats = result.get('stats', {})
            if not stats:
                return None
            
            return {
                "stats": stats,
                "start_date": start_date,
                "end_date": end_date
            }
        except Exception as e:
            logging.error(f"授精回数ごと受胎率統計計算エラー: {e}", exc_info=True)
            return None
    
    def _calculate_herd_summary(self, lactated_cows):
        """牛群要約を計算（平均分娩後日数、妊娠牛の割合、分娩間隔、予定分娩間隔）"""
        from datetime import datetime, timedelta
        
        try:
            dim_values = []
            pregnant_count = 0
            total_count = len(lactated_cows)
            cci_values = []  # 分娩間隔
            pcci_values = []  # 予定分娩間隔
            
            today = datetime.now()
            
            for cow in lactated_cows:
                cow_auto_id = cow.get('auto_id')
                if not cow_auto_id:
                    continue
                
                # 繁殖コードを取得
                rc = cow.get('rc')
                if rc in [5, 6]:  # 5: Pregnant, 6: Dry
                    pregnant_count += 1
                
                # 分娩日を取得してDIMを計算
                clvd = cow.get('clvd')
                if clvd:
                    try:
                        clvd_dt = datetime.strptime(clvd, '%Y-%m-%d')
                        dim = (today - clvd_dt).days
                        if dim >= 0:  # 分娩日が今日以前の場合のみ
                            dim_values.append(dim)
                    except (ValueError, TypeError):
                        pass
                
                # 分娩間隔（CCI）を計算
                # 分娩イベントが2回以上ある場合のみ計算
                events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
                calving_events = [e for e in events if e.get('event_number') == self.rule_engine.EVENT_CALV]
                calving_events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
                
                if len(calving_events) >= 2:
                    try:
                        prev_calv_date = calving_events[-2].get('event_date')
                        last_calv_date = calving_events[-1].get('event_date')
                        if prev_calv_date and last_calv_date:
                            prev_dt = datetime.strptime(prev_calv_date, '%Y-%m-%d')
                            last_dt = datetime.strptime(last_calv_date, '%Y-%m-%d')
                            cci = (last_dt - prev_dt).days
                            if cci > 0:
                                cci_values.append(cci)
                    except (ValueError, TypeError):
                        pass
                
                # 予定分娩間隔（PCCI）を計算
                # FormulaEngineのcalculateメソッドを使用（集計コマンドと同じロジック）
                if self.formula_engine:
                    try:
                        calculated = self.formula_engine.calculate(cow_auto_id, 'PCI')
                        if calculated and 'PCI' in calculated:
                            pcci = calculated['PCI']
                            if pcci is not None and isinstance(pcci, (int, float)) and pcci > 0:
                                pcci_values.append(int(pcci))
                    except Exception as e:
                        logging.debug(f"予定分娩間隔計算エラー: cow_auto_id={cow_auto_id}, error={e}")
                        pass
            
            # 平均分娩後日数を計算
            avg_dim = sum(dim_values) / len(dim_values) if dim_values else None
            
            # 妊娠牛の割合を計算
            pregnancy_rate = (pregnant_count / total_count * 100) if total_count > 0 else 0
            
            # 平均分娩間隔を計算
            avg_cci = sum(cci_values) / len(cci_values) if cci_values else None
            
            # 平均予定分娩間隔を計算
            avg_pcci = sum(pcci_values) / len(pcci_values) if pcci_values else None
            
            return {
                "avg_dim": avg_dim,
                "pregnancy_rate": pregnancy_rate,
                "pregnant_count": pregnant_count,
                "total_count": total_count,
                "avg_cci": avg_cci,
                "avg_pcci": avg_pcci
            }
        except Exception as e:
            logging.error(f"牛群要約計算エラー: {e}", exc_info=True)
            return {
                "avg_dim": None,
                "pregnancy_rate": 0,
                "pregnant_count": 0,
                "total_count": 0,
                "avg_cci": None,
                "avg_pcci": None
            }
    
    def _calculate_ai_et_conception_stats(self, lactated_cows):
        """経産牛の授精種類コード別受胎率を計算する（自然発情・WPG・CIDRなど）。"""
        import json as _json
        from settings_manager import SettingsManager
        # 授精種類コード→表示名マップを農場設定から取得
        type_names: dict = {}
        if self.farm_path:
            try:
                sm = SettingsManager(self.farm_path)
                type_names = sm.get("insemination_type_codes", {}) or {}
            except Exception:
                pass
        UNKNOWN = "不明"
        stats: dict = {}  # type_label -> {"count":0, "pregnant":0}

        for cow in lactated_cows:
            cow_auto_id = cow.get("auto_id")
            if not cow_auto_id:
                continue
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            clvd = cow.get("clvd") or ""
            ai_events = [
                e for e in events
                if e.get("event_number") in (200, 201)
                and e.get("event_date", "") >= clvd
            ]
            preg_events = [
                e for e in events
                if e.get("event_number") in (303, 307)
                and e.get("event_date", "") >= clvd
            ]
            neg_events = [
                e for e in events
                if e.get("event_number") in (302, 306)
                and e.get("event_date", "") >= clvd
            ]
            for ai in ai_events:
                ai_date = ai.get("event_date", "")
                # json_dataから授精種類コードを取得
                jd = ai.get("json_data") or {}
                if isinstance(jd, str):
                    try:
                        jd = _json.loads(jd)
                    except Exception:
                        jd = {}
                type_code = (
                    jd.get("insemination_type_code")
                    or jd.get("type")
                    or jd.get("ai_type")
                )
                label = type_names.get(str(type_code), UNKNOWN) if type_code else UNKNOWN
                if label not in stats:
                    stats[label] = {"count": 0, "pregnant": 0}
                stats[label]["count"] += 1
                # 受胎判定
                later_pos = [e for e in preg_events if e.get("event_date", "") > ai_date]
                later_neg = [e for e in neg_events  if e.get("event_date", "") > ai_date]
                if later_pos:
                    first_pos = min(e.get("event_date", "") for e in later_pos)
                    first_neg = min((e.get("event_date", "") for e in later_neg), default="9999")
                    if first_pos <= first_neg:
                        stats[label]["pregnant"] += 1

        # 授精数の多い順にソート
        result = []
        for label, s in sorted(stats.items(), key=lambda kv: -kv[1]["count"]):
            cr = round(s["pregnant"] / s["count"] * 100) if s["count"] > 0 else None
            result.append({"name": label, "count": s["count"], "pregnant": s["pregnant"], "cr": cr})
        return result

    def _calculate_first_ai_dim_scatter(self, lactated_cows):
        """初回授精DIM(Y) × 分娩月日(X) の散布図データを返す。"""
        from datetime import datetime
        points = []
        for cow in lactated_cows:
            cow_auto_id = cow.get("auto_id")
            clvd = cow.get("clvd") or ""
            cow_id = cow.get("cow_id", "")
            if not cow_auto_id or not clvd:
                continue
            try:
                clvd_dt = datetime.strptime(clvd, "%Y-%m-%d")
            except ValueError:
                continue
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            # 分娩日以降の最初のAI/ETイベント
            ai_events = sorted(
                [e for e in events if e.get("event_number") in (200, 201) and e.get("event_date","") >= clvd],
                key=lambda e: e.get("event_date","")
            )
            if not ai_events:
                continue
            first_ai_date = ai_events[0].get("event_date","")
            try:
                first_ai_dt = datetime.strptime(first_ai_date, "%Y-%m-%d")
            except ValueError:
                continue
            first_ai_dim = (first_ai_dt - clvd_dt).days
            if first_ai_dim < 0:
                continue
            rc = cow.get("rc", 4)
            lact = cow.get("lact", 1)
            points.append({
                "cow_id": cow_id,
                "auto_id": cow_auto_id,
                "x_date": clvd,
                "x_label": f"{clvd_dt.month}/{clvd_dt.day}",
                "y": first_ai_dim,
                "lact": lact,
                "rc": rc,
            })
        # 分娩日順にソート
        points.sort(key=lambda p: p["x_date"])
        return points

    def _calculate_rc_dim_scatter(self, lactated_cows):
        """RC(Y) × DIM(X) の散布図データを返す。"""
        from datetime import datetime
        today = datetime.now()
        points = []
        rc_order = [1, 2, 4, 3, 5, 6]
        rc_position = {rc: idx for idx, rc in enumerate(rc_order)}
        rc_labels = {1:"繁殖停止", 2:"Fresh", 3:"授精後", 4:"空胎", 5:"妊娠中", 6:"乾乳"}
        for cow in lactated_cows:
            clvd = cow.get("clvd") or ""
            cow_id = cow.get("cow_id","")
            auto_id = cow.get("auto_id")
            if not clvd or not auto_id:
                continue
            try:
                dim = (today - datetime.strptime(clvd, "%Y-%m-%d")).days
            except ValueError:
                continue
            if dim < 0:
                continue
            rc = cow.get("rc", 4)
            if rc not in rc_position:
                continue
            lact = cow.get("lact", 1)
            points.append({
                "cow_id": cow_id,
                "auto_id": auto_id,
                "x": dim,
                "y": rc_position[rc],
                "y_label": f"{rc}: {rc_labels.get(rc, str(rc))}",
                "rc_label": rc_labels.get(rc, str(rc)),
                "lact": lact,
            })
        return points

    def _calculate_dashboard_pr_hdr(self):
        """21日妊娠率（PR）と授精実施率（HDR）をメイン画面と同じ18サイクル遡りで計算する。"""
        try:
            from datetime import datetime
            from modules.reproduction_analysis import ReproductionAnalysis
            today = datetime.now()
            end_date = today.strftime("%Y-%m-%d")
            from constants import CONFIG_DEFAULT_DIR
            item_dict_path = CONFIG_DEFAULT_DIR / "item_dictionary.json"
            from modules.formula_engine import FormulaEngine
            fe = FormulaEngine(self.db, item_dict_path)
            analyzer = ReproductionAnalysis(
                db=self.db, rule_engine=self.rule_engine,
                formula_engine=fe, vwp=50
            )
            # period_start=None → DEFAULT_CYCLES(18)サイクル遡る（メイン画面と同じロジック）
            results = analyzer.analyze(None, end_date) or []
            total_br_el   = sum(getattr(r, "br_el",        0) or 0 for r in results)
            total_bred    = sum(getattr(r, "bred",          0) or 0 for r in results)
            total_preg_el = sum(getattr(r, "preg_eligible", 0) or 0 for r in results)
            total_preg    = sum(getattr(r, "preg",          0) or 0 for r in results)
            avg_hdr = round(total_bred    / total_br_el   * 100, 1) if total_br_el   else None
            avg_pr  = round(total_preg    / total_preg_el * 100, 1) if total_preg_el else None
            actual_start = getattr(results[0], "start_date", end_date) if results else end_date
            return {
                "avg_pr":  avg_pr,
                "avg_hdr": avg_hdr,
                "start_date": actual_start,
                "end_date":   end_date,
            }
        except Exception as e:
            logging.debug(f"PR/HDR計算スキップ: {e}")
            return {"avg_pr": None, "avg_hdr": None, "start_date": None, "end_date": None}

    def _calculate_dashboard_repro_detail(self):
        """21日サイクル別PR/HDR詳細とDIM別妊娠率をメイン画面と同じ18サイクル遡りで計算する。"""
        try:
            from datetime import datetime
            from modules.reproduction_analysis import ReproductionAnalysis
            from modules.formula_engine import FormulaEngine
            today = datetime.now()
            end_date = today.strftime("%Y-%m-%d")
            from constants import CONFIG_DEFAULT_DIR
            item_dict_path = CONFIG_DEFAULT_DIR / "item_dictionary.json"
            fe = FormulaEngine(self.db, item_dict_path)
            analyzer = ReproductionAnalysis(
                db=self.db, rule_engine=self.rule_engine,
                formula_engine=fe, vwp=50
            )
            # period_start=None → DEFAULT_CYCLES(18)サイクル遡る（メイン画面と同じロジック）
            cycle_results = analyzer.analyze(None, end_date) or []
            dim_results   = analyzer.analyze_by_dim(None, end_date, n_ranges=10) or []
            actual_start  = getattr(cycle_results[0], "start_date", end_date) if cycle_results else end_date
            return {
                "cycles":     cycle_results,
                "dim_pr":     dim_results,
                "start_date": actual_start,
                "end_date":   end_date,
            }
        except Exception as e:
            logging.debug(f"繁殖詳細計算スキップ: {e}")
            return None

    def _calculate_dashboard_scc_stats(self, lactated_cows):
        """体細胞要約を計算（リニアスコア階層、平均体細胞、平均リニアスコア）"""
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "avg_scc": None,
                "avg_ls": None,
                "avg_first_parity_ls": None,
                "avg_multiparous_ls": None,
                "latest_date": None,
                "ls_bins": {},
                "ls_values": [],
                "scc_values": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "avg_scc": None,
                "avg_ls": None,
                "avg_first_parity_ls": None,
                "avg_multiparous_ls": None,
                "latest_date": None,
                "ls_bins": {},
                "ls_values": [],
                "scc_values": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        # 体細胞とリニアスコアデータを取得
        scc_values = []
        ls_values = []
        first_parity_ls_values = []
        multiparous_ls_values = []
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得（産次を確認するため）
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 経産牛（LACT>=1）のみを対象
            lact = cow.get('lact', 0)
            if lact < 1:
                continue
            
            # json_dataから体細胞とリニアスコアを取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                scc_value = json_data.get('scc')
                ls_value = json_data.get('ls')
                
                if scc_value is not None:
                    try:
                        scc_num = float(scc_value)
                        if scc_num > 0:
                            scc_values.append(scc_num)
                    except (ValueError, TypeError):
                        pass
                
                if ls_value is not None:
                    try:
                        ls_num = float(ls_value)
                        if ls_num >= 0:
                            ls_values.append(ls_num)
                            
                            # 産次で分類
                            if lact == 1:
                                first_parity_ls_values.append(ls_num)
                            elif lact >= 2:
                                multiparous_ls_values.append(ls_num)
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        # 平均を計算
        avg_scc = sum(scc_values) / len(scc_values) if scc_values else None
        avg_ls = sum(ls_values) / len(ls_values) if ls_values else None
        avg_first_parity_ls = sum(first_parity_ls_values) / len(first_parity_ls_values) if first_parity_ls_values else None
        avg_multiparous_ls = sum(multiparous_ls_values) / len(multiparous_ls_values) if multiparous_ls_values else None
        
        # リニアスコア階層を計算（LS2以下、LS3-4、LS5以上）
        ls_bins = {
            "LS2以下": 0,
            "LS3-4": 0,
            "LS5以上": 0
        }
        
        for ls_value in ls_values:
            if ls_value <= 2:
                ls_bins["LS2以下"] += 1
            elif 3 <= ls_value <= 4:
                ls_bins["LS3-4"] += 1
            else:  # >= 5
                ls_bins["LS5以上"] += 1
        
        return {
            "avg_scc": avg_scc,
            "avg_ls": avg_ls,
            "avg_first_parity_ls": avg_first_parity_ls,
            "avg_multiparous_ls": avg_multiparous_ls,
            "latest_date": latest_milk_test_date,
            "ls_bins": ls_bins,
            "ls_values": ls_values,
            "scc_values": scc_values
        }
    
    def _calculate_dashboard_scatter_data(self, lactated_cows):
        """散布図データを計算（検定時のDIM、産次を含む）"""
        from datetime import datetime
        
        # 全イベントから乳検イベント（event_number=601）を取得
        all_events = self.db.get_events_by_number(601, include_deleted=False)
        
        if not all_events:
            return {
                "milk_points": [],
                "ls_points": []
            }
        
        # 最も新しい乳検日を特定
        latest_milk_test_date = max(event.get('event_date', '') for event in all_events if event.get('event_date'))
        
        if not latest_milk_test_date:
            return {
                "milk_points": [],
                "ls_points": []
            }
        
        # その乳検日のイベントを取得
        latest_milk_test_events = [
            event for event in all_events
            if event.get('event_date') == latest_milk_test_date
        ]
        
        # 分娩イベント：検定日以前で最新の分娩日を採用（検定時点のDIMを再現）
        calv_events = self.db.get_events_by_number(RuleEngine.EVENT_CALV, include_deleted=False)
        calv_dates_by_cow: Dict[int, List[str]] = {}
        for calv_event in calv_events:
            cow_auto_id = calv_event.get('cow_auto_id')
            event_date = calv_event.get('event_date')
            if not cow_auto_id or not event_date:
                continue
            try:
                cow_auto_id_int = int(cow_auto_id)
            except (ValueError, TypeError):
                continue
            date_str = str(event_date).strip()[:10]
            calv_dates_by_cow.setdefault(cow_auto_id_int, []).append(date_str)
        for cow_auto_id_int in calv_dates_by_cow:
            calv_dates_by_cow[cow_auto_id_int].sort()
        
        milk_points = []  # (dim, milk_yield, parity)
        ls_points = []    # (dim, ls, parity)
        
        def calc_dim_on_date(clvd_date: str, event_date: str) -> int:
            """検定時のDIMを計算"""
            if not clvd_date:
                return None
            try:
                clvd_dt = datetime.strptime(clvd_date, "%Y-%m-%d")
                event_dt = datetime.strptime(event_date, "%Y-%m-%d")
            except (ValueError, TypeError):
                return None
            if event_dt < clvd_dt:
                return None
            return (event_dt - clvd_dt).days
        
        def clvd_for_milk_dim(cow_auto_id: Any, cow: Optional[Dict[str, Any]]) -> Optional[str]:
            """検定日時点の直近分娩日。イベントが無い場合は cow.clvd を補助的に使用。"""
            try:
                cow_auto_id_int = int(cow_auto_id)
            except (ValueError, TypeError):
                return cow.get('clvd') if cow else None
            dates = calv_dates_by_cow.get(cow_auto_id_int, [])
            on_or_before = [d for d in dates if d <= latest_milk_test_date]
            if on_or_before:
                return max(on_or_before)
            return cow.get('clvd') if cow else None
        
        for event in latest_milk_test_events:
            cow_auto_id = event.get('cow_auto_id')
            if not cow_auto_id:
                continue
            
            # 除籍牛を除外
            if self._is_cow_disposed_for_dashboard(cow_auto_id):
                continue
            
            # 牛の情報を取得
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                continue
            
            # 産次は検定時点の値を優先（event_lact）。欠損時のみ cow.lact を使用
            lact_at_test = event.get('event_lact')
            if lact_at_test is None:
                lact_at_test = cow.get('lact', 0)
            try:
                lact = int(lact_at_test)
            except (ValueError, TypeError):
                continue
            
            # 経産牛（検定時産次 >= 1）のみを対象
            if lact < 1:
                continue
            
            # 産次を分類（1産、2産、3産以上）
            if lact == 1:
                parity = 1
            elif lact == 2:
                parity = 2
            else:  # lact >= 3
                parity = 3
            
            # 検定時のDIMを計算
            clvd = clvd_for_milk_dim(cow_auto_id, cow)
            dim = calc_dim_on_date(clvd, latest_milk_test_date)
            if dim is None:
                continue
            
            # json_dataから乳量とリニアスコアを取得
            json_data_str = event.get('json_data', '{}')
            try:
                json_data = json.loads(json_data_str) if isinstance(json_data_str, str) else json_data_str
                
                milk_yield = json_data.get('milk_yield')
                if milk_yield is not None:
                    try:
                        milk_value = float(milk_yield)
                        if milk_value > 0:
                            cow_id = str(cow.get("cow_id") or "")
                            jpn10 = str(cow.get("jpn10") or "")
                            milk_points.append((dim, milk_value, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass
                
                ls_value = json_data.get('ls')
                if ls_value is not None:
                    try:
                        ls_num = float(ls_value)
                        if ls_num >= 0:
                            cow_id = str(cow.get("cow_id") or "")
                            jpn10 = str(cow.get("jpn10") or "")
                            ls_points.append((dim, ls_num, parity, cow_id, jpn10))
                    except (ValueError, TypeError):
                        pass
            except (json.JSONDecodeError, TypeError):
                pass
        
        return {
            "milk_points": milk_points,
            "ls_points": ls_points
        }
    
    def _build_cumulative_pregnancy_section(self, data):
        """累積妊娠頭数グラフ＋目標値のHTML（SVG折れ線グラフ）を返す。"""
        import math as _math
        import calendar as _calendar
        from datetime import datetime as _dt

        if not data:
            return '<div class="section-title">累積妊娠頭数</div><div class="subnote">データがありません</div>'

        annual_required     = data.get("annual_required", 0) or 0
        current_year        = data.get("current_year", _dt.now().year)
        cum_cur             = data.get("cum_cur") or []
        cum_prev            = data.get("cum_prev") or []
        latest_dt           = data.get("latest_dt")
        current_pregnancies = data.get("current_pregnancies")
        month_per           = data.get("month_per", 0) or 0
        latest_ai_et_date   = data.get("latest_ai_et_date") or ""

        # ---- SVGチャート ----
        W, H          = 640, 240
        ML, MR, MT, MB = 46, 16, 16, 32
        PW = W - ML - MR
        PH = H - MT - MB

        # Y範囲
        max_y_raw = max(annual_required, 10.0)
        if cum_cur:  max_y_raw = max(max_y_raw, max(v for _, v in cum_cur))
        if cum_prev: max_y_raw = max(max_y_raw, max(v for _, v in cum_prev))
        max_y = (int(max_y_raw) // 10 + 1) * 10

        # 月末X座標（端数）: month m の月末 = cumulative_days / total_days
        total_days_year = 366 if _calendar.isleap(current_year) else 365
        month_end_fracs = []
        cum_d = 0
        for m in range(1, 13):
            _, days = _calendar.monthrange(current_year, m)
            cum_d += days
            month_end_fracs.append(cum_d / total_days_year)

        def xf(frac):
            return ML + frac * PW

        def yv(val):
            return MT + (1.0 - val / max_y) * PH if max_y > 0 else MT + PH

        # latest_dt の年内fractional位置
        latest_frac = None
        if latest_dt and latest_dt.year == current_year:
            jan1     = _dt(current_year, 1, 1)
            dec31    = _dt(current_year, 12, 31)
            tot_span = (dec31 - jan1).days
            if tot_span > 0:
                latest_frac = min(1.0, (latest_dt - jan1).days / tot_span)

        # 実績のうち確定済み終端インデックス
        solid_end = len(cum_cur)
        if latest_dt and latest_dt.year == current_year and cum_cur:
            for i, (d_str, _) in enumerate(cum_cur):
                if _dt.strptime(d_str, "%Y-%m-%d") > latest_dt:
                    solid_end = i
                    break

        s = []
        s.append(f'<svg viewBox="0 0 {W} {H}" style="width:100%;height:auto;display:block" xmlns="http://www.w3.org/2000/svg">')

        # グリッド & Y軸ラベル
        y_step = 10 if max_y <= 100 else 20 if max_y <= 200 else 50
        y_ticks = list(range(0, int(max_y) + 1, y_step))
        for val in y_ticks:
            y = yv(val)
            s.append(f'<line x1="{ML}" y1="{y:.1f}" x2="{ML+PW}" y2="{y:.1f}" stroke="#e9ecef" stroke-width="1"/>')
            s.append(f'<text x="{ML-4}" y="{y+4:.1f}" text-anchor="end" font-size="9" fill="#666">{val}</text>')

        # X軸ラベル（奇数月のみ）
        for m in range(1, 13):
            prev_frac = month_end_fracs[m-2] if m > 1 else 0.0
            x_mid = xf((prev_frac + month_end_fracs[m-1]) / 2)
            if m % 2 == 1:
                s.append(f'<text x="{x_mid:.1f}" y="{MT+PH+24}" text-anchor="middle" font-size="9" fill="#666">{m}月</text>')

        # 軸線
        s.append(f'<line x1="{ML}" y1="{MT}" x2="{ML}" y2="{MT+PH}" stroke="#adb5bd" stroke-width="1.5"/>')
        s.append(f'<line x1="{ML}" y1="{MT+PH:.1f}" x2="{ML+PW}" y2="{MT+PH:.1f}" stroke="#adb5bd" stroke-width="1.5"/>')

        # 前年実績（緑破線）
        if cum_prev:
            pts = " ".join(
                f"{xf(month_end_fracs[i]):.1f},{yv(v):.1f}"
                for i, (_, v) in enumerate(cum_prev)
            )
            s.append(f'<polyline points="{xf(0):.1f},{yv(0):.1f} {pts}" fill="none" stroke="#2e7d32" stroke-width="1.5" stroke-dasharray="4 2" stroke-opacity="0.7"/>')

        # 目標線（青）：latest_fracまで実線、以降破線
        x0_t, y0_t = xf(0), yv(0)
        x1_t, y1_t = xf(1.0), yv(annual_required)
        if latest_frac is not None and 0 < latest_frac < 1:
            x_lat = xf(latest_frac)
            y_lat = yv(annual_required * latest_frac)
            s.append(f'<line x1="{x0_t:.1f}" y1="{y0_t:.1f}" x2="{x_lat:.1f}" y2="{y_lat:.1f}" stroke="#1565C0" stroke-width="2"/>')
            s.append(f'<line x1="{x_lat:.1f}" y1="{y_lat:.1f}" x2="{x1_t:.1f}" y2="{y1_t:.1f}" stroke="#1565C0" stroke-width="1.5" stroke-dasharray="5 3" stroke-opacity="0.5"/>')
        else:
            s.append(f'<line x1="{x0_t:.1f}" y1="{y0_t:.1f}" x2="{x1_t:.1f}" y2="{y1_t:.1f}" stroke="#1565C0" stroke-width="2"/>')

        # 今年度実績（オレンジ）
        if cum_cur and solid_end > 0:
            cur_vals = [v for _, v in cum_cur]
            pts_list = [f"{xf(0):.1f},{yv(0):.1f}"] + [
                f"{xf(month_end_fracs[i]):.1f},{yv(cur_vals[i]):.1f}"
                for i in range(solid_end)
            ]
            s.append(f'<polyline points="{" ".join(pts_list)}" fill="none" stroke="#e65100" stroke-width="2.5"/>')
            for i in range(solid_end):
                cx_ = xf(month_end_fracs[i])
                cy_ = yv(cur_vals[i])
                s.append(f'<circle cx="{cx_:.1f}" cy="{cy_:.1f}" r="3" fill="none" stroke="#c62828" stroke-width="1.8"/>')
            # 最終点：塗りつぶし＋アノテーション
            lx = xf(month_end_fracs[solid_end - 1])
            ly = yv(cur_vals[solid_end - 1])
            lv = cur_vals[solid_end - 1]
            s.append(f'<circle cx="{lx:.1f}" cy="{ly:.1f}" r="4.5" fill="#c62828"/>')
            ann_x = min(lx + 7, ML + PW - 20)
            ann_y = max(ly - 5, MT + 10)
            s.append(f'<text x="{ann_x:.1f}" y="{ann_y:.1f}" font-size="10" font-weight="bold" fill="#c62828">{lv:.0f}</text>')

        # 凡例
        lx0 = ML + 6
        ly0 = MT + 12
        s.append(f'<line x1="{lx0}" y1="{ly0}" x2="{lx0+16}" y2="{ly0}" stroke="#1565C0" stroke-width="2"/>')
        s.append(f'<text x="{lx0+20}" y="{ly0+4}" font-size="9" fill="#333">目標</text>')
        s.append(f'<line x1="{lx0+50}" y1="{ly0}" x2="{lx0+66}" y2="{ly0}" stroke="#e65100" stroke-width="2.5"/>')
        s.append(f'<text x="{lx0+70}" y="{ly0+4}" font-size="9" fill="#333">実績</text>')
        s.append(f'<line x1="{lx0+100}" y1="{ly0}" x2="{lx0+116}" y2="{ly0}" stroke="#2e7d32" stroke-width="1.5" stroke-dasharray="4 2" stroke-opacity="0.7"/>')
        s.append(f'<text x="{lx0+120}" y="{ly0+4}" font-size="9" fill="#333">前年</text>')

        s.append('</svg>')
        chart_svg = '\n'.join(s)

        # ---- 目標値ボックス ----
        date_note = f"（{latest_ai_et_date[:10].replace('-','/')} 時点）" if latest_ai_et_date else ""
        pct_html = ""
        if current_pregnancies is not None and annual_required > 0:
            pct = round(current_pregnancies / annual_required * 100)
            color = "#2e7d32" if pct >= 80 else ("#e65100" if pct >= 50 else "#c62828")
            pct_html = (
                f'<div style="margin-top:10px;text-align:center">'
                f'<span style="font-size:22px;font-weight:bold;color:{color}">{pct}%</span>'
                f'<span style="font-size:10px;color:#666;margin-left:4px">達成</span>'
                f'</div>'
            )

        stats_rows = [
            ("年間目標",   f"{annual_required:.0f} 頭"),
            ("月あたり目標", f"{month_per} 頭"),
        ]
        if current_pregnancies is not None:
            stats_rows.append(("現在の実績", f'<span style="color:#c62828;font-weight:bold">{current_pregnancies:.0f} 頭</span>'))

        rows_html = "\n".join(
            f'<div style="display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid #f0f0f0">'
            f'<span style="color:#555;font-size:11px">{k}</span>'
            f'<span style="font-weight:bold;font-size:12px">{v}</span></div>'
            for k, v in stats_rows
        )
        if date_note:
            rows_html += f'<div style="font-size:9px;color:#888;text-align:right;margin-top:2px">{date_note}</div>'

        return f"""
<div class="section-title">累積妊娠頭数（{current_year}年）</div>
<div style="display:grid;grid-template-columns:150px 1fr;gap:16px;align-items:start">
  <div style="padding-top:4px">
    {rows_html}
    {pct_html}
  </div>
  <div>
    {chart_svg}
  </div>
</div>"""

    def _build_dashboard_html(self, farm_name, parity_segments, avg_parity, total, milk_stats, fertility_stats, monthly_fertility_stats, herd_summary, scc_summary, scatter_data, insemination_count_fertility_stats, herd_dynamics_data=None, pr_hdr_data=None, repro_detail=None, heifer_monthly_fertility_stats=None, cumulative_pregnancy_data=None, dim_parity_breakdown=None, ai_et_conception_stats=None, first_ai_scatter=None, rc_dim_scatter=None, milk_alerts=None, milk_trend_html=""):
        """ダッシュボードのHTMLを生成"""
        # 乳量・リニアスコア散布図（乳検レポートと同じ仕様：Plotly、ホバーでID表示・2グラフ連動、x_min/x_max で DIM 軸を固定）
        scatter_plotly_html = ""
        if _get_plotly_scatter_script and (scatter_data.get("milk_points") or scatter_data.get("ls_points")):
            def _unpack_pt(p):
                if len(p) >= 5:
                    return (p[0], p[1], p[2], str(p[3]), str(p[4]))
                return (p[0], p[1], p[2], "", "")
            configs = []
            milk_points = scatter_data.get("milk_points", [])
            ls_points = scatter_data.get("ls_points", [])
            if milk_points:
                m_pts = [_unpack_pt(p) for p in milk_points]
                configs.append({
                    "id": "scatter-dash-milk",
                    "data": {
                        "x": [int(p[0]) for p in m_pts],
                        "y": [float(p[1]) for p in m_pts],
                        "parity": [p[2] for p in m_pts],
                        "cow_ids": [p[3] for p in m_pts],
                        "jpn10s": [p[4] for p in m_pts],
                        "title": "乳量",
                        "xlabel": "DIM",
                        "ylabel": "乳量",
                        "x_min": 0,
                        "x_max": 400,
                        "y_start_zero": True,
                    },
                })
            if ls_points:
                l_pts = [_unpack_pt(p) for p in ls_points]
                configs.append({
                    "id": "scatter-dash-ls",
                    "data": {
                        "x": [int(p[0]) for p in l_pts],
                        "y": [float(p[1]) for p in l_pts],
                        "parity": [p[2] for p in l_pts],
                        "cow_ids": [p[3] for p in l_pts],
                        "jpn10s": [p[4] for p in l_pts],
                        "title": "リニアスコア",
                        "xlabel": "DIM",
                        "ylabel": "リニアスコア",
                        "x_min": 0,
                        "x_max": 400,
                        "y_start_zero": True,
                    },
                })
            if configs:
                config_json = json.dumps(configs, ensure_ascii=False).replace("</", "<\\/")
                script = _get_plotly_scatter_script(config_json)
                parts = []
                for c in configs:
                    tid = c["id"]
                    title = "乳量" if "milk" in tid else "リニアスコア"
                    parts.append(
                        f'<div class="scatter-card">'
                        f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">'
                        f'<div class="chart-title" style="margin:0">{html.escape(title)}</div>'
                        f'<button id="{html.escape(tid)}_xhbtn" style="font-size:10px;padding:3px 10px;'
                        f'border:1px solid #e53935;border-radius:4px;background:#fff;color:#e53935;'
                        f'cursor:pointer;white-space:nowrap">＋ 基準線</button>'
                        f'</div>'
                        f'<div class="scatter-chart"><div id="{html.escape(tid)}" class="plotly-scatter-div" style="min-height:260px;"></div></div>'
                        f'<div style="display:flex;justify-content:space-between;align-items:center">'
                        f'<div class="scatter-legend">'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                        f'<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                        f'</div>'
                        f'<div id="{html.escape(tid)}_xhlabel" style="font-size:9px;color:#e53935;display:none"></div>'
                        f'</div>'
                        f'</div>'
                    )
                # 基準線（十字）JS: Plotly shapes API を利用
                crosshair_js = """
<script>
(function(){
  var plotIds = ['scatter-dash-milk','scatter-dash-ls'];
  var xhMode = {};
  var fixed = {};  // plotId -> {x, y} or null

  plotIds.forEach(function(pid){
    xhMode[pid] = false;
    fixed[pid] = null;
  });

  // マウス座標 → データ座標変換（Plotlyの内部レイアウトを使用）
  function toDataCoords(pid, clientX, clientY){
    var gd = document.getElementById(pid);
    if(!gd || !gd._fullLayout) return null;
    var layout = gd._fullLayout;
    var plotEl = gd.querySelector('.nsewdrag');
    if(!plotEl) return null;
    var pb = plotEl.getBoundingClientRect();
    var mx = clientX - pb.left;
    var my = clientY - pb.top;
    if(mx < 0 || mx > pb.width || my < 0 || my > pb.height) return null;
    var xr = layout.xaxis.range;
    var yr = layout.yaxis.range;
    return {
      x: xr[0] + (mx / pb.width)  * (xr[1] - xr[0]),
      y: yr[0] + (1 - my / pb.height) * (yr[1] - yr[0])
    };
  }

  function applyShapes(pid, curCoords){
    var shapes = [];
    var annots = [];
    var f = fixed[pid];
    // 固定線（実線・濃いめ）
    if(f){
      shapes.push({type:'line',x0:f.x,x1:f.x,y0:0,y1:1,yref:'paper',
        line:{color:'#e53935',width:1.5,dash:'dot'}});
      shapes.push({type:'line',x0:0,x1:1,xref:'paper',y0:f.y,y1:f.y,
        line:{color:'#e53935',width:1.5,dash:'dot'}});
      annots.push({x:f.x,y:1,yref:'paper',
        text:f.x.toFixed(1),showarrow:false,
        font:{size:9,color:'#e53935'},xanchor:'left',yanchor:'top'});
      annots.push({x:1,xref:'paper',y:f.y,
        text:f.y.toFixed(2),showarrow:false,
        font:{size:9,color:'#e53935'},xanchor:'right'});
    }
    // カーソル追従線（薄め・破線）
    if(curCoords){
      shapes.push({type:'line',x0:curCoords.x,x1:curCoords.x,y0:0,y1:1,yref:'paper',
        opacity:0.5,line:{color:'#e53935',width:1,dash:'dash'}});
      shapes.push({type:'line',x0:0,x1:1,xref:'paper',y0:curCoords.y,y1:curCoords.y,
        opacity:0.5,line:{color:'#e53935',width:1,dash:'dash'}});
    }
    Plotly.relayout(pid, {shapes: shapes, annotations: annots});
  }

  function initChart(pid){
    var el = document.getElementById(pid);
    if(!el || !el.on) return;
    var btn = document.getElementById(pid+'_xhbtn');
    var lbl = document.getElementById(pid+'_xhlabel');

    // ボタン：モード切替
    if(btn){
      btn.addEventListener('click', function(e){
        e.stopPropagation();
        xhMode[pid] = !xhMode[pid];
        btn.style.background = xhMode[pid] ? '#e53935' : '#fff';
        btn.style.color      = xhMode[pid] ? '#fff'    : '#e53935';
        el.style.cursor      = xhMode[pid] ? 'crosshair': '';
        if(!xhMode[pid]){
          applyShapes(pid, null);
          if(lbl) lbl.style.display = 'none';
        }
      });
    }

    // mousemove：カーソル位置に十字線を追従
    el.addEventListener('mousemove', function(e){
      if(!xhMode[pid]) return;
      var coords = toDataCoords(pid, e.clientX, e.clientY);
      if(!coords) return;
      applyShapes(pid, coords);
      if(lbl){
        lbl.style.display = '';
        lbl.textContent = 'X: ' + coords.x.toFixed(1) + '  Y: ' + coords.y.toFixed(2);
      }
    });

    // mouseleave：カーソル線を消す（固定線は維持）
    el.addEventListener('mouseleave', function(){
      if(!xhMode[pid]) return;
      applyShapes(pid, null);
      if(lbl) lbl.style.display = 'none';
    });

    // click：現在位置に十字線を固定
    el.addEventListener('click', function(e){
      if(!xhMode[pid]) return;
      var coords = toDataCoords(pid, e.clientX, e.clientY);
      if(!coords) return;
      fixed[pid] = coords;
      applyShapes(pid, null);
    });

    // ダブルクリック：固定線をリセット
    el.on('plotly_doubleclick', function(){
      fixed[pid] = null;
      applyShapes(pid, null);
    });
  }

  function tryInit(){ plotIds.forEach(function(pid){ initChart(pid); }); }
  if(document.readyState === 'complete'){ setTimeout(tryInit, 600); }
  else{ window.addEventListener('load', function(){ setTimeout(tryInit, 400); }); }
})();
</script>"""
                scatter_plotly_html = '<div class="scatter-stack">' + "".join(parts) + '</div>\n' + script + crosshair_js
        if not scatter_plotly_html and (scatter_data.get("milk_points") or scatter_data.get("ls_points")):
            milk_pts = scatter_data.get("milk_points", [])
            ls_pts = scatter_data.get("ls_points", [])
            scatter_plotly_html = (
                '<div class="scatter-stack">'
                '<div class="scatter-card"><div class="chart-title">乳量</div><div class="scatter-chart">'
                + self._build_milk_scatter_svg([(p[0], p[1], p[2]) for p in milk_pts])
                + '</div><div class="scatter-legend">'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                '</div></div><div class="scatter-card"><div class="chart-title">リニアスコア</div><div class="scatter-chart">'
                + self._build_ls_scatter_svg([(p[0], p[1], p[2]) for p in ls_pts])
                + '</div><div class="scatter-legend">'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#E91E63;"></span>1産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#00ACC1;"></span>2産</span>'
                '<span class="legend-item"><span class="legend-dot" style="background-color:#26A69A;"></span>3産以上</span>'
                '</div></div></div>'
            )

        # 産次ドーナツチャート
        parity_donut_svg = self._build_parity_donut_svg(parity_segments, avg_parity, total)
        parity_legend = self._build_parity_legend(parity_segments, total)
        
        # 平均乳量統計と乳量階層ドーナツチャート
        milk_stats_html = ""
        if milk_stats["latest_date"]:
            # 乳量階層のドーナツチャート用データを準備
            milk_segments = []
            milk_bins = milk_stats.get("milk_bins", {})
            milk_yields = milk_stats.get("milk_yields", [])
            total_milk_count = len(milk_yields)
            
            if total_milk_count > 0:
                # 乳量階層の色定義（モダンなカラーパレット）
                milk_colors = {
                    "~20kg": "#FFA07A",      # ライトサルモン
                    "20kg台": "#98D8C8",     # ミントグリーン
                    "30kg台": "#6C5CE7",     # パープル
                    "40kg台": "#A29BFE",     # ライトパープル
                    "~50kg": "#FD79A8"       # ピンク
                }
                
                # 階層の順序
                milk_order = ["~20kg", "20kg台", "30kg台", "40kg台", "~50kg"]
                radius = 50
                circumference = 2 * 3.141592653589793 * radius
                offset = 0
                
                for label in milk_order:
                    count = milk_bins.get(label, 0)
                    if count > 0:
                        percent = (count / total_milk_count * 100) if total_milk_count > 0 else 0
                        length = (percent / 100) * circumference
                        milk_segments.append({
                            "label": label,
                            "count": count,
                            "percent": round(percent, 1),
                            "color": milk_colors.get(label, "#999999"),
                            "length": length,
                            "offset": -offset
                        })
                        offset += length
                
                # 乳量階層ドーナツチャートと凡例を生成
                milk_donut_svg = self._build_milk_donut_svg(milk_segments, milk_stats["avg_milk"], total_milk_count)
                milk_legend = self._build_milk_legend(milk_segments, total_milk_count)
            else:
                milk_donut_svg = '<div class="subnote">データなし</div>'
                milk_legend = ""
            
            milk_stats_html = f"""
            <div class="milk-section">
                <div class="section-title">乳量分布（直近乳検日 {milk_stats["latest_date"]}）</div>
                <div class="milk-chart-container">
                    <div class="milk-donut">
                        {milk_donut_svg}
                    </div>
                    <div class="legend">
                        {milk_legend}
                    </div>
                </div>
            <div class="milk-stats">
                <div class="stats-label"><span class="stats-label-text">平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_milk"]:.1f}kg</span></div>
                    {f'<div class="stats-label"><span class="stats-label-text">初産平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_first_parity"]:.1f}kg</span></div>' if milk_stats.get("avg_first_parity") else ''}
                    {f'<div class="stats-label"><span class="stats-label-text">2産以上平均乳量:</span> <span class="stats-label-value">{milk_stats["avg_multiparous"]:.1f}kg</span></div>' if milk_stats.get("avg_multiparous") else ''}
                </div>
            </div>
            """
        else:
            milk_stats_html = '<div class="milk-section"><div class="section-title">乳量要約</div><div class="milk-stats"><div class="stats-label">乳検データがありません。</div></div></div>'
        
        # 受胎率表
        fertility_table_html = ""
        if fertility_stats and fertility_stats.get("stats"):
            stats = fertility_stats["stats"]
            start_date = fertility_stats["start_date"]
            end_date = fertility_stats["end_date"]
            
            # 分類をソート（産次は数値順）
            def sort_key(classification):
                if classification == "合計":
                    return (2, 0)
                try:
                    return (0, int(classification))
                except (ValueError, TypeError):
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            table_rows = []
            for classification in sorted_classifications:
                stat = stats[classification]
                total_count = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(str(classification))}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">受胎率（産次別・経産牛）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>産次</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            fertility_table_html = '<div class="fertility-section"><div class="subnote">受胎率データがありません。</div></div>'
        
        # 月ごとの受胎率表
        monthly_fertility_table_html = ""
        if monthly_fertility_stats and monthly_fertility_stats.get("stats"):
            stats = monthly_fertility_stats["stats"]
            start_date = monthly_fertility_stats["start_date"]
            end_date = monthly_fertility_stats["end_date"]
            
            # 期間内のすべての月を生成
            from datetime import datetime, timedelta
            start_dt = datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            
            # 開始月から終了月までのすべての月を生成
            all_months = []
            current_dt = datetime(start_dt.year, start_dt.month, 1)
            end_month_dt = datetime(end_dt.year, end_dt.month, 1)
            
            while current_dt <= end_month_dt:
                month_str = current_dt.strftime('%Y-%m')
                all_months.append(month_str)
                # 次の月へ
                if current_dt.month == 12:
                    current_dt = datetime(current_dt.year + 1, 1, 1)
                else:
                    current_dt = datetime(current_dt.year, current_dt.month + 1, 1)
            
            table_rows = []
            for month_str in all_months:
                # データがある場合はそのデータを使用、ない場合は0で初期化
                if month_str in stats:
                    stat = stats[month_str]
                    total_count = stat['total']
                    conceived = stat['conceived']
                    not_conceived = stat['not_conceived']
                    other = stat['other']
                else:
                    total_count = 0
                    conceived = 0
                    not_conceived = 0
                    other = 0
                
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(month_str)}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            monthly_fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">受胎率（月次推移・経産牛）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>授精月</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            monthly_fertility_table_html = '<div class="fertility-section"><div class="subnote">月ごと受胎率データがありません。</div></div>'
        
        # 授精回数ごとの受胎率表
        insemination_count_fertility_table_html = ""
        if insemination_count_fertility_stats and insemination_count_fertility_stats.get("stats"):
            stats = insemination_count_fertility_stats["stats"]
            start_date = insemination_count_fertility_stats["start_date"]
            end_date = insemination_count_fertility_stats["end_date"]
            
            # 分類をソート（授精回数は数値順）
            def sort_key(classification):
                if classification == "合計":
                    return (2, 0)
                try:
                    return (0, int(classification))
                except (ValueError, TypeError):
                    return (1, classification)
            
            sorted_classifications = sorted(stats.keys(), key=sort_key)
            
            table_rows = []
            for classification in sorted_classifications:
                stat = stats[classification]
                total_count = stat['total']
                conceived = stat['conceived']
                not_conceived = stat['not_conceived']
                other = stat['other']
                # 受胎率＝受胎÷(受胎＋不受胎)。＊は付けない（無印）
                denominator = conceived + not_conceived
                if denominator > 0:
                    rate = (conceived / denominator) * 100
                    rate_str = f"{rate:.1f}%"
                else:
                    rate_str = "0.0%"
                other_str = str(other)
                
                table_rows.append(
                    f'<tr>'
                    f'<td class="num">{html.escape(str(classification))}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(other_str)}</td>'
                    f'<td class="num">{html.escape(str(total_count))}</td>'
                    f'</tr>'
                )
            
            # 合計行を追加
            total_conceived = sum(s['conceived'] for s in stats.values())
            total_not_conceived = sum(s['not_conceived'] for s in stats.values())
            total_other = sum(s['other'] for s in stats.values())
            total_all = sum(s['total'] for s in stats.values())
            total_denom = total_conceived + total_not_conceived
            if total_denom > 0:
                total_rate = (total_conceived / total_denom) * 100
                total_rate_str = f"{total_rate:.1f}%"
            else:
                total_rate_str = "0.0%"
            total_other_str = str(total_other)
            
            table_rows.append(
                f'<tr class="total-row">'
                f'<td class="num"><strong>合計</strong></td>'
                f'<td class="num"><strong>{html.escape(total_rate_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_not_conceived))}</strong></td>'
                f'<td class="num"><strong>{html.escape(total_other_str)}</strong></td>'
                f'<td class="num"><strong>{html.escape(str(total_all))}</strong></td>'
                f'</tr>'
            )
            
            insemination_count_fertility_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">受胎率（授精回数別・経産牛）</div>
                <div class="subheader">{start_date} ～ {end_date}</div>
                <table class="summary-table" style="width: 100%; max-width: 600px;">
                    <thead>
                        <tr>
                            <th>授精回数</th>
                            <th>受胎率</th>
                            <th>受胎</th>
                            <th>不受胎</th>
                            <th>その他</th>
                            <th>総数</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(table_rows)}
                    </tbody>
                </table>
            </div>
            """
        else:
            insemination_count_fertility_table_html = '<div class="fertility-section"><div class="subnote">授精回数ごと受胎率データがありません。</div></div>'
        
        # 体細胞要約
        scc_summary_html = ""
        if scc_summary and scc_summary.get("latest_date"):
            # リニアスコア階層のドーナツチャート用データを準備
            ls_segments = []
            ls_bins = scc_summary.get("ls_bins", {})
            ls_values = scc_summary.get("ls_values", [])
            total_ls_count = len(ls_values)
            
            if total_ls_count > 0:
                # リニアスコア階層の色定義（モダンなカラーパレット）
                ls_colors = {
                    "LS2以下": "#4169E1",      # ブルー
                    "LS3-4": "#FFD700",        # イエロー
                    "LS5以上": "#FF6B6B"        # レッド
                }
                
                # 階層の順序
                ls_order = ["LS2以下", "LS3-4", "LS5以上"]
                radius = 50
                circumference = 2 * 3.141592653589793 * radius
                offset = 0
                
                for label in ls_order:
                    count = ls_bins.get(label, 0)
                    if count > 0:
                        percent = (count / total_ls_count * 100) if total_ls_count > 0 else 0
                        length = (percent / 100) * circumference
                        ls_segments.append({
                            "label": label,
                            "count": count,
                            "percent": round(percent, 1),
                            "color": ls_colors.get(label, "#999999"),
                            "length": length,
                            "offset": -offset
                        })
                        offset += length
                
                # リニアスコア階層ドーナツチャートと凡例を生成
                ls_donut_svg = self._build_ls_donut_svg(ls_segments, scc_summary["avg_ls"], total_ls_count)
                ls_legend = self._build_ls_legend(ls_segments, total_ls_count)
            else:
                ls_donut_svg = '<div class="subnote">データなし</div>'
                ls_legend = ""
            
            # 平均体細胞を計算（千単位で表示）
            import math
            avg_scc_display = None
            if scc_summary.get("avg_scc") is not None:
                avg_scc_display = int(math.floor(scc_summary["avg_scc"] + 0.5))
            
            scc_summary_html = f"""
            <div class="scc-section">
                <div class="section-title">体細胞スコア分布（直近乳検日 {scc_summary["latest_date"]}）</div>
                <div class="milk-chart-container">
                    <div class="milk-donut">
                        {ls_donut_svg}
                    </div>
                    <div class="legend">
                        {ls_legend}
                    </div>
                </div>
                <div class="milk-stats">
                    <div class="stats-label"><span class="stats-label-text">平均体細胞:</span> <span class="stats-label-value">{f"{avg_scc_display}千" if avg_scc_display is not None else "-"}</span></div>
                    <div class="stats-label"><span class="stats-label-text">平均リニアスコア:</span> <span class="stats-label-value">{f"{scc_summary['avg_ls']:.1f}" if scc_summary.get("avg_ls") is not None else "-"}</span></div>
                    {f'<div class="stats-label"><span class="stats-label-text">初産平均リニアスコア:</span> <span class="stats-label-value">{scc_summary["avg_first_parity_ls"]:.1f}</span></div>' if scc_summary.get("avg_first_parity_ls") is not None else ''}
                    {f'<div class="stats-label"><span class="stats-label-text">2産以上平均リニアスコア:</span> <span class="stats-label-value">{scc_summary["avg_multiparous_ls"]:.1f}</span></div>' if scc_summary.get("avg_multiparous_ls") is not None else ''}
                </div>
            </div>
            """
        else:
            scc_summary_html = '<div class="scc-section"><div class="section-title">体細胞要約</div><div class="milk-stats"><div class="stats-label">乳検データがありません。</div></div></div>'
        
        # （PR/HDRは妊娠率カードで表示するため牛群要約からは除外）

        # 未経産牛 月別受胎率テーブル
        heifer_monthly_table_html = ""
        if heifer_monthly_fertility_stats and heifer_monthly_fertility_stats.get("stats"):
            h_stats = heifer_monthly_fertility_stats["stats"]
            h_start = heifer_monthly_fertility_stats["start_date"]
            h_end   = heifer_monthly_fertility_stats["end_date"]
            from datetime import datetime as _dt
            _s = _dt.strptime(h_start, '%Y-%m-%d')
            _e = _dt.strptime(h_end,   '%Y-%m-%d')
            _months = []
            _cur = _dt(_s.year, _s.month, 1)
            _end_m = _dt(_e.year, _e.month, 1)
            while _cur <= _end_m:
                _months.append(_cur.strftime('%Y-%m'))
                _cur = _dt(_cur.year + (_cur.month == 12), (_cur.month % 12) + 1, 1)
            h_rows = ""
            for ym in _months:
                st = h_stats.get(ym, {})
                conceived = st.get('conceived', 0)
                not_conceived = st.get('not_conceived', 0)
                other = st.get('other', 0)
                total_c = st.get('total', 0)
                denom = conceived + not_conceived
                rate_str = f"{conceived/denom*100:.1f}%" if denom > 0 else "0.0%"
                h_rows += (
                    f'<tr><td class="num">{html.escape(ym)}</td>'
                    f'<td class="num">{html.escape(rate_str)}</td>'
                    f'<td class="num">{html.escape(str(conceived))}</td>'
                    f'<td class="num">{html.escape(str(not_conceived))}</td>'
                    f'<td class="num">{html.escape(str(other))}</td>'
                    f'<td class="num">{html.escape(str(total_c))}</td></tr>\n'
                )
            heifer_monthly_table_html = f"""
            <div class="fertility-section">
                <div class="section-title">受胎率（月次推移・未経産牛）</div>
                <div class="subheader">{h_start} ～ {h_end}</div>
                <table class="summary-table" style="width:100%;max-width:600px">
                    <thead><tr>
                        <th>授精月</th><th>受胎率</th><th>受胎</th><th>不受胎</th><th>その他</th><th>総数</th>
                    </tr></thead>
                    <tbody>{h_rows}</tbody>
                </table>
            </div>"""
        else:
            heifer_monthly_table_html = '<div class="fertility-section"><div class="subnote">未経産牛受胎率データがありません。</div></div>'

        # 経産牛 妊娠状況 ドーナツグラフ + DIM×産次内訳テーブル
        import math as _math
        pregnancy_donut_html = ""
        _h_pregnant = herd_summary.get("pregnant_count", 0) if herd_summary else 0
        _h_total    = herd_summary.get("total_count", 0) if herd_summary else 0
        if _h_total > 0:
            _h_open  = _h_total - _h_pregnant
            _h_pct   = _h_pregnant / _h_total * 100
            _cx, _cy, _r, _sw = 80, 80, 52, 24
            _circ    = 2 * _math.pi * _r
            _preg_d  = _h_pregnant / _h_total * _circ
            # viewBox十分広く（凡例が切れないよう 280px）
            _donut_svg = (
                f'<svg viewBox="0 0 280 160" style="width:100%;max-width:280px;height:auto;display:block;margin:0 auto" xmlns="http://www.w3.org/2000/svg">'
                f'<circle cx="{_cx}" cy="{_cy}" r="{_r}" fill="none" stroke="#dee2e6" stroke-width="{_sw}"/>'
                f'<circle cx="{_cx}" cy="{_cy}" r="{_r}" fill="none" stroke="#1565C0" stroke-width="{_sw}" '
                f'stroke-dasharray="{_preg_d:.2f} {_circ:.2f}" transform="rotate(-90 {_cx} {_cy})"/>'
                f'<text x="{_cx}" y="{_cy - 8}" text-anchor="middle" font-size="18" font-weight="bold" fill="#1565C0">{_h_pct:.0f}%</text>'
                f'<text x="{_cx}" y="{_cy + 11}" text-anchor="middle" font-size="9" fill="#555">妊娠中</text>'
                f'<rect x="152" y="42" width="11" height="11" fill="#1565C0" rx="2"/>'
                f'<text x="167" y="52" font-size="10" fill="#333">妊娠・乾乳 {_h_pregnant}頭</text>'
                f'<rect x="152" y="60" width="11" height="11" fill="#dee2e6" rx="2"/>'
                f'<text x="167" y="70" font-size="10" fill="#333">未妊娠 {_h_open}頭</text>'
                f'<text x="167" y="86" font-size="10" fill="#888">計 {_h_total}頭</text>'
                f'</svg>'
            )
            # DIM別 妊娠ステータス バーチャート + 産次テーブル
            _dpb = dim_parity_breakdown
            if _dpb and _dpb.get("grand_total", 0) > 0:
                _bl  = _dpb["bin_labels"]
                _pl  = _dpb["parity_labels"]
                _dt  = _dpb["data"]
                _rt  = _dpb["row_totals"]
                _ct  = _dpb["col_totals"]
                _gt  = _dpb["grand_total"]
                _sd  = _dpb.get("status_data", [{"pregnant":0,"open":0,"dnb":0}]*len(_bl))
                _max_total = max(_rt) if _rt else 1

                # ── 横積み上げバーチャート ──
                _bar_rows = ""
                for i, blabel in enumerate(_bl):
                    if _rt[i] == 0:
                        continue
                    pg = _sd[i]["pregnant"]
                    op = _sd[i]["open"]
                    dn = _sd[i]["dnb"]
                    tot = _rt[i]
                    # 幅をmax_totalで正規化（最長バーが100%）
                    _scale = 100 / _max_total
                    pg_w = pg * _scale
                    op_w = op * _scale
                    dn_w = dn * _scale
                    # 各セグメント内にラベル（2頭以上のみ表示）
                    def _seg(w, color, count, label_color="#fff"):
                        if w < 0.1:
                            return ""
                        txt = str(count) if count >= 2 else ""
                        return (
                            f'<div style="width:{w:.1f}%;background:{color};display:flex;'
                            f'align-items:center;justify-content:center;'
                            f'font-size:9px;font-weight:600;color:{label_color};'
                            f'overflow:hidden;white-space:nowrap;min-width:0">{txt}</div>'
                        )
                    _bar_rows += (
                        f'<div style="display:flex;align-items:center;margin-bottom:3px">'
                        f'<div style="width:58px;font-size:9px;color:#546e7a;text-align:right;'
                        f'padding-right:7px;white-space:nowrap;flex-shrink:0">{blabel}日</div>'
                        f'<div style="flex:1;height:16px;display:flex;border-radius:3px;overflow:hidden;background:#f0f0f0">'
                        f'{_seg(pg_w,"#1565C0",pg)}'
                        f'{_seg(op_w,"#EF6C00",op)}'
                        f'{_seg(dn_w,"#9E9E9E",dn,"#fff")}'
                        f'</div>'
                        f'<div style="width:28px;font-size:9px;color:#333;text-align:right;'
                        f'padding-left:5px;flex-shrink:0;font-weight:600">{tot}</div>'
                        f'</div>'
                    )

                # 凡例
                _legend = (
                    f'<div style="display:flex;gap:12px;margin-bottom:6px;flex-wrap:wrap">'
                    f'<span style="display:flex;align-items:center;gap:4px;font-size:9px;color:#555">'
                    f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#1565C0"></span>妊娠</span>'
                    f'<span style="display:flex;align-items:center;gap:4px;font-size:9px;color:#555">'
                    f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#EF6C00"></span>未受胎</span>'
                    f'<span style="display:flex;align-items:center;gap:4px;font-size:9px;color:#555">'
                    f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#9E9E9E"></span>繁殖中止</span>'
                    f'</div>'
                )

                # ── 産次別テーブル（コンパクト） ──
                _th = "".join(
                    f'<th style="padding:2px 5px;font-size:9px;text-align:center;background:#f8f9fa;border:1px solid #dee2e6">{p}</th>'
                    for p in _pl
                )
                _tbody = ""
                for i, blabel in enumerate(_bl):
                    if _rt[i] == 0:
                        continue
                    _cells = "".join(
                        f'<td style="padding:2px 4px;font-size:9px;text-align:center;border:1px solid #e9ecef">{_dt[i][j] if _dt[i][j] else ""}</td>'
                        for j in range(len(_pl))
                    )
                    _tbody += (
                        f'<tr>'
                        f'<td style="padding:2px 4px;font-size:9px;white-space:nowrap;border:1px solid #e9ecef;background:#f8f9fa;color:#546e7a">{blabel}日</td>'
                        f'{_cells}'
                        f'<td style="padding:2px 4px;font-size:9px;text-align:center;font-weight:bold;border:1px solid #e9ecef">{_rt[i]}</td>'
                        f'</tr>'
                    )
                _foot_cells = "".join(
                    f'<td style="padding:2px 4px;font-size:9px;text-align:center;font-weight:bold;border:1px solid #dee2e6">{_ct[j] if _ct[j] else ""}</td>'
                    for j in range(len(_pl))
                )

                _dim_table_html = (
                    f'<div style="margin-top:12px">'
                    # バーチャートセクション
                    f'<div style="font-size:10px;color:#546e7a;font-weight:600;margin-bottom:6px">DIM別 妊娠ステータス</div>'
                    f'{_legend}'
                    f'{_bar_rows}'
                    # 産次テーブル
                    f'<div style="font-size:9px;color:#90a4ae;margin:10px 0 4px">DIM別・産次別内訳（頭）</div>'
                    f'<table style="border-collapse:collapse;width:100%">'
                    f'<thead><tr>'
                    f'<th style="padding:2px 4px;font-size:9px;background:#f8f9fa;border:1px solid #dee2e6">DIM</th>'
                    f'{_th}'
                    f'<th style="padding:2px 4px;font-size:9px;background:#f8f9fa;border:1px solid #dee2e6;text-align:center">計</th>'
                    f'</tr></thead>'
                    f'<tbody>{_tbody}</tbody>'
                    f'<tfoot><tr>'
                    f'<td style="padding:2px 4px;font-size:9px;font-weight:bold;background:#f8f9fa;border:1px solid #dee2e6">計</td>'
                    f'{_foot_cells}'
                    f'<td style="padding:2px 4px;font-size:9px;font-weight:bold;text-align:center;border:1px solid #dee2e6">{_gt}</td>'
                    f'</tr></tfoot>'
                    f'</table>'
                    f'</div>'
                )
            else:
                _dim_table_html = ""

            pregnancy_donut_html = (
                f'<div style="margin-top:10px;border-top:1px solid #e9ecef;padding-top:10px">'
                f'<div class="section-title" style="font-size:12px;margin-bottom:6px">経産牛 妊娠状況</div>'
                f'{_donut_svg}'
                f'{_dim_table_html}'
                f'</div>'
            )

        # 累積妊娠頭数グラフ（SVG版）
        cumulative_preg_section_html = self._build_cumulative_pregnancy_section(cumulative_pregnancy_data)

        # 授精種類別受胎率テーブル
        ai_et_table_html = ""
        if ai_et_conception_stats:
            rows = ""
            for r in ai_et_conception_stats:
                cr_str = f'{r["cr"]}%' if r["cr"] is not None else "—"
                cr_color = "#1565C0" if (r["cr"] or 0) >= 50 else ("#EF6C00" if (r["cr"] or 0) >= 30 else "#c62828")
                cr_style = f'color:{cr_color};font-weight:bold' if r["cr"] is not None else 'color:#999'
                rows += (
                    f'<tr>'
                    f'<td style="padding:5px 8px;font-size:11px;border-bottom:1px solid #f0f0f0">{r["name"]}</td>'
                    f'<td style="padding:5px 8px;font-size:11px;text-align:center;border-bottom:1px solid #f0f0f0">{r["count"]}</td>'
                    f'<td style="padding:5px 8px;font-size:11px;text-align:center;border-bottom:1px solid #f0f0f0">{r["pregnant"]}</td>'
                    f'<td style="padding:5px 8px;font-size:11px;text-align:center;border-bottom:1px solid #f0f0f0;{cr_style}">{cr_str}</td>'
                    f'</tr>'
                )
            ai_et_table_html = (
                f'<div class="section-title">授精種類別受胎率（経産牛）</div>'
                f'<table style="border-collapse:collapse;width:100%;margin-top:8px">'
                f'<thead><tr>'
                f'<th style="padding:5px 8px;font-size:11px;background:#f8f9fa;border-bottom:2px solid #dee2e6;text-align:left">授精種類</th>'
                f'<th style="padding:5px 8px;font-size:11px;background:#f8f9fa;border-bottom:2px solid #dee2e6;text-align:center">授精数</th>'
                f'<th style="padding:5px 8px;font-size:11px;background:#f8f9fa;border-bottom:2px solid #dee2e6;text-align:center">受胎数</th>'
                f'<th style="padding:5px 8px;font-size:11px;background:#f8f9fa;border-bottom:2px solid #dee2e6;text-align:center">受胎率</th>'
                f'</tr></thead>'
                f'<tbody>{rows}</tbody>'
                f'</table>'
            )

        # 産次別カラーヘルパー
        _parity_color = lambda p: "#E91E63" if p.get("lact") == 1 else "#00ACC1" if p.get("lact") == 2 else "#26A69A"
        _parity_legend = '<span style="color:#E91E63;font-size:9px">● 1産</span> <span style="color:#00ACC1;font-size:9px">● 2産</span> <span style="color:#26A69A;font-size:9px">● 3産以上</span>'

        # 初回授精DIM × 分娩月日 散布図
        first_ai_scatter_html = _build_inline_scatter(
            points=first_ai_scatter or [],
            chart_id="scatter_first_ai",
            title="初回授精日数 × 分娩月日",
            x_label="分娩月日",
            y_label="初回授精DIM（日）",
            x_key="x_label",
            y_key="y",
            color_fn=_parity_color,
            legend_html=_parity_legend,
            hline_default=None,
        )

        # 繁殖状態 × DIM 散布図
        rc_labels_map = {0:"繁殖停止", 1:"Fresh", 2:"空胎", 3:"授精後", 4:"妊娠中", 5:"乾乳"}
        rc_dim_scatter_html = _build_inline_scatter(
            points=rc_dim_scatter or [],
            chart_id="scatter_rc_dim",
            title="繁殖状態 × DIM",
            x_label="DIM（日）",
            y_label="繁殖状態",
            x_key="x",
            y_key="y",
            color_fn=_parity_color,
            legend_html=_parity_legend,
            y_tick_labels=rc_labels_map,
        )

        # 分娩予定月別・産子種類内訳グラフ（牛群動態）
        herd_dynamics_chart_html = self._build_herd_dynamics_chart_section(herd_dynamics_data)
        
        # 妊娠率サイクル表 HTML
        pr_cycle_table_html = ""
        pr_dim_chart_html = ""
        if repro_detail:
            cycles = repro_detail.get("cycles") or []
            dim_results = repro_detail.get("dim_pr") or []
            pd_start = repro_detail.get("start_date", "")
            pd_end = repro_detail.get("end_date", "")

            # --- サイクル表 ---
            rows_html = ""
            total_br_el = total_bred = total_preg_el = total_preg = total_loss = 0
            for r in cycles:
                sd = getattr(r, "start_date", "") or ""
                ed = getattr(r, "end_date", "") or ""
                def _short(d):
                    try:
                        parts = d.split("-")
                        return f"{int(parts[1])}/{int(parts[2])}"
                    except Exception:
                        return d
                br_el = getattr(r, "br_el", 0) or 0
                bred  = getattr(r, "bred", 0) or 0
                hdr_v = getattr(r, "hdr", None)
                preg_el = getattr(r, "preg_eligible", 0) or 0
                preg  = getattr(r, "preg", 0) or 0
                pr_v  = getattr(r, "pr", None)
                loss  = getattr(r, "loss", 0) or 0
                total_br_el += br_el; total_bred += bred
                total_preg_el += preg_el; total_preg += preg; total_loss += loss
                hdr_str = f"{int(hdr_v)}%" if hdr_v is not None else "-"
                pr_str  = f"{int(pr_v)}%"  if pr_v  is not None else "-"
                row_style = ' style="color:#aaa"' if preg_el == 0 else ""
                rows_html += (
                    f'<tr{row_style}>'
                    f'<td class="num">{html.escape(_short(sd))}</td>'
                    f'<td class="num">{html.escape(_short(ed))}</td>'
                    f'<td class="num">{br_el}</td>'
                    f'<td class="num">{bred}</td>'
                    f'<td class="num">{hdr_str}</td>'
                    f'<td class="num">{preg_el if preg_el else "-"}</td>'
                    f'<td class="num">{preg if preg_el else "-"}</td>'
                    f'<td class="num">{pr_str if preg_el else "-"}</td>'
                    f'<td class="num">{loss}</td>'
                    f'</tr>\n'
                )
            total_hdr_str = f"{round(total_bred/total_br_el*100)}%" if total_br_el else "-"
            total_pr_str  = f"{round(total_preg/total_preg_el*100)}%" if total_preg_el else "-"
            rows_html += (
                f'<tr class="total-row">'
                f'<td class="num" colspan="2"><strong>【合計】</strong></td>'
                f'<td class="num"><strong>{total_br_el}</strong></td>'
                f'<td class="num"><strong>{total_bred}</strong></td>'
                f'<td class="num"><strong>{total_hdr_str}</strong></td>'
                f'<td class="num"><strong>{total_preg_el}</strong></td>'
                f'<td class="num"><strong>{total_preg}</strong></td>'
                f'<td class="num"><strong>{total_pr_str}</strong></td>'
                f'<td class="num"><strong>{total_loss}</strong></td>'
                f'</tr>\n'
            )
            pr_cycle_table_html = f"""
<div class="section-title">妊娠率（21日サイクル）</div>
<div class="subheader">{html.escape(pd_start)} ～ {html.escape(pd_end)}　VWP=50日</div>
<table class="summary-table" style="width:100%;font-size:11px">
  <thead>
    <tr>
      <th>開始日</th><th>終了日</th><th>繁殖対象</th><th>授精</th><th>授精率%</th>
      <th>妊娠対象</th><th>妊娠</th><th>妊娠率%</th><th>損耗</th>
    </tr>
  </thead>
  <tbody>{rows_html}</tbody>
</table>"""

            # --- DIM別妊娠率複合チャート（積み上げ棒グラフ＋空胎残存率） ---
            if dim_results:
                import math as _math

                # SVG寸法
                CW, CH = 520, 300
                CPL, CPR, CPT, CPB = 44, 44, 28, 36
                CIW = CW - CPL - CPR
                CIH = CH - CPT - CPB

                # DIM範囲のX座標マッピング
                all_starts = [getattr(r, "dim_start", 0) for r in dim_results]
                all_ends   = [getattr(r, "dim_end",   9999) for r in dim_results]
                eff_ends = [ds + 30 if de >= 9999 else de for ds, de in zip(all_starts, all_ends)]
                x_dim_min = all_starts[0] if all_starts else 50
                x_dim_max = max(eff_ends) + 20 if eff_ends else 300
                x_dim_range = max(x_dim_max - x_dim_min, 1)

                def _cx(d): return CPL + (d - x_dim_min) / x_dim_range * CIW
                def _cy(p): return CPT + CIH * (1.0 - p / 100.0)

                # グリッドと両Y軸ラベル
                grid_svg = ""
                for pct in range(0, 101, 10):
                    yp = _cy(pct)
                    grid_svg += f'<line x1="{CPL:.1f}" y1="{yp:.1f}" x2="{CW-CPR:.1f}" y2="{yp:.1f}" stroke="#e0e0e0" stroke-width="0.5"/>\n'
                    grid_svg += f'<text x="{CPL-4}" y="{yp+3:.1f}" text-anchor="end" font-size="9" fill="#555">{pct}</text>\n'
                    grid_svg += f'<text x="{CW-CPR+4}" y="{yp+3:.1f}" text-anchor="start" font-size="9" fill="#e65100">{pct}</text>\n'
                # 軸線
                grid_svg += (
                    f'<line x1="{CPL}" y1="{CPT}" x2="{CPL}" y2="{CPT+CIH}" stroke="#aaa" stroke-width="1"/>\n'
                    f'<line x1="{CW-CPR}" y1="{CPT}" x2="{CW-CPR}" y2="{CPT+CIH}" stroke="#e65100" stroke-width="1"/>\n'
                    f'<line x1="{CPL}" y1="{CPT+CIH}" x2="{CW-CPR}" y2="{CPT+CIH}" stroke="#aaa" stroke-width="1"/>\n'
                )

                # 積み上げ棒グラフとエラーバー
                bars_svg = ""
                x_label_svg = ""
                surv_pts = [(x_dim_min, 100.0)]
                s_open = 100.0
                for r in dim_results:
                    ds   = getattr(r, "dim_start",     0)
                    de   = getattr(r, "dim_end",     9999)
                    de_e = ds + 30 if de >= 9999 else de
                    hdr_val  = getattr(r, "hdr",          0) or 0
                    pr_val   = getattr(r, "pr",           0) or 0
                    preg_el  = getattr(r, "preg_eligible",0) or 0
                    bw = (de_e - ds) * CIW / x_dim_range * 0.72
                    bx = _cx((ds + de_e) / 2) - bw / 2

                    # 薄青バー（HDR-PR の差分部分）
                    diff = max(hdr_val - pr_val, 0)
                    if diff > 0:
                        dh = diff / 100 * CIH
                        dy = _cy(hdr_val)
                        bars_svg += f'<rect x="{bx:.1f}" y="{dy:.1f}" width="{bw:.1f}" height="{dh:.1f}" fill="#90caf9"/>\n'
                    # 濃青バー（妊娠率）
                    if pr_val > 0:
                        ph = pr_val / 100 * CIH
                        py = _cy(pr_val)
                        bars_svg += f'<rect x="{bx:.1f}" y="{py:.1f}" width="{bw:.1f}" height="{ph:.1f}" fill="#1565c0"/>\n'
                    # 95%信頼区間エラーバー
                    if preg_el > 0 and pr_val > 0:
                        p_frac = pr_val / 100.0
                        ci = 1.96 * _math.sqrt(p_frac * (1 - p_frac) / preg_el) * 100
                        ci_top = min(100.0, pr_val + ci)
                        ci_bot = max(0.0,   pr_val - ci)
                        ex = bx + bw / 2
                        bars_svg += (
                            f'<line x1="{ex:.1f}" y1="{_cy(ci_top):.1f}" x2="{ex:.1f}" y2="{_cy(ci_bot):.1f}" stroke="#1565c0" stroke-width="1.5"/>\n'
                            f'<line x1="{ex-3:.1f}" y1="{_cy(ci_top):.1f}" x2="{ex+3:.1f}" y2="{_cy(ci_top):.1f}" stroke="#1565c0" stroke-width="1.5"/>\n'
                            f'<line x1="{ex-3:.1f}" y1="{_cy(ci_bot):.1f}" x2="{ex+3:.1f}" y2="{_cy(ci_bot):.1f}" stroke="#1565c0" stroke-width="1.5"/>\n'
                        )
                    # X軸ラベル（dim_start）
                    lx = _cx(ds)
                    x_label_svg += f'<text x="{lx:.1f}" y="{CPT+CIH+14}" text-anchor="middle" font-size="9" fill="#666">{ds}</text>\n'

                    # 空胎残存率の次のポイント
                    s_open *= (1 - pr_val / 100.0)
                    surv_pts.append((de_e, round(s_open, 1)))

                # 空胎残存率ライン
                surv_svg = ""
                surv_path = " ".join(
                    f"{'M' if i == 0 else 'L'}{_cx(p[0]):.1f},{_cy(p[1]):.1f}"
                    for i, p in enumerate(surv_pts)
                )
                surv_svg += f'<path d="{surv_path}" fill="none" stroke="#e65100" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>\n'
                surv_svg += "".join(
                    f'<circle cx="{_cx(p[0]):.1f}" cy="{_cy(p[1]):.1f}" r="3" fill="#e65100" stroke="#fff" stroke-width="1.2"/>\n'
                    for p in surv_pts
                )

                # 凡例
                legend_y = CPT - 10
                legend_svg = (
                    f'<rect x="{CPL}" y="{legend_y-8}" width="12" height="10" fill="#90caf9" stroke="#aaa" stroke-width="0.5"/>'
                    f'<text x="{CPL+15}" y="{legend_y}" font-size="9" fill="#333">授精率（差分）</text>'
                    f'<rect x="{CPL+95}" y="{legend_y-8}" width="12" height="10" fill="#1565c0"/>'
                    f'<text x="{CPL+110}" y="{legend_y}" font-size="9" fill="#333">妊娠率</text>'
                    f'<line x1="{CPL+160}" y1="{legend_y-4}" x2="{CPL+175}" y2="{legend_y-4}" stroke="#e65100" stroke-width="2.2"/>'
                    f'<circle cx="{CPL+168}" cy="{legend_y-4}" r="3" fill="#e65100"/>'
                    f'<text x="{CPL+178}" y="{legend_y}" font-size="9" fill="#333">空胎残存率</text>'
                )

                # Y軸タイトル・X軸タイトル
                axis_titles = (
                    f'<text x="10" y="{CPT + CIH//2}" text-anchor="middle" font-size="9" fill="#555" '
                    f'transform="rotate(-90,10,{CPT + CIH//2})">割合 (%)</text>\n'
                    f'<text x="{CW-8}" y="{CPT + CIH//2}" text-anchor="middle" font-size="9" fill="#e65100" '
                    f'transform="rotate(-90,{CW-8},{CPT + CIH//2})">空胎残存率 (%)</text>\n'
                    f'<text x="{CW//2}" y="{CH-2}" text-anchor="middle" font-size="9" fill="#666">泌乳日数（DIM）</text>\n'
                )

                # クロスヘア用プレースホルダー線
                dim_chart_id = "dim_pr_chart"
                crosshair_overlay = (
                    f'<line id="{dim_chart_id}_hline" x1="{CPL}" y1="-1000" x2="{CW-CPR}" y2="-1000" stroke="#e53935" stroke-width="1.2" stroke-dasharray="5 3" opacity="0.8"/>'
                    f'<line id="{dim_chart_id}_vline" x1="-1000" y1="{CPT}" x2="-1000" y2="{CPT+CIH}" stroke="#e53935" stroke-width="1.2" stroke-dasharray="5 3" opacity="0.8"/>'
                    f'<text id="{dim_chart_id}_hlabel" x="{CPL+CIW-2}" y="-1000" text-anchor="end" font-size="8" fill="#e53935"></text>'
                    f'<text id="{dim_chart_id}_vlabel" x="-1000" y="{CPT+11}" text-anchor="middle" font-size="8" fill="#e53935"></text>'
                    f'<line id="{dim_chart_id}_chline" x1="{CPL}" y1="-1000" x2="{CW-CPR}" y2="-1000" stroke="#e53935" stroke-width="1" stroke-dasharray="3 2" opacity="0.5"/>'
                    f'<line id="{dim_chart_id}_cvline" x1="-1000" y1="{CPT}" x2="-1000" y2="{CPT+CIH}" stroke="#e53935" stroke-width="1" stroke-dasharray="3 2" opacity="0.5"/>'
                )

                combo_svg = (
                    f'<svg id="{dim_chart_id}_svg" viewBox="0 0 {CW} {CH}" style="width:100%;height:auto;display:block;cursor:default" xmlns="http://www.w3.org/2000/svg">\n'
                    f'{grid_svg}{bars_svg}{surv_svg}{x_label_svg}{legend_svg}{axis_titles}{crosshair_overlay}'
                    f'</svg>'
                )

                dim_crosshair_js = f"""
<script>
(function(){{
  var cid = '{dim_chart_id}';
  var svgEl = document.getElementById(cid + '_svg');
  if (!svgEl) return;
  var active = false;
  var CPL={CPL}, CPR={CPR}, CPT={CPT}, CW={CW}, CH={CH}, CIW={CIW}, CIH={CIH};
  var xMin={x_dim_min}, xMax={x_dim_max};
  var btn = document.getElementById(cid + '_btn');

  function svgPt(e) {{
    var r = svgEl.getBoundingClientRect();
    return {{
      x: (e.clientX - r.left) * CW / r.width,
      y: (e.clientY - r.top)  * CH / r.height
    }};
  }}
  function toDataX(sx) {{ return xMin + (sx - CPL) / CIW * (xMax - xMin); }}
  function toDataY(sy) {{ return 100 * (1 - (sy - CPT) / CIH); }}

  if (btn) btn.addEventListener('click', function() {{
    active = !active;
    btn.style.background = active ? '#e53935' : '';
    btn.style.color      = active ? '#fff'    : '';
    svgEl.style.cursor   = active ? 'crosshair' : 'default';
    if (!active) {{
      document.getElementById(cid+'_chline').setAttribute('y1','-1000');
      document.getElementById(cid+'_chline').setAttribute('y2','-1000');
      document.getElementById(cid+'_cvline').setAttribute('x1','-1000');
      document.getElementById(cid+'_cvline').setAttribute('x2','-1000');
    }}
  }});

  svgEl.addEventListener('mousemove', function(e) {{
    if (!active) return;
    var p = svgPt(e);
    var sx = Math.max(CPL, Math.min(CPL+CIW, p.x));
    var sy = Math.max(CPT, Math.min(CPT+CIH, p.y));
    document.getElementById(cid+'_chline').setAttribute('y1', sy.toFixed(1));
    document.getElementById(cid+'_chline').setAttribute('y2', sy.toFixed(1));
    document.getElementById(cid+'_cvline').setAttribute('x1', sx.toFixed(1));
    document.getElementById(cid+'_cvline').setAttribute('x2', sx.toFixed(1));
  }});

  svgEl.addEventListener('mouseleave', function() {{
    if (!active) return;
    document.getElementById(cid+'_chline').setAttribute('y1','-1000');
    document.getElementById(cid+'_chline').setAttribute('y2','-1000');
    document.getElementById(cid+'_cvline').setAttribute('x1','-1000');
    document.getElementById(cid+'_cvline').setAttribute('x2','-1000');
  }});

  svgEl.addEventListener('click', function(e) {{
    if (!active) return;
    var p = svgPt(e);
    var sx = Math.max(CPL, Math.min(CPL+CIW, p.x));
    var sy = Math.max(CPT, Math.min(CPT+CIH, p.y));
    document.getElementById(cid+'_hline').setAttribute('y1', sy.toFixed(1));
    document.getElementById(cid+'_hline').setAttribute('y2', sy.toFixed(1));
    document.getElementById(cid+'_vline').setAttribute('x1', sx.toFixed(1));
    document.getElementById(cid+'_vline').setAttribute('x2', sx.toFixed(1));
    var dy = toDataY(sy);
    var dx = toDataX(sx);
    var hl = document.getElementById(cid+'_hlabel');
    hl.setAttribute('y', (sy - 3).toFixed(1));
    hl.textContent = dy.toFixed(1) + '%';
    var vl = document.getElementById(cid+'_vlabel');
    vl.setAttribute('x', sx.toFixed(1));
    vl.setAttribute('y', (CPT+11).toFixed(1));
    vl.textContent = Math.round(dx) + 'd';
  }});

  svgEl.addEventListener('dblclick', function() {{
    document.getElementById(cid+'_hline').setAttribute('y1','-1000');
    document.getElementById(cid+'_hline').setAttribute('y2','-1000');
    document.getElementById(cid+'_vline').setAttribute('x1','-1000');
    document.getElementById(cid+'_vline').setAttribute('x2','-1000');
    document.getElementById(cid+'_hlabel').textContent='';
    document.getElementById(cid+'_vlabel').textContent='';
  }});
}})();
</script>
"""

                dim_btn_html = (
                    f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">'
                    f'<div class="section-title" style="margin:0">妊娠率（DIM別）</div>'
                    f'<button id="{dim_chart_id}_btn" style="font-size:10px;padding:3px 10px;border:1px solid #e53935;'
                    f'border-radius:4px;background:#fff;color:#e53935;cursor:pointer;white-space:nowrap">＋ 基準線</button>'
                    f'</div>'
                )
                dim_hint = '<div style="font-size:8px;color:#bbb;margin-top:2px">基準線モード：クリックで固定 / ダブルクリックでリセット</div>'

                pr_dim_chart_html = (
                    dim_btn_html
                    + f'<div class="subheader">{html.escape(pd_start)} ～ {html.escape(pd_end)}　VWP=50日</div>\n'
                    + combo_svg
                    + dim_hint
                    + dim_crosshair_js
                )

        try:
            from modules.report_cow_bridge import DEFAULT_PORT as _report_open_cow_port
        except ImportError:
            _report_open_cow_port = 51985

        # ─── アラートHTML（BHB高値牛・乳量低下牛）────────────────────────────
        _milk_alerts = milk_alerts or {"bhb_rows": [], "milk_drop_rows": [], "latest_date": None}
        _bhb_rows  = _milk_alerts.get("bhb_rows", [])
        _drop_rows = _milk_alerts.get("milk_drop_rows", [])
        _alert_date = _milk_alerts.get("latest_date") or ""

        def _alert_tbl(rows, title, date_lbl):
            if not rows:
                return (f'<div class="section-title">{title}</div>'
                        f'<p class="subnote" style="margin:6px 0 0">該当なし</p>')
            trs = "".join(
                f'<tr>'
                f'<td class="num"><a href="#" class="report-cow-link" data-cow-id="{html.escape(str(r["cow_id"]))}">{html.escape(str(r["cow_id"]))}</a></td>'
                f'<td class="num">{html.escape(str(r["lact"]))}</td>'
                f'<td class="num">{html.escape(str(r["dim"]))}</td>'
                f'<td class="num">{html.escape(str(r["milk"]))}</td>'
                f'<td class="num">{html.escape(str(r["prev_milk"]))}</td>'
                f'<td class="num{r.get("scc_class","")}">{html.escape(str(r["scc"]))}</td>'
                f'<td class="num{r.get("bhb_class","")}">{html.escape(str(r["bhb"]))}</td>'
                f'</tr>'
                for r in rows
            )
            date_s = html.escape(str(date_lbl)) if date_lbl else ""
            return (
                f'<div class="section-title">{title}'
                f'{"（" + date_s + "）" if date_s else ""}'
                f'</div>'
                f'<div class="subheader">{len(rows)}頭</div>'
                f'<div style="overflow-x:auto">'
                f'<table class="summary-table" style="width:100%;font-size:12px">'
                f'<thead><tr><th>ID</th><th>産次</th><th>DIM</th>'
                f'<th>今月<br>乳量</th><th>前月<br>乳量</th><th>SCC</th><th>BHB</th></tr></thead>'
                f'<tbody>{trs}</tbody></table></div>'
            )

        bhb_alert_html  = _alert_tbl(_bhb_rows,  "BHB高値牛（≥0.13）",    _alert_date)
        drop_alert_html = _alert_tbl(_drop_rows, "乳量15%以上低下牛", _alert_date)

        # ─── 乳検トレンドCSS ────────────────────────────────────────────────
        _trend_css = ""
        try:
            from modules.milk_report_extras import MILK_REPORT_COMMENT_CSS as _trcss
            _trend_css = _trcss
        except Exception:
            pass

        # ─── KPI カード値 ────────────────────────────────────────────────────
        _kpi_total    = f"{total}頭"
        _kpi_avg_milk = f"{milk_stats['avg_milk']:.1f}kg"  if milk_stats.get("avg_milk")            else "-"
        _kpi_preg     = f"{herd_summary.get('pregnancy_rate',0):.1f}%" if herd_summary.get("pregnancy_rate") is not None else "-"
        _kpi_pr21     = f"{pr_hdr_data['avg_pr']:.0f}%"   if pr_hdr_data and pr_hdr_data.get("avg_pr")  is not None else "-"
        _kpi_avg_dim  = f"{herd_summary.get('avg_dim',0):.0f}日"       if herd_summary.get("avg_dim")  is not None else "-"
        _kpi_conception = "-"
        try:
            if fertility_stats and fertility_stats.get("stats"):
                _stats = fertility_stats.get("stats") or {}
                _tot_conceived = sum((s or {}).get("conceived", 0) for s in _stats.values())
                _tot_not_conceived = sum((s or {}).get("not_conceived", 0) for s in _stats.values())
                _den = _tot_conceived + _tot_not_conceived
                if _den > 0:
                    _kpi_conception = f"{(_tot_conceived / _den) * 100:.1f}%"
        except Exception:
            _kpi_conception = "-"
        _kpi_max_milk = f"{milk_stats['max_milk']:.1f}kg"  if milk_stats.get("max_milk")             else "-"
        _kpi_min_milk = f"{milk_stats['min_milk']:.1f}kg"  if milk_stats.get("min_milk")             else "-"
        _kpi_med_milk = f"{milk_stats['median_milk']:.1f}kg" if milk_stats.get("median_milk")        else "-"
        _milk_date    = milk_stats.get("latest_date") or ""

        html_content = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ダッシュボード</title>
    <style>
        /* ゲノム・乳検レポートと統一（フォント・配色・セクション見出し）。ダッシュボード用に可読性は維持 */
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Meiryo', 'Yu Gothic', sans-serif;
            font-size: 14px;
            margin: 20px;
            padding: 24px;
            background: #f8f9fa;
            color: #212529;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .dashboard-container {{
            max-width: 1400px;
            margin: 0 auto;
            background: #fff;
            padding: 32px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .header {{
            font-size: 1.4rem;
            font-weight: 700;
            margin-bottom: 4px;
            padding-bottom: 12px;
            border-bottom: 1px solid #e2e8f0;
            color: #1e293b;
        }}
        .header-meta {{
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 24px;
            letter-spacing: 0.02em;
        }}
        .dashboard-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
            margin-top: 20px;
        }}
        .dashboard-item {{
            display: flex;
            flex-direction: column;
            background: #fff;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
            transition: box-shadow 0.2s ease;
        }}
        .scatter-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 24px;
            margin-top: 0;
        }}
        .scatter-stack {{
            display: flex;
            flex-direction: row;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 16px;
            grid-column: 1 / -1;
            max-width: 100%;
            box-sizing: border-box;
        }}
        .scatter-stack .scatter-card {{
            flex: 1;
            min-width: 380px;
        }}
        .scatter-card {{
            border: 1px solid #dee2e6;
            border-radius: 6px;
            padding: 16px;
            background: #fff;
            overflow: hidden;
            box-sizing: border-box;
        }}
        .scatter-card .chart-title {{
            margin: 0 0 8px;
            font-size: 0.95rem;
            font-weight: 600;
            color: #212529;
        }}
        .scatter-card .scatter-chart {{
            width: 100%;
            max-width: 100%;
            overflow: hidden;
            box-sizing: border-box;
        }}
        .scatter-card .plotly-scatter-div {{
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box;
        }}
        .scatter-legend {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px 20px;
            margin-top: 10px;
            padding-bottom: 4px;
            font-size: 12px;
            color: #495057;
        }}
        .scatter-legend .legend-item {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}
        .scatter-legend .legend-dot {{
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            flex-shrink: 0;
        }}
        .fertility-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
            margin-top: 0;
        }}
        .heifer-fertility-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: 1fr;
            gap: 24px;
            margin-top: 0;
        }}
        .cumulative-preg-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 16px;
            margin-bottom: 0;
        }}
        .scatter-row-repro {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            margin-bottom: 0;
        }}
        .herd-dynamics-row {{
            grid-column: 1 / -1;
            margin-top: 0;
        }}
        .pr-detail-row {{
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: 1fr 2fr;
            gap: 24px;
            margin-top: 0;
        }}
        .pr-dim-card {{
            overflow: hidden;
        }}
        .dashboard-item:hover {{
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }}
        .parity-section {{
            margin-bottom: 0;
        }}
        .section-title {{
            font-size: 13px;
            font-weight: 700;
            padding-bottom: 8px;
            border-bottom: 1px solid #e2e8f0;
            margin-bottom: 16px;
            color: #334155;
            letter-spacing: 0.03em;
            text-transform: none;
        }}
        .chart-period {{
            font-size: 11px;
            color: #94a3b8;
            font-weight: 400;
            margin-left: 6px;
        }}
        .parity-chart-container {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
        }}
        .parity-donut {{
            flex-shrink: 0;
        }}
        .legend {{
            width: 100%;
        }}
        .legend-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 1px solid #e9ecef;
        }}
        .legend-row:last-child {{
            border-bottom: none;
        }}
        .legend-dot {{
            display: inline-block;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            margin-right: 10px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }}
        .legend-value {{
            display: flex;
            gap: 12px;
        }}
        .legend-count {{
            font-weight: bold;
        }}
        .legend-percent {{
            color: #666;
        }}
        .donut-center {{
            font-size: 12px;
            font-weight: bold;
            text-anchor: middle;
            fill: #333;
        }}
        .milk-section {{
            margin-bottom: 0;
        }}
        .milk-chart-container {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 15px;
            margin-bottom: 15px;
        }}
        .milk-donut {{
            flex-shrink: 0;
        }}
        .milk-stats {{
            margin: 15px 0 0 0;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
        }}
        .herd-summary-stats {{
            margin-top: 20px;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
        }}
        .summary-stat-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            font-size: 14px;
        }}
        .summary-stat-item:not(:last-child) {{
            border-bottom: 1px solid #e9ecef;
        }}
        .summary-stat-label {{
            color: #495057;
            font-weight: 500;
            font-size: 14px;
        }}
        .summary-stat-value {{
            color: #212529;
            font-weight: 600;
            font-size: 14px;
        }}
        .stats-label {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            font-size: 14px;
            margin-bottom: 0;
            font-weight: normal;
        }}
        .stats-label:not(:last-child) {{
            border-bottom: 1px solid #e9ecef;
        }}
        .stats-label-text {{
            color: #495057;
            font-weight: 500;
            font-size: 14px;
        }}
        .stats-label-value {{
            color: #212529;
            font-weight: 600;
            font-size: 14px;
        }}
        .stats-note {{
            font-size: 11px;
            color: #666;
            margin-top: 6px;
        }}
        .fertility-section {{
            margin-top: 0;
        }}
        .scatter-section {{
            margin-top: 0;
        }}
        .scatter-chart {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin-top: 10px;
        }}
        .scatter-legend {{
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #e9ecef;
        }}
        .scatter-legend .legend-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
            color: #495057;
        }}
        .scatter-legend .legend-dot {{
            display: inline-block;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            border: 1px solid rgba(0, 0, 0, 0.1);
        }}
        .subheader {{
            font-size: 12px;
            margin-bottom: 12px;
            color: #495057;
        }}
        .summary-table {{
            border-collapse: collapse;
            width: 100%;
            margin-top: 10px;
            font-size: 13px;
            border-radius: 8px;
            overflow: hidden;
        }}
        .summary-table th, .summary-table td {{
            border: 1px solid #dee2e6;
            padding: 10px 12px;
            text-align: left;
        }}
        .summary-table th {{
            background: #f1f5f9;
            font-weight: 700;
            text-align: center;
            font-size: 11px;
            color: #475569;
            letter-spacing: 0.03em;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .summary-table .num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .summary-table tbody tr:hover {{
            background: #eff6ff;
        }}
        .summary-table tbody tr:nth-child(even) {{
            background: #f9fafb;
        }}
        .summary-table tbody tr:nth-child(even):hover {{
            background: #eff6ff;
        }}
        .total-row {{
            background: #f1f5f9 !important;
            font-weight: 700;
            border-top: 1px solid #cbd5e1;
        }}
        .subnote {{
            color: #6c757d;
            font-size: 11px;
            font-style: italic;
        }}
        @media print {{
            body {{
                margin: 0;
                padding: 5px;
                background: #fff;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .dashboard-container {{
                max-width: 100%;
                padding: 5px;
                box-shadow: none;
                transform: scale(0.85);
                transform-origin: top left;
                width: 117.65%;  /* 1 / 0.85 = 1.1765 */
            }}
            .header {{
                font-size: 18px;
                margin-bottom: 15px;
            }}
            .dashboard-grid {{
                grid-template-columns: repeat(3, 1fr);
                gap: 12px;
                page-break-inside: avoid;
            }}
            .dashboard-item {{
                page-break-inside: avoid;
                padding: 16px;
                box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
            }}
            .section-title {{
                font-size: 15px;
                margin-bottom: 12px;
            }}
            .parity-chart-container {{
                gap: 10px;
                margin-bottom: 10px;
            }}
            .legend-row {{
                padding: 5px 0;
                font-size: 12px;
            }}
            .milk-stats, .herd-summary-stats {{
                padding: 12px;
                margin-top: 10px;
            }}
            .stats-label, .summary-stat-item {{
                font-size: 12px;
            }}
            .summary-stat-value {{
                font-size: 13px;
            }}
            .summary-table {{
                font-size: 11px;
            }}
            .summary-table th, .summary-table td {{
                padding: 6px 8px;
            }}
            .summary-table th {{
                font-size: 10px;
            }}
            .subheader {{
                font-size: 11px;
                margin-bottom: 8px;
            }}
            .herd-dynamics-chart {{
                margin: 12px 0;
            }}
            .herd-dynamics-chart svg {{
                font-family: 'Meiryo', 'Yu Gothic', sans-serif;
            }}
            .herd-dynamics-legend {{
                display: flex;
                flex-wrap: wrap;
                align-items: center;
                gap: 4px 16px;
                font-size: 12px;
                color: #546e7a;
                margin-top: 12px;
            }}
            .herd-dynamics-legend .legend-swatch {{
                display: inline-block;
                width: 14px;
                height: 14px;
                border: 1px solid #666;
                vertical-align: middle;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            .herd-dynamics-legend .legend-label {{
                vertical-align: middle;
            }}
            @page {{
                size: portrait;
                margin: 0.5cm;
            }}
        }}
        /* ── ナビゲーションバー ─────────────────────────────────────────── */
        .dash-nav {{
            position: sticky;
            top: 0;
            z-index: 200;
            background: #fff;
            border-bottom: 1px solid #dee2e6;
            box-shadow: 0 2px 6px rgba(0,0,0,.08);
            padding: 0 24px;
        }}
        .dash-nav .nav-inner {{
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            align-items: center;
            gap: 8px;
            height: 46px;
        }}
        .dash-nav .nav-farm {{
            font-weight: 700;
            color: #1e293b;
            font-size: 14px;
            white-space: nowrap;
            margin-right: 12px;
        }}
        .nav-buttons {{
            display: flex;
            gap: 4px;
            flex-wrap: wrap;
        }}
        .nav-btn {{
            padding: 5px 16px;
            border: 1px solid #dee2e6;
            border-radius: 20px;
            background: #fff;
            color: #495057;
            font-size: 13px;
            cursor: pointer;
            transition: all .15s;
            white-space: nowrap;
        }}
        .nav-btn:hover {{ background: #e9ecef; }}
        .nav-btn.active {{
            background: #1d4ed8;
            color: #fff;
            border-color: #1d4ed8;
            font-weight: 600;
        }}
        /* ── セクション ────────────────────────────────────────────────── */
        .dash-section {{
            margin-top: 36px;
            padding-top: 4px;
        }}
        .sec-label {{
            font-size: 1.05rem;
            font-weight: 700;
            color: #1e293b;
            border-left: 4px solid #1d4ed8;
            padding-left: 12px;
            margin-bottom: 20px;
        }}
        /* ── KPI カード ──────────────────────────────────────────────── */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 14px;
            margin-bottom: 24px;
        }}
        @media (max-width: 900px) {{
            .kpi-grid {{ grid-template-columns: repeat(3, 1fr); }}
        }}
        .kpi-card {{
            background: #fff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 16px 16px 14px;
            display: flex;
            flex-direction: column;
            gap: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,.04);
            border-left: 4px solid #1d4ed8;
        }}
        .kpi-card.kpi-green  {{ border-left-color: #059669; }}
        .kpi-card.kpi-orange {{ border-left-color: #d97706; }}
        .kpi-card.kpi-purple {{ border-left-color: #7c3aed; }}
        .kpi-card.kpi-teal   {{ border-left-color: #0891b2; }}
        .kpi-card.kpi-red    {{ border-left-color: #dc2626; }}
        .kpi-label {{
            font-size: 11px;
            color: #64748b;
            font-weight: 600;
            letter-spacing: 0.04em;
        }}
        .kpi-value {{
            font-size: 1.8rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.0;
            letter-spacing: -0.01em;
        }}
        .kpi-sub {{
            font-size: 11px;
            color: #94a3b8;
        }}
        /* ── 2/3カラムグリッド ───────────────────────────────────────���─ */
        .two-col-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }}
        .three-col-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 20px;
        }}
        /* ── アラートテーブル ─────────────────────────────────────────── */
        .bhb-alert {{ background: #fff3cd !important; font-weight: 600; }}
        .scc-alert {{ color: #dc3545; font-weight: 600; }}
        /* ── 乳量統計バー ────────────────────────────────────────────── */
        .milk-stat-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 7px 0;
            font-size: 13px;
            border-bottom: 1px solid #f0f0f0;
        }}
        .milk-stat-row:last-child {{ border-bottom: none; }}
        .milk-stat-lbl {{ color: #495057; }}
        .milk-stat-val {{ font-weight: 700; color: #212529; }}
        {_trend_css}
    </style>
    <script>var FALCON_OPEN_COW_PORT = {_report_open_cow_port};</script>
</head>
<body>
<!-- ────────────────── 固定ナビゲーションバー ────────────────── -->
<nav class="dash-nav" id="dash-nav">
  <div class="nav-inner">
    <span class="nav-farm">{html.escape(str(farm_name))}</span>
    <div class="nav-buttons">
      <button class="nav-btn active" data-sec="sec-overview"  onclick="jumpTo('sec-overview')">概要</button>
      <button class="nav-btn"        data-sec="sec-milk"      onclick="jumpTo('sec-milk')">乳生産</button>
      <button class="nav-btn"        data-sec="sec-repro"     onclick="jumpTo('sec-repro')">受胎率</button>
      <button class="nav-btn"        data-sec="sec-pr"        onclick="jumpTo('sec-pr')">妊娠率</button>
      <button class="nav-btn"        data-sec="sec-dynamics"  onclick="jumpTo('sec-dynamics')">牛群動態</button>
    </div>
  </div>
</nav>

<div class="dashboard-container">
  <div class="header">{html.escape(str(farm_name))}</div>
  <div class="header-meta">ダッシュボード　／　更新: {date.today().strftime('%Y年%m月%d日')}</div>

  <!-- ════════════════════════════════════════════════════
       セクション① 概要
       ════════════════════════════════════════════════════ -->
  <section id="sec-overview" class="dash-section">
    <div class="sec-label">概要</div>

    <!-- KPI カード (6枚) ─ 最重要指標を左上に配置（ガイドブック p.23） -->
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">経産牛頭数</div>
        <div class="kpi-value">{_kpi_total}</div>
        <div class="kpi-sub">平均産次 {avg_parity:.1f}</div>
      </div>
      <div class="kpi-card kpi-purple">
        <div class="kpi-label">21日妊娠率（PR）</div>
        <div class="kpi-value">{_kpi_pr21}</div>
        <div class="kpi-sub">直近18サイクル平均</div>
      </div>
      <div class="kpi-card kpi-teal">
        <div class="kpi-label">平均受胎率</div>
        <div class="kpi-value">{_kpi_conception}</div>
        <div class="kpi-sub">直近1年（経産牛）</div>
      </div>
      <div class="kpi-card kpi-orange">
        <div class="kpi-label">妊娠牛割合</div>
        <div class="kpi-value">{_kpi_preg}</div>
        <div class="kpi-sub">現在在籍経産牛</div>
      </div>
      <div class="kpi-card kpi-green">
        <div class="kpi-label">平均乳量</div>
        <div class="kpi-value">{_kpi_avg_milk}</div>
        <div class="kpi-sub">直近乳検 {_milk_date}</div>
      </div>
      <div class="kpi-card kpi-red">
        <div class="kpi-label">平均DIM</div>
        <div class="kpi-value">{_kpi_avg_dim}</div>
        <div class="kpi-sub">経産牛・分娩後日数平均</div>
      </div>
    </div>

    <!-- 産次構成 + 牛群サマリー -->
    <div class="two-col-grid">
      <div class="dashboard-item">
        <div class="parity-section">
          <div class="section-title">産次構成</div>
          <div class="parity-chart-container">
            <div class="parity-donut">{parity_donut_svg}</div>
            <div class="legend">{parity_legend}</div>
          </div>
          <div class="herd-summary-stats">
            {f'<div class="summary-stat-item"><span class="summary-stat-label">平均分娩後日数:</span><span class="summary-stat-value">{herd_summary.get("avg_dim",0):.1f}日</span></div>' if herd_summary.get("avg_dim") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">平均分娩後日数:</span><span class="summary-stat-value">-</span></div>'}
            {f'<div class="summary-stat-item"><span class="summary-stat-label">分娩間隔（実績）:</span><span class="summary-stat-value">{herd_summary.get("avg_cci",0):.1f}日</span></div>' if herd_summary.get("avg_cci") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">分娩間隔（実績）:</span><span class="summary-stat-value">-</span></div>'}
            {f'<div class="summary-stat-item"><span class="summary-stat-label">予定分娩間隔:</span><span class="summary-stat-value">{herd_summary.get("avg_pcci",0):.1f}日</span></div>' if herd_summary.get("avg_pcci") is not None else '<div class="summary-stat-item"><span class="summary-stat-label">予定分娩間隔:</span><span class="summary-stat-value">-</span></div>'}
          </div>
        </div>
      </div>
      <div class="dashboard-item">
        {scc_summary_html}
      </div>
    </div>
  </section>

  <!-- ════════════════════════════════════════════════════
       セクション② 乳生産
       ════════════════════════════════════════════════════ -->
  <section id="sec-milk" class="dash-section">
    <div class="sec-label">乳生産</div>

    <!-- 乳量要約 + 乳量統計 -->
    <div class="two-col-grid">
      <div class="dashboard-item">
        {milk_stats_html}
      </div>
      <div class="dashboard-item">
        <div class="section-title">乳量統計（直近乳検日{" " + html.escape(_milk_date) if _milk_date else ""}）</div>
        <div class="milk-stat-row"><span class="milk-stat-lbl">平均乳量</span><span class="milk-stat-val">{_kpi_avg_milk}</span></div>
        <div class="milk-stat-row"><span class="milk-stat-lbl">最高乳量</span><span class="milk-stat-val">{_kpi_max_milk}</span></div>
        <div class="milk-stat-row"><span class="milk-stat-lbl">最低乳量</span><span class="milk-stat-val">{_kpi_min_milk}</span></div>
        <div class="milk-stat-row"><span class="milk-stat-lbl">中央値</span><span class="milk-stat-val">{_kpi_med_milk}</span></div>
        {f'<div class="milk-stat-row"><span class="milk-stat-lbl">初産平均</span><span class="milk-stat-val">{milk_stats["avg_first_parity"]:.1f}kg</span></div>' if milk_stats.get("avg_first_parity") else ''}
        {f'<div class="milk-stat-row"><span class="milk-stat-lbl">経産（2産〜）平均</span><span class="milk-stat-val">{milk_stats["avg_multiparous"]:.1f}kg</span></div>' if milk_stats.get("avg_multiparous") else ''}
      </div>
    </div>

    <!-- 乳量・リニアスコア散布図 -->
    {scatter_plotly_html}

    <!-- 月次トレンドグラフ（12ヶ月） -->
    {milk_trend_html}

    <!-- BHB高値牛・乳量低下牛アラート -->
    <div class="two-col-grid" style="margin-top:20px">
      <div class="dashboard-item">{bhb_alert_html}</div>
      <div class="dashboard-item">{drop_alert_html}</div>
    </div>
  </section>

  <!-- ════════════════════════════════════════════════════
       セクション③ 受胎率
       ════════════════════════════════════════════════════ -->
  <section id="sec-repro" class="dash-section">
    <div class="sec-label">受胎率</div>

    <!-- 月別 / 産次別 / 授精回数別 -->
    <div class="three-col-grid">
      <div class="dashboard-item">{monthly_fertility_table_html}</div>
      <div class="dashboard-item">{fertility_table_html}</div>
      <div class="dashboard-item">{insemination_count_fertility_table_html}</div>
    </div>

    <!-- AI/ET比較 + 未経産牛受胎率 -->
    <div class="two-col-grid">
      <div class="dashboard-item">{ai_et_table_html}</div>
      <div class="dashboard-item">{heifer_monthly_table_html}</div>
    </div>

    <!-- 繁殖散布図 -->
    <div class="two-col-grid">
      <div class="dashboard-item">{first_ai_scatter_html}</div>
      <div class="dashboard-item">{rc_dim_scatter_html}</div>
    </div>
  </section>

  <!-- ════════════════════════════════════════════════════
       セクション④ 妊娠率（21日PR）
       ════════════════════════════════════════════════════ -->
  <section id="sec-pr" class="dash-section">
    <div class="sec-label">妊娠率（21日PR）</div>

    {f'''<div class="pr-detail-row">
      <div class="dashboard-item">{pr_cycle_table_html}</div>
      <div class="dashboard-item pr-dim-card">
        {pr_dim_chart_html}
        {pregnancy_donut_html}
      </div>
    </div>
    <div class="cumulative-preg-row" style="margin-top:20px">
      <div class="dashboard-item">{cumulative_preg_section_html}</div>
    </div>''' if repro_detail else '<p class="subnote" style="margin:8px 0">繁殖データがありません。</p>'}
  </section>

  <!-- ════════════════════════════════════════════════════
       セクション⑤ 牛群動態
       ════════════════════════════════════════════════════ -->
  <section id="sec-dynamics" class="dash-section">
    <div class="sec-label">牛群動態</div>
    <div class="dashboard-item">
      <div class="herd-dynamics-section">
        {herd_dynamics_chart_html}
      </div>
    </div>
  </section>

</div><!-- /.dashboard-container -->

<script>
(function() {{
  function jumpTo(id) {{
    var el = document.getElementById(id);
    if (el) el.scrollIntoView({{behavior: 'smooth', block: 'start'}});
  }}
  window.jumpTo = jumpTo;

  // IntersectionObserver でアクティブボタンを追跡
  var secs = document.querySelectorAll('.dash-section');
  var btns = document.querySelectorAll('.nav-btn');
  if (typeof IntersectionObserver !== 'undefined') {{
    var observer = new IntersectionObserver(function(entries) {{
      entries.forEach(function(entry) {{
        if (entry.isIntersecting) {{
          btns.forEach(function(b) {{
            b.classList.toggle('active', b.dataset.sec === entry.target.id);
          }});
        }}
      }});
    }}, {{rootMargin: '-10% 0px -60% 0px', threshold: 0}});
    secs.forEach(function(s) {{ observer.observe(s); }});
  }}
}})();
</script>
</body>
</html>"""
        return html_content
    
