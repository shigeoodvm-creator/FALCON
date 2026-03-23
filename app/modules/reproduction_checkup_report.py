"""
FALCON2 - 繁殖検診レポート（HTML）
レポート出力で「レポート」にチェックを入れたときに出力するHTMLを生成する。
のちのち項目を増やしていく想定でセクション単位で組み立てる。
"""

import calendar
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from db.db_handler import DBHandler
from modules.aggregation_service import AggregationService
from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager


def annual_required_pregnancies(
    parous_count: int,
    calving_interval_days: float,
    replacement_rate_percent: float,
    abortion_rate_percent: float,
) -> float:
    """
    年間必要妊娠頭数を計算する。

    年間必要妊娠頭数 ＝ 経産牛頭数 × (1 - 更新率) × 365 / 目標分娩間隔 × (1.0 + 流産率)
    """
    if calving_interval_days <= 0:
        return 0.0
    replacement = replacement_rate_percent / 100.0
    abortion = abortion_rate_percent / 100.0
    return (
        parous_count
        * (1.0 - replacement)
        * 365.0
        / calving_interval_days
        * (1.0 + abortion)
    )


def _is_cow_disposed(db: DBHandler, rule_engine: RuleEngine, cow_auto_id: Any) -> bool:
    if cow_auto_id is None:
        return True
    events = db.get_events_by_cow(cow_auto_id, include_deleted=False)
    for ev in events:
        if ev.get("event_number") in [rule_engine.EVENT_SOLD, rule_engine.EVENT_DEAD]:
            return True
    return False


def get_current_parous_count(db: DBHandler, rule_engine: RuleEngine) -> int:
    """現状の経産牛頭数（現存・LACT>=1）を取得"""
    all_cows = db.get_all_cows()
    existing = [c for c in all_cows if not _is_cow_disposed(db, rule_engine, c.get("auto_id"))]
    lactated = [
        c for c in existing
        if c.get("lact") is not None and (int(c.get("lact", 0)) if c.get("lact") is not None else 0) >= 1
    ]
    return len(lactated)


def get_first_parity_ratio_percent(db: DBHandler, rule_engine: RuleEngine) -> Optional[float]:
    """現状の初産割合（1産/経産牛）を%で返す。経産牛0頭の場合はNone"""
    all_cows = db.get_all_cows()
    existing = [c for c in all_cows if not _is_cow_disposed(db, rule_engine, c.get("auto_id"))]
    lactated = [
        c for c in existing
        if c.get("lact") is not None and (int(c.get("lact", 0)) if c.get("lact") is not None else 0) >= 1
    ]
    total = len(lactated)
    if total == 0:
        return None
    first_parity = sum(
        1 for c in lactated
        if (int(c.get("lact", 0)) if c.get("lact") is not None else 0) == 1
    )
    return round(100.0 * first_parity / total, 1)


def get_goal_values(settings_manager: SettingsManager) -> tuple:
    """農場目標から 目標分娩間隔・更新率・流産率 を取得 (interval_days, replacement_%, abortion_%)"""
    goals = settings_manager.get("farm_goals", {}) or {}
    interval = goals.get("calving_interval_days")
    if interval is None:
        interval = 420
    replacement = goals.get("first_lactation_ratio")
    if replacement is None:
        replacement = 30
    abortion = goals.get("abortion_rate_percent")
    if abortion is None:
        abortion = 10
    return interval, replacement, abortion


def get_farm_goal(settings_manager: SettingsManager, key: str, default: Any = None) -> Any:
    """農場目標の1項目を取得"""
    goals = settings_manager.get("farm_goals", {}) or {}
    return goals.get(key, default)


def get_parous_cows(db: DBHandler, rule_engine: RuleEngine) -> List[Dict[str, Any]]:
    """経産牛（現存・LACT>=1）のリストを返す"""
    all_cows = db.get_all_cows()
    existing = [c for c in all_cows if not _is_cow_disposed(db, rule_engine, c.get("auto_id"))]
    return [
        c for c in existing
        if c.get("lact") is not None and (int(c.get("lact", 0)) if c.get("lact") is not None else 0) >= 1
    ]


def _get_herd_repro_metrics(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
) -> Dict[str, Any]:
    """
    牛群の繁殖指標を計算する。
    妊娠率・発情発見率は未実装のため含めない。
    """
    from datetime import datetime as dt
    settings_manager = SettingsManager(farm_path)
    parous = get_parous_cows(db, rule_engine)
    n_parous = len(parous)
    if n_parous == 0:
        return {}

    try:
        check_dt = dt.strptime(checkup_date[:10], "%Y-%m-%d")
    except ValueError:
        check_dt = dt.now()

    result: Dict[str, Any] = {}

    # 搾乳日数（平均DIM）
    dims = []
    for c in parous:
        clvd = c.get("clvd")
        if not clvd:
            continue
        try:
            clvd_dt = dt.strptime(clvd[:10], "%Y-%m-%d")
            d = (check_dt - clvd_dt).days
            if d >= 0:
                dims.append(d)
        except (ValueError, TypeError):
            pass
    result["avg_dim"] = round(sum(dims) / len(dims)) if dims else None

    # 妊娠牛の割合（rc 5=妊娠中, 6=乾乳中）
    rc_vals = [c.get("rc") for c in parous if c.get("rc") is not None]
    pregnant_dry = sum(1 for r in rc_vals if r in (5, 6))
    result["pregnant_cow_ratio"] = round(pregnant_dry / n_parous * 100, 0) if n_parous else None
    result["pregnant_count"] = pregnant_dry
    result["non_pregnant_count"] = n_parous - pregnant_dry

    # 初回授精日数・初回授精受胎率・授精回数・分娩間隔・予定分娩間隔（FormulaEngine使用）
    formula_engine = None
    item_dict_path = Path(farm_path) / "config" / "item_dictionary.json"
    if not item_dict_path.exists():
        # プロジェクトルートの config_default を参照（app/modules から2つ上でFALCON）
        item_dict_path = Path(__file__).resolve().parent.parent.parent / "config_default" / "item_dictionary.json"
    if not item_dict_path.exists():
        item_dict_path = Path(__file__).resolve().parent.parent / "config_default" / "item_dictionary.json"
    try:
        from modules.formula_engine import FormulaEngine
        formula_engine = FormulaEngine(db, item_dict_path)
    except Exception:
        pass

    dimfai_list = []
    first_service_p = 0
    first_service_denom = 0
    ai_count_list = []
    cci_list = []
    pcci_list = []

    for c in parous:
        auto_id = c.get("auto_id")
        clvd = c.get("clvd")
        if not auto_id:
            continue
        events = db.get_events_by_cow(auto_id, include_deleted=False)
        events.sort(key=lambda e: (e.get("event_date", ""), e.get("id", 0)))

        if formula_engine:
            try:
                calc = formula_engine.calculate(auto_id)
                # DIMFAI（初回授精日数）の平均用
                v = calc.get("DIMFAI")
                if v is not None and isinstance(v, (int, float)):
                    dimfai_list.append(int(round(float(v))))
                # CINT（分娩間隔）の平均用
                v = calc.get("CINT")
                if v is not None and isinstance(v, (int, float)) and float(v) > 0:
                    cci_list.append(int(round(float(v))))
                # PCI（予定分娩間隔）の平均用（妊娠中・乾乳中の牛で計算される）
                v = calc.get("PCI")
                if v is not None and isinstance(v, (int, float)) and float(v) > 0:
                    pcci_list.append(int(round(float(v))))
            except Exception:
                pass

        # 初回授精の outcome（分娩後最初のAI/ET）
        ai_et_after_clvd = [
            e for e in events
            if e.get("event_number") in (200, 201) and e.get("event_date") and (not clvd or e.get("event_date", "") > clvd)
        ]
        if ai_et_after_clvd:
            first_ev = min(ai_et_after_clvd, key=lambda e: (e.get("event_date", ""), e.get("id", 0)))
            j = first_ev.get("json_data") or {}
            if isinstance(j, str):
                try:
                    import json
                    j = json.loads(j) if j else {}
                except Exception:
                    j = {}
            out = j.get("outcome")
            if out == "P":
                first_service_p += 1
                first_service_denom += 1
            elif out == "O":
                first_service_denom += 1

        if clvd:
            cnt = sum(1 for e in events if e.get("event_number") in (200, 201) and e.get("event_date", "") > clvd)
            if cnt >= 1:
                ai_count_list.append(cnt)

    result["avg_dimfai"] = round(sum(dimfai_list) / len(dimfai_list)) if dimfai_list else None
    result["first_service_conception_rate"] = round(first_service_p / first_service_denom * 100, 0) if first_service_denom else None
    result["avg_insemination_count"] = round(sum(ai_count_list) / len(ai_count_list), 1) if ai_count_list else None
    result["avg_cci"] = round(sum(cci_list) / len(cci_list)) if cci_list else None
    result["avg_pcci"] = round(sum(pcci_list) / len(pcci_list)) if pcci_list else None

    # 受胎率・直近月の受胎率（AggregationService）
    end_d = check_dt
    start_d = dt(check_dt.year - 1, check_dt.month, 1) if check_dt.month > 1 else dt(check_dt.year - 2, 12, 1)
    start_str = start_d.strftime("%Y-%m-%d")
    end_str = end_d.strftime("%Y-%m-%d")
    agg = AggregationService(db)
    try:
        rows = agg.conception_rate_by_month_and_lact(start_str, end_str)
        total_num = sum(r.get("total_numerator") or 0 for r in rows)
        total_den = sum(r.get("total_denominator") or 0 for r in rows)
        result["conception_rate"] = round(total_num / total_den * 100, 0) if total_den else None
        if rows:
            last = rows[-1]
            ln, ld = last.get("total_numerator") or 0, last.get("total_denominator") or 0
            result["conception_rate_recent_month"] = round(ln / ld * 100, 0) if ld else None
        else:
            result["conception_rate_recent_month"] = None
    except Exception:
        result["conception_rate"] = None
        result["conception_rate_recent_month"] = None

    return result


def get_latest_pregnancy_confirmed_ai_et_date(db: DBHandler) -> Optional[str]:
    """
    直近の妊娠確定AI/ETイベントの日付を返す。
    妊娠実績はコードP（outcome='P'）のAI(200)またはET(201)イベントの数で、
    その「年」はAI/ETの実施日の年（例：12月のAIの妊娠確定は1月になるため年が異なる）。
    """
    conn = db.connect()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT MAX(event_date) AS d
        FROM event
        WHERE event_number IN (200, 201)
          AND deleted = 0
          AND event_date IS NOT NULL
          AND (json_data LIKE '%"outcome":"P"%' OR json_data LIKE '%"outcome": "P"%')
    """)
    row = cursor.fetchone()
    conn.close()
    if row and row[0] is not None:
        return str(row[0])[:10]
    return None


def get_monthly_pregnancy_counts(
    db: DBHandler, start_date: str, end_date: str
) -> List[Tuple[str, int]]:
    """
    期間内の月別・妊娠確定頭数（AI/ET outcome=P）を返す。
    Returns: [(ym, count), ...] 例 [("2025-01", 4), ("2025-02", 3), ...]
    """
    agg = AggregationService(db)
    rows = agg.conception_rate_by_month_and_lact(start_date, end_date)
    return [(r["ym"], int(r.get("total_numerator") or 0)) for r in rows]


def _cumulative_from_monthly(monthly: List[Tuple[str, int]], year: int) -> List[Tuple[str, float]]:
    """
    月別頭数から累積系列を生成。各月の末日時点の累積値を返す。
    Returns: [(date_str, cumulative), ...]  date_str は "YYYY-MM-DD"（月末日）
    """
    ym_to_count = {ym: c for ym, c in monthly}
    result = []
    cum = 0
    for m in range(1, 13):
        ym = f"{year}-{m:02d}"
        cum += ym_to_count.get(ym, 0)
        _, last = calendar.monthrange(year, m)
        result.append((f"{year}-{m:02d}-{last:02d}", float(cum)))
    return result


def _configure_matplotlib_japanese() -> None:
    """matplotlib で日本語表示用フォントを設定"""
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


def _cumulative_pregnancy_chart_png(
    current_year: int,
    annual_required: float,
    latest_ai_et_date: Optional[str],
    monthly_current: List[Tuple[str, int]],
    monthly_prev: List[Tuple[str, int]],
) -> str:
    """累積妊娠頭数グラフをPNGで生成し base64 で返す。"""
    import base64
    import io
    try:
        import matplotlib
        matplotlib.use("Agg")
        from matplotlib.figure import Figure
        import matplotlib.dates as mdates
    except ImportError:
        return ""

    _configure_matplotlib_japanese()

    jan1 = datetime(current_year, 1, 1)
    dec31 = datetime(current_year, 12, 31)
    latest_dt = None
    if latest_ai_et_date:
        try:
            latest_dt = datetime.strptime(latest_ai_et_date[:10], "%Y-%m-%d")
        except ValueError:
            pass

    target_x = [jan1, dec31]
    target_y = [0.0, annual_required]

    cum_cur = _cumulative_from_monthly(monthly_current, current_year)
    prev_year = current_year - 1
    prev_dates = []
    prev_cum = []
    for m in range(1, 13):
        _, last = calendar.monthrange(current_year, m)  # X軸は当年のカレンダーに合わせる
        prev_dates.append(datetime(current_year, m, last))
        ym = f"{prev_year}-{m:02d}"
        c = sum(cnt for ymo, cnt in monthly_prev if ymo <= ym)
        prev_cum.append(float(c))

    fig = Figure(figsize=(7.2, 3.2), dpi=100)
    ax = fig.add_subplot(111)

    max_y = max(annual_required, 10.0)
    if cum_cur:
        max_y = max(max_y, max(c[1] for c in cum_cur))
    if prev_cum:
        max_y = max(max_y, max(prev_cum))
    max_y = (int(max_y) // 10 + 1) * 10

    # 目標線: 直近まで実線、以降はグラデーション（破線）
    if latest_dt and jan1 < latest_dt < dec31:
        frac = (latest_dt - jan1).days / (dec31 - jan1).days
        y_at_latest = annual_required * frac
        ax.plot([jan1, latest_dt], [0.0, y_at_latest], color="#1e88e5", linewidth=2, label="目標", zorder=3)
        ax.plot(
            [latest_dt, dec31],
            [y_at_latest, annual_required],
            color="#1e88e5", linestyle="--", alpha=0.5, linewidth=1.5, zorder=2,
        )
    else:
        ax.plot(target_x, target_y, color="#1e88e5", linewidth=2, label="目標", zorder=3)

    if cum_cur:
        cur_dates = [datetime.strptime(d, "%Y-%m-%d") for d, _ in cum_cur]
        cur_vals = [v for _, v in cum_cur]
        # 累積実績は直近の妊娠確定AIイベントの位置まで（12/31まで確定していなければその年は確定しない）
        solid_end = len(cur_dates)
        if latest_dt:
            for i, d in enumerate(cur_dates):
                if d > latest_dt:
                    solid_end = i
                    break
        confirmed_dates = cur_dates[:solid_end]
        confirmed_vals = cur_vals[:solid_end]

        # 実績: 直近妊娠確定AIイベント日まで実線＋プロット（その先は描かない）
        if confirmed_dates and confirmed_vals:
            ax.plot(confirmed_dates, confirmed_vals, color="#e65100", linewidth=2, zorder=4, label="実績")
            ax.scatter(confirmed_dates, confirmed_vals, facecolors="none", edgecolors="#c62828", s=60, linewidths=2, zorder=5)
            ax.scatter([confirmed_dates[-1]], [confirmed_vals[-1]], color="#c62828", s=80, zorder=6)
            ax.annotate(f"{confirmed_vals[-1]:.1f}", (confirmed_dates[-1], confirmed_vals[-1]), xytext=(5, 5), textcoords="offset points", fontsize=9, color="#c62828")

    if prev_dates and prev_cum:
        ax.plot(prev_dates, prev_cum, color="#2e7d32", linewidth=1.5, linestyle="-", alpha=0.9, label="前年実績", zorder=2)

    ax.set_xlim(jan1, dec31)
    ax.set_ylim(0, max_y)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.set_ylabel("累積頭数", fontsize=10)
    ax.set_xlabel("月", fontsize=10)
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _section_annual_required_pregnancies(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
) -> str:
    """年間必要妊娠頭数セクションのHTMLを返す"""
    settings_manager = SettingsManager(farm_path)
    parous = get_current_parous_count(db, rule_engine)
    interval, replacement, abortion = get_goal_values(settings_manager)
    result = annual_required_pregnancies(parous, float(interval), float(replacement), float(abortion))
    ref_first = get_first_parity_ratio_percent(db, rule_engine)
    ref_first_str = f"参考：現状の初産割合 {ref_first}％" if ref_first is not None else "参考：経産牛0頭のため算出不可"

    # 1月からの実績頭数と、いつまでの妊娠確定AIイベントか（同じ行の横に表示用）
    actual_from_jan_html = ""
    latest_ai_et_date = get_latest_pregnancy_confirmed_ai_et_date(db)
    if latest_ai_et_date:
        try:
            current_year = int(latest_ai_et_date[:10][:4])
            latest_dt = datetime.strptime(latest_ai_et_date[:10], "%Y-%m-%d")
        except ValueError:
            latest_dt = None
        if latest_dt is not None:
            start_current = f"{current_year}-01-01"
            end_current = f"{current_year}-12-31"
            monthly_current = get_monthly_pregnancy_counts(db, start_current, end_current)
            cum_cur = _cumulative_from_monthly(monthly_current, current_year)
            if cum_cur:
                cur_dates = [datetime.strptime(d, "%Y-%m-%d") for d, _ in cum_cur]
                cur_vals = [v for _, v in cum_cur]
                solid_end = len(cur_dates)
                for i, d in enumerate(cur_dates):
                    if d > latest_dt:
                        solid_end = i
                        break
                if solid_end > 0:
                    confirmed_cum = cur_vals[solid_end - 1]
                    date_disp = latest_ai_et_date[:10].replace("-", "/")
                    actual_from_jan_html = f"""
  <div class="result-box result-box-actual">
    <span class="result-label">1月からの実績頭数</span>
    <span class="result-value result-value-actual">{confirmed_cum:.0f} 頭</span>
    <span class="result-note">（{date_disp} までの妊娠確定AIイベント）</span>
  </div>"""

    return f"""
<section class="report-section">
  <h2 class="section-title">年間必要妊娠頭数</h2>
  <table class="param-table">
    <tr><th>想定経産牛頭数</th><td>{parous} 頭</td><td class="ref">（参考：現状 {parous} 頭）</td></tr>
    <tr><th>目標分娩間隔</th><td>{int(interval)} 日</td><td class="ref">（農場目標設定）</td></tr>
    <tr><th>更新率</th><td>{int(replacement)} ％</td><td class="ref">（{ref_first_str}）</td></tr>
    <tr><th>流産率</th><td>{int(abortion)} ％</td><td class="ref">（農場目標設定）</td></tr>
  </table>
  <div class="result-row">
  <div class="result-box">
    <span class="result-label">年間必要妊娠頭数</span>
    <span class="result-value">{result:.1f} 頭</span>
  </div>
  {actual_from_jan_html}
  </div>
</section>
"""


def _section_cumulative_pregnancy_chart(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
) -> str:
    """累積妊娠頭数グラフ＋目標値表のセクションHTMLを返す。失敗時は簡易メッセージのみ返す。"""
    import logging
    logger = logging.getLogger(__name__)
    try:
        settings_manager = SettingsManager(farm_path)
        parous = get_current_parous_count(db, rule_engine)
        interval, replacement, abortion = get_goal_values(settings_manager)
        annual_required = annual_required_pregnancies(parous, float(interval), float(replacement), float(abortion))

        latest_ai_et_date = get_latest_pregnancy_confirmed_ai_et_date(db)
        if latest_ai_et_date:
            try:
                current_year = int(latest_ai_et_date[:4])
            except ValueError:
                current_year = datetime.now().year
        else:
            current_year = datetime.now().year

        start_current = f"{current_year}-01-01"
        end_current = f"{current_year}-12-31"
        start_prev = f"{current_year - 1}-01-01"
        end_prev = f"{current_year - 1}-12-31"

        monthly_current = get_monthly_pregnancy_counts(db, start_current, end_current)
        monthly_prev = get_monthly_pregnancy_counts(db, start_prev, end_prev)

        png_b64 = _cumulative_pregnancy_chart_png(
            current_year=current_year,
            annual_required=annual_required,
            latest_ai_et_date=latest_ai_et_date,
            monthly_current=monthly_current,
            monthly_prev=monthly_prev,
        )

        month_per = round(annual_required / 12, 1)
        per_cycle = round(annual_required / 24, 1)

        # 当年の累積妊娠頭数（直近妊娠確定AI/ET日時点）
        cum_cur = _cumulative_from_monthly(monthly_current, current_year)
        current_pregnancies: Optional[float] = None
        if cum_cur and latest_ai_et_date:
            try:
                latest_dt = datetime.strptime(latest_ai_et_date[:10], "%Y-%m-%d")
                cur_dates = [datetime.strptime(d, "%Y-%m-%d") for d, _ in cum_cur]
                cur_vals = [v for _, v in cum_cur]
                solid_end = len(cur_dates)
                for i, d in enumerate(cur_dates):
                    if d > latest_dt:
                        solid_end = i
                        break
                if solid_end > 0:
                    current_pregnancies = cur_vals[solid_end - 1]
                else:
                    current_pregnancies = 0.0
            except (ValueError, TypeError):
                pass
        elif cum_cur:
            current_pregnancies = cum_cur[-1][1]

        chart_html = ""
        if png_b64:
            chart_html = f'<div class="chart-wrapper"><img src="data:image/png;base64,{png_b64}" alt="累積妊娠頭数" class="cumulative-chart-img" /></div>'

        # 目標達成表示（現在の妊娠頭数 / 年間必要妊娠頭数）
        target_achievement_html = ""
        if current_pregnancies is not None and annual_required and annual_required > 0:
            pct = round(current_pregnancies / annual_required * 100, 0)
            target_achievement_html = f"""
  <div class="chart-target-achievement">
    <span class="chart-target-achievement-label">目標達成</span>
    <span class="chart-target-achievement-value">{int(pct)}%</span>
    <span class="chart-target-achievement-pct">（{int(current_pregnancies)}/{int(annual_required)}）</span>
  </div>"""

        return f"""
<section class="report-section">
  <h2 class="section-title">累積妊娠頭数（1月～12月）</h2>
  {chart_html}
  <div class="chart-param-row">
  <table class="param-table chart-param-table">
    <tr><th>目標分娩間隔</th><td>{int(interval)} 日</td></tr>
    <tr><th>年間必要妊娠頭数</th><td>{annual_required:.0f} 頭</td></tr>
    <tr><th>月あたり必要妊娠頭数</th><td>{month_per} 頭</td></tr>
    <tr><th>一回あたり必要妊娠頭数</th><td>{per_cycle} 頭</td></tr>
  </table>
  {target_achievement_html}
  </div>
</section>
"""
    except Exception as e:
        logger.exception("累積妊娠頭数グラフの生成に失敗しました")
        return f"""
<section class="report-section">
  <h2 class="section-title">累積妊娠頭数（1月～12月）</h2>
  <p class="section-desc" style="color:#c62828;">グラフの生成に失敗しました。({e!s})</p>
</section>
"""


def _section_repro_indicators(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
) -> str:
    """繁殖指標セクション。妊娠率・発情発見率は空欄。"""
    import logging
    logger = logging.getLogger(__name__)
    try:
        settings_manager = SettingsManager(farm_path)
        goals = settings_manager.get("farm_goals", {}) or {}
        target_pregnant_ratio = goals.get("pregnant_cow_ratio", 50)
        target_first_conception = goals.get("first_service_conception_rate", 40)
        target_calving_interval = goals.get("calving_interval_days", 420)

        m = _get_herd_repro_metrics(db, rule_engine, farm_path, checkup_date)
        if not m:
            return ""

        def _row(label: str, value_str: str, target_str: str = "", below_target: bool = False, target_met: bool = False) -> str:
            target_part = f" （{target_str}）" if target_str else ""
            if below_target:
                return f"<tr><th>{label}</th><td><span class=\"metric-value-below\">{value_str}</span>{target_part}</td></tr>"
            if target_met:
                return f"<tr><th>{label}<span class=\"metric-check\">✓</span></th><td>{value_str}</td></tr>"
            return f"<tr><th>{label}</th><td>{value_str}{target_part}</td></tr>"

        rows = []
        v = m.get("avg_dim")
        rows.append(_row("搾乳日数", f"{v} 日" if v is not None else "—", ""))
        v = m.get("pregnant_cow_ratio")
        target_str = f"{int(target_pregnant_ratio)}%" if target_pregnant_ratio is not None else ""
        below = (v is not None and target_pregnant_ratio is not None and v < target_pregnant_ratio)
        rows.append(_row("妊娠牛の割合", f"{int(v)} %" if v is not None else "—", target_str, below_target=below))
        v = m.get("avg_dimfai")
        rows.append(_row("初回授精日数", f"{v} 日" if v is not None else "—", ""))
        v = m.get("first_service_conception_rate")
        target_str = f"{int(target_first_conception)}%" if target_first_conception is not None else ""
        below = (v is not None and target_first_conception is not None and v < target_first_conception)
        rows.append(_row("初回授精受胎率", f"{int(v)} %" if v is not None else "—", target_str, below_target=below))
        v = m.get("avg_insemination_count")
        rows.append(_row("授精回数", f"{v} 回" if v is not None else "—", ""))
        rows.append(_row("妊娠率", "", ""))
        rows.append(_row("発情発見率", "", ""))
        v = m.get("conception_rate")
        rows.append(_row("受胎率", f"{int(v)} %" if v is not None else "—", ""))
        v = m.get("conception_rate_recent_month")
        rows.append(_row("直近月の受胎率", f"{int(v)} %" if v is not None else "—", ""))
        v = m.get("avg_cci")
        target_met = (v is not None and target_calving_interval is not None and v <= target_calving_interval)
        rows.append(_row("分娩間隔", f"{v} 日" if v is not None else "—", "", target_met=target_met))
        v = m.get("avg_pcci")
        target_str = f"{int(target_calving_interval)} 日" if target_calving_interval else ""
        below = (v is not None and target_calving_interval is not None and v > target_calving_interval)
        rows.append(_row("予定分娩間隔", f"{v} 日" if v is not None else "—", target_str, below_target=below))

        rows_html = "\n    ".join(rows)
        pregnant_pct = m.get("pregnant_cow_ratio")
        if pregnant_pct is None:
            pregnant_pct = 0.0
        # 浮動小数点誤差で 100 を超えたりマイナスにならないようクランプ
        pregnant_pct = max(0.0, min(100.0, float(pregnant_pct)))
        non_pct = 100.0 - pregnant_pct
        pregnant_count = m.get("pregnant_count", 0)
        non_pregnant_count = m.get("non_pregnant_count", 0)
        # ドーナツ円グラフ（SVG）: 外周は非妊娠=オレンジで100%塗り、その上に妊娠=青を重ねる方式
        offset_top = 25  # 0%を上端に
        donut_svg = f"""<svg class="donut-chart" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
  <!-- まず非妊娠（オレンジ）でリング全体を描画 -->
  <circle cx="100" cy="100" r="65" fill="none" stroke="#e65100" stroke-width="28" stroke-dasharray="100 0" stroke-dashoffset="{offset_top}" pathLength="100" />
  <!-- その上に妊娠（青）の割合だけ重ね描きして、残りをオレンジとして見せる -->
  <circle cx="100" cy="100" r="65" fill="none" stroke="#1565c0" stroke-width="28" stroke-dasharray="{pregnant_pct:.1f} {100 - pregnant_pct:.1f}" stroke-dashoffset="{offset_top}" pathLength="100" />
  <circle cx="100" cy="100" r="35" fill="#fff" />
  <text x="100" y="92" text-anchor="middle" font-size="12" fill="#666">妊娠率 %</text>
  <text x="100" y="112" text-anchor="middle" font-size="18" font-weight="bold" fill="#1565c0">{int(round(pregnant_pct))}</text>
</svg>
  <div class="donut-legend">
    <span class="leg-pregnant">妊娠: {pregnant_count}頭（{int(round(pregnant_pct))}%）</span>
    <span class="leg-non">非妊娠: {non_pregnant_count}頭（{int(round(non_pct))}%）</span>
  </div>"""
        return f"""
<section class="report-section">
  <h2 class="section-title">繁殖指標</h2>
  <div class="repro-indicators-row">
  <table class="param-table metric-table">
    {rows_html}
  </table>
  <div class="repro-donut-wrap">
    {donut_svg}
  </div>
  </div>
</section>
"""
    except Exception as e:
        logger.exception("繁殖指標の算出に失敗しました")
        return ""


def build_report_html(
    db: DBHandler,
    rule_engine: RuleEngine,
    farm_path: Path,
    checkup_date: str,
    farm_name: str = "",
) -> str:
    """
    繁殖検診レポートのHTMLを組み立てる。
    のちのち項目を増やす場合は、sections リストにセクションHTMLを追加する。
    """
    settings_manager = SettingsManager(farm_path)
    if not farm_name:
        farm_name = settings_manager.get("farm_name", Path(farm_path).name)

    sections_html: list = []
    # セクション1: 年間必要妊娠頭数
    sections_html.append(
        _section_annual_required_pregnancies(db, rule_engine, farm_path, checkup_date)
    )
    # セクション2: 累積妊娠頭数グラフ
    sections_html.append(
        _section_cumulative_pregnancy_chart(db, rule_engine, farm_path, checkup_date)
    )
    # セクション3: 繁殖指標
    sections_html.append(
        _section_repro_indicators(db, rule_engine, farm_path, checkup_date)
    )
    # ここに将来のセクションを追加
    # sections_html.append(_section_xxx(...))

    body_sections = "\n".join(sections_html)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>繁殖検診レポート</title>
  <style>
    body {{ font-family: "Meiryo UI", "MS Gothic", sans-serif; font-size: 13px; margin: 24px; background: #f5f5f5; color: #263238; }}
    .report-container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 24px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
    h1 {{ font-size: 16px; margin-bottom: 4px; color: #1565c0; }}
    .report-date {{ font-size: 11px; color: #666; margin-bottom: 16px; }}
    .report-section {{ margin-bottom: 20px; }}
    .report-section:last-child {{ margin-bottom: 0; }}
    .section-title {{ font-size: 13px; font-weight: bold; margin-bottom: 6px; border-bottom: 1px solid #e0e0e0; padding-bottom: 4px; }}
    .param-table {{ border-collapse: collapse; width: 100%; max-width: 480px; margin-bottom: 12px; font-size: 12px; }}
    .param-table th, .param-table td {{ border: 1px solid #e0e0e0; padding: 6px 10px; text-align: left; }}
    .param-table th {{ background: #f5f5f5; width: 140px; }}
    .param-table td.ref {{ color: #666; font-size: 10px; }}
    .result-box {{ background: #e3f2fd; padding: 10px 14px; border-radius: 8px; margin-bottom: 12px; }}
    .result-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 12px; }}
    .result-row .result-box {{ margin-bottom: 0; flex: 1; min-width: 180px; }}
    .result-box-actual {{ background: #fff3e0; }}
    .result-box-actual .result-label {{ color: #bf360c; }}
    .result-value-actual {{ color: #e65100; }}
    .result-note {{ display: block; font-size: 10px; color: #666; margin-top: 2px; }}
    .result-label {{ font-size: 11px; color: #1565c0; display: block; margin-bottom: 2px; }}
    .result-value {{ font-size: 18px; font-weight: bold; color: #1565c0; }}
    .chart-wrapper {{ margin: 10px 0; }}
    .cumulative-chart-img {{ max-width: 100%; height: auto; }}
    .chart-param-table {{ margin-top: 8px; font-size: 12px; }}
    .chart-param-table th, .chart-param-table td {{ padding: 4px 8px; }}
    .chart-param-row {{ display: flex; flex-wrap: wrap; gap: 24px; align-items: flex-start; margin-top: 8px; }}
    .chart-param-row .chart-param-table {{ margin-top: 0; flex-shrink: 0; }}
    .chart-target-achievement {{ flex-shrink: 0; min-width: 160px; background: #e3f2fd; padding: 12px 16px; border-radius: 8px; }}
    .chart-target-achievement-label {{ display: block; font-size: 11px; color: #1565c0; margin-bottom: 4px; }}
    .chart-target-achievement-value {{ font-size: 18px; font-weight: bold; color: #1565c0; }}
    .chart-target-achievement-pct {{ font-size: 12px; color: #666; margin-left: 4px; }}
    .metric-table {{ font-size: 12px; }}
    .metric-value-below {{ background: #ffcdd2; padding: 2px 6px; }}
    .metric-check {{ color: #2e7d32; margin-left: 4px; font-weight: bold; }}
    .repro-indicators-row {{ display: flex; flex-wrap: wrap; gap: 24px; align-items: flex-start; }}
    .repro-indicators-row .param-table {{ margin-bottom: 0; flex-shrink: 0; }}
    .repro-donut-wrap {{ flex-shrink: 0; width: 200px; }}
    .repro-donut-wrap .donut-chart {{ display: block; width: 100%; height: auto; }}
    .repro-donut-wrap .donut-legend {{ margin-top: 4px; font-size: 11px; }}
    .repro-donut-wrap .donut-legend span {{ display: inline-block; margin-right: 12px; }}
    .repro-donut-wrap .donut-legend .leg-pregnant {{ color: #1565c0; }}
    .repro-donut-wrap .donut-legend .leg-non {{ color: #e65100; }}
    @media print {{
      body {{ font-size: 11px; margin: 0; background: #fff; }}
      .report-container {{ max-width: none; padding: 12px; box-shadow: none; }}
      h1 {{ font-size: 14px; margin-bottom: 2px; }}
      .report-date {{ font-size: 10px; margin-bottom: 8px; }}
      .report-section {{ margin-bottom: 12px; page-break-inside: avoid; }}
      .section-title {{ font-size: 12px; margin-bottom: 4px; padding-bottom: 2px; }}
      .param-table {{ font-size: 10px; margin-bottom: 8px; }}
      .param-table th, .param-table td {{ padding: 4px 6px; }}
      .result-box {{ padding: 8px 10px; margin-bottom: 8px; }}
      .result-row {{ margin-bottom: 8px; }}
      .result-value {{ font-size: 15px; }}
      .chart-wrapper {{ margin: 6px 0; }}
      .cumulative-chart-img {{ max-height: 220px; width: auto; object-fit: contain; }}
      .chart-param-table {{ font-size: 10px; margin-top: 4px; }}
      .chart-param-table th, .chart-param-table td {{ padding: 2px 6px; }}
      .chart-param-row {{ gap: 16px; margin-top: 4px; }}
      .chart-target-achievement {{ padding: 8px 12px; min-width: 120px; }}
      .chart-target-achievement-value {{ font-size: 15px; }}
      .repro-indicators-row {{ gap: 16px; }}
      .repro-donut-wrap {{ width: 160px; }}
      @page {{ size: A4; margin: 12mm; }}
    }}
  </style>
</head>
<body>
  <div class="report-container">
    <h1>繁殖検診レポート</h1>
    <p class="report-date">検診日：{checkup_date}　農場：{farm_name}</p>
    {body_sections}
  </div>
</body>
</html>
"""
