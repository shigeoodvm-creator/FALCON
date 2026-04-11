"""
FALCON2 - AnalysisSQLTemplates
Phase D - Step 3: 分析モード専用の公式SQLテンプレート集

【設計原則】
- 人間が定義した「正しいSQL」を再利用
- 集計定義のブレを完全に防ぐ
- event.event_lact / event.event_dim を唯一の事実として扱う
- cow.lact は使用禁止
"""

from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)


class AnalysisSQLTemplates:
    """
    分析モード専用のSQLテンプレート管理クラス
    
    Phase D - Step 3: 公式SQLテンプレート集
    - 辞書形式でSQLテンプレートを保持
    - key は分析目的の論理名
    - value は完成済みの SELECT 文（パラメータ付き）
    """
    
    # SQLテンプレート定義
    TEMPLATES: Dict[str, str] = {
        # 月別×産次別 分娩頭数（ピボット形式、期間内の全月を表示）
        "calving_by_month_lact": """
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(:start, 1, 7) AS ym,
                  date(:start, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(:end, '+1 day')
            )
            SELECT
              m.ym,
              COALESCE(SUM(CASE WHEN e.event_lact = 1 THEN 1 ELSE 0 END), 0) AS lact1,
              COALESCE(SUM(CASE WHEN e.event_lact = 2 THEN 1 ELSE 0 END), 0) AS lact2,
              COALESCE(SUM(CASE WHEN e.event_lact >= 3 THEN 1 ELSE 0 END), 0) AS lact3plus,
              COALESCE(COUNT(e.id), 0) AS total
            FROM months m
            LEFT JOIN event e ON substr(e.event_date, 1, 7) = m.ym
              AND e.event_number = 202
              AND e.event_date IS NOT NULL
              AND e.event_date >= :start
              AND e.event_date <= :end
              AND e.deleted = 0
            GROUP BY m.ym
            ORDER BY m.ym
        """,
        
        # 月別×産次別 授精頭数（AI+ET、ピボット形式、期間内の全月を表示）
        "insemination_by_month_lact": """
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(:start, 1, 7) AS ym,
                  date(:start, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(:end, '+1 day')
            )
            SELECT
              m.ym,
              COALESCE(SUM(CASE WHEN e.event_lact = 1 THEN 1 ELSE 0 END), 0) AS lact1,
              COALESCE(SUM(CASE WHEN e.event_lact = 2 THEN 1 ELSE 0 END), 0) AS lact2,
              COALESCE(SUM(CASE WHEN e.event_lact >= 3 THEN 1 ELSE 0 END), 0) AS lact3plus,
              COALESCE(COUNT(e.id), 0) AS total
            FROM months m
            LEFT JOIN event e ON substr(e.event_date, 1, 7) = m.ym
              AND e.event_number IN (200, 201)
              AND e.event_date IS NOT NULL
              AND e.event_date >= :start
              AND e.event_date <= :end
              AND e.deleted = 0
            GROUP BY m.ym
            ORDER BY m.ym
        """,
        
        # 月別×産次別 受胎率（ピボット形式、期間内の全月を表示）
        "conception_rate_by_month_lact": """
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(:start, 1, 7) AS ym,
                  date(:start, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(:end, '+1 day')
            )
            SELECT
              m.ym,
              COALESCE(SUM(CASE 
                WHEN e.event_lact = 1 
                  AND e.json_data IS NOT NULL 
                  AND (
                    e.json_data LIKE '%"outcome":"P"%' 
                    OR e.json_data LIKE '%"outcome": "P"%'
                  )
                THEN 1 
                ELSE 0 
              END), 0) AS lact1_numerator,
              COALESCE(SUM(CASE 
                WHEN e.event_lact = 1 
                  AND (e.json_data IS NULL 
                    OR (
                      e.json_data NOT LIKE '%"outcome":"R"%' 
                      AND e.json_data NOT LIKE '%"outcome": "R"%'
                    ))
                THEN 1 
                ELSE 0 
              END), 0) AS lact1_denominator,
              COALESCE(SUM(CASE 
                WHEN e.event_lact = 2 
                  AND e.json_data IS NOT NULL 
                  AND (
                    e.json_data LIKE '%"outcome":"P"%' 
                    OR e.json_data LIKE '%"outcome": "P"%'
                  )
                THEN 1 
                ELSE 0 
              END), 0) AS lact2_numerator,
              COALESCE(SUM(CASE 
                WHEN e.event_lact = 2 
                  AND (e.json_data IS NULL 
                    OR (
                      e.json_data NOT LIKE '%"outcome":"R"%' 
                      AND e.json_data NOT LIKE '%"outcome": "R"%'
                    ))
                THEN 1 
                ELSE 0 
              END), 0) AS lact2_denominator,
              COALESCE(SUM(CASE 
                WHEN e.event_lact >= 3 
                  AND e.json_data IS NOT NULL 
                  AND (
                    e.json_data LIKE '%"outcome":"P"%' 
                    OR e.json_data LIKE '%"outcome": "P"%'
                  )
                THEN 1 
                ELSE 0 
              END), 0) AS lact3plus_numerator,
              COALESCE(SUM(CASE 
                WHEN e.event_lact >= 3 
                  AND (e.json_data IS NULL 
                    OR (
                      e.json_data NOT LIKE '%"outcome":"R"%' 
                      AND e.json_data NOT LIKE '%"outcome": "R"%'
                    ))
                THEN 1 
                ELSE 0 
              END), 0) AS lact3plus_denominator,
              COALESCE(SUM(CASE 
                WHEN e.json_data IS NOT NULL 
                  AND (
                    e.json_data LIKE '%"outcome":"P"%' 
                    OR e.json_data LIKE '%"outcome": "P"%'
                  )
                THEN 1 
                ELSE 0 
              END), 0) AS total_numerator,
              COALESCE(SUM(CASE 
                WHEN e.json_data IS NULL 
                  OR (
                    e.json_data NOT LIKE '%"outcome":"R"%' 
                    AND e.json_data NOT LIKE '%"outcome": "R"%'
                  )
                THEN 1 
                ELSE 0 
              END), 0) AS total_denominator
            FROM months m
            LEFT JOIN event e ON substr(e.event_date, 1, 7) = m.ym
              AND e.event_number IN (200, 201)
              AND e.event_date IS NOT NULL
              AND e.event_date >= :start
              AND e.event_date <= :end
              AND e.deleted = 0
              AND EXISTS (
                SELECT 1 FROM cow c
                WHERE c.auto_id = e.cow_auto_id
                  AND (
                    c.entr IS NULL
                    OR TRIM(COALESCE(c.entr, '')) = ''
                    OR e.event_date >= c.entr
                  )
              )
            GROUP BY m.ym
            ORDER BY m.ym
        """,
        
        # 2産のみ 抽出（分娩）
        "calving_lact2_only": """
            SELECT
              substr(event_date, 1, 7) AS ym,
              COUNT(*) AS cnt
            FROM event
            WHERE event_number = 202
              AND event_lact = 2
              AND event_date IS NOT NULL
              AND event_date >= :start
              AND event_date <= :end
              AND deleted = 0
            GROUP BY ym
            ORDER BY ym
        """,
        
        # event_dim を使った DIM 分布
        "dim_distribution": """
            SELECT
              CASE
                WHEN event_dim IS NULL THEN 'NULL'
                WHEN event_dim < 0 THEN '異常値'
                WHEN event_dim <= 30 THEN '0-30日'
                WHEN event_dim <= 60 THEN '31-60日'
                WHEN event_dim <= 90 THEN '61-90日'
                WHEN event_dim <= 120 THEN '91-120日'
                WHEN event_dim <= 150 THEN '121-150日'
                WHEN event_dim <= 180 THEN '151-180日'
                WHEN event_dim <= 210 THEN '181-210日'
                WHEN event_dim <= 240 THEN '211-240日'
                WHEN event_dim <= 270 THEN '241-270日'
                WHEN event_dim <= 300 THEN '271-300日'
                ELSE '300日超'
              END AS dim_range,
              COUNT(*) AS cnt
            FROM event
            WHERE event_number = 200
              AND event_date IS NOT NULL
              AND event_date >= :start
              AND event_date <= :end
              AND deleted = 0
            GROUP BY dim_range
            ORDER BY 
              CASE dim_range
                WHEN 'NULL' THEN 0
                WHEN '異常値' THEN 1
                WHEN '0-30日' THEN 2
                WHEN '31-60日' THEN 3
                WHEN '61-90日' THEN 4
                WHEN '91-120日' THEN 5
                WHEN '121-150日' THEN 6
                WHEN '151-180日' THEN 7
                WHEN '181-210日' THEN 8
                WHEN '211-240日' THEN 9
                WHEN '241-270日' THEN 10
                WHEN '271-300日' THEN 11
                ELSE 12
              END
        """,
    }
    
    # テンプレートの説明（AI用）
    TEMPLATE_DESCRIPTIONS: Dict[str, str] = {
        "calving_by_month_lact": "月別×産次別 分娩頭数（ピボット形式：初産・2産・3産以上・合計）",
        "insemination_by_month_lact": "月別×産次別 授精頭数（ピボット形式：初産・2産・3産以上・合計、AI+ET）",
        "conception_rate_by_month_lact": "月別×産次別 受胎率（ピボット形式：初産・2産・3産以上・合計、outcomeで判定）",
        "calving_lact2_only": "2産のみの分娩頭数（月別、event_lact=2で抽出）",
        "dim_distribution": "DIM分布（event_dimを使用、AIイベントのみ）",
    }
    
    def __init__(self):
        """初期化"""
        pass
    
    def get_template(self, template_name: str) -> Optional[str]:
        """
        テンプレートを取得
        
        Args:
            template_name: テンプレート名
        
        Returns:
            SQLテンプレート（見つからない場合はNone）
        """
        return self.TEMPLATES.get(template_name)
    
    def list_templates(self) -> Dict[str, str]:
        """
        利用可能なテンプレート一覧を取得
        
        Returns:
            テンプレート名と説明の辞書
        """
        return self.TEMPLATE_DESCRIPTIONS.copy()
    
    def expand_template(self, template_name: str, params: Dict[str, Any]) -> Optional[str]:
        """
        テンプレートを展開（パラメータを置換）
        
        Args:
            template_name: テンプレート名
            params: パラメータ辞書（例: {"start": "2024-01-01", "end": "2024-12-31"}）
        
        Returns:
            展開されたSQL文（見つからない場合はNone）
        """
        template = self.get_template(template_name)
        if not template:
            return None
        
        # パラメータを置換（:param_name 形式）
        sql = template
        for key, value in params.items():
            placeholder = f":{key}"
            if placeholder in sql:
                # SQLインジェクション対策：日付形式のみ許可
                if isinstance(value, str) and self._is_safe_date_value(value):
                    sql = sql.replace(placeholder, f"'{value}'")
                elif isinstance(value, (int, float)):
                    sql = sql.replace(placeholder, str(value))
                else:
                    logger.warning(f"安全でないパラメータ値: {key}={value}")
                    return None
        
        return sql.strip()
    
    def _is_safe_date_value(self, value: str) -> bool:
        """
        日付値が安全かチェック（SQLインジェクション対策）
        
        Args:
            value: チェックする値
        
        Returns:
            安全な場合はTrue
        """
        # YYYY-MM-DD 形式のみ許可
        import re
        date_pattern = r'^\d{4}-\d{2}-\d{2}$'
        return bool(re.match(date_pattern, value))
    
    def get_template_list_for_ai(self) -> str:
        """
        AI用のテンプレート一覧を取得（プロンプト用）
        
        Returns:
            テンプレート一覧の文字列
        """
        lines = ["利用可能なSQLテンプレート:"]
        for name, description in self.TEMPLATE_DESCRIPTIONS.items():
            lines.append(f"  - {name}: {description}")
        return "\n".join(lines)

