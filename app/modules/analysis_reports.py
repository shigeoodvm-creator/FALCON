"""
FALCON2 - AnalysisReports
Phase E-1: 定型レポート定義

【設計原則】
- 定型レポート = 「名前付き分析ショートカット」
- 新しい集計ロジックは一切作らない
- SQLは analysis_sql_templates.py のもののみ使用
- Python側での集計・再計算は禁止
"""

from typing import Dict, Any, Optional
from datetime import datetime, timedelta
from calendar import monthrange
import re
import logging

logger = logging.getLogger(__name__)


class AnalysisReports:
    """
    定型レポート定義クラス
    
    Phase E-1: 定型レポート管理
    - 辞書形式で定型レポートを定義
    - 各レポートはSQLテンプレートとパラメータを持つ
    """
    
    # 定型レポート定義
    REPORTS: Dict[str, Dict[str, Any]] = {
        "calving_month_lact_report": {
            "display_name": "月別分娩頭数（産次別）",
            "description": "月別×産次別の分娩頭数を集計",
            "template_name": "calving_by_month_lact",
            "default_params": {
                "period": "last_12_months"
            },
            "default_output": ["screen"]
        },
        
        "insemination_month_lact_report": {
            "display_name": "月別授精頭数（産次別）",
            "description": "月別×産次別の授精頭数（AI+ET）を集計",
            "template_name": "insemination_by_month_lact",
            "default_params": {
                "period": "last_12_months"
            },
            "default_output": ["screen"]
        },
        
        "conception_rate_month_lact_report": {
            "display_name": "月別受胎率（産次別）",
            "description": "月別×産次別の受胎率を集計",
            "template_name": "conception_rate_by_month_lact",
            "default_params": {
                "period": "last_12_months"
            },
            "default_output": ["screen"]
        },
        
        "lact2_calving_report": {
            "display_name": "2産のみ分娩頭数",
            "description": "2産のみの分娩頭数を月別に集計",
            "template_name": "calving_lact2_only",
            "default_params": {
                "period": "last_12_months"
            },
            "default_output": ["screen"]
        },
        
        "dim_distribution_report": {
            "display_name": "DIM分布",
            "description": "DIM（分娩後日数）の分布を集計",
            "template_name": "dim_distribution",
            "default_params": {
                "period": "last_12_months"
            },
            "default_output": ["screen"]
        },
    }
    
    def __init__(self):
        """初期化"""
        pass
    
    def get_report(self, report_key: str) -> Optional[Dict[str, Any]]:
        """
        定型レポートを取得
        
        Args:
            report_key: レポートキー
        
        Returns:
            レポート定義（見つからない場合はNone）
        """
        return self.REPORTS.get(report_key)
    
    def list_reports(self) -> Dict[str, Dict[str, Any]]:
        """
        利用可能な定型レポート一覧を取得
        
        Returns:
            レポート定義の辞書
        """
        return self.REPORTS.copy()
    
    def find_report_by_display_name(self, display_name: str) -> Optional[str]:
        """
        表示名からレポートキーを検索
        
        Args:
            display_name: 表示名（部分一致可）
        
        Returns:
            レポートキー（見つからない場合はNone）
        """
        for key, report in self.REPORTS.items():
            if display_name in report.get("display_name", ""):
                return key
        return None
    
    def calculate_period(
        self,
        period_type: str,
        custom_start: Optional[str] = None,
        custom_end: Optional[str] = None
    ) -> Dict[str, str]:
        """
        期間を計算
        
        Args:
            period_type: 期間タイプ（"last_12_months", "this_year", "custom"）
            custom_start: カスタム開始日（YYYY-MM-DD形式、period_type="custom"の場合）
            custom_end: カスタム終了日（YYYY-MM-DD形式、period_type="custom"の場合）
        
        Returns:
            {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"}
        """
        today = datetime.now()
        
        if period_type == "last_12_months":
            # 直近12か月
            end_date = today
            start_date = today - timedelta(days=365)
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        elif period_type == "this_year":
            # 今年
            start_date = datetime(today.year, 1, 1)
            end_date = today
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        elif period_type == "custom":
            # カスタム期間
            if custom_start and custom_end:
                return {
                    "start": custom_start,
                    "end": custom_end
                }
            else:
                # カスタム期間が指定されていない場合は直近12か月
                return self.calculate_period("last_12_months")
        
        else:
            # デフォルト：直近12か月
            return self.calculate_period("last_12_months")
    
    def parse_period_from_text(self, text: str) -> Optional[Dict[str, str]]:
        """
        日本語テキストから期間を解析
        
        Args:
            text: 期間を表すテキスト（例：「2024年」「直近6か月」「3〜8月」）
        
        Returns:
            {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD"} または None
        """
        today = datetime.now()
        
        # 「2024年」形式
        year_pattern = r'(\d{4})年'
        match = re.search(year_pattern, text)
        if match:
            year = int(match.group(1))
            start_date = datetime(year, 1, 1)
            end_date = datetime(year, 12, 31)
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        # 「直近Nか月」形式
        months_pattern = r'直近(\d+)か月'
        match = re.search(months_pattern, text)
        if match:
            months = int(match.group(1))
            end_date = today
            start_date = today - timedelta(days=months * 30)  # 簡易計算
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        # 「N〜M月」形式（今年を前提）
        month_range_pattern = r'(\d+)〜(\d+)月'
        match = re.search(month_range_pattern, text)
        if match:
            start_month = int(match.group(1))
            end_month = int(match.group(2))
            year = today.year
            start_date = datetime(year, start_month, 1)
            # 終了月の最終日
            last_day = monthrange(year, end_month)[1]
            end_date = datetime(year, end_month, last_day)
            return {
                "start": start_date.strftime("%Y-%m-%d"),
                "end": end_date.strftime("%Y-%m-%d")
            }
        
        return None

