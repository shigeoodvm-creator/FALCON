"""
FALCON2 - Period Parser（期間解析モジュール）
自然な日本語で期間指定を解析し、start_date / end_date を確定する
"""

import re
import logging
from datetime import date, timedelta
from typing import Dict, Optional
from calendar import monthrange

logger = logging.getLogger(__name__)


def parse_period(text: str, base_date: date) -> Optional[Dict[str, date]]:
    """
    期間を解析
    
    Args:
        text: 期間指定文字列
        base_date: 基準日（デフォルトは今日）
    
    Returns:
        {
            "start": date,
            "end": date,
            "source": "explicit | year | month | relative | default"
        } または None（期間指定なし）
    """
    if not text:
        return None
    
    text = text.strip()
    
    # ① 明示指定（最優先）
    # YYYY[-/\.]MM[-/\.]DD\s*[~～\-]\s*(YYYY[-/\.])?MM[-/\.]DD
    explicit_pattern = re.search(
        r'(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})\s*[~～\-]\s*(\d{4})?[-/\.]?(\d{1,2})[-/\.](\d{1,2})',
        text
    )
    if explicit_pattern:
        start_year = int(explicit_pattern.group(1))
        start_month = int(explicit_pattern.group(2))
        start_day = int(explicit_pattern.group(3))
        end_year_str = explicit_pattern.group(4)
        end_month = int(explicit_pattern.group(5))
        end_day = int(explicit_pattern.group(6))
        
        # 終了日の年が省略された場合は開始日の年を使う
        if end_year_str:
            end_year = int(end_year_str)
        else:
            end_year = start_year
        
        try:
            start_date = date(start_year, start_month, start_day)
            end_date = date(end_year, end_month, end_day)
            
            if start_date > end_date:
                # 開始日が終了日より後の場合は無効
                logger.warning(f"[PeriodParser] 無効な期間指定: 開始日 > 終了日, text='{text}'")
                return None
            
            result = {
                "start": start_date,
                "end": end_date,
                "source": "explicit"
            }
            logger.info(f"[PeriodParser] source=explicit, start={start_date}, end={end_date}")
            return result
        except ValueError as e:
            logger.warning(f"[PeriodParser] 日付解析エラー: {e}, text='{text}'")
            return None
    
    # ② 年指定
    # 「今年」「本年」 → base_dateの年 2025-01-01 ～ 2025-12-31
    if "今年" in text or "本年" in text:
        year = base_date.year
        start_date = date(year, 1, 1)
        end_date = date(year, 12, 31)
        
        result = {
            "start": start_date,
            "end": end_date,
            "source": "year"
        }
        logger.info(f"[PeriodParser] source=year (今年/本年), start={start_date}, end={end_date}")
        return result
    
    # 「2025年」 → 2025-01-01 ～ 2025-12-31
    year_pattern = re.search(r'(\d{4})年', text)
    if year_pattern:
        year = int(year_pattern.group(1))
        start_date = date(year, 1, 1)
        end_date = date(year, 12, 31)
        
        result = {
            "start": start_date,
            "end": end_date,
            "source": "year"
        }
        logger.info(f"[PeriodParser] source=year, start={start_date}, end={end_date}")
        return result
    
    # ③ 月指定
    # 「2025年10月」 → 2025-10-01 ～ 2025-10-31
    # 「10月」 → base_date から最も近い10月
    year_month_pattern = re.search(r'(\d{4})年(\d{1,2})月', text)
    if year_month_pattern:
        year = int(year_month_pattern.group(1))
        month = int(year_month_pattern.group(2))
        if 1 <= month <= 12:
            start_date = date(year, month, 1)
            # 月の最後の日を取得
            _, last_day = monthrange(year, month)
            end_date = date(year, month, last_day)
            
            result = {
                "start": start_date,
                "end": end_date,
                "source": "month"
            }
            logger.info(f"[PeriodParser] source=month, start={start_date}, end={end_date}")
            return result
    
    # 「MM月」パターン（年が指定されていない場合）
    # 年指定なしの月指定は、業務ルールとして「直近の過去の月」を選択する
    month_only_pattern = re.search(r'(\d{1,2})月', text)
    if month_only_pattern:
        month = int(month_only_pattern.group(1))
        if 1 <= month <= 12:
            current_year = base_date.year
            current_month = base_date.month
            
            # 業務ルール：直近の過去の月を選択
            # month <= current_month の場合 → 当年（今年の該当月は既に完了しているか、現在の月）
            # month > current_month の場合 → 前年（今年の該当月は未来なので、前年の該当月を選択）
            if month <= current_month:
                year = current_year
            else:
                year = current_year - 1
            
            start_date = date(year, month, 1)
            _, last_day = monthrange(year, month)
            end_date = date(year, month, last_day)
            
            result = {
                "start": start_date,
                "end": end_date,
                "source": "month"
            }
            logger.info(f"[PeriodParser] source=month (年未指定), 基準日={base_date}, 決定年={year}, start={start_date}, end={end_date}")
            return result
    
    # ④ 相対期間
    # 「一年」「1年」「直近一年」「過去一年」→ base_date - 365日 ～ base_date
    # 「3ヶ月」「6ヶ月」も同様（月単位）
    
    # 年単位の相対期間
    relative_year_pattern = re.search(r'(?:直近|過去)?(?:一年|1年)', text)
    if relative_year_pattern:
        start_date = base_date - timedelta(days=365)
        end_date = base_date
        
        result = {
            "start": start_date,
            "end": end_date,
            "source": "relative"
        }
        logger.info(f"[PeriodParser] source=relative, start={start_date}, end={end_date}")
        return result
    
    # 月単位の相対期間（「3ヶ月」「6ヶ月」など）
    relative_month_pattern = re.search(r'(?:直近|過去)?(\d+)[ヶケ]月', text)
    if relative_month_pattern:
        months = int(relative_month_pattern.group(1))
        # 月数を日数に変換（簡易的に30日/月として計算）
        days = months * 30
        start_date = base_date - timedelta(days=days)
        end_date = base_date
        
        result = {
            "start": start_date,
            "end": end_date,
            "source": "relative"
        }
        logger.info(f"[PeriodParser] source=relative, start={start_date}, end={end_date}")
        return result
    
    # ⑤ 期間指定なし
    return None


