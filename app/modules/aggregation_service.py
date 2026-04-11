"""
FALCON2 - AggregationService
Phase C: event テーブル前提の公式・最小・堅牢な集計エンジン

【設計原則】
- event テーブルのみを参照
- event.event_lact / event.event_dim を使用（再計算禁止）
- SQL のみで集計（Python側でのGROUP BY禁止）
- 推論・再計算を一切行わない
- 再現性100%の公式数値を返す
"""

import sqlite3
import json
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path

from db.db_handler import DBHandler

logger = logging.getLogger(__name__)


class AggregationService:
    """
    集計サービス
    
    Phase C: event テーブル前提の公式集計エンジン
    - RuleEngine は使用しない
    - event.event_lact / event.event_dim を直接参照
    - SQL のみで集計
    """
    
    def __init__(self, db_handler: DBHandler):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
        """
        self.db = db_handler
    
    def calving_by_month_and_lact(
        self, 
        start_date: str, 
        end_date: str
    ) -> List[Dict[str, Any]]:
        """
        月別 × 産次別 分娩頭数（ピボット形式）
        
        Args:
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
        
        Returns:
            月別×産次別の分娩頭数リスト（ピボット形式）
            [
                {
                    "ym": "2025-01",
                    "lact1": 1,      # 初産
                    "lact2": 1,      # 2産
                    "lact3plus": 1,   # 3産以上
                    "total": 3        # 合計
                },
                ...
            ]
        """
        conn = self.db.connect()
        cursor = conn.cursor()
        
        # SQL: 月別×産次別分娩頭数（ピボット形式、期間内の全月を表示）
        cursor.execute("""
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(?, 1, 7) AS ym,
                  date(?, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(?, '+1 day')
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
              AND e.event_date >= ?
              AND e.event_date <= ?
              AND e.deleted = 0
            GROUP BY m.ym
            ORDER BY m.ym
        """, (start_date, start_date, end_date, start_date, end_date))
        
        rows = cursor.fetchall()
        
        result = []
        for row in rows:
            result.append({
                "ym": row[0],
                "lact1": row[1] or 0,      # 初産
                "lact2": row[2] or 0,      # 2産
                "lact3plus": row[3] or 0,   # 3産以上
                "total": row[4] or 0        # 合計
            })
        
        logger.info(
            f"calving_by_month_and_lact: {start_date} to {end_date}, "
            f"found {len(result)} months"
        )
        
        return result
    
    def insemination_count_by_month(
        self,
        start_date: str,
        end_date: str
    ) -> List[Dict[str, Any]]:
        """
        月別 授精頭数（AI + ET 合算）
        
        Args:
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
        
        Returns:
            月別の授精頭数リスト
            [
                {"ym": "2024-01", "count": 45},
                ...
            ]
        """
        conn = self.db.connect()
        cursor = conn.cursor()
        
        # SQL: 月別授精頭数（AI=200, ET=201）
        cursor.execute("""
            SELECT
              substr(event_date, 1, 7) AS ym,
              COUNT(*) AS count
            FROM event
            WHERE event_number IN (200, 201)
              AND event_date IS NOT NULL
              AND event_date >= ?
              AND event_date <= ?
              AND deleted = 0
            GROUP BY ym
            ORDER BY ym
        """, (start_date, end_date))
        
        rows = cursor.fetchall()
        
        result = []
        for row in rows:
            result.append({
                "ym": row[0],
                "count": row[1]
            })
        
        logger.info(
            f"insemination_count_by_month: {start_date} to {end_date}, "
            f"found {len(result)} months"
        )
        
        return result
    
    def insemination_by_month_and_lact(
        self,
        start_date: str,
        end_date: str
    ) -> List[Dict[str, Any]]:
        """
        月別 × 産次別 授精頭数（ピボット形式）
        
        Args:
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
        
        Returns:
            月別×産次別の授精頭数リスト（ピボット形式）
            [
                {
                    "ym": "2025-01",
                    "lact1": 1,      # 初産
                    "lact2": 1,      # 2産
                    "lact3plus": 1,   # 3産以上
                    "total": 3        # 合計
                },
                ...
            ]
        """
        conn = self.db.connect()
        cursor = conn.cursor()
        
        # SQL: 月別×産次別授精頭数（ピボット形式、期間内の全月を表示）
        cursor.execute("""
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(?, 1, 7) AS ym,
                  date(?, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(?, '+1 day')
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
              AND e.event_date >= ?
              AND e.event_date <= ?
              AND e.deleted = 0
            GROUP BY m.ym
            ORDER BY m.ym
        """, (start_date, start_date, end_date, start_date, end_date))
        
        rows = cursor.fetchall()
        
        result = []
        for row in rows:
            result.append({
                "ym": row[0],
                "lact1": row[1] or 0,      # 初産
                "lact2": row[2] or 0,      # 2産
                "lact3plus": row[3] or 0,   # 3産以上
                "total": row[4] or 0        # 合計
            })
        
        logger.info(
            f"insemination_by_month_and_lact: {start_date} to {end_date}, "
            f"found {len(result)} months"
        )
        
        return result
    
    def conception_rate_by_month_and_lact(
        self,
        start_date: str,
        end_date: str
    ) -> List[Dict[str, Any]]:
        """
        月別 × 産次別 受胎率（ピボット形式）
        
        定義（固定）:
        - 分母：AI + ET かつ outcome != 'R'（流産・再発情以外）
        - 分子：AI + ET かつ outcome = 'P'（妊娠確定）
        - cow.entr（登録日）が入っている個体は、event_date >= entr の授精のみ集計（メイン画面の受胎率と同一ルール）
        
        Args:
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
        
        Returns:
            月別×産次別の受胎率リスト（ピボット形式）
            [
                {
                    "ym": "2025-01",
                    "lact1_numerator": 5,
                    "lact1_denominator": 10,
                    "lact1_rate": 0.50,
                    "lact2_numerator": 8,
                    "lact2_denominator": 20,
                    "lact2_rate": 0.40,
                    "lact3plus_numerator": 3,
                    "lact3plus_denominator": 15,
                    "lact3plus_rate": 0.20,
                    "total_numerator": 16,
                    "total_denominator": 45,
                    "total_rate": 0.36
                },
                ...
            ]
        """
        conn = self.db.connect()
        cursor = conn.cursor()
        
        # SQL: 月別×産次別で分母（outcome != 'R'）と分子（outcome = 'P'）を集計（ピボット形式、期間内の全月を表示）
        cursor.execute("""
            WITH RECURSIVE months(ym, next_date) AS (
                -- 開始月
                SELECT 
                  substr(?, 1, 7) AS ym,
                  date(?, '+1 month') AS next_date
                UNION ALL
                -- 次の月を生成（終了日まで）
                SELECT 
                  substr(next_date, 1, 7) AS ym,
                  date(next_date, '+1 month') AS next_date
                FROM months
                WHERE next_date <= date(?, '+1 day')
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
              AND e.event_date >= ?
              AND e.event_date <= ?
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
        """, (start_date, start_date, end_date, start_date, end_date))
        
        rows = cursor.fetchall()
        
        result = []
        for row in rows:
            ym = row[0]
            lact1_num = row[1] or 0
            lact1_den = row[2] or 0
            lact2_num = row[3] or 0
            lact2_den = row[4] or 0
            lact3plus_num = row[5] or 0
            lact3plus_den = row[6] or 0
            total_num = row[7] or 0
            total_den = row[8] or 0
            
            # 受胎率を計算
            lact1_rate = lact1_num / lact1_den if lact1_den > 0 else None
            lact2_rate = lact2_num / lact2_den if lact2_den > 0 else None
            lact3plus_rate = lact3plus_num / lact3plus_den if lact3plus_den > 0 else None
            total_rate = total_num / total_den if total_den > 0 else None
            
            result.append({
                "ym": ym,
                "lact1_numerator": lact1_num,
                "lact1_denominator": lact1_den,
                "lact1_rate": lact1_rate,
                "lact2_numerator": lact2_num,
                "lact2_denominator": lact2_den,
                "lact2_rate": lact2_rate,
                "lact3plus_numerator": lact3plus_num,
                "lact3plus_denominator": lact3plus_den,
                "lact3plus_rate": lact3plus_rate,
                "total_numerator": total_num,
                "total_denominator": total_den,
                "total_rate": total_rate
            })
        
        logger.info(
            f"conception_rate_by_month_and_lact: {start_date} to {end_date}, "
            f"found {len(result)} months"
        )
        
        return result

