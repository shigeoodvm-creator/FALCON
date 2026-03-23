"""
Phase B 完全検証: テスト用分娩イベント投入
"""

import sys
import sqlite3
import json
from pathlib import Path

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


def insert_test_events(db_path: str, cow_auto_id: int):
    """
    テスト用イベントを投入
    
    Args:
        db_path: データベースパス
        cow_auto_id: 対象牛のauto_id
    """
    db = DBHandler(Path(db_path))
    rule_engine = RuleEngine(db)
    
    print(f"対象牛: cow_auto_id={cow_auto_id}")
    
    # イベント①: 分娩（1産目）
    print("\n[1] 分娩イベント（1産目）を投入...")
    event1 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 202,  # CALV
        'event_date': '2024-01-01',
        'json_data': {},
        'note': 'Phase B テスト: 1産目分娩'
    }
    event1_id = db.insert_event(event1)
    print(f"  投入完了: event_id={event1_id}")
    
    # 再計算
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print("  再計算完了")
    
    # イベント②: AI（泌乳1期）
    print("\n[2] AIイベント（泌乳1期）を投入...")
    event2 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 200,  # AI
        'event_date': '2024-02-01',
        'json_data': {},
        'note': 'Phase B テスト: 泌乳1期AI'
    }
    event2_id = db.insert_event(event2)
    print(f"  投入完了: event_id={event2_id}")
    
    # 再計算
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print("  再計算完了")
    
    # イベント③: 分娩（2産目）
    print("\n[3] 分娩イベント（2産目）を投入...")
    event3 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 202,  # CALV
        'event_date': '2025-01-10',
        'json_data': {},
        'note': 'Phase B テスト: 2産目分娩'
    }
    event3_id = db.insert_event(event3)
    print(f"  投入完了: event_id={event3_id}")
    
    # 再計算
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print("  再計算完了")
    
    # イベント④: AI（泌乳2期）
    print("\n[4] AIイベント（泌乳2期）を投入...")
    event4 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 200,  # AI
        'event_date': '2025-02-10',
        'json_data': {},
        'note': 'Phase B テスト: 泌乳2期AI'
    }
    event4_id = db.insert_event(event4)
    print(f"  投入完了: event_id={event4_id}")
    
    # 再計算
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print("  再計算完了")
    
    print("\n" + "=" * 80)
    print("テストイベント投入完了")
    print("=" * 80)
    
    return [event1_id, event2_id, event3_id, event4_id]


def verify_results(db_path: str, cow_auto_id: int):
    """
    検証SQLを実行して結果を確認
    
    Args:
        db_path: データベースパス
        cow_auto_id: 対象牛のauto_id
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    print("\n" + "=" * 80)
    print("検証SQL実行結果")
    print("=" * 80)
    
    # 【A】イベント時系列確認
    print("\n[A] イベント時系列確認")
    print("-" * 80)
    cursor.execute("""
        SELECT
          event_date,
          event_number,
          event_lact,
          event_dim
        FROM event
        WHERE cow_auto_id = ?
          AND deleted = 0
        ORDER BY event_date, id
    """, (cow_auto_id,))
    
    rows = cursor.fetchall()
    print(f"{'event_date':<12} {'event_number':<12} {'event_lact':<12} {'event_dim':<12}")
    print("-" * 50)
    for row in rows:
        event_lact = row['event_lact'] if row['event_lact'] is not None else 'NULL'
        event_dim = row['event_dim'] if row['event_dim'] is not None else 'NULL'
        print(f"{row['event_date']:<12} {row['event_number']:<12} {str(event_lact):<12} {str(event_dim):<12}")
    
    # 期待結果との比較
    print("\n期待結果:")
    print("2024-01-01  CALV(202)  event_lact=1  event_dim=0")
    print("2024-02-01  AI(200)    event_lact=1  event_dim=31")
    print("2025-01-10  CALV(202)  event_lact=2  event_dim=0")
    print("2025-02-10  AI(200)    event_lact=2  event_dim=31")
    
    # 【B】cow.lact の確認
    print("\n[B] cow.lact の確認")
    print("-" * 80)
    cursor.execute("SELECT lact FROM cow WHERE auto_id = ?", (cow_auto_id,))
    row = cursor.fetchone()
    lact = row['lact'] if row else None
    print(f"cow.lact = {lact}")
    print("期待結果: lact = 2")
    
    # 【C】月別×産次別 分娩集計
    print("\n[C] 月別×産次別 分娩集計（SQLのみ）")
    print("-" * 80)
    cursor.execute("""
        SELECT
          substr(event_date,1,7) AS ym,
          event_lact,
          COUNT(*) AS cnt
        FROM event
        WHERE event_number = 202
          AND deleted = 0
        GROUP BY ym, event_lact
        ORDER BY ym, event_lact
    """)
    
    rows = cursor.fetchall()
    print(f"{'年月':<10} {'産次':<8} {'頭数':<8}")
    print("-" * 30)
    for row in rows:
        event_lact = row['event_lact'] if row['event_lact'] is not None else 'NULL'
        print(f"{row['ym']:<10} {str(event_lact):<8} {row['cnt']:<8}")
    
    print("\n期待結果:")
    print("2024-01 | 1 | 1")
    print("2025-01 | 2 | 1")
    
    # 【D】2産のみ抽出
    print("\n[D] 2産のみ抽出（今回の最終ゴール）")
    print("-" * 80)
    cursor.execute("""
        SELECT
          substr(event_date,1,7) AS ym,
          COUNT(*) AS cnt
        FROM event
        WHERE event_number = 202
          AND event_lact = 2
          AND deleted = 0
        GROUP BY ym
        ORDER BY ym
    """)
    
    rows = cursor.fetchall()
    print(f"{'年月':<10} {'頭数':<8}")
    print("-" * 20)
    for row in rows:
        print(f"{row['ym']:<10} {row['cnt']:<8}")
    
    print("\n期待結果:")
    print("2025-01 | 1")
    
    conn.close()
    
    print("\n" + "=" * 80)
    print("検証完了")
    print("=" * 80)


if __name__ == "__main__":
    db_path = "C:/FARMS/DemoFarm3/farm.db"
    
    # 対象牛を選定（最初の1頭）
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT auto_id, cow_id FROM cow ORDER BY auto_id LIMIT 1")
    cow_row = cursor.fetchone()
    conn.close()
    
    if not cow_row:
        print("エラー: 対象牛が見つかりません")
        sys.exit(1)
    
    cow_auto_id = cow_row['auto_id']
    cow_id = cow_row['cow_id']
    
    print("=" * 80)
    print("Phase B 完全検証: テスト用分娩イベント投入")
    print("=" * 80)
    print(f"対象牛: cow_auto_id={cow_auto_id}, cow_id={cow_id}")
    print(f"データベース: {db_path}")
    print()
    
    # テストイベント投入
    event_ids = insert_test_events(db_path, cow_auto_id)
    
    # 検証
    verify_results(db_path, cow_auto_id)
    
    print(f"\n投入したイベントID: {event_ids}")
    print("検証完了しました。")

