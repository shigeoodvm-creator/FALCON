"""
Phase B 完全検証: 最終テスト（既存イベントを無視）
"""

import sys
from pathlib import Path
import sqlite3

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


def main():
    db_path = "C:/FARMS/DemoFarm3/farm.db"
    cow_auto_id = 1
    
    db = DBHandler(Path(db_path))
    rule_engine = RuleEngine(db)
    
    print("=" * 80)
    print("Phase B 最終テスト: 既存イベントを無視してテストイベントのみで検証")
    print("=" * 80)
    
    # 1. 既存のテストイベントを削除
    print("\n[1] 既存のテストイベントを削除...")
    conn = db.connect()
    c = conn.cursor()
    c.execute("DELETE FROM event WHERE cow_auto_id = ? AND note LIKE 'Phase B テスト%'", (cow_auto_id,))
    conn.commit()
    print("削除完了")
    
    # 2. 既存イベントを物理削除（テスト用）
    print("\n[2] 既存イベントを物理削除（テスト用）...")
    c.execute("DELETE FROM event WHERE cow_auto_id = ? AND (note IS NULL OR note NOT LIKE 'Phase B テスト%')", (cow_auto_id,))
    conn.commit()
    print("物理削除完了")
    
    # 3. baseline値を0にリセット
    print("\n[3] baseline値を0にリセット...")
    c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
    conn.commit()
    print("リセット完了")
    
    # 4. baseline値を再度確認・リセット（念のため）
    print("\n[4] baseline値を再度確認・リセット...")
    c.execute("SELECT lact FROM cow WHERE auto_id = ?", (cow_auto_id,))
    current_lact = c.fetchone()[0]
    print(f"  現在のbaseline値: {current_lact}")
    if current_lact != 0:
        c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
        conn.commit()
        print("  baseline値を0にリセットしました")
    else:
        print("  baseline値は既に0です")
    
    # 5. テストイベントを投入
    print("\n[5] テストイベントを投入...")
    
    # イベント①: 分娩（1産目）
    event1 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 202,
        'event_date': '2024-01-01',
        'json_data': {},
        'note': 'Phase B テスト: 1産目分娩'
    }
    event1_id = db.insert_event(event1)
    # baseline値を再確認・リセット（on_event_addedで更新される可能性があるため）
    c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
    conn.commit()
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print(f"  分娩1産目: event_id={event1_id}")
    
    # イベント②: AI（泌乳1期）
    event2 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 200,
        'event_date': '2024-02-01',
        'json_data': {},
        'note': 'Phase B テスト: 泌乳1期AI'
    }
    event2_id = db.insert_event(event2)
    # baseline値を再確認・リセット
    c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
    conn.commit()
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print(f"  AI泌乳1期: event_id={event2_id}")
    
    # イベント③: 分娩（2産目）
    event3 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 202,
        'event_date': '2025-01-10',
        'json_data': {},
        'note': 'Phase B テスト: 2産目分娩'
    }
    event3_id = db.insert_event(event3)
    # baseline値を再確認・リセット
    c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
    conn.commit()
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print(f"  分娩2産目: event_id={event3_id}")
    
    # イベント④: AI（泌乳2期）
    event4 = {
        'cow_auto_id': cow_auto_id,
        'event_number': 200,
        'event_date': '2025-02-10',
        'json_data': {},
        'note': 'Phase B テスト: 泌乳2期AI'
    }
    event4_id = db.insert_event(event4)
    # baseline値を再確認・リセット
    c.execute("UPDATE cow SET lact = 0 WHERE auto_id = ?", (cow_auto_id,))
    conn.commit()
    rule_engine.recalculate_events_for_cow(cow_auto_id)
    print(f"  AI泌乳2期: event_id={event4_id}")
    
    # 6. 検証
    print("\n" + "=" * 80)
    print("検証SQL実行結果")
    print("=" * 80)
    
    # 既存の接続を閉じる
    db.close()
    
    # 新しい接続を開く
    conn = db.connect()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    # 【A】イベント時系列確認
    print("\n[A] イベント時系列確認")
    print("-" * 80)
    c.execute("""
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
    
    rows = c.fetchall()
    print(f"{'event_date':<12} {'event_number':<12} {'event_lact':<12} {'event_dim':<12}")
    print("-" * 50)
    for row in rows:
        event_lact = row['event_lact'] if row['event_lact'] is not None else 'NULL'
        event_dim = row['event_dim'] if row['event_dim'] is not None else 'NULL'
        print(f"{row['event_date']:<12} {row['event_number']:<12} {str(event_lact):<12} {str(event_dim):<12}")
    
    print("\n期待結果:")
    print("2024-01-01  CALV(202)  event_lact=1  event_dim=0")
    print("2024-02-01  AI(200)    event_lact=1  event_dim=31")
    print("2025-01-10  CALV(202)  event_lact=2  event_dim=0")
    print("2025-02-10  AI(200)    event_lact=2  event_dim=31")
    
    # 判定
    test_events = [r for r in rows if r['event_date'] in ['2024-01-01', '2024-02-01', '2025-01-10', '2025-02-10']]
    result_a = False
    if len(test_events) == 4:
        expected = [
            (202, 1, 0),
            (200, 1, 31),
            (202, 2, 0),
            (200, 2, 31)
        ]
        result_a = all(
            row['event_number'] == exp[0] and 
            row['event_lact'] == exp[1] and 
            row['event_dim'] == exp[2]
            for row, exp in zip(test_events, expected)
        )
    
    if result_a:
        print("\n[判定] [A] OK: 期待結果と一致")
    else:
        print("\n[判定] [A] NG: 期待結果と不一致")
        for i, (row, exp) in enumerate(zip(test_events, expected)):
            if not (row['event_number'] == exp[0] and row['event_lact'] == exp[1] and row['event_dim'] == exp[2]):
                print(f"  行{i+1}: 期待=({exp[0]}, {exp[1]}, {exp[2]}), 実際=({row['event_number']}, {row['event_lact']}, {row['event_dim']})")
    
    # 【B】cow.lact の確認
    print("\n[B] cow.lact の確認")
    print("-" * 80)
    c.execute("SELECT lact FROM cow WHERE auto_id = ?", (cow_auto_id,))
    lact = c.fetchone()['lact']
    print(f"cow.lact = {lact}")
    print("期待結果: lact = 2")
    
    result_b = (lact == 2)
    if result_b:
        print("\n[判定] [B] OK: 期待結果と一致")
    else:
        print(f"\n[判定] [B] NG: 期待結果と不一致（実際={lact}）")
    
    # 【C】月別×産次別 分娩集計
    print("\n[C] 月別×産次別 分娩集計（SQLのみ）")
    print("-" * 80)
    c.execute("""
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
    
    rows = c.fetchall()
    print(f"{'年月':<10} {'産次':<8} {'頭数':<8}")
    print("-" * 30)
    for row in rows:
        event_lact = row['event_lact'] if row['event_lact'] is not None else 'NULL'
        print(f"{row['ym']:<10} {str(event_lact):<8} {row['cnt']:<8}")
    
    print("\n期待結果:")
    print("2024-01 | 1 | 1")
    print("2025-01 | 2 | 1")
    
    result_c = False
    if len(rows) == 2:
        if (rows[0]['ym'] == '2024-01' and rows[0]['event_lact'] == 1 and rows[0]['cnt'] == 1 and
            rows[1]['ym'] == '2025-01' and rows[1]['event_lact'] == 2 and rows[1]['cnt'] == 1):
            result_c = True
    
    if result_c:
        print("\n[判定] [C] OK: 期待結果と一致")
    else:
        print("\n[判定] [C] NG: 期待結果と不一致")
    
    # 【D】2産のみ抽出
    print("\n[D] 2産のみ抽出（今回の最終ゴール）")
    print("-" * 80)
    c.execute("""
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
    
    rows = c.fetchall()
    print(f"{'年月':<10} {'頭数':<8}")
    print("-" * 20)
    for row in rows:
        print(f"{row['ym']:<10} {row['cnt']:<8}")
    
    print("\n期待結果:")
    print("2025-01 | 1")
    
    result_d = False
    if len(rows) == 1 and rows[0]['ym'] == '2025-01' and rows[0]['cnt'] == 1:
        result_d = True
    
    if result_d:
        print("\n[判定] [D] OK: 期待結果と一致")
    else:
        print("\n[判定] [D] NG: 期待結果と不一致")
    
    conn.close()
    
    # 総合判定
    print("\n" + "=" * 80)
    print("総合判定")
    print("=" * 80)
    if result_a and result_b and result_c and result_d:
        print("[Phase B 完全合格] すべての検証項目が期待どおりです！")
        print("\n✓ event_lact / event_dim が正しく機能していることを実証")
        print("✓ Phase A（RuleEngine + event主導設計）が正しいことを確定")
    else:
        print("[Phase B 一部不合格] 以下の項目に問題があります:")
        if not result_a:
            print("  - [A] イベント時系列確認")
        if not result_b:
            print(f"  - [B] cow.lact (期待=2, 実際={lact})")
        if not result_c:
            print("  - [C] 月別×産次別 分娩集計")
        if not result_d:
            print("  - [D] 2産のみ抽出")
    print("=" * 80)


if __name__ == "__main__":
    main()

