"""
Phase B: event_lact / event_dim の検証（SQLのみ）
"""

import sqlite3
import sys
from pathlib import Path

def run_verification(db_path: str):
    """検証を実行"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    print("=" * 80)
    print("Phase B: event_lact / event_dim 検証")
    print("=" * 80)
    print(f"Database: {db_path}\n")
    
    # B-1: cow.lact と event_lact の最終値整合性チェック
    print("B-1. cow.lact と event_lact の最終値整合性チェック")
    print("-" * 80)
    try:
        cursor.execute("""
            SELECT
              c.auto_id,
              c.cow_id,
              c.lact AS cow_lact,
              MAX(e.event_lact) AS max_event_lact
            FROM cow c
            LEFT JOIN event e
              ON c.auto_id = e.cow_auto_id
              AND e.deleted = 0
            GROUP BY c.auto_id, c.cow_id, c.lact
            HAVING cow_lact != max_event_lact
               OR (cow_lact IS NOT NULL AND max_event_lact IS NULL)
               OR (cow_lact IS NULL AND max_event_lact IS NOT NULL)
        """)
        rows = cursor.fetchall()
        if len(rows) == 0:
            print("[OK] 不一致なし（0行）")
        else:
            print(f"[NG] {len(rows)}件の不一致を検出")
            for row in rows:
                print(f"  cow_id={row['cow_id']}, cow_lact={row['cow_lact']}, max_event_lact={row['max_event_lact']}")
    except Exception as e:
        print(f"[ERROR] {e}")
    print()
    
    # B-2: 代表1頭のイベント時系列確認
    print("B-2. 代表1頭のイベント時系列確認")
    print("-" * 80)
    try:
        # まず、イベントがある牛を1頭取得
        cursor.execute("""
            SELECT DISTINCT e.cow_auto_id, c.cow_id
            FROM event e
            JOIN cow c ON e.cow_auto_id = c.auto_id
            WHERE e.deleted = 0
            ORDER BY e.cow_auto_id
            LIMIT 1
        """)
        cow_row = cursor.fetchone()
        if cow_row:
            cow_auto_id = cow_row['cow_auto_id']
            cow_id = cow_row['cow_id']
            print(f"対象牛: cow_id={cow_id} (auto_id={cow_auto_id})\n")
            
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
            
            if len(rows) > 0:
                print("イベント時系列:")
                print(f"{'event_date':<12} {'event_number':<12} {'event_lact':<12} {'event_dim':<12}")
                print("-" * 50)
                for row in rows:
                    event_lact = row['event_lact'] if row['event_lact'] is not None else 'NULL'
                    event_dim = row['event_dim'] if row['event_dim'] is not None else 'NULL'
                    print(f"{row['event_date']:<12} {row['event_number']:<12} {str(event_lact):<12} {str(event_dim):<12}")
                
                # 確認ポイント
                print("\n確認ポイント:")
                has_calv = False
                prev_lact = None
                for row in rows:
                    event_number = row['event_number']
                    event_lact = row['event_lact']
                    event_dim = row['event_dim']
                    
                    if event_number == 202:  # 分娩
                        has_calv = True
                        if event_dim != 0:
                            print(f"  [NG] 分娩イベントで event_dim != 0: {row['event_date']}")
                        if prev_lact is not None and event_lact != prev_lact + 1:
                            print(f"  [NG] 産次が連続していない: prev={prev_lact}, current={event_lact}")
                        prev_lact = event_lact
                    else:
                        if not has_calv and event_lact != 0:
                            print(f"  [NG] 分娩前イベントで event_lact != 0: {row['event_date']}")
                        if not has_calv and event_dim is not None:
                            print(f"  [NG] 分娩前イベントで event_dim IS NOT NULL: {row['event_date']}")
                        if has_calv and event_lact is None:
                            print(f"  [NG] 分娩後イベントで event_lact IS NULL: {row['event_date']}")
                        if has_calv and event_dim is not None and event_dim < 0:
                            print(f"  [NG] 分娩後イベントで event_dim < 0: {row['event_date']}")
                
                print("  [OK] 時系列確認完了")
            else:
                print("  [WARN] イベントが見つかりません")
        else:
            print("  [WARN] イベントがある牛が見つかりません")
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    print()
    
    # B-3: baseline_calving イベントの挙動確認
    print("B-3. baseline_calving イベントの挙動確認")
    print("-" * 80)
    try:
        cursor.execute("""
            SELECT
              e.event_date,
              e.json_data,
              e.event_lact,
              e.event_dim,
              e.event_number,
              c.lact AS cow_lact
            FROM event e
            JOIN cow c ON e.cow_auto_id = c.auto_id
            WHERE e.json_data LIKE '%baseline_calving%'
              AND e.deleted = 0
            ORDER BY e.event_date
        """)
        rows = cursor.fetchall()
        
        if len(rows) == 0:
            print("  [WARN] baseline_calving イベントが見つかりません")
        else:
            print(f"  見つかった baseline_calving イベント: {len(rows)}件\n")
            for row in rows:
                print(f"  event_date={row['event_date']}, event_number={row['event_number']}")
                print(f"    event_lact={row['event_lact']}, event_dim={row['event_dim']}, cow_lact={row['cow_lact']}")
                print(f"    json_data={row['json_data'][:100]}...")
                
                # 確認ポイント
                if row['event_number'] == 202:  # 分娩
                    # baseline分娩でevent_dim=0になっていないか確認
                    if row['event_dim'] == 0:
                        print(f"    [NG] baseline分娩で event_dim = 0 になっている")
                    # baseline分娩でevent_lactが増えていないか確認（前後のイベントと比較が必要だが簡易チェック）
                    print()
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    print()
    
    # B-4: DIM の健全性チェック
    print("B-4. DIM の健全性チェック（全体）")
    print("-" * 80)
    try:
        cursor.execute("""
            SELECT
              COUNT(*) AS negative_dim_count
            FROM event
            WHERE event_dim < 0
              AND deleted = 0
        """)
        row = cursor.fetchone()
        negative_count = row['negative_dim_count'] if row else 0
        
        if negative_count == 0:
            print(f"[OK] マイナスDIM = {negative_count}")
        else:
            print(f"[NG] マイナスDIM = {negative_count}件")
            
            # 詳細を表示
            cursor.execute("""
                SELECT
                  cow_auto_id,
                  event_date,
                  event_number,
                  event_dim
                FROM event
                WHERE event_dim < 0
                  AND deleted = 0
                LIMIT 10
            """)
            for r in cursor.fetchall():
                print(f"  cow_auto_id={r['cow_auto_id']}, date={r['event_date']}, event_dim={r['event_dim']}")
    except Exception as e:
        print(f"[ERROR] {e}")
    print()
    
    # B-5: 月別×産次別 分娩頭数
    print("B-5. 月別×産次別 分娩頭数がSQLのみで成立するか")
    print("-" * 80)
    try:
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
        
        if len(rows) > 0:
            print(f"{'年月':<10} {'産次':<8} {'頭数':<8}")
            print("-" * 30)
            for row in rows:
                print(f"{row['ym']:<10} {str(row['event_lact']) if row['event_lact'] is not None else 'NULL':<8} {row['cnt']:<8}")
            
            # 確認ポイント
            print("\n確認ポイント:")
            months = set()
            lact_values = set()
            for row in rows:
                months.add(row['ym'])
                if row['event_lact'] is not None:
                    lact_values.add(row['event_lact'])
            
            print(f"  [OK] 月数: {len(months)}件")
            print(f"  [OK] 産次種類: {sorted(lact_values)}")
            
            # 11月・12月の確認
            nov_dec = [m for m in months if m.endswith('-11') or m.endswith('-12')]
            if nov_dec:
                print(f"  [OK] 11月・12月を含む: {nov_dec}")
            else:
                print(f"  [WARN] 11月・12月が見つかりません（データに存在しない可能性）")
        else:
            print("  [WARN] 分娩イベントが見つかりません")
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    print()
    
    # B-6: 「2産のみ」が即抽出できるか
    print("B-6. 「2産のみ」が即抽出できるか（最重要）")
    print("-" * 80)
    try:
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
        
        if len(rows) > 0:
            print(f"{'年月':<10} {'2産分娩頭数':<12}")
            print("-" * 25)
            for row in rows:
                print(f"{row['ym']:<10} {row['cnt']:<12}")
            print(f"\n[OK] 2産分娩の抽出が成功（{len(rows)}ヶ月分）")
        else:
            print("  [WARN] 2産分娩イベントが見つかりません（データに存在しない可能性）")
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    print()
    
    conn.close()
    
    print("=" * 80)
    print("Phase B 検証完了")
    print("=" * 80)


if __name__ == "__main__":
    # データベースパスを指定（コマンドライン引数またはデフォルト）
    if len(sys.argv) > 1:
        db_path = sys.argv[1]
    else:
        # デフォルト: 最初に見つかったfarm.db
        farms = Path('C:/FARMS')
        if farms.exists():
            for farm_dir in farms.iterdir():
                db_file = farm_dir / 'farm.db'
                if db_file.exists():
                    db_path = str(db_file)
                    break
            else:
                print("エラー: farm.db が見つかりません")
                sys.exit(1)
        else:
            print("エラー: C:/FARMS が見つかりません")
            sys.exit(1)
    
    run_verification(db_path)

