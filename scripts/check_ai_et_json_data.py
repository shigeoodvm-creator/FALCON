"""
AI/ETイベントのjson_dataを確認し、[P][O][R]の保存場所を特定
"""
import sqlite3
from pathlib import Path
import json

# デモファームのパス
farms = [
    Path("C:/FARMS/デモファーム"),
    Path("C:/FARMS/DemoFarm"),
    Path("C:/FARMS/DemoFarm2"),
    Path("C:/FARMS/DemoFarm3"),
    Path("C:/FARMS/DemoFarm4"),
]

print("=" * 80)
print("AI/ETイベントのjson_data確認")
print("=" * 80)
print()

# 1. eventテーブルの構造を確認
print("【1. eventテーブルの構造確認】")
print("-" * 80)

for farm_path in farms:
    db_path = farm_path / "farm.db"
    if not db_path.exists():
        continue
    
    print(f"\n{farm_path.name}:")
    try:
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()
        
        # テーブル構造を取得
        cursor.execute("PRAGMA table_info(event)")
        columns = cursor.fetchall()
        
        print("カラム一覧:")
        for col in columns:
            print(f"  - {col[1]} ({col[2]})")
        
        # outcome/result/pregnant/P/O/Rに関連するカラムがあるか確認
        column_names = [col[1].lower() for col in columns]
        related_columns = []
        for keyword in ['outcome', 'result', 'pregnant', 'p', 'o', 'r']:
            for col_name in column_names:
                if keyword in col_name:
                    related_columns.append(col_name)
        
        if related_columns:
            print(f"\n関連するカラム: {related_columns}")
        else:
            print("\n関連するカラム: なし")
        
        conn.close()
        break  # 最初のファームで確認すれば十分
        
    except Exception as e:
        print(f"エラー: {e}")

print("\n" + "=" * 80)
print("【2. AI/ETイベントのjson_data実データ確認】")
print("-" * 80)

# 2. AI/ETイベントのjson_dataを確認
for farm_path in farms:
    db_path = farm_path / "farm.db"
    if not db_path.exists():
        continue
    
    print(f"\n{farm_path.name}:")
    print("-" * 80)
    
    try:
        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 指定されたSQLを実行
        cursor.execute("""
            SELECT
                e.id,
                e.event_date,
                e.event_number,
                e.json_data
            FROM event e
            WHERE e.event_number IN (200, 201)
              AND e.deleted = 0
            ORDER BY e.event_date DESC
            LIMIT 20
        """)
        
        rows = cursor.fetchall()
        
        if not rows:
            print("AI/ETイベントが見つかりませんでした。")
            continue
        
        print(f"見つかったAI/ETイベント: {len(rows)}件\n")
        
        for i, row in enumerate(rows, 1):
            event_type = "AI" if row['event_number'] == 200 else "ET"
            print(f"--- {event_type}イベント {i} ---")
            print(f"ID: {row['id']}")
            print(f"日付: {row['event_date']}")
            print(f"イベント番号: {row['event_number']} ({event_type})")
            
            # json_dataをパース
            json_data = row['json_data']
            if json_data:
                try:
                    json_obj = json.loads(json_data)
                    print("json_data内容:")
                    print(json.dumps(json_obj, ensure_ascii=False, indent=2))
                    
                    # outcome/result/pregnant/P/O/Rに関連するキーを探す
                    keys = list(json_obj.keys())
                    related_keys = []
                    for keyword in ['outcome', 'result', 'pregnant', 'p', 'o', 'r']:
                        for key in keys:
                            if keyword.lower() in key.lower():
                                related_keys.append((key, json_obj[key]))
                    
                    if related_keys:
                        print("\n[P][O][R]に関連する可能性のあるキー:")
                        for key, value in related_keys:
                            print(f"  - {key}: {value}")
                    else:
                        print("\n[P][O][R]に関連するキー: 見つかりませんでした")
                    
                except json.JSONDecodeError as e:
                    print(f"JSON解析エラー: {e}")
                    print(f"生データ: {json_data}")
            else:
                print("json_data: (なし)")
            
            print()
        
        conn.close()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 80)
print("【3. まとめ】")
print("-" * 80)
print("確認結果:")
print("1. eventテーブルにoutcome/result/pregnant/P/O/Rのカラムがあるか")
print("2. json_dataにoutcome/result/pregnant/P/O/Rのキーがあるか")
print("3. 上記のいずれにも存在しない場合、UI表示専用ロジックの可能性が高い")

















