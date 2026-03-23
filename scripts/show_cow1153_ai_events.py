"""
牛ID 1153のAIイベントのJSONデータを表示
"""
import sqlite3
from pathlib import Path
import json

# デモファームのパスリスト
farms = [
    Path("C:/FARMS/デモファーム"),
    Path("C:/FARMS/DemoFarm"),
    Path("C:/FARMS/DemoFarm2"),
    Path("C:/FARMS/DemoFarm3"),
    Path("C:/FARMS/DemoFarm4"),
]

cow_id = "1153"
event_number = 200  # AIイベント

print(f"牛ID {cow_id} のAIイベント（event_number={event_number}）のJSONデータを検索中...\n")

found = False
for farm_path in farms:
    db_path = farm_path / "farm.db"
    if not db_path.exists():
        continue
    
    print(f"=== {farm_path.name} ===")
    
    try:
        conn = sqlite3.connect(str(db_path))
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 牛ID 1153のAIイベントを取得
        cursor.execute("""
            SELECT 
                c.cow_id,
                c.jpn10,
                e.id AS event_id,
                e.event_date,
                e.json_data,
                e.note
            FROM cow c
            JOIN event e ON c.auto_id = e.cow_auto_id
            WHERE c.cow_id = ?
              AND e.event_number = ?
              AND e.deleted = 0
            ORDER BY e.event_date
        """, (cow_id, event_number))
        
        rows = cursor.fetchall()
        
        if rows:
            found = True
            print(f"見つかりました: {len(rows)}件のAIイベント\n")
            
            for i, row in enumerate(rows, 1):
                print(f"--- AIイベント {i} ---")
                print(f"イベントID: {row['event_id']}")
                print(f"日付: {row['event_date']}")
                print(f"個体識別番号: {row['jpn10']}")
                print(f"備考: {row['note'] or '(なし)'}")
                
                # JSONデータをパースして表示
                json_data = row['json_data']
                if json_data:
                    try:
                        json_obj = json.loads(json_data)
                        print(f"JSONデータ:")
                        print(json.dumps(json_obj, ensure_ascii=False, indent=2))
                    except json.JSONDecodeError as e:
                        print(f"JSON解析エラー: {e}")
                        print(f"生データ: {json_data}")
                else:
                    print("JSONデータ: (なし)")
                print()
            
        else:
            print(f"牛ID {cow_id} のAIイベントは見つかりませんでした。\n")
        
        conn.close()
        
    except Exception as e:
        print(f"エラー: {e}\n")

if not found:
    print("すべてのデモファームで牛ID 1153のAIイベントが見つかりませんでした。")

















