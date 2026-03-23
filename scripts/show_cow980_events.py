"""
牛ID 980のイベントテーブルを表示
"""
import sqlite3
import sys
from pathlib import Path

# UTF-8出力を有効化
sys.stdout.reconfigure(encoding='utf-8')

db_path = Path("C:/FARMS/DemoFarm3/farm.db")

if not db_path.exists():
    print(f"データベースが見つかりません: {db_path}")
    exit(1)

conn = sqlite3.connect(str(db_path))
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# 牛ID 980を検索（auto_idを取得）
cursor.execute("""
    SELECT auto_id, cow_id FROM cow 
    WHERE ltrim(cow_id, '0') = '980' 
    LIMIT 1
""")
cow_row = cursor.fetchone()

if not cow_row:
    print("牛ID 980が見つかりませんでした")
    conn.close()
    exit(1)

cow_auto_id = cow_row['auto_id']
cow_id = cow_row['cow_id']

print(f"牛ID: {cow_id} (auto_id: {cow_auto_id})")
print("=" * 100)

# イベントテーブルを取得
cursor.execute("""
    SELECT 
        id,
        cow_auto_id,
        event_number,
        event_date,
        event_lact,
        event_dim,
        json_data,
        note,
        deleted
    FROM event
    WHERE cow_auto_id = ? AND deleted = 0
    ORDER BY event_date, id
""", (cow_auto_id,))

rows = cursor.fetchall()

if not rows:
    print("イベントが見つかりませんでした")
else:
    # カラム名を取得
    cols = [desc[0] for desc in cursor.description]
    
    # ヘッダーを表示
    header = " | ".join(f"{col:15}" for col in cols)
    print(header)
    print("-" * 100)
    
    # データを表示
    for row in rows:
        values = []
        for col in cols:
            val = row[col]
            if val is None:
                val = "NULL"
            elif col == 'json_data' and val:
                # json_dataは長い場合は省略
                val_str = str(val)
                if len(val_str) > 30:
                    val_str = val_str[:27] + "..."
                val = val_str
            else:
                val = str(val)
            values.append(f"{val:15}")
        print(" | ".join(values))

conn.close()

