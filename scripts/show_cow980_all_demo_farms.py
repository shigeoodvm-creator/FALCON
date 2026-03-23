"""
C:/FARMS内のすべてのデモファームで牛ID 980のイベントテーブルを表示
"""
import sqlite3
import sys
from pathlib import Path

# UTF-8出力を有効化
sys.stdout.reconfigure(encoding='utf-8')

farms_root = Path("C:/FARMS")
demo_farms = []

# デモファームを検索
for item in farms_root.iterdir():
    if item.is_dir():
        db_path = item / "farm.db"
        if db_path.exists():
            demo_farms.append((item.name, db_path))

if not demo_farms:
    print("デモファームが見つかりませんでした")
    exit(1)

# 各デモファームで牛ID 980のイベントを検索
for farm_name, db_path in sorted(demo_farms):
    print("\n" + "=" * 100)
    print(f"農場: {farm_name}")
    print(f"データベース: {db_path}")
    print("=" * 100)
    
    try:
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
            print(f"牛ID 980が見つかりませんでした")
            conn.close()
            continue
        
        cow_auto_id = cow_row['auto_id']
        cow_id = cow_row['cow_id']
        
        print(f"牛ID: {cow_id} (auto_id: {cow_auto_id})")
        print("-" * 100)
        
        # テーブルスキーマを確認（event_lactカラムの存在確認）
        cursor.execute("PRAGMA table_info(event)")
        columns_info = cursor.fetchall()
        column_names = [col[1] for col in columns_info]
        
        # カラムリストを動的に構築
        base_columns = ['id', 'cow_auto_id', 'event_number', 'event_date']
        if 'event_lact' in column_names:
            base_columns.append('event_lact')
        if 'event_dim' in column_names:
            base_columns.append('event_dim')
        base_columns.extend(['json_data', 'note', 'deleted'])
        
        # イベントテーブルを取得
        columns_str = ', '.join(base_columns)
        cursor.execute(f"""
            SELECT {columns_str}
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
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        continue

print("\n" + "=" * 100)
print("検索完了")
print("=" * 100)

