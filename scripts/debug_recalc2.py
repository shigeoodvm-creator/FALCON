"""デバッグ: 再計算ロジックの詳細確認"""
import sys
from pathlib import Path
import sqlite3
sys.path.insert(0, str(Path('app')))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

db = DBHandler(Path('C:/FARMS/DemoFarm3/farm.db'))

# baseline値を0にリセット
conn = db.connect()
c = conn.cursor()
c.execute('UPDATE cow SET lact=0 WHERE auto_id=1')
conn.commit()

# baseline値を確認
c.execute('SELECT lact FROM cow WHERE auto_id=1')
baseline_lact = c.fetchone()[0]
print(f'baseline値（リセット後）: cow.lact={baseline_lact}')
conn.close()
db.close()  # 接続をリセット

# イベントを確認
events = db.get_events_by_cow(1, include_deleted=False)
events.sort(key=lambda e: (e.get('event_date', ''), e.get('id', 0)))
print(f'\nイベント数: {len(events)}')
for e in events:
    print(f'  {e.get("event_date")} | {e.get("event_number")} | id={e.get("id")}')

# 再計算実行
rule_engine = RuleEngine(db)
print('\n再計算実行...')
rule_engine.recalculate_events_for_cow(1)

# 結果確認
conn = db.connect()
c = conn.cursor()
c.execute('SELECT event_date, event_number, event_lact, event_dim FROM event WHERE cow_auto_id=1 AND deleted=0 ORDER BY event_date, id')
rows = c.fetchall()
print('\n再計算後のイベント:')
for r in rows:
    print(f'  {r[0]} | {r[1]} | lact={r[2]} | dim={r[3]}')

c.execute('SELECT lact FROM cow WHERE auto_id=1')
print(f'\ncow.lact={c.fetchone()[0]}')
conn.close()

