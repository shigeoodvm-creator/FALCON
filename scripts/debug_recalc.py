"""デバッグ: 再計算ロジックの確認"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path('app')))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

db = DBHandler(Path('C:/FARMS/DemoFarm3/farm.db'))
cow = db.get_cow_by_auto_id(1)
print(f'baseline値: cow.lact={cow.get("lact")}')

rule_engine = RuleEngine(db)
print('\n再計算実行...')
rule_engine.recalculate_events_for_cow(1)

import sqlite3
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


















