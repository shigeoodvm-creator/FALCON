"""
分娩頭数のピボット形式レポートをテスト
"""
import sys
from pathlib import Path

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.aggregation_service import AggregationService

# UTF-8出力を有効化
sys.stdout.reconfigure(encoding='utf-8')

# デモファームのデータベースを使用
db_path = Path("C:/FARMS/DemoFarm3/farm.db")

if not db_path.exists():
    print(f"データベースが見つかりません: {db_path}")
    sys.exit(1)

# サービスを初期化
db_handler = DBHandler(db_path)
aggregation_service = AggregationService(db_handler)

# 分娩頭数を取得（ピボット形式）
start_date = "2024-12-24"
end_date = "2025-12-24"

print("=" * 80)
print("分娩頭数レポート（ピボット形式）")
print(f"期間: {start_date} ～ {end_date}")
print("=" * 80)
print()

results = aggregation_service.calving_by_month_and_lact(start_date, end_date)

# ヘッダーを表示
print(f"{'分娩月':<12} {'初産':<8} {'2産':<8} {'3産以上':<10} {'合計':<8}")
print("-" * 80)

# データを表示
for row in results:
    ym = row["ym"]
    lact1 = row["lact1"]
    lact2 = row["lact2"]
    lact3plus = row["lact3plus"]
    total = row["total"]
    
    print(f"{ym:<12} {lact1:<8} {lact2:<8} {lact3plus:<10} {total:<8}")

# 合計行を追加
if results:
    print("-" * 80)
    total_lact1 = sum(r['lact1'] for r in results)
    total_lact2 = sum(r['lact2'] for r in results)
    total_lact3plus = sum(r['lact3plus'] for r in results)
    total_all = sum(r['total'] for r in results)
    print(f"{'合計':<12} {total_lact1:<8} {total_lact2:<8} {total_lact3plus:<10} {total_all:<8}")

print("=" * 80)

