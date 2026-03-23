"""
すべての集計メソッドのピボット形式レポートをテスト
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

# 期間を設定
start_date = "2024-12-24"
end_date = "2025-12-24"

print("=" * 100)
print("集計レポート（ピボット形式）")
print(f"期間: {start_date} ～ {end_date}")
print("=" * 100)
print()

# 1. 分娩頭数
print("\n【1. 分娩頭数レポート】")
print("-" * 100)
results = aggregation_service.calving_by_month_and_lact(start_date, end_date)
print(f"{'分娩月':<12} {'初産':<8} {'2産':<8} {'3産以上':<10} {'合計':<8}")
print("-" * 100)
for row in results:
    print(f"{row['ym']:<12} {row['lact1']:<8} {row['lact2']:<8} {row['lact3plus']:<10} {row['total']:<8}")
if results:
    print("-" * 100)
    print(f"合計: {sum(r['total'] for r in results)} 件")

# 2. 授精頭数（産次別）
print("\n【2. 授精頭数レポート（産次別）】")
print("-" * 100)
results = aggregation_service.insemination_by_month_and_lact(start_date, end_date)
print(f"{'授精月':<12} {'初産':<8} {'2産':<8} {'3産以上':<10} {'合計':<8}")
print("-" * 100)
for row in results:
    print(f"{row['ym']:<12} {row['lact1']:<8} {row['lact2']:<8} {row['lact3plus']:<10} {row['total']:<8}")
if results:
    print("-" * 100)
    print(f"合計: {sum(r['total'] for r in results)} 件")

# 3. 受胎率（産次別）
print("\n【3. 受胎率レポート（産次別）】")
print("-" * 100)
results = aggregation_service.conception_rate_by_month_and_lact(start_date, end_date)
print(f"{'授精月':<12} {'初産':<12} {'2産':<12} {'3産以上':<14} {'合計':<12}")
print(f"{'':<12} {'分子/分母/率':<12} {'分子/分母/率':<12} {'分子/分母/率':<14} {'分子/分母/率':<12}")
print("-" * 100)
for row in results:
    lact1_str = f"{row['lact1_numerator']}/{row['lact1_denominator']}/{row['lact1_rate']:.2%}" if row['lact1_rate'] is not None else f"{row['lact1_numerator']}/{row['lact1_denominator']}/-"
    lact2_str = f"{row['lact2_numerator']}/{row['lact2_denominator']}/{row['lact2_rate']:.2%}" if row['lact2_rate'] is not None else f"{row['lact2_numerator']}/{row['lact2_denominator']}/-"
    lact3plus_str = f"{row['lact3plus_numerator']}/{row['lact3plus_denominator']}/{row['lact3plus_rate']:.2%}" if row['lact3plus_rate'] is not None else f"{row['lact3plus_numerator']}/{row['lact3plus_denominator']}/-"
    total_str = f"{row['total_numerator']}/{row['total_denominator']}/{row['total_rate']:.2%}" if row['total_rate'] is not None else f"{row['total_numerator']}/{row['total_denominator']}/-"
    print(f"{row['ym']:<12} {lact1_str:<12} {lact2_str:<12} {lact3plus_str:<14} {total_str:<12}")

print("\n" + "=" * 100)
print("すべてのレポート完了")
print("=" * 100)


















