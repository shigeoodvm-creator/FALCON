"""
Phase C: AggregationService のテスト
Phase B で投入したテストデータを使用
"""

import sys
from pathlib import Path

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.aggregation_service import AggregationService


def main():
    db_path = Path("C:/FARMS/DemoFarm3/farm.db")
    db = DBHandler(db_path)
    service = AggregationService(db)
    
    print("=" * 80)
    print("Phase C: AggregationService テスト")
    print("=" * 80)
    
    # Phase B で投入したテストデータの期間
    start_date = "2024-01-01"
    end_date = "2025-12-31"
    
    # [C-1] 月別 × 産次別 分娩頭数
    print("\n[C-1] 月別 × 産次別 分娩頭数")
    print("-" * 80)
    result = service.calving_by_month_and_lact(start_date, end_date)
    print(f"{'年月':<10} {'産次':<8} {'頭数':<8}")
    print("-" * 30)
    for row in result:
        lact_str = str(row['lact']) if row['lact'] is not None else 'NULL'
        print(f"{row['ym']:<10} {lact_str:<8} {row['count']:<8}")
    
    print("\n期待結果:")
    print("2024-01 | 1 | 1")
    print("2025-01 | 2 | 1")
    
    # 判定
    expected_calv = [
        {"ym": "2024-01", "lact": 1, "count": 1},
        {"ym": "2025-01", "lact": 2, "count": 1}
    ]
    result_c1 = all(
        any(r['ym'] == exp['ym'] and r['lact'] == exp['lact'] and r['count'] == exp['count'] 
            for r in result)
        for exp in expected_calv
    )
    if result_c1:
        print("\n[判定] [C-1] OK: 期待結果と一致")
    else:
        print("\n[判定] [C-1] NG: 期待結果と不一致")
    
    # [C-2] 月別 授精頭数
    print("\n[C-2] 月別 授精頭数（AI + ET 合算）")
    print("-" * 80)
    result = service.insemination_count_by_month(start_date, end_date)
    print(f"{'年月':<10} {'頭数':<8}")
    print("-" * 20)
    for row in result:
        print(f"{row['ym']:<10} {row['count']:<8}")
    
    print("\n期待結果:")
    print("2024-02 | 1 (AI)")
    print("2025-02 | 1 (AI)")
    
    # 判定
    expected_insem = [
        {"ym": "2024-02", "count": 1},
        {"ym": "2025-02", "count": 1}
    ]
    result_c2 = all(
        any(r['ym'] == exp['ym'] and r['count'] == exp['count'] for r in result)
        for exp in expected_insem
    )
    if result_c2:
        print("\n[判定] [C-2] OK: 期待結果と一致")
    else:
        print("\n[判定] [C-2] NG: 期待結果と不一致")
    
    # [C-3] 月別 × 産次別 受胎率
    print("\n[C-3] 月別 × 産次別 受胎率")
    print("-" * 80)
    result = service.conception_rate_by_month_and_lact(start_date, end_date)
    print(f"{'年月':<10} {'産次':<8} {'分子':<8} {'分母':<8} {'受胎率':<10}")
    print("-" * 50)
    for row in result:
        lact_str = str(row['lact']) if row['lact'] is not None else 'NULL'
        rate_str = f"{row['rate']:.2f}" if row['rate'] is not None else "NULL"
        print(f"{row['ym']:<10} {lact_str:<8} {row['numerator']:<8} {row['denominator']:<8} {rate_str:<10}")
    
    print("\n期待結果:")
    print("（テストデータにはoutcome情報がないため、分母=2, 分子=0が想定される）")
    
    # 総合判定
    print("\n" + "=" * 80)
    print("総合判定")
    print("=" * 80)
    if result_c1 and result_c2:
        print("[Phase C 基本機能 OK] 分娩頭数・授精頭数の集計が正しく動作")
        print("（受胎率はoutcome情報が必要なため、実データで確認が必要）")
    else:
        print("[Phase C 一部不合格] 以下の項目に問題があります:")
        if not result_c1:
            print("  - [C-1] 月別×産次別分娩頭数")
        if not result_c2:
            print("  - [C-2] 月別授精頭数")
    print("=" * 80)


if __name__ == "__main__":
    main()


















