# Phase C: AggregationService 実装サマリー

## 概要

Phase Cとして、event テーブル前提の「公式・最小・堅牢な集計エンジン」を実装しました。

**ファイル**: `app/modules/aggregation_service.py`

**クラス**: `AggregationService`

## 設計原則

1. **event テーブルのみを参照**
   - `event.event_lact` / `event.event_dim` を直接使用
   - `cow.lact` は参照しない
   - Python側での再計算は禁止

2. **SQL のみで集計**
   - Python側でのGROUP BY/再集計は禁止
   - 可読性重視（JOIN最小）

3. **推論・再計算を一切行わない**
   - 再現性100%の公式数値を返す
   - RuleEngine は使用しない

## 実装メソッド

### [C-1] calving_by_month_and_lact

**メソッド**: `calving_by_month_and_lact(start_date, end_date)`

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  event_lact AS lact,
  COUNT(*) AS count
FROM event
WHERE event_number = 202
  AND event_date IS NOT NULL
  AND event_date >= ?
  AND event_date <= ?
  AND deleted = 0
GROUP BY ym, lact
ORDER BY ym, lact
```

**戻り値**:
```python
[
  {"ym": "2024-01", "lact": 1, "count": 12},
  {"ym": "2024-01", "lact": 2, "count": 8},
  ...
]
```

**テスト結果**: ✅ OK（Phase Bテストデータで検証済み）

---

### [C-2] insemination_count_by_month

**メソッド**: `insemination_count_by_month(start_date, end_date)`

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  COUNT(*) AS count
FROM event
WHERE event_number IN (200, 201)
  AND event_date IS NOT NULL
  AND event_date >= ?
  AND event_date <= ?
  AND deleted = 0
GROUP BY ym
ORDER BY ym
```

**戻り値**:
```python
[
  {"ym": "2024-01", "count": 45},
  ...
]
```

**テスト結果**: ✅ OK（Phase Bテストデータで検証済み）

---

### [C-3] conception_rate_by_month_and_lact

**メソッド**: `conception_rate_by_month_and_lact(start_date, end_date)`

**定義（固定）**:
- **分母**: AI + ET かつ `outcome != 'R'`（流産・再発情以外）
- **分子**: AI + ET かつ `outcome = 'P'`（妊娠確定）

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  event_lact AS lact,
  SUM(CASE 
    WHEN json_data IS NOT NULL 
      AND (
        json_data LIKE '%"outcome":"P"%' 
        OR json_data LIKE '%"outcome": "P"%'
      )
    THEN 1 
    ELSE 0 
  END) AS numerator,
  SUM(CASE 
    WHEN json_data IS NULL 
      OR (
        json_data NOT LIKE '%"outcome":"R"%' 
        AND json_data NOT LIKE '%"outcome": "R"%'
      )
    THEN 1 
    ELSE 0 
  END) AS denominator
FROM event
WHERE event_number IN (200, 201)
  AND event_date IS NOT NULL
  AND event_date >= ?
  AND event_date <= ?
  AND deleted = 0
GROUP BY ym, lact
ORDER BY ym, lact
```

**戻り値**:
```python
[
  {
    "ym": "2024-01",
    "lact": 2,
    "numerator": 8,
    "denominator": 20,
    "rate": 0.40
  },
  ...
]
```

**注意**: `denominator=0` の場合は `rate=None`

**テスト結果**: ⚠️ テストデータにはoutcome情報がないため、実データで確認が必要

---

## 実装上の注意事項

1. **deleted フラグの除外**: すべてのクエリで `deleted = 0` を条件に含む
2. **event_date の NULL チェック**: `event_date IS NOT NULL` を条件に含む
3. **接続管理**: 各メソッドで独立して接続を開いて閉じる（`self.db.close()` → `self.db.connect()`）

## テスト結果

Phase B で投入したテストデータを使用して検証:

- ✅ [C-1] 月別×産次別分娩頭数: 期待結果と一致
- ✅ [C-2] 月別授精頭数: 期待結果と一致
- ⚠️ [C-3] 月別×産次別受胎率: outcome情報が必要（実データで確認が必要）

## 完了条件

- ✅ `aggregation_service.py` が単体で動作
- ✅ SQL のみで集計が成立
- ✅ Phase B のテストデータで正しい数値が返る

## 今後の拡張

- 実データでの受胎率計算の検証
- その他の集計メソッドの追加（必要に応じて）


















