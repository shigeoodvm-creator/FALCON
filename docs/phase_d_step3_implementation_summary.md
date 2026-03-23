# Phase D - Step 3: 分析モード専用の公式SQLテンプレート集 実装サマリー

## 概要

分析モード専用の「公式SQLテンプレート集」を実装しました。分析AIに自由にSQLを考えさせず、人間が定義した「正しいSQL」を再利用させることで、集計定義のブレを完全に防ぎます。

**実装ファイル**: 
- `app/modules/analysis_sql_templates.py`（新規）
- `app/ui/main_window.py`（修正）

## 実装内容

### 1. SQLテンプレート管理クラス

**ファイル**: `app/modules/analysis_sql_templates.py`

**クラス**: `AnalysisSQLTemplates`

**機能**:
- 辞書形式でSQLテンプレートを保持
- key は分析目的の論理名
- value は完成済みの SELECT 文（パラメータ付き）

**主要メソッド**:
- `get_template(template_name)`: テンプレートを取得
- `list_templates()`: 利用可能なテンプレート一覧を取得
- `expand_template(template_name, params)`: テンプレートを展開（パラメータを置換）
- `get_template_list_for_ai()`: AI用のテンプレート一覧を取得（プロンプト用）

### 2. 実装済みテンプレート（必須5つ）

#### ① 月別×産次別 分娩頭数

**テンプレート名**: `calving_by_month_lact`

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  event_lact AS lact,
  COUNT(*) AS cnt
FROM event
WHERE event_number = 202
  AND event_date IS NOT NULL
  AND event_date >= :start
  AND event_date <= :end
  AND deleted = 0
GROUP BY ym, event_lact
ORDER BY ym, event_lact
```

**パラメータ**: `start`, `end`（日付形式 YYYY-MM-DD）

---

#### ② 月別×産次別 授精頭数（AI+ET）

**テンプレート名**: `insemination_by_month_lact`

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  event_lact AS lact,
  COUNT(*) AS cnt
FROM event
WHERE event_number IN (200, 201)
  AND event_date IS NOT NULL
  AND event_date >= :start
  AND event_date <= :end
  AND deleted = 0
GROUP BY ym, event_lact
ORDER BY ym, event_lact
```

**パラメータ**: `start`, `end`

---

#### ③ 月別×産次別 受胎率

**テンプレート名**: `conception_rate_by_month_lact`

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
  AND event_date >= :start
  AND event_date <= :end
  AND deleted = 0
GROUP BY ym, event_lact
ORDER BY ym, event_lact
```

**パラメータ**: `start`, `end`

**定義**:
- 分母: AI + ET かつ `outcome != 'R'`
- 分子: AI + ET かつ `outcome = 'P'`

---

#### ④ 2産のみ 抽出（分娩）

**テンプレート名**: `calving_lact2_only`

**SQL**:
```sql
SELECT
  substr(event_date, 1, 7) AS ym,
  COUNT(*) AS cnt
FROM event
WHERE event_number = 202
  AND event_lact = 2
  AND event_date IS NOT NULL
  AND event_date >= :start
  AND event_date <= :end
  AND deleted = 0
GROUP BY ym
ORDER BY ym
```

**パラメータ**: `start`, `end`

---

#### ⑤ event_dim を使った DIM 分布

**テンプレート名**: `dim_distribution`

**SQL**:
```sql
SELECT
  CASE
    WHEN event_dim IS NULL THEN 'NULL'
    WHEN event_dim < 0 THEN '異常値'
    WHEN event_dim <= 30 THEN '0-30日'
    WHEN event_dim <= 60 THEN '31-60日'
    WHEN event_dim <= 90 THEN '61-90日'
    WHEN event_dim <= 120 THEN '91-120日'
    WHEN event_dim <= 150 THEN '121-150日'
    WHEN event_dim <= 180 THEN '151-180日'
    WHEN event_dim <= 210 THEN '181-210日'
    WHEN event_dim <= 240 THEN '211-240日'
    WHEN event_dim <= 270 THEN '241-270日'
    WHEN event_dim <= 300 THEN '271-300日'
    ELSE '300日超'
  END AS dim_range,
  COUNT(*) AS cnt
FROM event
WHERE event_number = 200
  AND event_date IS NOT NULL
  AND event_date >= :start
  AND event_date <= :end
  AND deleted = 0
GROUP BY dim_range
ORDER BY 
  CASE dim_range
    WHEN 'NULL' THEN 0
    WHEN '異常値' THEN 1
    WHEN '0-30日' THEN 2
    ...
  END
```

**パラメータ**: `start`, `end`

---

### 3. 分析モード時のAI制約を強化

分析モード用プロンプトに以下を追加：

```
【SQLテンプレート使用ルール（最重要）】

原則として SQL は以下のテンプレートから選択すること。

利用可能なSQLテンプレート:
  - calving_by_month_lact: 月別×産次別 分娩頭数（event_lactを使用）
  - insemination_by_month_lact: 月別×産次別 授精頭数（AI+ET、event_lactを使用）
  - conception_rate_by_month_lact: 月別×産次別 受胎率（event_lactを使用、outcomeで判定）
  - calving_lact2_only: 2産のみの分娩頭数（月別、event_lact=2で抽出）
  - dim_distribution: DIM分布（event_dimを使用、AIイベントのみ）

テンプレートを使用する場合は、以下の形式で宣言すること：

【使用テンプレート】
テンプレート名: <template_name>
使用理由: <理由>
パラメータ:
  - start: <開始日 YYYY-MM-DD>
  - end: <終了日 YYYY-MM-DD>

新規SQLを生成する場合は、「既存テンプレートでは対応できない理由」を明示すること。
```

### 4. 分析フローの固定

分析モード実行時の処理フロー：

1. **AIが宣言**:
   - 使用するテンプレート名
   - 使用理由
   - パラメータ（start, end 等）

2. **内部でテンプレートSQLを展開・実行**:
   - `_extract_template_info_from_response()` でテンプレート情報を抽出
   - `sql_templates.expand_template()` でSQLを展開
   - `_execute_sql_safely()` で実行

3. **結果をAIに渡す**:
   - SQL実行結果をフォーマット

4. **AIは出力**:
   - SQL（テンプレート名付き）
   - 結果
   - 結論

### 5. 実装メソッド

**新規追加**（`analysis_sql_templates.py`）:
- `AnalysisSQLTemplates` クラス
- `get_template()`: テンプレート取得
- `list_templates()`: テンプレート一覧取得
- `expand_template()`: テンプレート展開
- `get_template_list_for_ai()`: AI用一覧取得

**修正**（`main_window.py`）:
- `_get_analysis_mode_system_prompt()`: テンプレート一覧を追加
- `_extract_template_info_from_response()`: テンプレート情報を抽出（新規）
- `_extract_sql_from_response()`: テンプレート優先でSQL抽出
- `_handle_analysis_mode_result()`: テンプレート名を表示

## 安全性

- ✅ SQLインジェクション対策（日付形式のみ許可）
- ✅ パラメータ置換の安全性チェック
- ✅ テンプレートが見つからない場合の後方互換性（直接SQL抽出）

## 完了条件

- ✅ 分析モードのSQLがテンプレート中心になる
- ✅ 集計定義がコードとして固定される
- ✅ 設計者の意図がSQLとして残る

## 使用例

### 例1: テンプレートを使用

**ユーザー入力**: 「分析：2024年の月別×産次別分娩頭数を教えて」

**AI応答**:
```
【使用テンプレート】
テンプレート名: calving_by_month_lact
使用理由: 月別×産次別分娩頭数の集計に最適
パラメータ:
  - start: 2024-01-01
  - end: 2024-12-31
```

**システム処理**:
1. テンプレート `calving_by_month_lact` を展開
2. SQLを実行
3. 結果を表示

### 例2: テンプレート一覧の確認

AIは常に利用可能なテンプレート一覧を把握しており、適切なテンプレートを選択します。

## 今後の拡張

- 新しいテンプレートの追加（`TEMPLATES` 辞書に追加するだけ）
- パラメータの拡張（必要に応じて）


















