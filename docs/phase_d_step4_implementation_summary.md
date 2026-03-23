# Phase D - Step 4: 分析モードのCSV/Excel出力機能 実装サマリー

## 概要

分析モードのSQL実行結果を、CSVおよびExcel（.xlsx）として外部利用できる形で出力可能にしました。

**実装ファイル**: 
- `app/modules/analysis_exporter.py`（新規）
- `app/ui/main_window.py`（修正）

## 実装内容

### 1. 出力形式の指定方法

以下のいずれかで出力形式を指定可能：

#### UIチェックボックス
- 「CSV出力」チェックボックス
- 「Excel出力」チェックボックス
- 複数選択可能

#### 入力文の接頭辞
- 「CSV：」で始まる入力
- 「Excel：」で始まる入力
- 「CSV+Excel：」で始まる入力

**注意**: 指定がない場合は従来どおり画面表示のみ

### 2. 出力用ユーティリティ

**ファイル**: `app/modules/analysis_exporter.py`

**クラス**: `AnalysisResultExporter`

**実装メソッド**:

#### `export_to_csv(rows, columns, filepath)`
- SQL実行結果をCSVファイルに出力
- UTF-8 BOM付きで出力（Excelで正しく開ける）
- 列名・並び順・数値はSQL結果を完全に保持

#### `export_to_excel(rows, columns, filepath)`
- SQL実行結果をExcelファイル（.xlsx）に出力
- openpyxl を使用
- 1シート目にそのまま書き出し
- 書式装飾は行わない（純データ）

#### `generate_filename(template_name, start_date, end_date, extension)`
- 出力ファイル名を生成
- 形式: `<テンプレート名>_<開始日>_<終了日>.<csv|xlsx>`

### 3. ファイル保存ルール

**保存先**:
```
farms/<farm_name>/exports/analysis/
```

**ファイル名**:
```
<テンプレート名>_<開始日>_<終了日>.<csv|xlsx>
```

**例**:
- `calving_by_month_lact_2024-01-01_2024-12-31.csv`
- `calving_by_month_lact_2024-01-01_2024-12-31.xlsx`

### 4. 分析モード実行フローの拡張

分析モード時：

1. **SQLテンプレート選択**
2. **SQL実行**
3. **結果を内部保持**
4. **画面表示**（従来通り）
5. **出力指定があれば**:
   - CSV出力
   - Excel出力

**注意**: 出力処理は UI スレッドをブロックしない（同期的に実行）

### 5. 出力時のAI出力ルール

AIの出力に以下を追加：

```
【出力ファイル】
CSV: calving_by_month_lact_2024-01-01_2024-12-31.csv
保存先: farms/デモファーム/exports/analysis/
Excel: calving_by_month_lact_2024-01-01_2024-12-31.xlsx
保存先: farms/デモファーム/exports/analysis/
```

## 実装メソッド

### 新規追加（`analysis_exporter.py`）
- `AnalysisResultExporter` クラス
- `export_to_csv()`: CSV出力
- `export_to_excel()`: Excel出力
- `generate_filename()`: ファイル名生成

### 修正（`main_window.py`）
- `_check_export_format()`: 出力形式をチェック（新規）
- `_export_analysis_results()`: 分析結果をCSV/Excelに出力（新規）
- `_handle_analysis_mode_result()`: 出力処理を追加
- `_on_command_execute()`: 出力形式チェックを追加

### UI追加
- 「CSV出力」チェックボックス
- 「Excel出力」チェックボックス

## 安全性

- ✅ SQL実行結果をそのまま出力（加工禁止）
- ✅ 列名・並び順・数値はSQL結果を完全に保持
- ✅ 余計な列追加・変換は禁止
- ✅ 日付・数値は型を維持

## 完了条件

- ✅ 分析モード結果をCSV/Excelで即出力できる
- ✅ 出力内容がSQL結果と完全一致する
- ✅ Excelでそのままグラフ・加工に使える
- ✅ 通常モードの動作に影響がない

## 使用例

### 例1: UIチェックボックスで出力

1. 「分析モード」チェックボックスをON
2. 「CSV出力」チェックボックスをON
3. 「分析：2024年の月別×産次別分娩頭数」と入力
4. SQL実行後、CSVファイルが自動出力される

### 例2: 入力文の接頭辞で出力

1. 「分析：CSV+Excel：2024年の月別×産次別分娩頭数」と入力
2. SQL実行後、CSVとExcelファイルが自動出力される

### 例3: 出力ファイルの確認

出力されたファイルは以下の場所に保存される：
- `farms/<farm_name>/exports/analysis/calving_by_month_lact_2024-01-01_2024-12-31.csv`
- `farms/<farm_name>/exports/analysis/calving_by_month_lact_2024-01-01_2024-12-31.xlsx`

## 注意事項

- openpyxl が必要（Excel出力の場合）
- 出力処理は同期的に実行（大量データの場合は時間がかかる可能性）
- ファイル名に使用できない文字は自動的に処理される


















