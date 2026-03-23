# AI/ETイベントの[P][O][R]保存場所確認結果

## 確認日時
2025-12-24

## 確認内容
AI/ETイベントに表示されている[P][O][R]が、event.json_dataに保存されているのか、それともUI表示専用ロジックなのかを特定

## 確認結果

### 1. eventテーブルの構造
- **カラム一覧**:
  - id (INTEGER)
  - cow_auto_id (INTEGER)
  - event_number (INTEGER)
  - event_date (TEXT)
  - json_data (TEXT)
  - note (TEXT)
  - deleted (INTEGER)
  - event_lact (INTEGER)
  - event_dim (INTEGER)

- **outcome/result/pregnant/P/O/Rに関連するカラム**: **なし**

### 2. AI/ETイベントのjson_data実データ確認

**確認した20件のAI/ETイベントのjson_data内容**:

すべてのAI/ETイベントのjson_dataには以下のキーのみが存在：
- `sire`: 種雄牛ID
- `technician_code`: 授精師コード
- `insemination_type_code`: 授精種類コード

**outcome/result/pregnant/P/O/Rに関連するキー**: **見つかりませんでした**

### 3. UI表示ロジックの確認

**実装場所**: `app/modules/formula_engine.py` の `get_ai_conception_status()` メソッド

**判定ロジック**:
1. **P（受胎）**: 後続の妊娠イベント（PDP/PDP2/PAGP）から受胎したAI/ETイベントを特定
2. **R（一連のAI）**: 一週間以内に次のAI/ETイベントが入力された場合
3. **O（受胎なし）**: 妊娠マイナス（PDN）、PAGマイナス（PAGN）、または再度AI/ETが入力された場合
4. **N（受胎後のAI）**: PとなったAIイベント後のAI

**表示処理**: `app/ui/cow_card.py` の1746行目で `formula_engine.get_ai_conception_status()` を呼び出し、結果を `format_insemination_event()` に渡して表示

## 結論

**[P][O][R]はUI表示専用ロジックです。**

- `event.json_data`には保存されていない
- `event`テーブルにも関連カラムは存在しない
- 後続の妊娠鑑定イベント（PDP/PDP2/PAGP/PDN/PAGN）から動的に計算されている
- 計算ロジックは `FormulaEngine.get_ai_conception_status()` に実装されている

## 補足

一部のSQLテンプレート（`analysis_sql_templates.py`、`aggregation_service.py`）では、`json_data`に`outcome`キーが存在することを前提としたSQLが書かれていますが、実際のデータには存在しません。これらのSQLは現在正しく動作していない可能性があります。

















