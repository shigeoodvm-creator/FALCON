# Phase D - Step 2: 分析モード（DB限定）実装サマリー

## 概要

FALCON2に「分析モード（DB限定）」を追加しました。既存の通常モード（数値即答・集計専用）は一切変更していません。

**実装ファイル**: `app/ui/main_window.py`

## 実装内容

### 1. 分析モード用フラグの追加

- **内部状態**: `self.analysis_mode: bool = False`（デフォルトはFalse）
- **UI**: 分析モードチェックボックスを追加（コマンド入力欄の横）

### 2. 分析モードの有効化条件

以下のいずれかを満たした場合のみ `analysis_mode=True` とする：

- **UIで「分析モード」がON**: チェックボックスがONの場合
- **入力文の先頭が特定キーワードで始まる場合**:
  - 「分析：」
  - 「解説：」
  - 「理由：」

**注意**: 通常の日本語入力だけでは分析モードに入らない（明示的操作のみ）

### 3. 分析モード時のAIプロンプト

分析モード時のみ、以下の制約プロンプトをシステムメッセージとして追加：

```
---
【分析モード専用ルール】

あなたは FALCON2 の「分析AI」です。

・あなたは必ずデータベース（SQLite）から取得された結果のみを根拠にする
・仮定・一般論・経験則のみでの結論は禁止
・必ず最初に「使用するSQL（SELECT文のみ）」を明示する
・INSERT / UPDATE / DELETE / ALTER は一切禁止
・SQL結果が無い場合は「データ不足」と明示する
・event テーブルの event_lact / event_dim を唯一の産次・DIM定義として扱う
・cow.lact は使用禁止

【出力形式（必須）】
1. 使用したSQL（そのまま表示）
2. SQL結果（表 or CSV）
3. 結論（短く・DB結果に基づくもののみ）

「推測」「可能性」「一般論」は禁止
数値が出ない場合は「該当データなし」と明示
---
```

### 4. SQL実行インターフェースの分離

**実装メソッド**:
- `_extract_sql_from_response(response: str) -> Optional[str]`: AI応答からSQL文を抽出
- `_execute_sql_safely(sql: str) -> Optional[List[Dict[str, Any]]]`: SQLを安全に実行（SELECT文のみ）

**動作**:
- `analysis_mode=True` の場合のみ：
  - AIが生成したSELECT文を一旦ログ出力
  - 内部で実行
  - 実行結果（rows）をAIに渡す
- 通常モードではSQLをAIに生成させない

**安全性**:
- INSERT/UPDATE/DELETE/ALTER/DROP/CREATE/TRUNCATEを禁止
- SELECT文のみ許可
- SQLiteスレッド違反を回避（ワーカースレッド内で新しいDBHandlerを生成）

### 5. 分析モードの出力ルール

以下の順で必ず出力：

1. **使用したSQL**（そのまま表示）
   ```
   【使用したSQL】
   ```sql
   SELECT ...
   ```
   ```

2. **SQL結果**（表形式）
   ```
   【SQL結果】
   column1 | column2 | column3
   ---------------------------
   value1  | value2  | value3
   ```

3. **結論**（短く・DB結果に基づくもののみ）
   ```
   【結論】
   （AI応答からSQL部分を除いた部分）
   ```

**禁止事項**:
- 「推測」「可能性」「一般論」は禁止
- 数値が出ない場合は「該当データなし」と明示

## 実装メソッド一覧

### 新規追加メソッド

1. `_on_analysis_mode_toggle()`: 分析モードのトグル処理
2. `_check_analysis_mode_activation(raw_input: str) -> bool`: 分析モードの有効化条件をチェック
3. `_execute_analysis_mode(user_input: str)`: 分析モードの実行処理
4. `_run_analysis_mode_in_thread(user_input: str, job_id: int)`: バックグラウンドスレッドで分析モードを実行
5. `_extract_sql_from_response(response: str) -> Optional[str]`: AI応答からSQL文を抽出
6. `_execute_sql_safely(sql: str) -> Optional[List[Dict[str, Any]]]`: SQLを安全に実行
7. `_format_sql_result(result: List[Dict[str, Any]]) -> str`: SQL結果を表形式にフォーマット
8. `_handle_analysis_mode_result(...)`: 分析モードの結果を処理
9. `_handle_analysis_mode_error(job_id: int)`: 分析モードのエラーを処理
10. `_get_analysis_mode_system_prompt() -> str`: 分析モード時のシステムプロンプトを取得

### 修正メソッド

- `_on_command_execute()`: 分析モードのチェックを追加

## 動作フロー

### 通常モード（既存の動作）

1. ユーザーがコマンド入力
2. QueryRouterで解釈
3. 数値のみの場合：個体カードを開く
4. それ以外：何もしない

### 分析モード（新規）

1. ユーザーが「分析：」で始まる入力、または「分析モード」チェックボックスをON
2. `_check_analysis_mode_activation()` で分析モードを判定
3. `_execute_analysis_mode()` を実行
4. バックグラウンドスレッドでAIに問い合わせ
5. AI応答からSQL文を抽出
6. SQLを安全に実行（SELECT文のみ）
7. 結果をフォーマットして表示（SQL → 結果 → 結論）

## 安全性

- ✅ 通常モードの挙動が一切変わらない
- ✅ 分析モードは誤爆しない（明示的操作のみ）
- ✅ SQL実行はSELECT文のみ許可
- ✅ 危険なSQL文（INSERT/UPDATE/DELETE等）は実行禁止
- ✅ SQLiteスレッド違反を回避

## 完了条件

- ✅ 通常モードの挙動が一切変わらない
- ✅ 分析モード時のみ、SQL→結果→結論の流れが成立する
- ✅ 集計ロジックはPythonではなくSQLに集約される
- ✅ 分析モードは誤爆しない（明示的操作のみ）

## 使用例

### 例1: UIで分析モードをON

1. 「分析モード」チェックボックスをON
2. 「2024年の分娩頭数を教えて」と入力
3. AIがSQLを生成 → 実行 → 結果を表示

### 例2: キーワードで分析モードに入る

1. 「分析：2024年の分娩頭数を教えて」と入力
2. 自動的に分析モードに入る
3. AIがSQLを生成 → 実行 → 結果を表示

### 例3: 通常モード（既存動作）

1. 「101」と入力
2. 個体カードが開く（既存の動作）

## 注意事項

- 分析モードはChatGPT APIが必要
- SQL実行はSELECT文のみ許可
- event.event_lact / event.event_dim を唯一の事実として扱う
- cow.lact は使用禁止


















