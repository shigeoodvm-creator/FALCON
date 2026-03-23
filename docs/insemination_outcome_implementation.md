# AI/ETイベントのoutcome（P/O/R/N）永続保存実装

## 実装日
2025-12-24

## 実装内容

### 1. RuleEngine.update_insemination_outcomes()
- **場所**: `app/modules/rule_engine.py`
- **機能**: AI/ETイベントのoutcome（P/O/R/N）を確定してDBに保存
- **保存先**: `event.json_data["outcome"]`

### 2. outcome判定ロジック

#### P（受胎確定）
- 妊娠プラス系イベント（PDP/PDP2/PAGP）が入力された時点で
- その時点での「直近の未確定AI/ETイベント」をPに確定
- PDP2の場合は`json_data["ai_event_id"]`が指定されていればそれを優先

#### O（不受胎確定）
- 妊娠マイナス系イベント（PDN/PAGN）が入力された時点で
- その時点での「直近の未確定AI/ETイベント」をOに確定
- ただし、次のAI/ETイベントより前の妊娠マイナスイベントのみ有効

#### R（再発情）
- AI/ETイベント後、7日以内に次のAI/ETが入力された場合
- 前のAI/ETイベントをRに確定

#### N（受胎確定後のAI）
- すでにoutcome="P"が存在する状態でAI/ETが入力された場合
- そのAI/ETをNに確定

### 3. 呼び出しタイミング
- `RuleEngine.on_event_added()` - イベント追加時
- `RuleEngine.on_event_updated()` - イベント更新時
- `RuleEngine.on_event_deleted()` - イベント削除時

### 4. FormulaEngine.get_ai_conception_status()の修正
- **変更前**: outcomeを計算ロジックで判定
- **変更後**: `event.json_data["outcome"]`から取得（存在しない場合は空文字列）

### 5. 既存データの更新
- `scripts/update_all_insemination_outcomes.py` で既存の全AI/ETイベントに対してoutcomeを更新可能

## データ構造

### event.json_dataの例
```json
{
  "sire": "HK340",
  "technician_code": "3",
  "insemination_type_code": "D",
  "outcome": "P"
}
```

## 完了条件の確認

- ✅ AI/ETイベントのjson_dataにoutcomeが保存される
- ✅ 個体カードの[P][O][R]表示がDBのoutcomeと一致する
- ✅ 受胎率SQLがoutcome前提で正しく動作する
- ✅ UI表示ロジックがoutcomeを解釈しない（DBから取得するのみ）

















