# 授精後日数（DAI）の仕様

## 目的

**その産次における授精後日数**を表す。直近分娩以降に授精がなければ意味がないため、分娩でリセットし、分娩後に新たなAI/ETがあればそこから日数を表示する。

## 挙動

| 状況 | 表示 |
|------|------|
| 直近のイベントが分娩で、その後にAI/ETがない | **空欄** |
| 直近分娩の後にAI/ETがある | その授精日から本日までの日数（ETの場合は7日差し引いた授精日から計算） |
| 分娩イベントを削除した | 再計算され、前産次の「直近分娩」以降のAI/ETが再度対象になり、授精後日数が**復活**する |

## 実装

- **計算**: `app/modules/formula_engine.py` の `_calculate_days_after_insemination(events)`
  - イベントから直近の分娩日(202)を取得
  - その分娩日**より後**のAI/ET(200, 201)のうち、もっとも新しい授精日から本日までの日数を返す
  - 該当するAI/ETがなければ `None`（空欄）
- **表示**: 個体カードの「項目」タブなどで `formula_engine.calculate()` の結果の `DAI` を表示
- **分娩削除時の復活**: イベント削除後に `rule_engine.on_event_deleted()` が呼ばれ、個体カードは `refresh()` → `load_cow()` で再描画される。その際に `formula_engine.calculate()` が再度実行され、削除後のイベントだけで DAI が再計算されるため、前産次の授精が再度「直近分娩後」とみなされ表示が復活する。

## 参照

- 項目定義: `config_default/item_dictionary.json` の `DAI`（`formula: "days_after_insemination(events)"`）
- 繁殖検診など「当該産次のみ」で授精後日数を扱う箇所: `app/modules/reproduction_checkup_logic.py` の `filter_events_current_lactation` および `_calculate_dai`
