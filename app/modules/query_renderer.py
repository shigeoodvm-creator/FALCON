"""
FALCON2 - Query Renderer（クエリ結果のテキスト／表示名）
ExecutorV2 の結果辞書をテキスト化し、項目キーから表示名を解決する
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional, List


class QueryRenderer:
    """
    実行結果をテキスト化し、item_key → 表示名の解決を行う。
    """

    def __init__(
        self,
        item_dictionary_path: Optional[str] = None,
        event_dictionary_path: Optional[str] = None,
    ):
        self._item_dict: Dict[str, Any] = {}
        self._event_dict: Dict[str, Any] = {}
        if item_dictionary_path:
            path = Path(item_dictionary_path)
            if path.exists():
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        self._item_dict = json.load(f)
                except Exception:
                    pass
        if event_dictionary_path:
            path = Path(event_dictionary_path)
            if path.exists():
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        self._event_dict = json.load(f)
                except Exception:
                    pass

    def _get_item_display_name(self, item_key: str) -> str:
        """
        項目キーから表示名を返す。辞書に無い場合はキーをそのまま返す。
        """
        if not item_key:
            return ""
        entry = self._item_dict.get(item_key)
        if isinstance(entry, dict) and "display_name" in entry:
            return str(entry["display_name"])
        return item_key

    def render(self, result: Dict[str, Any], query_type: Optional[str] = None) -> str:
        """
        実行結果辞書をテキストに整形する。
        result: ExecutorV2 の戻り値 {"success", "data", "errors", "warnings"}
        query_type: "list", "agg", "eventcount", "graph", "scatter", "repro" など（省略可）
        """
        if not result:
            return "（結果がありません）"
        errors = result.get("errors") or []
        warnings = result.get("warnings") or []
        if not result.get("success"):
            lines = []
            if errors:
                lines.extend(errors)
            if warnings:
                lines.append("")
                lines.append("【警告】")
                lines.extend(warnings)
            return "\n".join(lines) if lines else "処理に失敗しました。"

        data = result.get("data")
        if not data:
            return "\n".join(errors) if errors else "（データなし）"

        kind = data.get("type")
        if kind == "table":
            return self._render_table(data)
        return self._format_generic(data, query_type)

    def _render_table(self, data: Dict[str, Any]) -> str:
        """テーブル形式の data をテキストに"""
        columns = data.get("columns") or []
        rows = data.get("rows") or []
        if not columns and not rows:
            return "（データなし）"
        # ヘッダーを表示名に
        headers = [self._get_item_display_name(c) for c in columns]
        lines = ["\t".join(headers)]
        for row in rows:
            if isinstance(row, dict):
                cells = [str(row.get(c, "")) for c in columns]
            else:
                cells = [str(x) for x in row]
            lines.append("\t".join(cells))
        return "\n".join(lines)

    def _format_generic(self, data: Dict[str, Any], query_type: Optional[str]) -> str:
        """その他の data を簡易テキスト化"""
        if query_type == "scatter" and "rows" in data:
            return self._render_table({"columns": data.get("columns", []), "rows": data["rows"]})
        return self._render_table(data) if "columns" in data and "rows" in data else str(data)


# モジュール外から _get_item_display_name を参照するため、インスタンスメソッドのまま公開
# main_window_scatter などは self.query_renderer._get_item_display_name(x_item_key) で利用
