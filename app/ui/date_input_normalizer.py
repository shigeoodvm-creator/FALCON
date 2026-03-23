"""
FALCON2 - 日付入力の全角→半角正規化（全ウィンドウ共通）
"""

import re
import unicodedata
import tkinter as tk
from tkinter import ttk


# 日付入力で使われる可能性が高い文字だけを対象にする
_DATE_LIKE_PATTERN = re.compile(r"^[0-9０-９/\-／－.,\s:：]+$")


def install_date_input_normalizer(root: tk.Misc) -> None:
    """Entry系ウィジェットのフォーカスアウト時に日付風入力を半角化する。"""

    def _on_focus_out(event: tk.Event) -> None:
        widget = event.widget
        if not isinstance(widget, (tk.Entry, ttk.Entry, ttk.Spinbox)):
            return

        try:
            raw = widget.get()
        except Exception:
            return

        if not raw:
            return
        if not _DATE_LIKE_PATTERN.fullmatch(raw):
            return

        normalized = unicodedata.normalize("NFKC", raw)
        if normalized == raw:
            return

        # 先頭末尾の空白は除去せず、入力体験を保ったまま文字種だけ統一する
        try:
            cursor_pos = widget.index(tk.INSERT)
        except Exception:
            cursor_pos = None

        widget.delete(0, tk.END)
        widget.insert(0, normalized)

        if cursor_pos is not None:
            try:
                widget.icursor(min(cursor_pos, len(normalized)))
            except Exception:
                pass

    # add='+' で既存バインドを壊さず共存させる
    root.bind_all("<FocusOut>", _on_focus_out, add="+")

