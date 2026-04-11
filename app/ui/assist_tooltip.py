"""
FALCON2 - アシスタントツールチップ

動作仕様:
  - ウィジェットにカーソルが入ると 700ms 後にポップアップ表示
  - ウィジェット または ポップアップ上にカーソルがある間は表示を維持
  - 両方からカーソルが離れると 300ms 後に自動クローズ
  - show_briefly() は指定時間後に自動クローズ（ステップ案内専用）
"""
import tkinter as tk
from typing import Optional

_assist_enabled: bool = True


def is_assist_enabled() -> bool:
    return _assist_enabled


def set_assist_enabled(value: bool):
    global _assist_enabled
    _assist_enabled = value


def dismiss_active_tooltip():
    """現在表示中のアシストポップアップを閉じる（タイプ切替・コンボ選択直後など）"""
    if FalconTooltip._active is not None:
        FalconTooltip._active._hide()


class FalconTooltip:
    """ウィジェットにホバーでガイドを表示する（カーソルが外れたら自動クローズ）"""

    _active: Optional["FalconTooltip"] = None

    HOVER_DELAY_MS = 700   # 表示開始までの遅延
    LEAVE_DELAY_MS = 300   # カーソルが外れてから閉じるまでの遅延

    def __init__(self, widget: tk.Widget, title: str, body: str, next_hint: str = ""):
        self.widget = widget
        self.title = title
        self.body = body
        self.next_hint = next_hint
        self._popup: Optional[tk.Toplevel] = None
        self._hover_after_id: Optional[str] = None
        self._leave_after_id: Optional[str] = None
        self._brief_after_id: Optional[str] = None
        self._brief_mode = False

        widget.bind("<Enter>", self._on_widget_enter, add="+")
        widget.bind("<Leave>", self._on_widget_leave, add="+")
        widget.bind("<FocusIn>", self._on_widget_focus_in, add="+")
        widget.bind("<FocusOut>", self._on_widget_focus_out, add="+")
        widget.bind("<Button-1>", self._on_widget_button1, add="+")

    # ── タイマー管理 ──────────────────────────────────

    def _cancel_hover_timer(self):
        if self._hover_after_id is not None:
            try:
                self.widget.after_cancel(self._hover_after_id)
            except Exception:
                pass
            self._hover_after_id = None

    def _cancel_leave_timer(self):
        if self._leave_after_id is not None:
            try:
                self.widget.after_cancel(self._leave_after_id)
            except Exception:
                pass
            self._leave_after_id = None

    def _cancel_brief_timer(self):
        if self._brief_after_id is not None:
            try:
                self.widget.after_cancel(self._brief_after_id)
            except Exception:
                pass
            self._brief_after_id = None

    # ── ウィジェット側イベント ───────────────────────

    def _on_widget_enter(self, event):
        if not _assist_enabled:
            return
        self._cancel_leave_timer()
        if self._popup is None:
            self._cancel_hover_timer()
            self._hover_after_id = self.widget.after(
                self.HOVER_DELAY_MS, self._open_hover_popup
            )

    def _on_widget_leave(self, event):
        self._cancel_hover_timer()
        if self._popup is not None and not self._brief_mode:
            # ポップアップが出ているときは遅延クローズ（ポップアップ上に移動する場合に備えて）
            self._schedule_leave_close()

    def _on_widget_focus_in(self, event):
        """キーボードやクリックでフォーカスが入ったときもマウスEnterと同様に案内を出す"""
        if not _assist_enabled:
            return
        self._cancel_leave_timer()
        if self._popup is None:
            self._cancel_hover_timer()
            self._hover_after_id = self.widget.after(
                self.HOVER_DELAY_MS, self._open_hover_popup
            )

    def _on_widget_focus_out(self, event):
        """フォーカスが外れたときはマウスLeaveと同様に閉じる処理へ"""
        if self._brief_mode:
            return
        self._cancel_hover_timer()
        if self._popup is not None:
            self._schedule_leave_close()

    def _on_widget_button1(self, event):
        """クリックでも閉じる"""
        self._hide()

    # ── ポップアップ側イベント ──────────────────────

    def _on_popup_enter(self, event):
        """ポップアップ内にカーソルが入ったらクローズタイマーをキャンセル"""
        self._cancel_leave_timer()

    def _on_popup_leave(self, event):
        """ポップアップからカーソルが出たら遅延クローズ"""
        if not self._brief_mode:
            self._schedule_leave_close()

    # ── クローズスケジューリング ────────────────────

    def _schedule_leave_close(self):
        self._cancel_leave_timer()
        self._leave_after_id = self.widget.after(self.LEAVE_DELAY_MS, self._on_leave_timeout)

    def _on_leave_timeout(self):
        self._leave_after_id = None
        self._hide()

    # ── ポップアップ構築 ────────────────────────────

    def _open_hover_popup(self):
        self._hover_after_id = None
        self._brief_mode = False
        self._build_popup()

    def _build_popup(self):
        if not _assist_enabled:
            return
        try:
            if not self.widget.winfo_exists():
                return
        except Exception:
            return

        if FalconTooltip._active and FalconTooltip._active is not self:
            FalconTooltip._active._hide()
        FalconTooltip._active = self

        master = self.widget.winfo_toplevel()
        popup = tk.Toplevel(master)
        popup.overrideredirect(True)
        try:
            popup.attributes("-topmost", True)
        except Exception:
            pass
        popup.configure(bg="#253040")

        inner = tk.Frame(popup, bg="#ffffff", padx=14, pady=10)
        inner.pack(padx=1, pady=1)

        tk.Label(
            inner,
            text=self.title,
            font=("Meiryo UI", 9, "bold"),
            bg="#ffffff",
            fg="#1a237e",
        ).pack(anchor=tk.W)

        tk.Frame(inner, bg="#e0e7ef", height=1).pack(fill=tk.X, pady=(4, 7))

        if self.body:
            tk.Label(
                inner,
                text=self.body,
                font=("Meiryo UI", 9),
                bg="#ffffff",
                fg="#37474f",
                justify=tk.LEFT,
                wraplength=300,
            ).pack(anchor=tk.W)

        if self.next_hint:
            tk.Frame(inner, bg="#e8eaf6", height=1).pack(fill=tk.X, pady=(8, 5))
            tk.Label(
                inner,
                text=self.next_hint,
                font=("Meiryo UI", 8),
                bg="#ffffff",
                fg="#3949ab",
                justify=tk.LEFT,
            ).pack(anchor=tk.W)

        # ポップアップ自体へのホバーイベントを登録
        popup.bind("<Enter>", self._on_popup_enter, add="+")
        popup.bind("<Leave>", self._on_popup_leave, add="+")
        for child in self._iter_all_children(popup):
            child.bind("<Enter>", self._on_popup_enter, add="+")
            child.bind("<Leave>", self._on_popup_leave, add="+")

        popup.update_idletasks()
        try:
            wx = self.widget.winfo_rootx()
            wy = self.widget.winfo_rooty()
            wh = self.widget.winfo_height()
        except Exception:
            try:
                popup.destroy()
            except Exception:
                pass
            if FalconTooltip._active is self:
                FalconTooltip._active = None
            return

        pw = popup.winfo_width()
        ph = popup.winfo_height()
        sw = popup.winfo_screenwidth()
        sh = popup.winfo_screenheight()

        x = wx
        y = wy + wh + 1
        if x + pw > sw:
            x = sw - pw - 8
        if x < 0:
            x = 4
        if y + ph > sh:
            y = wy - ph - 4

        popup.geometry(f"+{x}+{y}")
        self._popup = popup

    @staticmethod
    def _iter_all_children(widget):
        """ウィジェットの全子孫を列挙する"""
        for child in widget.winfo_children():
            yield child
            yield from FalconTooltip._iter_all_children(child)

    def _hide(self):
        # 表示前のホバー待ちもキャンセル（閉じた直後に再オープンしない）
        self._cancel_hover_timer()
        self._cancel_leave_timer()
        self._cancel_brief_timer()
        if self._popup:
            try:
                self._popup.destroy()
            except Exception:
                pass
            self._popup = None
        if FalconTooltip._active is self:
            FalconTooltip._active = None

    def show_briefly(self, duration_ms: int = 3000):
        """タイプ選択後など、一定時間だけ案内を出す（時間で閉じる）"""
        if not _assist_enabled:
            return
        self._cancel_hover_timer()
        self._brief_mode = True
        self._build_popup()
        if self._popup:
            self._brief_after_id = self.widget.after(duration_ms, self._hide)


def attach(widget: tk.Widget, title: str, body: str, next_hint: str = "") -> FalconTooltip:
    return FalconTooltip(widget, title, body, next_hint)
