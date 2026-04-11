"""
FALCON2 - ライセンス画面

・試用期間中のバナー（残日数表示）
・試用期限切れ時のアクティベーション画面
・ライセンス更新画面
"""

import tkinter as tk
from tkinter import messagebox
from typing import Callable, Optional

from modules.license_manager import get_license_manager, LicenseInfo, LicenseStatus

_FONT    = "Meiryo UI"
_BG      = "#f5f6fa"
_NAVY    = "#1e2a3a"
_NAVY_H  = "#2d3e50"
_TEXT    = "#1c2333"
_TEXT_S  = "#6b7280"
_WARN    = "#e65100"
_ERR     = "#c62828"
_GREEN   = "#2e7d32"
_BORDER  = "#dde1e7"

SUPPORT_CONTACT = "support@falcon2.jp"   # ← 実際の連絡先に変更すること


# ─────────────────────────────────────────────
#  試用期間バナー（メインウィンドウ上部に差し込む）
# ─────────────────────────────────────────────

def build_trial_banner(parent: tk.Widget, info: LicenseInfo,
                       on_activate: Optional[Callable] = None) -> Optional[tk.Frame]:
    """
    試用期間中のみバナーを返す。試用終了 or ライセンス済みは None を返す。
    on_activate: 「ライセンスを購入・入力」ボタン押下時のコールバック
    """
    if info.status != LicenseStatus.TRIAL_ACTIVE:
        return None

    remaining = info.trial_remaining_days
    if remaining > 7:
        bg, fg = "#fff8e1", "#f57f17"   # 薄黄
    else:
        bg, fg = "#fce4ec", _ERR        # 薄赤（警告）

    banner = tk.Frame(parent, bg=bg, padx=12, pady=6)

    msg = f"試用版 ― あと {remaining} 日間ご利用いただけます。"
    tk.Label(banner, text=msg, font=(_FONT, 9), bg=bg, fg=fg).pack(side=tk.LEFT)

    if on_activate:
        btn = tk.Label(
            banner, text="ライセンスを入力する →",
            font=(_FONT, 9, "underline"), bg=bg, fg=_NAVY, cursor="hand2"
        )
        btn.pack(side=tk.LEFT, padx=(12, 0))
        btn.bind("<Button-1>", lambda _: on_activate())

    return banner


# ─────────────────────────────────────────────
#  アクティベーション画面（期限切れ or 手動で開く）
# ─────────────────────────────────────────────

class LicenseActivationWindow:
    """
    試用期限切れ時 or メニューから開くライセンス入力ウィンドウ。
    on_activated: アクティベーション成功時のコールバック（引数なし）
    allow_close: Falseにすると×ボタンで閉じられない（期限切れ強制表示）
    """

    def __init__(self, parent: tk.Tk,
                 on_activated: Optional[Callable] = None,
                 allow_close: bool = True,
                 info: Optional[LicenseInfo] = None):
        self.parent       = parent
        self.on_activated = on_activated
        self.allow_close  = allow_close
        self.info         = info

        self.window = tk.Toplevel(parent)
        self.window.title("FALCON2 ライセンス")
        self.window.geometry("520x420")
        self.window.resizable(False, False)
        self.window.configure(bg=_BG)

        if not allow_close:
            self.window.protocol("WM_DELETE_WINDOW", self._on_force_close)

        self._build()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth()  // 2) - (self.window.winfo_width()  // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        self.window.lift()
        self.window.focus_force()

    def _build(self):
        # ── ヘッダー ──
        header = tk.Frame(self.window, bg=_NAVY, height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="FALCON2", font=(_FONT, 18, "bold"),
                 bg=_NAVY, fg="white").place(relx=0.5, rely=0.35, anchor=tk.CENTER)
        tk.Label(header, text="牛群管理システム",
                 font=(_FONT, 9), bg=_NAVY, fg="#c5cae9").place(relx=0.5, rely=0.72, anchor=tk.CENTER)

        # ── メッセージ ──
        body = tk.Frame(self.window, bg=_BG, padx=32, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        if self.info and self.info.status == LicenseStatus.TRIAL_EXPIRED:
            tk.Label(body, text="試用期間が終了しました",
                     font=(_FONT, 13, "bold"), bg=_BG, fg=_ERR).pack(anchor=tk.W)
            tk.Label(body,
                     text="引き続きご利用いただくには、ライセンスキーの入力が必要です。\n"
                          "ライセンスのご購入・お問い合わせは下記までご連絡ください。",
                     font=(_FONT, 9), bg=_BG, fg=_TEXT_S,
                     justify=tk.LEFT).pack(anchor=tk.W, pady=(6, 16))
        elif self.info and self.info.status == LicenseStatus.LICENSE_EXPIRED:
            tk.Label(body, text="ライセンスの有効期限が切れました",
                     font=(_FONT, 13, "bold"), bg=_BG, fg=_WARN).pack(anchor=tk.W)
            tk.Label(body,
                     text=f"有効期限: {self.info.expiry_date}\n"
                          "更新ライセンスキーをご入力ください。",
                     font=(_FONT, 9), bg=_BG, fg=_TEXT_S,
                     justify=tk.LEFT).pack(anchor=tk.W, pady=(6, 16))
        else:
            tk.Label(body, text="ライセンスキーの入力",
                     font=(_FONT, 13, "bold"), bg=_BG, fg=_TEXT).pack(anchor=tk.W)
            tk.Label(body, text="購入済みのライセンスキーを入力してください。",
                     font=(_FONT, 9), bg=_BG, fg=_TEXT_S).pack(anchor=tk.W, pady=(4, 16))

        # ── 連絡先 ──
        contact_frame = tk.Frame(body, bg="#edf0f3", padx=14, pady=10,
                                 highlightthickness=1, highlightbackground=_BORDER)
        contact_frame.pack(fill=tk.X, pady=(0, 16))
        tk.Label(contact_frame, text="購入・お問い合わせ",
                 font=(_FONT, 9, "bold"), bg="#edf0f3", fg=_TEXT).pack(anchor=tk.W)
        tk.Label(contact_frame, text=f"メール: {SUPPORT_CONTACT}",
                 font=(_FONT, 9), bg="#edf0f3", fg=_TEXT_S).pack(anchor=tk.W, pady=(2, 0))
        tk.Label(contact_frame, text="LINE: @falcon2support（営業時間内に返信）",
                 font=(_FONT, 9), bg="#edf0f3", fg=_TEXT_S).pack(anchor=tk.W)

        # ── キー入力 ──
        tk.Label(body, text="ライセンスキー",
                 font=(_FONT, 10, "bold"), bg=_BG, fg=_TEXT).pack(anchor=tk.W)
        tk.Label(body, text="（例: FALCON-XXXXXXXXXXXX-XXXXXX）",
                 font=(_FONT, 8), bg=_BG, fg=_TEXT_S).pack(anchor=tk.W)

        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(
            body, textvariable=self.key_var,
            font=(_FONT, 10), relief=tk.FLAT,
            highlightthickness=1, highlightbackground=_BORDER,
            highlightcolor=_NAVY, bg="white", fg=_TEXT
        )
        self.key_entry.pack(fill=tk.X, pady=(4, 4), ipady=7)
        self.key_entry.focus_set()
        self.key_entry.bind("<Return>", lambda _: self._on_activate())

        self.msg_label = tk.Label(body, text="", font=(_FONT, 9),
                                  bg=_BG, fg=_ERR)
        self.msg_label.pack(anchor=tk.W)

        # ── ボタン ──
        btn_frame = tk.Frame(body, bg=_BG)
        btn_frame.pack(fill=tk.X, pady=(12, 0))

        if self.allow_close:
            tk.Button(
                btn_frame, text="キャンセル",
                font=(_FONT, 10), relief=tk.FLAT,
                bg="#eaecf0", fg=_TEXT,
                cursor="hand2", padx=16, pady=7,
                command=self.window.destroy
            ).pack(side=tk.RIGHT, padx=(8, 0))

        tk.Button(
            btn_frame, text="  有効化する  ",
            font=(_FONT, 10, "bold"), relief=tk.FLAT,
            bg=_NAVY, fg="white", activebackground=_NAVY_H,
            cursor="hand2", padx=16, pady=7,
            command=self._on_activate
        ).pack(side=tk.RIGHT)

    def _on_activate(self):
        key = self.key_var.get().strip()
        if not key:
            self.msg_label.config(text="ライセンスキーを入力してください")
            return

        manager = get_license_manager()
        ok, msg = manager.activate(key)

        if ok:
            messagebox.showinfo("完了", msg, parent=self.window)
            self.window.destroy()
            if self.on_activated:
                self.on_activated()
        else:
            self.msg_label.config(text=f"エラー: {msg}")

    def _on_force_close(self):
        """期限切れ時は×で閉じると終了確認"""
        if messagebox.askyesno("終了確認",
                "ライセンスが有効化されていません。\nアプリを終了しますか？",
                parent=self.window):
            self.parent.quit()

    def show(self):
        self.window.wait_window()
