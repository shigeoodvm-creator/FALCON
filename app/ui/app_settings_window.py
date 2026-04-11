"""
FALCON2 - アプリ設定ウィンドウ
フォント、フォントサイズなどのアプリ全体の設定を行う（辞書・設定ダイアログと同一デザイン）
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
from typing import Optional, Callable

from modules.app_settings_manager import get_app_settings_manager

logger = logging.getLogger(__name__)


class AppSettingsWindow:
    """アプリ設定ウィンドウ（辞書・設定と統一イメージ）"""
    
    _FONT = "Meiryo UI"
    _BG = "#f5f5f5"
    _TITLE_FG = "#263238"
    _SUBTITLE_FG = "#607d8b"
    _DESC_FG = "#78909c"
    _BTN_PRIMARY_BG = "#3949ab"
    _BTN_PRIMARY_FG = "#ffffff"
    _BTN_SECONDARY_BG = "#fafafa"
    _BTN_SECONDARY_FG = "#546e7a"
    _BTN_SECONDARY_BD = "#b0bec5"
    
    def __init__(self, parent: tk.Tk, on_settings_changed: Optional[Callable] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            on_settings_changed: 設定変更時のコールバック関数
        """
        self.parent = parent
        self.on_settings_changed = on_settings_changed
        self.settings_manager = get_app_settings_manager()
        
        self.window = tk.Toplevel(parent)
        self.window.title("アプリ設定")
        self.window.geometry("520x580")
        self.window.minsize(480, 500)
        self.window.configure(bg=self._BG)
        
        self._create_widgets()
        self._load_current_settings()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _build_header(self, icon_char: str, title_text: str, subtitle_text: str):
        """統一デザインのヘッダー（アイコン・タイトル・サブタイトル）"""
        header = tk.Frame(self.window, bg=self._BG, pady=20, padx=24)
        header.pack(fill=tk.X)
        tk.Label(header, text=icon_char, font=(self._FONT, 24), bg=self._BG, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=self._BG)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text=title_text, font=(self._FONT, 16, "bold"), bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)
        tk.Label(title_frame, text=subtitle_text, font=(self._FONT, 10), bg=self._BG, fg=self._SUBTITLE_FG).pack(anchor=tk.W)
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        self._build_header("\u2699\ufe0f", "アプリ設定", "フォントサイズなどのアプリ全体の設定")
        
        content = tk.Frame(self.window, bg=self._BG, padx=24, pady=8)
        content.pack(fill=tk.BOTH, expand=True)
        
        # フォント設定
        font_frame = tk.Frame(content, bg=self._BG, padx=16, pady=16, highlightbackground="#e0e7ef", highlightthickness=1)
        font_frame.pack(fill=tk.X, pady=(0, 12))
        
        tk.Label(font_frame, text="フォント設定", font=(self._FONT, 11, "bold"), bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)
        tk.Label(
            font_frame, text="フォントはアプリで Meiryo UI に統一されています。",
            font=(self._FONT, 9), bg=self._BG, fg=self._DESC_FG
        ).pack(anchor=tk.W, pady=(2, 10))
        
        size_row = tk.Frame(font_frame, bg=self._BG)
        size_row.pack(fill=tk.X)
        tk.Label(size_row, text="フォントサイズ:", font=(self._FONT, 10), bg=self._BG, fg=self._TITLE_FG).pack(side=tk.LEFT, padx=(0, 10))
        self.font_size_var = tk.StringVar()
        font_sizes = [str(i) for i in range(6, 21)]
        self.font_size_combo = ttk.Combobox(
            size_row, textvariable=self.font_size_var, values=font_sizes, width=10, state="readonly"
        )
        self.font_size_combo.pack(side=tk.LEFT)
        self.font_size_var.trace("w", lambda *args: self._update_preview())
        
        # プレビュー
        preview_frame = tk.Frame(content, bg=self._BG, padx=16, pady=16, highlightbackground="#e0e7ef", highlightthickness=1)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))
        
        tk.Label(preview_frame, text="プレビュー", font=(self._FONT, 11, "bold"), bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)
        self.preview_label = tk.Label(
            preview_frame,
            text="これはプレビューです。\n文字の大きさの設定を確認できます。",
            font=(self._FONT, 10),
            bg=self._BG,
            fg=self._TITLE_FG,
            anchor=tk.CENTER,
            justify=tk.CENTER,
        )
        self.preview_label.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        
        # ライセンス情報
        lic_frame = tk.Frame(content, bg=self._BG, padx=16, pady=12,
                             highlightbackground="#e0e7ef", highlightthickness=1)
        lic_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(lic_frame, text="ライセンス", font=(self._FONT, 11, "bold"),
                 bg=self._BG, fg=self._TITLE_FG).pack(anchor=tk.W)

        try:
            import builtins
            from modules.license_manager import LicenseStatus as _LS
            _lic = getattr(builtins, "_falcon_license", None)
            lic_text = _lic.summary() if _lic else "（情報なし）"
            is_activatable = _lic and _lic.status in (
                _LS.TRIAL_ACTIVE, _LS.TRIAL_EXPIRED, _LS.LICENSE_EXPIRED
            )
        except Exception:
            lic_text = "（取得失敗）"
            is_activatable = False

        tk.Label(lic_frame, text=lic_text, font=(self._FONT, 9),
                 bg=self._BG, fg=self._DESC_FG, wraplength=440, justify=tk.LEFT
                 ).pack(anchor=tk.W, pady=(4, 0))

        if is_activatable:
            def _open_activation():
                from ui.license_window import LicenseActivationWindow
                LicenseActivationWindow(self.window, allow_close=True, info=_lic)
            tk.Button(
                lic_frame, text="ライセンスキーを入力する",
                font=(self._FONT, 9), relief=tk.FLAT,
                bg="#1e2a3a", fg="white", cursor="hand2", padx=12, pady=5,
                command=_open_activation
            ).pack(anchor=tk.W, pady=(8, 0))

        # フッター（ボタン）・余白を十分にとって切れないように
        footer = tk.Frame(self.window, bg=self._BG, pady=20)
        footer.pack(fill=tk.X)
        
        btn_frame = tk.Frame(footer, bg=self._BG)
        btn_frame.pack()
        
        tk.Button(
            btn_frame, text="保存", font=(self._FONT, 10),
            bg=self._BTN_PRIMARY_BG, fg=self._BTN_PRIMARY_FG,
            activebackground="#303f9f", activeforeground="#ffffff", relief=tk.FLAT,
            padx=20, pady=10, cursor="hand2", command=self._on_save
        ).pack(side=tk.LEFT, padx=6)
        tk.Button(
            btn_frame, text="キャンセル", font=(self._FONT, 10),
            bg=self._BTN_SECONDARY_BG, fg=self._BTN_SECONDARY_FG,
            activebackground="#eceff1", relief=tk.FLAT,
            padx=20, pady=10, highlightbackground=self._BTN_SECONDARY_BD, highlightthickness=1,
            cursor="hand2", command=self.window.destroy
        ).pack(side=tk.LEFT, padx=6)
        tk.Button(
            btn_frame, text="デフォルトに戻す", font=(self._FONT, 10),
            bg=self._BTN_SECONDARY_BG, fg=self._BTN_SECONDARY_FG,
            activebackground="#eceff1", relief=tk.FLAT,
            padx=20, pady=10, highlightbackground=self._BTN_SECONDARY_BD, highlightthickness=1,
            cursor="hand2", command=self._on_reset_defaults
        ).pack(side=tk.LEFT, padx=6)
    
    def _load_current_settings(self):
        """現在の設定を読み込んで表示"""
        font_size = self.settings_manager.get_font_size()
        self.font_size_var.set(str(font_size))
        self._update_preview()
    
    def _update_preview(self):
        """プレビューを更新（Meiryo UI + 選択中のサイズ）"""
        try:
            font_size = int(self.font_size_var.get())
            font = ("Meiryo UI", font_size)
            self.preview_label.config(font=font)
        except (ValueError, tk.TclError):
            pass
    
    def _on_save(self):
        """保存ボタンをクリック"""
        try:
            font_size = int(self.font_size_var.get())
            self.settings_manager.set_font_size(font_size)
            messagebox.showinfo("完了", "設定を保存しました。\nアプリを再起動すると反映されます。")
            if self.on_settings_changed:
                self.on_settings_changed()
            self.window.destroy()
        except ValueError:
            messagebox.showerror("エラー", "フォントサイズは数値で入力してください。")
        except Exception as e:
            logger.error(f"設定保存エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {e}")
    
    def _on_reset_defaults(self):
        """デフォルトに戻すボタンをクリック"""
        result = messagebox.askyesno(
            "確認",
            "フォントサイズをデフォルト（9）に戻しますか？"
        )
        if result:
            self.font_size_var.set("9")
            self._update_preview()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.focus_set()
