"""
FALCON2 - スプラッシュスクリーン

起動中に表示するロゴ画面。
farm_selector が表示可能になったら close() で閉じる。
"""

import tkinter as tk
from pathlib import Path

_NAVY   = "#1e2a3a"
_ACCENT = "#5c9ce6"
_WHITE  = "#ffffff"
_SUBTEXT = "#8faec8"
_FONT   = "Meiryo UI"


class SplashScreen:
    """起動スプラッシュ（overrideredirect で枠なし表示）"""

    W, H = 360, 220

    def __init__(self, root: tk.Tk, version: str = ""):
        self._root = root
        self.win   = tk.Toplevel(root)
        self.win.overrideredirect(True)          # タイトルバー非表示
        self.win.configure(bg=_NAVY)
        self.win.attributes("-topmost", True)

        # 中央配置
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x  = (sw - self.W) // 2
        y  = (sh - self.H) // 2
        self.win.geometry(f"{self.W}x{self.H}+{x}+{y}")

        self._build(version)

    def _build(self, version: str):
        # アイコン画像（あれば）
        try:
            ico_path = Path(__file__).parent.parent / "resources" / "falcon2.png"
            if ico_path.exists():
                from PIL import Image, ImageTk
                img = Image.open(ico_path).resize((64, 64), Image.LANCZOS)
                self._photo = ImageTk.PhotoImage(img)
                tk.Label(self.win, image=self._photo, bg=_NAVY).pack(pady=(28, 0))
        except Exception:
            # Pillow なし or ファイルなし → テキストで代替
            tk.Label(self.win, text="🦅", font=(_FONT, 36),
                     bg=_NAVY, fg=_ACCENT).pack(pady=(24, 0))

        tk.Label(self.win, text="FALCON2",
                 font=(_FONT, 22, "bold"), bg=_NAVY, fg=_WHITE).pack()

        tk.Label(self.win, text="牛群管理システム",
                 font=(_FONT, 10), bg=_NAVY, fg=_SUBTEXT).pack()

        if version:
            tk.Label(self.win, text=f"v{version}",
                     font=(_FONT, 8), bg=_NAVY, fg=_SUBTEXT).pack(pady=(2, 0))

        # アクセントライン
        tk.Frame(self.win, bg=_ACCENT, height=2).pack(fill=tk.X, padx=40, pady=12)

        self._status_var = tk.StringVar(value="起動中...")
        tk.Label(self.win, textvariable=self._status_var,
                 font=(_FONT, 8), bg=_NAVY, fg=_SUBTEXT).pack()

    def set_status(self, msg: str):
        """ステータス文字列を更新"""
        try:
            self._status_var.set(msg)
            self.win.update_idletasks()
        except Exception:
            pass

    def close(self):
        try:
            self.win.destroy()
        except Exception:
            pass
