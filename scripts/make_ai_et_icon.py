"""
矢印アイコン（水色背景）をアプリの青 #3949ab に調整して AI/ET 用アイコンを生成する。
使い方: python scripts/make_ai_et_icon.py [入力画像パス]
入力省略時は Cursor の既定パスを試す。
出力: assets/icons/ai_et_icon.png
"""
import sys
from pathlib import Path

# アプリのアクセント色
TARGET_R, TARGET_G, TARGET_B = 0x39, 0x49, 0xab  # #3949ab

def main():
    try:
        from PIL import Image
    except ImportError:
        print("Pillow が必要です: pip install Pillow")
        sys.exit(1)

    root = Path(__file__).resolve().parent.parent
    out_dir = root / "assets" / "icons"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "ai_et_icon.png"

    if len(sys.argv) >= 2:
        src = Path(sys.argv[1])
    else:
        src = root / "assets" / "icons" / "ai_et_icon_source.png"
        if not src.exists():
            src = Path(
                "C:/Users/user/.cursor/projects/c-FALCON/assets/"
                "c__Users_user_AppData_Roaming_Cursor_User_workspaceStorage_46886bcee3fcbb121a973e9b2035f05a_images_image-0272f134-7aeb-4ecf-a9f6-9c05a233173f.png"
            )

    if not src.exists():
        print(f"ファイルが見つかりません: {src}")
        print("使い方: python scripts/make_ai_et_icon.py <入力画像パス>")
        sys.exit(1)

    img = Image.open(src).convert("RGBA")
    w, h = img.size
    out = Image.new("RGBA", (w, h))
    for y in range(h):
        for x in range(w):
            r, g, b, a = img.getpixel((x, y))
            if a == 0:
                out.putpixel((x, y), (0, 0, 0, 0))
                continue
            # 白に近い部分はそのまま、青系は #3949ab に
            gray = (r + g + b) / 3
            if gray > 220 and max(r, g, b) - min(r, g, b) < 30:
                out.putpixel((x, y), (r, g, b, a))  # 白矢印は維持
            elif b > r and b > g:
                out.putpixel((x, y), (TARGET_R, TARGET_G, TARGET_B, a))
            else:
                out.putpixel((x, y), (r, g, b, a))

    size = 44
    out = out.resize((size, size), Image.Resampling.LANCZOS)
    out.save(out_path, "PNG")
    print(f"保存しました: {out_path}")

if __name__ == "__main__":
    main()
