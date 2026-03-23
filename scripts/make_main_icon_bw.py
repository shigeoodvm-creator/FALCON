"""
元画像を白黒にしてメイン画面用アイコンを生成する。
使い方: python scripts/make_main_icon_bw.py <入力画像パス>
出力: assets/icons/main_cow_icon_bw.png
"""
import sys
from pathlib import Path

def main():
    try:
        from PIL import Image
    except ImportError:
        print("Pillow が必要です: pip install Pillow")
        sys.exit(1)

    if len(sys.argv) < 2:
        print("使い方: python scripts/make_main_icon_bw.py <入力画像パス>")
        sys.exit(1)

    src = Path(sys.argv[1])
    if not src.exists():
        print(f"ファイルが見つかりません: {src}")
        sys.exit(1)

    root = Path(__file__).resolve().parent.parent
    out_dir = root / "assets" / "icons"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "main_cow_icon_bw.png"

    img = Image.open(src).convert("RGBA")
    w, h = img.size
    gray = img.convert("L")
    out = Image.new("RGBA", (w, h))
    for y in range(h):
        for x in range(w):
            r, g, b, a = img.getpixel((x, y))
            gv = gray.getpixel((x, y))
            v = 255 if gv > 140 else 0
            out.putpixel((x, y), (v, v, v, a))

    size = 32
    out = out.resize((size, size), Image.Resampling.LANCZOS)
    out.save(out_path, "PNG")
    print(f"保存しました: {out_path}")

if __name__ == "__main__":
    main()
