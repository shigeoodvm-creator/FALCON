"""
FALCONロゴの色を反転する: 黒背景+白F ⇔ 白背景+藍色F
使い方:
  python scripts/invert_falcon_logo.py         … 白背景+藍F に反転
  python scripts/invert_falcon_logo.py --revert … 元に戻す（黒背景+白F）
"""
import sys
from pathlib import Path

# 周囲の藍色（アプリで使用しているインディゴ）
INDIGO = (57, 73, 171)  # #3949ab
WHITE = (255, 255, 255)


def main():
    try:
        from PIL import Image
    except ImportError:
        print("Pillow が必要です: pip install Pillow")
        sys.exit(1)

    revert = "--revert" in sys.argv

    root = Path(__file__).resolve().parent.parent
    logo_path = root / "assets" / "falcon_logo.png"
    if not logo_path.exists():
        print(f"ロゴが見つかりません: {logo_path}")
        sys.exit(1)

    img = Image.open(logo_path).convert("RGBA")
    w, h = img.size
    gray = img.convert("L")
    out = Image.new("RGBA", (w, h))

    for y in range(h):
        for x in range(w):
            r, g, b, a = img.getpixel((x, y))
            gv = gray.getpixel((x, y))
            if revert:
                # 元に戻す: 明るい部分→黒、暗い部分→白（黒背景+白F）
                v = 255 - gv
                out.putpixel((x, y), (v, v, v, a))
            else:
                t = gv / 255.0
                nr = int((1 - t) * WHITE[0] + t * INDIGO[0])
                ng = int((1 - t) * WHITE[1] + t * INDIGO[1])
                nb = int((1 - t) * WHITE[2] + t * INDIGO[2])
                out.putpixel((x, y), (nr, ng, nb, a))

    out.save(logo_path, "PNG")
    print(f"保存しました: {logo_path}" + (" (元に戻しました)" if revert else ""))


if __name__ == "__main__":
    main()
