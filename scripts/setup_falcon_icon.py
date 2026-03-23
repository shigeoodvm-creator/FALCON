"""
Medical+V アイコンを assets/icons にコピー・リサイズする。
使い方: python scripts/setup_falcon_icon.py [入力画像パス]
"""
import sys
from pathlib import Path

def main():
    try:
        from PIL import Image
    except ImportError:
        print("Pillow が必要です: pip install Pillow")
        sys.exit(1)

    root = Path(__file__).resolve().parent.parent
    out_dir = root / "assets" / "icons"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "falcon_med_v_icon.png"

    if len(sys.argv) >= 2:
        src = Path(sys.argv[1])
    else:
        src = Path(
            "C:/Users/user/.cursor/projects/c-FALCON/assets/"
            "c__Users_user_AppData_Roaming_Cursor_User_workspaceStorage_46886bcee3fcbb121a973e9b2035f05a_images_image-19984631-29b6-4c23-a494-a96b7418ddcd.png"
        )

    if not src.exists():
        print(f"ファイルが見つかりません: {src}")
        print("使い方: python scripts/setup_falcon_icon.py <入力画像パス>")
        sys.exit(1)

    img = Image.open(src).convert("RGBA")
    w, h = img.size
    # Vの黒をマイルドなダークグレーに（#5c5c5c 程度）
    mild_r, mild_g, mild_b = 0x5c, 0x5c, 0x5c
    out = Image.new("RGBA", (w, h))
    for y in range(h):
        for x in range(w):
            r, g, b, a = img.getpixel((x, y))
            if a == 0:
                out.putpixel((x, y), (r, g, b, a))
                continue
            # 黒に近い部分（V）をマイルドなグレーに
            if r < 80 and g < 80 and b < 80:
                out.putpixel((x, y), (mild_r, mild_g, mild_b, a))
            else:
                out.putpixel((x, y), (r, g, b, a))
    img = out
    size = 52
    img = img.resize((size, size), Image.Resampling.LANCZOS)
    img.save(out_path, "PNG")
    print(f"保存しました: {out_path}")

if __name__ == "__main__":
    main()
