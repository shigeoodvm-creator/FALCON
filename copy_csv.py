"""CSVファイルをコピーするスクリプト"""
import shutil
from pathlib import Path

# 元のファイルパス（read_fileで読み込めるパス）
source = Path(r"c:\Users\user\Downloads\検定日速報_1219102_202511.csv")
dest = Path(r"C:\FALCON\temp_milk_test.csv")

try:
    # バイナリモードでコピー
    with open(source, 'rb') as f_src:
        with open(dest, 'wb') as f_dst:
            f_dst.write(f_src.read())
    print(f"ファイルをコピーしました: {dest}")
    print(f"ファイルサイズ: {dest.stat().st_size} bytes")
except Exception as e:
    print(f"エラー: {e}")
    import traceback
    traceback.print_exc()









































