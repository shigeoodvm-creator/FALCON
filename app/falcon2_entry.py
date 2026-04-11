"""
FALCON2 - PyInstaller EXE エントリーポイント

このファイルは app/ 内に置くことで、同じ app/ にある main.py を
ルートの main.py（CLIラッパー）と衝突せずに import できる。
通常の開発時は root の main.py を直接実行してください。
"""

from main import main

if __name__ == "__main__":
    main()
