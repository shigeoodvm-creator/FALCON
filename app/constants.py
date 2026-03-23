"""
FALCON2 - アプリ全体で共有する定数・パス定義

ハードコードされたパスをここに集約することで、
環境変更時の修正箇所を1ファイルに限定する。
"""

from pathlib import Path

# ========== ルートパス ==========

# FALCONアプリ本体のルート（このファイルの親フォルダ）
FALCON_ROOT: Path = Path(__file__).parent.parent

# アプリ（appフォルダ）ルート
APP_ROOT: Path = Path(__file__).parent

# ========== 農場データ ==========

# 農場データが置かれるルートフォルダ
FARMS_ROOT: Path = Path("C:/FARMS")

# ========== 設定フォルダ ==========

# アプリ設定（openai.json、app_settings.json など）
APP_CONFIG_DIR: Path = FALCON_ROOT / "config"

# デフォルト辞書（item_dictionary.json、event_dictionary.json）
CONFIG_DEFAULT_DIR: Path = FALCON_ROOT / "config_default"
