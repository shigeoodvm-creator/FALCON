"""
FALCON2 - アプリ全体で共有する定数・パス定義

ハードコードされたパスをここに集約することで、
環境変更時の修正箇所を1ファイルに限定する。
"""

import sys
import json
from pathlib import Path

# ========== バージョン ==========

APP_VERSION: str = "1.0.0"

# ========== ルートパス ==========

if getattr(sys, "frozen", False):
    # PyInstaller EXE 化時:
    #   sys.executable = EXE ファイルのパス（書き込み可能な場所）
    #   sys._MEIPASS   = バンドルファイルの展開先（読み取り専用）
    FALCON_ROOT: Path = Path(sys.executable).parent
    APP_ROOT: Path    = Path(sys._MEIPASS)
else:
    # 通常の開発実行時
    FALCON_ROOT: Path = Path(__file__).parent.parent
    APP_ROOT: Path    = Path(__file__).parent

# ========== 設定フォルダ ==========

# アプリ設定（app_settings.json など）― EXE 隣の config/ フォルダに書き込む
APP_CONFIG_DIR: Path = FALCON_ROOT / "config"

if getattr(sys, "frozen", False):
    # EXE モード: PyInstaller 6.x は _internal/ にバンドルするため APP_ROOT を使う
    # CONFIG_DEFAULT_DIR = dist/FALCON2/_internal/config_default
    CONFIG_DEFAULT_DIR: Path = APP_ROOT / "config_default"
    NORMALIZATION_DIR: Path  = APP_ROOT / "normalization"
else:
    # 開発モード: プロジェクトルート直下
    # CONFIG_DEFAULT_DIR = C:/FALCON/config_default
    CONFIG_DEFAULT_DIR: Path = FALCON_ROOT / "config_default"
    NORMALIZATION_DIR: Path  = FALCON_ROOT / "normalization"

# ========== 農場データ（起動時にconfigから読み込む） ==========

_APP_CONFIG_PATH: Path = APP_CONFIG_DIR / "app_config.json"
_DEFAULT_FARMS_ROOT: Path = Path("C:/FARMS")


def _load_farms_root() -> Path:
    """app_config.json から FARMS_ROOT を読み込む。なければデフォルト値を返す。"""
    try:
        with open(_APP_CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
        p = cfg.get("farms_root")
        if p:
            return Path(p)
    except Exception:
        pass
    return _DEFAULT_FARMS_ROOT


def set_farms_root(new_path: Path) -> None:
    """
    FARMS_ROOT を変更して app_config.json に保存する。
    farm_selector など UI から呼び出す。
    変更後は constants.FARMS_ROOT が即時更新される。
    """
    global FARMS_ROOT
    FARMS_ROOT = Path(new_path)

    APP_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    cfg: dict = {}
    try:
        with open(_APP_CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        pass
    cfg["farms_root"] = str(new_path)
    with open(_APP_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# モジュールロード時に確定する（以後 from constants import FARMS_ROOT で取得した値は
# この時点の値。変更時は constants.set_farms_root() を使うこと）
FARMS_ROOT: Path = _load_farms_root()
