"""
FALCON2 - アプリ設定管理
アプリ全体の設定（フォント、フォントサイズなど）を管理
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional

from constants import APP_CONFIG_DIR

logger = logging.getLogger(__name__)

# グローバル設定ファイルのパス
APP_SETTINGS_DIR = APP_CONFIG_DIR
APP_SETTINGS_FILE = APP_SETTINGS_DIR / "app_settings.json"


class AppSettingsManager:
    """アプリ設定管理クラス"""
    
    def __init__(self):
        """初期化"""
        self._settings: Optional[Dict[str, Any]] = None
        self._load_settings()
    
    def _load_settings(self):
        """設定を読み込む"""
        if APP_SETTINGS_FILE.exists():
            try:
                with open(APP_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    self._settings = json.load(f)
                logger.info(f"アプリ設定を読み込みました: {APP_SETTINGS_FILE}")
            except Exception as e:
                logger.error(f"アプリ設定の読み込みエラー: {e}")
                self._settings = self._get_default_settings()
        else:
            self._settings = self._get_default_settings()
            self.save()
    
    def _get_default_settings(self) -> Dict[str, Any]:
        """デフォルト設定を取得"""
        return {
            "font_family": "MS Gothic",  # デフォルトフォント
            "font_size": 9,  # デフォルトフォントサイズ
            "selected_month_1": None,  # 任意月の乳検データ1の指定月（1-12、Noneは未指定）
            "selected_month_2": None,  # 任意月の乳検データ2の指定月（1-12、Noneは未指定）
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """設定値を取得"""
        if self._settings is None:
            self._load_settings()
        return self._settings.get(key, default)
    
    def set(self, key: str, value: Any):
        """設定値を設定"""
        if self._settings is None:
            self._load_settings()
        self._settings[key] = value
        self.save()
    
    def save(self):
        """設定を保存"""
        if self._settings is None:
            return
        
        # ディレクトリが存在しない場合は作成
        APP_SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(APP_SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
            logger.info(f"アプリ設定を保存しました: {APP_SETTINGS_FILE}")
        except Exception as e:
            logger.error(f"アプリ設定の保存エラー: {e}")
    
    def get_font_family(self) -> str:
        """フォントファミリーを取得（アプリで Meiryo UI に固定）"""
        return "Meiryo UI"
    
    def set_font_family(self, font_family: str):
        """フォントファミリーを設定"""
        self.set("font_family", font_family)
    
    def get_font_size(self) -> int:
        """フォントサイズを取得"""
        return self.get("font_size", 9)
    
    def set_font_size(self, font_size: int):
        """フォントサイズを設定"""
        self.set("font_size", font_size)
    
    def get_all_settings(self) -> Dict[str, Any]:
        """全設定を取得"""
        if self._settings is None:
            self._load_settings()
        return self._settings.copy()
    
    def get_selected_month_1(self) -> Optional[int]:
        """任意月の乳検データ1の指定月を取得"""
        month = self.get("selected_month_1")
        if month is not None:
            try:
                month = int(month)
                if 1 <= month <= 12:
                    return month
            except (ValueError, TypeError):
                pass
        return None
    
    def set_selected_month_1(self, month: Optional[int]):
        """任意月の乳検データ1の指定月を設定"""
        if month is not None and (month < 1 or month > 12):
            raise ValueError("月は1-12の範囲で指定してください")
        self.set("selected_month_1", month)
    
    def get_selected_month_2(self) -> Optional[int]:
        """任意月の乳検データ2の指定月を取得"""
        month = self.get("selected_month_2")
        if month is not None:
            try:
                month = int(month)
                if 1 <= month <= 12:
                    return month
            except (ValueError, TypeError):
                pass
        return None
    
    def set_selected_month_2(self, month: Optional[int]):
        """任意月の乳検データ2の指定月を設定"""
        if month is not None and (month < 1 or month > 12):
            raise ValueError("月は1-12の範囲で指定してください")
        self.set("selected_month_2", month)


# グローバルインスタンス
_app_settings_manager: Optional[AppSettingsManager] = None


def get_app_settings_manager() -> AppSettingsManager:
    """アプリ設定マネージャーのシングルトンインスタンスを取得"""
    global _app_settings_manager
    if _app_settings_manager is None:
        _app_settings_manager = AppSettingsManager()
    return _app_settings_manager


