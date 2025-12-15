"""
FALCON2 - 辞書同期処理
event_dictionary.json / item_dictionary.json の同期を担当
設計書 第12章参照
"""

import json
import logging
import shutil
from pathlib import Path
from typing import Dict, Any, List, Tuple

# ロガーを設定
logger = logging.getLogger(__name__)


class DictionarySync:
    """辞書同期処理クラス"""
    
    def __init__(self, falcon_root: Path):
        """
        初期化
        
        Args:
            falcon_root: FALCONフォルダのルートパス（例: C:\FALCON）
        """
        self.falcon_root = Path(falcon_root)
        self.config_default_dir = self.falcon_root / "config_default"
        self.default_event_dict_path = self.config_default_dir / "event_dictionary.json"
        self.default_item_dict_path = self.config_default_dir / "item_dictionary.json"
    
    def sync_dictionaries(self, farm_path: Path) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """
        農場フォルダの辞書をテンプレートと同期
        
        Args:
            farm_path: 農場フォルダのパス（例: C:\FARMS\FarmA）
        
        Returns:
            (event_dictionary, item_dictionary) のタプル
        """
        farm_path = Path(farm_path)
        farm_event_dict_path = farm_path / "event_dictionary.json"
        farm_item_dict_path = farm_path / "item_dictionary.json"
        
        # event_dictionary.json の同期
        event_dict = self._sync_event_dictionary(farm_event_dict_path)
        
        # item_dictionary.json の同期
        item_dict = self._sync_item_dictionary(farm_item_dict_path)
        
        return event_dict, item_dict
    
    def _sync_event_dictionary(self, farm_dict_path: Path) -> Dict[str, Any]:
        """
        event_dictionary.json を同期
        
        Args:
            farm_dict_path: 農場側の event_dictionary.json のパス
        
        Returns:
            同期後の event_dictionary
        """
        # テンプレートを読み込む
        if not self.default_event_dict_path.exists():
            logger.warning(f"テンプレートが見つかりません: {self.default_event_dict_path}")
            return {}
        
        try:
            with open(self.default_event_dict_path, 'r', encoding='utf-8') as f:
                default_dict = json.load(f)
        except Exception as e:
            logger.error(f"テンプレート読み込みエラー: {e}")
            return {}
        
        # 農場側の辞書を読み込む（存在しない場合は空辞書）
        farm_dict = {}
        if farm_dict_path.exists():
            try:
                with open(farm_dict_path, 'r', encoding='utf-8') as f:
                    farm_dict = json.load(f)
                logger.info(f"Event dictionary loaded from: {farm_dict_path}")
            except Exception as e:
                logger.error(f"農場側辞書読み込みエラー: {e}")
                farm_dict = {}
        else:
            logger.info(f"Event dictionary not found, creating from template: {farm_dict_path}")
        
        # 同期前のイベント数
        before_count = len(farm_dict)
        
        # 不足している event_number を追加
        added_events: List[str] = []
        for event_number, event_data in default_dict.items():
            if event_number not in farm_dict:
                # テンプレートから完全にコピー（deprecatedも含む）
                farm_dict[event_number] = event_data.copy()
                added_events.append(event_number)
        
        # 同期後のイベント数
        after_count = len(farm_dict)
        
        # 変更があった場合は保存
        if added_events:
            # 農場フォルダが存在することを確認
            farm_dict_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                with open(farm_dict_path, 'w', encoding='utf-8') as f:
                    json.dump(farm_dict, f, ensure_ascii=False, indent=2)
                logger.info(f"Event dictionary saved to: {farm_dict_path}")
            except Exception as e:
                logger.error(f"Event dictionary保存エラー: {e}")
        
        # ログ出力
        logger.info(f"Event dictionary sync: before={before_count}, added={len(added_events)}, after={after_count}")
        if added_events:
            logger.info(f"Added events: {added_events}")
        
        return farm_dict
    
    def _sync_item_dictionary(self, farm_dict_path: Path) -> Dict[str, Any]:
        """
        item_dictionary.json を同期
        
        Args:
            farm_dict_path: 農場側の item_dictionary.json のパス
        
        Returns:
            同期後の item_dictionary
        """
        # テンプレートを読み込む
        if not self.default_item_dict_path.exists():
            logger.warning(f"テンプレートが見つかりません: {self.default_item_dict_path}")
            return {}
        
        try:
            with open(self.default_item_dict_path, 'r', encoding='utf-8') as f:
                default_dict = json.load(f)
        except Exception as e:
            logger.error(f"テンプレート読み込みエラー: {e}")
            return {}
        
        # 農場側の辞書を読み込む（存在しない場合は空辞書）
        farm_dict = {}
        if farm_dict_path.exists():
            try:
                with open(farm_dict_path, 'r', encoding='utf-8') as f:
                    farm_dict = json.load(f)
                logger.info(f"Item dictionary loaded from: {farm_dict_path}")
            except Exception as e:
                logger.error(f"農場側辞書読み込みエラー: {e}")
                farm_dict = {}
        else:
            logger.info(f"Item dictionary not found, creating from template: {farm_dict_path}")
        
        # 同期前の項目数
        before_count = len(farm_dict)
        
        # 不足している項目を追加
        added_items: List[str] = []
        for item_key, item_data in default_dict.items():
            if item_key not in farm_dict:
                # テンプレートから完全にコピー
                farm_dict[item_key] = item_data.copy()
                added_items.append(item_key)
        
        # 同期後の項目数
        after_count = len(farm_dict)
        
        # 変更があった場合は保存
        if added_items:
            # 農場フォルダが存在することを確認
            farm_dict_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                with open(farm_dict_path, 'w', encoding='utf-8') as f:
                    json.dump(farm_dict, f, ensure_ascii=False, indent=2)
                logger.info(f"Item dictionary saved to: {farm_dict_path}")
            except Exception as e:
                logger.error(f"Item dictionary保存エラー: {e}")
        
        # ログ出力
        logger.info(f"Item dictionary sync: before={before_count}, added={len(added_items)}, after={after_count}")
        if added_items:
            logger.info(f"Added items: {added_items}")
        
        return farm_dict



