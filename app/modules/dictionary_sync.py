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

from constants import FARMS_ROOT, CONFIG_DEFAULT_DIR

# ロガーを設定
logger = logging.getLogger(__name__)


class DictionarySync:
    """辞書同期処理クラス"""

    def __init__(self):
        """初期化（辞書パスは constants.CONFIG_DEFAULT_DIR から取得）"""
        self.config_default_dir = CONFIG_DEFAULT_DIR
        self.default_event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
        self.default_item_dict_path  = CONFIG_DEFAULT_DIR / "item_dictionary.json"

    def delete_farm_dictionary_files(self, farms_root: Path = None) -> Dict[str, Any]:
        """
        全農場フォルダから項目辞書・イベント辞書を削除する。
        辞書は本体（config_default）を参照するため、農場側のファイルは不要。
        """
        if farms_root is None:
            farms_root = FARMS_ROOT
        farms_root = Path(farms_root)
        if not farms_root.exists():
            logger.warning(f"農場フォルダが見つかりません: {farms_root}")
            return {'error': f'農場フォルダが見つかりません: {farms_root}'}
        deleted = []
        for farm_dir in farms_root.iterdir():
            if not farm_dir.is_dir():
                continue
            for name in ("item_dictionary.json", "event_dictionary.json"):
                path = farm_dir / name
                if path.exists():
                    try:
                        path.unlink()
                        deleted.append(str(path))
                        logger.info(f"農場側の辞書を削除: {path}")
                    except Exception as e:
                        logger.warning(f"削除をスキップ: {path} - {e}")
        return {'deleted': deleted}
    
    def sync_dictionaries(self, farm_path: Path) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """
        農場フォルダの辞書をテンプレートと同期
        
        【重要】同期対象は event_dictionary.json と item_dictionary.json のみ。
        農場設定に関するJSON（farm_settings.json、insemination_settings.json など）は
        農場特有のものなので、農場横断的に更新しない。
        
        Args:
            farm_path: 農場フォルダのパス（例: C:\\FARMS\\FarmA）
        
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
        
        # 不足している event_number を追加、既存のイベントを更新
        added_events: List[str] = []
        updated_events: List[str] = []
        
        for event_number, default_event_data in default_dict.items():
            if event_number not in farm_dict:
                # 新規イベント：テンプレートから完全にコピー
                farm_dict[event_number] = default_event_data.copy()
                added_events.append(event_number)
            else:
                # 既存イベント：マスターの定義で更新（マージ）
                farm_event_data = farm_dict[event_number]
                # マスターの定義で更新（ただし、カスタム設定は保持）
                merged_data, changed = self._merge_event_data(farm_event_data, default_event_data)
                if changed:
                    farm_dict[event_number] = merged_data
                    updated_events.append(event_number)
        
        # 同期後のイベント数
        after_count = len(farm_dict)
        
        # 項目・イベント辞書は本体参照のため農場へは保存しない
        
        # ログ出力
        logger.info(f"Event dictionary sync: before={before_count}, added={len(added_events)}, updated={len(updated_events)}, after={after_count}")
        if added_events:
            logger.info(f"Added events: {added_events}")
        if updated_events:
            logger.info(f"Updated events: {updated_events}")
        
        return farm_dict
    
    def _merge_event_data(self, farm_event: Dict[str, Any], default_event: Dict[str, Any]) -> Tuple[Dict[str, Any], bool]:
        """
        イベントデータをマージ（マスターの定義で更新、カスタム設定は保持）
        
        Args:
            farm_event: 農場側のイベントデータ
            default_event: マスターのイベントデータ
        
        Returns:
            (マージ後のイベントデータ, 変更があったかどうか) のタプル
        """
        merged = farm_event.copy()
        changed = False
        
        # マスターの定義で更新（基本的なフィールド）
        for key in ['alias', 'name_jp', 'category', 'input_code', 'input_fields']:
            if key in default_event:
                default_value = default_event[key]
                farm_value = merged.get(key)
                
                # 値が異なる場合は更新
                if farm_value != default_value:
                    merged[key] = default_value
                    changed = True
        
        # deprecatedフラグも更新
        if 'deprecated' in default_event:
            if merged.get('deprecated') != default_event['deprecated']:
                merged['deprecated'] = default_event['deprecated']
                changed = True
        
        return merged, changed
    
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
        
        # 不足している項目を追加、既存の項目を更新
        added_items: List[str] = []
        updated_items: List[str] = []
        removed_items: List[str] = []
        
        # 重複項目を削除（1STTDIM、2NDTDIMは1STTDI、2NDTDIと重複）
        duplicate_keys = ['1STTDIM', '2NDTDIM']
        for dup_key in duplicate_keys:
            if dup_key in farm_dict:
                del farm_dict[dup_key]
                removed_items.append(dup_key)
        
        for item_key, default_item_data in default_dict.items():
            if item_key not in farm_dict:
                # 新規項目：テンプレートから完全にコピー
                farm_dict[item_key] = default_item_data.copy()
                added_items.append(item_key)
            else:
                # 既存項目：マスターの定義で更新（マージ）
                farm_item_data = farm_dict[item_key]
                # マスターの定義で更新（ただし、カスタム設定は保持）
                merged_data, changed = self._merge_item_data(farm_item_data, default_item_data)
                if changed:
                    farm_dict[item_key] = merged_data
                    updated_items.append(item_key)
        
        # 同期後の項目数
        after_count = len(farm_dict)
        
        # 項目・イベント辞書は本体参照のため農場へは保存しない
        
        # マスターの項目数
        master_count = len(default_dict)
        
        # ログ出力（指定されたフォーマット）
        logger.info(f"Item dictionary sync:")
        logger.info(f"  master={master_count}")
        logger.info(f"  before={before_count}")
        logger.info(f"  added={len(added_items)}")
        logger.info(f"  updated={len(updated_items)}")
        logger.info(f"  removed={len(removed_items)}")
        logger.info(f"  after={after_count}")
        
        if added_items:
            logger.debug(f"Added items: {added_items}")
        if updated_items:
            logger.debug(f"Updated items: {updated_items}")
        if removed_items:
            logger.info(f"Removed duplicate items: {removed_items}")
        
        return farm_dict
    
    def _merge_item_data(self, farm_item: Dict[str, Any], default_item: Dict[str, Any]) -> Tuple[Dict[str, Any], bool]:
        """
        項目データをマージ（マスターの定義で更新、カスタム設定は保持）
        
        Args:
            farm_item: 農場側の項目データ
            default_item: マスターの項目データ
        
        Returns:
            (マージ後の項目データ, 変更があったかどうか) のタプル
        """
        merged = farm_item.copy()
        changed = False
        
        # マスターの定義で更新（基本的なフィールド）
        for key in ['label', 'data_type', 'origin', 'formula', 'source', 'editable', 'display_order', 'display_name', 'description', 'category']:
            if key in default_item:
                default_value = default_item[key]
                farm_value = merged.get(key)
                
                # 値が異なる場合は更新
                if farm_value != default_value:
                    merged[key] = default_value
                    changed = True
        
        return merged, changed
    
    def sync_all_farms(self, farms_root: Path = None) -> Dict[str, Dict[str, int]]:
        """
        全農場の辞書を同期
        
        【重要】同期対象は event_dictionary.json と item_dictionary.json のみ。
        農場設定に関するJSON（farm_settings.json、insemination_settings.json など）は
        農場特有のものなので、農場横断的に更新しない。
        
        Args:
            farms_root: 農場フォルダのルートパス（例: C:\\FARMS）。Noneの場合はデフォルト
        
        Returns:
            各農場の同期結果の辞書 {farm_name: {'events_added': int, 'events_updated': int, 'items_added': int, 'items_updated': int}}
        """
        if farms_root is None:
            farms_root = FARMS_ROOT

        farms_root = Path(farms_root)

        if not farms_root.exists():
            logger.warning(f"農場フォルダが見つかりません: {farms_root}")
            return {}

        results = {}
        
        # 全農場フォルダを取得
        for farm_dir in farms_root.iterdir():
            if not farm_dir.is_dir():
                continue
            
            # farm.dbが存在するフォルダのみを対象
            db_path = farm_dir / "farm.db"
            if not db_path.exists():
                continue
            
            farm_name = farm_dir.name
            logger.info(f"農場 '{farm_name}' の辞書を同期中...")
            
            try:
                # 辞書を同期
                event_dict, item_dict = self.sync_dictionaries(farm_dir)
                
                # 結果を記録（簡易版：詳細はログに記録済み）
                results[farm_name] = {
                    'events_added': len(event_dict),
                    'events_updated': 0,  # 詳細はログで確認
                    'items_added': len(item_dict),
                    'items_updated': 0  # 詳細はログで確認
                }
                
                logger.info(f"農場 '{farm_name}' の辞書同期が完了しました")
            except Exception as e:
                logger.error(f"農場 '{farm_name}' の辞書同期でエラーが発生しました: {e}")
                results[farm_name] = {
                    'error': str(e)
                }
        
        logger.info(f"全農場の辞書同期が完了しました。対象農場数: {len(results)}")
        return results
    
    def sync_with_latest_json(self, farms_root: Path = None) -> Dict[str, Any]:
        """
        全農場間で最新のJSONファイルを優先して同期化
        
        全農場のevent_dictionary.jsonとitem_dictionary.jsonの更新日時を比較し、
        最新のファイルを全農場にコピーする。
        
        Args:
            farms_root: 農場フォルダのルートパス（例: C:\\FARMS）。Noneの場合はデフォルト
        
        Returns:
            同期結果の辞書
        """
        if farms_root is None:
            farms_root = FARMS_ROOT

        farms_root = Path(farms_root)

        if not farms_root.exists():
            logger.warning(f"農場フォルダが見つかりません: {farms_root}")
            return {'error': f'農場フォルダが見つかりません: {farms_root}'}
        
        # 全農場フォルダを取得
        farm_dirs = []
        for farm_dir in farms_root.iterdir():
            if not farm_dir.is_dir():
                continue
            
            # farm.dbが存在するフォルダのみを対象
            db_path = farm_dir / "farm.db"
            if not db_path.exists():
                continue
            
            farm_dirs.append(farm_dir)
        
        if not farm_dirs:
            logger.warning("同期対象の農場が見つかりません")
            return {'error': '同期対象の農場が見つかりません'}
        
        logger.info(f"最新JSON優先同期を開始します。対象農場数: {len(farm_dirs)}")
        
        # event_dictionary.jsonの最新ファイルを探す
        latest_event_dict_path = self._find_latest_json(farm_dirs, "event_dictionary.json")
        # item_dictionary.jsonの最新ファイルを探す
        latest_item_dict_path = self._find_latest_json(farm_dirs, "item_dictionary.json")
        
        results = {
            'event_dictionary': {},
            'item_dictionary': {}
        }
        
        # event_dictionary.jsonを同期
        if latest_event_dict_path:
            logger.info(f"最新のevent_dictionary.json: {latest_event_dict_path}")
            results['event_dictionary'] = self._sync_json_to_all_farms(
                latest_event_dict_path, farm_dirs, "event_dictionary.json"
            )
        else:
            logger.warning("event_dictionary.jsonが見つかりません")
            results['event_dictionary'] = {'error': 'event_dictionary.jsonが見つかりません'}
        
        # item_dictionary.jsonを同期
        if latest_item_dict_path:
            logger.info(f"最新のitem_dictionary.json: {latest_item_dict_path}")
            results['item_dictionary'] = self._sync_json_to_all_farms(
                latest_item_dict_path, farm_dirs, "item_dictionary.json"
            )
        else:
            logger.warning("item_dictionary.jsonが見つかりません")
            results['item_dictionary'] = {'error': 'item_dictionary.jsonが見つかりません'}
        
        logger.info("最新JSON優先同期が完了しました")
        return results
    
    def _find_latest_json(self, farm_dirs: List[Path], json_filename: str) -> Path:
        """
        全農場の中から最新のJSONファイルを探す
        
        Args:
            farm_dirs: 農場フォルダのリスト
            json_filename: JSONファイル名（例: "event_dictionary.json"）
        
        Returns:
            最新のJSONファイルのパス（見つからない場合はNone）
        """
        latest_path = None
        latest_mtime = 0
        
        for farm_dir in farm_dirs:
            json_path = farm_dir / json_filename
            if json_path.exists():
                try:
                    mtime = json_path.stat().st_mtime
                    if mtime > latest_mtime:
                        latest_mtime = mtime
                        latest_path = json_path
                except Exception as e:
                    logger.warning(f"ファイル情報取得エラー: {json_path}, {e}")
        
        return latest_path
    
    def _sync_json_to_all_farms(self, source_path: Path, farm_dirs: List[Path], json_filename: str) -> Dict[str, Any]:
        """
        最新のJSONファイルを全農場にコピー
        
        Args:
            source_path: コピー元のJSONファイルのパス
            farm_dirs: 農場フォルダのリスト
            json_filename: JSONファイル名（例: "event_dictionary.json"）
        
        Returns:
            同期結果の辞書
        """
        if not source_path.exists():
            return {'error': f'コピー元ファイルが見つかりません: {source_path}'}
        
        results = {
            'source': str(source_path),
            'copied_to': [],
            'skipped': [],
            'errors': []
        }
        
        # ソースファイルを読み込む
        try:
            with open(source_path, 'r', encoding='utf-8') as f:
                source_data = json.load(f)
        except Exception as e:
            logger.error(f"ソースファイル読み込みエラー: {e}")
            results['errors'].append(f'ソースファイル読み込みエラー: {e}')
            return results
        
        # 各農場にコピー
        for farm_dir in farm_dirs:
            target_path = farm_dir / json_filename
            
            # ソースと同じファイルの場合はスキップ
            if target_path == source_path:
                results['skipped'].append(str(farm_dir.name))
                continue
            
            try:
                # 農場フォルダが存在することを確認
                farm_dir.mkdir(parents=True, exist_ok=True)
                
                # JSONファイルをコピー
                with open(target_path, 'w', encoding='utf-8') as f:
                    json.dump(source_data, f, ensure_ascii=False, indent=2)
                
                results['copied_to'].append(str(farm_dir.name))
                logger.info(f"{json_filename} を {farm_dir.name} にコピーしました")
            except Exception as e:
                error_msg = f"{farm_dir.name}: {e}"
                results['errors'].append(error_msg)
                logger.error(f"{json_filename} のコピーエラー ({farm_dir.name}): {e}")
        
        logger.info(f"{json_filename} 同期完了: コピー先={len(results['copied_to'])}, スキップ={len(results['skipped'])}, エラー={len(results['errors'])}")
        return results






