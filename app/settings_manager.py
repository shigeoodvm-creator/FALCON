"""
FALCON2 - 設定管理
農場ごとの farm_settings.json のロード/保存を担当
設計書 第8章・第19章参照
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional, List
from datetime import datetime


class SettingsManager:
    """農場設定管理クラス"""
    
    def __init__(self, farm_path: Path):
        """
        初期化
        
        Args:
            farm_path: 農場フォルダパス (例: C:/FARMS/FarmA)
        """
        self.farm_path = Path(farm_path)
        self.settings_file = self.farm_path / "farm_settings.json"
        self._settings: Optional[Dict[str, Any]] = None
    
    def load(self) -> Dict[str, Any]:
        """
        設定をロード
        
        Returns:
            設定辞書
        """
        if self._settings is not None:
            return self._settings
        
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self._settings = json.load(f)
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")
                self._settings = self._get_default_settings()
        else:
            self._settings = self._get_default_settings()
            self.save()
        
        return self._settings
    
    def save(self):
        """設定を保存"""
        if self._settings is None:
            return
        
        self.farm_path.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
    
    def _get_default_settings(self) -> Dict[str, Any]:
        """デフォルト設定を取得"""
        return {
            "farm_name": self.farm_path.name,
            "created_at": datetime.now().isoformat(),
            "db_version": "1.0",
            "repro_templates": [],  # 検診表テンプレート（設計書 第19章）
            "inseminator_codes": {  # 授精師コード辞書
                "1": "Sonoda",
                "2": "Tanaka"
            },
            "insemination_type_codes": {  # 授精種類コード辞書
                "1": "自然発情",
                "2": "同期化",
                "3": "ET"
            },
            "pen_settings": {}  # PENコード -> PEN名
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """設定値を取得"""
        settings = self.load()
        return settings.get(key, default)
    
    def set(self, key: str, value: Any):
        """設定値を設定"""
        settings = self.load()
        settings[key] = value
        self.save()
    
    # ========== 検診表テンプレート管理（設計書 第19章） ==========
    
    def get_repro_templates(self) -> List[Dict[str, Any]]:
        """検診表テンプレート一覧を取得"""
        return self.get("repro_templates", [])
    
    def add_repro_template(self, template: Dict[str, Any]) -> str:
        """
        検診表テンプレートを追加
        
        Args:
            template: テンプレート辞書（name, conditions など）
        
        Returns:
            テンプレートID
        """
        templates = self.get_repro_templates()
        
        # テンプレートIDを生成
        template_id = f"template_{len(templates) + 1}"
        template["id"] = template_id
        template["created_at"] = datetime.now().isoformat()
        
        templates.append(template)
        self.set("repro_templates", templates)
        
        return template_id
    
    def update_repro_template(self, template_id: str, template: Dict[str, Any]):
        """検診表テンプレートを更新"""
        templates = self.get_repro_templates()
        
        for i, t in enumerate(templates):
            if t.get("id") == template_id:
                template["id"] = template_id
                template["updated_at"] = datetime.now().isoformat()
                templates[i] = template
                self.set("repro_templates", templates)
                return
        
        raise ValueError(f"テンプレートが見つかりません: {template_id}")
    
    def delete_repro_template(self, template_id: str):
        """検診表テンプレートを削除"""
        templates = self.get_repro_templates()
        templates = [t for t in templates if t.get("id") != template_id]
        self.set("repro_templates", templates)
    
    def get_repro_template(self, template_id: str) -> Optional[Dict[str, Any]]:
        """検診表テンプレートを取得"""
        templates = self.get_repro_templates()
        for t in templates:
            if t.get("id") == template_id:
                return t
        return None
    
    # ========== 基本条件のデフォルト値 ==========
    
    @staticmethod
    def get_default_template_conditions() -> List[Dict[str, Any]]:
        """
        デフォルトの検診表テンプレート条件を取得
        設計書 第19章 E-1 参照
        """
        return [
            {
                "name": "VWP（日数）",
                "type": "vwp",
                "value": 60,
                "enabled": True
            },
            {
                "name": "妊娠鑑定日数",
                "type": "preg_check_days",
                "value": 35,
                "enabled": True
            },
            {
                "name": "再妊娠鑑定日数",
                "type": "recheck_days",
                "value": 60,
                "enabled": True
            },
            {
                "name": "育成牛月齢",
                "type": "heifer_age_months",
                "value": 13,
                "enabled": False
            },
            {
                "name": "育成牛妊娠鑑定日数",
                "type": "heifer_preg_check_days",
                "value": 35,
                "enabled": False
            },
            {
                "name": "育成牛再妊娠鑑定日数",
                "type": "heifer_recheck_days",
                "value": 60,
                "enabled": False
            }
        ]
    
    def create_default_template(self, name: str) -> str:
        """
        デフォルト条件を含むテンプレートを作成
        
        Args:
            name: テンプレート名
        
        Returns:
            テンプレートID
        """
        template = {
            "name": name,
            "conditions": self.get_default_template_conditions()
        }
        return self.add_repro_template(template)
    
    # ========== コード辞書管理（AI/ETイベント用） ==========
    
    def get_inseminator_codes(self) -> Dict[str, str]:
        """
        授精師コード辞書を取得
        
        Returns:
            コード辞書（例: {"1": "Sonoda", "2": "Tanaka"}）
        """
        return self.get("inseminator_codes", {})
    
    def set_inseminator_codes(self, codes: Dict[str, str]):
        """
        授精師コード辞書を設定
        
        Args:
            codes: コード辞書
        """
        self.set("inseminator_codes", codes)
    
    def get_insemination_type_codes(self) -> Dict[str, str]:
        """
        授精種類コード辞書を取得
        
        Returns:
            コード辞書（例: {"1": "自然発情", "2": "同期化"}）
        """
        return self.get("insemination_type_codes", {})
    
    def set_insemination_type_codes(self, codes: Dict[str, str]):
        """
        授精種類コード辞書を設定
        
        Args:
            codes: コード辞書
        """
        self.set("insemination_type_codes", codes)

    # ========== PEN 設定 ==========

    def load_pen_settings(self) -> Dict[str, str]:
        """
        PEN設定を取得

        Returns:
            {"1": "Lact1", "10": "Dry"} のような辞書
        """
        settings = self.load()
        pen_settings = settings.get("pen_settings")
        if pen_settings is None:
            pen_settings = {}
            settings["pen_settings"] = pen_settings
            self.save()
        return pen_settings

    def save_pen_settings(self, pen_settings: Dict[str, str]):
        """
        PEN設定を保存

        Args:
            pen_settings: コード文字列 -> 名称の辞書
        """
        settings = self.load()
        settings["pen_settings"] = pen_settings
        self.save()

