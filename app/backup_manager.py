"""
FALCON2 - バックアップ管理
農場フォルダの定期バックアップ（間隔・保存先はバックアップ設定で指定）
"""

import json
import logging
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

BACKUP_SETTINGS_FILENAME = "backup_settings.json"


def get_backup_settings_path(farm_path: Path) -> Path:
    """農場フォルダ内のバックアップ設定ファイルパスを返す"""
    return Path(farm_path) / BACKUP_SETTINGS_FILENAME


def load_backup_settings(farm_path: Path) -> Dict[str, Any]:
    """
    バックアップ設定を読み込む
    
    Returns:
        interval_months: int (1,2,3,6,12), destination_path: str, last_backup_ym: str|None
    """
    path = get_backup_settings_path(farm_path)
    if not path.exists():
        return {"interval_months": 1, "destination_path": "", "last_backup_ym": None}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "interval_months": data.get("interval_months", 1),
            "destination_path": data.get("destination_path", ""),
            "last_backup_ym": data.get("last_backup_ym"),
        }
    except Exception as e:
        logging.warning(f"バックアップ設定の読み込みに失敗: {e}")
        return {"interval_months": 1, "destination_path": "", "last_backup_ym": None}


def save_backup_settings(farm_path: Path, interval_months: int, destination_path: str):
    """バックアップ設定を保存（last_backup_ym は更新しない）"""
    path = get_backup_settings_path(farm_path)
    data = load_backup_settings(farm_path)
    data["interval_months"] = interval_months
    data["destination_path"] = destination_path
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"バックアップ設定の保存に失敗: {e}")
        raise


def _next_backup_ym(last_ym: str, interval_months: int) -> str:
    """最終バックアップ月から interval_months 経過後の次回バックアップ月 (YYYYMM) を返す"""
    y = int(last_ym) // 100
    m = int(last_ym) % 100
    m += interval_months
    while m > 12:
        m -= 12
        y += 1
    return f"{y}{m:02d}"


def should_run_backup(farm_path: Path) -> bool:
    """
    今月（または今回の間隔）でバックアップを実行すべきか判定する
    設定が無効（保存先が空）の場合は False
    """
    settings = load_backup_settings(farm_path)
    dest = (settings.get("destination_path") or "").strip()
    if not dest:
        return False
    interval = settings.get("interval_months", 1)
    if interval < 1:
        return False
    current_ym = datetime.now().strftime("%Y%m")
    last_ym = settings.get("last_backup_ym")
    if last_ym is None:
        return True  # 未実施なら今月実行
    next_ym = _next_backup_ym(last_ym, interval)
    return current_ym >= next_ym


def run_backup(farm_path: Path) -> Optional[Path]:
    """
    農場フォルダを指定先に日付付きフォルダ名でコピーする
    例: 20250205デモファーム
    
    Returns:
        作成したバックアップ先フォルダの Path。失敗時は None
    """
    settings = load_backup_settings(farm_path)
    dest_base = (settings.get("destination_path") or "").strip()
    if not dest_base:
        logging.warning("バックアップ先が設定されていません")
        return None
    dest_base_path = Path(dest_base)
    if not dest_base_path.is_dir():
        try:
            dest_base_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            logging.error(f"バックアップ先フォルダの作成に失敗: {e}")
            return None
    date_str = datetime.now().strftime("%Y%m%d")
    farm_name = farm_path.name
    backup_folder_name = f"{date_str}{farm_name}"
    backup_dest = dest_base_path / backup_folder_name
    if backup_dest.exists():
        logging.warning(f"バックアップ先が既に存在します: {backup_dest}")
        return None
    try:
        shutil.copytree(farm_path, backup_dest)
        logging.info(f"バックアップ完了: {backup_dest}")
    except Exception as e:
        logging.error(f"バックアップに失敗: {e}")
        return None
    # 最終バックアップ月を更新
    path = get_backup_settings_path(farm_path)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["last_backup_ym"] = datetime.now().strftime("%Y%m")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.warning(f"last_backup_ym の更新に失敗: {e}")
    return backup_dest


def run_backup_if_due(farm_path: Path) -> Optional[Path]:
    """
    バックアップが実行対象なら実行し、結果のパスを返す。対象でなければ None
    """
    if not should_run_backup(farm_path):
        return None
    return run_backup(farm_path)
