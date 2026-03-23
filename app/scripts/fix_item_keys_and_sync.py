"""
項目キーを6文字以下に制限し、全農場の辞書を同期化するスクリプト
"""
import json
import shutil
from pathlib import Path
from typing import Dict, Any

# 項目キーのマッピング（旧キー → 新キー）
KEY_MAPPING = {
    "1STMILK": "1STMIL",      # 7文字 → 6文字
    "2NDMILK": "2NDMIL",      # 7文字 → 6文字
    "3RDMILK": "3RDMIL",      # 7文字 → 6文字
    "1STTDIM": "1STTDI",      # 7文字 → 6文字
    "2NDTDIM": "2NDTDI",      # 7文字 → 6文字
    "AGEFCLV": "AGEFCL",      # 7文字 → 6文字
    "CALVMONTH": "CALVMO",    # 9文字 → 6文字
    "PCALVMONTH": "PCALVM",   # 10文字 → 6文字
}

FALCON_ROOT = Path("C:/FALCON")
FARMS_ROOT = Path("C:/FARMS")
MASTER_ITEM_DICT = FALCON_ROOT / "docs" / "item_dictionary.json"
CONFIG_DEFAULT_ITEM_DICT = FALCON_ROOT / "config_default" / "item_dictionary.json"


def fix_item_dictionary_keys(item_dict: Dict[str, Any]) -> Dict[str, Any]:
    """
    項目辞書のキーを6文字以下に変更
    
    Args:
        item_dict: 項目辞書
        
    Returns:
        キーを変更した項目辞書
    """
    new_dict = {}
    
    for old_key, item_data in item_dict.items():
        if old_key in KEY_MAPPING:
            new_key = KEY_MAPPING[old_key]
            print(f"キーを変更: {old_key} → {new_key}")
            new_dict[new_key] = item_data
        else:
            new_dict[old_key] = item_data
    
    return new_dict


def sync_to_all_farms():
    """
    マスターのitem_dictionary.jsonを全農場に同期化
    """
    # マスター辞書のパスを決定
    if MASTER_ITEM_DICT.exists():
        master_path = MASTER_ITEM_DICT
    elif CONFIG_DEFAULT_ITEM_DICT.exists():
        master_path = CONFIG_DEFAULT_ITEM_DICT
    else:
        print(f"エラー: マスター辞書が見つかりません")
        return
    
    print(f"マスター辞書: {master_path}")
    
    # マスター辞書を読み込む
    try:
        with open(master_path, 'r', encoding='utf-8') as f:
            master_dict = json.load(f)
        print(f"マスター辞書を読み込みました（項目数: {len(master_dict)}）")
    except Exception as e:
        print(f"エラー: マスター辞書の読み込みに失敗しました: {e}")
        return
    
    # 全農場フォルダを取得
    if not FARMS_ROOT.exists():
        print(f"エラー: 農場フォルダが見つかりません: {FARMS_ROOT}")
        return
    
    farm_dirs = []
    for farm_dir in FARMS_ROOT.iterdir():
        if not farm_dir.is_dir():
            continue
        
        # farm.dbが存在するフォルダのみを対象
        db_path = farm_dir / "farm.db"
        if db_path.exists():
            farm_dirs.append(farm_dir)
    
    if not farm_dirs:
        print("同期対象の農場が見つかりません")
        return
    
    print(f"\n同期対象農場数: {len(farm_dirs)}")
    
    # 各農場に同期
    for farm_dir in farm_dirs:
        farm_item_dict_path = farm_dir / "item_dictionary.json"
        
        try:
            # マスター辞書をコピー
            farm_dir.mkdir(parents=True, exist_ok=True)
            
            with open(farm_item_dict_path, 'w', encoding='utf-8') as f:
                json.dump(master_dict, f, ensure_ascii=False, indent=2)
            
            print(f"[OK] {farm_dir.name}: 同期完了")
        except Exception as e:
            print(f"[ERROR] {farm_dir.name}: エラー - {e}")
    
    print(f"\n全農場の辞書同期が完了しました")


def main():
    """メイン処理"""
    print("=" * 60)
    print("項目キー修正と辞書同期化スクリプト")
    print("=" * 60)
    
    # 1. マスター辞書のキーを修正
    print("\n[1] マスター辞書のキーを修正中...")
    
    master_path = MASTER_ITEM_DICT if MASTER_ITEM_DICT.exists() else CONFIG_DEFAULT_ITEM_DICT
    
    if not master_path.exists():
        print(f"エラー: マスター辞書が見つかりません: {master_path}")
        return
    
    # バックアップを作成
    backup_path = master_path.with_suffix('.json.backup')
    try:
        shutil.copy2(master_path, backup_path)
        print(f"バックアップを作成: {backup_path}")
    except Exception as e:
        print(f"警告: バックアップの作成に失敗しました: {e}")
    
    # 辞書を読み込む
    try:
        with open(master_path, 'r', encoding='utf-8') as f:
            item_dict = json.load(f)
        print(f"マスター辞書を読み込みました（項目数: {len(item_dict)}）")
    except Exception as e:
        print(f"エラー: マスター辞書の読み込みに失敗しました: {e}")
        return
    
    # キーを修正
    fixed_dict = fix_item_dictionary_keys(item_dict)
    
    # 6文字超過のキーが残っていないか確認
    keys_over_6 = [k for k in fixed_dict.keys() if len(k) > 6]
    if keys_over_6:
        print(f"警告: 6文字超過のキーが残っています: {keys_over_6}")
    else:
        print("[OK] 全てのキーが6文字以下になりました")
    
    # 保存
    try:
        with open(master_path, 'w', encoding='utf-8') as f:
            json.dump(fixed_dict, f, ensure_ascii=False, indent=2)
        print(f"[OK] マスター辞書を保存しました: {master_path}")
    except Exception as e:
        print(f"エラー: マスター辞書の保存に失敗しました: {e}")
        return
    
    # 2. 全農場に同期
    print("\n[2] 全農場の辞書を同期化中...")
    sync_to_all_farms()
    
    print("\n" + "=" * 60)
    print("処理が完了しました")
    print("=" * 60)


if __name__ == "__main__":
    main()

