"""
計算項目削除スクリプト
origin="calc" かつ item_key NOT IN ("DIM", "DAI", "DUE") の項目を削除
"""

import json
import sqlite3
from pathlib import Path
from typing import Dict, Any

# 削除対象外（残す項目）
KEEP_ITEMS = {"DIM", "DAI", "DUE"}


def should_delete_item(item_key: str, item_data: Dict[str, Any]) -> bool:
    """
    項目を削除すべきか判定
    
    Args:
        item_key: 項目キー
        item_data: 項目データ
        
    Returns:
        削除すべき場合はTrue
    """
    # 残す項目は削除しない
    if item_key in KEEP_ITEMS:
        return False
    
    # origin="calc" または type="calc" の項目を削除対象とする
    origin = item_data.get("origin")
    item_type = item_data.get("type")
    
    if origin == "calc" or item_type == "calc":
        return True
    
    return False


def remove_from_json(json_path: Path) -> int:
    """
    item_dictionary.jsonから削除対象項目を削除
    
    Args:
        json_path: item_dictionary.jsonのパス
        
    Returns:
        削除した項目数
    """
    if not json_path.exists():
        print(f"警告: {json_path} が見つかりません")
        return 0
    
    # JSONを読み込む
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 削除対象を特定
    to_delete = []
    for key, value in data.items():
        if should_delete_item(key, value):
            to_delete.append(key)
    
    # 削除実行
    deleted_count = 0
    for key in to_delete:
        del data[key]
        deleted_count += 1
        print(f"削除: {key}")
    
    # JSONを保存
    if deleted_count > 0:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"{json_path} から {deleted_count} 項目を削除しました")
    
    return deleted_count


def remove_from_db(db_path: Path) -> int:
    """
    DBのitem_dictionaryテーブルから削除対象項目を削除
    
    Args:
        db_path: farm.dbのパス
        
    Returns:
        削除した項目数（テーブルが存在しない場合は0）
    """
    if not db_path.exists():
        print(f"警告: {db_path} が見つかりません")
        return 0
    
    try:
        conn = sqlite3.connect(str(db_path))
        cursor = conn.cursor()
        
        # テーブルが存在するか確認
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='item_dictionary'
        """)
        
        if not cursor.fetchone():
            print(f"情報: {db_path} に item_dictionary テーブルは存在しません")
            conn.close()
            return 0
        
        # 削除対象を特定して削除
        # origin='calc' かつ item_key NOT IN ('DIM', 'DAI', 'DUE')
        cursor.execute("""
            DELETE FROM item_dictionary
            WHERE origin = 'calc'
              AND item_key NOT IN ('DIM', 'DAI', 'DUE')
        """)
        
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()
        
        if deleted_count > 0:
            print(f"{db_path} の item_dictionary テーブルから {deleted_count} 項目を削除しました")
        else:
            print(f"{db_path} の item_dictionary テーブルに削除対象項目はありませんでした")
        
        return deleted_count
        
    except Exception as e:
        print(f"エラー: {db_path} の処理中にエラーが発生しました: {e}")
        return 0


def main():
    """メイン処理"""
    print("=" * 60)
    print("計算項目削除スクリプト")
    print("=" * 60)
    print()
    
    total_deleted = 0
    
    # 1. docs/item_dictionary.json を処理
    docs_json = Path("docs/item_dictionary.json")
    if docs_json.exists():
        print(f"処理中: {docs_json}")
        deleted = remove_from_json(docs_json)
        total_deleted += deleted
        print()
    
    # 2. config_default/item_dictionary.json を処理
    config_json = Path("config_default/item_dictionary.json")
    if config_json.exists():
        print(f"処理中: {config_json}")
        deleted = remove_from_json(config_json)
        total_deleted += deleted
        print()
    
    # 3. 農場フォルダ内のitem_dictionary.jsonを処理
    farms_root = Path("C:/FARMS")
    if farms_root.exists():
        for farm_dir in farms_root.iterdir():
            if farm_dir.is_dir():
                farm_json = farm_dir / "item_dictionary.json"
                if farm_json.exists():
                    print(f"処理中: {farm_json}")
                    deleted = remove_from_json(farm_json)
                    total_deleted += deleted
                    print()
                
                # 4. 農場DBのitem_dictionaryテーブルを処理
                farm_db = farm_dir / "farm.db"
                if farm_db.exists():
                    print(f"処理中: {farm_db}")
                    deleted = remove_from_db(farm_db)
                    total_deleted += deleted
                    print()
    
    print("=" * 60)
    print(f"完了: 合計 {total_deleted} 項目を削除しました")
    print("=" * 60)
    print()
    print("残す項目:")
    for item in sorted(KEEP_ITEMS):
        print(f"  - {item}")
    print()
    print("アプリケーションを再起動してください。")


if __name__ == "__main__":
    main()






































