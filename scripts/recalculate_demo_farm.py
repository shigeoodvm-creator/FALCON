"""
C:\\FARMS\\デモファームの全牛の状態を再計算するスクリプト
"""
import sys
from pathlib import Path
import logging

# app ディレクトリをパスに追加
APP_DIR = Path(__file__).parent.parent / "app"
sys.path.insert(0, str(APP_DIR))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    # パスを直接指定
    farm_path = Path(r"C:\FARMS\デモファーム")
    db_path = farm_path / "farm.db"
    
    if not db_path.exists():
        logger.error(f"farm.dbが見つかりません: {db_path}")
        return
    
    logger.info(f"処理中: {farm_path.name}")
    logger.info(f"DBパス: {db_path}")
    
    try:
        db = DBHandler(db_path)
        rule_engine = RuleEngine(db)
        
        # 全牛を取得
        all_cows = db.get_all_cows()
        logger.info(f"  牛数: {len(all_cows)}")
        
        updated_count = 0
        error_count = 0
        
        for cow in all_cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            try:
                # 牛の状態を再計算してcowテーブルを更新
                rule_engine._recalculate_and_update_cow(cow_auto_id)
                updated_count += 1
                
                # 進捗表示（100頭ごと）
                if updated_count % 100 == 0:
                    logger.info(f"  進捗: {updated_count}/{len(all_cows)}頭を処理")
                    
            except Exception as e:
                error_count += 1
                logger.error(f"  エラー (cow_auto_id={cow_auto_id}, cow_id={cow.get('cow_id')}): {e}", exc_info=True)
        
        logger.info(f"  完了: {updated_count}頭を更新, {error_count}件のエラー")
        db.close()
        
        print()
        print("=" * 80)
        print(f"更新完了: {updated_count}頭を更新, {error_count}件のエラー")
        print("=" * 80)
        
    except Exception as e:
        logger.error(f"  エラー: {e}", exc_info=True)
        print()
        print("=" * 80)
        print("エラーが発生しました")
        print("=" * 80)

if __name__ == "__main__":
    print("=" * 80)
    print("C:\\FARMS\\デモファームの全牛の状態（lact, clvd, rc, pen）を再計算してDBを更新")
    print("=" * 80)
    print()
    print("このスクリプトは以下を更新します:")
    print("  - lact (産次)")
    print("  - clvd (最終分娩日)")
    print("  - rc (繁殖コード)")
    print("  - pen (群)")
    print()
    main()

