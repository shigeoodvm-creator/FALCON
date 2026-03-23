"""
既存の全AI/ETイベントに対してoutcomeを更新するスクリプト
"""
import sqlite3
from pathlib import Path
import logging
from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def update_all_farms():
    """全デモファームのAI/ETイベントのoutcomeを更新"""
    farms = [
        Path("C:/FARMS/デモファーム"),
        Path("C:/FARMS/DemoFarm"),
        Path("C:/FARMS/DemoFarm2"),
        Path("C:/FARMS/DemoFarm3"),
        Path("C:/FARMS/DemoFarm4"),
    ]
    
    for farm_path in farms:
        db_path = farm_path / "farm.db"
        if not db_path.exists():
            logger.info(f"スキップ: {farm_path.name} (farm.dbが見つかりません)")
            continue
        
        logger.info(f"処理中: {farm_path.name}")
        
        try:
            db = DBHandler(db_path)
            rule_engine = RuleEngine(db)
            
            # 全牛を取得
            all_cows = db.get_all_cows()
            logger.info(f"  牛数: {len(all_cows)}")
            
            updated_count = 0
            for cow in all_cows:
                cow_auto_id = cow.get('auto_id')
                if cow_auto_id:
                    try:
                        rule_engine.update_insemination_outcomes(cow_auto_id)
                        updated_count += 1
                    except Exception as e:
                        logger.error(f"  エラー (cow_auto_id={cow_auto_id}): {e}")
            
            logger.info(f"  完了: {updated_count}頭のoutcomeを更新")
            db.close()
            
        except Exception as e:
            logger.error(f"  エラー: {e}", exc_info=True)

if __name__ == "__main__":
    print("=" * 80)
    print("既存の全AI/ETイベントに対してoutcomeを更新")
    print("=" * 80)
    print()
    
    update_all_farms()
    
    print()
    print("=" * 80)
    print("更新完了")
    print("=" * 80)

















