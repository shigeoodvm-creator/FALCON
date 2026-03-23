"""
既存のAI/ETイベント全てにoutcome（P/O/R/N）を計算して記録するスクリプト

使用方法:
    python scripts/update_insemination_outcomes.py
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT))
sys.path.insert(0, str(APP_ROOT / "app"))

from app.db.db_handler import DBHandler
from app.modules.rule_engine import RuleEngine
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)


def main():
    """メイン処理"""
    # 農場パスを取得（環境変数またはデフォルト）
    import os
    farm_path = os.getenv('FARM_PATH', 'C:/FARMS')
    
    # 農場フォルダを検索
    farms_root = Path(farm_path)
    if not farms_root.exists():
        logger.error(f"農場フォルダが見つかりません: {farms_root}")
        return
    
    # 農場フォルダを取得
    farms = []
    for item in farms_root.iterdir():
        if item.is_dir():
            db_path = item / "farm.db"
            if db_path.exists():
                farms.append(item)
    
    if not farms:
        logger.error("農場が見つかりません")
        return
    
    # 各農場で処理
    for farm_path in farms:
        logger.info(f"農場を処理中: {farm_path.name}")
        
        db_path = farm_path / "farm.db"
        db = DBHandler(db_path)
        rule_engine = RuleEngine(db)
        
        # 全牛のoutcomeを更新
        conn = db.connect()
        cursor = conn.cursor()
        
        # 全牛のauto_idを取得
        cursor.execute("SELECT DISTINCT auto_id FROM cow")
        cows = cursor.fetchall()
        
        for cow_row in cows:
            cow_auto_id = cow_row['auto_id']
            try:
                rule_engine.update_insemination_outcomes(cow_auto_id)
            except Exception as e:
                logger.error(f"牛 auto_id={cow_auto_id} の処理エラー: {e}")
        
        conn.close()
        logger.info(f"農場 {farm_path.name} の処理完了")
    
    logger.info("全農場の処理完了")


if __name__ == "__main__":
    main()









