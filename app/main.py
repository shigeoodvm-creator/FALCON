import tkinter as tk
import logging
import json
from pathlib import Path
from ui.farm_selector import FarmSelectorWindow
from ui.main_window import MainWindow
from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.dictionary_sync import DictionarySync
from settings_manager import SettingsManager

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('falcon.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)


def main():
    print("DEBUG: app.main.main() が呼ばれました")
    logger.info("FALCON2 アプリケーションを起動しました")

    root = tk.Tk()
    root.title("FALCON2")
    root.geometry("600x400")

    # FALCONフォルダのルートパスを取得
    falcon_root = Path(__file__).parent.parent

    def on_farm_selected(farm_path: Path):
        """農場選択時のコールバック"""
        print(f"農場が選択されました: {farm_path}")
        logger.info(f"農場が選択されました: {farm_path}")
        
        # ========== 辞書同期処理 ==========
        try:
            sync = DictionarySync(falcon_root)
            event_dict, item_dict = sync.sync_dictionaries(farm_path)
            logger.info("辞書同期処理が完了しました")
        except Exception as e:
            logger.error(f"辞書同期処理でエラーが発生しました: {e}")
            print(f"辞書同期処理でエラーが発生しました: {e}")

        # 設定ファイル（farm_settings.json）を初期化・ロード
        try:
            settings_manager = SettingsManager(farm_path)
            settings_manager.load()  # pen_settings を含む初期キーを生成
        except Exception as e:
            logger.error(f"設定ファイル初期化に失敗しました: {e}")
        
        # データベースパス
        db_path = farm_path / "farm.db"
        
        # DBHandlerを初期化
        db_handler = DBHandler(db_path)
        
        # item_dictionary.jsonのパス（農場フォルダ側を優先）
        item_dict_path = farm_path / "item_dictionary.json"
        if not item_dict_path.exists():
            # 農場フォルダに存在しない場合はエラー（同期処理で作成されるはず）
            logger.warning(f"item_dictionary.json が見つかりません: {item_dict_path}")
            item_dict_path = None
        
        # FormulaEngineを初期化
        formula_engine = FormulaEngine(db_handler, item_dict_path)
        
        # RuleEngineを初期化
        rule_engine = RuleEngine(db_handler)
        
        # ========== 授精設定のロード ==========
        insemination_settings_file = farm_path / "insemination_settings.json"
        if not insemination_settings_file.exists():
            # ファイルが存在しない場合は空の初期ファイルを自動生成
            initial_data = {
                "technicians": {},
                "insemination_types": {}
            }
            try:
                farm_path.mkdir(parents=True, exist_ok=True)
                with open(insemination_settings_file, 'w', encoding='utf-8') as f:
                    json.dump(initial_data, f, ensure_ascii=False, indent=2)
                logger.info("授精設定ファイルを初期化しました")
            except Exception as e:
                logger.error(f"授精設定ファイルの初期化に失敗しました: {e}")
        else:
            # ファイルが存在する場合はロード
            try:
                with open(insemination_settings_file, 'r', encoding='utf-8') as f:
                    insemination_data = json.load(f)
                    technicians = insemination_data.get('technicians', {})
                    insemination_types = insemination_data.get('insemination_types', {})
                    logger.info("Insemination settings loaded")
                    logger.info(f"Technicians: {len(technicians)}")
                    logger.info(f"Insemination types: {len(insemination_types)}")
            except Exception as e:
                logger.error(f"授精設定ファイルの読み込みに失敗しました: {e}")
        
        # メインウィンドウを初期化して表示
        # 既存のウィジェットをクリア
        for widget in root.winfo_children():
            widget.destroy()
        
        # メインウィンドウを作成
        main_window = MainWindow(root, db_handler, formula_engine, rule_engine, farm_path)
        
        print("メインウィンドウを表示しました")
        logger.info("メインウィンドウを表示しました")

    # 最初の画面として FarmSelector を表示
    FarmSelectorWindow(root, on_farm_selected=on_farm_selected)

    root.mainloop()
