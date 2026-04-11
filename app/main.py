import tkinter as tk
from tkinter import messagebox
import logging
import logging.handlers
import json
import os
import sys
import traceback
from pathlib import Path
from typing import Optional
from ui.farm_selector import FarmSelectorWindow
from ui.date_input_normalizer import install_date_input_normalizer
from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from settings_manager import SettingsManager
from constants import FARMS_ROOT, CONFIG_DEFAULT_DIR

# ========== ログ設定 ==========
# EXE化後もサポートに送れるよう %APPDATA%/FALCON2/falcon2.log へ出力
_LOG_DIR  = Path(os.environ.get("APPDATA", Path.home())) / "FALCON2"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / "falcon2.log"

_log_formatter = logging.Formatter(
    "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
_file_handler = logging.handlers.RotatingFileHandler(
    _LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_file_handler.setFormatter(_log_formatter)
_file_handler.setLevel(logging.WARNING)

_console_handler = logging.StreamHandler()
_console_handler.setFormatter(_log_formatter)
_console_handler.setLevel(logging.WARNING)

logging.basicConfig(level=logging.WARNING, handlers=[_file_handler, _console_handler])

logger = logging.getLogger(__name__)


def _setup_excepthook():
    """未捕捉例外をログに記録し、ユーザーにログパスを案内するダイアログを出す"""
    def _handler(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logger.critical(f"未捕捉例外:\n{msg}")
        try:
            messagebox.showerror(
                "予期せぬエラー",
                f"アプリが予期せぬエラーで停止しました。\n\n"
                f"サポートへ以下のログファイルをお送りください:\n{_LOG_FILE}\n\n"
                f"エラー: {exc_value}"
            )
        except Exception:
            pass
    sys.excepthook = _handler

_setup_excepthook()


def main(cow_id: Optional[str] = None):
    """
    メイン関数
    
    Args:
        cow_id: 個体ID（4桁の管理番号、例: "0980"）。指定された場合、個体カードウィンドウを開く
    """
    logger = logging.getLogger(__name__)
    
    # ========== 起動時の重い処理を非同期化（バックグラウンドで実行） ==========
    # 辞書同期とnormalization辞書生成は、起動後にバックグラウンドで実行
    # これにより起動時間を大幅に短縮
    def _background_sync():
        """バックグラウンドで農場側の項目・イベント辞書を削除し、normalization辞書を生成"""
        try:
            from modules.dictionary_sync import DictionarySync
            sync = DictionarySync()
            sync.delete_farm_dictionary_files(FARMS_ROOT)
            logger.info("バックグラウンド：農場側の項目・イベント辞書の削除を完了しました")
        except Exception as e:
            logger.warning(f"バックグラウンド辞書削除でエラー: {e}")
        
        try:
            from modules.normalization_generator import generate_normalization_dict
            logger.info("バックグラウンド：normalization辞書を自動生成中...")
            generate_normalization_dict()
            logger.info("バックグラウンド：normalization辞書の自動生成が完了しました")
        except Exception as e:
            logger.warning(f"バックグラウンドnormalization辞書生成でエラー: {e}")
    
    # バックグラウンド処理をスレッドで実行（起動をブロックしない）
    import threading
    sync_thread = threading.Thread(target=_background_sync, daemon=True)
    sync_thread.start()

    root = tk.Tk()
    root.title("FALCON2")
    root.geometry("600x400")

    # アプリアイコン設定
    try:
        from pathlib import Path as _Path
        _ico = _Path(__file__).parent / "resources" / "falcon2.ico"
        if _ico.exists():
            root.iconbitmap(str(_ico))
    except Exception:
        pass
    # 全ウィンドウ共通：日付入力欄で全角数字/記号を半角に統一
    install_date_input_normalizer(root)
    # 起動時はルートを非表示にし、農場選択ウィンドウだけを前面に表示
    root.withdraw()

    # スプラッシュスクリーン表示
    try:
        from ui.splash_screen import SplashScreen
        from constants import APP_VERSION
        _splash = SplashScreen(root, version=APP_VERSION)
        root.update()
    except Exception as _e:
        logger.warning(f"スプラッシュ表示失敗: {_e}")
        _splash = None

    # ========== ライセンスチェック ==========
    from modules.license_manager import get_license_manager, LicenseStatus
    from ui.license_window import LicenseActivationWindow

    _license_info = get_license_manager().check()

    if _license_info.status == LicenseStatus.TRIAL_EXPIRED:
        # 試用期限切れ → アクティベーション必須（閉じると終了）
        LicenseActivationWindow(root, allow_close=False, info=_license_info).show()
        _license_info = get_license_manager().check()
        if not _license_info.is_usable:
            root.quit()
            return
    elif _license_info.status == LicenseStatus.LICENSE_EXPIRED:
        # ライセンス期限切れ → 更新を促す
        LicenseActivationWindow(root, allow_close=False, info=_license_info).show()
        _license_info = get_license_manager().check()
        if not _license_info.is_usable:
            root.quit()
            return
    elif _license_info.status == LicenseStatus.LICENSE_INVALID:
        messagebox.showerror(
            "ライセスエラー",
            "ライセンスファイルが不正です。\n"
            "ライセンスキーを再入力するか、サポートにご連絡ください。"
        )
        LicenseActivationWindow(root, allow_close=False, info=_license_info).show()
        _license_info = get_license_manager().check()
        if not _license_info.is_usable:
            root.quit()
            return

    # ライセンス情報をグローバルに保持（農場数チェック等で参照）
    import builtins
    builtins._falcon_license = _license_info

    # レポート散布図クリック→個体カードを開く用のローカルサーバを起動
    try:
        from modules.report_cow_bridge import start_server
        start_server()
    except Exception as e:
        logger.warning("レポート個体カードブリッジの起動をスキップ: %s", e)

    def on_farm_selected(farm_path: Path):
        """農場選択時のコールバック"""
        logger.info(f"農場が選択されました: {farm_path}")
        # ルートを再表示（農場選択後はメインウィンドウを表示するため）
        root.deiconify()
        
        # ========== 辞書同期処理 ==========
        # 【重要】起動時に最新JSON優先同期を実行済みのため、農場選択時の同期は不要。
        # 農場選択時にマスターから同期すると、最新JSON優先同期で更新した内容が上書きされてしまう。
        # 同期が必要な場合は、起動時の最新JSON優先同期で既に処理済み。

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
        
        # 項目辞書・イベント辞書は本体（config_default）のみ参照。農場側に残っているファイルは削除する。
        for name in ("item_dictionary.json", "event_dictionary.json"):
            farm_dict_file = farm_path / name
            if farm_dict_file.exists():
                try:
                    farm_dict_file.unlink()
                    logger.info(f"農場側の辞書を削除しました（本体参照のため）: {farm_dict_file}")
                except Exception as e:
                    logger.warning(f"農場側辞書の削除をスキップ: {farm_dict_file} - {e}")
        item_dict_path = CONFIG_DEFAULT_DIR / "item_dictionary.json"
        if not item_dict_path.exists():
            item_dict_path = None
        event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
        if not event_dict_path.exists():
            event_dict_path = None
        
        # FormulaEngineを初期化（本体の項目辞書を参照）
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
        
        # メインウィンドウを作成（for ループの外で必ず 1 回だけ実行）
        # 遅延import: 農場選択後に初めてMainWindowを読み込む（起動時間短縮）
        from ui.main_window import MainWindow
        main_window = MainWindow(root, db_handler, formula_engine, rule_engine, farm_path)
        
        # 産次が NULL だがイベントがある個体をイベントから産次・event_lact/event_dim を記録し直す
        try:
            fixed = rule_engine.fix_null_lact_cows()
            if fixed > 0:
                logger.info(f"産次NULLの個体を記録し直しました: {fixed} 頭")
        except Exception as e:
            logger.warning(f"産次NULLの個体の修正でエラー（続行）: {e}")
        
        # 産次が設定されている全頭で event_lact/event_dim を cow.lact に合わせて同期（最高乳量などの整合性のため）
        try:
            synced = rule_engine.sync_event_lact_for_all_cows()
            if synced > 0:
                logger.info(f"産次に合わせてイベントの産次を同期しました: {synced} 頭")
        except Exception as e:
            logger.warning(f"イベント産次同期でエラー（続行）: {e}")
        
        logger.info("メインウィンドウを表示しました")

        # ========== 月次チェック（長期空胎牛 → 分娩過期牛の順に表示） ==========
        def _run_monthly_checks():
            try:
                from modules.long_open_alert import show_alert_window as _open_alert
                _open_alert(root, farm_path, db_handler, formula_engine, rule_engine)
            except Exception as e:
                logger.warning(f"長期空胎牛チェックをスキップ: {e}")
            try:
                from modules.overdue_calving_alert import show_alert_window as _calv_alert
                _calv_alert(root, farm_path, db_handler, formula_engine, rule_engine)
            except Exception as e:
                logger.warning(f"分娩過期チェックをスキップ: {e}")

        root.after(1500, _run_monthly_checks)

        # ── バックアップリマインダー ────────────────────────────
        def _check_backup_reminder():
            try:
                from backup_manager import load_backup_settings, should_run_backup, run_backup
                settings = load_backup_settings(farm_path)
                dest = (settings.get("destination_path") or "").strip()

                if not dest:
                    # 保存先未設定
                    if messagebox.askyesno(
                        "バックアップ設定のお勧め",
                        "バックアップ先フォルダが設定されていません。\n"
                        "万一のデータ消失に備えて、バックアップ設定を行ってください。\n\n"
                        "今すぐ設定しますか？"
                    ):
                        from ui.backup_settings_window import BackupSettingsWindow
                        BackupSettingsWindow(root, farm_path).window.wait_window()
                elif should_run_backup(farm_path):
                    # バックアップ期限が来ている
                    interval = settings.get("interval_months", 1)
                    last_ym  = settings.get("last_backup_ym") or "未実施"
                    ans = messagebox.askyesno(
                        "バックアップのお知らせ",
                        f"バックアップの実施タイミングです（設定: {interval}か月ごと）。\n"
                        f"前回: {last_ym}\n\n"
                        f"今すぐバックアップを実行しますか？"
                    )
                    if ans:
                        result = run_backup(farm_path)
                        if result:
                            messagebox.showinfo("完了", f"バックアップが完了しました。\n保存先: {result}")
                        else:
                            messagebox.showwarning("失敗", "バックアップに失敗しました。\n保存先フォルダを確認してください。")
            except Exception as _e:
                logger.warning(f"バックアップリマインダーでエラー: {_e}")

        root.after(3000, _check_backup_reminder)

        # 個体IDが指定されている場合は、個体カードウィンドウを開く
        if cow_id:
            _open_cow_card_window(root, db_handler, formula_engine, rule_engine, farm_path, cow_id)

    # スプラッシュを閉じてから農場選択を表示
    if _splash:
        _splash.set_status("農場を選択してください")
        root.after(400, _splash.close)

    # 最初の画面として FarmSelector を表示
    FarmSelectorWindow(root, on_farm_selected=on_farm_selected)

    root.mainloop()


def _open_cow_card_window(root: tk.Tk, db_handler: DBHandler, 
                          formula_engine: FormulaEngine, rule_engine: RuleEngine,
                          farm_path: Path, cow_id: str):
    """
    個体カードウィンドウを開く
    
    Args:
        root: ルートウィンドウ
        db_handler: DBHandler インスタンス
        formula_engine: FormulaEngine インスタンス
        rule_engine: RuleEngine インスタンス
        farm_path: 農場フォルダのパス
        cow_id: 個体ID（4桁の管理番号、例: "0980"）
    """
    try:
        from ui.cow_card_window import CowCardWindow
        
        # 項目辞書・イベント辞書は本体（config_default）を参照
        event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
        if not event_dict_path.exists():
            event_dict_path = None
        item_dict_path = CONFIG_DEFAULT_DIR / "item_dictionary.json"
        if not item_dict_path.exists():
            item_dict_path = None
        
        # 個体IDで検索（4桁のcow_idで検索）
        cows = db_handler.get_cows_by_id(cow_id)
        
        if not cows:
            # 見つからない場合は、正規化されたIDで検索を試みる
            normalized_id = cow_id.lstrip('0')  # 左ゼロを除去
            cow = db_handler.get_cow_by_normalized_id(normalized_id)
            if cow:
                cows = [cow]
        
        if not cows:
            messagebox.showerror("エラー", f"ID {cow_id} の個体は見つかりませんでした")
            return
        
        # 複数の個体が見つかった場合は最初の1頭を表示
        selected_cow = cows[0]
        cow_auto_id = selected_cow.get('auto_id')
        
        if not cow_auto_id:
            messagebox.showerror("エラー", "個体情報の取得に失敗しました")
            return
        
        # 個体カードウィンドウを開く
        cow_card_window = CowCardWindow(
            parent=root,
            db_handler=db_handler,
            formula_engine=formula_engine,
            rule_engine=rule_engine,
            event_dictionary_path=event_dict_path,
            item_dictionary_path=item_dict_path,
            cow_auto_id=cow_auto_id
        )
        cow_card_window.show()
        
        logger.info(f"個体カードウィンドウを開きました: cow_id={cow_id}, auto_id={cow_auto_id}")
        
    except Exception as e:
        import traceback
        logger.error(f"個体カードウィンドウの表示に失敗しました: {e}")
        traceback.print_exc()
        messagebox.showerror("エラー", f"個体カードウィンドウの表示に失敗しました: {e}")
