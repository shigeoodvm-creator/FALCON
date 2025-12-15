"""
FALCON2 - RuleEngine テストスクリプト
イベント追加/削除時の状態再計算を検証
"""

import sys
import os
import json
import tempfile
from pathlib import Path
from datetime import datetime, timedelta

# アプリケーションパスを追加
APP_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(APP_ROOT / "app"))

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine


class EventDictionary:
    """event_dictionary.json を読み込んで alias から event_number を解決"""
    
    def __init__(self, dict_path: Path):
        """
        初期化
        
        Args:
            dict_path: event_dictionary.json のパス
        """
        self.dict_path = dict_path
        self.event_dict = {}
        self._load()
        # 設計書の定義も追加（event_dictionary.jsonにない場合のフォールバック）
        self._add_design_spec_events()
    
    def _load(self):
        """event_dictionary.json を読み込む"""
        if self.dict_path.exists():
            with open(self.dict_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # event_number をキーに、alias を値にマッピング
                for event_num_str, event_data in data.items():
                    if 'alias' in event_data:
                        alias = event_data['alias']
                        event_num = int(event_num_str)
                        self.event_dict[alias] = event_num
    
    def _add_design_spec_events(self):
        """設計書の定義を追加（event_dictionary.jsonにない場合のフォールバック）"""
        # 設計書 第3章の定義
        design_events = {
            'AI': 200,
            'ET': 201,
            'CALV': 202,
            'DRY': 203,
            'STOPR': 204,
            'SOLD': 205,
            'DEAD': 206,
            'FCHK': 300,
            'REPRO': 301,
            'PDN': 302,
            'PDP': 303,
            'PDP2': 304,
            'ABRT': 305,
            'PAGN': 306,
            'PAGP': 307,
            'MILK_TEST': 601,
            'MOVE': 611,
        }
        
        for alias, event_num in design_events.items():
            if alias not in self.event_dict:
                self.event_dict[alias] = event_num
    
    def get_event_number(self, alias: str) -> int:
        """
        alias から event_number を取得
        
        Args:
            alias: イベントの alias（例: "AI", "CALV", "PDP"）
        
        Returns:
            event_number
        
        Raises:
            ValueError: alias が見つからない場合
        """
        if alias not in self.event_dict:
            raise ValueError(f"イベント alias '{alias}' が見つかりません")
        return self.event_dict[alias]


def create_test_db() -> Tuple[DBHandler, Path]:
    """
    テスト用の一時DBを作成
    
    Returns:
        (DBHandler, temp_db_path)
    """
    # 一時ファイルとして farm.db を作成
    temp_dir = Path(tempfile.mkdtemp())
    temp_db_path = temp_dir / "farm.db"
    
    db = DBHandler(temp_db_path)
    db.create_tables()
    
    return db, temp_db_path


def create_test_cow(db: DBHandler) -> int:
    """
    テスト用の牛を1頭作成
    
    Args:
        db: DBHandler
    
    Returns:
        auto_id
    """
    cow_data = {
        'cow_id': '1234',
        'jpn10': '0123456789',
        'brd': 'ホルスタイン',
        'bthd': '2020-01-01',
        'entr': '2020-01-15',
        'lact': 0,
        'clvd': None,
        'rc': None,
        'pen': 'A1',
        'frm': 'TEST'
    }
    
    auto_id = db.insert_cow(cow_data)
    return auto_id


def add_event(db: DBHandler, cow_auto_id: int, event_number: int, 
              event_date: str, json_data: dict = None, note: str = None) -> int:
    """
    イベントを追加
    
    Args:
        db: DBHandler
        cow_auto_id: 牛の auto_id
        event_number: イベント番号
        event_date: イベント日付
        json_data: JSONデータ
        note: 備考
    
    Returns:
        event_id
    """
    event_data = {
        'cow_auto_id': cow_auto_id,
        'event_number': event_number,
        'event_date': event_date,
        'json_data': json_data or {},
        'note': note
    }
    
    event_id = db.insert_event(event_data)
    return event_id


def delete_event(db: DBHandler, event_id: int):
    """
    イベントを削除（論理削除）
    
    Args:
        db: DBHandler
        event_id: イベントID
    """
    db.delete_event(event_id, soft_delete=True)


def assert_cow_state(db: DBHandler, cow_auto_id: int, 
                     expected_rc: int = None,
                     expected_lact: int = None,
                     expected_clvd: str = None,
                     expected_last_ai_date: str = None,
                     expected_due_date: str = None,
                     message: str = ""):
    """
    牛の状態をアサート
    
    Args:
        db: DBHandler
        cow_auto_id: 牛の auto_id
        expected_rc: 期待する繁殖コード
        expected_lact: 期待する産次
        expected_clvd: 期待する最終分娩日
        expected_last_ai_date: 期待する最終AI日
        expected_due_date: 期待する分娩予定日
        message: アサートメッセージ
    """
    cow = db.get_cow_by_auto_id(cow_auto_id)
    assert cow is not None, f"牛が見つかりません: {message}"
    
    if expected_rc is not None:
        assert cow['rc'] == expected_rc, \
            f"{message}: rc が期待値と異なります。期待={expected_rc}, 実際={cow['rc']}"
    
    if expected_lact is not None:
        assert cow['lact'] == expected_lact, \
            f"{message}: lact が期待値と異なります。期待={expected_lact}, 実際={cow['lact']}"
    
    if expected_clvd is not None:
        assert cow['clvd'] == expected_clvd, \
            f"{message}: clvd が期待値と異なります。期待={expected_clvd}, 実際={cow['clvd']}"
    
    # last_ai_date と due_date は cow テーブルには保存されない（FormulaEngineで計算）
    # RuleEngine の apply_events で計算された状態を確認
    rule_engine = RuleEngine(db)
    state = rule_engine.apply_events(cow_auto_id)
    
    if expected_last_ai_date is not None:
        assert state.get('last_ai_date') == expected_last_ai_date, \
            f"{message}: last_ai_date が期待値と異なります。期待={expected_last_ai_date}, 実際={state.get('last_ai_date')}"
    
    if expected_due_date is not None:
        assert state.get('due_date') == expected_due_date, \
            f"{message}: due_date が期待値と異なります。期待={expected_due_date}, 実際={state.get('due_date')}"


def main():
    """メインテスト"""
    print("=" * 60)
    print("RuleEngine テスト開始")
    print("=" * 60)
    
    # event_dictionary.json を読み込む
    dict_path = APP_ROOT / "docs" / "event_dictionary.json"
    event_dict = EventDictionary(dict_path)
    
    # テスト用DBを作成
    db, temp_db_path = create_test_db()
    print(f"テストDB作成: {temp_db_path}")
    
    try:
        # RuleEngine を初期化
        rule_engine = RuleEngine(db)
        
        # テスト用の牛を1頭作成
        cow_auto_id = create_test_cow(db)
        print(f"テスト牛作成: auto_id={cow_auto_id}, cow_id=1234")
        
        # 日付を設定
        base_date = datetime(2024, 1, 1)
        
        # ========== (1) CALV 追加 ==========
        print("\n[1] 分娩イベント（CALV）を追加")
        calv_date = (base_date + timedelta(days=0)).strftime('%Y-%m-%d')
        calv_event_id = add_event(
            db, cow_auto_id, 
            event_dict.get_event_number('CALV'),
            calv_date
        )
        rule_engine.on_event_added(calv_event_id)
        
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_FRESH,
            expected_lact=1,
            expected_clvd=calv_date,
            message="[1] CALV追加後"
        )
        print("[OK] rc=Fresh, lact=1, clvd更新 を確認")
        
        # ========== (2) AI 追加 ==========
        print("\n[2] AIイベントを追加")
        ai_date = (base_date + timedelta(days=60)).strftime('%Y-%m-%d')
        ai_event_id = add_event(
            db, cow_auto_id,
            event_dict.get_event_number('AI'),
            ai_date,
            json_data={'sire': 'TEST001', 'technician': 'T01'}
        )
        rule_engine.on_event_added(ai_event_id)
        
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_BRED,
            expected_last_ai_date=ai_date,
            message="[2] AI追加後"
        )
        print("[OK] rc=Bred, last_ai_date更新 を確認")
        
        # ========== (3) PREG+ 追加 ==========
        print("\n[3] 妊娠鑑定プラス（PDP）を追加")
        preg_date = (base_date + timedelta(days=95)).strftime('%Y-%m-%d')
        preg_event_id = add_event(
            db, cow_auto_id,
            event_dict.get_event_number('PDP'),
            preg_date
        )
        rule_engine.on_event_added(preg_event_id)
        
        # 受胎日はAI日、分娩予定日は受胎日+280日
        expected_conception_date = ai_date
        expected_due_date = (datetime.strptime(ai_date, '%Y-%m-%d') + timedelta(days=280)).strftime('%Y-%m-%d')
        
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_PREGNANT,
            expected_due_date=expected_due_date,
            message="[3] PREG+追加後"
        )
        print(f"[OK] rc=Pregnant, due_date={expected_due_date} を確認")
        
        # ========== (4) PREG+ 削除 ==========
        print("\n[4] 妊娠鑑定プラス（PDP）を削除")
        delete_event(db, preg_event_id)
        rule_engine.on_event_deleted(preg_event_id)
        
        # PREG+削除後は、直近のイベント（AI）に基づいて Bred に戻る
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_BRED,
            expected_last_ai_date=ai_date,
            message="[4] PREG+削除後"
        )
        print("[OK] rc=Bred に戻ることを確認")
        
        # ========== (5) DRY 追加 ==========
        print("\n[5] 乾乳イベント（DRY）を追加")
        dry_date = (base_date + timedelta(days=200)).strftime('%Y-%m-%d')
        dry_event_id = add_event(
            db, cow_auto_id,
            event_dict.get_event_number('DRY'),
            dry_date
        )
        rule_engine.on_event_added(dry_event_id)
        
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_DRY,
            message="[5] DRY追加後"
        )
        print("[OK] rc=Dry を確認")
        
        # ========== (6) DRY 削除 ==========
        print("\n[6] 乾乳イベント（DRY）を削除")
        delete_event(db, dry_event_id)
        rule_engine.on_event_deleted(dry_event_id)
        
        # DRY削除後は、直近のイベント（AI）に基づいて Bred に戻る
        assert_cow_state(
            db, cow_auto_id,
            expected_rc=RuleEngine.RC_BRED,
            expected_last_ai_date=ai_date,
            message="[6] DRY削除後"
        )
        print("[OK] rc=Bred に戻ることを確認")
        
        print("\n" + "=" * 60)
        print("[OK] 全テスト成功！")
        print("=" * 60)
        
    except AssertionError as e:
        print(f"\n[ERROR] アサーションエラー: {e}")
        raise
    except Exception as e:
        print(f"\n[ERROR] エラー発生: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        # 一時DBをクリーンアップ
        db.close()
        if temp_db_path.exists():
            temp_db_path.unlink()
            temp_db_path.parent.rmdir()
            print(f"\nテストDB削除: {temp_db_path}")


if __name__ == "__main__":
    main()

