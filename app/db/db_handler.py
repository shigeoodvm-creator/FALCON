"""
FALCON2 - データベースハンドラ
SQLite (farm.db) の CRUD 操作を担当
設計書 第16章参照
"""

import sqlite3
import json
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any
from datetime import datetime


class DBHandler:
    """SQLite データベースハンドラ"""
    
    def __init__(self, db_path: Path):
        """
        初期化
        
        Args:
            db_path: farm.db のパス (例: C:/FARMS/FarmA/farm.db)
        """
        self.db_path = Path(db_path)
        self.conn: Optional[sqlite3.Connection] = None
        self._ensure_db_exists()
    
    def _ensure_db_exists(self):
        """データベースファイルが存在しない場合は作成し、既存DBにはマイグレーションを適用"""
        if not self.db_path.exists():
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self.create_tables()
        else:
            self._create_connection()
            self._run_migrations()
    
    def connect(self) -> sqlite3.Connection:
        """
        データベースに接続
        
        【スレッド安全性】
        SQLiteのスレッド制約を考慮し、check_same_thread=Falseを設定。
        ただし、同一接続の同時使用は避け、各操作は独立して実行すること。
        """
        # 接続が存在しない、または閉じられている場合は再接続
        if self.conn is None:
            self._create_connection()
        else:
            # 接続が閉じられているかチェック
            try:
                # 簡単なクエリで接続が有効か確認
                self.conn.execute("SELECT 1")
            except (sqlite3.ProgrammingError, sqlite3.OperationalError):
                # 接続が閉じられている場合は再接続
                logging.warning("データベース接続が閉じられていました。再接続します。")
                self.conn = None
                self._create_connection()
        
        return self.conn
    
    def _create_connection(self):
        """新しいデータベース接続を作成"""
        self.conn = sqlite3.connect(
            str(self.db_path),
            check_same_thread=False  # スレッドセーフ性のため
        )
        self.conn.row_factory = sqlite3.Row  # 辞書形式で取得
        # 外部キー制約を有効化
        self.conn.execute("PRAGMA foreign_keys = ON")
    
    def close(self):
        """データベース接続を閉じる"""
        if self.conn:
            self.conn.close()
            self.conn = None
    
    def create_tables(self):
        """テーブルを作成（設計書 第16章参照）"""
        conn = self.connect()
        cursor = conn.cursor()
        
        # cow テーブル（設計書 16.1）
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS cow (
                auto_id      INTEGER PRIMARY KEY AUTOINCREMENT,
                cow_id       TEXT NOT NULL,
                jpn10        TEXT NOT NULL,
                brd          TEXT,
                bthd         TEXT,
                entr         TEXT,
                lact         INTEGER,
                clvd         TEXT,
                rc           INTEGER,
                pen          TEXT,
                frm          TEXT,
                UNIQUE(cow_id, jpn10)
            )
        """)
        
        # event テーブル（設計書 16.2）
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS event (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                cow_auto_id   INTEGER NOT NULL,
                event_number  INTEGER NOT NULL,
                event_date    TEXT NOT NULL,
                json_data     TEXT,
                note          TEXT,
                deleted       INTEGER DEFAULT 0,
                event_dim     INTEGER,
                event_lact    INTEGER,
                FOREIGN KEY(cow_auto_id) REFERENCES cow(auto_id)
            )
        """)
        
        # event_dimとevent_lactカラムが存在しない場合は追加（マイグレーション）
        try:
            cursor.execute("ALTER TABLE event ADD COLUMN event_dim INTEGER")
        except sqlite3.OperationalError:
            # 既に存在する場合はスキップ
            pass
        
        try:
            cursor.execute("ALTER TABLE event ADD COLUMN event_lact INTEGER")
        except sqlite3.OperationalError:
            # 既に存在する場合はスキップ
            pass
        
        # baseline_lact: 分娩イベント0件時の産次。分娩削除で元の産次に戻すために使用。
        try:
            cursor.execute("ALTER TABLE cow ADD COLUMN baseline_lact INTEGER")
        except sqlite3.OperationalError:
            pass
        # 既存行で baseline_lact が NULL の場合は lact で埋める
        try:
            cursor.execute("UPDATE cow SET baseline_lact = lact WHERE baseline_lact IS NULL")
        except sqlite3.OperationalError:
            pass
        
        # item_value テーブル（設計書 16.3）
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS item_value (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                cow_auto_id  INTEGER NOT NULL,
                item_key     TEXT NOT NULL,
                value        TEXT,
                FOREIGN KEY(cow_auto_id) REFERENCES cow(auto_id)
            )
        """)
        
        # インデックス作成（設計書 16.4）
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_event_cow ON event(cow_auto_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_event_date ON event(event_date DESC)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_cow ON item_value(cow_auto_id)")
        
        conn.commit()
    
    # 現在のスキーマバージョン（新しいマイグレーション追加時はこの値をインクリメント）
    _SCHEMA_VERSION = 1

    def _get_schema_version(self) -> int:
        """SQLite user_version プラグマでスキーマバージョンを取得"""
        try:
            row = self.conn.execute("PRAGMA user_version").fetchone()
            return row[0] if row else 0
        except Exception:
            return 0

    def _set_schema_version(self, version: int):
        """スキーマバージョンを設定（user_version にはプレースホルダ不可）"""
        self.conn.execute(f"PRAGMA user_version = {int(version)}")

    def _run_migrations(self):
        """既存DBに対するマイグレーション（スキーマバージョン管理付き）"""
        if self.conn is None:
            return
        cursor = self.conn.cursor()

        # --- 軽量な構造変更（冪等なので毎回実行しても問題なし） ---
        try:
            cursor.execute("ALTER TABLE cow ADD COLUMN baseline_lact INTEGER")
        except sqlite3.OperationalError:
            pass
        try:
            cursor.execute("UPDATE cow SET baseline_lact = lact WHERE baseline_lact IS NULL")
        except sqlite3.OperationalError:
            pass
        self.conn.commit()

        # --- バージョン付きマイグレーション（各バージョンは1回だけ実行） ---
        current_version = self._get_schema_version()

        if current_version < 1:
            # v1: baseline_lact を「lact - 非baseline分娩数」に揃える
            self._migrate_baseline_lact_from_calv_count()
            self._set_schema_version(1)
            self.conn.commit()
            logging.info("DBマイグレーション v1 完了: baseline_lact を再計算")

    def _migrate_baseline_lact_from_calv_count(self):
        """各牛の baseline_lact を lact - 非baseline分娩イベント数 に再計算する（マイグレーション v1）"""
        try:
            cows = self.get_all_cows()
            for cow in cows:
                auto_id = cow.get('auto_id')
                if auto_id is None:
                    continue
                events = self.get_events_by_cow(auto_id, include_deleted=False)
                n = 0
                for e in events:
                    if e.get('event_number') != 202:
                        continue
                    j = e.get('json_data') or {}
                    if isinstance(j, str):
                        try:
                            j = json.loads(j)
                        except Exception:
                            j = {}
                    if j.get('baseline_calving', False):
                        continue
                    n += 1
                lact = cow.get('lact')
                if lact is None:
                    lact = 0
                baseline_lact = max(0, lact - n)
                self.update_cow(auto_id, {'baseline_lact': baseline_lact})
        except Exception as e:
            logging.warning(f"baseline_lact migration skip: {e}")
    
    # ========== cow テーブル操作 ==========
    
    def insert_cow(self, cow_data: Dict[str, Any]) -> int:
        """
        牛を追加
        
        Args:
            cow_data: 牛データ（cow_id, jpn10, brd, bthd, entr, lact, clvd, rc, pen, frm）
        
        Returns:
            追加された auto_id
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("""
            INSERT INTO cow (cow_id, jpn10, brd, bthd, entr, lact, clvd, rc, pen, frm, baseline_lact)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            cow_data.get('cow_id'),
            cow_data.get('jpn10'),
            cow_data.get('brd'),
            cow_data.get('bthd'),
            cow_data.get('entr'),
            cow_data.get('lact'),
            cow_data.get('clvd'),
            cow_data.get('rc'),
            cow_data.get('pen'),
            cow_data.get('frm'),
            cow_data.get('baseline_lact') if 'baseline_lact' in cow_data else cow_data.get('lact'),
        ))
        
        conn.commit()
        return cursor.lastrowid
    
    def get_cow_by_normalized_id(self, normalized_id: str) -> Optional[Dict[str, Any]]:
        """
        正規化された牛ID（左ゼロ除去）で牛を検索
        
        Args:
            normalized_id: 左ゼロを除去した牛ID（例: "980", "1"）
        
        Returns:
            牛の情報（辞書形式）、見つからない場合はNone
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # ltrim で左ゼロを除去して比較
        cursor.execute("""
            SELECT * FROM cow
            WHERE ltrim(cow_id, '0') = ?
            LIMIT 1
        """, (normalized_id,))
        
        row = cursor.fetchone()
        if row:
            return dict(row)
        return None
    
    def get_cow_by_id(self, cow_id: str) -> Optional[Dict[str, Any]]:
        """cow_id で牛を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow WHERE cow_id = ?", (cow_id,))
        row = cursor.fetchone()
        
        if row:
            return dict(row)
        return None
    
    def get_cows_by_id(self, cow_id: str) -> List[Dict[str, Any]]:
        """
        4桁のcow_idで牛を検索（複数件を返す）
        
        Args:
            cow_id: 4桁の牛ID（例: "0980"）
        
        Returns:
            牛の情報リスト（辞書形式）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow WHERE cow_id = ?", (cow_id,))
        rows = cursor.fetchall()
        
        return [dict(row) for row in rows]
    
    def get_cow_by_auto_id(self, auto_id: int) -> Optional[Dict[str, Any]]:
        """auto_id で牛を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow WHERE auto_id = ?", (auto_id,))
        row = cursor.fetchone()
        
        if row:
            return dict(row)
        return None
    
    def get_cows_by_jpn10(self, jpn10: str) -> List[Dict[str, Any]]:
        """
        JPN10（個体識別番号）で牛を検索
        
        Args:
            jpn10: 10桁の個体識別番号
        
        Returns:
            牛の情報リスト（辞書形式）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow WHERE jpn10 = ?", (jpn10,))
        rows = cursor.fetchall()
        
        return [dict(row) for row in rows]
    
    def update_cow(self, auto_id: int, cow_data: Dict[str, Any]):
        """
        牛データを更新（部分更新対応）
        
        Args:
            auto_id: 牛のauto_id
            cow_data: 更新するデータ（指定されたキーのみ更新）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # 更新可能なカラムのリスト
        updateable_columns = ['cow_id', 'jpn10', 'brd', 'bthd', 'entr', 'lact', 'clvd', 'rc', 'pen', 'frm', 'baseline_lact']
        
        # 更新するカラムと値を構築
        set_clauses = []
        values = []
        for col in updateable_columns:
            if col in cow_data:
                set_clauses.append(f"{col} = ?")
                values.append(cow_data[col])
        
        if not set_clauses:
            # 更新する項目がない場合は何もしない
            return
        
        # WHERE句にauto_idを追加
        values.append(auto_id)
        
        # SQLを構築
        sql = f"UPDATE cow SET {', '.join(set_clauses)} WHERE auto_id = ?"
        cursor.execute(sql, values)
        
        conn.commit()
    
    def get_all_cows(self) -> List[Dict[str, Any]]:
        """全牛を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow ORDER BY cow_id")
        return [dict(row) for row in cursor.fetchall()]
    
    def search_cows_by_id_prefix(self, prefix: str, limit: int = 50) -> List[Dict[str, Any]]:
        """
        牛IDの前方一致で牛を検索（複数件を返す）
        
        Args:
            prefix: 検索プレフィックス（例: "1", "12"）
            limit: 最大取得件数（デフォルト: 50）
        
        Returns:
            牛の情報リスト（辞書形式）、cow_id順にソート
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if not prefix:
            # 空文字の場合は空リストを返す
            return []
        
        # cow_idの前方一致で検索（左ゼロパディングを考慮）
        # 例: "1" で検索 → "0001", "0010", "0100", "1000" などにマッチ
        # 例: "12" で検索 → "0012", "0120", "1200" などにマッチ
        cursor.execute("""
            SELECT * FROM cow
            WHERE cow_id LIKE ? OR ltrim(cow_id, '0') LIKE ?
            ORDER BY cow_id
            LIMIT ?
        """, (f"{prefix}%", f"{prefix}%", limit))
        
        return [dict(row) for row in cursor.fetchall()]
    
    def delete_cow(self, auto_id: int):
        """
        牛を削除（物理削除）
        
        Args:
            auto_id: 牛の auto_id
        
        Note:
            外部キー制約により、関連するイベントとitem_valueも削除される
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # 関連するイベントを削除
        cursor.execute("DELETE FROM event WHERE cow_auto_id = ?", (auto_id,))
        
        # 関連するitem_valueを削除
        cursor.execute("DELETE FROM item_value WHERE cow_auto_id = ?", (auto_id,))
        
        # 個体を削除
        cursor.execute("DELETE FROM cow WHERE auto_id = ?", (auto_id,))
        
        conn.commit()
    
    # ========== event テーブル操作 ==========
    
    def insert_event(self, event_data: Dict[str, Any]) -> int:
        """
        イベントを追加
        
        Args:
            event_data: イベントデータ（cow_auto_id, event_number, event_date, json_data, note）
        
        Returns:
            追加された event id
        
        Note:
            DIMと産次は、RuleEngine.on_event_added()内でrecalculate_events_for_cow()により
            自動計算・保存されるため、ここでは計算しない
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        json_str = json.dumps(event_data.get('json_data'), ensure_ascii=False) if event_data.get('json_data') else None
        
        cursor.execute("""
            INSERT INTO event (cow_auto_id, event_number, event_date, json_data, note, deleted, event_dim, event_lact)
            VALUES (?, ?, ?, ?, ?, 0, NULL, NULL)
        """, (
            event_data.get('cow_auto_id'),
            event_data.get('event_number'),
            event_data.get('event_date'),
            json_str,
            event_data.get('note')
        ))
        
        conn.commit()
        return cursor.lastrowid
    
    def get_events_by_cow(self, cow_auto_id: int, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        牛のイベント一覧を取得（最新→過去順）
        
        Args:
            cow_auto_id: 牛の auto_id
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト（event_date DESC 順）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                WHERE cow_auto_id = ?
                ORDER BY event_date DESC, id DESC
            """, (cow_auto_id,))
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE cow_auto_id = ? AND deleted = 0
                ORDER BY event_date DESC, id DESC
            """, (cow_auto_id,))
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    # パースエラーの場合は空辞書に（ログ出力）
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}, raw={json_data_raw}")
                    event['json_data'] = {}
            else:
                # json_dataがNoneの場合は空辞書に
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def get_event_by_id(self, event_id: int, include_deleted: bool = False) -> Optional[Dict[str, Any]]:
        """
        イベントIDでイベントを取得
        
        Args:
            event_id: イベントID
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントデータ（辞書形式）、見つからない場合はNone
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("SELECT * FROM event WHERE id = ?", (event_id,))
        else:
            cursor.execute("SELECT * FROM event WHERE id = ? AND deleted = 0", (event_id,))
        
        row = cursor.fetchone()
        if row:
            event = dict(row)
            # json_data をパース
            if event.get('json_data'):
                try:
                    event['json_data'] = json.loads(event['json_data'])
                except:
                    event['json_data'] = {}
            return event
        
        return None
    
    def get_all_events(self, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        農場全体のイベント一覧を取得（最新→過去順）
        
        Args:
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト（event_date DESC 順）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                ORDER BY event_date DESC, id DESC
            """)
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE deleted = 0
                ORDER BY event_date DESC, id DESC
            """)
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    # パースエラーの場合は空辞書に（ログ出力）
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}, raw={json_data_raw}")
                    event['json_data'] = {}
            else:
                # json_dataがNoneの場合は空辞書に
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def get_events_by_number(self, event_number: int, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        イベント番号でイベントを取得
        
        Args:
            event_number: イベント番号
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_number = ?
                ORDER BY event_date DESC, id DESC
            """, (event_number,))
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_number = ? AND deleted = 0
                ORDER BY event_date DESC, id DESC
            """, (event_number,))
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}")
                    event['json_data'] = {}
            else:
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def get_events_by_note_prefix(self, prefix: str, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        note が指定プレフィックスで始まるイベントを取得（DC305取込一括削除などで使用）
        
        Args:
            prefix: note のプレフィックス（例: "DC305取込:"）
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト（event_date DESC 順）
        """
        conn = self.connect()
        cursor = conn.cursor()
        pattern = prefix + "%"
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                WHERE note LIKE ?
                ORDER BY event_date DESC, id DESC
            """, (pattern,))
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE note LIKE ? AND deleted = 0
                ORDER BY event_date DESC, id DESC
            """, (pattern,))
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}")
                    event['json_data'] = {}
            else:
                event['json_data'] = {}
            events.append(event)
        return events
    
    def get_events_by_number_and_period(self, event_number: int, start_date: str, end_date: str, 
                                        include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        イベント番号と期間でイベントを取得
        
        Args:
            event_number: イベント番号
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_number = ? AND event_date >= ? AND event_date <= ?
                ORDER BY event_date DESC, id DESC
            """, (event_number, start_date, end_date))
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_number = ? AND event_date >= ? AND event_date <= ? AND deleted = 0
                ORDER BY event_date DESC, id DESC
            """, (event_number, start_date, end_date))
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}")
                    event['json_data'] = {}
            else:
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def get_events_by_period(self, start_date: str, end_date: str, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        期間でイベントを取得
        
        Args:
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト（event_date DESC 順）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if include_deleted:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_date >= ? AND event_date <= ?
                ORDER BY event_date DESC, id DESC
            """, (start_date, end_date))
        else:
            cursor.execute("""
                SELECT * FROM event
                WHERE event_date >= ? AND event_date <= ? AND deleted = 0
                ORDER BY event_date DESC, id DESC
            """, (start_date, end_date))
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}")
                    event['json_data'] = {}
            else:
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def get_events_by_cow_id_and_period(self, cow_id: str, start_date: str, end_date: str, include_deleted: bool = False) -> List[Dict[str, Any]]:
        """
        個体IDと期間でイベントを取得
        
        Args:
            cow_id: 4桁の牛ID（例: "0980"）
            start_date: 開始日（YYYY-MM-DD形式）
            end_date: 終了日（YYYY-MM-DD形式）
            include_deleted: 削除済みを含むか
        
        Returns:
            イベントリスト（event_date DESC 順）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # まず、cow_idで牛を取得
        cursor.execute("SELECT auto_id FROM cow WHERE cow_id = ?", (cow_id,))
        cow_rows = cursor.fetchall()
        
        if not cow_rows:
            return []
        
        cow_auto_ids = [row[0] for row in cow_rows]
        
        # イベントを取得
        placeholders = ','.join(['?'] * len(cow_auto_ids))
        if include_deleted:
            cursor.execute(f"""
                SELECT e.* FROM event e
                WHERE e.cow_auto_id IN ({placeholders}) AND e.event_date >= ? AND e.event_date <= ?
                ORDER BY e.event_date DESC, e.id DESC
            """, cow_auto_ids + [start_date, end_date])
        else:
            cursor.execute(f"""
                SELECT e.* FROM event e
                WHERE e.cow_auto_id IN ({placeholders}) AND e.event_date >= ? AND e.event_date <= ? AND e.deleted = 0
                ORDER BY e.event_date DESC, e.id DESC
            """, cow_auto_ids + [start_date, end_date])
        
        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # json_data をパース
            json_data_raw = event.get('json_data')
            if json_data_raw:
                try:
                    event['json_data'] = json.loads(json_data_raw)
                except Exception as e:
                    logging.warning(f"json_data parse error for event_id={event.get('id')}: {e}")
                    event['json_data'] = {}
            else:
                event['json_data'] = {}
            events.append(event)
        
        return events
    
    def update_event(self, event_id: int, event_data: Dict[str, Any]):
        """
        イベントを更新（DIMと産次を自動再計算）
        
        Args:
            event_id: イベントID
            event_data: 更新するデータ（含まれるフィールドのみが更新される）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # 既存のイベントを取得
        existing_event = self.get_event_by_id(event_id, include_deleted=True)
        if not existing_event:
            return
        
        # 更新後の値を決定
        new_event_date = event_data.get('event_date', existing_event.get('event_date'))
        new_event_number = event_data.get('event_number', existing_event.get('event_number'))
        cow_auto_id = existing_event.get('cow_auto_id')
        
        # event_dataに含まれるフィールドのみを更新
        update_fields = []
        update_values = []
        
        if 'event_number' in event_data:
            update_fields.append('event_number = ?')
            update_values.append(event_data['event_number'])
        
        if 'event_date' in event_data:
            update_fields.append('event_date = ?')
            update_values.append(event_data['event_date'])
        
        if 'json_data' in event_data:
            update_fields.append('json_data = ?')
            json_str = json.dumps(event_data['json_data'], ensure_ascii=False) if event_data['json_data'] else None
            update_values.append(json_str)
        
        if 'note' in event_data:
            update_fields.append('note = ?')
            update_values.append(event_data.get('note'))
        
        # イベント日またはイベント番号が変更された場合は、DIMと産次をNULLにして
        # RuleEngineの再計算に任せる
        if 'event_date' in event_data or 'event_number' in event_data:
            update_fields.append('event_dim = NULL')
            update_fields.append('event_lact = NULL')
        
        if not update_fields:
            # 更新するフィールドがない場合は何もしない
            return
        
        # WHERE句の値を追加
        update_values.append(event_id)
        
        # UPDATEクエリを構築
        query = f"UPDATE event SET {', '.join(update_fields)} WHERE id = ?"
        cursor.execute(query, tuple(update_values))
        
        conn.commit()
    
    def delete_event(self, event_id: int, soft_delete: bool = True):
        """
        イベントを削除
        
        Args:
            event_id: イベントID
            soft_delete: True の場合は論理削除（deleted=1）、False の場合は物理削除
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        if soft_delete:
            cursor.execute("UPDATE event SET deleted = 1 WHERE id = ?", (event_id,))
        else:
            cursor.execute("DELETE FROM event WHERE id = ?", (event_id,))
        
        conn.commit()
    
    # ========== item_value テーブル操作 ==========
    
    def set_item_value(self, cow_auto_id: int, item_key: str, value: Any):
        """
        カスタムアイテムの値を設定
        
        Args:
            cow_auto_id: 牛の auto_id
            item_key: アイテムキー
            value: 値（文字列に変換して保存）
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        # 既存レコードを確認
        cursor.execute("""
            SELECT id FROM item_value
            WHERE cow_auto_id = ? AND item_key = ?
        """, (cow_auto_id, item_key))
        
        existing = cursor.fetchone()
        
        if existing:
            # 更新
            cursor.execute("""
                UPDATE item_value SET value = ?
                WHERE cow_auto_id = ? AND item_key = ?
            """, (str(value), cow_auto_id, item_key))
        else:
            # 新規追加
            cursor.execute("""
                INSERT INTO item_value (cow_auto_id, item_key, value)
                VALUES (?, ?, ?)
            """, (cow_auto_id, item_key, str(value)))
        
        conn.commit()
    
    def get_item_value(self, cow_auto_id: int, item_key: str) -> Optional[str]:
        """カスタムアイテムの値を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT value FROM item_value
            WHERE cow_auto_id = ? AND item_key = ?
        """, (cow_auto_id, item_key))
        
        row = cursor.fetchone()
        return row['value'] if row else None
    
    def get_all_item_values(self, cow_auto_id: int) -> Dict[str, str]:
        """牛の全カスタムアイテム値を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT item_key, value FROM item_value
            WHERE cow_auto_id = ?
        """, (cow_auto_id,))
        
        return {row['item_key']: row['value'] for row in cursor.fetchall()}
    
    def __enter__(self):
        """コンテキストマネージャー：with 文で使用"""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """コンテキストマネージャー：終了処理"""
        self.close()

