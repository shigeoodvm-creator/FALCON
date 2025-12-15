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
        """データベースファイルが存在しない場合は作成"""
        if not self.db_path.exists():
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self.create_tables()
    
    def connect(self) -> sqlite3.Connection:
        """データベースに接続"""
        if self.conn is None:
            self.conn = sqlite3.connect(str(self.db_path))
            self.conn.row_factory = sqlite3.Row  # 辞書形式で取得
            # 外部キー制約を有効化
            self.conn.execute("PRAGMA foreign_keys = ON")
        return self.conn
    
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
                FOREIGN KEY(cow_auto_id) REFERENCES cow(auto_id)
            )
        """)
        
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
            INSERT INTO cow (cow_id, jpn10, brd, bthd, entr, lact, clvd, rc, pen, frm)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            cow_data.get('frm')
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
    
    def update_cow(self, auto_id: int, cow_data: Dict[str, Any]):
        """牛データを更新"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("""
            UPDATE cow SET
                cow_id = ?, jpn10 = ?, brd = ?, bthd = ?, entr = ?,
                lact = ?, clvd = ?, rc = ?, pen = ?, frm = ?
            WHERE auto_id = ?
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
            auto_id
        ))
        
        conn.commit()
    
    def get_all_cows(self) -> List[Dict[str, Any]]:
        """全牛を取得"""
        conn = self.connect()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM cow ORDER BY cow_id")
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
        """
        conn = self.connect()
        cursor = conn.cursor()
        
        json_str = json.dumps(event_data.get('json_data'), ensure_ascii=False) if event_data.get('json_data') else None
        
        cursor.execute("""
            INSERT INTO event (cow_auto_id, event_number, event_date, json_data, note, deleted)
            VALUES (?, ?, ?, ?, ?, 0)
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
    
    def update_event(self, event_id: int, event_data: Dict[str, Any]):
        """イベントを更新"""
        conn = self.connect()
        cursor = conn.cursor()
        
        json_str = json.dumps(event_data.get('json_data'), ensure_ascii=False) if event_data.get('json_data') else None
        
        cursor.execute("""
            UPDATE event SET
                event_number = ?, event_date = ?, json_data = ?, note = ?
            WHERE id = ?
        """, (
            event_data.get('event_number'),
            event_data.get('event_date'),
            json_str,
            event_data.get('note'),
            event_id
        ))
        
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

