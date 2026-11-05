import sqlite3
from pathlib import Path

from services.database import connect


class TestConnect:
    def test_creates_schema(self, tmp_path: Path) -> None:
        """新しい DB にテーブルを作り、スキーマのバージョンを設定する"""
        conn = connect(str(tmp_path / 'sub' / 'new.db'))
        try:
            tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
            assert {'documents', 'list_items'} <= tables
            assert conn.execute('PRAGMA user_version').fetchone()[0] == 2
            assert conn.execute('PRAGMA journal_mode').fetchone()[0] == 'delete'
        finally:
            conn.close()

    def test_keeps_existing_data(self, tmp_path: Path) -> None:
        """既存の DB を開いてもデータを消さない"""
        db_path = str(tmp_path / 'existing.db')
        conn = connect(db_path)
        with conn:
            conn.execute("INSERT INTO list_items (category, name, sort_order) VALUES ('staff', '山本', 0)")
        conn.close()

        conn = connect(db_path)
        try:
            assert conn.execute("SELECT name FROM list_items").fetchall()[0][0] == '山本'
        finally:
            conn.close()

    def test_migrates_version_1(self, tmp_path: Path) -> None:
        """版 1 の DB には deleted 列を追加し、既存の行は未削除のまま残す"""
        db_path = str(tmp_path / 'v1.db')
        conn = sqlite3.connect(db_path)
        conn.executescript(
            "CREATE TABLE documents (id INTEGER PRIMARY KEY, received_date TEXT NOT NULL);"
            "INSERT INTO documents (received_date) VALUES ('2025-01-10');"
            "PRAGMA user_version = 1;"
        )
        conn.close()

        conn = connect(db_path)
        try:
            assert conn.execute('PRAGMA user_version').fetchone()[0] == 2
            assert conn.execute("SELECT deleted FROM documents").fetchone()[0] == 0
        finally:
            conn.close()

        # 移行済みの DB を開き直しても列を二重に追加しない
        connect(db_path).close()

    def test_row_factory(self, connection: sqlite3.Connection) -> None:
        """列名で値を取り出せる"""
        row = connection.execute("SELECT 1 AS value").fetchone()
        assert row['value'] == 1
