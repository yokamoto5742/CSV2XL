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
            assert conn.execute('PRAGMA user_version').fetchone()[0] == 1
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

    def test_row_factory(self, connection: sqlite3.Connection) -> None:
        """列名で値を取り出せる"""
        row = connection.execute("SELECT 1 AS value").fetchone()
        assert row['value'] == 1
