import datetime
import sqlite3
from pathlib import Path

SCHEMA_PATH = Path(__file__).parent / 'schema.sql'
BUSY_TIMEOUT_SECONDS = 5


def connect(db_path: str) -> sqlite3.Connection:
    """DB に接続する。テーブルが無ければ作成する"""
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(db_path, timeout=BUSY_TIMEOUT_SECONDS)
    connection.row_factory = sqlite3.Row
    # ネットワーク共有上では WAL が使えないため、従来のロールバックジャーナルを使う
    connection.execute('PRAGMA journal_mode=DELETE')
    schema_version = connection.execute('PRAGMA user_version').fetchone()[0]
    if schema_version == 0:
        connection.executescript(SCHEMA_PATH.read_text(encoding='utf-8'))
    elif schema_version == 1:
        migrate_to_version_2(connection)
    return connection


def migrate_to_version_2(connection: sqlite3.Connection) -> None:
    """版 1 の DB に削除フラグ(deleted 列)を追加する"""
    with connection:
        connection.execute('BEGIN IMMEDIATE')
        # 別の端末が先に移行していることがあるため、ロックを取ってから版を確かめる
        if connection.execute('PRAGMA user_version').fetchone()[0] == 1:
            connection.execute("ALTER TABLE documents ADD COLUMN deleted INTEGER NOT NULL DEFAULT 0")
            connection.execute('PRAGMA user_version = 2')


def now_text() -> str:
    """現在日時を DB 保存用の文字列にする"""
    return datetime.datetime.now().isoformat(timespec='seconds')
