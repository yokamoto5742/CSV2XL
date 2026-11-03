import sqlite3

STAFF = 'staff'
EXCLUDE_DOCS = 'exclude_doc'
EXCLUDE_DOCTORS = 'exclude_doctor'


class ListRepository:
    """担当者名・除外する文書名・除外する医師名のリスト(list_items テーブル)の読み書き"""

    def __init__(self, connection: sqlite3.Connection) -> None:
        self.connection = connection

    def get_items(self, category: str) -> list[str]:
        rows = self.connection.execute(
            "SELECT name FROM list_items WHERE category = ? ORDER BY sort_order", (category,)
        ).fetchall()
        return [row['name'] for row in rows]

    def set_items(self, category: str, items: list[str]) -> None:
        """リストを置き換える(空白は除き、重複は1つにする)"""
        names = list(dict.fromkeys(item.strip() for item in items if item.strip()))
        with self.connection:
            self.connection.execute("DELETE FROM list_items WHERE category = ?", (category,))
            self.connection.executemany(
                "INSERT INTO list_items (category, name, sort_order) VALUES (?, ?, ?)",
                [(category, name, order) for order, name in enumerate(names)]
            )
