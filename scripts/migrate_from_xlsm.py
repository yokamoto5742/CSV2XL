"""医療文書担当一覧.xlsm のデータを SQLite に移す(最初に1回だけ実行する)

使い方:
    uv run python -m scripts.migrate_from_xlsm <xlsmのパス> [DBのパス]

DB のパスを省略すると config.ini の [Paths] database_path を使う。
担当一覧シートの全行・設定シートの担当者・config.ini の除外リストを移す。
"""
import datetime
import sqlite3
import sys
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from services.database import connect, now_text
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository
from services.validators import parse_date
from utils.config_manager import ConfigManager

DOCUMENT_SHEET = "担当一覧"
SETTINGS_SHEET = "設定"
STAFF_LABEL = "担当"


def to_date_text(value: Any) -> str | None:
    """セルの日付を 'YYYY-MM-DD' にする。日付として読めない文字列はそのまま返す"""
    if value is None or value == '':
        return None
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    try:
        return parse_date(str(value))
    except ValueError:
        return str(value)


def to_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def read_document_rows(workbook: Workbook) -> list[tuple[Any, ...]]:
    """担当一覧シートの A〜I 列を DB に入れる形で読む(A 列が空の行で終わり)"""
    rows = []
    sheet = workbook[DOCUMENT_SHEET]
    for row in sheet.iter_rows(min_row=2, max_col=9, values_only=True):
        if row[0] is None:
            break
        rows.append((
            to_date_text(row[0]), int(str(row[1]).replace(',', '')), to_text(row[2]) or '',
            to_text(row[3]) or '', to_text(row[4]) or '', to_text(row[5]) or '',
            to_text(row[6]), to_date_text(row[7]), to_text(row[8]),
        ))
    return rows


def read_staff_names(workbook: Workbook) -> list[str]:
    """設定シートの「担当」の行(B 列はカンマ区切りの担当者名)を読む"""
    if SETTINGS_SHEET not in workbook.sheetnames:
        return []
    for row in workbook[SETTINGS_SHEET].iter_rows(max_col=2, values_only=True):
        if row[0] == STAFF_LABEL and row[1]:
            return [name.strip() for name in str(row[1]).split(',') if name.strip()]
    return []


def migrate(connection: sqlite3.Connection, workbook: Workbook, config: ConfigManager) -> int:
    """移行して移した行数を返す。DB に行があるときは二重登録を防ぐため中止する"""
    if connection.execute("SELECT COUNT(*) FROM documents").fetchone()[0] > 0:
        raise RuntimeError("データベースにすでに行があるため、移行を中止しました。")

    rows = read_document_rows(workbook)
    now = now_text()
    with connection:
        connection.executemany(
            "INSERT INTO documents (received_date, patient_id, patient_name, document_name, department,"
            " doctor_name, staff_name, requested_date, note, created_at, updated_at)"
            " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [(*row, now, now) for row in rows]
        )

    lists = ListRepository(connection)
    for category, items in ((STAFF, read_staff_names(workbook)),
                            (EXCLUDE_DOCS, config.get_list('ExcludeDocs')),
                            (EXCLUDE_DOCTORS, config.get_list('ExcludeDoctors'))):
        if not lists.get_items(category):
            lists.set_items(category, items)
    return len(rows)


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        print(__doc__)
        sys.exit(1)
    config = ConfigManager()
    db_path = argv[2] if len(argv) > 2 else config.get_database_path()
    workbook = load_workbook(argv[1], read_only=True, data_only=True)
    connection = connect(db_path)
    try:
        count = migrate(connection, workbook, config)
        lists = ListRepository(connection)
        print(f"移行先: {db_path}")
        print(f"担当一覧: {count}件")
        print(f"担当者名: {lists.get_items(STAFF)}")
        print(f"除外する文書名: {len(lists.get_items(EXCLUDE_DOCS))}件")
        print(f"除外する医師名: {len(lists.get_items(EXCLUDE_DOCTORS))}件")
    finally:
        connection.close()
        workbook.close()


if __name__ == "__main__":
    main(sys.argv)
