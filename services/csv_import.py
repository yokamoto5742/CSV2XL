import sqlite3
from pathlib import Path

from services.csv_processor import (
    convert_date_format,
    dataframe_to_documents,
    find_latest_csv,
    process_completed_csv,
    process_csv_data,
    read_csv_with_encoding,
)
from services.document_repository import DocumentRepository
from services.file_manager import backup_database, cleanup_old_csv_files, ensure_directories_exist
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, ListRepository
from utils.config_manager import ConfigManager


def import_csv_to_database(connection: sqlite3.Connection) -> int:
    """ダウンロードフォルダの最新 CSV を DB に取り込み、追加した件数を返す

    失敗時は例外を送出する(メッセージ表示は呼び出し側で行う)
    """
    config = ConfigManager()
    downloads_path = config.get_downloads_path()
    processed_dir = Path(config.get_processed_path())

    ensure_directories_exist()

    cleanup_old_csv_files(processed_dir)

    latest_csv = find_latest_csv(downloads_path)
    if not latest_csv:
        raise FileNotFoundError("ダウンロードフォルダにCSVファイルが見つかりません。")

    df = read_csv_with_encoding(latest_csv)
    if df is None:
        raise ValueError("CSVファイルの読み込みに失敗しました。")
    lists = ListRepository(connection)
    df = process_csv_data(df, lists.get_items(EXCLUDE_DOCS), lists.get_items(EXCLUDE_DOCTORS))
    df = convert_date_format(df)

    added_count = DocumentRepository(connection).add_new_documents(dataframe_to_documents(df))
    process_completed_csv(latest_csv)
    backup_database(connection)
    return added_count
