from pathlib import Path

from services.csv_processor import (
    find_latest_csv,
    read_csv_with_encoding,
    process_csv_data,
    convert_date_format,
    process_completed_csv
)
from services.excel_processor import write_data_to_excel, open_and_sort_excel
from services.file_manager import backup_excel_file, cleanup_old_csv_files, ensure_directories_exist
from utils.config_manager import ConfigManager


def transfer_csv_to_excel() -> None:
    """ダウンロードフォルダからCSVファイルを読み込みExcelファイルに転記

    失敗時は例外を送出する(メッセージ表示は呼び出し側で行う)
    """
    config = ConfigManager()
    downloads_path = config.get_downloads_path()
    excel_path = config.get_excel_path()
    processed_dir = Path(config.get_processed_path())

    ensure_directories_exist()

    cleanup_old_csv_files(processed_dir)

    latest_csv = find_latest_csv(downloads_path)
    if not latest_csv:
        raise FileNotFoundError("ダウンロードフォルダにCSVファイルが見つかりません。")

    df = read_csv_with_encoding(latest_csv)
    if df is None:
        raise ValueError("CSVファイルの読み込みに失敗しました。")
    df = process_csv_data(df)
    df = convert_date_format(df)

    write_data_to_excel(excel_path, df)
    process_completed_csv(latest_csv)
    backup_excel_file(excel_path)
    open_and_sort_excel(excel_path)
