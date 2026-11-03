import datetime
import sqlite3
from contextlib import closing
from pathlib import Path

from utils.config_manager import ConfigManager


def backup_database(connection: sqlite3.Connection) -> Path:
    """DB のバックアップをバックアップフォルダに作成し、作成したファイルのパスを返す

    ファイルコピーではなく SQLite の backup API を使い、書き込み途中の状態を写さないようにする
    """
    config = ConfigManager()
    backup_dir = Path(config.get_backup_path())
    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M')
    backup_path = backup_dir / f"医療文書担当一覧_{timestamp}.db"
    with closing(sqlite3.connect(backup_path)) as backup_connection:
        connection.backup(backup_connection)
    print(f"バックアップを作成しました: {backup_path}")
    return backup_path


def cleanup_old_csv_files(processed_dir: Path) -> None:
    """指定した日数より前の処理済みCSVファイルを削除

    Args:
        processed_dir: 処理済みCSVファイルの格納ディレクトリ
    """
    config = ConfigManager()
    retention_days = config.get_backup_retention_days()
    current_time = datetime.datetime.now()
    for file in processed_dir.glob("*.csv"):
        file_time = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        if (current_time - file_time).days >= retention_days:
            try:
                file.unlink()
            except Exception as e:
                print(f"ファイル削除中にエラーが発生しました: {file} - {str(e)}")


def cleanup_old_backup_files() -> None:
    """指定した日数より前のバックアップファイルを削除"""
    config = ConfigManager()
    backup_dir = Path(config.get_backup_path())
    retention_days = config.get_backup_retention_days()

    if not backup_dir.exists():
        return

    current_time = datetime.datetime.now()
    for file in backup_dir.glob("医療文書担当一覧_*.db"):
        file_time = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        if (current_time - file_time).days >= retention_days:
            try:
                file.unlink()
                print(f"古いバックアップを削除しました: {file}")
            except Exception as e:
                print(f"バックアップ削除中にエラーが発生しました: {file} - {str(e)}")


def ensure_directories_exist() -> None:
    """設定に指定されたディレクトリが存在しない場合は作成

    ダウンロード、バックアップ、処理済みCSVディレクトリを確認・作成
    """
    config = ConfigManager()
    directories = [
        Path(config.get_downloads_path()),
        Path(config.get_backup_path()),
        Path(config.get_processed_path())
    ]

    for directory in directories:
        if not directory.exists():
            directory.mkdir(parents=True, exist_ok=True)
