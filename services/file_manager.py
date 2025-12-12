import shutil
import datetime
from pathlib import Path

from utils.config_manager import ConfigManager


def backup_excel_file(excel_path: str) -> None:
    """Excelファイルのバックアップを作成

    Args:
        excel_path: バックアップ対象のExcelファイルパス
    """
    config = ConfigManager()
    backup_dir = Path(config.get_backup_path())

    if not backup_dir.exists():
        backup_dir.mkdir(parents=True)

    # 現在の日時を取得してファイル名を生成
    timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M')
    backup_filename = f"医療文書担当一覧_{timestamp}.xlsm"
    backup_path = backup_dir / backup_filename

    try:
        shutil.copy2(excel_path, backup_path)
        print(f"バックアップを作成しました: {backup_path}")
    except Exception as e:
        print(f"バックアップ作成中にエラーが発生しました: {str(e)}")
        raise


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
    for file in backup_dir.glob("医療文書担当一覧_*.xlsm"):
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
            try:
                directory.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"ディレクトリの作成中にエラーが発生しました: {directory} - {str(e)}")
                raise
