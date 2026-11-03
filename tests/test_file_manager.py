import sqlite3
from contextlib import closing

import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from services.file_manager import backup_database, cleanup_old_csv_files, ensure_directories_exist


class TestFileManager:
    @patch('services.file_manager.ConfigManager')
    def test_backup_database(self, mock_config_manager, connection: sqlite3.Connection, tmp_path: Path):
        """バックアップフォルダを作成し、DB の内容をバックアップファイルに写す"""
        backup_dir = tmp_path / 'backup'
        mock_config_manager.return_value.get_backup_path.return_value = str(backup_dir)
        with connection:
            connection.execute("INSERT INTO list_items (category, name, sort_order) VALUES ('staff', '山本', 0)")

        backup_path = backup_database(connection)

        assert backup_path.parent == backup_dir
        assert backup_path.name.startswith('医療文書担当一覧_') and backup_path.suffix == '.db'
        with closing(sqlite3.connect(backup_path)) as backup_connection:
            assert backup_connection.execute("SELECT name FROM list_items").fetchone()[0] == '山本'

    @patch('services.file_manager.ConfigManager')
    @patch('services.file_manager.datetime.datetime')
    def test_cleanup_old_csv_files(self, mock_dt, mock_config_manager):
        """古いCSVファイルの削除テスト"""
        # ConfigManagerのモック設定（保持期間14日）
        mock_config = MagicMock()
        mock_config.get_backup_retention_days.return_value = 14
        mock_config_manager.return_value = mock_config

        # 現在時刻を固定
        current_time = MagicMock()
        current_time.now = MagicMock(return_value=current_time)
        mock_dt.now.return_value = current_time

        # モックディレクトリと3つのファイル
        processed_dir = MagicMock(spec=Path)

        file1 = MagicMock()  # 新しいファイル
        file2 = MagicMock()  # 古いファイル
        file3 = MagicMock()  # 非CSVファイル

        file1.suffix = '.csv'
        file2.suffix = '.csv'
        file3.suffix = '.txt'

        # glob()がCSVファイルだけを返すように設定（実際の動作に合わせる）
        processed_dir.glob.return_value = [file1, file2]

        # fromtimestampメソッドをモック
        file1_time = MagicMock()
        file2_time = MagicMock()

        # 日付比較の結果を設定
        diff1 = MagicMock()
        diff1.days = 10  # 10日前（削除しない）

        diff2 = MagicMock()
        diff2.days = 15  # 15日前（削除する - 保持期間14日以上）

        # 引数に応じて異なる日付オブジェクトを返す
        def from_timestamp_side_effect(timestamp):
            if timestamp is file1.stat().st_mtime:
                return file1_time
            elif timestamp is file2.stat().st_mtime:
                return file2_time
            return MagicMock()

        mock_dt.fromtimestamp.side_effect = from_timestamp_side_effect

        # 日付の差分計算結果を設定
        current_time.__sub__.side_effect = lambda other: diff1 if other is file1_time else diff2

        # 関数を実行
        cleanup_old_csv_files(processed_dir)

        # ConfigManagerが呼ばれたことを確認
        mock_config_manager.assert_called_once()
        mock_config.get_backup_retention_days.assert_called_once()

        # 古いCSVファイル（14日以上前）のみが削除されることを確認
        file1.unlink.assert_not_called()  # 新しいファイルは削除されない
        file2.unlink.assert_called_once()  # 古いファイルは削除される
        file3.unlink.assert_not_called()  # 非CSVファイルは削除されない

    @patch('services.file_manager.ConfigManager')
    def test_ensure_directories_exist(self, mock_config_manager):
        """必要なディレクトリの存在確認と作成テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = 'C:/Downloads'
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config.get_processed_path.return_value = 'C:/Processed'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定
        with patch('services.file_manager.Path') as mock_path:
            # 各ディレクトリの存在状態をモック
            mock_downloads = MagicMock()
            mock_downloads.exists.return_value = True  # 既に存在する

            mock_backup = MagicMock()
            mock_backup.exists.return_value = False  # 存在しない

            mock_processed = MagicMock()
            mock_processed.exists.return_value = False  # 存在しない

            # Pathのインスタンス化の戻り値を設定
            mock_path.side_effect = [mock_downloads, mock_backup, mock_processed]

            # 関数実行
            ensure_directories_exist()

            # 結果を確認
            mock_config_manager.assert_called_once()
            mock_config.get_downloads_path.assert_called_once()
            mock_config.get_backup_path.assert_called_once()
            mock_config.get_processed_path.assert_called_once()

            # 存在するディレクトリは作成されないことを確認
            mock_downloads.mkdir.assert_not_called()

            # 存在しないディレクトリは作成されることを確認
            mock_backup.mkdir.assert_called_once_with(parents=True, exist_ok=True)
            mock_processed.mkdir.assert_called_once_with(parents=True, exist_ok=True)

    @patch('services.file_manager.ConfigManager')
    def test_ensure_directories_exist_exception(self, mock_config_manager):
        """ディレクトリ作成時の例外処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = 'C:/Downloads'
        mock_config.get_backup_path.return_value = 'C:/Backup'
        mock_config.get_processed_path.return_value = 'C:/Processed'
        mock_config_manager.return_value = mock_config

        # Pathのモック設定
        with patch('services.file_manager.Path') as mock_path:
            # ディレクトリ作成で例外を発生させる
            mock_dir = MagicMock()
            mock_dir.exists.return_value = False
            mock_dir.mkdir.side_effect = PermissionError("アクセス拒否")

            # Pathのインスタンス化の戻り値を設定（すべて同じモックを返す）
            mock_path.return_value = mock_dir

            # 関数実行（例外が発生するはず）
            with pytest.raises(Exception) as excinfo:
                ensure_directories_exist()

            # 例外内容を確認
            # 元のコードが例外をラップしている場合はこれに合わせる
            assert isinstance(excinfo.value, Exception)
            # エラーメッセージにはPermissionErrorの内容が含まれるはず
            assert "アクセス拒否" in str(excinfo.value)
