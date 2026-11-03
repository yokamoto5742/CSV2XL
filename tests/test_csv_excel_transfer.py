import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path

from services.csv_excel_transfer import transfer_csv_to_excel


@pytest.fixture
def mock_config_manager():
    """ConfigManagerをモックしてパスを固定するフィクスチャ"""
    with patch('services.csv_excel_transfer.ConfigManager') as mock_cls:
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_cls.return_value = mock_config
        yield mock_cls


@pytest.mark.usefixtures("mock_config_manager")
class TestCsvExcelTransfer:
    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.read_csv_with_encoding')
    @patch('services.csv_excel_transfer.convert_date_format')
    @patch('services.csv_excel_transfer.process_csv_data')
    @patch('services.csv_excel_transfer.write_data_to_excel')
    @patch('services.csv_excel_transfer.backup_excel_file')
    @patch('services.csv_excel_transfer.process_completed_csv')
    @patch('services.csv_excel_transfer.open_and_sort_excel')
    def test_transfer_csv_to_excel_success(self, mock_open_sort, mock_process_csv,
                                           mock_backup, mock_write, mock_process_data,
                                           mock_convert_date, mock_read_csv, mock_find_csv,
                                           mock_cleanup, mock_ensure_dirs):
        """CSVからExcelへの正常な転送処理のテスト"""
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = "mock_dataframe"
        mock_convert_date.return_value = "mock_dataframe_with_date"
        mock_process_data.return_value = "mock_processed_dataframe"

        transfer_csv_to_excel()

        mock_ensure_dirs.assert_called_once()
        mock_cleanup.assert_called_once_with(Path("C:/Processed"))
        mock_find_csv.assert_called_once_with("C:/Downloads")
        mock_read_csv.assert_called_once_with("C:/Downloads/test.csv")
        mock_process_data.assert_called_once_with("mock_dataframe")
        mock_convert_date.assert_called_once_with("mock_processed_dataframe")
        mock_write.assert_called_once_with("C:/Excel/test.xlsm", "mock_dataframe_with_date")
        mock_backup.assert_called_once_with("C:/Excel/test.xlsm")
        mock_process_csv.assert_called_once_with("C:/Downloads/test.csv")
        mock_open_sort.assert_called_once_with("C:/Excel/test.xlsm")

    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    def test_transfer_csv_to_excel_no_csv(self, mock_find_csv, mock_cleanup, mock_ensure_dirs):
        """CSVファイルが見つからない場合は例外を送出する"""
        mock_find_csv.return_value = None

        with pytest.raises(FileNotFoundError, match="CSVファイルが見つかりません"):
            transfer_csv_to_excel()

    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.read_csv_with_encoding')
    def test_transfer_csv_to_excel_read_failure(self, mock_read_csv, mock_find_csv,
                                                mock_cleanup, mock_ensure_dirs):
        """CSVファイルを読み込めない場合は例外を送出する"""
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = None

        with pytest.raises(ValueError, match="CSVファイルの読み込みに失敗しました"):
            transfer_csv_to_excel()

    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.read_csv_with_encoding')
    @patch('services.csv_excel_transfer.convert_date_format')
    @patch('services.csv_excel_transfer.process_csv_data')
    @patch('services.csv_excel_transfer.write_data_to_excel')
    @patch('services.csv_excel_transfer.backup_excel_file')
    @patch('services.csv_excel_transfer.process_completed_csv')
    def test_transfer_csv_to_excel_write_error(self, mock_process_csv, mock_backup, mock_write,
                                               mock_process_data, mock_convert_date, mock_read_csv,
                                               mock_find_csv, mock_cleanup, mock_ensure_dirs):
        """Excel書き込みに失敗したら後続処理を行わずに例外を送出する"""
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_write.side_effect = PermissionError("別のプロセスで開かれています")

        with pytest.raises(PermissionError):
            transfer_csv_to_excel()

        mock_process_csv.assert_not_called()
        mock_backup.assert_not_called()
