import sys
import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path

import polars as pl
from PyQt6.QtWidgets import QApplication, QMessageBox

from services.csv_excel_transfer import transfer_csv_to_excel
from services.csv_processor import process_csv_data
from utils.config_manager import ConfigManager


@pytest.fixture
def app():
    """テスト用のQApplicationを提供するフィクスチャ"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


class TestCsvExcelTransfer:
    @patch('services.csv_excel_transfer.ConfigManager')
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
                                           mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """CSVからExcelへの正常な転送処理のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 各関数のモック戻り値設定
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = "mock_dataframe"
        mock_convert_date.return_value = "mock_dataframe_with_date"
        mock_process_data.return_value = "mock_processed_dataframe"
        mock_write.return_value = True  # 書き込み成功

        # 関数実行
        transfer_csv_to_excel()

        # 各関数が正しく呼ばれたことを確認
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

    @patch('services.csv_excel_transfer.ConfigManager')
    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.QMessageBox.warning')
    def test_transfer_csv_to_excel_no_csv(self, mock_warning, mock_find_csv,
                                          mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """CSVファイルが見つからない場合のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # CSVファイルが見つからない場合
        mock_find_csv.return_value = None

        # 関数実行
        transfer_csv_to_excel()

        # 警告が表示されることを確認
        mock_warning.assert_called_once()
        args = mock_warning.call_args[0]
        assert args[1] == "警告"
        assert "CSVファイルが見つかりません" in args[2]

    @patch('services.csv_excel_transfer.ConfigManager')
    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.read_csv_with_encoding')
    @patch('services.csv_excel_transfer.convert_date_format')
    @patch('services.csv_excel_transfer.process_csv_data')
    @patch('services.csv_excel_transfer.write_data_to_excel')
    @patch('services.csv_excel_transfer.QMessageBox.critical')
    def test_transfer_csv_to_excel_write_error(self, mock_critical, mock_write, mock_process_data,
                                               mock_convert_date, mock_read_csv, mock_find_csv,
                                               mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """Excel書き込みエラーのテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 各関数のモック戻り値設定
        mock_find_csv.return_value = "C:/Downloads/test.csv"
        mock_read_csv.return_value = "mock_dataframe"
        mock_convert_date.return_value = "mock_dataframe_with_date"
        mock_process_data.return_value = "mock_processed_dataframe"
        mock_write.return_value = False  # 書き込み失敗

        # 関数実行
        transfer_csv_to_excel()

        # バックアップや後続処理が呼ばれないことを確認
        assert not mock_critical.called

    @patch('services.csv_excel_transfer.ConfigManager')
    @patch('services.csv_excel_transfer.ensure_directories_exist')
    @patch('services.csv_excel_transfer.cleanup_old_csv_files')
    @patch('services.csv_excel_transfer.find_latest_csv')
    @patch('services.csv_excel_transfer.QMessageBox.critical')
    def test_transfer_csv_to_excel_exception(self, mock_critical, mock_find_csv,
                                             mock_cleanup, mock_ensure_dirs, mock_config_manager, app):
        """例外発生時のテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_downloads_path.return_value = "C:/Downloads"
        mock_config.get_excel_path.return_value = "C:/Excel/test.xlsm"
        mock_config.get_processed_path.return_value = "C:/Processed"
        mock_config_manager.return_value = mock_config

        # 例外を発生させる
        mock_find_csv.side_effect = Exception("テストエラー")

        # 関数実行
        transfer_csv_to_excel()

        # エラーメッセージが表示されることを確認
        mock_critical.assert_called_once()
        args = mock_critical.call_args[0]
        assert args[1] == "エラー"
        assert "テストエラー" in args[2]


class TestCsvProcessor:
    """CSV処理機能のテストクラス"""
    
    def create_test_dataframe(self):
        """テスト用のDataFrameを作成（実際のCSVの11列構造を模擬）"""
        return pl.DataFrame({
            "患者ID": [1, 2, 3, 4],
            "col1": ["A", "B", "C", "D"],
            "col2": ["E", "F", "G", "H"],
            "col3": ["I", "J", "K", "L"],
            "文書名": ["検査結果 A *", "診断書 B", "処方箋 C", "除外文書*"],  # スペースとアスタリスクを含むデータ
            "日付": ["20240101", "20240102", "20240103", "20240104"],
            "医師名": ["田中 医師 *", "除外 医師", "佐藤医師", "鈴木 医師"],  # スペースとアスタリスクを含むデータ
            "col7": ["内容1", "内容2", "内容3", "内容4"],
            "col8": ["削除予定", "削除予定", "削除予定", "削除予定"],  # K列（インデックス8）
            "col9": ["データ", "データ", "データ", "データ"],
            "col10": ["削除予定", "削除予定", "削除予定", "削除予定"]  # I列（インデックス10）
        })
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_no_exclusions(self, mock_config_manager):
        """除外設定なしでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = []
        mock_config.get_exclude_doctors.return_value = []
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外されていないこと）
        assert len(result_df) == 4  # 全ての行が残る
        # 処理後は文書名がcol_4_文書名（インデックス1）、医師名がcol_6_医師名（インデックス3）
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        # スペースとアスタリスクが除去されることを確認
        doc_values = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_values = result_df[result_df.columns[doctor_col_idx]].to_list()
        assert "除外文書" in doc_values  # *が除去された後
        assert "除外医師" in doctor_values  # スペースが除去された後
        # 具体的な除去が行われていることを確認
        assert "田中医師" in doctor_values  # スペースとアスタリスクが除去されている
        assert "検査結果A" in doc_values  # スペースとアスタリスクが除去されている
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_doc_exclusions(self, mock_config_manager):
        """文書除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["除外文書"]
        mock_config.get_exclude_doctors.return_value = []
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外文書の行が削除されていること）
        assert len(result_df) == 3
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        assert "除外文書" not in result_df[result_df.columns[doc_col_idx]].to_list()
        assert "除外医師" in result_df[result_df.columns[doctor_col_idx]].to_list()  # 医師は残る
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_doctor_exclusions(self, mock_config_manager):
        """医師除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = []
        mock_config.get_exclude_doctors.return_value = ["除外医師"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（除外医師の行が削除されていること）
        assert len(result_df) == 3
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        assert "除外文書" in result_df[result_df.columns[doc_col_idx]].to_list()  # 文書は残る
        assert "除外医師" not in result_df[result_df.columns[doctor_col_idx]].to_list()
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_with_both_exclusions(self, mock_config_manager):
        """文書と医師両方の除外設定ありでのCSVデータ処理テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["除外文書"]
        mock_config.get_exclude_doctors.return_value = ["除外医師"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（両方の除外条件に該当する行が削除されていること）
        assert len(result_df) == 2
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        assert "除外文書" not in result_df[result_df.columns[doc_col_idx]].to_list()
        assert "除外医師" not in result_df[result_df.columns[doctor_col_idx]].to_list()
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_partial_match_exclusions(self, mock_config_manager):
        """部分一致での除外機能テスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["検査"]  # 「検査結果 A」に部分一致
        mock_config.get_exclude_doctors.return_value = ["田中"]  # 「田中医師」に部分一致
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（部分一致で除外されていること）
        assert len(result_df) == 3  # "検査結果 A"と"田中医師"の行が除外される
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        doc_list = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_list = result_df[result_df.columns[doctor_col_idx]].to_list()
        
        # 「検査」を含む文書が除外されていることを確認
        assert not any("検査" in doc for doc in doc_list)
        # 「田中」を含む医師が除外されていることを確認
        assert not any("田中" in doctor for doctor in doctor_list)
    
    @patch('services.csv_processor.ConfigManager')
    def test_process_csv_data_multiple_exclusions(self, mock_config_manager):
        """複数の除外条件でのテスト"""
        # ConfigManagerのモック設定
        mock_config = MagicMock()
        mock_config.get_exclude_docs.return_value = ["検査", "診断"]
        mock_config.get_exclude_doctors.return_value = ["田中", "佐藤"]
        mock_config_manager.return_value = mock_config
        
        # テストデータ作成
        df = self.create_test_dataframe()
        
        # 処理実行
        result_df = process_csv_data(df)
        
        # 結果確認（複数の除外条件に該当する行が全て削除されていること）
        assert len(result_df) == 1  # 「処方箋 C」と「佐藤医師」の組み合わせも除外される
        doc_col_idx = 1  # col_4_文書名
        doctor_col_idx = 3  # col_6_医師名
        doc_list = result_df[result_df.columns[doc_col_idx]].to_list()
        doctor_list = result_df[result_df.columns[doctor_col_idx]].to_list()
        
        # 除外条件に該当しない行のみ残っていることを確認
        assert "除外文書" in doc_list
        assert "鈴木医師" in doctor_list
