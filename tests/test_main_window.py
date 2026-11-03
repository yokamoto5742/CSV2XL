import sys
import pytest
import configparser
from unittest.mock import patch, MagicMock

from PyQt6.QtWidgets import QApplication, QPushButton

from app.main_window import MainWindow
from utils.config_manager import ConfigManager, CONFIG_PATH
from app import __version__


def restore_config(config, original_config):
    """configを元の状態に復元するヘルパーメソッド"""
    for section in config.sections():
        config.remove_section(section)
    for section in original_config.sections():
        if not config.has_section(section):
            config.add_section(section)
        for key, value in original_config[section].items():
            config[section][key] = value


@pytest.fixture
def app():
    """テスト用のQApplicationを提供するフィクスチャ"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


@pytest.fixture
def backup_config():
    """設定ファイルのバックアップと復元を行うフィクスチャ"""
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH, encoding='utf-8')
    # 現在の設定内容を保存
    original_config = configparser.ConfigParser()
    for section in config.sections():
        original_config.add_section(section)
        for key, value in config[section].items():
            original_config[section][key] = value

    yield original_config

    # テスト後に設定を元に戻す
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH, encoding='utf-8')
    restore_config(config, original_config)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        config.write(f)


@pytest.fixture
def db_connect():
    """本物の DB に接続しないよう connect をモックする"""
    with patch('app.main_window.connect') as mock_connect:
        yield mock_connect


class TestMainWindow:
    def test_init(self, app, backup_config):
        """MainWindowの初期化テスト"""
        window = MainWindow()

        # ウィンドウタイトルを確認
        assert window.windowTitle() == f"CSV取込アプリ v{__version__}"

        # 設定に基づくフォントサイズとウィンドウサイズの確認
        config = ConfigManager()
        font_size = config.get_font_size()
        assert window.font().pointSize() == font_size

        window_size = config.get_window_size()
        assert window.width() == window_size[0]
        assert window.height() == window_size[1]

        # 必要なウィジェットが存在することを確認
        central_widget = window.centralWidget()
        assert central_widget is not None

        # レイアウト内のボタンの数を確認
        layout = central_widget.layout()
        assert layout is not None

        # (タイトルラベル、CSVボタン、担当一覧ボタン、設定ラベル、除外文書ボタン、除外医師ボタン、
        #  担当者名ボタン、外観ボタン、フォルダボタン、閉じるボタン)
        assert layout.count() == 10

        # 起動時には DB に接続しない
        assert window.connection is None

    @patch('app.main_window.MainWindow.show_document_window')
    @patch('app.main_window.QMessageBox.information')
    @patch('app.main_window.import_csv_to_database', return_value=3)
    def test_import_csv_success(self, mock_import, mock_info, mock_show, app, backup_config, db_connect):
        """CSVインポート成功時は件数を表示して担当一覧を開く"""
        window = MainWindow()

        window.import_csv()

        mock_import.assert_called_once_with(db_connect.return_value)
        assert "3件" in mock_info.call_args[0][2]
        mock_show.assert_called_once()

    @patch('app.main_window.import_csv_to_database')
    @patch('app.main_window.QMessageBox.critical')
    def test_import_csv_error(self, mock_critical, mock_import, app, backup_config, db_connect):
        """CSVインポートエラーのテスト"""
        mock_import.side_effect = Exception("テストエラー")

        window = MainWindow()
        window.import_csv()

        mock_critical.assert_called_once()
        args = mock_critical.call_args[0]
        assert args[0] == window  # 親ウィンドウ
        assert args[1] == "エラー"  # タイトル
        assert "テストエラー" in args[2]  # エラーメッセージ

    def test_connection_is_reused(self, app, backup_config, db_connect):
        """DB には1回だけ接続し、閉じたら次は接続し直す"""
        window = MainWindow()

        first = window.get_connection()
        assert window.get_connection() is first
        db_connect.assert_called_once()

        window.close_connection()
        window.get_connection()
        assert db_connect.call_count == 2

    @patch('app.main_window.DocumentListWindow')
    def test_show_document_window(self, mock_window_class, app, backup_config, db_connect):
        """初回は担当一覧を作って表示し、2回目は読み直す"""
        window = MainWindow()

        window.show_document_window()
        window.show_document_window()

        mock_window_class.assert_called_once_with(db_connect.return_value, window)
        document_window = mock_window_class.return_value
        document_window.load_initial.assert_called_once()
        document_window.reload.assert_called_once()
        assert document_window.showMaximized.call_count == 2

    @pytest.mark.parametrize("method_name, dialog_name", [
        ("show_exclude_docs_dialog", "ExcludeDocsDialog"),
        ("show_exclude_doctors_dialog", "ExcludeDoctorsDialog"),
        ("show_staff_names_dialog", "StaffNamesDialog"),
    ])
    def test_show_list_dialogs(self, method_name, dialog_name, app, backup_config, db_connect):
        """リスト編集ダイアログを DB のリポジトリ付きで開く"""
        with patch(f'app.main_window.{dialog_name}') as mock_dialog:
            window = MainWindow()
            getattr(window, method_name)()

        repository, parent = mock_dialog.call_args[0]
        assert repository.connection is db_connect.return_value
        assert parent is window
        mock_dialog.return_value.exec.assert_called_once()

    @patch('app.main_window.AppearanceDialog')
    def test_show_appearance_dialog(self, mock_dialog, app, backup_config):
        """外観設定ダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_appearance_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    @patch('app.main_window.FolderPathDialog')
    def test_show_folder_path_dialog(self, mock_dialog, app, backup_config):
        """フォルダパスダイアログ表示テスト"""
        mock_instance = MagicMock()
        mock_dialog.return_value = mock_instance

        window = MainWindow()
        window.show_folder_path_dialog()

        # ダイアログが作成され、execが呼ばれたことを確認
        mock_dialog.assert_called_once_with(window)
        mock_instance.exec.assert_called_once()

    @patch('app.main_window.QMainWindow.close')
    def test_close_button(self, mock_close, app, backup_config):
        """閉じるボタンのテスト"""
        window = MainWindow()

        # 閉じるボタンを探して押下
        close_button = None
        for button in window.findChildren(QPushButton):
            if button.text() == "閉じる":
                close_button = button
                break

        assert close_button is not None
        close_button.click()

        # closeメソッドが呼ばれたことを確認
        mock_close.assert_called_once()
