import sys
import pytest
import configparser
from unittest.mock import patch, MagicMock

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QApplication, QPushButton, QWidget

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


class FakeDocumentListWidget(QWidget):
    """DB を読まない担当一覧の代わり"""
    search_status_changed = pyqtSignal(str)

    def __init__(self, connection):
        super().__init__()
        self.connection = connection
        self.load_initial = MagicMock()
        self.reload = MagicMock()
        self.search = MagicMock()
        self.show_next_hit = MagicMock()
        self.show_previous_hit = MagicMock()
        self.refresh_timer = MagicMock()


@pytest.fixture
def db_connect():
    """本物の DB に接続しないよう connect と担当一覧をモックする"""
    with patch('app.main_window.connect') as mock_connect, \
            patch('app.main_window.DocumentListWidget', FakeDocumentListWidget):
        yield mock_connect


def fake_widget(window: MainWindow) -> FakeDocumentListWidget:
    document_widget = window.document_widget
    assert isinstance(document_widget, FakeDocumentListWidget)
    return document_widget


def find_button(window: MainWindow, text: str) -> QPushButton:
    for button in window.findChildren(QPushButton):
        if button.text() == text:
            return button
    raise AssertionError(f"{text} ボタンがありません")


class TestMainWindow:
    def test_init(self, app, backup_config, db_connect):
        """起動時に設定を反映し、DB に接続して担当一覧を同じ画面に表示する"""
        window = MainWindow()

        assert window.windowTitle() == f"医療文書担当一覧 v{__version__}"
        assert window.font().pointSize() == ConfigManager().get_font_size()

        db_connect.assert_called_once()
        document_widget = fake_widget(window)
        assert document_widget.connection is db_connect.return_value
        document_widget.load_initial.assert_called_once()
        assert window.content_layout.indexOf(document_widget) >= 0

        for text in ("検索", "前へ", "次へ", "CSV取り込み", "文書名", "医師名",
                     "担当者名", "画面サイズ", "フォルダパス", "閉じる"):
            find_button(window, text)

    @patch('app.main_window.QMessageBox.critical')
    def test_open_document_list_error(self, mock_critical, app, backup_config, db_connect):
        """DB に接続できなくても画面は開き、エラーを表示する"""
        db_connect.side_effect = Exception("接続エラー")

        window = MainWindow()

        assert window.document_widget is None
        assert "接続エラー" in mock_critical.call_args[0][2]

    @patch('app.main_window.QMessageBox.information')
    @patch('app.main_window.import_csv_to_database', return_value=3)
    def test_import_csv_success(self, mock_import, mock_info, app, backup_config, db_connect):
        """CSVインポート成功時は件数を表示して担当一覧を読み直す"""
        window = MainWindow()

        window.import_csv()

        mock_import.assert_called_once_with(db_connect.return_value)
        assert "3件" in mock_info.call_args[0][2]
        fake_widget(window).reload.assert_called_once()

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
        assert window.document_widget is None
        window.get_connection()
        assert db_connect.call_count == 2

    def test_search(self, app, backup_config, db_connect):
        """検索欄の患者IDで担当一覧を検索し、件数を左側に表示する"""
        window = MainWindow()
        document_widget = fake_widget(window)

        window.patient_id_input.setText('100')
        find_button(window, "検索").click()
        find_button(window, "次へ").click()
        find_button(window, "前へ").click()
        document_widget.search_status_changed.emit("1/2件")

        document_widget.search.assert_called_once_with('100')
        document_widget.show_next_hit.assert_called_once()
        document_widget.show_previous_hit.assert_called_once()
        assert window.search_result_label.text() == "1/2件"

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
    def test_show_appearance_dialog(self, mock_dialog, app, backup_config, db_connect):
        """外観設定ダイアログ表示テスト"""
        window = MainWindow()
        window.show_appearance_dialog()

        mock_dialog.assert_called_once_with(window)
        mock_dialog.return_value.exec.assert_called_once()

    @patch('app.main_window.FolderPathDialog')
    def test_show_folder_path_dialog(self, mock_dialog, app, backup_config, db_connect):
        """フォルダの場所を保存したら、DB に接続し直して担当一覧を作り直す"""
        mock_dialog.return_value.exec.return_value = 1
        window = MainWindow()
        first_widget = window.document_widget

        window.show_folder_path_dialog()

        mock_dialog.assert_called_once_with(window)
        assert db_connect.call_count == 2
        assert window.document_widget is not None
        assert window.document_widget is not first_widget

    @patch('app.main_window.QMainWindow.close')
    def test_close_button(self, mock_close, app, backup_config, db_connect):
        """閉じるボタンのテスト"""
        window = MainWindow()

        find_button(window, "閉じる").click()

        mock_close.assert_called_once()


