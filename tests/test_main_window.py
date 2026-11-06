import sys
import pytest
import configparser
from unittest.mock import patch, MagicMock

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QApplication, QLabel, QPushButton, QWidget

from app.main_window import MainWindow, apply_app_font_size
from services.xlsm_import import XlsmImportResult
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
        self.clear_search = MagicMock()
        self.set_column_widths = MagicMock()
        self.column_widths = MagicMock(return_value=[90, 70, 110, 300, 100, 90, 90, 100, 200])
        self.apply_font_size = MagicMock()
        self.refresh_timer = MagicMock()
        self.add_document = MagicMock()
        self.delete_selected = MagicMock()
        self.export_csv = MagicMock()


@pytest.fixture
def db_connect():
    """本物の DB に接続しないよう connect と担当一覧をモックする"""
    with patch('app.main_window.connect') as mock_connect, \
            patch('app.main_window.DocumentListWidget', FakeDocumentListWidget), \
            patch('app.main_window.backup_database_daily'):
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
        assert window.splitter.indexOf(document_widget) == 1
        assert QApplication.font().pointSize() == ConfigManager().get_font_size()

        assert all(__version__ not in label.text() for label in window.findChildren(QLabel))
        for text in ("検索", "絞り込み解除", "CSV取り込み", "行の追加", "行の削除", "文書名", "医師名", "設定", "閉じる"):
            find_button(window, text)
        # xlsm取り込み・CSV出力・フォルダパス・担当者名は設定ダイアログに移した
        texts = [button.text() for button in window.findChildren(QPushButton)]
        assert not {"xlsm取り込み", "CSV出力", "フォルダパス", "担当者名", "前へ", "次へ"} & set(texts)

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

    @patch('app.main_window.QMessageBox.information')
    @patch('app.main_window.import_xlsm_to_database')
    def test_import_xlsm_success(self, mock_import, mock_info, app, backup_config, db_connect):
        """xlsm 取り込み後は件数と取り込めなかった行を表示して担当一覧を読み直す"""
        mock_import.return_value = XlsmImportResult('医療文書担当一覧_202501101200.xlsm', 3, 2, ['5行目: 預り日が空です'])
        window = MainWindow()

        window.import_xlsm()

        mock_import.assert_called_once_with(db_connect.return_value)
        message = mock_info.call_args[0][2]
        assert "3件を追加し、2件の空欄を埋めました" in message
        assert "5行目: 預り日が空です" in message
        fake_widget(window).reload.assert_called_once()

    @patch('app.main_window.import_xlsm_to_database', side_effect=FileNotFoundError("xlsmなし"))
    @patch('app.main_window.QMessageBox.critical')
    def test_import_xlsm_error(self, mock_critical, mock_import, app, backup_config, db_connect):
        window = MainWindow()

        window.import_xlsm()

        assert "xlsmなし" in mock_critical.call_args[0][2]

    def test_backup_daily_on_startup(self, app, backup_config, db_connect):
        with patch('app.main_window.backup_database_daily') as mock_backup:
            MainWindow()

        mock_backup.assert_called_once_with(db_connect.return_value)

    @patch('app.main_window.QMessageBox.warning')
    def test_backup_daily_error(self, mock_warning, app, backup_config, db_connect):
        with patch('app.main_window.backup_database_daily', side_effect=OSError("書き込み不可")):
            MainWindow()

        assert "書き込み不可" in mock_warning.call_args[0][2]

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
        """検索欄の患者IDで担当一覧を絞り込み、件数を左側に表示する。絞り込み解除で検索欄も空にする"""
        window = MainWindow()
        document_widget = fake_widget(window)

        window.patient_id_input.setText('100')
        find_button(window, "検索").click()
        document_widget.search_status_changed.emit("患者ID 100: 2件")

        document_widget.search.assert_called_once_with('100')
        assert window.search_result_label.text() == "患者ID 100: 2件"

        find_button(window, "絞り込み解除").click()
        document_widget.clear_search.assert_called_once()
        assert window.patient_id_input.text() == ''

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

    @patch('app.main_window.SettingsDialog')
    def test_show_settings_dialog(self, mock_dialog, app, backup_config, db_connect):
        """設定ダイアログに xlsm取り込み・CSV出力・フォルダパス・フォントサイズ・担当者名を渡す"""
        window = MainWindow()

        find_button(window, "設定").click()

        actions, parent = mock_dialog.call_args[0]
        assert actions == [("xlsm取り込み", window.import_xlsm), ("CSV出力", window.export_csv),
                           ("フォルダパス", window.show_folder_path_dialog),
                           ("フォントサイズ", window.show_font_size_dialog),
                           ("担当者名", window.show_staff_names_dialog)]
        assert parent is window
        mock_dialog.return_value.exec.assert_called_once()

    @patch('app.main_window.FontSizeDialog')
    def test_font_size_is_applied_and_saved(self, mock_dialog, app, backup_config, db_connect):
        """OK なら保存し、作成済みの画面(スタイルシートを設定した左パネルのボタンも)にすぐ反映する"""
        window = MainWindow()
        original_size = ConfigManager().get_font_size()
        new_size = original_size + 2
        mock_dialog.return_value.exec.return_value = 1
        mock_dialog.return_value.font_size.return_value = new_size

        window.show_font_size_dialog()

        assert ConfigManager().get_font_size() == new_size
        assert find_button(window, "CSV取り込み").font().pointSize() == new_size
        fake_widget(window).apply_font_size.assert_called_once()
        apply_app_font_size(original_size)

    @patch('app.main_window.FontSizeDialog')
    def test_font_size_cancel(self, mock_dialog, app, backup_config, db_connect):
        window = MainWindow()
        original_size = ConfigManager().get_font_size()
        mock_dialog.return_value.exec.return_value = 0

        window.show_font_size_dialog()

        assert ConfigManager().get_font_size() == original_size
        fake_widget(window).apply_font_size.assert_not_called()

    def test_column_widths_are_saved_on_close(self, app, backup_config, db_connect):
        """終了時に列幅を保存し、次回はその列幅を担当一覧に渡す"""
        window = MainWindow()
        window.close()

        assert ConfigManager().get_column_widths() == [90, 70, 110, 300, 100, 90, 90, 100, 200]

        reopened = MainWindow()
        fake_widget(reopened).set_column_widths.assert_called_once_with([90, 70, 110, 300, 100, 90, 90, 100, 200])

    def test_window_geometry_is_saved_on_close(self, app, backup_config, db_connect):
        """終了時の位置と大きさを保存し、次回はその位置と大きさで開く"""
        window = MainWindow()
        window.setGeometry(120, 130, 900, 700)
        window.show()
        window.close()

        config = ConfigManager()
        assert config.get_window_position() == (120, 130)
        assert config.get_window_size() == (900, 700)
        assert config.get_start_maximized() is False

        reopened = MainWindow()
        assert (reopened.x(), reopened.y()) == (120, 130)
        assert (reopened.width(), reopened.height()) == (900, 700)

    def test_window_geometry_saved_while_maximized(self, app, backup_config, db_connect):
        """最大化したまま終了したら、最大化を解除したときの大きさと最大化の状態を保存する"""
        window = MainWindow()
        window.setGeometry(120, 130, 900, 700)
        window.showMaximized()
        app.processEvents()
        window.close()

        config = ConfigManager()
        assert config.get_start_maximized() is True
        assert config.get_window_size() == (900, 700)

    def test_offscreen_position_is_ignored(self, app, backup_config, db_connect):
        """保存した位置がどの画面にも入らなければ、大きさだけを戻す"""
        ConfigManager().set_window_geometry(-50000, -50000, 900, 700, False)

        window = MainWindow()

        assert (window.x(), window.y()) != (-50000, -50000)
        assert (window.width(), window.height()) == (900, 700)

    @patch('app.main_window.AddDocumentDialog')
    def test_add_document(self, mock_dialog, app, backup_config, db_connect):
        mock_dialog.return_value.exec.return_value = 1
        window = MainWindow()

        find_button(window, "行の追加").click()

        fake_widget(window).add_document.assert_called_once_with(mock_dialog.return_value.new_document.return_value)

    @patch('app.main_window.AddDocumentDialog')
    def test_add_document_cancelled(self, mock_dialog, app, backup_config, db_connect):
        mock_dialog.return_value.exec.return_value = 0
        window = MainWindow()

        window.add_document()

        fake_widget(window).add_document.assert_not_called()

    def test_delete_document(self, app, backup_config, db_connect):
        window = MainWindow()

        find_button(window, "行の削除").click()

        fake_widget(window).delete_selected.assert_called_once()

    @patch('app.main_window.QMessageBox.information')
    @patch('app.main_window.QFileDialog.getSaveFileName', return_value=('C:/out/list.csv', ''))
    def test_export_csv(self, mock_save, mock_info, app, backup_config, db_connect):
        window = MainWindow()

        window.export_csv()

        assert mock_save.call_args[0][2].endswith('.csv')
        fake_widget(window).export_csv.assert_called_once_with('C:/out/list.csv')
        mock_info.assert_called_once()

    @patch('app.main_window.QFileDialog.getSaveFileName', return_value=('', ''))
    def test_export_csv_cancelled(self, mock_save, app, backup_config, db_connect):
        window = MainWindow()

        window.export_csv()

        fake_widget(window).export_csv.assert_not_called()

    @patch('app.main_window.QMessageBox.critical')
    @patch('app.main_window.QFileDialog.getSaveFileName', return_value=('C:/out/list.csv', ''))
    def test_export_csv_error(self, mock_save, mock_critical, app, backup_config, db_connect):
        window = MainWindow()
        fake_widget(window).export_csv.side_effect = OSError("書き込めません")

        window.export_csv()

        assert "書き込めません" in mock_critical.call_args[0][2]

    def test_side_panel_width_is_saved_on_close(self, app, backup_config, db_connect):
        """左パネルの幅は終了時に保存し、次回はその幅で開く"""
        # 左パネルの最小幅より広い値にする(狭すぎる値は Qt が最小幅に広げる)
        ConfigManager().set_side_panel_width(400)
        window = MainWindow()
        window.resize(1400, 600)
        window.show()
        assert window.splitter.sizes()[0] == 400

        window.splitter.setSizes([450, 950])
        window.close()

        assert ConfigManager().get_side_panel_width() == 450

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


