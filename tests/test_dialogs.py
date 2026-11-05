import sys
import pytest
import configparser
from unittest.mock import MagicMock, patch

from PyQt6.QtCore import QDate, Qt
from PyQt6.QtWidgets import QApplication, QDialog, QListWidget, QPushButton

from app.dialogs import (
    AddDocumentDialog, ColumnFilterDialog, DocumentDetailDialog, ExcludeDocsDialog, ExcludeDoctorsDialog,
    FolderPathDialog, ListItemDialog, SettingsDialog, StaffNamesDialog
)
from services.document_repository import Document, NewDocument
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository
from utils.config_manager import ConfigManager, CONFIG_PATH


def get_item_text(list_widget: QListWidget, row: int) -> str:
    item = list_widget.item(row)
    assert item is not None
    return item.text()


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
def list_repository(connection) -> ListRepository:
    return ListRepository(connection)


class TestListItemDialog:
    def test_init(self, app, list_repository):
        """ListItemDialogの初期化テスト"""
        dialog = ListItemDialog("テストダイアログ", STAFF, list_repository)
        assert dialog.windowTitle() == "テストダイアログ"
        assert dialog.category == STAFF
        assert dialog.item_list is not None

    def test_add_item(self, app, list_repository):
        """項目追加機能のテスト"""
        dialog = ListItemDialog("テストダイアログ", STAFF, list_repository)

        dialog.input_field.setText("テスト項目1")
        dialog.add_item()

        assert dialog.item_list.count() == 1
        assert get_item_text(dialog.item_list, 0) == "テスト項目1"

        # 空白の項目は追加されないことを確認
        dialog.input_field.setText("  ")
        dialog.add_item()
        assert dialog.item_list.count() == 1

    def test_delete_selected(self, app, list_repository):
        """項目削除機能のテスト"""
        dialog = ListItemDialog("テストダイアログ", STAFF, list_repository)
        dialog.input_field.setText("テスト項目1")
        dialog.add_item()
        dialog.input_field.setText("テスト項目2")
        dialog.add_item()

        dialog.item_list.setCurrentRow(0)
        dialog.delete_selected()

        assert dialog.item_list.count() == 1
        assert get_item_text(dialog.item_list, 0) == "テスト項目2"

    def test_load_items(self, app, list_repository):
        """DB からの項目読み込みテスト"""
        list_repository.set_items(STAFF, ['アイテム1', 'アイテム2', 'アイテム3'])

        dialog = ListItemDialog("テストダイアログ", STAFF, list_repository)

        assert [get_item_text(dialog.item_list, i) for i in range(3)] == ['アイテム1', 'アイテム2', 'アイテム3']

    def test_accept(self, app, list_repository):
        """ダイアログ確定時に DB へ保存するテスト"""
        dialog = ListItemDialog("テストダイアログ", STAFF, list_repository)
        dialog.input_field.setText("新しい項目1")
        dialog.add_item()
        dialog.input_field.setText("新しい項目2")
        dialog.add_item()

        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()
            mock_accept.assert_called_once()

        assert list_repository.get_items(STAFF) == ["新しい項目1", "新しい項目2"]


class TestListDialogSubclasses:
    @pytest.mark.parametrize("dialog_class, title, category", [
        (ExcludeDocsDialog, "除外する文書名", EXCLUDE_DOCS),
        (ExcludeDoctorsDialog, "除外する医師名", EXCLUDE_DOCTORS),
        (StaffNamesDialog, "担当者名", STAFF),
    ])
    def test_init(self, app, list_repository, dialog_class, title, category):
        dialog = dialog_class(list_repository)
        assert dialog.windowTitle() == title
        assert dialog.category == category


CONFIG_LIST_DIALOGS = [
    (StaffNamesDialog, STAFF, 'get_staff_names', "[Staff] ordered_names"),
    (ExcludeDocsDialog, EXCLUDE_DOCS, 'get_exclude_docs', "[ExcludeDocs] list"),
    (ExcludeDoctorsDialog, EXCLUDE_DOCTORS, 'get_exclude_doctors', "[ExcludeDoctors] list"),
]


class TestLoadFromConfig:
    @pytest.mark.parametrize("dialog_class, category, getter, config_key", CONFIG_LIST_DIALOGS)
    def test_load_from_config(self, app, list_repository, dialog_class, category, getter, config_key):
        """config のリストで一覧を置き換え、OK で DB に保存する"""
        list_repository.set_items(category, ['旧項目'])
        dialog = dialog_class(list_repository)

        with patch.object(ConfigManager, getter, return_value=['植田', '花﨑']):
            dialog.load_from_config()
        with patch.object(QDialog, 'accept'):
            dialog.accept()

        assert list_repository.get_items(category) == ['植田', '花﨑']

    @pytest.mark.parametrize("dialog_class, category, getter, config_key", CONFIG_LIST_DIALOGS)
    @patch('app.dialogs.QMessageBox.warning')
    def test_load_from_config_empty(self, mock_warning, app, list_repository, dialog_class, category, getter,
                                    config_key):
        """config のリストが空なら警告し、一覧はそのまま"""
        list_repository.set_items(category, ['山本'])
        dialog = dialog_class(list_repository)

        with patch.object(ConfigManager, getter, return_value=[]):
            dialog.load_from_config()

        assert config_key in mock_warning.call_args[0][2]
        assert dialog.item_list.count() == 1


class TestAddDocumentDialog:
    def fill(self, dialog: AddDocumentDialog, patient_id: str = '１２３') -> None:
        dialog.received_date_input.setDate(QDate(2025, 1, 5))
        dialog.patient_id_input.setText(patient_id)
        for field, value in (("patient_name", "山田　太郎"), ("document_name", "診断書"),
                             ("department", "内科"), ("doctor_name", "田中")):
            dialog.text_inputs[field].setText(value)

    def test_new_document(self, app):
        dialog = AddDocumentDialog()
        self.fill(dialog)

        assert dialog.new_document() == NewDocument('2025-01-05', 123, '山田　太郎', '診断書', '内科', '田中')

    @pytest.mark.parametrize("patient_id", ["", "abc"])
    @patch('app.dialogs.QMessageBox.warning')
    def test_rejects_invalid_patient_id(self, mock_warning, patient_id, app):
        dialog = AddDocumentDialog()
        self.fill(dialog, patient_id)

        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()

        mock_accept.assert_not_called()
        mock_warning.assert_called_once()

    def test_rejects_empty_field(self, app):
        dialog = AddDocumentDialog()
        self.fill(dialog)
        dialog.text_inputs["doctor_name"].setText("  ")

        assert dialog.new_document() is None


class TestFolderPathDialog:
    def test_init(self, app, backup_config):
        """FolderPathDialogの初期化テスト"""
        dialog = FolderPathDialog()
        assert dialog.windowTitle() == "フォルダの場所"
        assert dialog.downloads_path is not None
        assert dialog.database_path is not None
        assert dialog.backup_path is not None

    def test_load_paths(self, app, backup_config):
        """パス読み込み機能のテスト"""
        # 設定ファイルにテスト用パスを設定
        config = ConfigManager()
        if 'Paths' not in config.config:
            config.config['Paths'] = {}
        config.config['Paths']['downloads_path'] = 'C:\\Test\\Downloads'
        config.config['Paths']['database_path'] = 'C:\\Test\\DB\\test.db'
        config.config['Paths']['backup_path'] = 'C:\\Test\\Backup'
        config.save_config()

        # ダイアログを初期化してパスが読み込まれるか確認
        dialog = FolderPathDialog()

        assert dialog.downloads_path.text() == 'C:\\Test\\Downloads'
        assert dialog.database_path.text() == 'C:\\Test\\DB\\test.db'
        assert dialog.backup_path.text() == 'C:\\Test\\Backup'

    @patch('app.dialogs.QFileDialog.getExistingDirectory')
    def test_browse_folder_downloads(self, mock_get_dir, app, backup_config):
        """ダウンロードフォルダブラウズ機能のテスト"""
        mock_get_dir.return_value = 'C:\\New\\Downloads'

        dialog = FolderPathDialog()
        dialog.browse_folder('downloads')

        mock_get_dir.assert_called_once()
        assert dialog.downloads_path.text() == 'C:\\New\\Downloads'

    @patch('app.dialogs.QFileDialog.getExistingDirectory')
    def test_browse_folder_backup(self, mock_get_dir, app, backup_config):
        """バックアップフォルダブラウズ機能のテスト"""
        mock_get_dir.return_value = 'C:\\New\\Backup'

        dialog = FolderPathDialog()
        dialog.browse_folder('backup')

        mock_get_dir.assert_called_once()
        assert dialog.backup_path.text() == 'C:\\New\\Backup'

    @patch('app.dialogs.QFileDialog.getExistingDirectory')
    def test_browse_folder_xlsm(self, mock_get_dir, app, backup_config):
        mock_get_dir.return_value = 'C:\\New\\Xlsm'

        dialog = FolderPathDialog()
        dialog.browse_folder('xlsm')

        assert dialog.xlsm_path.text() == 'C:\\New\\Xlsm'

    @patch('app.dialogs.QFileDialog.getSaveFileName')
    def test_browse_folder_database(self, mock_get_file, app, backup_config):
        """データベースファイルブラウズ機能のテスト"""
        mock_get_file.return_value = ('C:\\New\\DB\\new.db', 'SQLite データベース (*.db)')

        dialog = FolderPathDialog()
        dialog.browse_folder('database')

        mock_get_file.assert_called_once()
        assert dialog.database_path.text() == 'C:\\New\\DB\\new.db'

    @patch('app.dialogs.os.path.exists')
    @patch('app.dialogs.os.startfile')
    def test_open_folder_exists(self, mock_startfile, mock_exists, app, backup_config):
        """フォルダオープン機能のテスト（フォルダが存在する場合）"""
        mock_exists.return_value = True

        dialog = FolderPathDialog()
        dialog.open_folder('C:\\Test\\Folder')

        mock_exists.assert_called_once_with('C:\\Test\\Folder')
        mock_startfile.assert_called_once_with('C:\\Test\\Folder')

    @patch('app.dialogs.os.path.exists')
    @patch('app.dialogs.QMessageBox.warning')
    def test_open_folder_not_exists(self, mock_warning, mock_exists, app, backup_config):
        """フォルダオープン機能のテスト（フォルダが存在しない場合）"""
        mock_exists.return_value = False

        dialog = FolderPathDialog()
        dialog.open_folder('C:\\Test\\NonExistingFolder')

        mock_exists.assert_called_once_with('C:\\Test\\NonExistingFolder')
        mock_warning.assert_called_once()

    def test_accept(self, app, backup_config):
        """ダイアログ確定時の設定保存テスト"""
        dialog = FolderPathDialog()

        # パスを設定
        dialog.downloads_path.setText('C:\\New\\Downloads')
        dialog.database_path.setText('C:\\New\\DB\\new.db')
        dialog.backup_path.setText('C:\\New\\Backup')
        dialog.xlsm_path.setText('C:\\New\\Xlsm')

        # モック化したsuperクラスのacceptメソッドでテスト
        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()
            mock_accept.assert_called_once()

        # 設定が保存されたことを確認
        config = ConfigManager()
        assert config.get_downloads_path() == 'C:\\New\\Downloads'
        assert config.get_database_path() == 'C:\\New\\DB\\new.db'
        assert config.get_backup_path() == 'C:\\New\\Backup'
        assert config.get_xlsm_path() == 'C:\\New\\Xlsm'


class TestSettingsDialog:
    def test_button_closes_dialog_then_runs_action(self, app):
        action = MagicMock()
        dialog = SettingsDialog([("xlsm取り込み", action), ("CSV出力", MagicMock())])
        button = next(b for b in dialog.findChildren(QPushButton) if b.text() == "xlsm取り込み")

        with patch.object(QDialog, 'accept') as mock_accept:
            button.click()

        mock_accept.assert_called_once()
        action.assert_called_once()

    def test_buttons(self, app):
        dialog = SettingsDialog([("xlsm取り込み", MagicMock()), ("CSV出力", MagicMock())])

        assert [b.text() for b in dialog.findChildren(QPushButton)] == ["xlsm取り込み", "CSV出力", "閉じる"]


def make_document(staff_name: str | None = None, requested_date: str | None = None,
                  note: str | None = None) -> Document:
    return Document(1, '2025-01-10', 100, '山田　太郎', 'とても長い文書名' * 5, '内科', '田中',
                    staff_name, requested_date, note, 3)


class TestDocumentDetailDialog:
    def test_shows_all_values(self, app):
        dialog = DocumentDetailDialog(make_document('退職者', '2025-01-15', '至急'), ['山本'])

        assert dialog.staff_input.currentText() == '退職者'
        assert dialog.requested_date_input.date() == QDate(2025, 1, 15)
        assert not dialog.no_requested_date_input.isChecked()
        assert dialog.note_input.text() == '至急'
        assert dialog.edited_document() == make_document('退職者', '2025-01-15', '至急')

    def test_empty_requested_date(self, app):
        dialog = DocumentDetailDialog(make_document(), ['山本'])

        assert dialog.no_requested_date_input.isChecked()
        assert not dialog.requested_date_input.isEnabled()
        assert dialog.edited_document().requested_date is None

    def test_edited_document(self, app):
        dialog = DocumentDetailDialog(make_document(), ['山本'])
        dialog.staff_input.setCurrentText('山本')
        dialog.no_requested_date_input.setChecked(False)
        dialog.requested_date_input.setDate(QDate(2025, 2, 3))
        dialog.note_input.setText(' メモ ')

        edited = dialog.edited_document()

        assert (edited.staff_name, edited.requested_date, edited.note) == ('山本', '2025-02-03', 'メモ')
        assert (edited.id, edited.version) == (1, 3)


def check(dialog: ColumnFilterDialog, row: int, checked: bool) -> None:
    item = dialog.value_list.item(row)
    assert item is not None
    item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)


class TestColumnFilterDialog:
    def test_all_checked_means_no_filter(self, app):
        dialog = ColumnFilterDialog("担当者名", ['山本', ''], None)

        assert get_item_text(dialog.value_list, 1) == "(空白)"
        assert dialog.select_all_input.isChecked()
        assert dialog.selected_values() is None

    def test_selected_values(self, app):
        dialog = ColumnFilterDialog("担当者名", ['山本', '植田', ''], {'山本'})

        assert dialog.selected_values() == {'山本'}
        check(dialog, 2, True)
        assert dialog.selected_values() == {'山本', ''}

    def test_select_all(self, app):
        dialog = ColumnFilterDialog("担当者名", ['山本', '植田'], {'山本'})

        dialog.set_all_checked(True)
        assert dialog.selected_values() is None

        dialog.set_all_checked(False)
        assert dialog.checked_values() == set()

    @patch('app.dialogs.QMessageBox.warning')
    def test_rejects_nothing_checked(self, mock_warning, app):
        dialog = ColumnFilterDialog("担当者名", ['山本'], None)
        check(dialog, 0, False)

        with patch.object(QDialog, 'accept') as mock_accept:
            dialog.accept()

        mock_accept.assert_not_called()
        mock_warning.assert_called_once()
