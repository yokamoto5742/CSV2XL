import dataclasses
import os
from collections.abc import Callable
from pathlib import Path

from PyQt6.QtCore import QDate, Qt
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QDialogButtonBox, QFileDialog,
    QMessageBox, QCheckBox, QComboBox, QDateEdit, QFormLayout, QWidget
)

from app.document_table_model import COLUMNS, NOTE_COLUMN, format_cell
from services.document_repository import Document, NewDocument
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository
from services.validators import parse_patient_id
from utils.config_manager import ConfigManager

DISPLAY_DATE_FORMAT = 'yyyy/MM/dd'
EMPTY_VALUE_LABEL = "(空白)"
DETAIL_DIALOG_MIN_WIDTH = 480


def create_date_edit(parent: QWidget | None = None) -> QDateEdit:
    """▼でカレンダーを開ける日付欄"""
    date_edit = QDateEdit(parent)
    date_edit.setCalendarPopup(True)
    date_edit.setDisplayFormat(DISPLAY_DATE_FORMAT)
    return date_edit


def create_ok_cancel_buttons(dialog: QDialog) -> QDialogButtonBox:
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    return buttons


class ListItemDialog(QDialog):
    """DB に保存するリスト(除外する文書名・医師名、担当者名)の編集"""

    def __init__(self, title, category, repository: ListRepository, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.category = category
        self.repository = repository

        layout = QVBoxLayout()

        self.input_field = QLineEdit()
        layout.addWidget(QLabel(f"{title}を入力:"))
        layout.addWidget(self.input_field)

        self.item_list = QListWidget()
        layout.addWidget(QLabel(f"登録済み{title}:"))
        layout.addWidget(self.item_list)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_item)

        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.load_items()

    def load_items(self):
        self.item_list.addItems(self.repository.get_items(self.category))

    def add_item(self):
        item_name = self.input_field.text().strip()
        if item_name:
            self.item_list.addItem(item_name)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.item_list.currentItem()
        if current_item:
            self.item_list.takeItem(self.item_list.row(current_item))

    def accept(self):
        items = []
        for i in range(self.item_list.count()):
            item = self.item_list.item(i)
            if item is not None:
                items.append(item.text())

        self.repository.set_items(self.category, items)
        super().accept()


class ConfigListItemDialog(ListItemDialog):
    """config.ini のカンマ区切りのリストで一覧を置き換えられるリストの編集"""

    def __init__(self, title: str, category: str, config_key: str,
                 read_config: Callable[[ConfigManager], list[str]], repository: ListRepository,
                 parent: QWidget | None = None) -> None:
        super().__init__(title, category, repository, parent)
        self.title = title
        self.config_key = config_key
        self.read_config = read_config
        load_button = QPushButton("configから読み込み")
        load_button.clicked.connect(self.load_from_config)
        layout = self.layout()
        if isinstance(layout, QVBoxLayout):
            # OK・キャンセルの上に置く
            layout.insertWidget(layout.count() - 1, load_button)

    def load_from_config(self) -> None:
        """一覧を config.ini のリストで置き換える(OK で DB に保存される)"""
        names = self.read_config(ConfigManager())
        if not names:
            QMessageBox.warning(self, "警告", f"config.ini の {self.config_key} に{self.title}がありません。")
            return
        self.item_list.clear()
        self.item_list.addItems(names)


class ExcludeDocsDialog(ConfigListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("除外する文書名", EXCLUDE_DOCS, "[ExcludeDocs] list",
                         lambda config: config.get_exclude_docs(), repository, parent)


class ExcludeDoctorsDialog(ConfigListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("除外する医師名", EXCLUDE_DOCTORS, "[ExcludeDoctors] list",
                         lambda config: config.get_exclude_doctors(), repository, parent)


class StaffNamesDialog(ConfigListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("担当者名", STAFF, "[Staff] ordered_names",
                         lambda config: config.get_staff_names(), repository, parent)


class AddDocumentDialog(QDialog):
    """手動で追加する行(預り日〜医師名)の入力"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("行の追加")
        self.setModal(True)

        layout = QFormLayout()
        self.received_date_input = create_date_edit()
        self.received_date_input.setDate(QDate.currentDate())
        layout.addRow("預り日:", self.received_date_input)
        self.patient_id_input = QLineEdit()
        layout.addRow("患者ID:", self.patient_id_input)
        self.text_inputs: dict[str, QLineEdit] = {}
        for field, label in (("patient_name", "患者名"), ("document_name", "文書名"),
                             ("department", "診療科"), ("doctor_name", "医師名")):
            self.text_inputs[field] = QLineEdit()
            layout.addRow(f"{label}:", self.text_inputs[field])

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        self.setLayout(layout)

    def new_document(self) -> NewDocument | None:
        """入力内容。未入力の項目や数字でない患者IDがあれば None"""
        patient_id = parse_patient_id(self.patient_id_input.text())
        texts = {field: line_edit.text().strip() for field, line_edit in self.text_inputs.items()}
        if patient_id is None or not all(texts.values()):
            return None
        return NewDocument(
            received_date=self.received_date_input.date().toString('yyyy-MM-dd'),
            patient_id=patient_id,
            **texts,
        )

    def accept(self):
        if self.new_document() is None:
            QMessageBox.warning(self, "警告", "すべての項目を入力してください。\n患者IDは数字で入力してください。")
            return
        super().accept()


class FolderPathDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォルダの場所")
        self.setModal(True)

        config = ConfigManager()
        dialog_width, dialog_height = config.get_folder_dialog_size()
        self.resize(dialog_width, dialog_height)

        layout = QVBoxLayout()

        self.downloads_path = self.create_path_section(
            layout, "ダウンロードフォルダ:",
            lambda: self.browse_folder('downloads'),
            lambda: self.open_folder(self.downloads_path.text())
        )

        self.database_path = self.create_path_section(
            layout, "データベースファイル:",
            lambda: self.browse_folder('database'),
            lambda: self.open_folder(str(Path(self.database_path.text()).parent))
        )

        self.backup_path = self.create_path_section(
            layout, "バックアップフォルダ:",
            lambda: self.browse_folder('backup'),
            lambda: self.open_folder(self.backup_path.text())
        )

        self.xlsm_path = self.create_path_section(
            layout, "xlsmフォルダ:",
            lambda: self.browse_folder('xlsm'),
            lambda: self.open_folder(self.xlsm_path.text())
        )

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.config = ConfigManager()
        self.load_paths()

    @staticmethod
    def create_path_section(layout, label_text, browse_callback, open_callback):
        layout.addWidget(QLabel(label_text))
        path_layout = QHBoxLayout()

        path_field = QLineEdit()
        path_layout.addWidget(path_field)

        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(browse_callback)
        path_layout.addWidget(browse_button)

        open_button = QPushButton("開く")
        open_button.clicked.connect(open_callback)
        path_layout.addWidget(open_button)

        layout.addLayout(path_layout)
        return path_field

    def open_folder(self, path: str):
        if path and os.path.exists(path):
            os.startfile(path)
        else:
            QMessageBox.warning(self, "警告", "指定されたフォルダが存在しません。")

    def load_paths(self):
        self.downloads_path.setText(self.config.get_downloads_path())
        self.database_path.setText(self.config.get_database_path())
        self.backup_path.setText(self.config.get_backup_path())
        self.xlsm_path.setText(self.config.get_xlsm_path())

    def browse_folder(self, path_type):
        folders = {
            'downloads': ("ダウンロードフォルダの選択", self.downloads_path),
            'backup': ("バックアップフォルダの選択", self.backup_path),
            'xlsm': ("xlsmフォルダの選択", self.xlsm_path),
        }
        if path_type in folders:
            title, path_field = folders[path_type]

            folder = QFileDialog.getExistingDirectory(
                self,
                title,
                path_field.text()
            )
            if folder:
                path_field.setText(folder)
        else:
            # 新しい場所を指定することもあるため、存在しないファイル名も選べる保存ダイアログを使う
            file, _ = QFileDialog.getSaveFileName(
                self,
                "データベースファイルの選択",
                self.database_path.text(),
                "SQLite データベース (*.db)",
                options=QFileDialog.Option.DontConfirmOverwrite
            )
            if file:
                self.database_path.setText(file)

    def accept(self):
        self.config.set_downloads_path(self.downloads_path.text())
        self.config.set_database_path(self.database_path.text())
        self.config.set_backup_path(self.backup_path.text())
        self.config.set_xlsm_path(self.xlsm_path.text())
        super().accept()


class SettingsDialog(QDialog):
    """xlsm取り込み・CSV出力・フォルダパスのボタンをまとめたダイアログ。押したボタンの処理はダイアログを閉じてから行う"""

    def __init__(self, actions: list[tuple[str, Callable[[], object]]], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setModal(True)
        layout = QVBoxLayout(self)
        for text, action in actions:
            button = QPushButton(text)
            button.clicked.connect(lambda _checked, action=action: self.run(action))
            layout.addWidget(button)
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.reject)
        layout.addWidget(close_button)

    def run(self, action: Callable[[], object]) -> None:
        self.accept()
        action()


class DocumentDetailDialog(QDialog):
    """1行の全項目を表示し、担当者名・医師依頼日・備考を編集する"""

    def __init__(self, document: Document, staff_names: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("行の詳細")
        self.setModal(True)
        self.setMinimumWidth(DETAIL_DIALOG_MIN_WIDTH)
        self.document = document

        layout = QFormLayout(self)
        # 長い文書名なども折り返して全文を表示する(選んでコピーもできる)
        for column, (heading, _) in enumerate(COLUMNS[:6]):
            value_label = QLabel(format_cell(document, column))
            value_label.setWordWrap(True)
            value_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            layout.addRow(f"{heading}:", value_label)

        self.staff_input = QComboBox()
        names = ['', *staff_names]
        # リストから外れた担当者が入力済みの行でも、今の値を消さない
        if document.staff_name and document.staff_name not in names:
            names.append(document.staff_name)
        self.staff_input.addItems(names)
        self.staff_input.setCurrentText(document.staff_name or '')
        layout.addRow("担当者名:", self.staff_input)

        self.requested_date_input = create_date_edit()
        self.no_requested_date_input = QCheckBox("未設定")
        self.no_requested_date_input.toggled.connect(self.requested_date_input.setDisabled)
        requested_date = QDate.fromString(document.requested_date or '', 'yyyy-MM-dd')
        self.requested_date_input.setDate(requested_date if requested_date.isValid() else QDate.currentDate())
        self.no_requested_date_input.setChecked(document.requested_date is None)
        requested_date_row = QHBoxLayout()
        requested_date_row.addWidget(self.requested_date_input)
        requested_date_row.addWidget(self.no_requested_date_input)
        layout.addRow("医師依頼日:", requested_date_row)

        self.note_input = QLineEdit(document.note or '')
        layout.addRow(f"{COLUMNS[NOTE_COLUMN][0]}:", self.note_input)

        layout.addRow(create_ok_cancel_buttons(self))

    def edited_document(self) -> Document:
        """入力内容を反映した行"""
        requested_date = (None if self.no_requested_date_input.isChecked()
                          else self.requested_date_input.date().toString('yyyy-MM-dd'))
        return dataclasses.replace(
            self.document,
            staff_name=self.staff_input.currentText().strip() or None,
            requested_date=requested_date,
            note=self.note_input.text().strip() or None,
        )


class ColumnFilterDialog(QDialog):
    """列の値をチェックボックスで選んで絞り込む(Excel のオートフィルタに近い)"""

    def __init__(self, heading: str, values: list[str], selected: set[str] | None,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{heading}のフィルタ")
        self.setModal(True)
        self.values = values

        layout = QVBoxLayout(self)
        self.select_all_input = QCheckBox("すべて選択")
        self.select_all_input.setChecked(selected is None)
        self.select_all_input.clicked.connect(self.set_all_checked)
        layout.addWidget(self.select_all_input)

        self.value_list = QListWidget()
        for value in values:
            item = QListWidgetItem(value or EMPTY_VALUE_LABEL)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            checked = selected is None or value in selected
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            self.value_list.addItem(item)
        layout.addWidget(self.value_list)

        layout.addWidget(create_ok_cancel_buttons(self))

    def set_all_checked(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for row in range(self.value_list.count()):
            item = self.value_list.item(row)
            if item is not None:
                item.setCheckState(state)

    def checked_values(self) -> set[str]:
        checked = set()
        for row, value in enumerate(self.values):
            item = self.value_list.item(row)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                checked.add(value)
        return checked

    def selected_values(self) -> set[str] | None:
        """表示する値。すべてチェックしていれば None(絞り込まない)"""
        checked = self.checked_values()
        return None if len(checked) == len(self.values) else checked

    def accept(self):
        if not self.checked_values():
            QMessageBox.warning(self, "警告", "表示する値を1つ以上選んでください。")
            return
        super().accept()
