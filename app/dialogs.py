import os
from pathlib import Path

from PyQt6.QtCore import QDate
from PyQt6.QtGui import QIntValidator
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QListWidget, QDialogButtonBox, QFileDialog,
    QMessageBox, QCheckBox, QFormLayout
)

from app.document_window import create_date_edit
from services.document_repository import NewDocument
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository
from services.validators import parse_patient_id
from utils.config_manager import ConfigManager

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


class ExcludeDocsDialog(ListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("除外する文書名", EXCLUDE_DOCS, repository, parent)


class ExcludeDoctorsDialog(ListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("除外する医師名", EXCLUDE_DOCTORS, repository, parent)


class StaffNamesDialog(ListItemDialog):
    def __init__(self, repository: ListRepository, parent=None):
        super().__init__("担当者名", STAFF, repository, parent)
        load_button = QPushButton("configから読み込み")
        load_button.clicked.connect(self.load_from_config)
        layout = self.layout()
        if isinstance(layout, QVBoxLayout):
            # OK・キャンセルの上に置く
            layout.insertWidget(layout.count() - 1, load_button)

    def load_from_config(self) -> None:
        """一覧を config.ini の [Staff] ordered_names で置き換える(OK で DB に保存される)"""
        names = ConfigManager().get_staff_names()
        if not names:
            QMessageBox.warning(self, "警告", "config.ini の [Staff] ordered_names に担当者名がありません。")
            return
        self.item_list.clear()
        self.item_list.addItems(names)


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

    def browse_folder(self, path_type):
        if path_type in ['downloads', 'backup']:
            title = "ダウンロードフォルダの選択" if path_type == 'downloads' else "バックアップフォルダの選択"
            path_field = self.downloads_path if path_type == 'downloads' else self.backup_path

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
        super().accept()


class AppearanceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォントとウインドウサイズ")
        self.setModal(True)

        layout = QVBoxLayout()

        layout.addWidget(QLabel("フォントサイズ:"))
        self.font_size_input = QLineEdit()
        self.font_size_input.setValidator(QIntValidator(6, 72))
        layout.addWidget(self.font_size_input)

        layout.addWidget(QLabel("ウィンドウの幅:"))
        self.window_width_input = QLineEdit()
        self.window_width_input.setValidator(QIntValidator(200, 3840))
        layout.addWidget(self.window_width_input)

        layout.addWidget(QLabel("ウィンドウの高さ:"))
        self.window_height_input = QLineEdit()
        self.window_height_input.setValidator(QIntValidator(150, 2160))
        layout.addWidget(self.window_height_input)

        self.start_maximized_input = QCheckBox("起動時に最大化する")
        layout.addWidget(self.start_maximized_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        self.config = ConfigManager()
        self.load_settings()

    def load_settings(self):
        self.font_size_input.setText(str(self.config.get_font_size()))
        window_size = self.config.get_window_size()
        self.window_width_input.setText(str(window_size[0]))
        self.window_height_input.setText(str(window_size[1]))
        self.start_maximized_input.setChecked(self.config.get_start_maximized())

    def accept(self):
        inputs = [self.font_size_input, self.window_width_input, self.window_height_input]
        if not all(field.hasAcceptableInput() for field in inputs):
            QMessageBox.warning(self, "警告", "入力値が範囲外です。")
            return

        self.config.set_font_size(int(self.font_size_input.text()))
        self.config.set_window_size(
            int(self.window_width_input.text()),
            int(self.window_height_input.text())
        )
        self.config.set_start_maximized(self.start_maximized_input.isChecked())
        super().accept()
