import os
from pathlib import Path

from PyQt6.QtGui import QIntValidator
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QListWidget, QDialogButtonBox, QFileDialog,
    QMessageBox
)

from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository
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
        self.window_width_input.setValidator(QIntValidator(200, 1000))
        layout.addWidget(self.window_width_input)

        layout.addWidget(QLabel("ウィンドウの高さ:"))
        self.window_height_input = QLineEdit()
        self.window_height_input.setValidator(QIntValidator(150, 800))
        layout.addWidget(self.window_height_input)

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
        QMessageBox.information(self, "設定完了", "設定を保存しました。\n変更を適用するにはアプリケーションを再起動してください。")
        super().accept()
