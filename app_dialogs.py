import os
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QLineEdit, QListWidget, QDialogButtonBox, QFileDialog,
    QMessageBox
)
from PyQt6.QtGui import QIntValidator
from config_manager import ConfigManager


class ExcludeItemDialog(QDialog):
    def __init__(self, title, item_label, config_section, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.config_section = config_section

        layout = QVBoxLayout()

        self.input_field = QLineEdit()
        layout.addWidget(QLabel(f"{item_label}を入力:"))
        layout.addWidget(self.input_field)

        self.item_list = QListWidget()
        layout.addWidget(QLabel(f"登録済み{item_label}:"))
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

        self.config = ConfigManager()
        self.load_items()

    def load_items(self):
        if self.config_section in self.config.config:
            items = self.config.config[self.config_section].get('list', '').split(',')
            for item in items:
                if item.strip():
                    self.item_list.addItem(item.strip())

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
            items.append(self.item_list.item(i).text())

        if self.config_section not in self.config.config:
            self.config.config[self.config_section] = {}
        self.config.config[self.config_section]['list'] = ','.join(items)
        self.config.save_config()
        super().accept()


class ExcludeDocsDialog(ExcludeItemDialog):
    def __init__(self, parent=None):
        super().__init__("除外する文書名", "除外する文書名", "ExcludeDocs", parent)


class ExcludeDoctorsDialog(ExcludeItemDialog):
    def __init__(self, parent=None):
        super().__init__("除外する医師名", "除外する医師名", "ExcludeDoctors", parent)


class FolderPathDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォルダの場所")
        self.setModal(True)

        config = ConfigManager()
        dialog_width, dialog_height = config.get_folder_dialog_size()
        self.resize(dialog_width, dialog_height)

        layout = QVBoxLayout()

        # ダウンロードパス設定
        layout.addWidget(QLabel("ダウンロードフォルダ:"))
        downloads_layout = QHBoxLayout()
        self.downloads_path = QLineEdit()
        downloads_layout.addWidget(self.downloads_path)
        downloads_browse = QPushButton("参照...")
        downloads_browse.clicked.connect(lambda: self.browse_folder('downloads'))
        downloads_layout.addWidget(downloads_browse)
        downloads_open = QPushButton("開く")
        downloads_open.clicked.connect(lambda: self.open_folder(self.downloads_path.text()))
        downloads_layout.addWidget(downloads_open)
        layout.addLayout(downloads_layout)

        # Excelファイルパス設定
        layout.addWidget(QLabel("Excelファイルパス:"))
        excel_layout = QHBoxLayout()
        self.excel_path = QLineEdit()
        excel_layout.addWidget(self.excel_path)
        excel_browse = QPushButton("参照...")
        excel_browse.clicked.connect(lambda: self.browse_folder('excel'))
        excel_layout.addWidget(excel_browse)
        excel_open = QPushButton("開く")
        excel_open.clicked.connect(lambda: self.open_folder(str(Path(self.excel_path.text()).parent)))
        excel_layout.addWidget(excel_open)
        layout.addLayout(excel_layout)

        # バックアップフォルダ設定
        layout.addWidget(QLabel("バックアップフォルダ:"))
        backup_layout = QHBoxLayout()
        self.backup_path = QLineEdit()
        backup_layout.addWidget(self.backup_path)
        backup_browse = QPushButton("参照...")
        backup_browse.clicked.connect(lambda: self.browse_folder('backup'))
        backup_layout.addWidget(backup_browse)
        backup_open = QPushButton("開く")
        backup_open.clicked.connect(lambda: self.open_folder(self.backup_path.text()))
        backup_layout.addWidget(backup_open)
        layout.addLayout(backup_layout)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_paths()

    def open_folder(self, path: str) -> None:
        """指定されたパスのフォルダをエクスプローラーで開く"""
        if path and os.path.exists(path):
            os.startfile(path)
        else:
            QMessageBox.warning(self, "警告", "指定されたフォルダが存在しません。")

    def load_paths(self):
        self.downloads_path.setText(self.config.get_downloads_path())
        self.excel_path.setText(self.config.get_excel_path())
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
            file, _ = QFileDialog.getOpenFileName(
                self,
                "Excelファイルの選択",
                self.excel_path.text(),
                "Excel Files (*.xlsx *.xlsm)"
            )
            if file:
                self.excel_path.setText(file)

    def accept(self):
        # 設定の保存
        self.config.set_downloads_path(self.downloads_path.text())
        self.config.set_excel_path(self.excel_path.text())
        self.config.set_backup_path(self.backup_path.text())
        super().accept()


class AppearanceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォントとウインドウサイズ")
        self.setModal(True)

        layout = QVBoxLayout()

        # フォントサイズ設定
        layout.addWidget(QLabel("フォントサイズ:"))
        self.font_size_input = QLineEdit()
        self.font_size_input.setValidator(QIntValidator(6, 72))
        layout.addWidget(self.font_size_input)

        # ウィンドウサイズ設定
        layout.addWidget(QLabel("ウィンドウの幅:"))
        self.window_width_input = QLineEdit()
        self.window_width_input.setValidator(QIntValidator(200, 1000))
        layout.addWidget(self.window_width_input)

        layout.addWidget(QLabel("ウィンドウの高さ:"))
        self.window_height_input = QLineEdit()
        self.window_height_input.setValidator(QIntValidator(150, 800))
        layout.addWidget(self.window_height_input)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_settings()

    def load_settings(self):
        self.font_size_input.setText(str(self.config.get_font_size()))
        window_size = self.config.get_window_size()
        self.window_width_input.setText(str(window_size[0]))
        self.window_height_input.setText(str(window_size[1]))

    def accept(self):
        # 設定の保存
        self.config.set_font_size(int(self.font_size_input.text()))
        self.config.set_window_size(
            int(self.window_width_input.text()),
            int(self.window_height_input.text())
        )
        QMessageBox.information(self, "設定完了", "設定を保存しました。\n変更を適用するにはアプリケーションを再起動してください。")
        super().accept()
