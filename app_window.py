import os
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,QHBoxLayout,
    QPushButton, QLabel, QDialog, QLineEdit,
    QListWidget, QDialogButtonBox, QFileDialog,
    QMessageBox
)
from PyQt6.QtGui import QIntValidator
from PyQt6.QtCore import Qt, QTimer
from config_manager import ConfigManager
from service_csv_excel_transfer import transfer_csv_to_excel
from coordinate_tracker import CoordinateTracker
from version import VERSION


class ExcludeDocsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("除外する文書名")
        self.setModal(True)

        layout = QVBoxLayout()

        # 文書名入力フィールド
        self.input_field = QLineEdit()
        layout.addWidget(QLabel("除外する文書名を入力:"))
        layout.addWidget(self.input_field)

        # 登録済み文書名リスト
        self.doc_list = QListWidget()
        layout.addWidget(QLabel("登録済み文書名:"))
        layout.addWidget(self.doc_list)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # 追加ボタン
        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_document)

        # 削除ボタン
        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_documents()

    def load_documents(self):
        if 'ExcludeDocs' in self.config.config:
            docs = self.config.config['ExcludeDocs'].get('list', '').split(',')
            for doc in docs:
                if doc.strip():
                    self.doc_list.addItem(doc.strip())

    def add_document(self):
        doc_name = self.input_field.text().strip()
        if doc_name:
            self.doc_list.addItem(doc_name)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.doc_list.currentItem()
        if current_item:
            self.doc_list.takeItem(self.doc_list.row(current_item))

    def accept(self):
        # 設定の保存
        docs = []
        for i in range(self.doc_list.count()):
            docs.append(self.doc_list.item(i).text())

        if 'ExcludeDocs' not in self.config.config:
            self.config.config['ExcludeDocs'] = {}
        self.config.config['ExcludeDocs']['list'] = ','.join(docs)
        self.config.save_config()
        super().accept()


class ExcludeDoctorsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("除外する医師名")
        self.setModal(True)

        layout = QVBoxLayout()

        # 医師名入力フィールド
        self.input_field = QLineEdit()
        layout.addWidget(QLabel("除外する医師名を入力:"))
        layout.addWidget(self.input_field)

        # 登録済み医師名リスト
        self.doc_list = QListWidget()
        layout.addWidget(QLabel("登録済み医師名:"))
        layout.addWidget(self.doc_list)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # 追加ボタン
        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_doctor)

        # 削除ボタン
        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_doctors()

    def load_doctors(self):
        if 'ExcludeDoctors' in self.config.config:
            docs = self.config.config['ExcludeDoctors'].get('list', '').split(',')
            for doc in docs:
                if doc.strip():
                    self.doc_list.addItem(doc.strip())

    def add_doctor(self):
        doc_name = self.input_field.text().strip()
        if doc_name:
            self.doc_list.addItem(doc_name)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.doc_list.currentItem()
        if current_item:
            self.doc_list.takeItem(self.doc_list.row(current_item))

    def accept(self):
        # 設定の保存
        docs = []
        for i in range(self.doc_list.count()):
            docs.append(self.doc_list.item(i).text())

        if 'ExcludeDoctors' not in self.config.config:
            self.config.config['ExcludeDoctors'] = {}
        self.config.config['ExcludeDoctors']['list'] = ','.join(docs)
        self.config.save_config()
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
        super().accept()
        self.config.set_backup_path(self.backup_path.text())


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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.tracker = CoordinateTracker()
        self.config = ConfigManager()
        font = self.font()
        font.setPointSize(self.config.get_font_size())
        self.setFont(font)
        window_size = self.config.get_window_size()
        self.setFixedSize(*window_size)
        self.setWindowTitle(f"CSV取込アプリ v{VERSION}")

        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # レイアウト
        layout = QVBoxLayout()

        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        # タイトル
        title_label = QLabel("Papyrus書類受付リスト")
        layout.addWidget(title_label)

        # CSVファイル取り込みボタン
        # CSVファイル取り込みボタン
        csv_button = QPushButton("CSVファイル取り込み")
        csv_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border: 2px solid #45a049;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        csv_button.clicked.connect(self.import_csv)
        layout.addWidget(csv_button)

        settings_label = QLabel("設定")
        layout.addWidget(settings_label)

        exclude_docs_button = QPushButton("除外する文書名")
        exclude_docs_button.clicked.connect(self.show_exclude_docs_dialog)
        layout.addWidget(exclude_docs_button)

        exclude_doctors_button = QPushButton("除外する医師名")
        exclude_doctors_button.clicked.connect(self.show_exclude_doctors_dialog)
        layout.addWidget(exclude_doctors_button)

        appearance_button = QPushButton("フォントとウインドウサイズ")
        appearance_button.clicked.connect(self.show_appearance_dialog)
        layout.addWidget(appearance_button)

        coordinate_button = QPushButton("画面の座標表示")
        coordinate_button.clicked.connect(self.show_coordinate_tracker)
        layout.addWidget(coordinate_button)

        folder_path_button = QPushButton("フォルダの場所")
        folder_path_button.clicked.connect(self.show_folder_path_dialog)
        layout.addWidget(folder_path_button)

        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        main_widget.setLayout(layout)

    def import_csv(self):
        try:
            transfer_csv_to_excel()
        except Exception as e:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")

    def show_exclude_docs_dialog(self):
        dialog = ExcludeDocsDialog(self)
        dialog.exec()

    def show_exclude_doctors_dialog(self):
        dialog = ExcludeDoctorsDialog(self)
        dialog.exec()

    def show_appearance_dialog(self):
        dialog = AppearanceDialog(self)
        dialog.exec()

    def show_coordinate_tracker(self):
        self.tracker.show()

    def show_folder_path_dialog(self):
        dialog = FolderPathDialog(self)
        dialog.exec()
