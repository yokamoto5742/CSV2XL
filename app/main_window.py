import sqlite3

from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QMessageBox
)

from app.dialogs import ExcludeDocsDialog, ExcludeDoctorsDialog, StaffNamesDialog, AppearanceDialog, FolderPathDialog
from app.document_window import DocumentListWindow
from utils.config_manager import ConfigManager
from services.csv_import import import_csv_to_database
from services.database import connect
from services.file_manager import cleanup_old_backup_files
from services.list_repository import ListRepository
from app import __version__


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager()
        # DB はネットワーク上に置くことがあるため、使うときに初めて接続する
        self.connection: sqlite3.Connection | None = None
        self.document_window: DocumentListWindow | None = None

        cleanup_old_backup_files()

        font = self.font()
        font.setPointSize(self.config.get_font_size())
        self.setFont(font)
        window_size = self.config.get_window_size()
        self.setFixedSize(*window_size)
        self.setWindowTitle(f"CSV取込アプリ v{__version__}")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        layout = QVBoxLayout()

        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        title_label = QLabel("Papyrus書類受付リスト")
        layout.addWidget(title_label)

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

        document_list_button = QPushButton("担当一覧を開く")
        document_list_button.clicked.connect(self.show_document_window)
        layout.addWidget(document_list_button)

        settings_label = QLabel("設定")
        layout.addWidget(settings_label)

        exclude_docs_button = QPushButton("除外する文書名")
        exclude_docs_button.clicked.connect(self.show_exclude_docs_dialog)
        layout.addWidget(exclude_docs_button)

        exclude_doctors_button = QPushButton("除外する医師名")
        exclude_doctors_button.clicked.connect(self.show_exclude_doctors_dialog)
        layout.addWidget(exclude_doctors_button)

        staff_names_button = QPushButton("担当者名")
        staff_names_button.clicked.connect(self.show_staff_names_dialog)
        layout.addWidget(staff_names_button)

        appearance_button = QPushButton("フォントとウインドウサイズ")
        appearance_button.clicked.connect(self.show_appearance_dialog)
        layout.addWidget(appearance_button)

        folder_path_button = QPushButton("フォルダの場所")
        folder_path_button.clicked.connect(self.show_folder_path_dialog)
        layout.addWidget(folder_path_button)

        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        main_widget.setLayout(layout)

    def get_connection(self) -> sqlite3.Connection:
        if self.connection is None:
            self.connection = connect(ConfigManager().get_database_path())
        return self.connection

    def close_connection(self) -> None:
        if self.document_window is not None:
            self.document_window.close()
            self.document_window = None
        if self.connection is not None:
            self.connection.close()
            self.connection = None

    def import_csv(self):
        try:
            added_count = import_csv_to_database(self.get_connection())
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")
            return
        QMessageBox.information(self, "取り込み完了", f"{added_count}件を追加しました。")
        self.show_document_window()

    def show_document_window(self):
        try:
            if self.document_window is None:
                self.document_window = DocumentListWindow(self.get_connection(), self)
                self.document_window.setFont(self.font())
                self.document_window.load_initial()
            else:
                self.document_window.reload()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"担当一覧を開けませんでした:\n{str(e)}")
            return
        self.document_window.showMaximized()
        self.document_window.activateWindow()

    def show_list_dialog(self, dialog_class: type[ExcludeDocsDialog | ExcludeDoctorsDialog | StaffNamesDialog]):
        try:
            dialog = dialog_class(ListRepository(self.get_connection()), self)
            dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"設定の読み書きに失敗しました:\n{str(e)}")

    def show_exclude_docs_dialog(self):
        self.show_list_dialog(ExcludeDocsDialog)

    def show_exclude_doctors_dialog(self):
        self.show_list_dialog(ExcludeDoctorsDialog)

    def show_staff_names_dialog(self):
        self.show_list_dialog(StaffNamesDialog)
        if self.document_window is not None:
            self.document_window.reload()

    def show_appearance_dialog(self):
        dialog = AppearanceDialog(self)
        dialog.exec()

    def show_folder_path_dialog(self):
        dialog = FolderPathDialog(self)
        if dialog.exec():
            # DB の場所が変わった可能性があるため、次に使うときに接続し直す
            self.close_connection()

    def closeEvent(self, event: QCloseEvent | None) -> None:
        self.close_connection()
        super().closeEvent(event)
