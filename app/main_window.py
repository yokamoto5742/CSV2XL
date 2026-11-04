import sqlite3
from collections.abc import Callable

from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QLineEdit, QMessageBox
)

from app.dialogs import ExcludeDocsDialog, ExcludeDoctorsDialog, StaffNamesDialog, AppearanceDialog, FolderPathDialog
from app.document_window import DocumentListWidget
from utils.config_manager import ConfigManager
from services.csv_import import import_csv_to_database
from services.database import connect
from services.file_manager import cleanup_old_backup_files
from services.list_repository import ListRepository
from app import __version__

CSV_BUTTON_STYLE = """
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
"""
SECTION_SPACING = 12


def add_button(layout: QVBoxLayout | QHBoxLayout, text: str, slot: Callable[[], object]) -> QPushButton:
    button = QPushButton(text)
    button.clicked.connect(slot)
    layout.addWidget(button)
    return button


class MainWindow(QMainWindow):
    """左に操作ボタン、右に担当一覧を並べた1画面のウィンドウ"""

    def __init__(self) -> None:
        super().__init__()
        self.config = ConfigManager()
        self.connection: sqlite3.Connection | None = None
        self.document_widget: DocumentListWidget | None = None

        cleanup_old_backup_files()

        font = self.font()
        font.setPointSize(self.config.get_font_size())
        self.setFont(font)
        # 起動時は最大化する。この大きさは最大化を解除したときに使う
        self.resize(*self.config.get_window_size())
        title = f"医療文書担当一覧 v{__version__}"
        self.setWindowTitle(title)
        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        main_widget = QWidget()
        self.content_layout = QHBoxLayout(main_widget)
        self.content_layout.addLayout(self.create_side_panel(title))
        self.setCentralWidget(main_widget)

        self.open_document_list()

    def create_side_panel(self, title: str) -> QVBoxLayout:
        panel = QVBoxLayout()
        panel.addWidget(QLabel(title))
        panel.addSpacing(SECTION_SPACING)
        self.add_search_section(panel)
        panel.addSpacing(SECTION_SPACING)

        panel.addWidget(QLabel("Papyrus"))
        add_button(panel, "CSV取り込み", self.import_csv).setStyleSheet(CSV_BUTTON_STYLE)
        panel.addSpacing(SECTION_SPACING)

        panel.addWidget(QLabel("除外項目"))
        add_button(panel, "文書名", self.show_exclude_docs_dialog)
        add_button(panel, "医師名", self.show_exclude_doctors_dialog)
        panel.addSpacing(SECTION_SPACING)

        add_button(panel, "担当者名", self.show_staff_names_dialog)
        add_button(panel, "画面サイズ", self.show_appearance_dialog)
        add_button(panel, "フォルダパス", self.show_folder_path_dialog)
        panel.addSpacing(SECTION_SPACING)

        add_button(panel, "閉じる", self.close)
        panel.addStretch()
        return panel

    def add_search_section(self, panel: QVBoxLayout) -> None:
        panel.addWidget(QLabel("患者ID検索"))
        self.patient_id_input = QLineEdit()
        self.patient_id_input.returnPressed.connect(self.search)
        panel.addWidget(self.patient_id_input)

        search_buttons = QHBoxLayout()
        add_button(search_buttons, "検索", self.search)
        add_button(search_buttons, "前へ", self.show_previous_hit)
        add_button(search_buttons, "次へ", self.show_next_hit)
        panel.addLayout(search_buttons)

        self.search_result_label = QLabel("")
        panel.addWidget(self.search_result_label)

    def get_connection(self) -> sqlite3.Connection:
        if self.connection is None:
            self.connection = connect(ConfigManager().get_database_path())
        return self.connection

    def open_document_list(self) -> None:
        """担当一覧を表示する。表示済みなら DB から読み直す"""
        if self.document_widget is not None:
            self.document_widget.reload()
            return
        try:
            document_widget = DocumentListWidget(self.get_connection())
            document_widget.load_initial()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"担当一覧を開けませんでした:\n{str(e)}")
            return
        document_widget.search_status_changed.connect(self.search_result_label.setText)
        self.content_layout.addWidget(document_widget, stretch=1)
        self.document_widget = document_widget

    def close_connection(self) -> None:
        if self.document_widget is not None:
            self.document_widget.refresh_timer.stop()
            self.content_layout.removeWidget(self.document_widget)
            self.document_widget.deleteLater()
            self.document_widget = None
            self.search_result_label.setText("")
        if self.connection is not None:
            self.connection.close()
            self.connection = None

    def search(self) -> None:
        if self.document_widget is not None:
            self.document_widget.search(self.patient_id_input.text())

    def show_next_hit(self) -> None:
        if self.document_widget is not None:
            self.document_widget.show_next_hit()

    def show_previous_hit(self) -> None:
        if self.document_widget is not None:
            self.document_widget.show_previous_hit()

    def import_csv(self) -> None:
        try:
            added_count = import_csv_to_database(self.get_connection())
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")
            return
        QMessageBox.information(self, "取り込み完了", f"{added_count}件を追加しました。")
        self.open_document_list()

    def show_list_dialog(self, dialog_class: type[ExcludeDocsDialog | ExcludeDoctorsDialog | StaffNamesDialog]) -> None:
        try:
            dialog = dialog_class(ListRepository(self.get_connection()), self)
            dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"設定の読み書きに失敗しました:\n{str(e)}")

    def show_exclude_docs_dialog(self) -> None:
        self.show_list_dialog(ExcludeDocsDialog)

    def show_exclude_doctors_dialog(self) -> None:
        self.show_list_dialog(ExcludeDoctorsDialog)

    def show_staff_names_dialog(self) -> None:
        self.show_list_dialog(StaffNamesDialog)
        if self.document_widget is not None:
            self.document_widget.reload()

    def show_appearance_dialog(self) -> None:
        dialog = AppearanceDialog(self)
        dialog.exec()

    def show_folder_path_dialog(self) -> None:
        dialog = FolderPathDialog(self)
        if dialog.exec():
            # DB の場所が変わった可能性があるため、接続し直して一覧を作り直す
            self.close_connection()
            self.open_document_list()

    def closeEvent(self, event: QCloseEvent | None) -> None:
        self.close_connection()
        super().closeEvent(event)
