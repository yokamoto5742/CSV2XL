import datetime
import sqlite3
from collections.abc import Callable
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QLineEdit, QMessageBox, QFileDialog, QSplitter
)

from app.dialogs import (
    AddDocumentDialog, ExcludeDocsDialog, ExcludeDoctorsDialog, StaffNamesDialog, AppearanceDialog, FolderPathDialog
)
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
MAX_TABLE_WIDTH = 10000


def apply_app_font_size(point_size: int) -> None:
    """一覧・左パネル・ダイアログなど、アプリ全体のフォントサイズを変える"""
    font = QApplication.font()
    font.setPointSize(point_size)
    QApplication.setFont(font)


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

        apply_app_font_size(self.config.get_font_size())
        # 最大化して起動するときは、この大きさを最大化を解除したときに使う
        self.resize(*self.config.get_window_size())
        title = f"医療文書担当一覧 v{__version__}"
        self.setWindowTitle(title)
        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        # 左パネルと一覧の境目をドラッグして幅を変えられるようにする
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        side_panel = QWidget()
        side_panel.setLayout(self.create_side_panel(title))
        self.splitter.addWidget(side_panel)
        self.setCentralWidget(self.splitter)
        self.side_panel_width = self.config.get_side_panel_width()

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

        panel.addWidget(QLabel("データ"))
        add_button(panel, "行の追加", self.add_document)
        add_button(panel, "行の削除", self.delete_document)
        add_button(panel, "CSV出力", self.export_csv)
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
        self.splitter.addWidget(document_widget)
        self.splitter.setStretchFactor(1, 1)
        # 一覧側は残りの幅いっぱいに広がるので、大きめの値を渡しておく
        self.splitter.setSizes([self.side_panel_width, MAX_TABLE_WIDTH])
        self.document_widget = document_widget

    def close_connection(self) -> None:
        if self.document_widget is not None:
            self.side_panel_width = self.splitter.sizes()[0]
            self.document_widget.refresh_timer.stop()
            self.document_widget.setParent(None)
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

    def add_document(self) -> None:
        if self.document_widget is None:
            return
        dialog = AddDocumentDialog(self)
        if dialog.exec():
            new_document = dialog.new_document()
            if new_document is not None:
                self.document_widget.add_document(new_document)

    def delete_document(self) -> None:
        if self.document_widget is not None:
            self.document_widget.delete_selected()

    def export_csv(self) -> None:
        if self.document_widget is None:
            return
        file_name = f"医療文書担当一覧_{datetime.datetime.now():%Y%m%d%H%M}.csv"
        default_path = str(Path(self.config.get_downloads_path()) / file_name)
        path, _ = QFileDialog.getSaveFileName(self, "CSV出力", default_path, "CSV ファイル (*.csv)")
        if not path:
            return
        try:
            self.document_widget.export_csv(path)
        except OSError as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの出力に失敗しました:\n{str(e)}")
            return
        QMessageBox.information(self, "出力完了", "CSVファイルを出力しました。")

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
        if dialog.exec():
            self.apply_appearance()

    def apply_appearance(self) -> None:
        """保存した外観の設定を再起動せずに反映する"""
        self.config = ConfigManager()
        apply_app_font_size(self.config.get_font_size())
        if self.document_widget is not None:
            self.document_widget.apply_font_size()
        if not self.isMaximized():
            self.resize(*self.config.get_window_size())

    def show_folder_path_dialog(self) -> None:
        dialog = FolderPathDialog(self)
        if dialog.exec():
            # DB の場所が変わった可能性があるため、接続し直して一覧を作り直す
            self.close_connection()
            self.open_document_list()

    def closeEvent(self, event: QCloseEvent | None) -> None:
        self.close_connection()
        self.config.set_side_panel_width(self.side_panel_width)
        super().closeEvent(event)
