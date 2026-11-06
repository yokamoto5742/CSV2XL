import datetime
import sqlite3
from collections.abc import Callable
from pathlib import Path

from PyQt6.QtCore import QPoint, QSize, Qt
from PyQt6.QtGui import QCloseEvent, QGuiApplication
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QLineEdit, QMessageBox, QFileDialog, QSplitter
)

from app.dialogs import (
    AddDocumentDialog, ExcludeDocsDialog, ExcludeDoctorsDialog, StaffNamesDialog, FolderPathDialog, FontSizeDialog,
    SettingsDialog
)
from app.document_window import DocumentListWidget
from utils.config_manager import ConfigManager
from services.csv_import import import_csv_to_database
from services.database import connect
from services.file_manager import backup_database_daily, cleanup_old_backup_files
from services.list_repository import ListRepository
from services.xlsm_import import import_xlsm_to_database
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
# ボタンの文字の左右に取る余白(CSV_BUTTON_STYLE の padding と border を含む)
BUTTON_TEXT_PADDING = 24
# 取り込めなかった行の一覧に出す最大の行数
MAX_ERROR_LINES = 50


def apply_app_font_size(point_size: int) -> None:
    """一覧・左パネル・ダイアログなど、アプリ全体のフォントサイズを変える"""
    font = QApplication.font()
    font.setPointSize(point_size)
    QApplication.setFont(font)


class CompactButton(QPushButton):
    """Windows の既定の最小幅(約80px)より狭くできるボタン。左パネルを文字の幅まで狭くするために使う"""

    def sizeHint(self) -> QSize:
        width = self.fontMetrics().horizontalAdvance(self.text()) + BUTTON_TEXT_PADDING
        return QSize(width, super().sizeHint().height())


def add_button(layout: QVBoxLayout | QHBoxLayout, text: str, slot: Callable[[], object]) -> QPushButton:
    button = CompactButton(text)
    button.clicked.connect(slot)
    layout.addWidget(button)
    return button


class MainWindow(QMainWindow):
    """左に操作ボタン、右に担当一覧を並べた1画面のウィンドウ"""

    def __init__(self) -> None:
        config = ConfigManager()
        # 作成済みのウィンドウにはフォントの変更がすぐには反映されないため、ウィンドウを作る前に設定する
        apply_app_font_size(config.get_font_size())
        super().__init__()
        self.config = config
        self.connection: sqlite3.Connection | None = None
        self.document_widget: DocumentListWidget | None = None

        cleanup_old_backup_files()

        self.restore_window_geometry()
        title = f"医療文書担当一覧 v{__version__}"
        self.setWindowTitle(title)
        self.setStyleSheet("QMainWindow { border: 5px solid darkgreen; }")

        # 左パネルと一覧の境目をドラッグして幅を変えられるようにする
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        side_panel = QWidget()
        side_panel.setLayout(self.create_side_panel())
        self.splitter.addWidget(side_panel)
        self.setCentralWidget(self.splitter)
        self.side_panel_width = self.config.get_side_panel_width()
        self.column_widths = self.config.get_column_widths()

        self.open_document_list()
        self.backup_daily()

    def create_side_panel(self) -> QVBoxLayout:
        panel = QVBoxLayout()
        self.add_search_section(panel)
        panel.addSpacing(SECTION_SPACING)

        panel.addWidget(QLabel("Papyrus"))
        add_button(panel, "CSV取り込み", self.import_csv).setStyleSheet(CSV_BUTTON_STYLE)
        panel.addSpacing(SECTION_SPACING)

        panel.addWidget(QLabel("データ"))
        add_button(panel, "行の追加", self.add_document)
        add_button(panel, "行の削除", self.delete_document)
        panel.addSpacing(SECTION_SPACING)

        panel.addWidget(QLabel("除外項目"))
        add_button(panel, "文書名", self.show_exclude_docs_dialog)
        add_button(panel, "医師名", self.show_exclude_doctors_dialog)
        panel.addSpacing(SECTION_SPACING)

        add_button(panel, "設定", self.show_settings_dialog)
        panel.addSpacing(SECTION_SPACING)

        add_button(panel, "閉じる", self.close)
        panel.addStretch()
        return panel

    def add_search_section(self, panel: QVBoxLayout) -> None:
        panel.addWidget(QLabel("患者ID検索"))
        self.patient_id_input = QLineEdit()
        self.patient_id_input.returnPressed.connect(self.search)
        # 左パネルを狭くするため、ボタンを入力欄の横と次の行に分ける
        search_input = QHBoxLayout()
        search_input.addWidget(self.patient_id_input)
        add_button(search_input, "検索", self.search)
        panel.addLayout(search_input)

        add_button(panel, "絞り込み解除", self.clear_search)

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
        document_widget.set_column_widths(self.column_widths)
        self.splitter.addWidget(document_widget)
        self.splitter.setStretchFactor(1, 1)
        # 一覧側は残りの幅いっぱいに広がるので、大きめの値を渡しておく
        self.splitter.setSizes([self.side_panel_width, MAX_TABLE_WIDTH])
        self.document_widget = document_widget

    def close_connection(self) -> None:
        if self.document_widget is not None:
            self.side_panel_width = self.splitter.sizes()[0]
            self.column_widths = self.document_widget.column_widths()
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

    def clear_search(self) -> None:
        self.patient_id_input.clear()
        if self.document_widget is not None:
            self.document_widget.clear_search()

    def import_csv(self) -> None:
        try:
            added_count = import_csv_to_database(self.get_connection())
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")
            return
        QMessageBox.information(self, "取り込み完了", f"{added_count}件を追加しました。")
        self.open_document_list()

    def import_xlsm(self) -> None:
        try:
            result = import_xlsm_to_database(self.get_connection())
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"xlsmファイルの取り込み中にエラーが発生しました:\n{str(e)}")
            return
        message = f"{result.file_name}\n{result.added_count}件を追加し、{result.updated_count}件の空欄を埋めました。"
        if result.errors:
            message += f"\n\n取り込めなかった行({len(result.errors)}件):\n" + "\n".join(result.errors[:MAX_ERROR_LINES])
            if len(result.errors) > MAX_ERROR_LINES:
                message += "\n…"
        QMessageBox.information(self, "取り込み完了", message)
        self.open_document_list()

    def backup_daily(self) -> None:
        """起動時に、その日のバックアップがまだ無ければ作る"""
        if self.connection is None:
            return
        try:
            backup_database_daily(self.connection)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"データベースのバックアップに失敗しました:\n{str(e)}")

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

    def show_settings_dialog(self) -> None:
        dialog = SettingsDialog([
            ("xlsm取り込み", self.import_xlsm),
            ("CSV出力", self.export_csv),
            ("フォルダパス", self.show_folder_path_dialog),
            ("フォントサイズ", self.show_font_size_dialog),
            ("担当者名", self.show_staff_names_dialog),
        ], self)
        dialog.exec()

    def show_font_size_dialog(self) -> None:
        """入力したフォントサイズを保存し、すぐにアプリ全体に反映する"""
        config = ConfigManager()
        dialog = FontSizeDialog(config.get_font_size(), self)
        if not dialog.exec():
            return
        font_size = dialog.font_size()
        if font_size is None:
            return
        config.set_font_size(font_size)
        apply_app_font_size(font_size)
        # スタイルシートを設定したウィンドウの子にはアプリのフォントが届かないため、設定し直して反映させる
        self.setStyleSheet(self.styleSheet())
        if self.document_widget is not None:
            self.document_widget.apply_font_size()

    def restore_window_geometry(self) -> None:
        """前回終了時の位置と大きさにする。その位置がどの画面にも入らなければ、大きさだけを戻す"""
        self.resize(*self.config.get_window_size())
        position = self.config.get_window_position()
        if position is not None and QGuiApplication.screenAt(QPoint(*position)) is not None:
            self.setGeometry(*position, self.width(), self.height())

    def save_window_geometry(self) -> None:
        # 最大化しているときは、最大化を解除したときの位置と大きさを保存する
        maximized = self.isMaximized()
        rect = self.normalGeometry() if maximized else self.geometry()
        self.config.set_window_geometry(rect.x(), rect.y(), rect.width(), rect.height(), maximized)

    def show_folder_path_dialog(self) -> None:
        dialog = FolderPathDialog(self)
        if dialog.exec():
            # DB の場所が変わった可能性があるため、接続し直して一覧を作り直す
            self.close_connection()
            self.open_document_list()

    def closeEvent(self, event: QCloseEvent | None) -> None:
        self.close_connection()
        # フォルダパスのダイアログが保存した値を古い内容で上書きしないよう、読み直してから保存する
        self.config = ConfigManager()
        self.config.set_side_panel_width(self.side_panel_width)
        self.config.set_column_widths(self.column_widths)
        self.save_window_geometry()
        super().closeEvent(event)
