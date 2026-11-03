import sqlite3

from PyQt6.QtCore import QAbstractItemModel, QModelIndex, QPersistentModelIndex, Qt, QTimer
from PyQt6.QtWidgets import (
    QAbstractItemView, QComboBox, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QStyledItemDelegate, QStyleOptionViewItem, QTableView,
    QVBoxLayout, QWidget
)

from app.document_table_model import STAFF_COLUMN, DocumentTableModel
from services.document_repository import DocumentRepository
from services.list_repository import STAFF, ListRepository
from services.validators import parse_patient_id

# Excel 版の列幅に近い値(ピクセル)
COLUMN_WIDTHS = [100, 80, 120, 320, 120, 100, 100, 110, 260]
AUTO_REFRESH_INTERVAL_MS = 5000


class StaffComboDelegate(QStyledItemDelegate):
    """担当者名をドロップダウンで選ぶ"""

    def __init__(self, window: 'DocumentListWindow') -> None:
        super().__init__(window)
        self.window = window

    def createEditor(self, parent: QWidget | None, option: QStyleOptionViewItem,
                     index: QModelIndex | QPersistentModelIndex) -> QWidget:
        combo = QComboBox(parent)
        # 選んだらすぐ確定する
        combo.activated.connect(lambda: self.commit_and_close(combo))
        return combo

    def commit_and_close(self, combo: QComboBox) -> None:
        self.commitData.emit(combo)
        self.closeEditor.emit(combo)

    def setEditorData(self, editor: QWidget | None, index: QModelIndex | QPersistentModelIndex) -> None:
        if not isinstance(editor, QComboBox):
            return
        current = str(index.data(Qt.ItemDataRole.EditRole) or '')
        names = ['', *self.window.staff_names]
        # リストから外れた担当者が入力済みの行でも、今の値を消さない
        if current not in names:
            names.append(current)
        editor.addItems(names)
        editor.setCurrentText(current)

    def setModelData(self, editor: QWidget | None, model: QAbstractItemModel | None,
                     index: QModelIndex | QPersistentModelIndex) -> None:
        if isinstance(editor, QComboBox) and model is not None:
            model.setData(model.index(index.row(), index.column()), editor.currentText(), Qt.ItemDataRole.EditRole)


class DocumentListWindow(QMainWindow):
    """医療文書担当一覧の表示・編集・患者ID検索"""

    def __init__(self, connection: sqlite3.Connection, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("医療文書担当一覧")
        self.list_repository = ListRepository(connection)
        self.repository = DocumentRepository(connection)
        self.model = DocumentTableModel(self.repository)
        self.staff_names: list[str] = []
        self.search_hits: list[int] = []
        self.search_position = 0
        self.last_data_version = -1

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        layout.addLayout(self.create_toolbar())
        self.table = self.create_table()
        layout.addWidget(self.table)
        self.setCentralWidget(central_widget)

        # セルの確定処理の途中で表を読み直さないよう、確定が終わってから処理する
        self.model.error_occurred.connect(lambda message: QTimer.singleShot(0, lambda: self.show_error(message)))
        self.model.conflict_detected.connect(lambda: QTimer.singleShot(0, self.handle_conflict))
        self.model.modelReset.connect(self.clear_search)
        self.model.layoutChanged.connect(self.clear_search)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_if_changed)
        self.refresh_timer.start(AUTO_REFRESH_INTERVAL_MS)

    def create_toolbar(self) -> QHBoxLayout:
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("患者ID:"))
        self.patient_id_input = QLineEdit()
        self.patient_id_input.setFixedWidth(120)
        self.patient_id_input.returnPressed.connect(self.search)
        toolbar.addWidget(self.patient_id_input)

        for text, slot in (("検索", self.search), ("前へ", self.show_previous_hit), ("次へ", self.show_next_hit)):
            button = QPushButton(text)
            button.clicked.connect(slot)
            toolbar.addWidget(button)

        self.search_result_label = QLabel("")
        toolbar.addWidget(self.search_result_label)
        toolbar.addStretch()

        reload_button = QPushButton("再読み込み")
        reload_button.clicked.connect(self.reload)
        toolbar.addWidget(reload_button)
        return toolbar

    def create_table(self) -> QTableView:
        table = QTableView()
        table.setModel(self.model)
        table.setItemDelegateForColumn(STAFF_COLUMN, StaffComboDelegate(self))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked
                              | QAbstractItemView.EditTrigger.EditKeyPressed
                              | QAbstractItemView.EditTrigger.AnyKeyPressed)
        table.setAlternatingRowColors(True)
        header = table.horizontalHeader()
        if header is not None:
            # 起動直後は既定の並び(預り日 → 診療科 → 患者ID)にするため、並び替えの列を未選択にしておく
            header.setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
            header.setStretchLastSection(True)
        table.setSortingEnabled(True)
        for column, width in enumerate(COLUMN_WIDTHS):
            table.setColumnWidth(column, width)
        return table

    def reload(self) -> None:
        """DB から読み直す(スクロール位置は保つ)"""
        scroll_bar = self.table.verticalScrollBar()
        scroll_value = scroll_bar.value() if scroll_bar is not None else 0
        try:
            self.last_data_version = self.repository.data_version()
            self.staff_names = self.list_repository.get_items(STAFF)
            self.model.load()
        except sqlite3.Error as e:
            self.show_error(f"データベースの読み込みに失敗しました。\n{e}")
            return
        if scroll_bar is not None:
            scroll_bar.setValue(scroll_value)

    def load_initial(self) -> None:
        """初回表示。最新の行が見えるよう末尾までスクロールする"""
        self.reload()
        self.table.scrollToBottom()

    def refresh_if_changed(self) -> None:
        """他の端末で変更があれば読み直す(編集中は待つ)"""
        if self.table.state() == QAbstractItemView.State.EditingState:
            return
        try:
            changed = self.repository.data_version() != self.last_data_version
        except sqlite3.Error as e:
            print(f"更新の確認に失敗しました: {e}")
            return
        if changed:
            self.reload()

    def search(self) -> None:
        patient_id = parse_patient_id(self.patient_id_input.text())
        if patient_id is None:
            QMessageBox.warning(self, "患者ID検索", "数値を入力してください。")
            return
        hits = self.model.find_rows(patient_id)
        if not hits:
            self.clear_search()
            QMessageBox.information(self, "患者ID検索", "該当する患者IDは見つかりませんでした。")
            return
        self.search_hits = hits
        self.search_position = 0
        self.show_current_hit()

    def show_next_hit(self) -> None:
        if self.search_hits:
            self.search_position = (self.search_position + 1) % len(self.search_hits)
            self.show_current_hit()

    def show_previous_hit(self) -> None:
        if self.search_hits:
            self.search_position = (self.search_position - 1) % len(self.search_hits)
            self.show_current_hit()

    def show_current_hit(self) -> None:
        row = self.search_hits[self.search_position]
        self.table.selectRow(row)
        self.table.scrollTo(self.model.index(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)
        self.search_result_label.setText(f"{self.search_position + 1}/{len(self.search_hits)}件")

    def clear_search(self) -> None:
        # 並び替えや読み直しで行の位置が変わるため、検索結果を捨てる
        self.search_hits = []
        self.search_position = 0
        self.search_result_label.setText("")

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, "エラー", message)

    def handle_conflict(self) -> None:
        QMessageBox.warning(
            self, "更新の競合",
            "他の端末でこの行が更新されたため、保存できませんでした。\n最新の内容を読み込みます。"
        )
        self.reload()
