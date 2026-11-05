import sqlite3

from PyQt6.QtCore import QAbstractItemModel, QDate, QModelIndex, QPersistentModelIndex, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView, QComboBox, QDateEdit, QMessageBox, QStyledItemDelegate, QStyleOptionViewItem,
    QTableView, QVBoxLayout, QWidget
)

from app.document_table_model import COLUMNS, REQUESTED_DATE_COLUMN, STAFF_COLUMN, DocumentTableModel
from services.csv_export import write_csv
from services.document_repository import DocumentRepository, NewDocument
from services.list_repository import STAFF, ListRepository
from services.validators import parse_patient_id

# Excel 版の列幅に近い値(ピクセル)
COLUMN_WIDTHS = [100, 80, 120, 320, 120, 100, 100, 110, 260]
AUTO_REFRESH_INTERVAL_MS = 5000
DISPLAY_DATE_FORMAT = 'yyyy/MM/dd'
# 行の高さ = 文字の高さ + 上下の余白(ピクセル)
ROW_PADDING = 8


def create_date_edit(parent: QWidget | None = None) -> QDateEdit:
    """▼でカレンダーを開ける日付欄"""
    date_edit = QDateEdit(parent)
    date_edit.setCalendarPopup(True)
    date_edit.setDisplayFormat(DISPLAY_DATE_FORMAT)
    return date_edit


class CommitOnSelectDelegate(QStyledItemDelegate):
    """選んだらすぐ確定する編集欄の共通処理"""

    def commit_and_close(self, editor: QWidget) -> None:
        self.commitData.emit(editor)
        self.closeEditor.emit(editor)


class StaffComboDelegate(CommitOnSelectDelegate):
    """担当者名をドロップダウンで選ぶ"""

    def __init__(self, window: 'DocumentListWidget') -> None:
        super().__init__(window)
        self.window = window

    def createEditor(self, parent: QWidget | None, option: QStyleOptionViewItem,
                     index: QModelIndex | QPersistentModelIndex) -> QWidget:
        combo = QComboBox(parent)
        combo.activated.connect(lambda: self.commit_and_close(combo))
        return combo

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


class RequestedDateDelegate(CommitOnSelectDelegate):
    """医師依頼日を日付欄とカレンダーで入力する"""

    def createEditor(self, parent: QWidget | None, option: QStyleOptionViewItem,
                     index: QModelIndex | QPersistentModelIndex) -> QWidget:
        date_edit = create_date_edit(parent)
        calendar = date_edit.calendarWidget()
        if calendar is not None:
            calendar.clicked.connect(lambda date: self.select_date(date_edit, date))
        return date_edit

    def select_date(self, date_edit: QDateEdit, date: QDate) -> None:
        date_edit.setDate(date)
        self.commit_and_close(date_edit)

    def setEditorData(self, editor: QWidget | None, index: QModelIndex | QPersistentModelIndex) -> None:
        if not isinstance(editor, QDateEdit):
            return
        # 空欄なら今日から選び始める
        date = QDate.fromString(str(index.data(Qt.ItemDataRole.EditRole) or ''), DISPLAY_DATE_FORMAT)
        editor.setDate(date if date.isValid() else QDate.currentDate())

    def setModelData(self, editor: QWidget | None, model: QAbstractItemModel | None,
                     index: QModelIndex | QPersistentModelIndex) -> None:
        if isinstance(editor, QDateEdit) and model is not None:
            model.setData(model.index(index.row(), index.column()),
                          editor.date().toString(DISPLAY_DATE_FORMAT), Qt.ItemDataRole.EditRole)


class DocumentListWidget(QWidget):
    """医療文書担当一覧の表示・編集・患者ID検索"""

    # 検索結果の位置(例: "1/3件")。検索結果が無いときは空文字
    search_status_changed = pyqtSignal(str)

    def __init__(self, connection: sqlite3.Connection, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.list_repository = ListRepository(connection)
        self.repository = DocumentRepository(connection)
        self.model = DocumentTableModel(self.repository)
        self.staff_names: list[str] = []
        self.search_hits: list[int] = []
        self.search_position = 0
        self.last_data_version = -1

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.table = self.create_table()
        layout.addWidget(self.table)
        self.apply_font_size()

        # セルの確定処理の途中で表を読み直さないよう、確定が終わってから処理する
        self.model.error_occurred.connect(lambda message: QTimer.singleShot(0, lambda: self.show_error(message)))
        self.model.conflict_detected.connect(lambda: QTimer.singleShot(0, self.handle_conflict))
        self.model.modelReset.connect(self.clear_search)
        self.model.layoutChanged.connect(self.clear_search)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_if_changed)
        self.refresh_timer.start(AUTO_REFRESH_INTERVAL_MS)

    def create_table(self) -> QTableView:
        table = QTableView()
        table.setModel(self.model)
        table.setItemDelegateForColumn(STAFF_COLUMN, StaffComboDelegate(self))
        table.setItemDelegateForColumn(REQUESTED_DATE_COLUMN, RequestedDateDelegate(self))
        # 日付欄は空にできないため、セルを選んで Delete キーで消去する(編集中は日付欄側で処理される)
        clear_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), table)
        clear_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        clear_shortcut.activated.connect(self.clear_requested_date)
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

    def search(self, text: str) -> None:
        patient_id = parse_patient_id(text)
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
        self.search_status_changed.emit(f"{self.search_position + 1}/{len(self.search_hits)}件")

    def clear_search(self) -> None:
        # 並び替えや読み直しで行の位置が変わるため、検索結果を捨てる
        self.search_hits = []
        self.search_position = 0
        self.search_status_changed.emit("")

    def clear_requested_date(self) -> None:
        index = self.table.currentIndex()
        if index.isValid() and index.column() == REQUESTED_DATE_COLUMN:
            self.model.setData(index, '')

    def apply_font_size(self) -> None:
        """フォントを変えたときに行の高さを合わせる"""
        header = self.table.verticalHeader()
        if header is not None:
            header.setDefaultSectionSize(self.table.fontMetrics().height() + ROW_PADDING)

    def add_document(self, document: NewDocument) -> None:
        """手動で追加し、追加した行を選ぶ"""
        try:
            self.repository.add_document(document)
        except sqlite3.Error as e:
            self.show_error(f"行の追加に失敗しました。\n{e}")
            return
        self.reload()
        new_id = max(doc.id for doc in self.model.documents)
        row = self.model.find_row_by_id(new_id)
        if row is not None:
            self.table.selectRow(row)
            self.table.scrollTo(self.model.index(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)

    def delete_selected(self) -> None:
        """選んだ行を確認してから削除する"""
        index = self.table.currentIndex()
        if not index.isValid():
            QMessageBox.information(self, "行の削除", "削除する行を選んでください。")
            return
        document = self.model.document(index.row())
        answer = QMessageBox.question(
            self, "行の削除",
            f"次の行を削除しますか?\n\n患者ID: {document.patient_id}  {document.patient_name}\n文書名: {document.document_name}"
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            deleted = self.repository.delete_document(document)
        except sqlite3.Error as e:
            self.show_error(f"行の削除に失敗しました。\n{e}")
            return
        if not deleted:
            self.handle_conflict("削除")
            return
        self.reload()

    def export_csv(self, path: str) -> None:
        """画面の並び順のまま CSV に書き出す。失敗時は OSError"""
        write_csv(path, [heading for heading, _ in COLUMNS], self.model.export_rows())

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, "エラー", message)

    def handle_conflict(self, action: str = "保存") -> None:
        QMessageBox.warning(
            self, "更新の競合",
            f"他の端末でこの行が更新されたため、{action}できませんでした。\n最新の内容を読み込みます。"
        )
        self.reload()
