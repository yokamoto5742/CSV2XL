import sqlite3

from PyQt6.QtCore import (
    QAbstractItemModel, QDate, QEvent, QModelIndex, QObject, QPersistentModelIndex, QPoint, Qt, QTimer, pyqtSignal
)
from PyQt6.QtGui import QKeyEvent, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractButton, QAbstractItemView, QApplication, QComboBox, QDateEdit, QMenu, QMessageBox, QStyle,
    QStyledItemDelegate, QStyleOptionHeader, QStyleOptionViewItem, QStylePainter, QTableView, QVBoxLayout, QWidget
)

from app.dialogs import DISPLAY_DATE_FORMAT, ColumnFilterDialog, DocumentDetailDialog, create_date_edit
from app.document_table_model import (
    COLUMNS, EDITABLE_COLUMNS, REQUESTED_DATE_COLUMN, STAFF_COLUMN, DocumentTableModel
)
from services.csv_export import write_csv
from services.document_repository import Document, DocumentRepository, NewDocument
from services.list_repository import STAFF, ListRepository
from services.validators import parse_patient_id

# Excel 版の列幅に近い値(ピクセル)
COLUMN_WIDTHS = [100, 80, 120, 320, 120, 100, 100, 110, 260]
AUTO_REFRESH_INTERVAL_MS = 5000
# 行の高さ = 文字の高さ + 上下の余白(ピクセル)
ROW_PADDING = 8
# 行番号の見出し(表の左上の角に表示する)と、その左右の余白(ピクセル)
CORNER_LABEL = "番号"
CORNER_LABEL_PADDING = 12
# 医師依頼日の編集中に Delete キーで消去したことを日付欄に記録するプロパティ名
CLEARED_PROPERTY = 'cleared'
# 削除の確認で表示する行数の上限
DELETE_CONFIRM_MAX_ROWS = 10


def delete_confirmation_message(documents: list[Document]) -> str:
    """削除の確認文。行が多いときは先頭の数行だけを表示する"""
    lines = [f"患者ID: {doc.patient_id}  {doc.patient_name}  文書名: {doc.document_name}"
             for doc in documents[:DELETE_CONFIRM_MAX_ROWS]]
    hidden_count = len(documents) - DELETE_CONFIRM_MAX_ROWS
    if hidden_count > 0:
        lines.append(f"ほか {hidden_count} 件")
    return f"次の {len(documents)} 行を削除しますか?\n\n" + "\n".join(lines)


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
        # 編集を始めたらすぐに選択肢を開く(表示されてから開く)
        QTimer.singleShot(0, combo.showPopup)
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

    def eventFilter(self, watched: QObject | None, event: QEvent | None) -> bool:
        # 日付欄は空にできないため、編集中に Delete キーを押したら空欄として保存する
        if (isinstance(watched, QDateEdit) and isinstance(event, QKeyEvent)
                and event.type() == QEvent.Type.KeyPress and event.key() == Qt.Key.Key_Delete):
            watched.setProperty(CLEARED_PROPERTY, True)
            self.commit_and_close(watched)
            return True
        return super().eventFilter(watched, event)

    def setEditorData(self, editor: QWidget | None, index: QModelIndex | QPersistentModelIndex) -> None:
        if not isinstance(editor, QDateEdit):
            return
        # 空欄なら今日から選び始める
        date = QDate.fromString(str(index.data(Qt.ItemDataRole.EditRole) or ''), DISPLAY_DATE_FORMAT)
        editor.setDate(date if date.isValid() else QDate.currentDate())

    def setModelData(self, editor: QWidget | None, model: QAbstractItemModel | None,
                     index: QModelIndex | QPersistentModelIndex) -> None:
        if isinstance(editor, QDateEdit) and model is not None:
            text = '' if editor.property(CLEARED_PROPERTY) else editor.date().toString(DISPLAY_DATE_FORMAT)
            model.setData(model.index(index.row(), index.column()), text, Qt.ItemDataRole.EditRole)


class CornerLabelPainter(QObject):
    """表の左上の角(行番号の上)に見出しの文字を描く。Qt の角のボタンは文字を描かないため"""

    def eventFilter(self, watched: QObject | None, event: QEvent | None) -> bool:
        if isinstance(watched, QAbstractButton) and event is not None and event.type() == QEvent.Type.Paint:
            option = QStyleOptionHeader()
            option.initFrom(watched)
            option.text = CORNER_LABEL
            option.textAlignment = Qt.AlignmentFlag.AlignCenter
            option.position = QStyleOptionHeader.SectionPosition.OnlyOneSection
            QStylePainter(watched).drawControl(QStyle.ControlElement.CE_Header, option)
            return True
        return super().eventFilter(watched, event)


class DocumentListWidget(QWidget):
    """医療文書担当一覧の表示・編集・患者ID検索"""

    # 患者ID検索の結果(例: "患者ID 100: 3件")。絞り込んでいないときは空文字
    search_status_changed = pyqtSignal(str)

    def __init__(self, connection: sqlite3.Connection, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.list_repository = ListRepository(connection)
        self.repository = DocumentRepository(connection)
        self.model = DocumentTableModel(self.repository)
        self.staff_names: list[str] = []
        self.last_data_version = -1

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.table = self.create_table()
        layout.addWidget(self.table)
        self.apply_font_size()

        # セルの確定処理の途中で表を読み直さないよう、確定が終わってから処理する
        self.model.error_occurred.connect(lambda message: QTimer.singleShot(0, lambda: self.show_error(message)))
        self.model.conflict_detected.connect(lambda: QTimer.singleShot(0, self.handle_conflict))
        # 読み直しやフィルタの変更で件数が変わるため、そのたびに表示し直す
        self.model.modelReset.connect(self.update_search_status)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_if_changed)
        self.refresh_timer.start(AUTO_REFRESH_INTERVAL_MS)

    def create_table(self) -> QTableView:
        table = QTableView()
        table.setModel(self.model)
        table.setItemDelegateForColumn(STAFF_COLUMN, StaffComboDelegate(self))
        table.setItemDelegateForColumn(REQUESTED_DATE_COLUMN, RequestedDateDelegate(self))
        # 日付欄は空にできないため、セルを選んで Delete キーで消去する(編集中は RequestedDateDelegate で処理する)
        clear_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), table)
        clear_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        clear_shortcut.activated.connect(self.clear_requested_date)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        # Ctrl+クリックで1行ずつ、Shift+クリックで範囲を追加して選べる(一括削除用)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        # ダブルクリックは詳細ダイアログを開くため、セルの直接編集は左クリック・F2 キー・文字の入力で始める
        table.setEditTriggers(QAbstractItemView.EditTrigger.EditKeyPressed
                              | QAbstractItemView.EditTrigger.AnyKeyPressed)
        table.clicked.connect(self.start_cell_edit)
        table.doubleClicked.connect(self.show_detail)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_cell_menu)
        table.setAlternatingRowColors(True)
        header = table.horizontalHeader()
        if header is not None:
            # 起動直後は既定の並び(預り日 → 診療科 → 患者ID)にするため、並び替えの列を未選択にしておく
            header.setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
            header.setStretchLastSection(True)
            # 見出しの左クリックは並び替え、右クリックはフィルタ
            header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            header.customContextMenuRequested.connect(self.show_column_filter)
        table.setSortingEnabled(True)
        for column, width in enumerate(COLUMN_WIDTHS):
            table.setColumnWidth(column, width)
        corner_button = table.findChild(QAbstractButton)
        if corner_button is not None:
            corner_button.installEventFilter(CornerLabelPainter(table))
        return table

    def column_widths(self) -> list[int]:
        return [self.table.columnWidth(column) for column in range(len(COLUMNS))]

    def set_column_widths(self, widths: list[int]) -> None:
        """保存した列幅にする。列の数が合わなければ(列を増減した後など)既定の幅のままにする"""
        if len(widths) != len(COLUMNS):
            return
        for column, width in enumerate(widths):
            self.table.setColumnWidth(column, width)

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
        """その患者IDの行だけを表示する(列のフィルタも併せて効く)。空欄なら絞り込みを外す"""
        if not text.strip():
            self.clear_search()
            return
        patient_id = parse_patient_id(text)
        if patient_id is None:
            QMessageBox.warning(self, "患者ID検索", "数値を入力してください。")
            return
        previous = self.model.patient_id_filter
        self.model.set_patient_id_filter(patient_id)
        if self.model.rowCount() == 0:
            self.model.set_patient_id_filter(previous)
            QMessageBox.information(self, "患者ID検索", "該当する患者IDは見つかりませんでした。")
            return
        self.table.selectRow(0)
        self.table.scrollToTop()

    def clear_search(self) -> None:
        self.model.set_patient_id_filter(None)

    def update_search_status(self) -> None:
        patient_id = self.model.patient_id_filter
        status = "" if patient_id is None else f"患者ID {patient_id}: {self.model.rowCount()}件"
        self.search_status_changed.emit(status)

    def clear_requested_date(self) -> None:
        index = self.table.currentIndex()
        if index.isValid() and index.column() == REQUESTED_DATE_COLUMN:
            self.model.setData(index, '')

    def apply_font_size(self) -> None:
        """フォントを変えたときに行の高さと行番号の幅(角の「番号」が収まる幅)を合わせる"""
        header = self.table.verticalHeader()
        if header is not None:
            metrics = self.table.fontMetrics()
            header.setDefaultSectionSize(metrics.height() + ROW_PADDING)
            header.setMinimumWidth(metrics.horizontalAdvance(CORNER_LABEL) + CORNER_LABEL_PADDING)

    def add_document(self, document: NewDocument) -> None:
        """手動で追加し、追加した行を選ぶ"""
        try:
            self.repository.add_document(document)
        except sqlite3.Error as e:
            self.show_error(f"行の追加に失敗しました。\n{e}")
            return
        self.reload()
        # フィルタで隠れた行も含めて、いちばん新しい行を探す
        new_id = max(doc.id for doc in self.model.all_documents)
        row = self.model.find_row_by_id(new_id)
        if row is not None:
            self.table.selectRow(row)
            self.table.scrollTo(self.model.index(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)

    def selected_documents(self) -> list[Document]:
        """選んだ行(画面の並び順)"""
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return []
        rows = sorted(index.row() for index in selection_model.selectedRows())
        return [self.model.document(row) for row in rows]

    def delete_selected(self) -> None:
        """選んだ行(複数可)を確認してから削除する"""
        documents = self.selected_documents()
        if not documents:
            QMessageBox.information(self, "行の削除", "削除する行を選んでください。")
            return
        answer = QMessageBox.question(self, "行の削除", delete_confirmation_message(documents))
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            deleted = self.repository.delete_documents(documents)
        except sqlite3.Error as e:
            self.show_error(f"行の削除に失敗しました。\n{e}")
            return
        if not deleted:
            self.handle_conflict("削除")
            return
        self.reload()

    def show_detail(self, index: QModelIndex) -> None:
        """行の全項目をダイアログで表示し、OK なら担当者名・医師依頼日・備考を保存する"""
        if not index.isValid():
            return
        document = self.model.document(index.row())
        dialog = DocumentDetailDialog(document, self.staff_names, self)
        if not dialog.exec():
            return
        edited = dialog.edited_document()
        if edited != document:
            self.model.save_document(edited)

    def start_cell_edit(self, index: QModelIndex) -> None:
        """担当者名・医師依頼日・備考のセルは左クリックでそのセルだけ編集する(Ctrl/Shift+クリックは行の選択)"""
        selecting_rows = QApplication.keyboardModifiers() & (Qt.KeyboardModifier.ControlModifier
                                                             | Qt.KeyboardModifier.ShiftModifier)
        if index.column() in EDITABLE_COLUMNS and not selecting_rows:
            self.table.edit(index)

    def show_cell_menu(self, position: QPoint) -> None:
        index = self.table.indexAt(position)
        if not index.isValid():
            return
        menu = QMenu(self)
        copy_action = menu.addAction("コピー")
        viewport = self.table.viewport()
        if viewport is not None and menu.exec(viewport.mapToGlobal(position)) is copy_action:
            self.copy_cell(index)

    def copy_cell(self, index: QModelIndex) -> None:
        clipboard = QApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(str(index.data() or ''))

    def show_column_filter(self, position: QPoint) -> None:
        """見出しを右クリックした列のフィルタの条件を選んで絞り込む"""
        header = self.table.horizontalHeader()
        if header is None:
            return
        column = header.logicalIndexAt(position)
        if column < 0:
            return
        dialog = ColumnFilterDialog(column, self.model.column_values(column), self.model.filters.get(column), self)
        dialog.move(header.mapToGlobal(position))
        if dialog.exec():
            self.model.set_filter(column, dialog.column_filter())

    def export_csv(self, path: str) -> None:
        """表示中の行を画面の並び順のまま CSV に書き出す。失敗時は OSError"""
        write_csv(path, [heading for heading, _ in COLUMNS], self.model.export_rows())

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, "エラー", message)

    def handle_conflict(self, action: str = "保存") -> None:
        QMessageBox.warning(
            self, "更新の競合",
            f"他の端末でこの行が更新されたため、{action}できませんでした。\n最新の内容を読み込みます。"
        )
        self.reload()
