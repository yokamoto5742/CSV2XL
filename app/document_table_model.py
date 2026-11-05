import dataclasses
import sqlite3
from typing import Any

from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QPersistentModelIndex, Qt, pyqtSignal

from services.document_repository import Document, DocumentRepository
from services.validators import parse_date

# (見出し, Document の属性名)
COLUMNS = [
    ("預り日", "received_date"),
    ("患者ID", "patient_id"),
    ("患者名", "patient_name"),
    ("文書名", "document_name"),
    ("診療科", "department"),
    ("医師名", "doctor_name"),
    ("担当者名", "staff_name"),
    ("医師依頼日", "requested_date"),
    ("備考", "note"),
]
STAFF_COLUMN = 6
REQUESTED_DATE_COLUMN = 7
NOTE_COLUMN = 8
EDITABLE_COLUMNS = {STAFF_COLUMN, REQUESTED_DATE_COLUMN, NOTE_COLUMN}
DATE_COLUMNS = {0, REQUESTED_DATE_COLUMN}
LEFT_ALIGNED_COLUMNS = {2, 3, NOTE_COLUMN}
DATE_ERROR_MESSAGE = "医師依頼日は YYYY/MM/DD の形式で入力してください。\n(例: 2025/01/05)"

ModelIndex = QModelIndex | QPersistentModelIndex


def default_sort_key(document: Document) -> tuple[str, str, int, int]:
    """既定の並び: 預り日 → 診療科 → 患者ID(同じなら登録順)"""
    return document.received_date, document.department, document.patient_id, document.id


def sort_documents(documents: list[Document], column: int, descending: bool) -> list[Document]:
    """column が負なら既定の並び、それ以外はその列で並べる(空欄は常に末尾)"""
    result = sorted(documents, key=default_sort_key)
    if column < 0:
        return result
    field = COLUMNS[column][1]
    filled = [doc for doc in result if getattr(doc, field) not in (None, '')]
    empty = [doc for doc in result if getattr(doc, field) in (None, '')]
    return sorted(filled, key=lambda doc: getattr(doc, field), reverse=descending) + empty


def format_cell(document: Document, column: int) -> str:
    value = getattr(document, COLUMNS[column][1])
    if value is None:
        return ''
    if column in DATE_COLUMNS:
        return str(value).replace('-', '/')
    return str(value)


def matches_filters(document: Document, filters: dict[int, set[str]]) -> bool:
    """すべての列で、表示する値に含まれているか"""
    return all(format_cell(document, column) in values for column, values in filters.items())


class DocumentTableModel(QAbstractTableModel):
    """担当一覧の表。編集した値はすぐ DB に保存する"""

    error_occurred = pyqtSignal(str)
    conflict_detected = pyqtSignal()

    def __init__(self, repository: DocumentRepository) -> None:
        super().__init__()
        self.repository = repository
        # all_documents は全行、documents はフィルタで残った表示中の行(どちらも現在の並び順)
        self.all_documents: list[Document] = []
        self.documents: list[Document] = []
        # 列ごとの表示する値(表示形式の文字列)。フィルタの無い列は含めない
        self.filters: dict[int, set[str]] = {}
        self.sort_column = -1
        self.sort_descending = False

    def load(self) -> None:
        """DB から読み直し、現在の並び順とフィルタで表示する"""
        documents = self.repository.fetch_all()
        self.beginResetModel()
        self.all_documents = sort_documents(documents, self.sort_column, self.sort_descending)
        self.documents = self.filtered_documents()
        self.endResetModel()

    def filtered_documents(self) -> list[Document]:
        return [doc for doc in self.all_documents if matches_filters(doc, self.filters)]

    def set_filter(self, column: int, values: set[str] | None) -> None:
        """列のフィルタを変える。values が None ならその列のフィルタを外す"""
        if values is None:
            self.filters.pop(column, None)
        else:
            self.filters[column] = values
        self.beginResetModel()
        self.documents = self.filtered_documents()
        self.endResetModel()
        self.headerDataChanged.emit(Qt.Orientation.Horizontal, column, column)

    def column_values(self, column: int) -> list[str]:
        """フィルタで選べる値。ほかの列のフィルタで残る行の値を、並べ替えと同じく空欄を末尾にして返す"""
        other_filters = {col: values for col, values in self.filters.items() if col != column}
        field = COLUMNS[column][1]
        raw_values = {format_cell(doc, column): getattr(doc, field)
                      for doc in self.all_documents if matches_filters(doc, other_filters)}
        return sorted(raw_values, key=lambda text: (raw_values[text] in (None, ''), raw_values[text] or ''))

    def document(self, row: int) -> Document:
        return self.documents[row]

    def find_rows(self, patient_id: int) -> list[int]:
        return [row for row, doc in enumerate(self.documents) if doc.patient_id == patient_id]

    def find_row_by_id(self, document_id: int) -> int | None:
        return next((row for row, doc in enumerate(self.documents) if doc.id == document_id), None)

    def export_rows(self) -> list[list[str]]:
        """画面の並び順・表示形式のままの全行"""
        return [[format_cell(doc, column) for column in range(len(COLUMNS))] for doc in self.documents]

    def rowCount(self, parent: ModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.documents)

    def columnCount(self, parent: ModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(COLUMNS)

    def headerData(self, section: int, orientation: Qt.Orientation,
                   role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            # フィルタをかけている列は見出しに ▼ を付ける
            return f"{COLUMNS[section][0]} ▼" if section in self.filters else COLUMNS[section][0]
        return super().headerData(section, orientation, role)

    def data(self, index: ModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            return None
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            return format_cell(self.documents[index.row()], index.column())
        if role == Qt.ItemDataRole.TextAlignmentRole:
            horizontal = (Qt.AlignmentFlag.AlignLeft if index.column() in LEFT_ALIGNED_COLUMNS
                          else Qt.AlignmentFlag.AlignHCenter)
            return horizontal | Qt.AlignmentFlag.AlignVCenter
        return None

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        flags = super().flags(index)
        if index.column() in EDITABLE_COLUMNS:
            flags |= Qt.ItemFlag.ItemIsEditable
        return flags

    def setData(self, index: ModelIndex, value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool:
        if role != Qt.ItemDataRole.EditRole or index.column() not in EDITABLE_COLUMNS:
            return False
        current = self.documents[index.row()]
        try:
            edited = self.apply_edit(current, index.column(), str(value or ''))
        except ValueError:
            self.error_occurred.emit(DATE_ERROR_MESSAGE)
            return False
        if edited == current:
            return False
        return self.save_document(edited)

    def save_document(self, edited: Document) -> bool:
        """担当者名・医師依頼日・備考を DB に保存し、表の行を置き換える

        Qt から呼ばれる関数なので、例外は外に出さずにシグナルで知らせる
        """
        try:
            saved = self.repository.update_editable_fields(edited)
        except sqlite3.Error as e:
            self.error_occurred.emit(f"保存に失敗しました。\n{e}")
            return False
        if saved is None:
            self.conflict_detected.emit()
            return False

        # 保存した値がフィルタに合わなくなっても、読み直すまでは行を残す
        self.all_documents = [saved if doc.id == saved.id else doc for doc in self.all_documents]
        row = self.find_row_by_id(saved.id)
        if row is not None:
            self.documents[row] = saved
            self.dataChanged.emit(self.index(row, 0), self.index(row, len(COLUMNS) - 1))
        return True

    @staticmethod
    def apply_edit(document: Document, column: int, text: str) -> Document:
        """入力値を反映した Document を返す。日付が不正なら ValueError"""
        text = text.strip()
        if column == STAFF_COLUMN:
            return dataclasses.replace(document, staff_name=text or None)
        if column == REQUESTED_DATE_COLUMN:
            return dataclasses.replace(document, requested_date=parse_date(text))
        return dataclasses.replace(document, note=text or None)

    def sort(self, column: int, order: Qt.SortOrder = Qt.SortOrder.AscendingOrder) -> None:
        self.sort_column = column
        self.sort_descending = order == Qt.SortOrder.DescendingOrder
        self.layoutAboutToBeChanged.emit()
        self.all_documents = sort_documents(self.all_documents, self.sort_column, self.sort_descending)
        self.documents = self.filtered_documents()
        self.layoutChanged.emit()
