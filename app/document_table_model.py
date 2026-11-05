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
    ("患者氏名", "patient_name"),
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
# 文字を入力して部分一致で絞り込める列(患者氏名・文書名・医師名・担当者名・備考)
TEXT_FILTER_COLUMNS = {2, 3, 5, STAFF_COLUMN, NOTE_COLUMN}
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


@dataclasses.dataclass(frozen=True)
class ColumnFilter:
    """列のフィルタの条件。指定した条件をすべて満たす行を表示する"""
    # 表示する値(表示形式の文字列)。None ならチェックボックスでは絞らない
    values: frozenset[str] | None = None
    # 部分一致で絞り込む文字。空なら絞らない
    text: str = ''
    # 日付の範囲(YYYY-MM-DD)。None ならその側は絞らない
    date_from: str | None = None
    date_to: str | None = None

    def is_empty(self) -> bool:
        return self == ColumnFilter()

    def matches(self, document: Document, column: int) -> bool:
        cell = format_cell(document, column)
        if self.values is not None and cell not in self.values:
            return False
        if self.text and self.text not in cell:
            return False
        if self.date_from is None and self.date_to is None:
            return True
        # 日付の範囲を指定したときは、空欄の行は表示しない
        date = getattr(document, COLUMNS[column][1])
        if not date:
            return False
        return (self.date_from is None or date >= self.date_from) and (self.date_to is None or date <= self.date_to)


def matches_filters(document: Document, filters: dict[int, ColumnFilter]) -> bool:
    """すべての列のフィルタの条件を満たすか"""
    return all(column_filter.matches(document, column) for column, column_filter in filters.items())


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
        # 列ごとのフィルタの条件。フィルタの無い列は含めない
        self.filters: dict[int, ColumnFilter] = {}
        # 患者ID検索で絞り込んでいる患者ID
        self.patient_id_filter: int | None = None
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
        return [doc for doc in self.all_documents if self.is_visible(doc, self.filters)]

    def is_visible(self, document: Document, filters: dict[int, ColumnFilter]) -> bool:
        if self.patient_id_filter is not None and document.patient_id != self.patient_id_filter:
            return False
        return matches_filters(document, filters)

    def apply_filters(self) -> None:
        self.beginResetModel()
        self.documents = self.filtered_documents()
        self.endResetModel()

    def set_filter(self, column: int, column_filter: ColumnFilter | None) -> None:
        """列のフィルタを変える。None か条件が空ならその列のフィルタを外す"""
        if column_filter is None or column_filter.is_empty():
            self.filters.pop(column, None)
        else:
            self.filters[column] = column_filter
        self.apply_filters()
        self.headerDataChanged.emit(Qt.Orientation.Horizontal, column, column)

    def set_patient_id_filter(self, patient_id: int | None) -> None:
        """患者ID検索の絞り込みを変える。None なら外す"""
        self.patient_id_filter = patient_id
        self.apply_filters()

    def column_values(self, column: int) -> list[str]:
        """フィルタで選べる値。ほかの列のフィルタで残る行の値を、並べ替えと同じく空欄を末尾にして返す"""
        other_filters = {col: condition for col, condition in self.filters.items() if col != column}
        field = COLUMNS[column][1]
        raw_values = {format_cell(doc, column): getattr(doc, field)
                      for doc in self.all_documents if self.is_visible(doc, other_filters)}
        return sorted(raw_values, key=lambda text: (raw_values[text] in (None, ''), raw_values[text] or ''))

    def document(self, row: int) -> Document:
        return self.documents[row]

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
