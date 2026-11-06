import dataclasses
import sqlite3
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtCore import QDate, QEvent, QModelIndex, QPoint, Qt
from PyQt6.QtGui import QKeyEvent
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDateEdit, QMessageBox, QStyleOptionViewItem
)

from app.document_table_model import NOTE_COLUMN, REQUESTED_DATE_COLUMN, STAFF_COLUMN, ColumnFilter
from app.document_window import (
    COLUMN_WIDTHS, CORNER_LABEL_PADDING, ROW_PADDING, DocumentListWidget, RequestedDateDelegate, StaffComboDelegate
)
from services.database import connect
from services.document_repository import DocumentRepository
from services.list_repository import STAFF, ListRepository
from tests.conftest import make_new_document


@pytest.fixture
def app():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


@pytest.fixture
def window(app, connection: sqlite3.Connection) -> DocumentListWidget:
    DocumentRepository(connection).add_new_documents([
        make_new_document(100, '2025-01-10'),
        make_new_document(200, '2025-01-11'),
        make_new_document(100, '2025-01-12'),
    ])
    ListRepository(connection).set_items(STAFF, ['山本', '植田'])
    list_window = DocumentListWidget(connection)
    list_window.refresh_timer.stop()
    list_window.load_initial()
    return list_window


def selected_row(window: DocumentListWidget) -> int:
    selection_model = window.table.selectionModel()
    assert selection_model is not None
    return selection_model.selectedRows()[0].row()


@pytest.fixture
def statuses(window: DocumentListWidget) -> list[str]:
    """search_status_changed で通知された文字列"""
    received: list[str] = []
    window.search_status_changed.connect(received.append)
    return received


def visible_patient_ids(window: DocumentListWidget) -> list[int]:
    return [window.model.document(row).patient_id for row in range(window.model.rowCount())]


class TestSearch:
    def test_search_shows_only_matching_rows(self, window: DocumentListWidget, statuses: list[str]):
        window.search('100')

        assert visible_patient_ids(window) == [100, 100]
        assert selected_row(window) == 0
        assert statuses[-1] == "患者ID 100: 2件"

    def test_clear_search(self, window: DocumentListWidget, statuses: list[str]):
        window.search('100')

        window.clear_search()

        assert visible_patient_ids(window) == [100, 200, 100]
        assert statuses[-1] == ""

    def test_empty_text_clears_search(self, window: DocumentListWidget, statuses: list[str]):
        window.search('100')

        window.search(' ')

        assert window.model.rowCount() == 3
        assert statuses[-1] == ""

    @patch('app.document_window.QMessageBox.information')
    def test_not_found_keeps_previous_search(self, mock_info, window: DocumentListWidget, statuses: list[str]):
        window.search('200')

        window.search('999')

        assert mock_info.call_args[0][2] == "該当する患者IDは見つかりませんでした。"
        assert visible_patient_ids(window) == [200]
        assert statuses[-1] == "患者ID 200: 1件"

    @patch('app.document_window.QMessageBox.warning')
    def test_not_a_number(self, mock_warning, window: DocumentListWidget):
        window.search('abc')

        assert mock_warning.call_args[0][2] == "数値を入力してください。"

    def test_search_kept_after_sort_and_reload(self, window: DocumentListWidget, statuses: list[str]):
        window.search('100')

        window.model.sort(1)
        window.reload()

        assert visible_patient_ids(window) == [100, 100]
        assert statuses[-1] == "患者ID 100: 2件"


class TestRefresh:
    def test_refresh_when_other_connection_changes(self, window: DocumentListWidget, tmp_path: Path):
        """他の端末(別の接続)で追加された行を自動更新で読み込む"""
        other = connect(str(tmp_path / 'test.db'))
        try:
            DocumentRepository(other).add_new_documents([make_new_document(300, '2025-01-13')])
        finally:
            other.close()

        window.refresh_if_changed()

        assert window.model.rowCount() == 4

    def test_no_reload_without_changes(self, window: DocumentListWidget):
        with patch.object(window.model, 'load') as mock_load:
            window.refresh_if_changed()
        mock_load.assert_not_called()

    @patch('app.document_window.QMessageBox.warning')
    def test_conflict_reloads(self, mock_warning, window: DocumentListWidget, tmp_path: Path):
        """他の端末が先に更新した行を保存しようとしたら、警告して読み直す"""
        other = connect(str(tmp_path / 'test.db'))
        try:
            other_repository = DocumentRepository(other)
            target = next(doc for doc in other_repository.fetch_all() if doc.id == window.model.document(0).id)
            other_repository.update_editable_fields(dataclasses.replace(target, note='他端末'))
        finally:
            other.close()

        assert not window.model.setData(window.model.index(0, 8), '自端末')
        QApplication.processEvents()

        mock_warning.assert_called_once()
        assert window.model.document(0).note == '他端末'


class TestStaffComboDelegate:
    def test_editor_lists_staff_and_keeps_current_value(self, window: DocumentListWidget):
        """リストに無い担当者が入力済みでも、その値を選択肢に残す"""
        window.model.setData(window.model.index(0, STAFF_COLUMN), '退職者')
        delegate = StaffComboDelegate(window)
        combo = QComboBox()

        delegate.setEditorData(combo, window.model.index(0, STAFF_COLUMN))

        assert [combo.itemText(i) for i in range(combo.count())] == ['', '山本', '植田', '退職者']
        assert combo.currentText() == '退職者'

    def test_set_model_data(self, window: DocumentListWidget):
        delegate = StaffComboDelegate(window)
        combo = QComboBox()
        combo.addItems(['', '山本'])
        combo.setCurrentText('山本')
        model = MagicMock()
        index = window.model.index(0, STAFF_COLUMN)

        delegate.setModelData(combo, model, index)

        model.setData.assert_called_once()
        assert model.setData.call_args[0][1] == '山本'


class TestRequestedDateDelegate:
    def index(self, window: DocumentListWidget):
        return window.model.index(0, REQUESTED_DATE_COLUMN)

    def test_editor_starts_today_when_empty(self, window: DocumentListWidget):
        delegate = RequestedDateDelegate(window)
        date_edit = QDateEdit()

        delegate.setEditorData(date_edit, self.index(window))

        assert date_edit.date() == QDate.currentDate()

    def test_editor_shows_current_value(self, window: DocumentListWidget):
        window.model.setData(self.index(window), '2025/01/15')
        delegate = RequestedDateDelegate(window)
        date_edit = QDateEdit()

        delegate.setEditorData(date_edit, self.index(window))

        assert date_edit.date() == QDate(2025, 1, 15)

    def test_calendar_click_saves_date(self, window: DocumentListWidget):
        """カレンダーで日付を選ぶとすぐ保存する"""
        delegate = RequestedDateDelegate(window)
        editor = delegate.createEditor(window.table, QStyleOptionViewItem(), self.index(window))
        assert isinstance(editor, QDateEdit)
        delegate.commitData.connect(lambda _: delegate.setModelData(editor, window.model, self.index(window)))

        delegate.select_date(editor, QDate(2025, 2, 3))

        assert window.model.document(0).requested_date == '2025-02-03'

    def test_delete_key_clears_date(self, window: DocumentListWidget):
        window.model.setData(self.index(window), '2025/01/15')
        window.table.setCurrentIndex(self.index(window))

        window.clear_requested_date()

        assert window.model.document(0).requested_date is None

    def test_delete_key_while_editing_clears_date(self, window: DocumentListWidget):
        """左クリックで編集中の日付欄でも Delete キーで消去できる"""
        window.model.setData(self.index(window), '2025/01/15')
        window.table.setCurrentIndex(self.index(window))
        window.table.edit(self.index(window))
        editor = window.table.findChild(QDateEdit)
        assert isinstance(editor, QDateEdit)

        QApplication.sendEvent(editor, QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_Delete,
                                                 Qt.KeyboardModifier.NoModifier))

        assert window.model.document(0).requested_date is None

    def test_delete_key_ignores_other_columns(self, window: DocumentListWidget):
        window.model.setData(window.model.index(0, STAFF_COLUMN), '山本')
        window.table.setCurrentIndex(window.model.index(0, STAFF_COLUMN))

        window.clear_requested_date()

        assert window.model.document(0).staff_name == '山本'


class TestAddAndDelete:
    def test_add_document_selects_new_row(self, window: DocumentListWidget):
        window.add_document(make_new_document(999, '2025-01-01'))

        assert window.model.rowCount() == 4
        assert window.model.document(selected_row(window)).patient_id == 999

    @patch('app.document_window.QMessageBox.question', return_value=QMessageBox.StandardButton.Yes)
    def test_delete_selected(self, mock_question, window: DocumentListWidget):
        window.table.selectRow(1)
        target = window.model.document(1)

        window.delete_selected()

        assert window.model.rowCount() == 2
        assert window.model.find_row_by_id(target.id) is None

    @patch('app.document_window.QMessageBox.question', return_value=QMessageBox.StandardButton.No)
    def test_delete_cancelled(self, mock_question, window: DocumentListWidget):
        window.table.selectRow(1)

        window.delete_selected()

        assert window.model.rowCount() == 3

    @patch('app.document_window.QMessageBox.information')
    def test_delete_without_selection(self, mock_info, window: DocumentListWidget):
        window.table.setCurrentIndex(QModelIndex())

        window.delete_selected()

        mock_info.assert_called_once()
        assert window.model.rowCount() == 3

    @patch('app.document_window.QMessageBox.warning')
    @patch('app.document_window.QMessageBox.question', return_value=QMessageBox.StandardButton.Yes)
    def test_delete_conflict_reloads(self, mock_question, mock_warning, window: DocumentListWidget,
                                     tmp_path: Path):
        """他の端末が先に更新した行は削除せず、警告して読み直す"""
        window.table.selectRow(0)
        other = connect(str(tmp_path / 'test.db'))
        try:
            other_repository = DocumentRepository(other)
            target = next(doc for doc in other_repository.fetch_all() if doc.id == window.model.document(0).id)
            other_repository.update_editable_fields(dataclasses.replace(target, note='他端末'))
        finally:
            other.close()

        window.delete_selected()

        assert "削除できませんでした" in mock_warning.call_args[0][2]
        assert window.model.rowCount() == 3
        assert window.model.document(0).note == '他端末'


class TestExportCsv:
    def test_export_in_display_order(self, window: DocumentListWidget, tmp_path: Path):
        path = tmp_path / 'export.csv'

        window.export_csv(str(path))

        lines = path.read_text(encoding='utf-8-sig').splitlines()
        assert lines[0] == '預り日,患者ID,患者氏名,文書名,診療科,医師名,担当者名,医師依頼日,備考'
        assert [line.split(',')[0] for line in lines[1:]] == ['2025/01/10', '2025/01/11', '2025/01/12']


class TestDetailDialog:
    @patch('app.document_window.DocumentDetailDialog')
    def test_double_click_saves_edits(self, mock_dialog, window: DocumentListWidget):
        """ダブルクリックで詳細ダイアログを開き、OK なら編集した値を保存する"""
        target = window.model.document(1)
        mock_dialog.return_value.exec.return_value = 1
        mock_dialog.return_value.edited_document.return_value = dataclasses.replace(target, note='確認済み')

        window.table.doubleClicked.emit(window.model.index(1, 3))

        assert mock_dialog.call_args[0][:2] == (target, ['山本', '植田'])
        assert window.model.document(1).note == '確認済み'

    @patch('app.document_window.DocumentDetailDialog')
    def test_cancel_does_not_save(self, mock_dialog, window: DocumentListWidget):
        mock_dialog.return_value.exec.return_value = 0

        with patch.object(window.model, 'save_document') as mock_save:
            window.show_detail(window.model.index(0, 0))

        mock_save.assert_not_called()

    def test_double_click_does_not_start_inline_edit(self, window: DocumentListWidget):
        triggers = window.table.editTriggers()
        assert not triggers & QAbstractItemView.EditTrigger.DoubleClicked
        assert triggers & QAbstractItemView.EditTrigger.EditKeyPressed


class TestClickEdit:
    @pytest.mark.parametrize("column", [STAFF_COLUMN, REQUESTED_DATE_COLUMN, NOTE_COLUMN])
    def test_click_editable_column_starts_edit(self, window: DocumentListWidget, column: int):
        """担当者名・医師依頼日・備考のセルは左クリックで編集を始める"""
        window.table.clicked.emit(window.model.index(0, column))

        assert window.table.state() == QAbstractItemView.State.EditingState

    def test_click_other_column_does_not_start_edit(self, window: DocumentListWidget):
        window.table.clicked.emit(window.model.index(0, 3))

        assert window.table.state() != QAbstractItemView.State.EditingState


def test_copy_cell(window: DocumentListWidget):
    window.copy_cell(window.model.index(0, 2))

    clipboard = QApplication.clipboard()
    assert clipboard is not None
    assert clipboard.text() == '山田　太郎'


class TestColumnFilter:
    @patch('app.document_window.ColumnFilterDialog')
    def test_header_right_click_filters(self, mock_dialog, window: DocumentListWidget):
        mock_dialog.return_value.exec.return_value = 1
        mock_dialog.return_value.column_filter.return_value = ColumnFilter(values=frozenset({'100'}))
        header = window.table.horizontalHeader()
        assert header is not None
        window.resize(1200, 400)
        position = QPoint(header.sectionViewportPosition(1) + 5, 5)

        window.show_column_filter(position)

        assert mock_dialog.call_args[0][:3] == (1, ['100', '200'], None)
        assert window.model.rowCount() == 2
        assert window.model.filters == {1: ColumnFilter(values=frozenset({'100'}))}

    @patch('app.document_window.ColumnFilterDialog')
    def test_cancel_keeps_filter(self, mock_dialog, window: DocumentListWidget):
        mock_dialog.return_value.exec.return_value = 0

        window.show_column_filter(QPoint(5, 5))

        assert window.model.rowCount() == 3

    def test_search_only_visible_rows(self, window: DocumentListWidget, statuses: list[str]):
        window.model.set_filter(0, ColumnFilter(values=frozenset({'2025/01/12'})))

        window.search('100')

        assert statuses[-1] == "患者ID 100: 1件"

    def test_export_only_visible_rows(self, window: DocumentListWidget, tmp_path: Path):
        window.model.set_filter(1, ColumnFilter(values=frozenset({'200'})))
        path = tmp_path / 'export.csv'

        window.export_csv(str(path))

        assert len(path.read_text(encoding='utf-8-sig').splitlines()) == 2

    def test_add_document_while_hidden_by_filter(self, window: DocumentListWidget):
        """追加した行がフィルタで隠れても、エラーにならない"""
        window.model.set_filter(1, ColumnFilter(values=frozenset({'100'})))

        window.add_document(make_new_document(999, '2025-01-01'))

        assert len(window.model.all_documents) == 4
        assert window.model.rowCount() == 2


def test_apply_font_size_sets_row_height(window: DocumentListWidget):
    header = window.table.verticalHeader()
    assert header is not None

    window.apply_font_size()

    assert header.defaultSectionSize() == window.table.fontMetrics().height() + ROW_PADDING
    assert header.minimumWidth() == window.table.fontMetrics().horizontalAdvance("番号") + CORNER_LABEL_PADDING


class TestColumnWidths:
    def test_set_and_get(self, window: DocumentListWidget):
        widths = [90, 70, 110, 300, 100, 90, 90, 100, 200]

        window.set_column_widths(widths)

        assert window.column_widths()[:-1] == widths[:-1]

    def test_wrong_count_keeps_default(self, window: DocumentListWidget):
        window.set_column_widths([50, 50])

        assert window.column_widths()[:-1] == COLUMN_WIDTHS[:-1]
