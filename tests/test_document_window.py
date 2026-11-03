import dataclasses
import sqlite3
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtWidgets import QApplication, QComboBox

from app.document_table_model import STAFF_COLUMN
from app.document_window import DocumentListWindow, StaffComboDelegate
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
def window(app, connection: sqlite3.Connection) -> DocumentListWindow:
    DocumentRepository(connection).add_new_documents([
        make_new_document(100, '2025-01-10'),
        make_new_document(200, '2025-01-11'),
        make_new_document(100, '2025-01-12'),
    ])
    ListRepository(connection).set_items(STAFF, ['山本', '植田'])
    list_window = DocumentListWindow(connection)
    list_window.refresh_timer.stop()
    list_window.load_initial()
    return list_window


def selected_row(window: DocumentListWindow) -> int:
    selection_model = window.table.selectionModel()
    assert selection_model is not None
    return selection_model.selectedRows()[0].row()


class TestSearch:
    def test_search_and_cycle(self, window: DocumentListWindow):
        """見つかった行を選び、次へ・前へで順に移る(端まで行ったら先頭に戻る)"""
        window.patient_id_input.setText('100')
        window.search()
        assert selected_row(window) == 0
        assert window.search_result_label.text() == "1/2件"

        window.show_next_hit()
        assert selected_row(window) == 2
        assert window.search_result_label.text() == "2/2件"

        window.show_next_hit()
        assert window.search_result_label.text() == "1/2件"

        window.show_previous_hit()
        assert window.search_result_label.text() == "2/2件"

    @patch('app.document_window.QMessageBox.information')
    def test_not_found(self, mock_info, window: DocumentListWindow):
        window.patient_id_input.setText('999')
        window.search()

        assert mock_info.call_args[0][2] == "該当する患者IDは見つかりませんでした。"
        assert window.search_result_label.text() == ""

    @patch('app.document_window.QMessageBox.warning')
    def test_not_a_number(self, mock_warning, window: DocumentListWindow):
        window.patient_id_input.setText('abc')
        window.search()

        assert mock_warning.call_args[0][2] == "数値を入力してください。"

    def test_sort_clears_search(self, window: DocumentListWindow):
        window.patient_id_input.setText('100')
        window.search()

        window.model.sort(1)

        assert window.search_hits == []
        assert window.search_result_label.text() == ""


class TestRefresh:
    def test_refresh_when_other_connection_changes(self, window: DocumentListWindow, tmp_path: Path):
        """他の端末(別の接続)で追加された行を自動更新で読み込む"""
        other = connect(str(tmp_path / 'test.db'))
        try:
            DocumentRepository(other).add_new_documents([make_new_document(300, '2025-01-13')])
        finally:
            other.close()

        window.refresh_if_changed()

        assert window.model.rowCount() == 4

    def test_no_reload_without_changes(self, window: DocumentListWindow):
        with patch.object(window.model, 'load') as mock_load:
            window.refresh_if_changed()
        mock_load.assert_not_called()

    @patch('app.document_window.QMessageBox.warning')
    def test_conflict_reloads(self, mock_warning, window: DocumentListWindow, tmp_path: Path):
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
    def test_editor_lists_staff_and_keeps_current_value(self, window: DocumentListWindow):
        """リストに無い担当者が入力済みでも、その値を選択肢に残す"""
        window.model.setData(window.model.index(0, STAFF_COLUMN), '退職者')
        delegate = StaffComboDelegate(window)
        combo = QComboBox()

        delegate.setEditorData(combo, window.model.index(0, STAFF_COLUMN))

        assert [combo.itemText(i) for i in range(combo.count())] == ['', '山本', '植田', '退職者']
        assert combo.currentText() == '退職者'

    def test_set_model_data(self, window: DocumentListWindow):
        delegate = StaffComboDelegate(window)
        combo = QComboBox()
        combo.addItems(['', '山本'])
        combo.setCurrentText('山本')
        model = MagicMock()
        index = window.model.index(0, STAFF_COLUMN)

        delegate.setModelData(combo, model, index)

        model.setData.assert_called_once()
        assert model.setData.call_args[0][1] == '山本'
