import dataclasses
import sqlite3
import sys
from unittest.mock import MagicMock, patch

import pytest
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

from app.document_table_model import (
    NOTE_COLUMN, REQUESTED_DATE_COLUMN, STAFF_COLUMN, DocumentTableModel, sort_documents
)
from services.document_repository import Document, DocumentRepository
from tests.conftest import make_new_document


@pytest.fixture
def app():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


@pytest.fixture
def model(app, connection: sqlite3.Connection) -> DocumentTableModel:
    repository = DocumentRepository(connection)
    repository.add_new_documents([
        make_new_document(300, '2025-01-11', '内科'),
        make_new_document(200, '2025-01-10', '眼科'),
        make_new_document(100, '2025-01-10', '内科'),
    ])
    table_model = DocumentTableModel(repository)
    table_model.load()
    return table_model


def make_document(doc_id: int, received_date: str, department: str, patient_id: int,
                  note: str | None = None) -> Document:
    return Document(doc_id, received_date, patient_id, '患者', '文書', department, '医師',
                    None, None, note, 1)


class TestSortDocuments:
    def test_default_order(self):
        """既定は 預り日 → 診療科 → 患者ID の昇順"""
        documents = [
            make_document(1, '2025-01-11', 'A', 1),
            make_document(2, '2025-01-10', 'B', 1),
            make_document(3, '2025-01-10', 'A', 9),
            make_document(4, '2025-01-10', 'A', 5),
        ]
        assert [doc.id for doc in sort_documents(documents, -1, False)] == [4, 3, 2, 1]

    def test_column_order_keeps_empty_last(self):
        """列で並べたとき、空欄は昇順・降順とも末尾"""
        documents = [
            make_document(1, '2025-01-10', 'A', 1, note=None),
            make_document(2, '2025-01-10', 'A', 2, note='b'),
            make_document(3, '2025-01-10', 'A', 3, note='a'),
        ]
        assert [doc.id for doc in sort_documents(documents, NOTE_COLUMN, False)] == [3, 2, 1]
        assert [doc.id for doc in sort_documents(documents, NOTE_COLUMN, True)] == [2, 3, 1]


class TestDocumentTableModel:
    def test_load_uses_default_order(self, model: DocumentTableModel):
        assert [model.document(row).patient_id for row in range(model.rowCount())] == [100, 200, 300]

    def test_display_formats(self, model: DocumentTableModel):
        """日付は YYYY/MM/DD、空欄は空文字で表示する"""
        assert model.data(model.index(0, 0)) == '2025/01/10'
        assert model.data(model.index(0, 1)) == '100'
        assert model.data(model.index(0, STAFF_COLUMN)) == ''

    def test_only_g_to_i_are_editable(self, model: DocumentTableModel):
        editable = [col for col in range(model.columnCount())
                    if model.flags(model.index(0, col)) & Qt.ItemFlag.ItemIsEditable]
        assert editable == [STAFF_COLUMN, REQUESTED_DATE_COLUMN, NOTE_COLUMN]

    def test_header(self, model: DocumentTableModel):
        headers = [model.headerData(col, Qt.Orientation.Horizontal) for col in range(model.columnCount())]
        assert headers == ["預り日", "患者ID", "患者名", "文書名", "診療科", "医師名", "担当者名", "医師依頼日", "備考"]

    def test_set_data_saves_to_database(self, model: DocumentTableModel):
        assert model.setData(model.index(0, STAFF_COLUMN), '山本')
        assert model.setData(model.index(0, REQUESTED_DATE_COLUMN), '2025/1/15')
        assert model.setData(model.index(0, NOTE_COLUMN), ' 至急 ')

        saved = next(doc for doc in model.repository.fetch_all() if doc.patient_id == 100)
        assert (saved.staff_name, saved.requested_date, saved.note) == ('山本', '2025-01-15', '至急')
        assert model.data(model.index(0, REQUESTED_DATE_COLUMN)) == '2025/01/15'

    def test_set_data_clears_value(self, model: DocumentTableModel):
        model.setData(model.index(0, REQUESTED_DATE_COLUMN), '2025/01/15')

        assert model.setData(model.index(0, REQUESTED_DATE_COLUMN), '')
        assert model.document(0).requested_date is None

    def test_invalid_date_is_rejected(self, model: DocumentTableModel):
        errors = MagicMock()
        model.error_occurred.connect(errors)

        assert not model.setData(model.index(0, REQUESTED_DATE_COLUMN), '2025/02/30')

        errors.assert_called_once()
        assert "YYYY/MM/DD" in errors.call_args[0][0]
        assert model.document(0).requested_date is None

    def test_conflict_emits_signal(self, model: DocumentTableModel):
        conflicts = MagicMock()
        model.conflict_detected.connect(conflicts)

        with patch.object(model.repository, 'update_editable_fields', return_value=None):
            assert not model.setData(model.index(0, NOTE_COLUMN), 'メモ')

        conflicts.assert_called_once()

    def test_database_error_emits_signal(self, model: DocumentTableModel):
        errors = MagicMock()
        model.error_occurred.connect(errors)

        with patch.object(model.repository, 'update_editable_fields',
                          side_effect=sqlite3.OperationalError('database is locked')):
            assert not model.setData(model.index(0, NOTE_COLUMN), 'メモ')

        assert "database is locked" in errors.call_args[0][0]

    def test_sort_by_column(self, model: DocumentTableModel):
        model.sort(1, Qt.SortOrder.DescendingOrder)
        assert [model.document(row).patient_id for row in range(3)] == [300, 200, 100]

        # 読み直しても並び順を保つ
        model.load()
        assert [model.document(row).patient_id for row in range(3)] == [300, 200, 100]

    def test_find_rows(self, model: DocumentTableModel):
        assert model.find_rows(200) == [1]
        assert model.find_rows(999) == []

    def test_find_row_by_id(self, model: DocumentTableModel):
        assert model.find_row_by_id(model.document(2).id) == 2
        assert model.find_row_by_id(999) is None

    def test_export_rows(self, model: DocumentTableModel):
        """画面の並び順・表示形式のまま返す"""
        model.setData(model.index(0, REQUESTED_DATE_COLUMN), '2025/01/15')

        rows = model.export_rows()

        assert [row[1] for row in rows] == ['100', '200', '300']
        assert rows[0] == ['2025/01/10', '100', '山田　太郎', '診断書', '内科', '田中', '', '2025/01/15', '']

    def test_save_document_updates_row(self, model: DocumentTableModel):
        edited = dataclasses.replace(model.document(1), staff_name='山本', note='至急')

        assert model.save_document(edited)

        assert model.document(1).staff_name == '山本'
        assert model.document(1).version == edited.version + 1
        saved = next(doc for doc in model.repository.fetch_all() if doc.id == edited.id)
        assert (saved.staff_name, saved.note) == ('山本', '至急')


def patient_ids(model: DocumentTableModel) -> list[int]:
    return [model.document(row).patient_id for row in range(model.rowCount())]


class TestFilter:
    def test_filter_and_clear(self, model: DocumentTableModel):
        model.set_filter(4, {'内科'})
        assert patient_ids(model) == [100, 300]
        assert model.headerData(4, Qt.Orientation.Horizontal) == "診療科 ▼"

        model.set_filter(4, None)
        assert patient_ids(model) == [100, 200, 300]
        assert model.headerData(4, Qt.Orientation.Horizontal) == "診療科"

    def test_filters_on_several_columns(self, model: DocumentTableModel):
        model.set_filter(4, {'内科'})
        model.set_filter(0, {'2025/01/11'})

        assert patient_ids(model) == [300]

    def test_empty_value(self, model: DocumentTableModel):
        model.setData(model.index(0, STAFF_COLUMN), '山本')

        model.set_filter(STAFF_COLUMN, {''})

        assert patient_ids(model) == [200, 300]

    def test_filter_kept_after_sort_and_load(self, model: DocumentTableModel):
        model.set_filter(4, {'内科'})

        model.sort(1, Qt.SortOrder.DescendingOrder)
        assert patient_ids(model) == [300, 100]

        model.load()
        assert patient_ids(model) == [300, 100]

    def test_column_values(self, model: DocumentTableModel):
        """ほかの列のフィルタで残る行の値を、空欄を末尾にして並べる"""
        model.setData(model.index(0, STAFF_COLUMN), '山本')

        assert model.column_values(1) == ['100', '200', '300']
        assert model.column_values(STAFF_COLUMN) == ['山本', '']

        model.set_filter(4, {'眼科'})
        assert model.column_values(1) == ['200']
        # 自分の列のフィルタでは絞らない
        assert model.column_values(4) == ['内科', '眼科']

    def test_column_values_sorts_numbers(self, app, connection: sqlite3.Connection):
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(20), make_new_document(100)])
        table_model = DocumentTableModel(repository)
        table_model.load()

        assert table_model.column_values(1) == ['20', '100']

    def test_find_and_export_use_visible_rows(self, model: DocumentTableModel):
        model.set_filter(4, {'内科'})

        assert model.find_rows(200) == []
        assert [row[1] for row in model.export_rows()] == ['100', '300']

    def test_saved_row_stays_until_reload(self, model: DocumentTableModel):
        """保存した値がフィルタに合わなくなっても、読み直すまで行を残す"""
        model.set_filter(STAFF_COLUMN, {''})

        assert model.setData(model.index(0, STAFF_COLUMN), '山本')
        assert model.rowCount() == 3
        assert model.all_documents[0].staff_name == '山本'

        model.load()
        assert model.rowCount() == 2
