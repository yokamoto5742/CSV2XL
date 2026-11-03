import dataclasses
import sqlite3
from pathlib import Path

from services.database import connect
from services.document_repository import DocumentRepository
from tests.conftest import make_new_document


class TestAddNewDocuments:
    def test_adds_documents(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)

        added = repository.add_new_documents([make_new_document(1), make_new_document(2)])

        documents = repository.fetch_all()
        assert added == 2
        assert [doc.patient_id for doc in documents] == [1, 2]
        assert documents[0].staff_name is None
        assert documents[0].version == 1

    def test_skips_existing_documents(self, connection: sqlite3.Connection) -> None:
        """A〜F 列が既存行と同じものは追加しない"""
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])

        added = repository.add_new_documents([make_new_document(1), make_new_document(2)])

        assert added == 1
        assert len(repository.fetch_all()) == 2

    def test_keeps_duplicates_in_same_batch(self, connection: sqlite3.Connection) -> None:
        """同じ CSV 内の重複は除外しない(Excel 版と同じ動き)"""
        repository = DocumentRepository(connection)

        added = repository.add_new_documents([make_new_document(1), make_new_document(1)])

        assert added == 2

    def test_empty_list(self, connection: sqlite3.Connection) -> None:
        assert DocumentRepository(connection).add_new_documents([]) == 0


class TestUpdateEditableFields:
    def test_updates_fields(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document()])
        document = repository.fetch_all()[0]
        edited = dataclasses.replace(document, staff_name='山本', requested_date='2025-01-15', note='至急')

        saved = repository.update_editable_fields(edited)

        assert saved is not None
        assert saved.version == 2
        reloaded = repository.fetch_all()[0]
        assert (reloaded.staff_name, reloaded.requested_date, reloaded.note) == ('山本', '2025-01-15', '至急')
        assert reloaded.version == 2

    def test_conflict_returns_none(self, tmp_path: Path) -> None:
        """別端末が先に更新していたら保存しない"""
        db_path = str(tmp_path / 'shared.db')
        conn_a = connect(db_path)
        conn_b = connect(db_path)
        try:
            repo_a = DocumentRepository(conn_a)
            repo_b = DocumentRepository(conn_b)
            repo_a.add_new_documents([make_new_document()])
            doc_a = repo_a.fetch_all()[0]
            doc_b = repo_b.fetch_all()[0]

            assert repo_b.update_editable_fields(dataclasses.replace(doc_b, note='B')) is not None
            assert repo_a.update_editable_fields(dataclasses.replace(doc_a, note='A')) is None
            assert repo_a.fetch_all()[0].note == 'B'
        finally:
            conn_a.close()
            conn_b.close()


class TestDataVersion:
    def test_changes_when_other_connection_commits(self, tmp_path: Path) -> None:
        db_path = str(tmp_path / 'shared.db')
        conn_a = connect(db_path)
        conn_b = connect(db_path)
        try:
            repo_a = DocumentRepository(conn_a)
            before = repo_a.data_version()

            DocumentRepository(conn_b).add_new_documents([make_new_document()])

            assert repo_a.data_version() != before
        finally:
            conn_a.close()
            conn_b.close()
