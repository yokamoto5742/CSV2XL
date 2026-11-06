import dataclasses
import sqlite3
from pathlib import Path

from services.database import connect
from services.document_repository import DocumentRepository, ImportedDocument
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


class TestMergeDocuments:
    def test_adds_new_document_with_editable_fields(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)

        counts = repository.merge_documents([ImportedDocument(make_new_document(1), '山本', '2025-01-15', '至急')])

        document = repository.fetch_all()[0]
        assert counts == (1, 0)
        assert (document.staff_name, document.requested_date, document.note) == ('山本', '2025-01-15', '至急')

    def test_fills_only_blank_fields(self, connection: sqlite3.Connection) -> None:
        """既存行の空欄だけを埋め、入力済みの値は上書きしない"""
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])
        existing = repository.fetch_all()[0]
        repository.update_editable_fields(dataclasses.replace(existing, staff_name='植田'))

        counts = repository.merge_documents([ImportedDocument(make_new_document(1), '山本', '2025-01-15', None)])

        document = repository.fetch_all()[0]
        assert counts == (0, 1)
        assert (document.staff_name, document.requested_date, document.note) == ('植田', '2025-01-15', None)
        assert document.version == 3

    def test_skips_duplicate_without_new_values(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])

        counts = repository.merge_documents([ImportedDocument(make_new_document(1), None, None, None)])

        assert counts == (0, 0)
        assert repository.fetch_all()[0].version == 1

    def test_pairs_same_key_rows_in_order(self, connection: sqlite3.Connection) -> None:
        """A〜F 列が同じ行は登録順に1対1で対応させ、余った行は追加する"""
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])

        counts = repository.merge_documents([ImportedDocument(make_new_document(1), '山本', None, None),
                                             ImportedDocument(make_new_document(1), '植田', None, None)])

        assert counts == (1, 1)
        assert [doc.staff_name for doc in repository.fetch_all()] == ['山本', '植田']

    def test_does_not_restore_deleted_document(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])
        repository.delete_document(repository.fetch_all()[0])

        counts = repository.merge_documents([ImportedDocument(make_new_document(1), '山本', None, None)])

        assert counts == (0, 0)
        assert repository.fetch_all() == []


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


class TestAddDocument:
    def test_adds_even_if_duplicate(self, connection: sqlite3.Connection) -> None:
        """手動の追加は重複チェックをしない"""
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])

        repository.add_document(make_new_document(1))

        assert [doc.patient_id for doc in repository.fetch_all()] == [1, 1]


class TestDeleteDocument:
    def test_hides_deleted_document(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1), make_new_document(2)])

        assert repository.delete_document(repository.fetch_all()[0]) is True

        assert [doc.patient_id for doc in repository.fetch_all()] == [2]

    def test_csv_does_not_restore_deleted_document(self, connection: sqlite3.Connection) -> None:
        """削除した行と同じ内容を取り込んでも追加しない"""
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1)])
        repository.delete_document(repository.fetch_all()[0])

        assert repository.add_new_documents([make_new_document(1)]) == 0
        assert repository.fetch_all() == []

    def test_conflict_returns_false(self, tmp_path: Path) -> None:
        """別端末が先に更新していたら削除しない"""
        db_path = str(tmp_path / 'shared.db')
        conn_a = connect(db_path)
        conn_b = connect(db_path)
        try:
            repo_a = DocumentRepository(conn_a)
            repo_b = DocumentRepository(conn_b)
            repo_a.add_new_documents([make_new_document()])
            doc_a = repo_a.fetch_all()[0]
            repo_b.update_editable_fields(dataclasses.replace(repo_b.fetch_all()[0], note='B'))

            assert repo_a.delete_document(doc_a) is False
            assert len(repo_a.fetch_all()) == 1
        finally:
            conn_a.close()
            conn_b.close()


class TestDeleteDocuments:
    def test_deletes_all(self, connection: sqlite3.Connection) -> None:
        repository = DocumentRepository(connection)
        repository.add_new_documents([make_new_document(1), make_new_document(2), make_new_document(3)])
        documents = repository.fetch_all()

        assert repository.delete_documents([documents[0], documents[2]]) is True

        assert [doc.patient_id for doc in repository.fetch_all()] == [2]

    def test_conflict_deletes_nothing(self, tmp_path: Path) -> None:
        """1行でも別端末が先に更新していたら、1行も削除しない"""
        db_path = str(tmp_path / 'shared.db')
        conn_a = connect(db_path)
        conn_b = connect(db_path)
        try:
            repo_a = DocumentRepository(conn_a)
            repo_b = DocumentRepository(conn_b)
            repo_a.add_new_documents([make_new_document(1), make_new_document(2)])
            docs_a = repo_a.fetch_all()
            repo_b.update_editable_fields(dataclasses.replace(repo_b.fetch_all()[1], note='B'))

            assert repo_a.delete_documents(docs_a) is False
            assert len(repo_a.fetch_all()) == 2
            assert len(repo_b.fetch_all()) == 2
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
