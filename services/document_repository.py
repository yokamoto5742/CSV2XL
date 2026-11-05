import dataclasses
import sqlite3
from dataclasses import dataclass

from services.database import now_text

DocumentKey = tuple[str, int, str, str, str, str]
INSERT_SQL = (
    "INSERT INTO documents (received_date, patient_id, patient_name, document_name,"
    " department, doctor_name, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
)


@dataclass(frozen=True)
class NewDocument:
    """CSV から取り込む行(A〜F 列)"""
    received_date: str
    patient_id: int
    patient_name: str
    document_name: str
    department: str
    doctor_name: str

    def key(self) -> DocumentKey:
        """重複チェック用のキー"""
        return (self.received_date, self.patient_id, self.patient_name,
                self.document_name, self.department, self.doctor_name)


@dataclass(frozen=True)
class Document:
    """担当一覧の1行"""
    id: int
    received_date: str
    patient_id: int
    patient_name: str
    document_name: str
    department: str
    doctor_name: str
    staff_name: str | None
    requested_date: str | None
    note: str | None
    version: int


class DocumentRepository:
    """documents テーブルの読み書き"""

    def __init__(self, connection: sqlite3.Connection) -> None:
        self.connection = connection

    def fetch_all(self) -> list[Document]:
        rows = self.connection.execute(
            "SELECT id, received_date, patient_id, patient_name, document_name, department,"
            " doctor_name, staff_name, requested_date, note, version FROM documents"
            " WHERE deleted = 0 ORDER BY id"
        ).fetchall()
        return [Document(**dict(row)) for row in rows]

    def add_document(self, document: NewDocument) -> None:
        """手動で1行追加する(重複していても追加する)"""
        now = now_text()
        with self.connection:
            self.connection.execute(INSERT_SQL, (*document.key(), now, now))

    def delete_document(self, document: Document) -> bool:
        """削除済みにする。読み込み後に別端末で更新されていた場合は削除せずに False を返す"""
        with self.connection:
            cursor = self.connection.execute(
                "UPDATE documents SET deleted = 1, version = version + 1, updated_at = ?"
                " WHERE id = ? AND version = ? AND deleted = 0",
                (now_text(), document.id, document.version)
            )
        return cursor.rowcount > 0

    def add_new_documents(self, documents: list[NewDocument]) -> int:
        """既存行と A〜F 列が同じものを除いて追加し、追加した件数を返す

        同じ CSV 内の重複は既存の Excel 版と同じく除外しない
        """
        now = now_text()
        with self.connection:
            # 別端末の同時取り込みで重複しないよう、読み込みの前に書き込みロックを取る
            self.connection.execute('BEGIN IMMEDIATE')
            existing_keys = {tuple(row) for row in self.connection.execute(
                "SELECT received_date, patient_id, patient_name, document_name, department, doctor_name"
                " FROM documents"
            )}
            new_documents = [doc for doc in documents if doc.key() not in existing_keys]
            self.connection.executemany(INSERT_SQL, [(*doc.key(), now, now) for doc in new_documents])
        return len(new_documents)

    def update_editable_fields(self, document: Document) -> Document | None:
        """担当者名・医師依頼日・備考を保存する

        読み込み後に別端末で更新されていた場合は保存せずに None を返す
        """
        with self.connection:
            cursor = self.connection.execute(
                "UPDATE documents SET staff_name = ?, requested_date = ?, note = ?,"
                " version = version + 1, updated_at = ? WHERE id = ? AND version = ?",
                (document.staff_name, document.requested_date, document.note,
                 now_text(), document.id, document.version)
            )
        if cursor.rowcount == 0:
            return None
        return dataclasses.replace(document, version=document.version + 1)

    def data_version(self) -> int:
        """別の接続(他の端末)がコミットするたびに変わる値"""
        return self.connection.execute('PRAGMA data_version').fetchone()[0]
