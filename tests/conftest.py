import sqlite3
from collections.abc import Iterator
from pathlib import Path

import pytest

from services.database import connect
from services.document_repository import NewDocument


@pytest.fixture
def connection(tmp_path: Path) -> Iterator[sqlite3.Connection]:
    """一時フォルダの DB に接続する"""
    conn = connect(str(tmp_path / 'test.db'))
    yield conn
    conn.close()


def make_new_document(patient_id: int = 1001, received_date: str = '2025-01-10',
                      department: str = '内科', document_name: str = '診断書') -> NewDocument:
    return NewDocument(
        received_date=received_date,
        patient_id=patient_id,
        patient_name='山田　太郎',
        document_name=document_name,
        department=department,
        doctor_name='田中',
    )
