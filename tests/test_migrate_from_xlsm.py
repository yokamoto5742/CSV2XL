import datetime
import sqlite3
from unittest.mock import MagicMock

import pytest
from openpyxl import Workbook

from scripts.migrate_from_xlsm import migrate
from services.document_repository import DocumentRepository
from services.list_repository import EXCLUDE_DOCS, EXCLUDE_DOCTORS, STAFF, ListRepository


def make_workbook() -> Workbook:
    """医療文書担当一覧.xlsm と同じ構成(担当一覧・設定シート)のブックを作る"""
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "担当一覧"
    sheet.append(["預り日", "患者ID", "患者名", "文書名", "診療科", "医師名", "担当者名", "医師依頼日", "備考"])
    row = [datetime.datetime(2025, 1, 10), 1001, "山田　太郎", "診断書", "内科", "田中"]
    sheet.append([*row, "山本", datetime.datetime(2025, 1, 15), "至急"])
    sheet.append([*row, None, None, None])  # 実データにある A〜F 列が同じ行
    settings = workbook.create_sheet("設定")
    settings.append(["担当", "植田,山本", None, None, "診療科"])
    settings.append(["植田", None, None, None, "内科"])
    return workbook


@pytest.fixture
def config() -> MagicMock:
    mock_config = MagicMock()
    mock_config.get_list.side_effect = lambda section: {'ExcludeDocs': ['紹介状'], 'ExcludeDoctors': ['清水']}[section]
    return mock_config


class TestMigrate:
    def test_migrates_all_rows_and_lists(self, connection: sqlite3.Connection, config: MagicMock):
        count = migrate(connection, make_workbook(), config)

        documents = DocumentRepository(connection).fetch_all()
        assert count == 2
        assert len(documents) == 2
        first = documents[0]
        assert (first.received_date, first.patient_id, first.staff_name, first.requested_date, first.note) == (
            '2025-01-10', 1001, '山本', '2025-01-15', '至急'
        )
        assert documents[1].staff_name is None

        lists = ListRepository(connection)
        assert lists.get_items(STAFF) == ['植田', '山本']
        assert lists.get_items(EXCLUDE_DOCS) == ['紹介状']
        assert lists.get_items(EXCLUDE_DOCTORS) == ['清水']

    def test_refuses_second_run(self, connection: sqlite3.Connection, config: MagicMock):
        migrate(connection, make_workbook(), config)

        with pytest.raises(RuntimeError, match="すでに行がある"):
            migrate(connection, make_workbook(), config)
