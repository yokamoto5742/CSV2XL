import datetime
import sqlite3
from pathlib import Path
from unittest.mock import patch

import pytest
from openpyxl import Workbook

from services.document_repository import DocumentRepository
from services.list_repository import STAFF, ListRepository
from services.xlsm_import import XlsmImportResult, find_latest_xlsm, import_xlsm_to_database, read_xlsm, to_date_text

HEADERS = ["預り日", "患者ID", "患者名", "文書名", "診療科", "医師名", "担当者名", "医師依頼日", "備考"]
STAFF_NAMES = {'山本', '植田'}


def save_xlsm(path: Path, rows: list[list[object]], headers: list[str] = HEADERS) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


def make_row(staff_name: object = '山本', requested_date: object = datetime.datetime(2025, 1, 15),
             received_date: object = datetime.datetime(2025, 1, 10), patient_id: object = 1001) -> list[object]:
    return [received_date, patient_id, '山田　太郎', '診断書', '内科', '田中', staff_name, requested_date, '至急']


class TestFindLatestXlsm:
    def test_returns_newest_by_file_name(self, tmp_path: Path) -> None:
        for name in ('医療文書担当一覧_202501101200.xlsm', '医療文書担当一覧_202502011200.xlsm',
                     '~$医療文書担当一覧_202503011200.xlsm', '医療文書担当一覧.xlsm'):
            (tmp_path / name).touch()

        assert find_latest_xlsm(str(tmp_path)) == tmp_path / '医療文書担当一覧_202502011200.xlsm'

    def test_returns_none_when_missing(self, tmp_path: Path) -> None:
        assert find_latest_xlsm(str(tmp_path)) is None


class TestToDateText:
    @pytest.mark.parametrize('value, expected', [
        (datetime.datetime(2025, 1, 10, 9, 30), '2025-01-10'),
        (datetime.date(2025, 1, 10), '2025-01-10'),
        (45667, '2025-01-10'),
        ('2025/1/10', '2025-01-10'),
        ('2025/01/10 0:00:00', '2025-01-10'),
        (None, None),
        ('  ', None),
    ])
    def test_converts(self, value: object, expected: str | None) -> None:
        assert to_date_text(value) == expected

    def test_invalid_text(self) -> None:
        with pytest.raises(ValueError):
            to_date_text('不明')


class TestReadXlsm:
    def test_reads_rows(self, tmp_path: Path) -> None:
        path = save_xlsm(tmp_path / 'test.xlsm', [make_row(), make_row(staff_name=None, requested_date=None)])
        result = XlsmImportResult(path.name)

        documents = read_xlsm(path, STAFF_NAMES, result)

        assert result.errors == []
        assert documents[0].new_document.key() == ('2025-01-10', 1001, '山田　太郎', '診断書', '内科', '田中')
        assert (documents[0].staff_name, documents[0].requested_date, documents[0].note) == (
            '山本', '2025-01-15', '至急')
        assert (documents[1].staff_name, documents[1].requested_date) == (None, None)

    def test_skips_empty_rows(self, tmp_path: Path) -> None:
        path = save_xlsm(tmp_path / 'test.xlsm', [[None] * 9, make_row()])

        assert len(read_xlsm(path, STAFF_NAMES, XlsmImportResult(path.name))) == 1

    @pytest.mark.parametrize('row, reason', [
        (make_row(received_date=None), '預り日が空です'),
        (make_row(received_date='不明'), '預り日が日付ではありません'),
        (make_row(patient_id='abc'), '患者IDが数値ではありません'),
        (make_row(requested_date='不明'), '医師依頼日が日付ではありません'),
        (make_row(requested_date=datetime.datetime(2025, 1, 9)), '医師依頼日が預り日より前です'),
        (make_row(staff_name='佐藤'), '担当者名「佐藤」が担当者名のリストにありません'),
    ])
    def test_reports_invalid_rows(self, tmp_path: Path, row: list[object], reason: str) -> None:
        path = save_xlsm(tmp_path / 'test.xlsm', [row])
        result = XlsmImportResult(path.name)

        assert read_xlsm(path, STAFF_NAMES, result) == []
        assert result.errors == [f'2行目: {reason}']

    def test_missing_column(self, tmp_path: Path) -> None:
        path = save_xlsm(tmp_path / 'test.xlsm', [], headers=HEADERS[:-1])

        with pytest.raises(ValueError, match='備考'):
            read_xlsm(path, STAFF_NAMES, XlsmImportResult(path.name))


class TestImportXlsmToDatabase:
    @patch('services.xlsm_import.backup_database')
    @patch('services.xlsm_import.ConfigManager')
    def test_imports_latest_xlsm(self, mock_config_manager, mock_backup, connection: sqlite3.Connection,
                                 tmp_path: Path) -> None:
        mock_config_manager.return_value.get_xlsm_path.return_value = str(tmp_path)
        ListRepository(connection).set_items(STAFF, ['山本'])
        save_xlsm(tmp_path / '医療文書担当一覧_202501101200.xlsm', [make_row(), make_row(staff_name='佐藤')])

        result = import_xlsm_to_database(connection)

        assert (result.file_name, result.added_count, result.updated_count) == (
            '医療文書担当一覧_202501101200.xlsm', 1, 0)
        assert len(result.errors) == 1
        assert DocumentRepository(connection).fetch_all()[0].staff_name == '山本'
        mock_backup.assert_called_once_with(connection)

    @patch('services.xlsm_import.ConfigManager')
    def test_no_xlsm(self, mock_config_manager, connection: sqlite3.Connection, tmp_path: Path) -> None:
        mock_config_manager.return_value.get_xlsm_path.return_value = str(tmp_path)

        with pytest.raises(FileNotFoundError):
            import_xlsm_to_database(connection)
