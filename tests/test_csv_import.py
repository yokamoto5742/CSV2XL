import sqlite3
from collections.abc import Iterator
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from services.csv_import import import_csv_to_database
from services.document_repository import DocumentRepository
from services.list_repository import EXCLUDE_DOCS, ListRepository

CSV_HEADER = "c0,c1,c2,預り日,患者ID,患者名,文書名,診療科,c8,医師名,c10"


def write_csv(path: Path, rows: list[str]) -> None:
    """Papyrus の CSV と同じく、先頭3行の後にヘッダーが来る CSV を Shift-JIS で書く"""
    lines = ["タイトル", "", "", CSV_HEADER, *rows]
    path.write_text("\n".join(lines) + "\n", encoding='shift-jis')


@pytest.fixture
def folders(tmp_path: Path) -> Iterator[dict[str, Path]]:
    """ConfigManager をモックし、ダウンロード・処理済み・バックアップを一時フォルダにする"""
    paths = {name: tmp_path / name for name in ('downloads', 'processed', 'backup')}
    mock_config = MagicMock()
    mock_config.get_downloads_path.return_value = str(paths['downloads'])
    mock_config.get_processed_path.return_value = str(paths['processed'])
    mock_config.get_backup_path.return_value = str(paths['backup'])
    mock_config.get_backup_retention_days.return_value = 14
    with patch('services.csv_import.ConfigManager', return_value=mock_config), \
            patch('services.csv_processor.ConfigManager', return_value=mock_config), \
            patch('services.file_manager.ConfigManager', return_value=mock_config):
        yield paths


class TestImportCsvToDatabase:
    def test_imports_latest_csv(self, connection: sqlite3.Connection, folders: dict[str, Path]) -> None:
        """CSV を取り込み、除外・重複除外をして、CSV の移動とバックアップを行う"""
        folders['downloads'].mkdir()
        csv_path = folders['downloads'] / '1234_20250105120000.csv'
        write_csv(csv_path, [
            "x,x,x,20250105,1001,山田　太郎,診断書 *,内科,x,田中,x",
            "x,x,x,20250105,1002,佐藤　花子,紹介状,外科,x,鈴木,x",
        ])
        ListRepository(connection).set_items(EXCLUDE_DOCS, ['紹介状'])

        added = import_csv_to_database(connection)

        documents = DocumentRepository(connection).fetch_all()
        assert added == 1
        assert [(d.received_date, d.patient_id, d.document_name) for d in documents] == [
            ('2025-01-05', 1001, '診断書')
        ]
        assert not csv_path.exists()
        assert (folders['processed'] / csv_path.name).exists()
        assert len(list(folders['backup'].glob('医療文書担当一覧_*.db'))) == 1

    def test_second_import_adds_nothing(self, connection: sqlite3.Connection, folders: dict[str, Path]) -> None:
        folders['downloads'].mkdir()
        row = "x,x,x,20250105,1001,山田　太郎,診断書,内科,x,田中,x"
        write_csv(folders['downloads'] / '1234_20250105120000.csv', [row])
        import_csv_to_database(connection)
        write_csv(folders['downloads'] / '1234_20250105130000.csv', [row])

        assert import_csv_to_database(connection) == 0

    def test_no_csv(self, connection: sqlite3.Connection, folders: dict[str, Path]) -> None:
        """CSVファイルが見つからない場合は例外を送出する"""
        with pytest.raises(FileNotFoundError, match="CSVファイルが見つかりません"):
            import_csv_to_database(connection)

    @patch('services.csv_import.read_csv_with_encoding', return_value=None)
    def test_read_failure(self, _mock_read: MagicMock, connection: sqlite3.Connection,
                          folders: dict[str, Path]) -> None:
        """CSVファイルを読み込めない場合は例外を送出し、CSV を移動しない"""
        folders['downloads'].mkdir()
        csv_path = folders['downloads'] / '1234_20250105120000.csv'
        write_csv(csv_path, [])

        with pytest.raises(ValueError, match="CSVファイルの読み込みに失敗しました"):
            import_csv_to_database(connection)
        assert csv_path.exists()
