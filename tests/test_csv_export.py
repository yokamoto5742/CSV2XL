from pathlib import Path

from services.csv_export import write_csv


def test_write_csv_with_bom(tmp_path: Path) -> None:
    """BOM 付き UTF-8 で見出しと行を書き出す(CP932 の外字も化けない)"""
    path = tmp_path / 'export.csv'

    write_csv(str(path), ['担当者名', '備考'], [['花﨑', 'a,b'], ['', '']])

    raw = path.read_bytes()
    assert raw.startswith(b'\xef\xbb\xbf')
    assert raw.decode('utf-8-sig') == '担当者名,備考\r\n花﨑,"a,b"\r\n,\r\n'
