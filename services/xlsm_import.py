import datetime
import re
import sqlite3
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from services.document_repository import DocumentRepository, ImportedDocument, NewDocument
from services.file_manager import backup_database
from services.list_repository import STAFF, ListRepository
from services.validators import parse_date
from utils.config_manager import ConfigManager

# 医療文書担当一覧_YYYYMMDDhhmm.xlsm の日時の部分
XLSM_FILE_PATTERN = re.compile(r'^医療文書担当一覧_(\d{12})\.xlsm$', re.IGNORECASE)
REQUIRED_COLUMNS = ('預り日', '患者ID', '患者名', '文書名', '診療科', '医師名', '担当者名', '医師依頼日', '備考')


@dataclass
class XlsmImportResult:
    file_name: str
    added_count: int = 0
    updated_count: int = 0
    # 取り込まなかった行(「n行目: 理由」)
    errors: list[str] = field(default_factory=list)


def find_latest_xlsm(directory: str) -> Path | None:
    """ファイル名の日時が最も新しい医療文書担当一覧_YYYYMMDDhhmm.xlsm を返す"""
    candidates = [(match.group(1), path) for path in Path(directory).glob('*.xlsm')
                  if (match := XLSM_FILE_PATTERN.match(path.name))]
    if not candidates:
        return None
    return max(candidates)[1]


def to_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def to_date_text(value: object) -> str | None:
    """日付・Excel のシリアル値・'YYYY/MM/DD' の文字列を 'YYYY-MM-DD' にする。読めなければ ValueError"""
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return from_excel(value).date().isoformat()
    text = to_text(value)
    if text is None:
        return None
    # 時刻が付いた文字列('2025/01/10 0:00:00')は日付の部分だけを使う
    return parse_date(text.split()[0])


def to_document(cells: Mapping[str, object], staff_names: set[str]) -> ImportedDocument:
    """1行を取り込む形にする。取り込めない行は理由を入れた ValueError"""
    try:
        received_date = to_date_text(cells['預り日'])
    except ValueError:
        raise ValueError('預り日が日付ではありません') from None
    if received_date is None:
        raise ValueError('預り日が空です')
    patient_id_text = (to_text(cells['患者ID']) or '').replace(',', '')
    if not patient_id_text.isdecimal():
        raise ValueError('患者IDが数値ではありません')
    try:
        requested_date = to_date_text(cells['医師依頼日'])
    except ValueError:
        raise ValueError('医師依頼日が日付ではありません') from None
    if requested_date is not None and requested_date < received_date:
        raise ValueError('医師依頼日が預り日より前です')
    staff_name = to_text(cells['担当者名'])
    if staff_name is not None and staff_name not in staff_names:
        raise ValueError(f'担当者名「{staff_name}」が担当者名のリストにありません')

    new_document = NewDocument(
        received_date=received_date,
        patient_id=int(patient_id_text),
        patient_name=to_text(cells['患者名']) or '',
        document_name=to_text(cells['文書名']) or '',
        department=to_text(cells['診療科']) or '',
        doctor_name=to_text(cells['医師名']) or '',
    )
    return ImportedDocument(new_document, staff_name, requested_date, to_text(cells['備考']))


def read_xlsm(path: Path, staff_names: set[str], result: XlsmImportResult) -> list[ImportedDocument]:
    """1行目を見出しとして読み、取り込む行を返す。取り込めない行は result.errors に入れる"""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        assert sheet is not None
        rows = sheet.iter_rows(values_only=True)
        headers = [to_text(value) for value in next(rows, ())]
        missing = [column for column in REQUIRED_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"{path.name} に列がありません: {', '.join(missing)}")

        documents = []
        for row_number, row in enumerate(rows, start=2):
            cells = {column: row[headers.index(column)] if headers.index(column) < len(row) else None
                     for column in REQUIRED_COLUMNS}
            if all(to_text(value) is None for value in cells.values()):
                continue
            try:
                documents.append(to_document(cells, staff_names))
            except ValueError as e:
                result.errors.append(f"{row_number}行目: {e}")
        return documents
    finally:
        workbook.close()


def import_xlsm_to_database(connection: sqlite3.Connection) -> XlsmImportResult:
    """xlsm フォルダの最新の医療文書担当一覧を DB に取り込む

    失敗時は例外を送出する(メッセージ表示は呼び出し側で行う)
    """
    xlsm_dir = ConfigManager().get_xlsm_path()
    latest_xlsm = find_latest_xlsm(xlsm_dir)
    if latest_xlsm is None:
        raise FileNotFoundError(f"{xlsm_dir} に医療文書担当一覧_年月日時分.xlsm が見つかりません。")

    result = XlsmImportResult(latest_xlsm.name)
    staff_names = set(ListRepository(connection).get_items(STAFF))
    documents = read_xlsm(latest_xlsm, staff_names, result)
    result.added_count, result.updated_count = DocumentRepository(connection).merge_documents(documents)
    backup_database(connection)
    return result
