import shutil
from pathlib import Path
from typing import Optional

import polars as pl

from services.document_repository import NewDocument
from utils.config_manager import ConfigManager

# 元CSVの列位置(0始まり)
DOC_NAME_COL = 6  # G列: 文書名
DOCTOR_NAME_COL = 9  # J列: 医師名
DROP_COLS = {0, 1, 2, 8, 10}  # A〜C列, I列, K列(Excelに出力しない)


def read_csv_with_encoding(file_path: str) -> Optional[pl.DataFrame]:
    """複数のエンコーディングを試してCSVファイルを読み込む"""
    encodings = ['shift-jis', 'cp932', 'utf-8']

    for encoding in encodings:
        try:
            schema = {
                "患者ID": pl.Int64,
            }
            df = pl.read_csv(
                file_path,
                encoding=encoding,
                separator=',',
                skip_rows=3,  # 最初の3行をスキップ
                has_header=True,  # 4行目をヘッダーとして使用
                infer_schema_length=0,
                schema_overrides=schema
            )

            if len(df.columns) > 1:
                print(f"エンコーディング {encoding} で正常に読み込みました")
                print(f"列数: {len(df.columns)}")
                print(f"行数: {len(df)}")
                print(f"列名: {df.columns}")
                return df
        except Exception as e:
            print(f"{encoding}での読み込み試行中にエラー: {str(e)}")
            continue

    print("すべてのエンコーディングでの読み込みに失敗しました")
    return None


def process_csv_data(df: pl.DataFrame, exclude_docs: list[str], exclude_doctors: list[str]) -> pl.DataFrame:
    """CSVデータをExcel出力用に加工
    列名の一意化、スペースと*の除去、指定列の削除、除外データのフィルタリング"""
    # 列名を一意にするためインデックスと元の名前を組み合わせて識別子を作成
    df = df.rename({name: f"col_{i}_{name}" for i, name in enumerate(df.columns)})

    doc_col = pl.col(df.columns[DOC_NAME_COL])
    doctor_col = pl.col(df.columns[DOCTOR_NAME_COL])

    # 文書名と医師名のスペースと全角スペース、*を除去
    df = df.with_columns([
        doc_col.cast(pl.String).str.replace_all(r'[\s*　]', ''),
        doctor_col.cast(pl.String).str.replace_all(r'[\s*　]', ''),
    ])

    for doc in exclude_docs:
        df = df.filter(~doc_col.str.contains(doc, literal=True))
    for doctor in exclude_doctors:
        df = df.filter(~doctor_col.str.contains(doctor, literal=True))

    columns_to_keep = [name for i, name in enumerate(df.columns) if i not in DROP_COLS]
    return df.select(columns_to_keep)


def convert_date_format(df: pl.DataFrame) -> pl.DataFrame:
    """日付列をYYYYMMDD形式からDate型に変換"""
    try:
        date_col = df.columns[0]
        df = df.with_columns([
            pl.col(date_col).str.strptime(pl.Date, format="%Y%m%d")
            .alias(date_col)
        ])
        return df
    except Exception as e:
        print(f"日付変換中にエラーが発生しましたが、処理を継続します: {str(e)}")
        return df


def dataframe_to_documents(df: pl.DataFrame) -> list[NewDocument]:
    """加工済みの DataFrame(預り日・患者ID・患者名・文書名・診療科・医師名の順)を取り込み用の行にする"""
    rows = df.select(pl.all().cast(pl.String)).rows()
    return [
        NewDocument(
            received_date=row[0] or '',
            patient_id=int(row[1]),
            patient_name=row[2] or '',
            document_name=row[3] or '',
            department=row[4] or '',
            doctor_name=row[5] or '',
        )
        for row in rows
    ]


def process_completed_csv(csv_path: str) -> None:
    """処理済みCSVファイルを指定ディレクトリに移動"""
    csv_file = Path(csv_path)
    if not csv_file.exists():
        return

    config = ConfigManager()
    processed_dir = Path(config.get_processed_path())
    processed_dir.mkdir(exist_ok=True, parents=True)

    new_path = processed_dir / csv_file.name
    shutil.move(str(csv_file), str(new_path))


def find_latest_csv(downloads_path: str) -> Optional[str]:
    """ダウンロードフォルダから最新のCSVファイルを取得

    職員ID_YYYYMMDDHHmmss形式のファイルを検索して最も新しいものを返す

    Args:
        downloads_path: ダウンロードフォルダのパス

    Returns:
        最新CSVファイルのパス、または見つからない場合はNone
    """
    csv_files = [f for f in Path(downloads_path).glob('*.csv')
                 if len(f.name.split('_')) == 2 and
                 (0 <= len(f.name.split('_')[0]) <= 5) and
                 len(f.name.split('_')[1].split('.')[0]) == 14]

    if not csv_files:
        return None

    return str(max(csv_files, key=lambda f: f.stat().st_mtime))
