import shutil
from pathlib import Path

import polars as pl

from utils.config_manager import ConfigManager


def read_csv_with_encoding(file_path):
    encodings = [
        'shift-jis', 'utf-8', 'cp932', 'euc-jp', 'iso-2022-jp',
        'utf-8-sig', 'latin1', 'shift_jis_2004', 'shift_jisx0213'
    ]

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

    # 最後の手段：エラー処理を含む読み込み
    for encoding in ['cp932', 'shift-jis', 'utf-8']:
        try:
            schema = {
                "患者ID": pl.Int64,
            }
            df = pl.read_csv(
                file_path,
                encoding=encoding,
                separator=',',
                skip_rows=3,
                has_header=True,
                infer_schema_length=0,
                schema_overrides=schema,
                ignore_errors=True  # エラーを無視して読み込み
            )
            
            if len(df.columns) > 1:
                print(f"エンコーディング {encoding} (エラー無視) で読み込みました")
                return df
        except Exception as e:
            print(f"{encoding} (エラー無視) での読み込み試行中にエラー: {str(e)}")
            continue

    raise Exception("CSVファイルの読み込みに失敗しました")


def process_csv_data(df):
    try:
        # 列名を一意にする
        original_columns = df.columns
        unique_columns = []
        for i, col in enumerate(original_columns):
            unique_columns.append(f"col_{i}_{col}")
        df = df.select([
            pl.col(old_name).alias(new_name)
            for old_name, new_name in zip(original_columns, unique_columns)
        ])

        # K列とI列を削除 (インデックスベースで削除)
        columns_to_keep = [i for i in range(len(df.columns)) if i not in [8, 10]]
        df = df.select([df.columns[i] for i in columns_to_keep])

        # A列からC列を削除
        df = df.select(df.columns[3:])

        # 文書名（D列→削除後は1列目）と医師名（F列→削除後は3列目）のスペースと*を常に除去
        df = df.with_columns([
            pl.col(df.columns[1]).str.replace_all(r'[\s*]', ''),  # 文書名
            pl.col(df.columns[3]).str.replace_all(r'[\s*]', '')  # 医師名
        ])

        config = ConfigManager()
        exclude_docs = config.get_exclude_docs()
        exclude_doctors = config.get_exclude_doctors()

        # 除外する文書名のフィルタリング
        if exclude_docs:
            for doc in exclude_docs:
                df = df.filter(~pl.col(df.columns[1]).str.contains(doc))

        # 除外する医師名のフィルタリング
        if exclude_doctors:
            for doctor in exclude_doctors:
                df = df.filter(~pl.col(df.columns[3]).str.contains(doctor))

        return df

    except Exception as e:
        print(f"データ処理中にエラーが発生しました: {str(e)}")
        raise


def convert_date_format(df):
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


def process_completed_csv(csv_path: str):
    try:
        csv_file = Path(csv_path)
        if not csv_file.exists():
            return

        config = ConfigManager()
        processed_dir = Path(config.get_processed_path())
        processed_dir.mkdir(exist_ok=True, parents=True)

        new_path = processed_dir / csv_file.name
        shutil.move(str(csv_file), str(new_path))

    except Exception as e:
        print(f"CSVファイルの処理中にエラーが発生しました: {str(e)}")
        raise


def find_latest_csv(downloads_path):
    # ファイル名の形式がYYYYMMDD_HHmmss.csvのファイルを検索
    csv_files = [f for f in Path(downloads_path).glob('*.csv')
                 if len(f.name.split('_')) == 2 and
                 (3 <= len(f.name.split('_')[0]) <= 4) and
                 len(f.name.split('_')[1].split('.')[0]) == 14]

    if not csv_files:
        return None

    return str(max(csv_files, key=lambda f: f.stat().st_mtime))
