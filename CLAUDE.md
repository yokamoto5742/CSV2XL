# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.必ず日本語で回答してください。

コーディング規約・テスト・コミット・回答スタイルは `.claude/rules/` を参照する。

## コマンド

```bash
uv run python main.py    # アプリ起動
uv run pyright           # 型チェック(設定は pyproject.toml の [tool.pyright])
python build.py          # PyInstaller で dist/CSV2XL.exe を作成(バージョンは自動で上がらない)
```

## アーキテクチャ

PyQt6 製の Windows 専用 GUI。ダウンロードフォルダの最新 CSV(`XXXX_YYYYMMDDhhmmss.csv`)を polars で読み込んで SQLite に取り込み、担当一覧画面で表示・検索・編集する。CSV は processed フォルダへ移動する。

- `app/`: GUI(メインウィンドウ・担当一覧・設定ダイアログ)
- `services/csv_import.py`: CSV 取り込みの流れを管理する
- `services/schema.sql`・`document_repository.py`・`list_repository.py`: DB のスキーマと読み書き
- `utils/config_manager.py`: `utils/config.ini` の読み書き
- `docs/spec.md`: 仕様書。将来 C# (.NET) で作り直すため、動作を変えたらここも更新する

## 注意点

- DB は複数 PC から共有フォルダ越しに使う。WAL は使わず、書き込みは短いトランザクションにする。G〜I 列の更新は `version` 列で楽観ロックする。
- 担当者名・除外する文書名・除外する医師名は DB の `list_items` に保存する(全端末共通)。config.ini には端末ごとの設定だけを置く。
- テストの DB は `tests/conftest.py` の `connection` fixture(一時フォルダ)を使い、本番の DB パスに接続しない。
- CSV の読み込みは Shift-JIS → CP932 → UTF-8 の順にエンコーディングを試す。
- `utils/config.ini` は exe に同梱される。exe 実行時は exe と同じフォルダの `config.ini` を読み書きし、無ければ同梱版(`sys._MEIPASS` 内)をコピーする。設定の既定値を変えたら再ビルドが必要。
- `utils/config.ini` のパスはこの PC 固有の値。勝手に書き換えない。
