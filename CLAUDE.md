# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

コーディング規約・テスト・コミット・回答スタイルは `.claude/rules/` を参照する。

## コマンド

```bash
uv run python main.py    # アプリ起動
uv run pyright           # 型チェック(設定は pyproject.toml の [tool.pyright])
python build.py          # PyInstaller で dist/CSV2XL.exe を作成(バージョンは自動で上がらない)
```

## アーキテクチャ

PyQt6 製の Windows 専用 GUI。ダウンロードフォルダの最新 CSV(`XXXX_YYYYMMDDhhmmss.csv`)を polars で読み込み、openpyxl で `.xlsm` に書き込んだあと、CSV を processed フォルダへ移動する。

- `app/`: GUI(メインウィンドウ・設定ダイアログ)
- `services/csv_excel_transfer.py`: 処理全体の流れを管理する
- `utils/config_manager.py`: `utils/config.ini` の読み書き

## 注意点

- `services/excel_processor.py` の `open_and_sort_excel` は win32com で実際の Excel を操作し、pyautogui で `[ButtonPosition]` の座標をクリックする。テストでは win32com・pyautogui を必ずモックする。
- CSV の読み込みは Shift-JIS → CP932 → UTF-8 の順にエンコーディングを試す。
- `utils/config.ini` は exe に同梱され、exe 実行時は `sys._MEIPASS` 内のコピーを読む。設定の既定値を変えたら再ビルドが必要。
- `utils/config.ini` のパスはこの PC 固有の値。勝手に書き換えない。
