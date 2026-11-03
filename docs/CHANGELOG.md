# CHANGELOG.md

このプロジェクトの全変更は、このファイルに記録されます。

フォーマットは[Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)に基づいており、
バージョン管理は[Semantic Versioning](https://semver.org/lang/ja/)に従っています。


## [2.0.0] - 2026-11-03

医療文書担当一覧.xlsm を廃止し、データを SQLite に保存してアプリだけで完結するようにした。

### 追加

- **担当一覧画面**: 担当一覧の表示、列見出しのクリックでの並び替え(既定は 預り日 → 診療科 → 患者ID)、担当者名(ドロップダウン)・医師依頼日・備考の入力
- **患者ID検索**: Excel のマクロと同じく、該当する行を「n/m件」で順に表示する。「前へ」「次へ」で移動する
- **医師依頼日の入力チェック**: `YYYY/MM/DD` / `YYYY-MM-DD` 以外や存在しない日付は保存しない
- **複数端末での共有**: DB を共有フォルダに置けるよう、ロールバックジャーナル・ロック待ち・楽観ロック(`version` 列)で同時編集に対応。5 秒ごとに他の端末の変更を読み込む
- **担当者名の設定**: 担当者名のリストを編集するダイアログを追加
- **移行スクリプト**: `scripts/migrate_from_xlsm.py` で、ブックの全行・担当者名・config.ini の除外リストを DB に移す
- **仕様書**: C# 版への移植用に `docs/spec.md` を追加

### 変更

- **CSV 取り込み**: 取り込み先を Excel から DB に変更(`csv_excel_transfer.py` → `csv_import.py`)。取り込み後に追加件数を表示して担当一覧を開く
- **除外リスト**: 除外する文書名・医師名を config.ini から DB に移し、全端末で同じ条件で取り込むようにした
- **バックアップ**: xlsm のコピーから、SQLite の backup API による `.db` のバックアップに変更
- **フォルダの場所**: 「Excelファイルパス」を「データベースファイル」に変更(既定は `C:\Shinseikai\CSV2XL\医療文書担当一覧.db`)
- **依存関係**: openpyxl を移行スクリプト用の開発依存に移した

### 削除

- **Excel 操作**: win32com による Excel の起動・並び替え・フィルタ解除、pyautogui によるブック共有ボタンのクリック(`excel_processor.py`)
- **画面の座標表示**: ブック共有ボタンの座標を調べる機能(`coordinate_tracker.py`)と `[ButtonPosition]` 設定
- **依存関係**: pywin32、pyautogui

## [1.2.0] - 2026-11-03

### 変更

- **ConfigManager(M1)**: getter の既定値の二重定義を削除し、setter を `_set` に共通化。既定フォルダを `DEFAULT_BASE_DIR` 定数にまとめた
- **ConfigManager(M2)**: カンマ区切りリストを読み書きする `get_list` / `set_list` を追加し、除外ダイアログが `config.config` を直接触らないようにした。`ExcludeItemDialog` の引数 `item_label` を `title` に統合
- **ConfigManager**: 保存失敗時の例外メッセージを `Failed to save config` に修正
- **エラー処理(M3)**: services 層から `QMessageBox` を除き、失敗時は例外を送出するように変更。エラー表示は `MainWindow.import_csv` の 1 箇所に集約（CSV 未検出やファイルロック時も同じエラーダイアログで表示）
- **エラー処理(M4)**: ログ出力して再送出するだけの `try/except` を `process_csv_data` / `process_completed_csv` / `sort_excel_data` / `backup_excel_file` / `ensure_directories_exist` から削除
- **テスト**: `test_csv_processor.py` にあった `transfer_csv_to_excel` の重複テストを削除（`test_csv_excel_transfer.py` に集約）
- **CSV 加工(M5)**: 元 CSV の列位置を `DOC_NAME_COL` / `DOCTOR_NAME_COL` / `DROP_COLS` 定数にし、列削除の前に除外フィルタを済ませるようにした（列番号が途中で変わらない）
- **Excel 書き込み(M6)**: 約 125 行の `write_data_to_excel` を `make_row_key` / `read_existing_keys` / `filter_new_rows` / `write_rows` に分割。重複判定の日付は Excel 側・CSV 側とも `YYYY-MM-DD` にそろえた

### 削除

- **ConfigManager**: 未使用の `set_processed_path` / `set_share_button_position` / `set_share_button_wait_time` / `set_backup_retention_days` / `_ensure_section` を削除
- **依存関係**: M6 で `to_numpy()` を使わなくなったため `numpy` を削除

### 修正

- **設定ファイル(H1)**: exe 実行時は exe と同じフォルダの `config.ini` を読み書きするように変更（無ければ同梱版をコピー）。ダイアログで変えた設定が再ビルドや onefile 構成で消えないようにするため
- **Excel 書式(H2)**: A〜F 列の垂直中央配置が水平配置の設定で上書きされて消えていた不具合を修正
- **フォントとウインドウサイズ(H3)**: 空欄や範囲外の値で OK を押すとアプリが落ちる/極小ウィンドウになる不具合を修正。不正値のときは警告して保存しない
- **Excel ファイル未検出(H4)**: Excel ファイルが無い/`.xlsm` でない場合に何も表示されなかったため、エラーメッセージを表示するように変更
- **ウィンドウ最小化(H5)**: Excel の処理に失敗したときに、前面にある別のウィンドウを最小化してしまう不具合を修正

## [1.1.5] - 2026-11-03

### 修正

- **テスト**: pyright の型エラー42件を解消（未使用の import・変数を削除し、None になり得る戻り値をチェックするように変更）
- **.gitignore**: `__pycache__/` を除外対象に追加

## [1.1.4] - 2026-11-03

### 変更

- **パッケージ管理**: uv に移行し、依存関係を pyproject.toml と uv.lock で管理するように変更
- **README.md**: uv でのインストール・実行手順、設定ファイルの `[Backup]` セクション、対応 Python バージョン（3.13 以上）を実態に合わせて更新
- **Stop フック**: 通知処理を `.claude/hooks/stop-notification.ps1` の呼び出しに統一

### 修正

- **pyright 設定**: 型チェック対象の `service` を `services` に修正し、services ディレクトリが型チェックされるようにした

### 削除

- **requirements.txt**: pyproject.toml に統合したため削除

## [1.1.3] - 2025-12-11

### 追加

- **Python型チェックツール**: CoverageおよびPython型チェック用ツールを追加（requirements.txt更新）

### 変更

- **pyrightconfig.json**: servicesディレクトリを型チェック対象に含める設定を更新
- **excel_processor テスト**: pathlib.Pathを使用した実装に変更
- **テスト(test_file_manager.py)**: ConfigManager をモック化し、バックアップ保持期間の設定値を使用するテストケースに更新

### 修正

- **excel_processor**: セル値取得時のNoneチェックを追加し、None値処理の堅牢性を向上

## [1.1.2] - 2025-12-10

### 追加

- **Excelフィルタ解除機能**: Excelファイル処理時に既存フィルタを自動解除する処理を追加
- **セットアップドキュメント**: インストールおよび使用方法をREADME.mdで詳細記載

### 変更

- **docs/README.md**: 完全改訂。主な機能、動作環境、インストール手順、使い方、プロジェクト構成、設定ファイル説明、開発情報を追加
- **.gitignore**: IDE設定ファイル、環境変数ファイル、ログファイルを除外対象に追加
- **requirements.txt**: パッケージ管理の整理

### 削除

- **古い関数一覧ドキュメント**: 重複ドキュメントを削除し、README.mdに統合

## [1.1.1] - 2025-09-04

### 修正

- バージョン情報の更新

## [1.1.0] - 以前のバージョン

詳細は個別のコミット履歴を参照してください。

---

[未リリース (Unreleased)]: https://github.com/yokamura/CSV2XL/compare/v1.1.3...HEAD
[1.1.3]: https://github.com/yokamura/CSV2XL/compare/v1.1.2...v1.1.3
[1.1.2]: https://github.com/yokamura/CSV2XL/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/yokamura/CSV2XL/releases/tag/v1.1.1
