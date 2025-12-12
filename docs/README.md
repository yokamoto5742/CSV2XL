# CSV2XL - 医療文書管理アプリケーション

CSVからExcelへの自動転送を行うPyQt6ベースのデスクトップアプリケーションです。医療文書管理システム「Papyrus」からのCSVデータを効率的にExcelファイルに転記します。

## バージョン情報

- **現在のバージョン**: 1.1.3
- **最終更新日**: 2025年12月11日

## 主な機能

- **自動CSVファイル検出**: 最新のCSVファイルを自動で検出し処理
- **エンコーディング自動判定**: Shift-JIS、UTF-8、CP932に対応
- **日付形式自動変換**: CSV内の日付を目的地の形式に自動変換
- **柔軟なフィルタリング**: 除外する文書名・医師名を設定可能
- **重複チェック**: Excelへのデータ転記時に重複を自動検出
- **自動バックアップ**: Excel処理前に自動でバックアップを作成
- **UI カスタマイズ**: フォント、ウィンドウサイズを自由に調整
- **座標トラッキングツール**: 自動化機能の座標設定をサポート

## 動作環境

| 項目 | 要件 |
|------|------|
| **OS** | Windows 11 |
| **Python** | 3.11 以上 |
| **メモリ** | 4GB 以上（推奨） |
| **ディスク** | 100MB 以上の空き容量（推奨） |

## インストール

### 1. リポジトリをクローン
```bash
git clone <repository-url>
cd CSV2XL
```

### 2. 仮想環境を作成
```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. 依存パッケージをインストール
```bash
pip install -r requirements.txt
```

### 4. アプリケーションを起動
```bash
python main.py
```

## 使い方

```bash
python main.py
```

### 基本的なワークフロー

1. **CSVファイルを準備**: Papyrusからダウンロードフォルダにエクスポート
2. **CSVファイル取り込みボタンをクリック**: 以下を自動実行
   - 最新のCSVを検出・読み込み
   - データを処理・変換
   - Excelに転記
   - バックアップを作成
   - 処理済みCSVを移動
3. **Excelが自動で開く**: データ確認と追加操作が可能

### 設定メニュー

| 項目 | 説明 | 既定値 |
|------|------|--------|
| **除外する文書名** | 転記対象外の文書種別（カンマ区切り） | - |
| **除外する医師名** | 転記対象外の医師（カンマ区切り） | - |
| **フォントサイズ** | UIのフォントサイズ | 11 |
| **ウィンドウサイズ** | 幅×高さ | 350×330 |
| **画面の座標表示** | マウス座標をリアルタイム表示し、自動化の座標を設定 | - |
| **ダウンロードフォルダ** | CSVファイルのソースパス | - |
| **Excelファイルパス** | データ転記先のExcelファイル | - |
| **バックアップフォルダ** | バックアップの保存先 | - |
| **処理済みCSVフォルダ** | 処理済みCSVの移動先 | - |

## プロジェクト構成

```
CSV2XL/
├── app/                          # UIレイヤー
│   ├── __init__.py              # バージョン情報
│   ├── main_window.py           # メインウィンドウ
│   └── dialogs.py               # 設定ダイアログ
├── services/                    # ビジネスロジック
│   ├── csv_excel_transfer.py   # 中核的なデータ転送処理
│   ├── csv_processor.py        # CSV読込・処理
│   ├── excel_processor.py      # Excel操作
│   ├── file_manager.py         # ファイル・ディレクトリ管理
│   └── coordinate_tracker.py   # 座標トラッキングUI
├── utils/                       # ユーティリティ
│   ├── config_manager.py       # 設定管理
│   └── config.ini              # 設定ファイル
├── scripts/                     # ビルド・補助スクリプト
│   └── version_manager.py      # バージョン管理
├── tests/                       # ユニットテスト
├── main.py                      # エントリーポイント
├── build.py                     # 実行ファイル生成
└── requirements.txt             # 依存パッケージ
```

## 主要機能の詳細

### CSVデータ処理

アプリケーションは以下のエンコーディングで自動判定してCSVを読み込みます：
- Shift-JIS（日本語対応の標準エンコーディング）
- UTF-8
- CP932（Windows日本語）

```python
from services.csv_processor import read_csv_with_encoding, process_csv_data, convert_date_format

# エンコーディング自動判定で読み込み、データ変換、日付形式の自動変換
df = read_csv_with_encoding("path/to/file.csv")
df = process_csv_data(df)
df = convert_date_format(df)
```

### Excel 操作と ファイル管理

```python
from services.excel_processor import write_data_to_excel
from services.file_manager import backup_excel_file, cleanup_old_csv_files

# Excel に書き込み（重複自動検出・セル書式の自動適用）
success = write_data_to_excel("path/to/file.xlsx", df)

# バックアップと古いファイルのクリーンアップ
backup_excel_file("path/to/file.xlsx")
cleanup_old_csv_files("processed_folder_path")
```

## 設定ファイル（config.ini）

```ini
[Appearance]
font_size = 11
window_width = 350
window_height = 330

[ExcludeDocs]
list = 訪問看護指示書,紹介状

[ExcludeDoctors]
list = 清水,菅原,寺井

[Paths]
downloads_path = C:\Users\...\Downloads
excel_path = C:/path/to/file.xlsm
backup_path = C:/path/to/backup
processed_path = C:\path\to\processed

[ButtonPosition]
share_button_x = 1450
share_button_y = 160
share_button_wait_time = 1
```

## 開発情報

### テスト実行

```bash
python -m pytest              # 全テストを実行
python -m pytest tests/test_csv_processor.py  # 特定のテストファイルを実行
```

### 実行ファイル生成

```bash
python build.py
```

このコマンドはバージョンを自動更新してPyInstallerで実行ファイルを生成し、`dist/` フォルダに出力します。

### バージョン管理

バージョン情報は `app/__init__.py` で管理され、ビルド時に自動更新されます：
```python
__version__ = "1.1.3"
__date__ = "2025-12-11"
```

### 主要な依存パッケージ

詳細は `requirements.txt` を参照：

- **PyQt6** (6.8.0): GUI フレームワーク
- **openpyxl** (3.1.5): Excel ファイル操作
- **polars** (1.17.1): CSV/データ処理
- **PyAutoGUI** (0.9.54): 自動化機能
- **pyinstaller** (6.11.1): 実行ファイル生成

## CSVファイル名規則

アプリケーションは以下のパターンのCSVファイルのみを処理します：

```
XXXX_YYYYMMDDhhmmss.csv

例: 0001_20250101120000.csv
```

## トラブルシューティング

| 問題 | 解決方法 |
|------|--------|
| **CSVファイルが見つからない** | ダウンロードフォルダパスを確認。ファイル名が `XXXX_YYYYMMDDhhmmss.csv` 形式と一致することを確認 |
| **エンコーディングエラー** | CSVを Shift-JIS または UTF-8 で再エクスポート。対応形式: Shift-JIS、UTF-8、CP932 |
| **Excelが開けない** | ファイルパスが正しいか、他のアプリで開かれていないか確認。ファイルが存在するか確認 |
| **座標設定が機能しない** | 「画面の座標表示」ツールで正確な座標を取得し、`config.ini` の `ButtonPosition` を更新 |
| **パフォーマンス低下** | 処理済みCSVとバックアップフォルダをクリーンアップ。古いファイルを削除 |

## ライセンス

このプロジェクトのライセンス条件については、`LICENSE` ファイルを参照してください。

## サポート

問題や機能リクエストについては、GitHubのIssuesセクションにお問い合わせください。
