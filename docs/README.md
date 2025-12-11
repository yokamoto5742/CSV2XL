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

**必須条件:**
- Windows11
- Python3.11以上

**推奨:**
- メモリ: 4GB以上
- ディスク: 空き容量100MB以上

## インストール

### 1. リポジトリをクローン
```bash
git clone <repository-url>
cd CSV2XL
```

### 2. 仮想環境を作成（推奨）
```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. 依存パッケージをインストール
```bash
pip install -r requirements.txt
```

### 4. アプリケーションの起動
```bash
python main.py
```

## 使い方

### アプリケーションの起動

```bash
python main.py
```

メインウィンドウが表示され、以下の機能が利用可能になります。

### 基本的なワークフロー

1. **CSVファイルを準備**: Papyrusからダウンロードフォルダにエクスポート
2. **CSVファイル取り込みボタンをクリック**: アプリが以下を自動実行
   - 最新のCSVファイルを検出
   - 指定されたエンコーディングで読み込み
   - データを処理・変換
   - Excelファイルに転記
   - バックアップを作成
   - 処理済みCSVを移動
3. **Excel ファイルが自動で開く**: データ確認と追加操作が可能

### 設定メニュー

#### 除外する文書名
特定の文書種別を転記対象外にします。
- 例: 訪問看護指示書、紹介状など

#### 除外する医師名
特定の医師を転記対象外にします。
- 例: 清水、菅原、寺井など

#### フォントとウィンドウサイズ
- **フォントサイズ**: アプリUIのフォントサイズ（既定値: 11）
- **ウィンドウサイズ**: 幅×高さを指定（既定値: 350×330）

#### 画面の座標表示
自動化機能で使用される座標を設定するツール。マウスカーソル位置のリアルタイム表示で正確な座標値を取得できます。

#### フォルダの場所
- **ダウンロードフォルダ**: CSVファイルのソースパス
- **Excelファイルパス**: データ転記先のExcelファイル
- **バックアップフォルダ**: バックアップの保存先
- **処理済みCSVフォルダ**: 処理済みファイルの移動先

## プロジェクト構成

```
CSV2XL/
├── app/                          # UIレイヤー
│   ├── __init__.py              # バージョン情報
│   ├── main_window.py           # メインウィンドウ（主要UI）
│   └── dialogs.py               # 設定ダイアログ
│
├── services/                    # ビジネスロジック
│   ├── csv_excel_transfer.py   # 転送処理の中核
│   ├── csv_processor.py        # CSV読込・処理
│   ├── excel_processor.py      # Excel操作
│   ├── file_manager.py         # ファイル・ディレクトリ管理
│   └── coordinate_tracker.py   # 座標トラッキングUI
│
├── utils/                       # ユーティリティ
│   ├── config_manager.py       # 設定管理
│   └── config.ini              # 設定ファイル
│
├── scripts/                     # ビルド・補助スクリプト
│   └── version_manager.py      # バージョン管理
│
├── tests/                       # テスト
│   └── test_*.py               # ユニットテスト
│
├── main.py                      # エントリーポイント
├── build.py                     # 実行ファイル生成
└── requirements.txt             # 依存パッケージ
```

## 主要機能の詳細

### CSVデータ処理

**読込エンコーディング自動判定**

```python
from services.csv_processor import read_csv_with_encoding

# エンコーディングを自動判定して読み込み
df = read_csv_with_encoding("path/to/file.csv")
```

CSV読込時に以下のエンコーディングを自動試行：
- Shift-JIS（日本語対応の標準エンコーディング）
- UTF-8
- CP932（Windows日本語）

**データ変換処理**

```python
from services.csv_processor import process_csv_data, convert_date_format

# CSV データを処理
df = process_csv_data(df)

# 日付形式を変換
df = convert_date_format(df)
```

### Excel 操作

**データの書き込みと重複チェック**

```python
from services.excel_processor import write_data_to_excel

# Excelに書き込み（重複は自動検出）
success = write_data_to_excel("path/to/file.xlsx", df)
```

**自動フォーマット適用**

- セル書式の自動適用
- データの自動ソート
- 既存フォーマットの保持

### ファイル管理

**バックアップと クリーンアップ**

```python
from services.file_manager import backup_excel_file, cleanup_old_csv_files

# Excel ファイルをバックアップ
backup_excel_file("path/to/file.xlsx")

# 古い処理済みCSVを削除
cleanup_old_csv_files("processed_folder_path")
```

## 設定ファイル（config.ini）

設定は `utils/config.ini` で管理されます。

```ini
[Appearance]
font_size = 11                    # フォントサイズ
window_width = 350                # ウィンドウ幅
window_height = 330               # ウィンドウ高さ

[ExcludeDocs]
list = 訪問看護指示書,紹介状    # 除外する文書名

[ExcludeDoctors]
list = 清水,菅原,寺井             # 除外する医師名

[Paths]
downloads_path = C:\Users\...\Downloads    # CSVソース
excel_path = C:/path/to/file.xlsm          # Excel 保存先
backup_path = C:/path/to/backup            # バックアップ先
processed_path = C:\path\to\processed      # 処理済みCSV先

[ButtonPosition]
share_button_x = 1450             # 自動化ボタンX座標
share_button_y = 160              # 自動化ボタンY座標
share_button_wait_time = 1        # 待機時間（秒）
```

## 開発情報

### テスト実行

```bash
# 全テストを実行
python -m pytest

# 特定のテストファイルを実行
python -m pytest tests/test_csv_processor.py
```

### 実行ファイル生成

```bash
# PyInstallerで実行ファイルを生成（自動的にバージョンを更新）
python build.py
```

このコマンドは以下を実行します：
- `scripts/version_manager.py` でバージョンを自動更新
- PyInstallerで実行ファイル生成
- `dist/` フォルダに実行ファイルを出力

### バージョン管理

バージョン情報は `app/__init__.py` で管理されます：

```python
__version__ = "1.1.2"
__date__ = "2025-12-10"
```

ビルド実行時に自動更新されます。

### 主要な依存パッケージ

- **PyQt6** (6.8.0): GUI フレームワーク
- **openpyxl** (3.1.5): Excel ファイル操作
- **polars** (1.17.1): 高速 CSV/データ処理
- **PyAutoGUI** (0.9.54): 自動化機能
- **pyinstaller** (6.11.1): 実行ファイル生成

詳細は `requirements.txt` を参照してください。

## CSVファイル名規則

アプリケーションは特定のパターンに一致したCSVファイルのみを処理します：

```
[3-4桁の数字]_[14桁の数字].csv

例: 0001_20250101120000.csv
```

別のパターンのファイルは処理対象外となります。

## トラブルシューティング

### CSVファイルが見つからない
- ダウンロードフォルダのパス設定を確認
- ファイル名が指定の形式に一致しているか確認
- `utils/config.ini` の `downloads_path` を確認

### Excelファイルを開けない
- ファイルパスが正しく設定されているか確認
- ファイルが他のアプリケーションで開かれていないか確認
- ファイルが存在するか確認

### エンコーディングエラー
- CSVファイルのエンコーディングを確認
- 対応エンコーディング: Shift-JIS、UTF-8、CP932
- 必要に応じてファイルを再エクスポート

### 座標設定が機能しない
- 「画面の座標表示」ツールで正確な座標を取得
- `config.ini` の `ButtonPosition` セクションを更新
- マウス位置とウィンドウの状態を確認

### パフォーマンスが低下
- 処理済みCSVフォルダの容量を確認
- バックアップフォルダの容量を確認
- 必要に応じて古いファイルを削除

## ライセンス

このプロジェクトのライセンス条件については、`LICENSE` ファイルを参照してください。

## サポート

問題や機能リクエストについては、GitHubのIssuesセクションにお問い合わせください。
