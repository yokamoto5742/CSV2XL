# CSV2XL - 医療文書管理アプリケーション

医療文書管理システム「Papyrus」からのCSVデータを、PyQt6ベースのGUIで効率的にExcelファイルに自動転記するデスクトップアプリケーションです。

## 主な機能

- 最新のCSVファイルを自動検出・処理
- 複数エンコーディング対応（Shift-JIS、UTF-8、CP932）
- CSV内の日付を自動変換
- 除外する文書名・医師名をカスタマイズ可能
- Excelへの転記時に重複を自動検出
- 処理前に自動バックアップを作成
- UIのフォント・ウィンドウサイズをカスタマイズ
- 自動化機能の座標設定をサポート

## 前提条件

- **OS**: Windows 11
- **Python**: 3.11 以上
- **メモリ**: 4GB 以上（推奨）

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

### アプリケーション起動
```bash
python main.py
```

### 基本的なワークフロー

1. PapyrusからダウンロードフォルダにCSVファイルをエクスポート
2. **CSVファイル取り込み**ボタンをクリック
3. 以下が自動で実行されます：
   - 最新のCSVを検出・読み込み
   - データを処理・変換
   - Excelに転記
   - バックアップを作成
   - 処理済みCSVを指定フォルダに移動
4. Excelファイルが自動で開きます

### 設定項目

**フィルタリング**:
- 除外する文書名（カンマ区切り）
- 除外する医師名（カンマ区切り）

**UI設定**:
- フォントサイズ（既定値：11）
- ウィンドウサイズ（既定値：350×330）

**ファイル・フォルダパス**:
- ダウンロードフォルダ（CSVファイルの検索元）
- Excelファイルパス（データ転記先）
- バックアップフォルダ
- 処理済みCSVフォルダ

**その他**:
- 画面の座標表示：マウス座標をリアルタイム表示し、自動化機能の座標を取得

## プロジェクト構造

```
CSV2XL/
├── app/                       # UI レイヤー
│   ├── __init__.py           # バージョン情報
│   ├── main_window.py        # メインウィンドウ
│   └── dialogs.py            # 設定ダイアログ
├── services/                 # ビジネスロジック
│   ├── csv_excel_transfer.py # CSVからExcelへの転送処理
│   ├── csv_processor.py      # CSV読込・エンコーディング判定
│   ├── excel_processor.py    # Excel書込・ソート・フォーマット
│   ├── file_manager.py       # バックアップ・クリーンアップ
│   └── coordinate_tracker.py # 座標トラッキング機能
├── utils/                    # ユーティリティ
│   ├── config_manager.py     # 設定ファイル管理
│   └── config.ini            # 設定ファイル
├── scripts/                  # ビルド・補助スクリプト
│   └── version_manager.py    # バージョン自動更新
├── tests/                    # ユニットテスト
├── main.py                   # エントリーポイント
├── build.py                  # 実行ファイル生成スクリプト
└── requirements.txt          # Python依存パッケージ
```

## 主要機能の詳細

### CSVデータ処理

複数のエンコーディング対応により、様々な形式のCSVファイルを自動判定して読み込みます：

```python
from services.csv_processor import read_csv_with_encoding, process_csv_data, convert_date_format

# エンコーディング自動判定（Shift-JIS → CP932 → UTF-8の順で試行）
df = read_csv_with_encoding("path/to/file.csv")
# 列名の一意化、スペース・特殊文字の除去
df = process_csv_data(df)
# 日付形式を自動変換
df = convert_date_format(df)
```

### Excel操作とファイル管理

```python
from services.excel_processor import write_data_to_excel, open_and_sort_excel
from services.file_manager import backup_excel_file, cleanup_old_csv_files

# Excelに書込（重複自動検出・セル書式自動適用）
success = write_data_to_excel("path/to/file.xlsx", df)

# バックアップ作成と古いファイル削除
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
list = 田中,山田,中村

[Paths]
downloads_path = C:\Users\...\Downloads
excel_path = C:\path\to\file.xlsm
backup_path = C:\path\to\backup
processed_path = C:\path\to\processed

[FileRetention]
backup_retention_days = 14

[ButtonPosition]
share_button_x = 1450
share_button_y = 160
share_button_wait_time = 1
```

- **Appearance**: UI外観設定（フォントサイズ、ウィンドウサイズ）
- **ExcludeDocs/ExcludeDoctors**: フィルタリング対象
- **Paths**: ファイル・フォルダパス
- **FileRetention**: バックアップ・処理済みCSVの保持期間（日数）
- **ButtonPosition**: 自動化機能の座標設定

## 開発情報

### テスト実行

```bash
python -m pytest                        # 全テストを実行
python -m pytest tests/test_csv_processor.py  # 特定テストを実行
python -m pytest --cov                 # カバレッジ付きで実行
```

### 実行ファイル生成

```bash
python build.py
```

バージョンを自動更新し、PyInstallerで実行ファイルを生成します。出力先は `dist/` フォルダです。

### バージョン管理

`app/__init__.py` で管理し、ビルド時に自動更新：
```python
__version__ = "1.1.3"
__date__ = "2025-12-11"
```

### 主要な依存パッケージ

詳細は `requirements.txt` を参照：

- **PyQt6** (6.8.0): GUIフレームワーク
- **openpyxl** (3.1.5): Excel操作
- **polars** (1.17.1): CSV/データ処理
- **PyAutoGUI** (0.9.54): UI自動化
- **pyinstaller** (6.11.1): 実行ファイル生成

## CSVファイル名規則

アプリケーションは以下のパターンのCSVファイルのみを処理します：

```
XXXX_YYYYMMDDhhmmss.csv
例: 0001_20250101120000.csv
```

最新のファイルが自動で検出されます。

## トラブルシューティング

**CSVファイルが見つからない**
- ダウンロードフォルダパスが `config.ini` で正しく設定されているか確認
- ファイル名が `XXXX_YYYYMMDDhhmmss.csv` 形式と一致するか確認

**エンコーディングエラー**
- CSVを Shift-JIS または UTF-8 で再度エクスポート
- 対応エンコーディング: Shift-JIS、CP932、UTF-8（この順で試行）

**Excelが開けない / データが転記されない**
- ファイルパスが正しいか、他のアプリで開かれていないか確認
- ファイルが読み取り可能な状態か確認
- Excelファイルのフォーマットが .xlsm であることを確認

**座標設定が機能しない**
- UI設定メニューの「画面の座標表示」ツールで正確な座標を取得
- `config.ini` の `[ButtonPosition]` セクションを更新

**パフォーマンス低下**
- バックアップフォルダと処理済みCSVフォルダをクリーンアップ
- 保持期間は `config.ini` の `backup_retention_days` で設定可能（デフォルト：14日）

## ライセンス

このプロジェクトのライセンス条件については、`LICENSE` ファイルを参照してください。
