# チェンジログ

このプロジェクトの全変更は、このファイルに記録されます。

フォーマットは[Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)に基づいており、
バージョン管理は[Semantic Versioning](https://semver.org/lang/ja/)に従っています。

## [未リリース (Unreleased)]

### 追加

- **Python型チェックツール**: CoverageおよびPython型チェック用ツールを追加（requirements.txt更新）

### 変更

- **pyrightconfig.json**: servicesディレクトリを型チェック対象に含める設定を更新
- **excel_processor テスト**: pathlib.Pathを使用した実装に変更
- **app/__init__.py**: バージョン情報と日付を更新

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

[未リリース (Unreleased)]: https://github.com/yokamura/CSV2XL/compare/v1.1.2...HEAD
[1.1.2]: https://github.com/yokamura/CSV2XL/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/yokamura/CSV2XL/releases/tag/v1.1.1
