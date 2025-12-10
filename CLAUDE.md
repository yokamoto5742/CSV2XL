# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## House Rules:
- 文章ではなくパッチの差分を返す。
- コードの変更範囲は最小限に抑える。
- コードの修正は直接適用する。
- Pythonのコーディング規約はPEP8に従います。
- KISSの原則に従い、できるだけシンプルなコードにします。
- 可読性を優先します。一度読んだだけで理解できるコードが最高のコードです。
- Pythonのコードのimport文は以下の適切な順序に並べ替えてください。
標準ライブラリ
サードパーティライブラリ
カスタムモジュール 
それぞれアルファベット順に並べます。importが先でfromは後です。

## CHANGELOG
このプロジェクトにおけるすべての重要な変更は日本語でdcos/CHANGELOG.mdに記録します。
フォーマットは[Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)に基づきます。

## Automatic Notifications (Hooks)
自動通知は`.claude/settings.local.json` で設定済：
- **Stop Hook**: ユーザーがClaude Codeを停止した時に「作業が完了しました」と通知
- **SessionEnd Hook**: セッション終了時に「Claude Code セッションが終了しました」と通知

## クリーンコードガイドライン
- 関数のサイズ：関数は50行以下に抑えることを目標にしてください。関数の処理が多すぎる場合は、より小さなヘルパー関数に分割してください。
- 単一責任：各関数とモジュールには明確な目的が1つあるようにします。無関係なロジックをまとめないでください。
- 命名：説明的な名前を使用してください。`tmp` 、`data`、`handleStuff`のような一般的な名前は避けてください。例えば、`doCalc`よりも`calculateInvoiceTotal` の方が適しています。
- DRY原則：コードを重複させないでください。類似のロジックが2箇所に存在する場合は、共有関数にリファクタリングしてください。それぞれに独自の実装が必要な場合はその理由を明確にしてください。
- コメント:分かりにくいロジックについては説明を加えます。説明不要のコードには過剰なコメントはつけないでください。
- コメントとdocstringは必要最小限に日本語で記述します。文末に"。"や"."をつけないでください。

## Commands

### Running the Application
```bash
python main.py
```

### Testing
```bash
python -m pytest                    # Run all tests
python -m pytest tests/test_*.py    # Run specific test file
```

### Building Executable
```bash
python build.py                     # Builds executable using PyInstaller with version bump
```

## Architecture

This is a PyQt6-based CSV-to-Excel transfer application with Japanese interface text. The architecture follows a layered pattern:

**Application Layer (`app/`)**:
- `main_window.py`: Main GUI window with CSV import functionality
- `dialogs.py`: Configuration dialogs for exclusions, appearance, and folder paths

**Services Layer (`services/`)**:
- `csv_excel_transfer.py`: Main orchestration service that coordinates the CSV-to-Excel workflow
- `csv_processor.py`: CSV file reading, encoding detection, and data processing
- `excel_processor.py`: Excel file operations using openpyxl
- `file_manager.py`: File backup, cleanup, and directory management
- `coordinate_tracker.py`: Screen coordinate tracking for UI automation

**Utils Layer (`utils/`)**:
- `config_manager.py`: Configuration management using INI files

**Version Management**:
- Version info stored in `app/__init__.py` 
- `scripts/version_manager.py` handles automatic version bumping
- Build process automatically updates version before creating executable

**Key Dependencies**:
- PyQt6 for GUI framework
- openpyxl for Excel file manipulation
- polars for CSV processing
- pyinstaller for executable building

**Data Flow**:
1. User triggers CSV import from main window
2. `csv_excel_transfer.py` orchestrates the process
3. Latest CSV file is found and processed with encoding detection
4. Data is transformed and written to Excel file using configured templates
5. Original CSV is moved to processed folder after successful transfer