# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## House Rules:
- 文章ではなくパッチの差分を返すこと。Return patch diffs, not prose.
- 不明な点がある場合は、トレードオフを明記した2つの選択肢を提案すること（80語以内）。
- 変更範囲は最小限に抑えること：範囲拡大の指示がない限り最大3ファイルまで。
- 3ファイルを超える場合は、理由と対象ファイルを明記すること。
- Pythonコードのimport文は以下の適切な順序に並べ替えてください。
標準ライブラリ
サードパーティライブラリ
カスタムモジュール 
それぞれアルファベット順に並べます。importが先でfromは後です。

## Automatic Notifications (Hooks)
自動通知は`.claude/settings.local.json` で設定済：

- **Stop Hook**: ユーザーがClaude Codeを停止した時に「作業が完了しました」と通知
- **SessionEnd Hook**: セッション終了時に「Claude Code セッションが終了しました」と通知

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