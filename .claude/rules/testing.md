---
description: テスト実行コマンドとテスト方針
---

## テスト実行コマンド

テストは pytest を uv 経由で実行する(`.venv` を直接指定しない)。

```bash
# 全件
uv run pytest tests/ -v --tb=short

# 単一ファイル
uv run pytest tests/test_<モジュール>.py -v

# 単一テスト
uv run pytest tests/test_<モジュール>.py::test_<テスト名> -v

# カバレッジ付き(pytest-cov を一時的に使う。依存には追加しない)
uv run --with pytest-cov pytest tests/ --cov --cov-report=term-missing
```

## テストの配置

- テストは `tests/` 直下にモジュールごとに置く(例: `services/foo.py` → `tests/test_foo.py`)。
- コードを変更したら、コミット前に全件を実行してパスを確認する。
