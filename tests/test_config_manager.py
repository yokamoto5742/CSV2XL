import sys
from pathlib import Path

import pytest

from utils.config_manager import get_config_path


class TestGetConfigPath:
    def test_not_frozen(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """通常実行時は utils/config.ini を使う"""
        monkeypatch.delattr(sys, 'frozen', raising=False)
        expected = Path(__file__).resolve().parent.parent / 'utils' / 'config.ini'
        assert get_config_path() == expected

    def test_frozen_copies_bundled_config(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """exe 実行時は exe と同じフォルダに同梱版をコピーして使う"""
        bundle_dir = tmp_path / '_internal'
        bundle_dir.mkdir()
        (bundle_dir / 'config.ini').write_text('[Appearance]\nfont_size = 9\n', encoding='utf-8')
        monkeypatch.setattr(sys, 'frozen', True, raising=False)
        monkeypatch.setattr(sys, '_MEIPASS', str(bundle_dir), raising=False)
        monkeypatch.setattr(sys, 'executable', str(tmp_path / 'CSV2XL.exe'))

        config_path = get_config_path()

        assert config_path == tmp_path / 'config.ini'
        assert config_path.read_text(encoding='utf-8') == '[Appearance]\nfont_size = 9\n'

    def test_frozen_keeps_existing_config(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """exe の隣に config.ini があれば同梱版で上書きしない"""
        bundle_dir = tmp_path / '_internal'
        bundle_dir.mkdir()
        (bundle_dir / 'config.ini').write_text('bundled', encoding='utf-8')
        (tmp_path / 'config.ini').write_text('user', encoding='utf-8')
        monkeypatch.setattr(sys, 'frozen', True, raising=False)
        monkeypatch.setattr(sys, '_MEIPASS', str(bundle_dir), raising=False)
        monkeypatch.setattr(sys, 'executable', str(tmp_path / 'CSV2XL.exe'))

        assert get_config_path().read_text(encoding='utf-8') == 'user'
