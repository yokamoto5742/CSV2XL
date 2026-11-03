import sys
from pathlib import Path

import pytest

from utils.config_manager import ConfigManager, get_config_path


@pytest.fixture
def config(tmp_path: Path) -> ConfigManager:
    config_file = tmp_path / 'config.ini'
    config_file.write_text('', encoding='utf-8')
    return ConfigManager(config_file)


class TestConfigManager:
    def test_defaults_without_sections(self, config: ConfigManager) -> None:
        """セクションが無くても既定値を返す"""
        assert config.get_font_size() == 9
        assert config.get_window_size() == (300, 200)
        assert config.get_share_button_position() == (1450, 180)
        assert config.get_backup_retention_days() == 7
        assert config.get_exclude_docs() == []
        assert config.get_backup_path() == r"C:\Shinseikai\CSV2XL\backup"

    def test_set_creates_section_and_saves(self, config: ConfigManager) -> None:
        """setter はセクションを作成してファイルに保存する"""
        config.set_window_size(400, 300)
        config.set_backup_path(r"D:\backup")

        reloaded = ConfigManager(config.config_file)
        assert reloaded.get_window_size() == (400, 300)
        assert reloaded.get_backup_path() == r"D:\backup"

    def test_list_round_trip(self, config: ConfigManager) -> None:
        """リストの保存と読み込み(空白や空要素は除く)"""
        config.set_list('ExcludeDoctors', ['田中', ' 山田 ', ''])

        reloaded = ConfigManager(config.config_file)
        assert reloaded.get_exclude_doctors() == ['田中', '山田']


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
