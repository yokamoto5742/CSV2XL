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
        assert config.get_backup_retention_days() == 7
        assert config.get_list('ExcludeDocs') == []
        assert config.get_backup_path() == r"C:\Shinseikai\CSV2XL\backup"
        assert config.get_database_path() == r"C:\Shinseikai\CSV2XL\医療文書担当一覧.db"

    def test_set_creates_section_and_saves(self, config: ConfigManager) -> None:
        """setter はセクションを作成してファイルに保存する"""
        config.set_window_size(400, 300)
        config.set_backup_path(r"D:\backup")
        config.set_database_path(r"\\server\share\test.db")

        reloaded = ConfigManager(config.config_file)
        assert reloaded.get_window_size() == (400, 300)
        assert reloaded.get_backup_path() == r"D:\backup"
        assert reloaded.get_database_path() == r"\\server\share\test.db"

    def test_get_list(self, config: ConfigManager) -> None:
        """カンマ区切りのリストを読み込む(空白や空要素は除く)"""
        config.config_file.write_text('[ExcludeDoctors]\nlist = 田中, 山田 ,,\n', encoding='utf-8')

        assert ConfigManager(config.config_file).get_list('ExcludeDoctors') == ['田中', '山田']

    def test_get_staff_names(self, config: ConfigManager) -> None:
        """[Staff] ordered_names を順番どおりに読む(CP932 の外字も化けない)"""
        config.config_file.write_text('[Staff]\nordered_names = 植田, 花﨑,,山本\n', encoding='utf-8')

        assert ConfigManager(config.config_file).get_staff_names() == ['植田', '花﨑', '山本']

    def test_appearance_defaults_and_save(self, config: ConfigManager) -> None:
        assert config.get_start_maximized() is True
        assert config.get_side_panel_width() == 200

        config.set_start_maximized(False)
        config.set_side_panel_width(320)

        reloaded = ConfigManager(config.config_file)
        assert reloaded.get_start_maximized() is False
        assert reloaded.get_side_panel_width() == 320


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
