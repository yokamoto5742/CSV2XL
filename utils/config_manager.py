import configparser
import os
import shutil
import sys
from pathlib import Path

DEFAULT_BASE_DIR = Path(r"C:\Shinseikai\CSV2XL")


def get_config_path() -> Path:
    if getattr(sys, 'frozen', False):
        # exe 実行時は exe と同じフォルダの config.ini を使う(無ければ同梱版をコピーする)
        config_path = Path(sys.executable).parent / 'config.ini'
        if not config_path.exists():
            bundled_path = Path(sys._MEIPASS) / 'config.ini'  # type: ignore[attr-defined]
            shutil.copy2(bundled_path, config_path)
        return config_path
    return Path(os.path.dirname(os.path.abspath(__file__))) / 'config.ini'


CONFIG_PATH = get_config_path()


class ConfigManager:
    def __init__(self, config_file: Path | str = CONFIG_PATH) -> None:
        self.config_file: Path = Path(config_file)
        self.config: configparser.ConfigParser = configparser.ConfigParser()
        self.load_config()

    def load_config(self) -> None:
        if not self.config_file.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_file}")

        try:
            self.config.read(self.config_file, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                content: str = self.config_file.read_bytes().decode('cp932')
                self.config.read_string(content)
            except (UnicodeDecodeError, OSError) as e:
                raise OSError(f"Failed to load config: {e}") from e

    def save_config(self) -> None:
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except (IOError, OSError) as e:
            raise OSError(f"Failed to save config: {e}") from e

    def _set(self, section: str, values: dict[str, str]) -> None:
        if section not in self.config:
            self.config[section] = {}
        self.config[section].update(values)
        self.save_config()

    def get_list(self, section: str) -> list[str]:
        """カンマ区切りの 'list' キーをリストとして取得"""
        items = self.config.get(section, 'list', fallback='')
        return [item.strip() for item in items.split(',') if item.strip()]

    def set_list(self, section: str, items: list[str]) -> None:
        self._set(section, {'list': ','.join(items)})

    def get_exclude_docs(self) -> list[str]:
        return self.get_list('ExcludeDocs')

    def get_exclude_doctors(self) -> list[str]:
        return self.get_list('ExcludeDoctors')

    def get_downloads_path(self) -> str:
        return self.config.get('Paths', 'downloads_path', fallback=str(Path.home() / "Downloads"))

    def set_downloads_path(self, path: str) -> None:
        self._set('Paths', {'downloads_path': path})

    def get_excel_path(self) -> str:
        return self.config.get('Paths', 'excel_path', fallback=str(DEFAULT_BASE_DIR / "医療文書担当一覧.xlsm"))

    def set_excel_path(self, path: str) -> None:
        self._set('Paths', {'excel_path': path})

    def get_backup_path(self) -> str:
        return self.config.get('Paths', 'backup_path', fallback=str(DEFAULT_BASE_DIR / "backup"))

    def set_backup_path(self, path: str) -> None:
        self._set('Paths', {'backup_path': path})

    def get_processed_path(self) -> str:
        return self.config.get('Paths', 'processed_path', fallback=str(DEFAULT_BASE_DIR / "processed"))

    def get_folder_dialog_size(self) -> tuple[int, int]:
        width = self.config.getint('DialogSize', 'folder_dialog_width', fallback=400)
        height = self.config.getint('DialogSize', 'folder_dialog_height', fallback=200)
        return width, height

    def get_font_size(self) -> int:
        return self.config.getint('Appearance', 'font_size', fallback=9)

    def set_font_size(self, size: int) -> None:
        self._set('Appearance', {'font_size': str(size)})

    def get_window_size(self) -> tuple[int, int]:
        width = self.config.getint('Appearance', 'window_width', fallback=300)
        height = self.config.getint('Appearance', 'window_height', fallback=200)
        return width, height

    def set_window_size(self, width: int, height: int) -> None:
        self._set('Appearance', {'window_width': str(width), 'window_height': str(height)})

    def get_share_button_position(self) -> tuple[int, int]:
        x = self.config.getint('ButtonPosition', 'share_button_x', fallback=1450)
        y = self.config.getint('ButtonPosition', 'share_button_y', fallback=180)
        return x, y

    def get_share_button_wait_time(self) -> int:
        return self.config.getint('ButtonPosition', 'share_button_wait_time', fallback=3)

    def get_backup_retention_days(self) -> int:
        """バックアップファイルの保持日数を取得"""
        return self.config.getint('Backup', 'retention_days', fallback=7)
