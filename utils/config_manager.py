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


def split_comma_list(text: str) -> list[str]:
    """カンマ区切りの文字列をリストにする(空白や空要素は除く)"""
    return [item.strip() for item in text.split(',') if item.strip()]


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

    def get_downloads_path(self) -> str:
        return self.config.get('Paths', 'downloads_path', fallback=str(Path.home() / "Downloads"))

    def set_downloads_path(self, path: str) -> None:
        self._set('Paths', {'downloads_path': path})

    def get_database_path(self) -> str:
        return self.config.get('Paths', 'database_path', fallback=str(DEFAULT_BASE_DIR / "医療文書担当一覧.db"))

    def set_database_path(self, path: str) -> None:
        self._set('Paths', {'database_path': path})

    def get_backup_path(self) -> str:
        return self.config.get('Paths', 'backup_path', fallback=str(DEFAULT_BASE_DIR / "backup"))

    def set_backup_path(self, path: str) -> None:
        self._set('Paths', {'backup_path': path})

    def get_xlsm_path(self) -> str:
        """医療文書担当一覧_YYYYMMDDhhmm.xlsm を置くフォルダ"""
        return self.config.get('Paths', 'xlsm_path', fallback=str(DEFAULT_BASE_DIR))

    def set_xlsm_path(self, path: str) -> None:
        self._set('Paths', {'xlsm_path': path})

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

    def get_column_widths(self) -> list[int]:
        """担当一覧の列幅(ピクセル)。保存されていなければ空のリスト"""
        return [int(width) for width in split_comma_list(self.config.get('Appearance', 'column_widths', fallback=''))]

    def set_column_widths(self, widths: list[int]) -> None:
        self._set('Appearance', {'column_widths': ','.join(str(width) for width in widths)})

    def get_window_size(self) -> tuple[int, int]:
        width = self.config.getint('Appearance', 'window_width', fallback=300)
        height = self.config.getint('Appearance', 'window_height', fallback=200)
        return width, height

    def get_window_position(self) -> tuple[int, int] | None:
        """前回終了時のウィンドウの位置。保存されていなければ None"""
        if not (self.config.has_option('Appearance', 'window_x') and self.config.has_option('Appearance', 'window_y')):
            return None
        return self.config.getint('Appearance', 'window_x'), self.config.getint('Appearance', 'window_y')

    def get_start_maximized(self) -> bool:
        """前回終了時に最大化していたか"""
        return self.config.getboolean('Appearance', 'start_maximized', fallback=True)

    def set_window_geometry(self, x: int, y: int, width: int, height: int, maximized: bool) -> None:
        """終了時のウィンドウの位置・大きさ(最大化を解除したときのもの)と最大化の状態を保存する"""
        self._set('Appearance', {
            'window_x': str(x), 'window_y': str(y), 'window_width': str(width), 'window_height': str(height),
            'start_maximized': str(maximized),
        })

    def get_side_panel_width(self) -> int:
        return self.config.getint('Appearance', 'side_panel_width', fallback=200)

    def set_side_panel_width(self, width: int) -> None:
        self._set('Appearance', {'side_panel_width': str(width)})

    def get_staff_names(self) -> list[str]:
        """担当者名を一括登録するためのカンマ区切りのリスト"""
        return split_comma_list(self.config.get('Staff', 'ordered_names', fallback=''))

    def get_exclude_docs(self) -> list[str]:
        """除外する文書名を一括登録するためのカンマ区切りのリスト"""
        return split_comma_list(self.config.get('ExcludeDocs', 'list', fallback=''))

    def get_exclude_doctors(self) -> list[str]:
        """除外する医師名を一括登録するためのカンマ区切りのリスト"""
        return split_comma_list(self.config.get('ExcludeDoctors', 'list', fallback=''))

    def get_backup_retention_days(self) -> int:
        """バックアップファイルの保持日数を取得"""
        return self.config.getint('Backup', 'retention_days', fallback=7)
