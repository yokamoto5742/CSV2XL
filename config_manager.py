import configparser
import os
import sys
from pathlib import Path
from typing import Final, List

def get_config_path() -> Path:
    # 実行ファイルのディレクトリを取得
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(os.path.dirname(os.path.abspath(__file__)))
    return base_path / 'config.ini'

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

    def get_exclude_docs(self) -> List[str]:
        if 'ExcludeDocs' not in self.config:
            return []
        docs = self.config.get('ExcludeDocs', 'list', fallback='')
        return [doc.strip() for doc in docs.split(',') if doc.strip()]

    def get_downloads_path(self) -> str:
        if 'Paths' not in self.config:
            return str(Path.home() / "Downloads")
        return self.config.get('Paths', 'downloads_path', fallback=str(Path.home() / "Downloads"))

    def get_excel_path(self) -> str:
        if 'Paths' not in self.config:
            return r"C:\Shinseikai\CSV2XL\医療文書担当一覧.xlsm"
        return self.config.get('Paths', 'excel_path', fallback=r"C:\Shinseikai\CSV2XL\医療文書担当一覧.xlsm")

    def set_downloads_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['downloads_path'] = path
        self.save_config()

    def set_excel_path(self, path: str) -> None:
        if 'Paths' not in self.config:
            self.config['Paths'] = {}
        self.config['Paths']['excel_path'] = path
        self.save_config()

    def save_config(self) -> None:
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except (IOError, OSError) as e:
            raise OSError(f"Failed to load config: {e}") from e

    def _ensure_section(self, section: str) -> None:
        """設定セクションが存在することを確認し、必要に応じて作成する"""
        if section not in self.config:
            self.config[section] = {}
