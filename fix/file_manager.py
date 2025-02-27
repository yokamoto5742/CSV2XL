# file_manager.py - ファイル操作に関する機能
import os
from pathlib import Path
import shutil
import datetime


class FileManager:
    """ファイル操作を行うクラス"""
    
    def __init__(self, config):
        self.config = config
    
    def get_csv_files(self, downloads_path):
        """処理対象のCSVファイルを取得"""
        csv_files = [
            f for f in os.listdir(downloads_path)
            if f.endswith('.csv') and
               len(f.split('_')) == 2 and
               (3 <= len(f.split('_')[0]) <= 4) and
               len(f.split('_')[1].split('.')[0]) == 14
        ]
        
        if not csv_files:
            return None
            
        return max([os.path.join(downloads_path, f) for f in csv_files], key=os.path.getmtime)
    
    def backup_excel_file(self, excel_path):
        """Excelファイルのバックアップを作成"""
        if self.config:
            backup_dir = Path(self.config.get_backup_path())
        else:
            # configが利用できない場合は現在の日付をバックアップディレクトリとして使用
            backup_dir = Path(os.path.dirname(excel_path)) / "backup" / datetime.datetime.now().strftime("%Y%m%d")
            
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        backup_path = backup_dir / f"backup_{Path(excel_path).name}"

        try:
            shutil.copy2(excel_path, backup_path)
            print(f"バックアップを作成しました: {backup_path}")
        except Exception as e:
            print(f"バックアップ作成中にエラーが発生しました: {str(e)}")
            raise
    
    def process_completed_csv(self, csv_path):
        """処理済みCSVファイルを移動して古いファイルをクリーンアップ"""
        try:
            csv_file = Path(csv_path)
            if not csv_file.exists():
                print(f"処理対象のCSVファイルが見つかりません: {csv_path}")
                return

            processed_dir = Path(self.config.get_processed_path())
            processed_dir.mkdir(exist_ok=True, parents=True)

            new_path = processed_dir / csv_file.name
            shutil.move(str(csv_file), str(new_path))
            print(f"CSVファイルを移動しました: {new_path}")
            self._cleanup_old_files(processed_dir)

        except Exception as e:
            print(f"CSVファイルの処理中にエラーが発生しました: {str(e)}")
            raise
    
    def _cleanup_old_files(self, processed_dir):
        """3日以上経過した古いファイルを削除"""
        current_time = datetime.datetime.now()
        for file in processed_dir.glob("*.csv"):
            file_time = datetime.datetime.fromtimestamp(file.stat().st_mtime)
            if (current_time - file_time).days >= 3:
                try:
                    file.unlink()
                    print(f"古いファイルを削除しました: {file}")
                except Exception as e:
                    print(f"ファイル削除中にエラーが発生しました: {file} - {str(e)}")
