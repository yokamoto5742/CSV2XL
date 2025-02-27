import os
from PyQt6.QtWidgets import QMessageBox
import traceback
from csv_processor import CSVProcessor
from excel_processor import ExcelProcessor
from file_manager import FileManager
from config_manager import ConfigManager

def transfer_csv_to_excel():
    """CSVデータをExcelに転送するメイン関数"""
    try:
        config = ConfigManager()
        downloads_path = config.get_downloads_path()
        excel_path = config.get_excel_path()

        # ファイル検索
        file_manager = FileManager(config)
        latest_csv = file_manager.get_csv_files(downloads_path)
        if not latest_csv:
            QMessageBox.warning(None, "警告", "ダウンロードフォルダにCSVファイルが見つかりません。")
            return

        print(f"処理するCSVファイル: {latest_csv}")

        # CSVデータの処理
        csv_processor = CSVProcessor()
        df = csv_processor.read_csv_with_encoding(latest_csv)
        df = csv_processor.convert_date_format(df)
        df = csv_processor.process_csv_data(df, config)

        # Excelデータの処理
        excel_processor = ExcelProcessor()
        result = excel_processor.process_excel_file(excel_path, df)
        if not result:
            return

        # 処理済みCSVファイルの移動
        file_manager.process_completed_csv(latest_csv)

        # Excelファイルを開いて操作
        excel_processor.open_and_process_excel(excel_path, config)

    except Exception as e:
        QMessageBox.critical(None, "エラー", f"処理中にエラーが発生しました: {str(e)}")
        print(f"エラーが発生しました: {str(e)}")
        print("詳細なエラー情報:")
        print(traceback.format_exc())
