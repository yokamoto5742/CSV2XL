import traceback
from PyQt6.QtWidgets import QMessageBox
from config_manager import ConfigManager
from service_csv_processor import find_latest_csv, read_csv_with_encoding, process_csv_data, convert_date_format, process_completed_csv
from service_excel_processor import write_data_to_excel, open_and_sort_excel
from service_file_manager import backup_excel_file, cleanup_old_files, ensure_directories_exist
from pathlib import Path


def transfer_csv_to_excel():
    try:
        # 初期化と設定の取得
        config = ConfigManager()
        downloads_path = config.get_downloads_path()
        excel_path = config.get_excel_path()
        processed_dir = Path(config.get_processed_path())
        
        # 必要なディレクトリが存在することを確認
        ensure_directories_exist()
        
        # 古いファイルのクリーンアップ
        cleanup_old_files(processed_dir)

        # CSVファイルの検索
        latest_csv = find_latest_csv(downloads_path)
        if not latest_csv:
            QMessageBox.warning(None, "警告", "ダウンロードフォルダにCSVファイルが見つかりません。")
            return

        print(f"処理するCSVファイル: {latest_csv}")

        # CSVファイルの読み込みと処理
        df = read_csv_with_encoding(latest_csv)
        df = convert_date_format(df)
        df = process_csv_data(df)

        # Excelファイルへの書き込み
        if write_data_to_excel(excel_path, df):
            # バックアップの作成
            backup_excel_file(excel_path)
            
            # 処理済みCSVファイルの移動
            process_completed_csv(latest_csv)
            
            # Excelファイルを開いてソートし、共有ボタンをクリック
            open_and_sort_excel(excel_path)
            
            QMessageBox.information(None, "完了", "CSVデータをExcelに転送しました。")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        print("詳細なエラー情報:")
        print(traceback.format_exc())
        QMessageBox.critical(None, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")


if __name__ == "__main__":
    # 単体テスト用
    transfer_csv_to_excel()
