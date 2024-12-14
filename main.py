import sys
from app_window import MainWindow, QApplication
from service_csv_excel_transfer import transfer_csv_to_excel
from version import VERSION, LAST_UPDATED


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
