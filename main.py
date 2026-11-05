import sys

from PyQt6.QtWidgets import QApplication

from app.main_window import MainWindow


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    if window.config.get_start_maximized():
        window.showMaximized()
    else:
        window.show()
    sys.exit(app.exec())
