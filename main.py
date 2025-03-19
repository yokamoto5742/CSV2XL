import sys

from ui.main_window import MainWindow, QApplication
from utils.version import VERSION, LAST_UPDATED


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
