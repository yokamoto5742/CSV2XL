import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QShortcut, QKeySequence
import pyautogui


class CoordinateTracker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("画面の座標表示")
        self.setFixedSize(300, 150)

        # メインウィジェット
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 座標表示ラベル
        self.coord_label = QLabel("座標: (0, 0)")
        self.coord_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.coord_label)

        # 説明ラベル
        info_label = QLabel("Escキーで終了\nSpaceキーで座標をコピー")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info_label)

        # タイマーで座標更新
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_coordinates)
        self.timer.start(100)  # 100ミリ秒ごとに更新

        # ショートカットキー
        QShortcut(QKeySequence(Qt.Key.Key_Space), self).activated.connect(self.copy_coordinates)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self).activated.connect(self.close)

    def update_coordinates(self):
        x, y = pyautogui.position()
        self.coord_label.setText(f"座標: ({x}, {y})")

    def copy_coordinates(self):
        x, y = pyautogui.position()
        QApplication.clipboard().setText(f"{x}, {y}")


def main():
    app = QApplication(sys.argv)
    tracker = CoordinateTracker()
    tracker.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()