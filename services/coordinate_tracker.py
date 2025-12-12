import pyautogui
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QShortcut, QKeySequence
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel


class CoordinateTracker(QMainWindow):
    """マウス座標を表示するトラッキングウィンドウ"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("画面の座標")
        self.setFixedSize(250, 100)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.coord_label = QLabel("座標: (0, 0)")
        self.coord_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.coord_label)

        info_label = QLabel("Spaceキーで座標をコピー")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(info_label)

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_coordinates)
        self.timer.start(100)  # 100ミリ秒ごとに座標を更新

        QShortcut(QKeySequence(Qt.Key.Key_Space), self).activated.connect(self.copy_coordinates)

    def update_coordinates(self):
        """現在のマウス座標を取得して表示を更新"""
        x, y = pyautogui.position()
        self.coord_label.setText(f"座標: ({x}, {y})")

    @staticmethod
    def copy_coordinates():
        """現在のマウス座標をクリップボードにコピー"""
        x, y = pyautogui.position()
        clipboard = QApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(f"{x}, {y}")
