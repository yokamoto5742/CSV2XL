import pyautogui
import tkinter as tk
from tkinter import ttk
import keyboard


class CoordinateTracker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("座標トラッカー")
        self.root.geometry("300x150")

        # メインフレーム
        self.frame = ttk.Frame(self.root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 座標表示ラベル
        self.coord_var = tk.StringVar(value="座標: (0, 0)")
        self.coord_label = ttk.Label(self.frame, textvariable=self.coord_var)
        self.coord_label.grid(row=0, column=0, pady=10)

        # 説明ラベル
        self.info_label = ttk.Label(
            self.frame,
            text="Escキーで終了\nSpaceキーで座標をコピー"
        )
        self.info_label.grid(row=1, column=0, pady=10)

        # 更新処理
        self.update_coordinates()

        # キーバインド
        keyboard.on_press_key("space", self.copy_coordinates)
        keyboard.on_press_key("esc", self.quit_app)

    def update_coordinates(self):
        x, y = pyautogui.position()
        self.coord_var.set(f"座標: ({x}, {y})")
        self.root.after(100, self.update_coordinates)

    def copy_coordinates(self, _):
        x, y = pyautogui.position()
        self.root.clipboard_clear()
        self.root.clipboard_append(f"{x}, {y}")

    def quit_app(self, _):
        self.root.quit()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    tracker = CoordinateTracker()
    tracker.run()
