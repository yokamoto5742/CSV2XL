import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QDialog, QLineEdit,
    QListWidget, QDialogButtonBox, QFileDialog
)
from PyQt6.QtCore import Qt
from config_manager import ConfigManager
from main import transfer_csv_to_excel


class ExcludeDocsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("除外する文書名の設定")
        self.setModal(True)

        layout = QVBoxLayout()

        # 文書名入力フィールド
        self.input_field = QLineEdit()
        layout.addWidget(QLabel("除外する文書名:"))
        layout.addWidget(self.input_field)

        # 登録済み文書名リスト
        self.doc_list = QListWidget()
        layout.addWidget(QLabel("登録済み文書名:"))
        layout.addWidget(self.doc_list)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # 追加ボタン
        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_document)

        # 削除ボタン
        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_documents()

    def load_documents(self):
        if 'ExcludeDocs' in self.config.config:
            docs = self.config.config['ExcludeDocs'].get('list', '').split(',')
            for doc in docs:
                if doc.strip():
                    self.doc_list.addItem(doc.strip())

    def add_document(self):
        doc_name = self.input_field.text().strip()
        if doc_name:
            self.doc_list.addItem(doc_name)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.doc_list.currentItem()
        if current_item:
            self.doc_list.takeItem(self.doc_list.row(current_item))

    def accept(self):
        # 設定の保存
        docs = []
        for i in range(self.doc_list.count()):
            docs.append(self.doc_list.item(i).text())

        if 'ExcludeDocs' not in self.config.config:
            self.config.config['ExcludeDocs'] = {}
        self.config.config['ExcludeDocs']['list'] = ','.join(docs)
        self.config.save_config()
        super().accept()


class FolderPathDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("フォルダパスの設定")
        self.setModal(True)

        layout = QVBoxLayout()

        # パス入力フィールド
        self.input_field = QLineEdit()
        layout.addWidget(QLabel("フォルダパス:"))
        layout.addWidget(self.input_field)

        # 参照ボタン
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.browse_folder)
        layout.addWidget(browse_button)

        # パスリスト
        self.path_list = QListWidget()
        layout.addWidget(QLabel("登録済みパス:"))
        layout.addWidget(self.path_list)

        # ボタン
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # 追加ボタン
        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_path)

        # 削除ボタン
        delete_button = QPushButton("削除")
        delete_button.clicked.connect(self.delete_selected)

        layout.addWidget(add_button)
        layout.addWidget(delete_button)
        layout.addWidget(buttons)

        self.setLayout(layout)

        # 設定の読み込み
        self.config = ConfigManager()
        self.load_paths()

    def load_paths(self):
        directories = self.config.get_directories()
        for directory in directories:
            if directory.strip():
                self.path_list.addItem(directory.strip())

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "フォルダの選択")
        if folder:
            self.input_field.setText(folder)

    def add_path(self):
        path = self.input_field.text().strip()
        if path:
            self.path_list.addItem(path)
            self.input_field.clear()

    def delete_selected(self):
        current_item = self.path_list.currentItem()
        if current_item:
            self.path_list.takeItem(self.path_list.row(current_item))

    def accept(self):
        # 設定の保存
        paths = []
        for i in range(self.path_list.count()):
            paths.append(self.path_list.item(i).text())
        self.config.set_directories(paths)
        super().accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV2XL v1.0.0")
        self.setFixedSize(300, 200)

        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # レイアウト
        layout = QVBoxLayout()

        # タイトル
        title_label = QLabel("Papyrus書類受付リスト")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # CSVファイル取り込みボタン
        csv_button = QPushButton("CSVファイル取り込み")
        csv_button.clicked.connect(self.import_csv)
        layout.addWidget(csv_button)

        # 設定セクションのラベル
        settings_label = QLabel("設定")
        layout.addWidget(settings_label)

        # 除外する文書名ボタン
        exclude_docs_button = QPushButton("除外する文書名")
        exclude_docs_button.clicked.connect(self.show_exclude_docs_dialog)
        layout.addWidget(exclude_docs_button)

        # フォルダパスボタン
        folder_path_button = QPushButton("フォルダパス")
        folder_path_button.clicked.connect(self.show_folder_path_dialog)
        layout.addWidget(folder_path_button)

        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        main_widget.setLayout(layout)

    def import_csv(self):
        try:
            transfer_csv_to_excel()
        except Exception as e:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "エラー", f"CSVファイルの取り込み中にエラーが発生しました:\n{str(e)}")

    def show_exclude_docs_dialog(self):
        dialog = ExcludeDocsDialog(self)
        dialog.exec()

    def show_folder_path_dialog(self):
        dialog = FolderPathDialog(self)
        dialog.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
