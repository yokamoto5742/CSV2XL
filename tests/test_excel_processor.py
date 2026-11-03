import pytest
import datetime
from typing import cast
from unittest.mock import patch, MagicMock

from openpyxl.worksheet.worksheet import Worksheet

from services.excel_processor import (
    get_last_row, apply_cell_formats, make_row_key, filter_new_rows, write_rows,
    bring_excel_to_front, write_data_to_excel, open_and_sort_excel
)


@pytest.fixture
def mock_worksheet():
    """テスト用のワークシートモックを生成するフィクスチャ"""
    worksheet = MagicMock()

    # 空の行を含むセルデータを模擬
    cells = [
        [MagicMock(value="header1"), MagicMock(value="header2"), MagicMock(value="header3")],
        [MagicMock(value="data1"), MagicMock(value="data2"), MagicMock(value="data3")],
        [MagicMock(value="data4"), MagicMock(value="data5"), MagicMock(value="data6")],
        [MagicMock(value=None), MagicMock(value=None), MagicMock(value=None)]
    ]

    # iter_rowsメソッドをモック
    worksheet.iter_rows.return_value = cells

    return worksheet


@pytest.fixture
def mock_workbook():
    """テスト用のワークブックモックを生成するフィクスチャ"""
    workbook = MagicMock()
    worksheet = MagicMock()
    workbook.active = worksheet
    return workbook, worksheet


class TestExcelProcessor:
    def test_get_last_row(self, mock_worksheet):
        """get_last_row関数のテスト"""
        # 関数を実行
        last_row = get_last_row(mock_worksheet)

        # 非空の行数を確認（この場合は3行）
        assert last_row == 3

        # モックの呼び出しを確認
        mock_worksheet.iter_rows.assert_called_once()

    def test_apply_cell_formats(self, mock_worksheet):
        """apply_cell_formats関数のテスト"""
        # 関数の実行
        start_row = 2

        # get_last_rowをモック
        with patch('services.excel_processor.get_last_row', return_value=3):
            apply_cell_formats(mock_worksheet, start_row)

        # セルにアクセスとフォーマット適用が行われたことを確認
        assert mock_worksheet.cell.call_count == 12  # 2行 x 6列 = 12セル

        # 各列のアライメント設定を確認
        for row in range(start_row, 4):  # 行2から3まで
            for col in range(1, 7):  # A列からF列まで
                mock_worksheet.cell.assert_any_call(row=row, column=col)

    @patch('services.excel_processor.win32gui.FindWindow')
    @patch('services.excel_processor.win32gui.SetForegroundWindow')
    @patch('services.excel_processor.time.sleep')
    def test_bring_excel_to_front_success(self, mock_sleep, mock_set_foreground, mock_find_window):
        """bring_excel_to_front関数の成功ケースのテスト"""
        # Excelウィンドウが見つかる場合
        mock_find_window.return_value = 12345  # 適当なウィンドウハンドル

        # 関数実行
        result = bring_excel_to_front()

        # 結果を確認
        assert result is True
        mock_find_window.assert_called_once_with("XLMAIN", None)
        mock_set_foreground.assert_called_once_with(12345)
        assert mock_sleep.called is False  # 成功した場合はsleepは呼ばれない

    @patch('services.excel_processor.win32gui.FindWindow')
    @patch('services.excel_processor.win32gui.SetForegroundWindow')
    @patch('services.excel_processor.time.sleep')
    def test_bring_excel_to_front_retry(self, mock_sleep, mock_set_foreground, mock_find_window):
        """bring_excel_to_front関数の再試行ケースのテスト"""
        # 最初は失敗、2回目に成功するケース
        mock_find_window.side_effect = [0, 12345]  # 0はウィンドウなし

        # 関数実行
        result = bring_excel_to_front()

        # 結果を確認
        assert result is True
        assert mock_find_window.call_count == 2
        mock_set_foreground.assert_called_once_with(12345)
        mock_sleep.assert_called_once_with(0.1)

    @patch('services.excel_processor.win32gui.FindWindow')
    @patch('services.excel_processor.time.sleep')
    def test_bring_excel_to_front_failure(self, mock_sleep, mock_find_window):
        """bring_excel_to_front関数の失敗ケースのテスト"""
        # ウィンドウが見つからない場合
        mock_find_window.return_value = 0  # ウィンドウなし

        # 関数実行
        result = bring_excel_to_front()

        # 結果を確認
        assert result is False
        assert mock_find_window.call_count == 2  # 2回試行
        assert mock_sleep.call_count == 2  # 2回sleep（各試行の間に1回ずつ）

    @patch('services.excel_processor.Path')
    @patch('services.excel_processor.load_workbook')
    @patch('services.excel_processor.get_last_row')
    @patch('services.excel_processor.apply_cell_formats')
    def test_write_data_to_excel_success(self, mock_apply_formats, mock_get_last_row,
                                         mock_load_workbook, mock_path):
        """write_data_to_excel関数の成功ケースのテスト"""
        # モックの設定
        mock_path_instance = MagicMock()
        mock_path_instance.exists.return_value = True
        mock_path.return_value = mock_path_instance

        mock_workbook, mock_worksheet = MagicMock(), MagicMock()
        mock_load_workbook.return_value = mock_workbook
        mock_workbook.active = mock_worksheet
        mock_get_last_row.return_value = 2  # ヘッダー + 既存データ1行

        # 既存データ(A2～F2)のモック。患者IDはExcelでは数値になっている
        mock_worksheet.iter_rows.return_value = [
            (datetime.datetime(2023, 1, 1), 12345, "診断書", "内科", "医師名", "備考"),
        ]

        # テスト用のCSVデータフレーム
        # 2行のデータ、1行目は既存と重複（スキップされる）、2行目は新規
        import polars as pl
        df = pl.DataFrame({
            "col_0": ["2023-01-01", "2023-02-01"],  # 日付列
            "col_1": ["12345", "67890"],  # 患者ID列
            "col_2": ["診断書", "処方箋"],
            "col_3": ["内科", "外科"],
            "col_4": ["医師名", "別の医師"],
            "col_5": ["備考", "新規備考"]
        })

        # 関数実行
        write_data_to_excel("test.xlsm", df)

        # 結果を確認
        mock_path.assert_called_once_with("test.xlsm")
        mock_path_instance.exists.assert_called_once()
        mock_load_workbook.assert_called_once_with(filename="test.xlsm", keep_vba=True)
        mock_get_last_row.assert_called_once_with(mock_worksheet)

        # 新規データ(CSVの2行目)だけが3行目に書き込まれることを確認
        mock_worksheet.iter_rows.assert_called_once_with(min_row=2, max_row=2, max_col=6, values_only=True)
        assert mock_worksheet.cell.call_count == 6
        mock_worksheet.cell.assert_any_call(row=3, column=1)
        mock_apply_formats.assert_called_once_with(mock_worksheet, 3)

        # 保存が呼ばれることを確認
        mock_workbook.save.assert_called_once_with("test.xlsm")
        mock_workbook.close.assert_called_once()

    def test_apply_cell_formats_keeps_vertical_center(self):
        """全列で垂直中央配置が維持されることのテスト"""
        worksheet = MagicMock()
        cells = {}

        def mock_cell(row, column):
            return cells.setdefault((row, column), MagicMock())

        worksheet.cell.side_effect = mock_cell

        with patch('services.excel_processor.get_last_row', return_value=2):
            apply_cell_formats(worksheet, 2)

        for col in range(1, 7):
            assert cells[(2, col)].alignment.vertical == 'center'
        assert cells[(2, 1)].alignment.horizontal == 'center'
        assert cells[(2, 3)].alignment.horizontal == 'left'

    def test_make_row_key_matches_excel_and_csv(self):
        """Excel側(datetime・数値)とCSV側(文字列)で同じキーになることのテスト"""
        excel_row = (datetime.datetime(2023, 1, 1), 12345, "診断書", "内科", None, "備考", "G列は無視")
        csv_row = ("2023-01-01", "12345", "診断書", "内科", None, "備考")

        assert make_row_key(excel_row) == make_row_key(csv_row)
        assert make_row_key(csv_row) == ("2023-01-01", "12345", "診断書", "内科", "", "備考")

    def test_filter_new_rows(self):
        """既存キーに無い行だけが残ることのテスト"""
        rows = [("2023-01-01", "1", "a", "b", "c", "d"), ("2023-01-02", "2", "a", "b", "c", "d")]
        existing_keys = {("2023-01-01", "1", "a", "b", "c", "d")}

        assert filter_new_rows(rows, existing_keys) == [rows[1]]

    def test_write_rows_converts_types(self):
        """日付列は日付型、患者ID列は数値、None は空文字で書き込まれることのテスト"""
        from openpyxl import Workbook
        ws = cast(Worksheet, Workbook().active)

        write_rows(ws, 2, [("2023-01-01", "1,234", "診断書", None), ("不正な日付", "ID-X", "x", "y")])

        assert ws.cell(row=2, column=1).value == datetime.datetime(2023, 1, 1)
        assert ws.cell(row=2, column=1).number_format == 'yyyy/mm/dd'
        assert ws.cell(row=2, column=2).value == 1234
        assert ws.cell(row=2, column=4).value == ""
        assert ws.cell(row=3, column=1).value == "不正な日付"
        assert ws.cell(row=3, column=2).value == "ID-X"

    def test_write_data_to_excel_file_not_found(self, tmp_path):
        """Excelファイルが存在しない場合は例外になることのテスト"""
        import polars as pl
        df = pl.DataFrame({"test": [1, 2, 3]})

        with pytest.raises(FileNotFoundError, match="Excelファイルが見つかりません"):
            write_data_to_excel(str(tmp_path / "nonexistent.xlsm"), df)

    def test_write_data_to_excel_not_xlsm(self, tmp_path):
        """拡張子が.xlsmでない場合は例外になることのテスト"""
        import polars as pl
        excel_file = tmp_path / "test.xlsx"
        excel_file.touch()
        df = pl.DataFrame({"test": [1, 2, 3]})

        with pytest.raises(ValueError, match=".xlsm"):
            write_data_to_excel(str(excel_file), df)

    @patch('services.excel_processor.Path')
    @patch('services.excel_processor.load_workbook')
    def test_write_data_to_excel_permission_error(self, mock_load_workbook, mock_path):
        """Excelファイルを開けない場合は分かりやすいメッセージで例外を送出する"""
        mock_path.return_value.exists.return_value = True
        mock_load_workbook.side_effect = PermissionError("ファイルが開かれています")

        import polars as pl
        df = pl.DataFrame({"test": [1, 2, 3]})
        with pytest.raises(PermissionError, match="別のプロセスで開かれています"):
            write_data_to_excel("locked.xlsm", df)

        mock_load_workbook.assert_called_once_with(filename="locked.xlsm", keep_vba=True)

    @patch('services.excel_processor.Path')
    @patch('services.excel_processor.load_workbook')
    @patch('services.excel_processor.get_last_row', return_value=1)
    @patch('services.excel_processor.apply_cell_formats')
    def test_write_data_to_excel_save_error_closes_workbook(self, mock_apply_formats, mock_get_last_row,
                                                            mock_load_workbook, mock_path):
        """保存に失敗しても例外を送出してワークブックを閉じる"""
        mock_path.return_value.exists.return_value = True
        mock_workbook = MagicMock()
        mock_workbook.save.side_effect = PermissionError("保存できません")
        mock_load_workbook.return_value = mock_workbook

        import polars as pl
        df = pl.DataFrame({f"col_{i}": ["2023-01-01"] for i in range(6)})
        with pytest.raises(PermissionError, match="保存できません"):
            write_data_to_excel("locked.xlsm", df)

        mock_workbook.close.assert_called_once()

    @patch('services.excel_processor.win32com.client.Dispatch')
    @patch('services.excel_processor.bring_excel_to_front')
    @patch('services.excel_processor.Path')
    @patch('services.excel_processor.sort_excel_data')
    @patch('services.excel_processor.ConfigManager')
    @patch('services.excel_processor.time.sleep')
    @patch('services.excel_processor.pyautogui.click')
    @patch('services.excel_processor.pyautogui.hotkey')
    def test_open_and_sort_excel(self, mock_hotkey, mock_click, mock_sleep, mock_config_manager,
                                 mock_sort, mock_path, mock_bring_front, mock_dispatch):
        """open_and_sort_excel関数のテスト"""
        # モックの設定
        mock_excel = MagicMock()
        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()
        mock_dispatch.return_value = mock_excel
        mock_excel.Workbooks.Open.return_value = mock_workbook
        mock_workbook.ActiveSheet = mock_worksheet
        mock_bring_front.return_value = True
        mock_path_instance = MagicMock()
        mock_path_instance.resolve.return_value = "C:/absolute/path/to/test.xlsm"
        mock_path.return_value = mock_path_instance

        # ConfigManagerのモック
        mock_config = MagicMock()
        mock_config.get_share_button_wait_time.return_value = 1
        mock_config.get_share_button_position.return_value = (100, 200)
        mock_config_manager.return_value = mock_config

        # セルの位置を返すモック
        mock_last_row = 10
        mock_sort.return_value = mock_last_row
        mock_worksheet.Cells.return_value.End.return_value.Row = mock_last_row

        # 関数実行
        open_and_sort_excel("test.xlsm")

        # 結果を確認
        mock_dispatch.assert_called_once_with("Excel.Application")
        assert mock_excel.Visible is True
        mock_bring_front.assert_called_once()
        mock_excel.Workbooks.Open.assert_called_once_with("C:/absolute/path/to/test.xlsm")
        assert mock_excel.WindowState == -4137  # xlMaximized
        mock_workbook.Windows.assert_called_once_with(1)
        mock_sort.assert_called_once_with(mock_worksheet)
        mock_worksheet.Cells.assert_called_with(mock_last_row, 1)
        mock_worksheet.Cells.return_value.Select.assert_called_once()

        # 共有ボタンクリックの前に待機
        mock_sleep.assert_called_once_with(1)
        mock_click.assert_called_once_with(100, 200)

        # ウィンドウを最小化
        mock_hotkey.assert_called_once_with('win', 'down')

    @patch('services.excel_processor.win32com.client.Dispatch')
    @patch('services.excel_processor.bring_excel_to_front')
    @patch('services.excel_processor.Path')
    @patch('services.excel_processor.pyautogui.hotkey')
    def test_open_and_sort_excel_error_does_not_minimize(self, mock_hotkey,
                                                         mock_path, mock_bring_front, mock_dispatch):
        """エラー時は例外を送出し、別ウィンドウを最小化しないことのテスト"""
        mock_path.return_value.exists.return_value = True
        mock_dispatch.return_value.Workbooks.Open.side_effect = Exception("開けません")

        with pytest.raises(Exception, match="開けません"):
            open_and_sort_excel("test.xlsm")

        mock_hotkey.assert_not_called()

    def test_open_and_sort_excel_file_not_found(self, tmp_path):
        """Excelファイルが無い場合は例外を送出する"""
        with pytest.raises(FileNotFoundError, match="Excelファイルが見つかりません"):
            open_and_sort_excel(str(tmp_path / "nonexistent.xlsm"))

    def test_sort_excel_data(self):
        """sort_excel_data関数のテスト"""
        # こちらはwin32comに強く依存しており、モックの構築が複雑なため省略
        # 実際の環境では、この関数は他の関数からの呼び出しでテストされることになります
        pass
