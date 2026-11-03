import datetime
import time
from pathlib import Path
from typing import Any, Sequence, cast

import polars as pl
import pyautogui
import win32com.client
import win32gui
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

from utils.config_manager import ConfigManager


def get_last_row(worksheet: Worksheet) -> int:
    """ワークシートの最後のデータ行番号を取得

    Args:
        worksheet: openpyxlのワークシートオブジェクト

    Returns:
        最後のデータが存在する行番号
    """
    last_row = 0
    for row in worksheet.iter_rows():
        if all(cell.value is None for cell in row):
            break
        last_row += 1
    return last_row


def apply_cell_formats(worksheet: Worksheet, start_row: int) -> None:
    """指定された行からセルのフォーマットと配置を適用

    Args:
        worksheet: openpyxlのワークシートオブジェクト
        start_row: フォーマット開始行番号
    """
    last_row = get_last_row(worksheet)

    # A～F列に対して、垂直中央配置を設定し、特定列の水平配置を調整
    for row in range(start_row, last_row + 1):
        for col in range(1, 7):
            cell = worksheet.cell(row=row, column=col)

            if col in [3, 4]:  # C列とD列
                cell.alignment = Alignment(vertical='center', horizontal='left', shrink_to_fit=True)
            else:
                cell.alignment = Alignment(vertical='center', horizontal='center')


def sort_excel_data(worksheet: Any) -> int:
    """Excelデータを預り日、診療科、患者IDでソート

    Args:
        worksheet: win32comのワークシートオブジェクト

    Returns:
        ソート後の最終行番号
    """
    last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row

    # ソートの範囲を設定
    sort_range = worksheet.Range(f"A2:I{last_row}")

    sort_range.Sort(
        Key1=worksheet.Range("A2"),  # A列（預り日）
        Order1=1,  # 1=昇順
        Key2=worksheet.Range("E2"),  # E列（診療科）
        Order2=1,  # 1=昇順
        Key3=worksheet.Range("B2"),  # B列（患者ID）
        Order3=1,  # 1=昇順
        Header=1,  # 1=ヘッダーあり
        OrderCustom=1,
        MatchCase=False,
        Orientation=1)

    return last_row


def bring_excel_to_front() -> bool:
    """Excelウィンドウを最前面に表示

    Returns:
        成功時はTrue、失敗時はFalse
    """
    for _ in range(2):
        hwnd = win32gui.FindWindow("XLMAIN", None)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            return True
        time.sleep(0.1)
    return False


def make_row_key(row: Sequence[Any]) -> tuple[str, ...]:
    """重複チェック用のキー(A～F列の値を文字列にしたもの)を作成

    日付は Excel 側(datetime)と CSV 側(文字列)で型が違うため YYYY-MM-DD にそろえる
    """
    date_value = row[0]
    if isinstance(date_value, datetime.date):
        date_str = date_value.strftime('%Y-%m-%d')
    else:
        date_str = str(date_value or '')
    return (date_str, *(str(value or '') for value in row[1:6]))


def read_existing_keys(ws: Worksheet, last_row: int) -> set[tuple[str, ...]]:
    """Excelの既存データ(2行目以降)の重複チェック用キーを取得"""
    rows = ws.iter_rows(min_row=2, max_row=last_row, max_col=6, values_only=True)
    return {make_row_key(row) for row in rows}


def filter_new_rows(rows: list[tuple[Any, ...]], existing_keys: set[tuple[str, ...]]) -> list[tuple[Any, ...]]:
    """既存データに無い行だけを返す"""
    return [row for row in rows if make_row_key(row) not in existing_keys]


def write_rows(ws: Worksheet, start_row: int, rows: list[tuple[Any, ...]]) -> None:
    """行データをセルに書き込み、日付列は日付型、患者ID列は数値に変換する"""
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            cell = cast(Any, ws.cell(row=start_row + i, column=j + 1))

            if j == 0:  # 日付列を日付型に変換
                try:
                    cell.value = datetime.datetime.strptime(value, '%Y-%m-%d')
                    cell.number_format = 'yyyy/mm/dd'
                except ValueError:
                    cell.value = value
            elif j == 1:  # 患者ID列
                try:
                    cell.value = int(value.replace(',', ''))
                    cell.number_format = '0'
                except ValueError:
                    cell.value = value
            else:
                cell.value = value if value is not None else ""


def write_data_to_excel(excel_path: str, df: pl.DataFrame) -> None:
    """DataFrameのデータをExcelファイルに重複排除して書き込み

    既存データを確認して重複していないデータのみを追加。日付と患者IDの形式変換も実施

    Args:
        excel_path: Excelファイルのパス
        df: 書き込むpolarsのDataFrame
    """
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")
    if not excel_path.endswith('.xlsm'):
        raise ValueError(f"Excelファイルは .xlsm 形式を指定してください: {excel_path}")

    try:
        wb = load_workbook(filename=excel_path, keep_vba=True)
    except PermissionError as e:
        raise PermissionError(
            "Excelファイルが別のプロセスで開かれています。\nファイルを閉じてから再度実行してください。"
        ) from e

    ws = cast(Worksheet, wb.active)
    last_row = get_last_row(ws)

    # DataFrameの値はすべて文字列にする(日付列は YYYY-MM-DD になる)
    csv_rows = df.select(pl.col('*').cast(pl.String)).rows()
    new_rows = filter_new_rows(csv_rows, read_existing_keys(ws, last_row))

    write_rows(ws, last_row + 1, new_rows)
    apply_cell_formats(ws, last_row + 1)

    try:
        wb.save(excel_path)
    except PermissionError as e:
        raise PermissionError(
            "Excelファイルが別のプロセスで開かれているため、保存できません。\nファイルを閉じてから再度実行してください。"
        ) from e
    finally:
        wb.close()


def clear_all_filters(worksheet: Any, workbook: Any) -> None:
    """ワークシートのフィルタをすべてクリアして解除"""
    try:
        is_shared = False
        try:
            # 共有ブックかどうかを確認
            is_shared = workbook.MultiUserEditing
        except Exception:
            pass

        if not worksheet.AutoFilterMode:
            print("オートフィルタは設定されていません")
            return

        # フィルタ条件が適用されていればすべてのデータを表示
        try:
            if worksheet.FilterMode:
                worksheet.ShowAllData()
                print("フィルタ条件をクリアしました")
        except Exception as e:
            print(f"ShowAllData実行中にエラー: {str(e)}")

        # 念のため各列のフィルタを個別にクリアする
        try:
            if worksheet.AutoFilterMode:
                auto_filter = worksheet.AutoFilter
                if auto_filter is not None:
                    filters = auto_filter.Filters
                    for i in range(1, filters.Count + 1):
                        try:
                            if filters.Item(i).On:
                                # 各列のフィルタ条件をクリア
                                auto_filter.Range.AutoFilter(Field=i)
                        except Exception:
                            pass
                    print("個別のフィルタ条件をクリアしました")
        except Exception as e:
            print(f"個別フィルタクリア中にエラー: {str(e)}")

        # 共有ブックでない場合のみオートフィルタ自体を解除
        if not is_shared:
            try:
                worksheet.AutoFilterMode = False
                print("オートフィルタを解除しました")
            except Exception as e:
                print(f"AutoFilterMode解除中にエラー（共有ブックの可能性）: {str(e)}")
        else:
            print("共有ブックのため、オートフィルタの解除はスキップしました（フィルタ条件はクリア済み）")

    except Exception as e:
        print(f"フィルタ解除中にエラーが発生しました: {str(e)}")


def open_and_sort_excel(excel_path: str) -> None:
    """Excelファイルを開いてデータをソート、共有ボタンをクリック

    Args:
        excel_path: Excelファイルのパス
    """
    excel_path_obj = Path(excel_path)

    if not excel_path_obj.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    bring_excel_to_front()

    # Excelファイルを開く
    workbook = excel.Workbooks.Open(str(excel_path_obj.resolve()))

    if workbook is None:
        raise RuntimeError("Excelファイルを開くことができませんでした。")

    excel.WindowState = -4137  # xlMaximized
    workbook.Windows(1).Activate()

    worksheet = workbook.ActiveSheet

    # フィルタがかかっていればクリア
    clear_all_filters(worksheet, workbook)

    # データをソート
    sort_excel_data(worksheet)

    # 最終行にカーソルを移動
    last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row
    worksheet.Cells(last_row, 1).Select()

    # 設定に従って共有ボタンをクリック
    config = ConfigManager()
    wait_time = config.get_share_button_wait_time()
    time.sleep(wait_time)
    share_x, share_y = config.get_share_button_position()
    pyautogui.click(share_x, share_y)

    pyautogui.hotkey('win', 'down')  # ウィンドウを最小化
