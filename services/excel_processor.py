import datetime
import time
from pathlib import Path

import polars as pl
import pyautogui
import win32com.client
import win32gui
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from PyQt6.QtWidgets import QMessageBox

from utils.config_manager import ConfigManager


def get_last_row(worksheet):
    last_row = 0
    for row in worksheet.iter_rows():
        if all(cell.value is None for cell in row):
            break
        last_row += 1
    return last_row


def apply_cell_formats(worksheet, start_row):
    last_row = get_last_row(worksheet)

    # A列からF列までの範囲を設定
    for row in range(start_row, last_row + 1):
        for col in range(1, 7):
            cell = worksheet.cell(row=row, column=col)

            cell.alignment = Alignment(vertical='center')

            if col in [1, 2, 5, 6]:  # A,B列とE,F列
                cell.alignment = Alignment(horizontal='center')

            elif col in [3, 4]:  # C列とD列
                cell.alignment = Alignment(horizontal='left', shrink_to_fit=True)


def sort_excel_data(worksheet):
    try:
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

    except Exception as e:
        print(f"ソート中にエラーが発生しました: {str(e)}")
        raise


def bring_excel_to_front():
    # Excelの表示を最前面にする（最大3回まで試行）
    for _ in range(2):
        hwnd = win32gui.FindWindow("XLMAIN", None)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            return True
        time.sleep(0.1)
    return False


def write_data_to_excel(excel_path, df):
    if not Path(excel_path).exists() or not excel_path.endswith('.xlsm'):
        print(f"Excelファイルが見つかりません: {excel_path}")
        return False

    try:
        wb = load_workbook(filename=excel_path, keep_vba=True)
    except PermissionError:
        QMessageBox.critical(None,
                             "エラー",
                             "Excelファイルが別のプロセスで開かれています。\nファイルを閉じてから再度実行してください。"
                             )
        return False

    ws = wb.active

    # 実際のデータが存在する最終行を取得
    last_row = get_last_row(ws)

    # 既存データを保持するセットを作成（A列からF列までの値をキーとして使用）
    existing_data = set()
    for row in range(2, last_row + 1):  # ヘッダー行をスキップ
        # 日付をYYYYMMDD形式の文字列として取得
        cell1 = ws.cell(row=row, column=1)  # type: ignore[misc]
        date_value = cell1.value if cell1 else None
        if isinstance(date_value, datetime.datetime):
            date_str = date_value.strftime('%Y%m%d')
        else:
            date_str = str(date_value or '')

        # A列からF列までの値を取得（日付は数値形式で保持）
        cell2 = ws.cell(row=row, column=2)  # type: ignore[misc]
        cell3 = ws.cell(row=row, column=3)  # type: ignore[misc]
        cell4 = ws.cell(row=row, column=4)  # type: ignore[misc]
        cell5 = ws.cell(row=row, column=5)  # type: ignore[misc]
        cell6 = ws.cell(row=row, column=6)  # type: ignore[misc]
        row_data = (
            date_str,  # 日付を8桁の数値文字列として保持
            str(cell2.value or '') if cell2 else '',
            str(cell3.value or '') if cell3 else '',
            str(cell4.value or '') if cell4 else '',
            str(cell5.value or '') if cell5 else '',
            str(cell6.value or '') if cell6 else ''
        )
        existing_data.add(row_data)

    # CSVデータを文字列に変換
    temp_df = df.select([
        pl.col('*').cast(pl.String)
    ])
    data_to_write = temp_df.to_numpy().tolist()

    # 重複していないデータのみを抽出
    unique_data = []
    for row in data_to_write:
        # CSVの日付を8桁の数値文字列に変換
        csv_date = row[0]
        if isinstance(csv_date, str):
            try:
                # YYYY-MM-DD形式をYYYYMMDD形式に変換
                date_obj = datetime.datetime.strptime(csv_date, '%Y-%m-%d')
                date_str = date_obj.strftime('%Y%m%d')
            except ValueError:
                date_str = csv_date
        else:
            date_str = str(csv_date)

        # 比較用のタプルを作成
        row_data = (
            date_str,
            str(row[1] or ''),
            str(row[2] or ''),
            str(row[3] or ''),
            str(row[4] or ''),
            str(row[5] or '')
        )

        if row_data not in existing_data:
            unique_data.append(row)

    # 重複しないデータのみを書き込む
    for i, row in enumerate(unique_data):
        for j, value in enumerate(row):
            cell = ws.cell(row=last_row + 1 + i, column=j + 1)  # type: ignore[misc]

            if j == 0:  # 日付列
                try:
                    date_value = datetime.datetime.strptime(value, '%Y-%m-%d')
                    cell.value = date_value  # type: ignore[attr-defined]
                    cell.number_format = 'yyyy/mm/dd'
                except ValueError:
                    cell.value = value  # type: ignore[attr-defined]
            elif j == 1:  # 患者ID列
                try:
                    cell.value = int(value.replace(',', ''))  # type: ignore[attr-defined]
                    cell.number_format = '0'
                except ValueError:
                    cell.value = value  # type: ignore[attr-defined]
            else:
                cell.value = value if value is not None else ""  # type: ignore[attr-defined]

    apply_cell_formats(ws, last_row + 1)

    try:
        wb.save(excel_path)
        wb.close()
        return True
    except PermissionError:
        QMessageBox.critical(None,
                             "エラー",
                             "Excelファイルが別のプロセスで開かれているため、保存できません。\nファイルを閉じてから再度実行してください。"
                             )
        if 'wb' in locals():
            wb.close()
        return False


def clear_all_filters(worksheet, workbook):
    """
    すべてのフィルタを解除する
    共有ブックの場合はAutoFilterModeの変更ができないため、
    ShowAllDataのみでフィルタ条件をクリアする
    """
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

        # フィルタ条件が適用されていれば、すべてのデータを表示
        try:
            if worksheet.FilterMode:
                worksheet.ShowAllData()
                print("フィルタ条件をクリアしました")
        except Exception as e:
            print(f"ShowAllData実行中にエラー: {str(e)}")

        # 各列のフィルタを個別にクリアする（ShowAllDataが効かない場合の対策）
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

        # 共有ブックでない場合のみ、オートフィルタ自体を解除
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


def open_and_sort_excel(excel_path):
    excel_path_obj = Path(excel_path)

    # ファイルの存在確認
    if not excel_path_obj.exists():
        QMessageBox.critical(None, "エラー", f"Excelファイルが見つかりません: {excel_path}")
        return

    excel_path_str = str(excel_path_obj.resolve())
    excel = None
    workbook = None

    try:
        # Excel アプリケーションを起動
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        bring_excel_to_front()

        # ファイルを開く
        workbook = excel.Workbooks.Open(excel_path_str)

        # workbook が正常に開かれたか確認
        if workbook is None:
            QMessageBox.critical(None, "エラー", "Excelファイルを開くことができませんでした。")
            return

        excel.WindowState = -4137  # xlMaximized
        workbook.Windows(1).Activate()

        worksheet = workbook.ActiveSheet

        # フィルタがかかっていれば解除（共有ブック対応）
        clear_all_filters(worksheet, workbook)

        sort_excel_data(worksheet)

        last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row
        worksheet.Cells(last_row, 1).Select()

        config = ConfigManager()
        wait_time = config.get_share_button_wait_time()
        time.sleep(wait_time)
        share_x, share_y = config.get_share_button_position()
        pyautogui.click(share_x, share_y)

    except Exception as e:
        error_msg = f"Excelファイルの処理中にエラーが発生しました: {str(e)}"
        print(error_msg)
        QMessageBox.critical(None, "エラー", error_msg)
    finally:
        # Excelは開いたままにするが、エラー処理は行う
        pyautogui.hotkey('win', 'down')  # ウィンドウを最小化
