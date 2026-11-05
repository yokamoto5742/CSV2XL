import csv


def write_csv(path: str, headers: list[str], rows: list[list[str]]) -> None:
    """Excel で文字化けせずに開けるよう、BOM 付き UTF-8 で書き出す"""
    with open(path, 'w', encoding='utf-8-sig', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(headers)
        writer.writerows(rows)
