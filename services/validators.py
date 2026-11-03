import datetime
import unicodedata

DATE_INPUT_FORMATS = ('%Y/%m/%d', '%Y-%m-%d')


def normalize_input(text: str) -> str:
    """全角の数字や記号を半角にして前後の空白を除く"""
    return unicodedata.normalize('NFKC', text).strip()


def parse_date(text: str) -> str | None:
    """YYYY/MM/DD または YYYY-MM-DD の入力を 'YYYY-MM-DD' にする

    空欄なら None を返し、日付として正しくなければ ValueError を送出する
    """
    normalized = normalize_input(text)
    if not normalized:
        return None
    for date_format in DATE_INPUT_FORMATS:
        try:
            return datetime.datetime.strptime(normalized, date_format).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"日付の形式が正しくありません: {text}")


def parse_patient_id(text: str) -> int | None:
    """患者IDの入力を数値にする。数字でなければ None を返す"""
    normalized = normalize_input(text)
    if not normalized.isdecimal():
        return None
    return int(normalized)
