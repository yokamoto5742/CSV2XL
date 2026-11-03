import pytest

from services.validators import parse_date, parse_patient_id


class TestParseDate:
    @pytest.mark.parametrize("text, expected", [
        ("2025/01/05", "2025-01-05"),
        ("2025/1/5", "2025-01-05"),
        ("2025-01-05", "2025-01-05"),
        (" 2025/01/05 ", "2025-01-05"),
        ("２０２５／０１／０５", "2025-01-05"),
    ])
    def test_valid(self, text: str, expected: str) -> None:
        assert parse_date(text) == expected

    @pytest.mark.parametrize("text", ["", "  "])
    def test_empty_returns_none(self, text: str) -> None:
        assert parse_date(text) is None

    @pytest.mark.parametrize("text", ["2025/02/30", "2025/13/01", "20250105", "1/5", "abc", "2025.01.05"])
    def test_invalid_raises(self, text: str) -> None:
        with pytest.raises(ValueError):
            parse_date(text)


class TestParsePatientId:
    @pytest.mark.parametrize("text, expected", [("12345", 12345), (" 42 ", 42), ("１２３", 123)])
    def test_valid(self, text: str, expected: int) -> None:
        assert parse_patient_id(text) == expected

    @pytest.mark.parametrize("text", ["", "abc", "12.5", "-1", "1 2"])
    def test_invalid_returns_none(self, text: str) -> None:
        assert parse_patient_id(text) is None
