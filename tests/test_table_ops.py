import pytest

from src.excel_efficiency_toolkit.table_ops import (
    build_split_targets,
    normalize_header_row,
    parse_column_index,
    resolve_source_sheet_name,
    validate_row_numbers,
)


def test_normalize_header_row_converts_values_to_stripped_strings():
    assert normalize_header_row([" 姓名 ", None, 123, ""]) == ["姓名", "", "123", ""]


def test_validate_row_numbers_accepts_valid_rows():
    validate_row_numbers(1, 2)
    validate_row_numbers(3, 10)


@pytest.mark.parametrize(
    ("header_row", "data_start_row"),
    [
        (0, 2),
        (-1, 2),
        (1, 1),
        (2, 1),
    ],
)
def test_validate_row_numbers_rejects_invalid_rows(header_row, data_start_row):
    with pytest.raises(ValueError):
        validate_row_numbers(header_row, data_start_row)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("A", 1),
        ("B", 2),
        ("Z", 26),
        ("AA", 27),
        ("AB", 28),
        ("a", 1),
    ],
)
def test_parse_column_index_accepts_excel_letters(value, expected):
    assert parse_column_index(value) == expected


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("1", 1),
        ("2", 2),
        ("27", 27),
        (" 3 ", 3),
    ],
)
def test_parse_column_index_accepts_numbers(value, expected):
    assert parse_column_index(value) == expected


@pytest.mark.parametrize("value", ["", "   ", "0", "-1", "A1", "A-B", "*", "中文"])
def test_parse_column_index_rejects_invalid_values(value):
    with pytest.raises(ValueError):
        parse_column_index(value)


def test_build_split_targets_deduplicates_and_keeps_first_seen_order():
    assert build_split_targets(["A", "B", "A", " C ", "C"]) == ["A", "B", "C"]


def test_build_split_targets_converts_blank_values_to_blank_label():
    assert build_split_targets([None, "", "   ", "A", None]) == ["空白", "A"]


def test_resolve_source_sheet_name_uses_single_sheet_when_input_is_empty():
    assert resolve_source_sheet_name("", ["Data"]) == "Data"
    assert resolve_source_sheet_name(None, ["Data"]) == "Data"


def test_resolve_source_sheet_name_requires_name_for_multiple_sheets():
    with pytest.raises(ValueError, match="多个工作表"):
        resolve_source_sheet_name("", ["Data", "Summary"])


def test_resolve_source_sheet_name_matches_case_insensitive_name():
    assert resolve_source_sheet_name("data", ["Data", "Summary"]) == "Data"


def test_resolve_source_sheet_name_rejects_missing_name():
    with pytest.raises(ValueError, match="未找到源 sheet"):
        resolve_source_sheet_name("Missing", ["Data"])
