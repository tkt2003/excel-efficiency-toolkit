from src.excel_efficiency_toolkit.sheet_ops import (
    _sheet_index_clear_end_row,
    _sheet_names_for_index,
    _sheet_a1_sub_address,
)


def test_sheet_a1_sub_address_escapes_single_quote():
    assert _sheet_a1_sub_address("O'Brien") == "'O''Brien'!A1"


def test_sheet_index_clear_end_row_covers_old_and_new_directory_area():
    assert _sheet_index_clear_end_row(sheet_count=3, used_last_row=2) == 4
    assert _sheet_index_clear_end_row(sheet_count=3, used_last_row=10) == 10
    assert _sheet_index_clear_end_row(sheet_count=0, used_last_row=0) == 1


def test_sheet_names_for_index_excludes_index_sheet_case_insensitive():
    assert _sheet_names_for_index(["工作表目录", "Data", "Summary"], "工作表目录") == ["Data", "Summary"]
    assert _sheet_names_for_index(["Index", "Data"], "index") == ["Data"]
