from src.excel_efficiency_toolkit.delete_sheet_ops import (
    build_rule_table_rows,
    collect_unique_sheet_names,
    get_unique_output_path,
    is_excel_file,
    normalize_sheet_name_for_compare,
)


def test_is_excel_file_accepts_supported_extensions_case_insensitive():
    assert is_excel_file("book.xlsx") is True
    assert is_excel_file("book.XLSM") is True
    assert is_excel_file("book.xls") is True


def test_is_excel_file_rejects_temporary_excel_files():
    assert is_excel_file("~$book.xlsx") is False
    assert is_excel_file("C:/tmp/~$book.xlsm") is False


def test_is_excel_file_rejects_non_excel_files():
    assert is_excel_file("book.csv") is False
    assert is_excel_file("book.txt") is False


def test_normalize_sheet_name_for_compare_strips_and_lowers():
    assert normalize_sheet_name_for_compare(None) == ""
    assert normalize_sheet_name_for_compare(" Data ") == "data"


def test_collect_unique_sheet_names_keeps_first_seen_order_and_original_name():
    sheet_name_lists = [
        [" Data ", "Summary", "data", ""],
        ["SUMMARY", "Detail"],
    ]

    assert collect_unique_sheet_names(sheet_name_lists) == [" Data ", "Summary", "Detail"]


def test_collect_unique_sheet_names_ignores_blank_names():
    assert collect_unique_sheet_names([["", "   ", "A"]]) == ["A"]


def test_build_rule_table_rows_returns_expected_table_shape():
    assert build_rule_table_rows(["A", "B"]) == [
        ["所有表格名", "保留表格名", "删除表格名", "排除文件名关键词"],
        ["A", "", "", ""],
        ["B", "", "", ""],
    ]


def test_get_unique_output_path_returns_original_when_not_exists(tmp_path):
    output_path = tmp_path / "规则表.xlsx"

    assert get_unique_output_path(str(output_path)) == str(output_path)


def test_get_unique_output_path_appends_suffix_when_exists(tmp_path):
    output_path = tmp_path / "规则表.xlsx"
    output_path.write_text("", encoding="utf-8")
    (tmp_path / "规则表_2.xlsx").write_text("", encoding="utf-8")

    assert get_unique_output_path(str(output_path)) == str(tmp_path / "规则表_3.xlsx")
