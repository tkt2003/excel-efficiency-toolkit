from src.excel_efficiency_toolkit.delete_sheet_ops import (
    build_rule_table_rows,
    build_source_path_rows,
    collect_unique_sheet_names,
    copy_source_to_output,
    get_workbook_sheet_names,
    get_worksheet_by_name,
    get_unique_output_path,
    infer_delete_mode_from_rule_values,
    is_excel_file,
    normalize_sheet_name_for_compare,
    normalize_delete_mode,
    parse_source_paths_from_rows,
    plan_sheets_to_delete,
    read_rule_values_from_rows,
    should_exclude_file,
)


class FakeSheet:
    def __init__(self, name):
        self.Name = name


class FakeWorkbook:
    def __init__(self, sheet_names):
        self.Worksheets = [FakeSheet(name) for name in sheet_names]


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


def test_build_source_path_rows_returns_source_file_sheet_rows():
    assert build_source_path_rows(["C:/a.xlsx", "D:/b.xlsx"]) == [
        ["源文件完整路径"],
        ["C:/a.xlsx"],
        ["D:/b.xlsx"],
    ]


def test_get_worksheet_by_name_finds_sheet_without_relying_on_order():
    workbook = FakeWorkbook(["源文件清单", "批量删除规则"])

    assert get_worksheet_by_name(workbook, "批量删除规则").Name == "批量删除规则"


def test_get_worksheet_by_name_rejects_missing_sheet():
    workbook = FakeWorkbook(["源文件清单"])

    try:
        get_worksheet_by_name(workbook, "批量删除规则")
    except ValueError as e:
        assert "批量删除规则" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_get_workbook_sheet_names_returns_names_in_order():
    workbook = FakeWorkbook(["批量删除规则", "源文件清单"])

    assert get_workbook_sheet_names(workbook) == ["批量删除规则", "源文件清单"]


def test_get_unique_output_path_returns_original_when_not_exists(tmp_path):
    output_path = tmp_path / "规则表.xlsx"

    assert get_unique_output_path(str(output_path)) == str(output_path)


def test_get_unique_output_path_appends_suffix_when_exists(tmp_path):
    output_path = tmp_path / "规则表.xlsx"
    output_path.write_text("", encoding="utf-8")
    (tmp_path / "规则表_2.xlsx").write_text("", encoding="utf-8")

    assert get_unique_output_path(str(output_path)) == str(tmp_path / "规则表_3.xlsx")


def test_read_rule_values_from_rows_reads_b_c_d_and_ignores_a_column():
    rows = [
        ["所有表格名", "保留表格名", "删除表格名", "排除文件名关键词"],
        ["参考A", " 保留1 ", " 删除1 ", " 客户A "],
        ["参考B", "保留2", "", "客户B"],
    ]

    assert read_rule_values_from_rows(rows) == {
        "keep_names": ["保留1", "保留2"],
        "delete_names": ["删除1"],
        "exclude_keywords": ["客户A", "客户B"],
    }


def test_read_rule_values_from_rows_ignores_blank_cells():
    rows = [
        ["所有表格名", "保留表格名", "删除表格名", "排除文件名关键词"],
        ["A", None, "   ", ""],
        ["B", "保留", None, "  排除  "],
    ]

    assert read_rule_values_from_rows(rows) == {
        "keep_names": ["保留"],
        "delete_names": [],
        "exclude_keywords": ["排除"],
    }


def test_parse_source_paths_from_rows_skips_header_and_blank_cells():
    rows = [
        ["源文件完整路径"],
        [" C:/a.xlsx "],
        [""],
        [None],
        ["D:/b.xlsx"],
    ]

    assert parse_source_paths_from_rows(rows) == ["C:/a.xlsx", "D:/b.xlsx"]


def test_infer_delete_mode_from_rule_values_returns_keep_when_only_keep_names_exist():
    assert infer_delete_mode_from_rule_values({
        "keep_names": ["A"],
        "delete_names": [],
        "exclude_keywords": [],
    }) == "keep"


def test_infer_delete_mode_from_rule_values_returns_delete_when_only_delete_names_exist():
    assert infer_delete_mode_from_rule_values({
        "keep_names": [],
        "delete_names": ["A"],
        "exclude_keywords": [],
    }) == "delete"


def test_infer_delete_mode_from_rule_values_returns_none_when_keep_and_delete_both_exist():
    assert infer_delete_mode_from_rule_values({
        "keep_names": ["A"],
        "delete_names": ["B"],
        "exclude_keywords": [],
    }) is None


def test_infer_delete_mode_from_rule_values_rejects_empty_keep_and_delete_names():
    try:
        infer_delete_mode_from_rule_values({
            "keep_names": [" "],
            "delete_names": [],
            "exclude_keywords": ["客户"],
        })
    except ValueError as e:
        assert "B 列填写保留表格名" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_should_exclude_file_matches_keyword_case_insensitive():
    assert should_exclude_file("客户A_报表.xlsx", ["客户a"]) is True
    assert should_exclude_file("summary.xlsx", ["客户a"]) is False


def test_plan_sheets_to_delete_keep_mode_keeps_only_requested_sheets():
    assert plan_sheets_to_delete(["A", "B", "C"], "keep", ["B"], []) == ["A", "C"]


def test_plan_sheets_to_delete_delete_mode_deletes_only_requested_sheets():
    assert plan_sheets_to_delete(["A", "B", "C"], "delete", [], ["B"]) == ["B"]


def test_plan_sheets_to_delete_matches_case_insensitive_and_stripped():
    assert plan_sheets_to_delete([" Data ", "Summary"], "delete", [], ["data"]) == [" Data "]
    assert plan_sheets_to_delete(["Data", "Summary"], "keep", [" summary "], []) == ["Data"]


def test_plan_sheets_to_delete_rejects_deleting_all_sheets():
    try:
        plan_sheets_to_delete(["A"], "delete", [], ["A"])
    except ValueError as e:
        assert "没有任何工作表" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_plan_sheets_to_delete_rejects_empty_keep_rules():
    try:
        plan_sheets_to_delete(["A", "B"], "keep", [], [])
    except ValueError as e:
        assert "B 列不能为空" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_plan_sheets_to_delete_rejects_empty_delete_rules():
    try:
        plan_sheets_to_delete(["A", "B"], "delete", [], [])
    except ValueError as e:
        assert "C 列不能为空" in str(e)
    else:
        raise AssertionError("expected ValueError")


def test_normalize_delete_mode_accepts_number_chinese_and_english():
    assert normalize_delete_mode("1") == "keep"
    assert normalize_delete_mode("保留") == "keep"
    assert normalize_delete_mode("keep") == "keep"
    assert normalize_delete_mode("保留模式") == "keep"
    assert normalize_delete_mode("2") == "delete"
    assert normalize_delete_mode("删除") == "delete"
    assert normalize_delete_mode("delete") == "delete"
    assert normalize_delete_mode("删除模式") == "delete"


def test_copy_source_to_output_copies_file_and_avoids_overwrite(tmp_path):
    source_path = tmp_path / "源.xlsx"
    source_path.write_text("source", encoding="utf-8")
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    existing_path = output_dir / "源.xlsx"
    existing_path.write_text("existing", encoding="utf-8")

    copied_path = copy_source_to_output(str(source_path), str(output_dir))

    assert copied_path == str(output_dir / "源_2.xlsx")
    assert (output_dir / "源_2.xlsx").read_text(encoding="utf-8") == "source"
    assert existing_path.read_text(encoding="utf-8") == "existing"
