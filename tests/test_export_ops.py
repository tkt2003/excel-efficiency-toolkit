from src.excel_efficiency_toolkit.export_ops import (
    clean_filename,
    get_safe_sheet_filename,
    get_unique_filepath,
    get_unique_name,
)


def test_clean_filename_removes_windows_invalid_chars_and_strips_spaces():
    assert clean_filename('  a\\b/c:d*e?f"g<h>i|j  ') == "abcdefghij"


def test_clean_filename_none_and_empty_string():
    assert clean_filename(None) == ""
    assert clean_filename("") == ""
    assert clean_filename("   ") == ""


def test_get_safe_sheet_filename_falls_back_when_cleaned_name_is_empty():
    assert get_safe_sheet_filename(" /:*?\"<>| ", 1) == "Sheet_1"
    assert get_safe_sheet_filename(None, 2) == "Sheet_2"


def test_get_safe_sheet_filename_truncates_long_name():
    long_name = "A" * 120
    assert get_safe_sheet_filename(long_name, 1) == "A" * 100


def test_get_unique_name_adds_name_to_used_names():
    used_names = set()

    assert get_unique_name("Sheet", used_names) == "Sheet"
    assert used_names == {"Sheet"}


def test_get_unique_name_appends_sequence_for_duplicates():
    used_names = {"Sheet", "Sheet_2"}

    assert get_unique_name("Sheet", used_names) == "Sheet_3"
    assert used_names == {"Sheet", "Sheet_2", "Sheet_3"}


def test_get_unique_filepath_returns_base_when_target_does_not_exist(tmp_path):
    filepath = get_unique_filepath(str(tmp_path), "Sheet")

    assert filepath == str(tmp_path / "Sheet.xlsx")


def test_get_unique_filepath_avoids_existing_files(tmp_path):
    (tmp_path / "Sheet.xlsx").write_text("", encoding="utf-8")
    (tmp_path / "Sheet_2.xlsx").write_text("", encoding="utf-8")

    filepath = get_unique_filepath(str(tmp_path), "Sheet")

    assert filepath == str(tmp_path / "Sheet_3.xlsx")


def test_get_unique_filepath_accepts_ext_with_or_without_dot(tmp_path):
    assert get_unique_filepath(str(tmp_path), "Book", ".xlsm") == str(tmp_path / "Book.xlsm")
    assert get_unique_filepath(str(tmp_path), "Book", "xlsm") == str(tmp_path / "Book.xlsm")
