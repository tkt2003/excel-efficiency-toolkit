# -*- coding: utf-8 -*-
from src.excel_efficiency_toolkit.text_display import (
    char_display_width,
    compact_log_cell,
    display_width,
    fit_display_cell,
    format_display_row,
    middle_truncate_display,
    take_display_prefix,
    take_display_suffix,
)


def test_cjk_char_counts_as_double_width():
    assert char_display_width("中") == 2
    assert char_display_width("A") == 1
    assert display_width("中A文b") == 6


def test_take_display_prefix_and_suffix_respect_cjk_width():
    assert take_display_prefix("中文abc", 3) == "中"  # 再放"文"会超宽
    assert take_display_prefix("中文abc", 4) == "中文"
    assert take_display_suffix("abc中文", 3) == "文"
    assert take_display_suffix("abc中文", 5) == "c中文"


def test_middle_truncate_display_keeps_head_and_tail():
    result = middle_truncate_display("abcdefghij", 7)
    assert result == "ab...ij"
    assert display_width(result) == 7


def test_middle_truncate_display_tiny_width_returns_dots():
    assert middle_truncate_display("abcdef", 3) == "..."
    assert middle_truncate_display("abcdef", 2) == ".."
    assert middle_truncate_display("abcdef", 0) == ""


def test_fit_display_cell_pads_to_width():
    cell = fit_display_cell("ab", 6)
    assert cell == "ab    "
    assert display_width(cell) == 6


def test_fit_display_cell_truncates_long_text():
    cell = fit_display_cell("a" * 30, 10)
    assert display_width(cell) == 10
    assert "..." in cell


def test_compact_log_cell_flattens_newlines_and_truncates():
    assert compact_log_cell("a\r\nb\nc") == "a  b c"
    assert len(compact_log_cell("x" * 200)) == 80


def test_format_display_row_aligns_columns():
    columns = [("列一", 6), ("列二", 4)]
    row = format_display_row(["ab", "中文超长文本"], columns)
    # fit("ab", 6) -> "ab    "；fit("中文超长文本", 4) -> "... "（截断后补齐）
    assert row == "ab      ... "
