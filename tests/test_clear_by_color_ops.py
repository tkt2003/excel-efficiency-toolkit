from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from src.excel_efficiency_toolkit.clear_by_color_ops import (
    build_clear_by_color_backup_path,
    build_clear_detail_log_record,
    build_clear_ranges_from_cells,
    build_clear_summary_log_record,
    clear_loaded_workbook_cells,
    get_load_workbook_options_for_color_clear,
    is_supported_openpyxl_color_workbook,
    scan_workbook_color_matches,
    should_keep_vba_for_workbook,
    write_clear_by_color_log_workbook,
)
from src.excel_efficiency_toolkit.color_sum_ops import get_openpyxl_fill_key


def _build_scan_workbook(path: Path):
    yellow = PatternFill(fill_type="solid", fgColor="FFFF00")
    green = PatternFill(fill_type="solid", fgColor="00FF00")

    workbook = Workbook()
    visible = workbook.active
    visible.title = "可见表1"
    visible["A1"] = "样本"
    visible["A1"].fill = yellow
    visible["B1"] = 1
    visible["B1"].fill = yellow
    visible["C1"] = 2
    visible["C1"].fill = green
    visible["A2"] = 3
    visible["A2"].fill = yellow

    visible2 = workbook.create_sheet("可见表2")
    visible2["D4"] = 4
    visible2["D4"].fill = yellow

    hidden = workbook.create_sheet("隐藏表")
    hidden.sheet_state = "hidden"
    hidden["A1"] = 5
    hidden["A1"].fill = yellow

    workbook.save(path)
    workbook.close()


def _build_clear_workbook(path: Path):
    yellow = PatternFill(fill_type="solid", fgColor="FFFF00")

    workbook = Workbook()
    sheet1 = workbook.active
    sheet1.title = "Sheet1"
    sheet1["A1"] = 100
    sheet1["A1"].fill = yellow
    sheet1["B1"] = "=SUM(1,2)"
    sheet1["B1"].fill = yellow
    sheet1["C1"] = "保留"
    sheet1["C1"].fill = PatternFill(fill_type="solid", fgColor="00FF00")

    sheet2 = workbook.create_sheet("Sheet2")
    sheet2["D4"] = 200
    sheet2["D4"].fill = yellow
    sheet2["E4"] = "=1+1"
    sheet2["E4"].fill = yellow
    sheet2.merge_cells("G1:H1")
    sheet2["G1"] = "合并值"
    sheet2["G1"].fill = yellow

    workbook.save(path)
    workbook.close()


def test_is_supported_openpyxl_color_workbook_accepts_xlsx_xlsm_and_rejects_xls():
    assert is_supported_openpyxl_color_workbook("book.xlsx")
    assert is_supported_openpyxl_color_workbook("book.xlsm")
    assert not is_supported_openpyxl_color_workbook("book.xls")


def test_build_clear_by_color_backup_path_does_not_overwrite(tmp_path):
    target_path = tmp_path / "目标.xlsx"
    target_path.write_text("a", encoding="utf-8")

    first = build_clear_by_color_backup_path(str(target_path), timestamp="20260612_120000")
    Path(first).write_text("b", encoding="utf-8")
    second = build_clear_by_color_backup_path(str(target_path), timestamp="20260612_120000")

    assert first != second


def test_scan_workbook_color_matches_finds_same_color_and_skips_hidden_sheet(tmp_path):
    path = tmp_path / "扫描.xlsx"
    _build_scan_workbook(path)

    workbook = load_workbook(path)
    selected_color_key = get_openpyxl_fill_key(workbook["可见表1"]["A1"])
    workbook.close()

    plan = scan_workbook_color_matches(str(path), selected_color_key)

    assert plan["matched_sheet_count"] == 2
    assert plan["matched_cell_count"] == 4
    names = [item["sheet_name"] for item in plan["sheet_plans"]]
    assert "隐藏表" not in names


def test_clear_loaded_workbook_cells_clears_values_and_formulas_but_keeps_fill(tmp_path):
    path = tmp_path / "清空.xlsx"
    _build_clear_workbook(path)
    workbook = load_workbook(path)

    result = clear_loaded_workbook_cells(
        workbook,
        str(path),
        [{"sheet_name": "Sheet1", "cells": ["A1", "B1"]}],
    )
    workbook.save(path)
    workbook.close()

    reloaded = load_workbook(path)
    assert reloaded["Sheet1"]["A1"].value is None
    assert reloaded["Sheet1"]["B1"].value is None
    assert reloaded["Sheet1"]["A1"].fill.fgColor.rgb == "00FFFF00"
    assert reloaded["Sheet1"]["B1"].fill.fgColor.rgb == "00FFFF00"
    assert result["cleared_cell_count"] == 2
    reloaded.close()


def test_clear_loaded_workbook_cells_supports_multiple_sheets(tmp_path):
    path = tmp_path / "多sheet.xlsx"
    _build_clear_workbook(path)
    workbook = load_workbook(path)

    clear_loaded_workbook_cells(
        workbook,
        str(path),
        [
            {"sheet_name": "Sheet1", "cells": ["A1"]},
            {"sheet_name": "Sheet2", "cells": ["D4", "E4"]},
        ],
    )
    workbook.save(path)
    workbook.close()

    reloaded = load_workbook(path)
    assert reloaded["Sheet1"]["A1"].value is None
    assert reloaded["Sheet2"]["D4"].value is None
    assert reloaded["Sheet2"]["E4"].value is None
    reloaded.close()


def test_clear_loaded_workbook_cells_skips_non_top_left_merged_cell_without_crashing(tmp_path):
    path = tmp_path / "合并.xlsx"
    _build_clear_workbook(path)
    workbook = load_workbook(path)
    detail_records = []

    result = clear_loaded_workbook_cells(
        workbook,
        str(path),
        [{"sheet_name": "Sheet2", "cells": ["H1"]}],
        detail_records=detail_records,
    )

    assert result["cleared_cell_count"] == 0
    assert result["skipped_merged_cell_count"] == 1
    assert detail_records[0]["状态"] == "跳过"
    workbook.close()


def test_build_clear_ranges_from_cells_groups_same_row_contiguous_columns():
    ranges = build_clear_ranges_from_cells([(2, 1), (2, 2), (2, 4), (3, 1)])
    assert ranges == ["A2:B2", "D2", "A3"]


def test_build_clear_summary_log_record_structure_is_correct():
    record = build_clear_summary_log_record(
        "D:/a.xlsx",
        2,
        10,
        8,
        "D:/backup.xlsx",
        "成功",
        "说明",
    )
    assert record["匹配工作表数量"] == 2
    assert record["清空单元格数量"] == 8


def test_build_clear_detail_log_record_structure_is_correct():
    record = build_clear_detail_log_record("D:/a.xlsx", "Sheet1", "C3", "跳过", "合并单元格")
    assert record["工作表名"] == "Sheet1"
    assert record["单元格地址"] == "C3"


def test_write_clear_by_color_log_workbook_generates_file(tmp_path):
    log_path = write_clear_by_color_log_workbook(
        str(tmp_path),
        [build_clear_summary_log_record("D:/a.xlsx", 1, 2, 2, "D:/bak.xlsx", "成功", "完成")],
        [build_clear_detail_log_record("D:/a.xlsx", "Sheet1", "A1", "成功", "")],
        timestamp="20260612_120000",
    )

    assert Path(log_path).exists()
    workbook = load_workbook(log_path)
    assert workbook.sheetnames == ["处理汇总", "处理明细"]
    workbook.close()


def test_get_load_workbook_options_for_color_clear_uses_keep_vba_for_xlsm():
    assert should_keep_vba_for_workbook("book.xlsm") is True
    assert should_keep_vba_for_workbook("book.xlsx") is False
    assert get_load_workbook_options_for_color_clear("book.xlsm") == {"keep_vba": True}
