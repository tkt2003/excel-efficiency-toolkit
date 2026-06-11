from openpyxl import Workbook

from src.excel_efficiency_toolkit.report_generate_ops import (
    DEFAULT_CHECKLIST_SHEET_NAME,
    build_unique_output_path,
    read_report_checklist,
    sanitize_output_stem,
    summarize_report_records,
)


def test_read_report_checklist_reads_fixed_columns_and_skips_empty_path(tmp_path):
    template_path = tmp_path / "template.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = DEFAULT_CHECKLIST_SHEET_NAME
    sheet.append(["新 TB / 新链接文件路径", "输出文件名主体", "公司名称", "报表类型"])
    sheet.append([str(tmp_path / "new_tb_1.xlsx"), "子公司1", "公司1", "月报"])
    sheet.append(["", "空路径", "公司空", "月报"])
    sheet.append([str(tmp_path / "new_tb_2.xlsx"), "子公司2", "公司2", "年报"])
    workbook.save(template_path)
    workbook.close()

    records = read_report_checklist(str(template_path))

    assert len(records) == 2
    assert records[0].row_number == 2
    assert records[0].tb_link_path.endswith("new_tb_1.xlsx")
    assert records[0].output_name_raw == "子公司1"
    assert records[0].company_name == "公司1"
    assert records[0].report_type == "月报"
    assert records[1].row_number == 4


def test_read_report_checklist_requires_sheet(tmp_path):
    template_path = tmp_path / "template.xlsx"
    workbook = Workbook()
    workbook.save(template_path)
    workbook.close()

    try:
        read_report_checklist(str(template_path), "missing")
    except ValueError as error:
        assert "缺少清单 sheet" in str(error)
    else:
        raise AssertionError("expected missing sheet error")


def test_sanitize_output_stem_removes_illegal_chars_and_extension():
    assert sanitize_output_stem(' 子公司:1/报表.xlsx ', row_number=2) == "子公司_1_报表"
    assert sanitize_output_stem("CONSOL", row_number=3) == "CONSOL"


def test_sanitize_output_stem_falls_back_to_row_number():
    assert sanitize_output_stem("   ", row_number=8) == "报表_第8行"


def test_build_unique_output_path_adds_sequence_without_overwrite(tmp_path):
    first_path = tmp_path / "子公司1.xlsx"
    second_path = tmp_path / "子公司1_2.xlsx"
    first_path.write_text("exists", encoding="utf-8")
    second_path.write_text("exists", encoding="utf-8")

    output_path = build_unique_output_path(str(tmp_path), "子公司1", ".xlsx")

    assert output_path == str(tmp_path / "子公司1_3.xlsx")
    assert first_path.read_text(encoding="utf-8") == "exists"
    assert second_path.read_text(encoding="utf-8") == "exists"


def test_summarize_report_records_counts_statuses():
    records = [
        {"status": "成功"},
        {"status": "成功"},
        {"status": "跳过"},
        {"status": "失败"},
    ]

    assert summarize_report_records(records) == {
        "success_count": 2,
        "skipped_count": 1,
        "failed_count": 1,
        "total_count": 4,
    }
