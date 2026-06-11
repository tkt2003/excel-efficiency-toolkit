import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_to_tuple


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
RESULT_SHEET_NAME = "数据穿透结果"
RESULT_HEADERS = [
    "序号",
    "源文件名",
    "源文件路径",
    "目标 Sheet 名",
    "单元格地址",
    "取值",
    "是否为空",
    "是否错误值",
    "状态",
    "说明",
]
EXCEL_ERROR_VALUES = {
    "#VALUE!",
    "#DIV/0!",
    "#REF!",
    "#NAME?",
    "#N/A",
    "#NULL!",
    "#NUM!",
}


def is_supported_openpyxl_workbook(path: str) -> bool:
    return os.path.splitext(str(path))[1].lower() in SUPPORTED_EXTENSIONS


def is_office_temp_file(path: str) -> bool:
    return os.path.basename(str(path)).startswith("~$")


def build_unique_output_path(directory: str, base_name: str = "数据穿透结果", extension: str = ".xlsx") -> str:
    output_dir = os.path.abspath(str(directory))
    normalized_extension = extension if extension.startswith(".") else f".{extension}"
    candidate = os.path.join(output_dir, f"{base_name}{normalized_extension}")
    if not os.path.exists(candidate):
        return candidate

    counter = 2
    while True:
        candidate = os.path.join(output_dir, f"{base_name}_{counter}{normalized_extension}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def normalize_sheet_lookup_name(name: str) -> str:
    return str(name or "").strip().casefold()


def resolve_sheet_name(workbook, target_sheet_name: str) -> tuple[str | None, list[str], str]:
    available_sheet_names = list(workbook.sheetnames)
    normalized_target_name = normalize_sheet_lookup_name(target_sheet_name)
    if not normalized_target_name:
        return None, available_sheet_names, "目标 Sheet 名为空"

    for sheet_name in available_sheet_names:
        if normalize_sheet_lookup_name(sheet_name) == normalized_target_name:
            return sheet_name, available_sheet_names, "匹配同名 Sheet"

    if available_sheet_names:
        available_text = "、".join(available_sheet_names)
        return None, available_sheet_names, f"缺少同名 Sheet：{target_sheet_name}；可用 Sheet：{available_text}"
    return None, available_sheet_names, f"缺少同名 Sheet：{target_sheet_name}；源文件没有工作表"


def is_excel_error_value(value) -> bool:
    return isinstance(value, str) and value.strip() in EXCEL_ERROR_VALUES


def read_drill_cell_from_workbook(source_path: str, sheet_name: str, cell_address: str) -> dict:
    source_path = os.path.abspath(str(source_path))
    filename = os.path.basename(source_path)
    record = _new_record(source_path, sheet_name, cell_address)

    skip_message = _get_source_skip_message(source_path)
    if skip_message is not None:
        status = "失败" if skip_message == "文件不存在" else "跳过"
        return _finish_record(record, None, status, skip_message)

    value_workbook = None
    formula_workbook = None
    try:
        value_workbook = load_workbook(source_path, read_only=True, data_only=True)
        formula_workbook = load_workbook(source_path, read_only=True, data_only=False)

        actual_sheet_name, available_sheet_names, message = resolve_sheet_name(value_workbook, sheet_name)
        if actual_sheet_name is None:
            return _finish_record(record, None, "跳过", message)

        value_sheet = value_workbook[actual_sheet_name]
        formula_sheet = formula_workbook[actual_sheet_name]
        row_index, column_index = coordinate_to_tuple(str(cell_address).replace("$", "").strip())
        cell_missing = row_index > value_sheet.max_row or column_index > value_sheet.max_column
        value = _read_cell_value(value_sheet, row_index, column_index)
        formula_value = _read_cell_value(formula_sheet, row_index, column_index)

        if _is_formula_value(formula_value) and value is None:
            return _finish_record(
                record,
                value,
                "成功",
                "公式缓存为空，取值可能需要先打开源文件计算并保存",
            )
        if cell_missing:
            return _finish_record(record, None, "成功", "源文件缺少指定单元格，已按空值记录")

        return _finish_record(record, value, "成功", "读取成功")
    except Exception as e:
        return _finish_record(record, None, "失败", f"读取异常：{e}")
    finally:
        if value_workbook is not None:
            value_workbook.close()
        if formula_workbook is not None:
            formula_workbook.close()


def build_data_drill_records(source_paths: list[str], sheet_name: str, cell_address: str, logger=None) -> list[dict]:
    records = []
    for index, source_path in enumerate(source_paths, start=1):
        _log(logger, "info", f"正在读取源文件 {index}/{len(source_paths)}：{source_path}")
        record = read_drill_cell_from_workbook(source_path, sheet_name, cell_address)
        records.append(record)
        _log(logger, "info", f"{record['source_file_name']}：{record['status']}，{record['message']}")
    return records


def summarize_data_drill_records(records: list[dict]) -> dict:
    summary = {"success_count": 0, "skipped_count": 0, "failed_count": 0, "total_count": len(records)}
    for record in records:
        status = record.get("status")
        if status == "成功":
            summary["success_count"] += 1
        elif status == "跳过":
            summary["skipped_count"] += 1
        elif status == "失败":
            summary["failed_count"] += 1
    return summary


def write_data_drill_result_workbook(
    records: list[dict],
    output_path: str,
    source_sheet_name: str,
    cell_address: str,
) -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = RESULT_SHEET_NAME
    sheet.append(RESULT_HEADERS)

    for index, record in enumerate(records, start=1):
        value = record.get("value")
        sheet.append(
            [
                index,
                record.get("source_file_name", ""),
                record.get("source_file_path", ""),
                record.get("target_sheet_name", source_sheet_name),
                record.get("cell_address", cell_address),
                value,
                "是" if _is_empty_value(value) else "否",
                "是" if is_excel_error_value(value) else "否",
                record.get("status", ""),
                record.get("message", ""),
            ]
        )

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    _set_result_column_widths(sheet)
    workbook.save(output_path)
    return output_path


def _get_source_skip_message(source_path: str) -> str | None:
    filename = os.path.basename(source_path)
    if is_office_temp_file(source_path):
        return "临时文件已跳过"
    if not os.path.exists(source_path):
        return "文件不存在"

    extension = os.path.splitext(filename)[1].lower()
    if extension == ".xls":
        return "请另存为 xlsx/xlsm 后再处理"
    if extension not in SUPPORTED_EXTENSIONS:
        return "不支持的文件格式"
    return None


def _new_record(source_path: str, sheet_name: str, cell_address: str) -> dict:
    return {
        "source_file_name": os.path.basename(source_path),
        "source_file_path": source_path,
        "target_sheet_name": sheet_name,
        "cell_address": str(cell_address).replace("$", "").strip(),
        "value": None,
        "status": "",
        "message": "",
    }


def _finish_record(record: dict, value, status: str, message: str) -> dict:
    record["value"] = value
    record["is_empty"] = _is_empty_value(value)
    record["is_error_value"] = is_excel_error_value(value)
    record["status"] = status
    record["message"] = message
    return record


def _read_cell_value(sheet, row_index: int, column_index: int):
    for row_values in sheet.iter_rows(
        min_row=row_index,
        max_row=row_index,
        min_col=column_index,
        max_col=column_index,
        values_only=True,
    ):
        return row_values[0] if row_values else None
    return None


def _is_formula_value(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _is_empty_value(value) -> bool:
    return value is None or value == ""


def _set_result_column_widths(sheet) -> None:
    widths = {
        "A": 8,
        "B": 24,
        "C": 52,
        "D": 18,
        "E": 14,
        "F": 18,
        "G": 10,
        "H": 12,
        "I": 10,
        "J": 48,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width


def _log(logger, level: str, message: str) -> None:
    if logger is None:
        return
    log_func = getattr(logger, level, None)
    if log_func is None:
        return
    log_func(message)
