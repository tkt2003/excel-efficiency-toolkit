from __future__ import annotations

import os
from datetime import datetime

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries

from .rename_sheet_ops import is_sheet_hidden_by_visible_value


RESULT_SHEET_BASE_NAME = "数据穿透查询结果"
EXCEL_ERROR_TEXTS = {
    "#VALUE!",
    "#DIV/0!",
    "#REF!",
    "#NAME?",
    "#N/A",
    "#NULL!",
    "#NUM!",
}
EXCEL_ERROR_CODE_MAP = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
}


def normalize_range_address(range_address: str) -> str:
    return str(range_address or "").replace("$", "").strip().upper()


def is_multi_area_range(range_address: str) -> bool:
    normalized = normalize_range_address(range_address)
    return "," in normalized


def expand_range_addresses(range_address: str) -> list[str]:
    normalized = normalize_range_address(range_address)
    if not normalized:
        raise ValueError("选区地址为空")
    if is_multi_area_range(normalized):
        raise ValueError("不支持多区域选区")

    min_col, min_row, max_col, max_row = range_boundaries(normalized)
    addresses = []
    for row_index in range(min_row, max_row + 1):
        for column_index in range(min_col, max_col + 1):
            addresses.append(f"{get_column_letter(column_index)}{row_index}")
    return addresses


def build_unique_result_sheet_name(
    existing_sheet_names: list[str],
    base_name: str = RESULT_SHEET_BASE_NAME,
) -> str:
    existing_names = {str(name or "").strip() for name in existing_sheet_names}
    if base_name not in existing_names:
        return base_name

    counter = 2
    while True:
        candidate = f"{base_name}_{counter}"
        if candidate not in existing_names:
            return candidate
        counter += 1


def should_skip_history_result_sheet(
    sheet_name: str,
    base_name: str = RESULT_SHEET_BASE_NAME,
) -> bool:
    return str(sheet_name or "").strip().startswith(base_name)


def stringify_excel_error_value(value):
    if isinstance(value, str):
        text = value.strip()
        if text in EXCEL_ERROR_TEXTS:
            return text
        return value

    if isinstance(value, (int, float)):
        error_text = EXCEL_ERROR_CODE_MAP.get(int(value))
        if error_text is not None:
            return error_text
    return value


def build_single_file_result_headers(range_address: str) -> list[str]:
    return ["工作表名", "工作表可见状态", *expand_range_addresses(range_address)]


def build_multi_file_result_headers(range_address: str) -> list[str]:
    return ["源文件名", "源文件路径", "工作表名", *expand_range_addresses(range_address), "状态", "说明"]


def build_single_file_result_row(
    sheet_name: str,
    visible_status: str,
    values_by_address: dict[str, object],
    range_address: str,
) -> list[object]:
    addresses = expand_range_addresses(range_address)
    return [
        sheet_name,
        visible_status,
        *[stringify_excel_error_value(values_by_address.get(address)) for address in addresses],
    ]


def build_multi_file_result_row(record: dict, range_address: str) -> list[object]:
    addresses = expand_range_addresses(range_address)
    values_by_address = record.get("values_by_address", {})
    return [
        record.get("source_file_name", ""),
        record.get("source_file_path", ""),
        record.get("target_sheet_name", ""),
        *[stringify_excel_error_value(values_by_address.get(address)) for address in addresses],
        record.get("status", ""),
        record.get("message", ""),
    ]


def write_single_file_result_sheet(
    sheet,
    base_sheet_name: str,
    range_address: str,
    created_at_text: str,
    records: list[dict],
) -> None:
    sheet.cell(row=1, column=1).value = (
        f"基准工作表：{base_sheet_name}    基准选区：{normalize_range_address(range_address)}    生成时间：{created_at_text}"
    )

    headers = build_single_file_result_headers(range_address)
    for column_index, header in enumerate(headers, start=1):
        sheet.cell(row=3, column=column_index).value = header

    for row_offset, record in enumerate(records, start=4):
        row_values = build_single_file_result_row(
            sheet_name=record.get("sheet_name", ""),
            visible_status=record.get("visible_status", ""),
            values_by_address=record.get("values_by_address", {}),
            range_address=range_address,
        )
        for column_index, value in enumerate(row_values, start=1):
            sheet.cell(row=row_offset, column=column_index).value = value


def write_single_workbook_drill_result_to_com_sheet(
    sheet,
    base_sheet_name: str,
    range_address: str,
    created_at_text: str,
    records: list[dict],
) -> None:
    sheet.Cells(1, 1).Value = (
        f"基准工作表：{base_sheet_name}    基准选区：{normalize_range_address(range_address)}    生成时间：{created_at_text}"
    )

    headers = build_single_file_result_headers(range_address)
    for column_index, header in enumerate(headers, start=1):
        sheet.Cells(3, column_index).Value = header

    for row_offset, record in enumerate(records, start=4):
        row_values = build_single_file_result_row(
            sheet_name=record.get("sheet_name", ""),
            visible_status=record.get("visible_status", ""),
            values_by_address=record.get("values_by_address", {}),
            range_address=range_address,
        )
        for column_index, value in enumerate(row_values, start=1):
            sheet.Cells(row_offset, column_index).Value = value


def write_multi_file_result_sheet(
    sheet,
    records: list[dict],
    range_address: str,
) -> None:
    headers = build_multi_file_result_headers(range_address)
    for column_index, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=column_index).value = header

    for row_offset, record in enumerate(records, start=2):
        row_values = build_multi_file_result_row(record, range_address)
        for column_index, value in enumerate(row_values, start=1):
            sheet.cell(row=row_offset, column=column_index).value = value


def range_value_to_address_map(range_address: str, raw_value) -> dict[str, object]:
    addresses = expand_range_addresses(range_address)
    if len(addresses) == 1:
        return {addresses[0]: stringify_excel_error_value(raw_value)}

    rows = _normalize_range_values_matrix(raw_value, range_address)
    values_by_address = {}
    address_index = 0
    for row_values in rows:
        for value in row_values:
            values_by_address[addresses[address_index]] = stringify_excel_error_value(value)
            address_index += 1
    return values_by_address


def is_excel_error_text(value) -> bool:
    return isinstance(value, str) and value.strip() in EXCEL_ERROR_TEXTS


def _normalize_range_values_matrix(raw_value, range_address: str) -> list[list[object]]:
    normalized = normalize_range_address(range_address)
    min_col, min_row, max_col, max_row = range_boundaries(normalized)
    width = max_col - min_col + 1
    height = max_row - min_row + 1

    if raw_value is None:
        return [[None for _ in range(width)] for _ in range(height)]

    if isinstance(raw_value, tuple):
        rows = []
        for row_values in raw_value:
            if isinstance(row_values, tuple):
                rows.append(list(row_values))
            else:
                rows.append([row_values])
        return _pad_matrix(rows, width, height)

    return _pad_matrix([[raw_value]], width, height)


def _pad_matrix(rows: list[list[object]], width: int, height: int) -> list[list[object]]:
    normalized_rows = []
    for row_index in range(height):
        source_row = rows[row_index] if row_index < len(rows) else []
        normalized_row = list(source_row[:width])
        if len(normalized_row) < width:
            normalized_row.extend([None] * (width - len(normalized_row)))
        normalized_rows.append(normalized_row)
    return normalized_rows


# ---------------------------------------------------------------------------
# Excel COM 活动会话读取（原 app.py 数据穿透逻辑）
# ---------------------------------------------------------------------------
def format_drill_context_steps(steps: list[str]) -> str:
    return "、".join(steps) if steps else "尚未连接 Excel"


def get_selection_range_address(excel) -> str:
    """读取当前选区地址并规范化；多区域选区抛错。"""
    try:
        selection = excel.Selection
    except Exception as e:
        raise RuntimeError(f"已连接 Excel，但无法读取当前 Selection：{e}") from e

    if selection is None:
        raise RuntimeError("已连接 Excel，但当前没有可识别的选区。")
    try:
        areas = selection.Areas.Count
    except Exception:
        areas = 1
    if areas and int(areas) > 1:
        raise RuntimeError("不支持多区域选区，请只选择一个连续区域后重试。")

    try:
        address_member = selection.Address
        if callable(address_member):
            address = address_member(False, False)
        else:
            address = address_member
    except Exception:
        address = selection.Address(False, False)
    normalized_address = normalize_range_address(address)
    if is_multi_area_range(normalized_address):
        raise RuntimeError("不支持多区域选区，请只选择一个连续区域后重试。")
    return normalized_address


def get_active_drill_context(require_saved_workbook: bool) -> dict:
    """在独立 COM 会话中读取活动工作簿/工作表/选区上下文。"""
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

    pythoncom.CoInitialize()
    steps = []
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception as e:
            raise RuntimeError(
                "未检测到正在运行的 Excel，请先打开目标/合并工作簿后重试。"
                "当前步骤：尚未连接 Excel。"
            ) from e
        steps.append("已连接 Excel")

        try:
            workbook = excel.ActiveWorkbook
        except Exception as e:
            raise RuntimeError(
                "Excel 中没有活动工作簿，请先打开目标/合并工作簿后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}。"
            ) from e
        if workbook is None:
            raise RuntimeError(
                "Excel 中没有活动工作簿，请先打开目标/合并工作簿后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}。"
            )
        steps.append("已取得 ActiveWorkbook")

        workbook_name = str(workbook.Name)
        workbook_dir = str(workbook.Path or "").strip()
        if require_saved_workbook and not workbook_dir:
            raise RuntimeError(
                "当前活动工作簿尚未保存，无法确定结果文件输出目录。请先保存当前工作簿后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}。"
            )
        workbook_path = str(workbook.FullName or "").strip()
        if workbook_path and not os.path.dirname(workbook_path) and workbook_dir:
            workbook_path = os.path.join(workbook_dir, workbook_name)
        workbook_path_text = workbook_path or workbook_name
        steps.append("已取得工作簿保存路径")

        try:
            active_sheet = excel.ActiveSheet
            sheet_name = str(active_sheet.Name)
        except Exception as e:
            raise RuntimeError(
                "无法读取当前活动 Sheet，请先切换到目标工作表后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}。"
            ) from e
        if not sheet_name:
            raise RuntimeError(
                "无法读取当前活动 Sheet，请先切换到目标工作表后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}。"
            )
        steps.append("已取得 ActiveSheet")

        range_address = get_selection_range_address(excel)
        if not range_address:
            raise RuntimeError(
                "没有有效选区，请先选中一个需要追查的单元格或区域后重试。"
                f"当前步骤：{format_drill_context_steps(steps)}，但未取得 ActiveCell 或可用 Selection。"
            )
        steps.append("已取得 Selection")

        return {
            "workbook_name": workbook_name,
            "workbook_path": workbook_path,
            "workbook_path_text": workbook_path_text,
            "output_dir": workbook_dir,
            "workbook_dir": workbook_dir,
            "sheet_name": sheet_name,
            "range_address": range_address,
        }
    finally:
        pythoncom.CoUninitialize()


def read_range_values_from_com_sheet(sheet, range_address: str, logger=None) -> dict:
    """读取 COM 工作表指定区域的值映射；读取异常时每个地址填入异常说明。"""
    try:
        raw_value = sheet.Range(range_address).Value
        return range_value_to_address_map(range_address, raw_value)
    except Exception as e:
        message = f"读取异常：{e}"
        if logger is not None:
            logger.error(f"{sheet.Name}!{range_address} 读取失败：{message}")
        return {address: message for address in expand_range_addresses(range_address)}


def execute_single_workbook_drill(excel, logger=None) -> dict:
    """对活动工作簿执行单文件数据穿透：遍历可见工作表读取同选区值并写入结果工作表。

    返回结果摘要 dict，结果提示对话框由调用方负责。工作簿不会被自动保存。
    """
    workbook = excel.ActiveWorkbook
    if workbook is None:
        raise RuntimeError("Excel 中没有活动工作簿，请先打开目标工作簿后重试。")

    active_sheet = excel.ActiveSheet
    if active_sheet is None or not str(active_sheet.Name):
        raise RuntimeError("无法读取当前活动 Sheet，请先切换到目标工作表后重试。")

    range_address = get_selection_range_address(excel)
    workbook_name = str(workbook.Name)
    workbook_path = str(workbook.FullName or "").strip()
    workbook_path_text = workbook_path or workbook_name
    sheet_name = str(active_sheet.Name)

    if logger is not None:
        logger.info(f"当前工作簿：{workbook_path_text}")
        logger.info(f"当前工作表：{sheet_name}")
        logger.info(f"当前选区：{range_address}")

    records = []
    visible_sheet_count = 0
    for sheet in workbook.Worksheets:
        if is_sheet_hidden_by_visible_value(sheet.Visible):
            continue
        if should_skip_history_result_sheet(str(sheet.Name), RESULT_SHEET_BASE_NAME):
            continue

        visible_sheet_count += 1
        values_by_address = read_range_values_from_com_sheet(sheet, range_address, logger=logger)
        records.append(
            {
                "sheet_name": str(sheet.Name),
                "visible_status": "可见",
                "values_by_address": values_by_address,
            }
        )

    result_sheet_name = build_unique_result_sheet_name([str(sheet.Name) for sheet in workbook.Worksheets])
    result_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    result_sheet.Name = result_sheet_name
    write_single_workbook_drill_result_to_com_sheet(
        sheet=result_sheet,
        base_sheet_name=sheet_name,
        range_address=range_address,
        created_at_text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        records=records,
    )

    if logger is not None:
        logger.info(f"读取工作表数量：{visible_sheet_count}")
        logger.info(f"结果工作表：{result_sheet_name}")
        logger.info("数据穿透查询（单文件）完成。")
        logger.info("当前活动工作簿已新增结果工作表，但未自动保存。")

    return {
        "workbook_name": workbook_name,
        "workbook_path_text": workbook_path_text,
        "sheet_name": sheet_name,
        "range_address": range_address,
        "visible_sheet_count": visible_sheet_count,
        "result_sheet_name": result_sheet_name,
    }
