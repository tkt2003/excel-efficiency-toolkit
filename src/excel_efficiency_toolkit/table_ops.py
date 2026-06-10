import os

from .excel_com import get_active_excel
from .name_utils import get_safe_sheet_name, get_unique_sheet_name


XL_SHEET_VISIBLE = -1


def normalize_header_row(values: list[object]) -> list[str]:
    return ["" if value is None else str(value).strip() for value in values]


def validate_row_numbers(header_row: int, data_start_row: int) -> None:
    if header_row < 1:
        raise ValueError("表头行号必须大于等于 1。")
    if data_start_row <= header_row:
        raise ValueError("数据起始行号必须大于表头行号。")


def parse_column_index(value: str) -> int:
    column = str(value).strip()
    if not column:
        raise ValueError("列不能为空。")

    if column.isdigit():
        index = int(column)
        if index < 1:
            raise ValueError("列号必须大于等于 1。")
        return index

    column = column.upper()
    if not column.isalpha() or not column.isascii():
        raise ValueError("列必须是 Excel 列字母或正整数。")

    index = 0
    for char in column:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def build_split_targets(values: list[object]) -> list[str]:
    targets: list[str] = []
    seen: set[str] = set()

    for value in values:
        target = "空白" if value is None else str(value).strip()
        if not target:
            target = "空白"
        if target not in seen:
            seen.add(target)
            targets.append(target)

    return targets


def resolve_source_sheet_name(source_sheet_name: str | None, available_names: list[str]) -> str:
    if not available_names:
        raise ValueError("工作簿中没有可用工作表。")

    requested_name = "" if source_sheet_name is None else str(source_sheet_name).strip()
    if not requested_name:
        if len(available_names) == 1:
            return available_names[0]
        raise ValueError("工作簿包含多个工作表，请输入源 sheet 名。")

    lower_names = {name.lower(): name for name in available_names}
    matched_name = lower_names.get(requested_name.lower())
    if matched_name:
        return matched_name

    raise ValueError(f"未找到源 sheet：{requested_name}。")


def _normalize_split_target(value: object) -> str:
    target = "空白" if value is None else str(value).strip()
    return target or "空白"


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


def _get_workbook_sheet_names(workbook) -> set[str]:
    return {sheet.Name for sheet in workbook.Worksheets}


def _get_or_open_workbook(source_path: str, logger=None):
    if not source_path:
        raise ValueError("请选择源 Excel 工作簿。")

    abs_path = os.path.abspath(source_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"源工作簿不存在：{abs_path}")

    _log(logger, "info", "尝试连接当前运行的 Excel 实例...")
    excel = get_active_excel()
    if not excel:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True

    normalized_path = os.path.normcase(abs_path)
    for workbook in excel.Workbooks:
        try:
            if os.path.normcase(os.path.abspath(workbook.FullName)) == normalized_path:
                _log(logger, "info", f"复用已打开工作簿：{workbook.Name}")
                return workbook
        except Exception:
            continue

    try:
        _log(logger, "info", f"正在打开工作簿：{abs_path}")
        return excel.Workbooks.Open(abs_path, UpdateLinks=0)
    except Exception as e:
        raise RuntimeError(f"打开源工作簿失败，请确认文件未损坏且未被其他 Excel 实例占用：{e}") from e


def _get_last_used_row_and_col(sheet) -> tuple[int, int]:
    used_range = sheet.UsedRange
    last_row = used_range.Row + used_range.Rows.Count - 1
    last_col = used_range.Column + used_range.Columns.Count - 1
    return last_row, last_col


def _normalize_2d_range_values(values, row_count: int, col_count: int) -> list[list[object]]:
    if row_count == 1 and col_count == 1:
        return [[values]]
    if row_count == 1:
        return [list(values[0] if values and isinstance(values[0], tuple) else values)]
    if col_count == 1:
        return [[row[0] if isinstance(row, tuple) else row] for row in values]
    return [list(row) for row in values]


def _read_range_values(sheet, start_row: int, start_col: int, end_row: int, end_col: int) -> list[list[object]]:
    row_count = end_row - start_row + 1
    col_count = end_col - start_col + 1
    values = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col)).Value
    return _normalize_2d_range_values(values, row_count, col_count)


def _write_range_values(sheet, start_row: int, start_col: int, values: list[list[object]]) -> None:
    if not values:
        return
    row_count = len(values)
    col_count = len(values[0])
    sheet.Range(
        sheet.Cells(start_row, start_col),
        sheet.Cells(start_row + row_count - 1, start_col + col_count - 1),
    ).Value = tuple(tuple(row) for row in values)


def _merge_workbook_to_new_sheet(
    workbook,
    header_row: int = 1,
    data_start_row: int = 2,
    result_sheet_name: str = "合并结果",
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)

    source_sheets = [sheet for sheet in workbook.Worksheets if sheet.Visible == XL_SHEET_VISIBLE]
    if not source_sheets:
        raise RuntimeError("当前工作簿没有可见工作表。")

    first_source = None
    header_values: list[str] = []
    header_col_count = 0
    sheet_infos = []

    for sheet in source_sheets:
        last_row, last_col = _get_last_used_row_and_col(sheet)
        if last_row < data_start_row or last_col < 1:
            _log(logger, "info", f"跳过无有效数据的工作表：{sheet.Name}")
            continue
        if first_source is None:
            first_source = sheet
            header_col_count = last_col
            header_values = normalize_header_row(_read_range_values(sheet, header_row, 1, header_row, last_col)[0])
        sheet_infos.append((sheet, last_row))

    if first_source is None:
        raise RuntimeError("没有找到包含有效数据的可见工作表。")

    participating_names = [sheet.Name for sheet, _ in sheet_infos]
    _log(logger, "info", f"参与合并的工作表：{', '.join(participating_names)}")

    existing_names = _get_workbook_sheet_names(workbook)
    safe_name = get_safe_sheet_name(result_sheet_name, fallback="合并结果")
    unique_name = get_unique_sheet_name(safe_name, existing_names)

    result_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    result_sheet.Name = unique_name
    _write_range_values(result_sheet, 1, 1, [["来源工作表", *header_values]])

    next_row = 2
    source_sheet_count = 0
    appended_row_count = 0

    for sheet, last_row in sheet_infos:
        source_values = _read_range_values(sheet, data_start_row, 1, last_row, header_col_count)
        output_values = [[sheet.Name, *row] for row in source_values]
        _write_range_values(result_sheet, next_row, 1, output_values)
        copied_rows = len(output_values)
        next_row += copied_rows
        source_sheet_count += 1
        appended_row_count += copied_rows
        _log(logger, "info", f"已合并工作表：{sheet.Name}，追加 {copied_rows} 行。")

    return {
        "workbook_name": workbook.Name,
        "result_sheet_name": unique_name,
        "source_sheet_count": source_sheet_count,
        "appended_row_count": appended_row_count,
    }


def merge_visible_sheets_to_new_sheet(
    header_row: int = 1,
    data_start_row: int = 2,
    result_sheet_name: str = "合并结果",
    logger=None,
) -> dict:
    _log(logger, "info", "尝试连接当前运行的 Excel 实例...")
    excel = get_active_excel()
    if not excel:
        raise RuntimeError("未检测到正在运行的 Excel。请先打开 Excel。")

    workbook = excel.ActiveWorkbook
    if not workbook:
        raise RuntimeError("没有打开的工作簿。请先打开或新建一个 Excel 文件。")

    return _merge_workbook_to_new_sheet(
        workbook,
        header_row=header_row,
        data_start_row=data_start_row,
        result_sheet_name=result_sheet_name,
        logger=logger,
    )


def merge_workbook_sheets_to_new_sheet(
    source_path: str,
    header_row: int = 1,
    data_start_row: int = 2,
    result_sheet_name: str = "合并结果",
    logger=None,
) -> dict:
    workbook = _get_or_open_workbook(source_path, logger)
    return _merge_workbook_to_new_sheet(
        workbook,
        header_row=header_row,
        data_start_row=data_start_row,
        result_sheet_name=result_sheet_name,
        logger=logger,
    )


def _split_workbook_sheet_by_column(
    workbook,
    source_sheet,
    column_input: str,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    column_index = parse_column_index(column_input)

    _log(logger, "info", f"准备拆分工作表：{source_sheet.Name}，拆分列：{column_input.strip()}（第 {column_index} 列）")

    last_row, last_col = _get_last_used_row_and_col(source_sheet)
    if last_row < data_start_row:
        raise RuntimeError("当前活动工作表没有可拆分的数据。")
    if column_index > last_col:
        raise RuntimeError("拆分列超出当前工作表的有效区域。")

    header_values = _read_range_values(source_sheet, header_row, 1, header_row, last_col)
    data_values = _read_range_values(source_sheet, data_start_row, 1, last_row, last_col)
    split_values = [row[column_index - 1] for row in data_values]
    targets = build_split_targets(split_values)

    grouped_rows = {target: [] for target in targets}
    for row in data_values:
        grouped_rows[_normalize_split_target(row[column_index - 1])].append(row)

    existing_names = _get_workbook_sheet_names(workbook)
    created_sheet_count = 0
    copied_row_count = 0

    for target in targets:
        safe_name = get_safe_sheet_name(target, fallback="空白")
        unique_name = get_unique_sheet_name(safe_name, existing_names)
        new_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        new_sheet.Name = unique_name

        _write_range_values(new_sheet, 1, 1, header_values)
        rows = grouped_rows[target]
        _write_range_values(new_sheet, 2, 1, rows)

        created_sheet_count += 1
        copied_row_count += len(rows)
        _log(logger, "info", f"已生成工作表：{unique_name}，复制 {len(rows)} 行。")

    _log(logger, "info", f"拆分完成，共生成 {created_sheet_count} 个工作表，复制 {copied_row_count} 行。")
    return {
        "workbook_name": workbook.Name,
        "source_sheet_name": source_sheet.Name,
        "created_sheet_count": created_sheet_count,
        "copied_row_count": copied_row_count,
    }


def split_active_sheet_by_column(
    column_input: str,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    _log(logger, "info", "尝试连接当前运行的 Excel 实例...")
    excel = get_active_excel()
    if not excel:
        raise RuntimeError("未检测到正在运行的 Excel。请先打开 Excel。")

    workbook = excel.ActiveWorkbook
    if not workbook:
        raise RuntimeError("没有打开的工作簿。请先打开或新建一个 Excel 文件。")

    source_sheet = excel.ActiveSheet
    if not source_sheet:
        raise RuntimeError("没有活动的工作表。")

    return _split_workbook_sheet_by_column(
        workbook,
        source_sheet,
        column_input=column_input,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )


def split_workbook_sheet_by_column(
    source_path: str,
    source_sheet_name: str | None,
    column_input: str,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    workbook = _get_or_open_workbook(source_path, logger)
    sheet_names = [sheet.Name for sheet in workbook.Worksheets]
    resolved_name = resolve_source_sheet_name(source_sheet_name, sheet_names)
    source_sheet = workbook.Worksheets(resolved_name)

    return _split_workbook_sheet_by_column(
        workbook,
        source_sheet,
        column_input=column_input,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )
