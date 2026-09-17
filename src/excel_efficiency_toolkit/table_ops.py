from dataclasses import dataclass
import os

from .excel_com import get_active_excel
from .export_ops import EXCEL_FILE_FORMAT_XLSX
from .name_utils import get_safe_sheet_name, get_unique_sheet_name


XL_SHEET_VISIBLE = -1
_INVALID_SPLIT_FILENAME_CHARS = frozenset('<>:"/\\|?*')
_WINDOWS_RESERVED_FILE_NAMES = frozenset(
    {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{index}" for index in range(1, 10)),
        *(f"LPT{index}" for index in range(1, 10)),
    }
)
_MAX_SPLIT_FILENAME_BASE_LENGTH = 250


def normalize_header_row(values: list[object]) -> list[str]:
    return ["" if value is None else str(value).strip() for value in values]


def validate_row_numbers(header_row: int, data_start_row: int) -> None:
    if header_row < 1:
        raise ValueError("表头行号必须大于等于 1。")
    if data_start_row <= header_row:
        raise ValueError("数据起始行号必须大于表头行号。")


def validate_rows_per_part(rows_per_part: int) -> int:
    if isinstance(rows_per_part, bool) or not isinstance(rows_per_part, int):
        raise ValueError("每份行数必须是大于等于 1 的整数。")
    if rows_per_part < 1:
        raise ValueError("每份行数必须大于等于 1。")
    return rows_per_part


@dataclass(frozen=True)
class RowSplitPart:
    name: str
    keep_start_row: int
    keep_end_row: int


def build_row_split_plans(
    data_start_row: int,
    last_row: int,
    rows_per_part: int = 1,
) -> list[RowSplitPart]:
    validate_rows_per_part(rows_per_part)
    if last_row < data_start_row:
        raise RuntimeError("当前工作表没有可拆分的数据。")

    plans: list[RowSplitPart] = []
    batch_index = 1
    current_start = data_start_row
    while current_start <= last_row:
        current_end = min(current_start + rows_per_part - 1, last_row)
        name = f"数据{batch_index}" if rows_per_part == 1 else f"第{batch_index}批"
        plans.append(
            RowSplitPart(
                name=name,
                keep_start_row=current_start,
                keep_end_row=current_end,
            )
        )
        batch_index += 1
        current_start = current_end + 1
    return plans


def _get_row_split_rows_to_delete(
    data_start_row: int,
    last_row: int,
    keep_start_row: int,
    keep_end_row: int,
) -> list[int]:
    rows: list[int] = []
    for r in range(last_row, keep_end_row, -1):
        rows.append(r)
    for r in range(keep_start_row - 1, data_start_row - 1, -1):
        rows.append(r)
    return rows


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


def _get_split_rows_to_delete(
    split_values: list[object],
    target: str,
    data_start_row: int,
) -> list[int]:
    return [
        data_start_row + offset
        for offset in range(len(split_values) - 1, -1, -1)
        if _normalize_split_target(split_values[offset]) != target
    ]


def clean_split_filename(name: str | None) -> str:
    """清理拆分值中的 Windows 文件名非法字符和尾部空格/句点。"""
    if name is None:
        return ""

    return "".join(
        char
        for char in str(name).strip()
        if ord(char) >= 32 and char not in _INVALID_SPLIT_FILENAME_CHARS
    ).rstrip(" .")


def _is_windows_reserved_file_name(name: str) -> bool:
    first_component = name.split(".", 1)[0].upper()
    return first_component in _WINDOWS_RESERVED_FILE_NAMES


def get_safe_split_filename(name: str | None, fallback: str = "空白") -> str:
    """返回可用于输出文件主名的拆分值。"""
    cleaned = clean_split_filename(name)
    if not cleaned:
        cleaned = clean_split_filename(fallback)
    if not cleaned:
        cleaned = "split"

    cleaned = cleaned[:_MAX_SPLIT_FILENAME_BASE_LENGTH].rstrip(" .")
    if _is_windows_reserved_file_name(cleaned):
        cleaned = f"_{cleaned}"[:_MAX_SPLIT_FILENAME_BASE_LENGTH].rstrip(" .")
    return cleaned or "split"


def get_unique_split_filepath(
    directory: str,
    base_name: str,
    extension: str = ".xlsx",
    reserved_paths: set[str] | None = None,
) -> str:
    """生成不覆盖已有文件、且不与本次运行已预留路径冲突的绝对路径。"""
    normalized_extension = str(extension)
    if not normalized_extension.startswith("."):
        normalized_extension = f".{normalized_extension}"

    reserved = reserved_paths if reserved_paths is not None else set()
    counter = 1
    while True:
        suffix = "" if counter == 1 else f"_{counter}"
        filepath = os.path.abspath(
            os.path.join(directory, f"{base_name}{suffix}{normalized_extension}")
        )
        normalized_path = os.path.normcase(filepath)
        if not os.path.exists(filepath) and normalized_path not in reserved:
            reserved.add(normalized_path)
            return filepath
        counter += 1


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


def _get_workbook_sheet_names(workbook) -> set[str]:
    return {sheet.Name for sheet in workbook.Worksheets}


def _validate_output_directory(output_dir: str) -> str:
    if output_dir is None:
        raise ValueError("输出目录不能为空。")

    try:
        raw_path = os.fsdecode(os.fspath(output_dir))
    except TypeError as e:
        raise ValueError("输出目录无效。") from e

    if not raw_path.strip():
        raise ValueError("输出目录不能为空。")

    normalized_path = os.path.abspath(raw_path)
    if not os.path.exists(normalized_path):
        raise FileNotFoundError(f"输出目录不存在：{normalized_path}")
    if not os.path.isdir(normalized_path):
        raise NotADirectoryError(f"输出路径不是目录：{normalized_path}")
    if not os.access(normalized_path, os.W_OK):
        raise PermissionError(f"输出目录不可写：{normalized_path}")
    return normalized_path


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


def _collect_split_values(
    source_sheet,
    column_input: str,
    header_row: int,
    data_start_row: int,
    logger=None,
) -> tuple[list[object], list[str]]:
    validate_row_numbers(header_row, data_start_row)
    column_index = parse_column_index(column_input)

    _log(logger, "info", f"准备拆分工作表：{source_sheet.Name}，拆分列：{column_input.strip()}（第 {column_index} 列）")

    last_row, last_col = _get_last_used_row_and_col(source_sheet)
    if last_row < data_start_row:
        raise RuntimeError("当前活动工作表没有可拆分的数据。")
    if column_index > last_col:
        raise RuntimeError("拆分列超出当前工作表的有效区域。")

    split_values = [
        row[0]
        for row in _read_range_values(
            source_sheet,
            data_start_row,
            column_index,
            last_row,
            column_index,
        )
    ]
    return split_values, build_split_targets(split_values)


def _delete_non_target_rows(
    sheet,
    split_values: list[object],
    target: str,
    data_start_row: int,
) -> int:
    rows_to_delete = _get_split_rows_to_delete(
        split_values,
        target,
        data_start_row,
    )
    for row_number in rows_to_delete:
        sheet.Rows(row_number).Delete()
    return len(split_values) - len(rows_to_delete)


def _close_workbook_without_saving(workbook) -> None:
    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        pass


def _is_same_workbook(first_workbook, second_workbook) -> bool:
    if first_workbook is second_workbook:
        return True

    try:
        if first_workbook == second_workbook:
            return True
    except Exception:
        pass

    try:
        first_full_name = str(first_workbook.FullName)
        second_full_name = str(second_workbook.FullName)
    except Exception:
        return False

    if not first_full_name or not second_full_name:
        return False
    return os.path.normcase(os.path.abspath(first_full_name)) == os.path.normcase(
        os.path.abspath(second_full_name)
    )


def _get_new_workbook_after_sheet_copy(source_workbook):
    try:
        excel = source_workbook.Application
        output_workbook = excel.ActiveWorkbook
    except Exception as e:
        raise RuntimeError("复制工作表后无法取得 Excel 新建的输出工作簿。") from e

    if not output_workbook:
        raise RuntimeError("复制工作表后没有取得输出工作簿。")
    if _is_same_workbook(output_workbook, source_workbook):
        raise RuntimeError("复制工作表没有创建新的输出工作簿，已停止以避免误操作源工作簿。")
    return output_workbook


def _copy_sheet_and_delete_non_target_rows(
    source_workbook,
    source_sheet,
    split_values: list[object],
    target: str,
    data_start_row: int,
    *,
    destination_workbook=None,
    sheet_name: str | None = None,
):
    copied_workbook = destination_workbook
    try:
        if destination_workbook is None:
            source_sheet.Copy()
            copied_workbook = _get_new_workbook_after_sheet_copy(source_workbook)
            new_sheet = copied_workbook.Worksheets(copied_workbook.Worksheets.Count)
        else:
            source_sheet.Copy(
                None,
                destination_workbook.Worksheets(destination_workbook.Worksheets.Count),
            )
            new_sheet = destination_workbook.Worksheets(destination_workbook.Worksheets.Count)

        if sheet_name is not None:
            new_sheet.Name = sheet_name

        kept_row_count = _delete_non_target_rows(
            new_sheet,
            split_values,
            target,
            data_start_row,
        )
        return copied_workbook, new_sheet, kept_row_count
    except Exception:
        if destination_workbook is None and copied_workbook is not None:
            _close_workbook_without_saving(copied_workbook)
        raise


def _copy_sheet_and_keep_row_range(
    source_workbook,
    source_sheet,
    data_start_row: int,
    last_row: int,
    keep_start_row: int,
    keep_end_row: int,
    *,
    destination_workbook=None,
    sheet_name: str | None = None,
):
    copied_workbook = destination_workbook
    try:
        if destination_workbook is None:
            source_sheet.Copy()
            copied_workbook = _get_new_workbook_after_sheet_copy(source_workbook)
            new_sheet = copied_workbook.Worksheets(copied_workbook.Worksheets.Count)
        else:
            source_sheet.Copy(
                None,
                destination_workbook.Worksheets(destination_workbook.Worksheets.Count),
            )
            new_sheet = destination_workbook.Worksheets(destination_workbook.Worksheets.Count)

        if sheet_name is not None:
            new_sheet.Name = sheet_name

        rows_to_delete = _get_row_split_rows_to_delete(
            data_start_row=data_start_row,
            last_row=last_row,
            keep_start_row=keep_start_row,
            keep_end_row=keep_end_row,
        )
        for row_number in rows_to_delete:
            new_sheet.Rows(row_number).Delete()

        kept_row_count = keep_end_row - keep_start_row + 1
        return copied_workbook, new_sheet, kept_row_count
    except Exception:
        if destination_workbook is None and copied_workbook is not None:
            _close_workbook_without_saving(copied_workbook)
        raise


def _split_workbook_sheet_by_column(
    workbook,
    source_sheet,
    column_input: str,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    split_values, targets = _collect_split_values(
        source_sheet,
        column_input,
        header_row,
        data_start_row,
        logger,
    )

    existing_names = _get_workbook_sheet_names(workbook)
    created_sheet_count = 0
    copied_row_count = 0

    for target in targets:
        safe_name = get_safe_sheet_name(target, fallback="空白")
        unique_name = get_unique_sheet_name(safe_name, existing_names)
        _, _, kept_row_count = _copy_sheet_and_delete_non_target_rows(
            workbook,
            source_sheet,
            split_values,
            target,
            data_start_row,
            destination_workbook=workbook,
            sheet_name=unique_name,
        )

        created_sheet_count += 1
        copied_row_count += kept_row_count
        _log(logger, "info", f"已生成工作表：{unique_name}，复制 {kept_row_count} 行。")

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


def _split_workbook_sheet_by_rows(
    workbook,
    source_sheet,
    rows_per_part: int = 1,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    validate_rows_per_part(rows_per_part)

    last_row, _ = _get_last_used_row_and_col(source_sheet)
    plans = build_row_split_plans(
        data_start_row=data_start_row,
        last_row=last_row,
        rows_per_part=rows_per_part,
    )

    _log(
        logger,
        "info",
        f"准备按行拆分工作表：{source_sheet.Name}，数据行范围：{data_start_row}~{last_row}，每份 {rows_per_part} 行，共 {len(plans)} 份",
    )

    existing_names = _get_workbook_sheet_names(workbook)
    created_sheet_count = 0
    copied_row_count = 0

    for part in plans:
        safe_name = get_safe_sheet_name(part.name, fallback="数据")
        unique_name = get_unique_sheet_name(safe_name, existing_names)
        _, _, kept_row_count = _copy_sheet_and_keep_row_range(
            workbook,
            source_sheet,
            data_start_row=data_start_row,
            last_row=last_row,
            keep_start_row=part.keep_start_row,
            keep_end_row=part.keep_end_row,
            destination_workbook=workbook,
            sheet_name=unique_name,
        )

        created_sheet_count += 1
        copied_row_count += kept_row_count
        _log(
            logger,
            "info",
            f"已生成工作表：{unique_name}，保留第 {part.keep_start_row}~{part.keep_end_row} 行（共 {kept_row_count} 行）。",
        )

    _log(logger, "info", f"按行拆分完成，共生成 {created_sheet_count} 个工作表，复制 {copied_row_count} 行。")
    return {
        "workbook_name": workbook.Name,
        "source_sheet_name": source_sheet.Name,
        "created_sheet_count": created_sheet_count,
        "copied_row_count": copied_row_count,
        "rows_per_part": rows_per_part,
    }


def split_active_sheet_by_rows(
    rows_per_part: int = 1,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    validate_rows_per_part(rows_per_part)

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

    return _split_workbook_sheet_by_rows(
        workbook,
        source_sheet,
        rows_per_part=rows_per_part,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )


def split_workbook_sheet_by_rows(
    source_path: str,
    source_sheet_name: str | None,
    rows_per_part: int = 1,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    validate_rows_per_part(rows_per_part)

    workbook = _get_or_open_workbook(source_path, logger)
    sheet_names = [sheet.Name for sheet in workbook.Worksheets]
    resolved_name = resolve_source_sheet_name(source_sheet_name, sheet_names)
    source_sheet = workbook.Worksheets(resolved_name)

    return _split_workbook_sheet_by_rows(
        workbook,
        source_sheet,
        rows_per_part=rows_per_part,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )


def _get_split_output_settings(source_path: str, op_name: str = "按列拆分") -> tuple[str, int]:
    if not source_path:
        raise ValueError("请选择源 Excel 工作簿。")

    extension = os.path.splitext(os.fsdecode(os.fspath(source_path)))[1].lower()
    if extension == ".xlsx":
        return ".xlsx", EXCEL_FILE_FORMAT_XLSX
    if extension == ".xlsm":
        raise ValueError(
            f"{op_name}为多个文件暂不支持 .xlsm：Worksheet.Copy 创建的新工作簿无法可靠保留 VBA 工程。"
        )
    if extension == ".xls":
        raise ValueError(
            f"{op_name}为多个文件暂不支持 .xls：当前没有可靠的旧格式保存处理，请先另存为 .xlsx。"
        )
    raise ValueError(
        f"{op_name}为多个文件暂不支持源文件格式：{extension or '无扩展名'}。目前仅支持 .xlsx。"
    )


def _split_workbook_sheet_by_column_to_files(
    source_workbook,
    source_sheet,
    column_input: str,
    output_dir: str,
    output_extension: str,
    output_file_format: int,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    split_values, targets = _collect_split_values(
        source_sheet,
        column_input,
        header_row,
        data_start_row,
        logger,
    )

    output_paths: list[str] = []
    reserved_paths: set[str] = set()
    copied_row_count = 0

    for target in targets:
        sheet_name = get_safe_sheet_name(target, fallback="空白")
        file_base_name = get_safe_split_filename(target, fallback="空白")
        output_path = get_unique_split_filepath(
            output_dir,
            file_base_name,
            output_extension,
            reserved_paths,
        )
        _log(logger, "info", f"正在生成文件：{source_sheet.Name} -> {output_path}")

        output_workbook = None
        try:
            output_workbook, _, kept_row_count = _copy_sheet_and_delete_non_target_rows(
                source_workbook,
                source_sheet,
                split_values,
                target,
                data_start_row,
                sheet_name=sheet_name,
            )
            output_workbook.SaveAs(
                os.path.abspath(output_path),
                FileFormat=output_file_format,
            )
            output_workbook.Close(SaveChanges=False)
            output_workbook = None
        except Exception:
            if output_workbook is not None:
                _close_workbook_without_saving(output_workbook)
            raise

        output_paths.append(output_path)
        copied_row_count += kept_row_count
        _log(logger, "info", f"已生成文件：{output_path}，复制 {kept_row_count} 行。")

    _log(logger, "info", f"拆分完成，共生成 {len(output_paths)} 个文件，复制 {copied_row_count} 行。")
    return {
        "workbook_name": source_workbook.Name,
        "source_sheet_name": source_sheet.Name,
        "created_file_count": len(output_paths),
        "copied_row_count": copied_row_count,
        "output_dir": output_dir,
        "output_paths": output_paths,
    }


def split_workbook_sheet_by_column_to_files(
    source_path: str,
    source_sheet_name: str | None,
    column_input: str,
    output_dir: str,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    output_extension, output_file_format = _get_split_output_settings(source_path)
    normalized_output_dir = _validate_output_directory(output_dir)

    source_workbook = _get_or_open_workbook(source_path, logger)
    sheet_names = [sheet.Name for sheet in source_workbook.Worksheets]
    resolved_name = resolve_source_sheet_name(source_sheet_name, sheet_names)
    source_sheet = source_workbook.Worksheets(resolved_name)

    return _split_workbook_sheet_by_column_to_files(
        source_workbook,
        source_sheet,
        column_input=column_input,
        output_dir=normalized_output_dir,
        output_extension=output_extension,
        output_file_format=output_file_format,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )

def _split_workbook_sheet_by_rows_to_files(
    source_workbook,
    source_sheet,
    output_dir: str,
    output_extension: str,
    output_file_format: int,
    rows_per_part: int = 1,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    validate_rows_per_part(rows_per_part)

    last_row, _ = _get_last_used_row_and_col(source_sheet)
    plans = build_row_split_plans(
        data_start_row=data_start_row,
        last_row=last_row,
        rows_per_part=rows_per_part,
    )

    _log(
        logger,
        "info",
        f"准备按行拆分为多个文件：{source_sheet.Name}，数据行范围：{data_start_row}~{last_row}，每份 {rows_per_part} 行，共 {len(plans)} 个文件",
    )

    output_paths: list[str] = []
    reserved_paths: set[str] = set()
    copied_row_count = 0

    for part in plans:
        sheet_name = get_safe_sheet_name(part.name, fallback="数据")
        file_base_name = get_safe_split_filename(part.name, fallback="数据")
        output_path = get_unique_split_filepath(
            output_dir,
            file_base_name,
            output_extension,
            reserved_paths,
        )
        _log(logger, "info", f"正在生成文件：{source_sheet.Name} -> {output_path}")

        output_workbook = None
        try:
            output_workbook, _, kept_row_count = _copy_sheet_and_keep_row_range(
                source_workbook,
                source_sheet,
                data_start_row=data_start_row,
                last_row=last_row,
                keep_start_row=part.keep_start_row,
                keep_end_row=part.keep_end_row,
                sheet_name=sheet_name,
            )
            output_workbook.SaveAs(
                os.path.abspath(output_path),
                FileFormat=output_file_format,
            )
            output_workbook.Close(SaveChanges=False)
            output_workbook = None
        except Exception:
            if output_workbook is not None:
                _close_workbook_without_saving(output_workbook)
            raise

        output_paths.append(output_path)
        copied_row_count += kept_row_count
        _log(logger, "info", f"已生成文件：{output_path}，复制 {kept_row_count} 行。")

    _log(logger, "info", f"按行拆分完成，共生成 {len(output_paths)} 个文件，复制 {copied_row_count} 行。")
    return {
        "workbook_name": source_workbook.Name,
        "source_sheet_name": source_sheet.Name,
        "created_file_count": len(output_paths),
        "copied_row_count": copied_row_count,
        "rows_per_part": rows_per_part,
        "output_dir": output_dir,
        "output_paths": output_paths,
    }


def split_workbook_sheet_by_rows_to_files(
    source_path: str,
    source_sheet_name: str | None,
    output_dir: str,
    rows_per_part: int = 1,
    header_row: int = 1,
    data_start_row: int = 2,
    logger=None,
) -> dict:
    validate_row_numbers(header_row, data_start_row)
    validate_rows_per_part(rows_per_part)

    output_extension, output_file_format = _get_split_output_settings(
        source_path, op_name="按行拆分"
    )
    normalized_output_dir = _validate_output_directory(output_dir)

    source_workbook = _get_or_open_workbook(source_path, logger)
    sheet_names = [sheet.Name for sheet in source_workbook.Worksheets]
    resolved_name = resolve_source_sheet_name(source_sheet_name, sheet_names)
    source_sheet = source_workbook.Worksheets(resolved_name)

    return _split_workbook_sheet_by_rows_to_files(
        source_workbook,
        source_sheet,
        output_dir=normalized_output_dir,
        output_extension=output_extension,
        output_file_format=output_file_format,
        rows_per_part=rows_per_part,
        header_row=header_row,
        data_start_row=data_start_row,
        logger=logger,
    )
