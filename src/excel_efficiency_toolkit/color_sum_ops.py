import os


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
XL_COLOR_INDEX_NONE = -4142
XL_PATTERN_NONE = -4142


def column_number_to_letter(col: int) -> str:
    if col < 1:
        raise ValueError("列号必须大于等于 1。")

    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def build_cell_address(row: int, col: int) -> str:
    if row < 1:
        raise ValueError("行号必须大于等于 1。")
    return f"{column_number_to_letter(col)}{row}"


def group_cells_by_contiguous_rows(cells: list[tuple[int, int]]) -> list[list[tuple[int, int]]]:
    groups: list[list[tuple[int, int]]] = []
    for row, col in sorted(cells):
        if groups and groups[-1][-1][0] == row and groups[-1][-1][1] + 1 == col:
            groups[-1].append((row, col))
        else:
            groups.append([(row, col)])
    return groups


def get_value_from_used_range_array(
    values,
    used_range_first_row: int,
    used_range_first_col: int,
    row: int,
    col: int,
):
    row_index = row - used_range_first_row
    col_index = col - used_range_first_col
    if row_index < 0 or col_index < 0:
        return None

    if not isinstance(values, (tuple, list)):
        return values if row_index == 0 and col_index == 0 else None

    if not values:
        return None

    first_item = values[0]
    if isinstance(first_item, (tuple, list)):
        if row_index >= len(values):
            return None
        row_values = values[row_index]
        if isinstance(row_values, (tuple, list)):
            if col_index >= len(row_values):
                return None
            return row_values[col_index]
        return row_values if col_index == 0 else None

    if row_index == 0 and col_index < len(values):
        return values[col_index]
    if col_index == 0 and row_index < len(values):
        return values[row_index]
    return None


def is_valid_number(value: object) -> bool:
    return to_number_or_none(value) is not None


def to_number_or_none(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def normalize_write_mode(user_input: str) -> str:
    normalized = str(user_input).strip().lower()
    if normalized in {"1", "公式", "formula", "写入公式", "写入求和公式"}:
        return "formula"
    if normalized in {"2", "数值", "value", "只保留数值", "只写入汇总数值"}:
        return "value"
    raise ValueError("写入方式无效，请输入 1/公式 或 2/数值。")


def build_external_sum_formula(source_refs: list[tuple[str, str, str]]) -> str:
    if not source_refs:
        raise ValueError("公式引用不能为空。")

    refs = [
        _build_external_cell_ref(source_file_path, sheet_name, cell_address)
        for source_file_path, sheet_name, cell_address in source_refs
    ]
    return "=" + "+".join(refs)


def same_excel_color(color1, color2) -> bool:
    if color1 is None or color2 is None:
        return False
    return color1 == color2


def is_excel_file(path: str) -> bool:
    filename = os.path.basename(path)
    if filename.startswith("~$"):
        return False

    _, ext = os.path.splitext(filename)
    return ext.lower() in EXCEL_EXTENSIONS


def sum_current_sheet_by_fill_color(
    source_paths: list[str],
    write_mode: str,
    target_sheet_name: str | None = None,
    logger=None,
) -> dict:
    normalized_write_mode = normalize_write_mode(write_mode)
    valid_source_paths = [os.path.abspath(path) for path in source_paths if is_excel_file(path)]
    if not valid_source_paths:
        raise ValueError("请选择至少一个有效的 Excel 文件。")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    target_excel = None
    target_excel_state = {}
    source_excel = None
    source_excel_state = {}
    try:
        target_excel = _get_active_excel(win32com.client)
        target_workbook, active_sheet, active_cell = _get_active_target_context(target_excel)
        selected_color = _get_cell_fill_color(active_cell)
        if selected_color is None:
            raise ValueError("请先选中一个带填充色的单元格。")

        target_sheet = _resolve_target_sheet(target_workbook, active_sheet, target_sheet_name)
        target_sheet_actual_name = target_sheet.Name
        matched_cells = _collect_matching_cells(target_sheet, selected_color)
        if not matched_cells:
            raise ValueError(f"目标 sheet 中没有找到与当前选中单元格同色的单元格：{target_sheet_actual_name}。")

        _log(logger, "info", f"目标工作簿：{target_workbook.Name}")
        _log(logger, "info", f"目标工作表：{target_sheet_actual_name}")
        _log(logger, "info", f"写入方式：{_get_write_mode_label(normalized_write_mode)}")
        _log(logger, "info", f"同色单元格数量：{len(matched_cells)}")
        _log(logger, "info", f"源文件数量：{len(valid_source_paths)}")

        source_excel = win32com.client.DispatchEx("Excel.Application")
        source_excel_state = _set_excel_fast_mode(source_excel)
        source_excel.Visible = False

        target_excel_state = _set_excel_fast_mode(target_excel)

        matched_sources: list[tuple[str, str]] = []
        missing_sheet_file_count = 0
        ignored_non_numeric_count = 0
        totals = {cell: 0.0 for cell in matched_cells}

        source_file_count = len(valid_source_paths)
        for index, source_path in enumerate(valid_source_paths, start=1):
            source_workbook = None
            source_filename = os.path.basename(source_path)
            try:
                _log(logger, "info", f"正在读取源文件 {index}/{source_file_count}：{source_filename}")
                source_workbook = source_excel.Workbooks.Open(
                    source_path,
                    ReadOnly=True,
                    UpdateLinks=0,
                )
                source_sheet = _find_worksheet_by_name(source_workbook, target_sheet_actual_name)
                if source_sheet is None:
                    missing_sheet_file_count += 1
                    _log(
                        logger,
                        "info",
                        f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：缺少同名工作表",
                    )
                    continue

                _log(logger, "info", f"找到同名工作表：{source_sheet.Name}")
                if normalized_write_mode == "value":
                    source_totals = {cell: 0.0 for cell in matched_cells}
                    source_ignored_count = _accumulate_source_sheet_values_from_used_range(
                        source_sheet=source_sheet,
                        cells=matched_cells,
                        totals=source_totals,
                        source_path=source_path,
                        logger=logger,
                    )
                    for cell in matched_cells:
                        totals[cell] += source_totals[cell]
                    ignored_non_numeric_count += source_ignored_count

                matched_sources.append((source_path, source_sheet.Name))
            except Exception as e:
                _log(
                    logger,
                    "error",
                    f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：读取失败：{e}",
                )
            finally:
                if source_workbook is not None:
                    try:
                        source_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass

        if not matched_sources:
            raise RuntimeError(f"所有有效源文件都没有找到同名 sheet：{target_sheet_actual_name}。")

        write_groups = group_cells_by_contiguous_rows(matched_cells)
        _log(
            logger,
            "info",
            f"正在写入目标工作表：{target_sheet_actual_name}；待写入单元格 {len(matched_cells)} 个，批量区域 {len(write_groups)} 组。",
        )
        written_cell_count = _write_target_cells(
            target_sheet=target_sheet,
            cell_groups=write_groups,
            write_mode=normalized_write_mode,
            totals=totals,
            matched_sources=matched_sources,
        )

        return {
            "target_workbook_name": target_workbook.Name,
            "target_sheet_name": target_sheet_actual_name,
            "matched_cell_count": len(matched_cells),
            "source_file_count": len(valid_source_paths),
            "matched_source_file_count": len(matched_sources),
            "missing_sheet_file_count": missing_sheet_file_count,
            "written_cell_count": written_cell_count,
            "ignored_non_numeric_count": ignored_non_numeric_count,
            "write_mode": normalized_write_mode,
        }

    finally:
        try:
            _restore_excel_state(target_excel, target_excel_state)
        except Exception:
            pass
        try:
            _restore_excel_state(source_excel, source_excel_state)
        except Exception:
            pass
        if source_excel is not None:
            try:
                source_excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _escape_excel_quote(value: object) -> str:
    return str(value).replace("'", "''")


def _build_external_cell_ref(source_file_path: str, sheet_name: str, cell_address: str) -> str:
    abs_path = os.path.abspath(source_file_path)
    directory, filename = os.path.split(abs_path)
    workbook_ref = os.path.join(directory, f"[{filename}]") if directory else f"[{filename}]"
    quoted_sheet_ref = _escape_excel_quote(f"{workbook_ref}{sheet_name}")
    return f"'{quoted_sheet_ref}'!{str(cell_address).strip()}"


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


def _set_excel_fast_mode(excel) -> dict:
    if excel is None:
        return {}

    state = {}
    for attr, value in {
        "ScreenUpdating": False,
        "EnableEvents": False,
        "DisplayAlerts": False,
    }.items():
        try:
            state[attr] = getattr(excel, attr)
            setattr(excel, attr, value)
        except Exception:
            pass
    return state


def _restore_excel_state(excel, state: dict) -> None:
    if excel is None:
        return
    for attr, value in state.items():
        try:
            setattr(excel, attr, value)
        except Exception:
            pass


def _get_write_mode_label(write_mode: str) -> str:
    if write_mode == "formula":
        return "公式"
    return "数值"


def _get_active_excel(win32com_client):
    try:
        return win32com_client.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError("请先打开目标合并工作簿，并选中一个带颜色的单元格。") from e


def _get_active_target_context(excel):
    try:
        target_workbook = excel.ActiveWorkbook
        active_sheet = excel.ActiveSheet
        active_cell = excel.ActiveCell
    except Exception as e:
        raise RuntimeError("请先打开目标合并工作簿，并选中一个带颜色的单元格。") from e

    if target_workbook is None or active_sheet is None or active_cell is None:
        raise RuntimeError("请先打开目标合并工作簿，并选中一个带颜色的单元格。")

    return target_workbook, active_sheet, active_cell


def _normalize_sheet_name(sheet_name: object) -> str:
    if sheet_name is None:
        return ""
    return str(sheet_name).strip().lower()


def _find_worksheet_by_name(workbook, sheet_name: str):
    expected_name = _normalize_sheet_name(sheet_name)
    for sheet in workbook.Worksheets:
        if _normalize_sheet_name(sheet.Name) == expected_name:
            return sheet
    return None


def _resolve_target_sheet(workbook, active_sheet, target_sheet_name: str | None):
    requested_name = "" if target_sheet_name is None else str(target_sheet_name).strip()
    if not requested_name:
        return active_sheet

    target_sheet = _find_worksheet_by_name(workbook, requested_name)
    if target_sheet is None:
        raise ValueError(f"目标工作簿中不存在 sheet：{requested_name}。")
    return target_sheet


def _get_cell_fill_color(cell):
    try:
        if cell.Interior.ColorIndex == XL_COLOR_INDEX_NONE:
            return None
    except Exception:
        pass

    try:
        if cell.Interior.Pattern == XL_PATTERN_NONE:
            return None
    except Exception:
        pass

    try:
        return cell.Interior.Color
    except Exception:
        return None


def _collect_matching_cells(sheet, selected_color) -> list[tuple[int, int]]:
    used_range = sheet.UsedRange
    first_row = used_range.Row
    first_column = used_range.Column
    row_count = used_range.Rows.Count
    column_count = used_range.Columns.Count
    cells: list[tuple[int, int]] = []

    for row in range(first_row, first_row + row_count):
        for column in range(first_column, first_column + column_count):
            cell = sheet.Cells(row, column)
            if same_excel_color(_get_cell_fill_color(cell), selected_color):
                cells.append((row, column))

    return cells


def _cell_address_from_row_column(row: int, column: int) -> str:
    return build_cell_address(row, column)


def _column_number_to_letters(column: int) -> str:
    return column_number_to_letter(column)


def _is_blank_value(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _read_used_range_values(sheet):
    used_range = sheet.UsedRange
    try:
        values = used_range.Value2
    except Exception:
        values = used_range.Value
    values = _normalize_used_range_values(values, used_range.Rows.Count, used_range.Columns.Count)
    return values, used_range.Row, used_range.Column


def _normalize_used_range_values(values, row_count: int, col_count: int):
    if row_count == 1 and col_count == 1:
        return values
    if not isinstance(values, (tuple, list)):
        return values
    if row_count == 1 and col_count > 1 and (not values or not isinstance(values[0], (tuple, list))):
        return (tuple(values),)
    if col_count == 1 and row_count > 1 and (not values or not isinstance(values[0], (tuple, list))):
        return tuple((value,) for value in values)
    return values


def _accumulate_source_sheet_values_from_used_range(
    source_sheet,
    cells: list[tuple[int, int]],
    totals: dict[tuple[int, int], float],
    source_path: str,
    logger=None,
) -> int:
    ignored_non_numeric_count = 0
    values, first_row, first_col = _read_used_range_values(source_sheet)
    for row, col in cells:
        value = get_value_from_used_range_array(values, first_row, first_col, row, col)
        number = to_number_or_none(value)
        if number is None:
            if not _is_blank_value(value):
                ignored_non_numeric_count += 1
                address = build_cell_address(row, col)
                _log(logger, "info", f"非数字已忽略：{os.path.basename(source_path)} {source_sheet.Name}!{address} = {value}")
            continue
        totals[(row, col)] += number
    return ignored_non_numeric_count


def _write_target_cells(
    target_sheet,
    cell_groups: list[list[tuple[int, int]]],
    write_mode: str,
    totals: dict[tuple[int, int], float],
    matched_sources: list[tuple[str, str]],
) -> int:
    written_cell_count = 0
    for group in cell_groups:
        row = group[0][0]
        start_col = group[0][1]
        end_col = group[-1][1]
        row_values = [_build_write_value(row, col, write_mode, totals, matched_sources) for row, col in group]

        if len(group) == 1:
            target_cell = target_sheet.Cells(row, start_col)
            _write_single_cell(target_cell, write_mode, row_values[0])
        else:
            target_range = target_sheet.Range(
                target_sheet.Cells(row, start_col),
                target_sheet.Cells(row, end_col),
            )
            _write_range_row(target_range, write_mode, row_values)
        written_cell_count += len(group)

    return written_cell_count


def _build_write_value(
    row: int,
    col: int,
    write_mode: str,
    totals: dict[tuple[int, int], float],
    matched_sources: list[tuple[str, str]],
):
    if write_mode == "formula":
        address = build_cell_address(row, col)
        source_refs = [
            (source_path, source_sheet_name, address)
            for source_path, source_sheet_name in matched_sources
        ]
        return build_external_sum_formula(source_refs)
    return totals[(row, col)]


def _write_single_cell(target_cell, write_mode: str, value) -> None:
    if write_mode == "formula":
        target_cell.Formula = value
    else:
        try:
            target_cell.Value2 = value
        except Exception:
            target_cell.Value = value


def _write_range_row(target_range, write_mode: str, values: list[object]) -> None:
    row_values = (tuple(values),)
    if write_mode == "formula":
        target_range.Formula = row_values
    else:
        try:
            target_range.Value2 = row_values
        except Exception:
            target_range.Value = row_values
