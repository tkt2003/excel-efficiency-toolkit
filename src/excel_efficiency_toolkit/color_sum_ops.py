import os


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm"}
XL_COLOR_INDEX_NONE = -4142
XL_PATTERN_NONE = -4142


class MissingSheetError(ValueError):
    pass


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


def can_read_with_openpyxl(path: str) -> bool:
    if not is_excel_file(path):
        return False

    _, ext = os.path.splitext(os.path.basename(path))
    return ext.lower() in OPENPYXL_EXTENSIONS


def sheet_exists_with_openpyxl(path: str, sheet_name: str) -> bool:
    workbook = _load_openpyxl_workbook(path)
    try:
        return _find_openpyxl_sheet_name(workbook.sheetnames, sheet_name) is not None
    finally:
        workbook.close()


def read_sheet_values_with_openpyxl(
    path: str,
    sheet_name: str,
    cells: list[tuple[int, int]],
) -> dict[tuple[int, int], object]:
    workbook = _load_openpyxl_workbook(path)
    try:
        actual_sheet_name = _find_openpyxl_sheet_name(workbook.sheetnames, sheet_name)
        if actual_sheet_name is None:
            raise MissingSheetError(f"源文件缺少同名 sheet：{sheet_name}")

        sheet = workbook[actual_sheet_name]
        return _read_openpyxl_cells(sheet, cells)
    finally:
        workbook.close()


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
        target_excel = _get_active_excel(win32com.client, pythoncom, logger=logger)
        target_workbook, active_sheet, active_cell = _get_active_target_context(target_excel, logger=logger)
        active_cell_fill_info = _get_cell_fill_color_info(active_cell)
        _log_active_cell_diagnostics(logger, target_workbook, active_sheet, active_cell, active_cell_fill_info)
        selected_color = active_cell_fill_info["color"] if active_cell_fill_info["has_fill"] else None
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

        target_excel_state = _set_excel_fast_mode(target_excel)

        matched_sources: list[tuple[str, str]] = []
        missing_sheet_file_count = 0
        ignored_non_numeric_count = 0
        totals = {cell: 0.0 for cell in matched_cells}

        source_file_count = len(valid_source_paths)
        for index, source_path in enumerate(valid_source_paths, start=1):
            source_filename = os.path.basename(source_path)
            _log(logger, "info", f"正在读取源文件 {index}/{source_file_count}：{source_filename}")

            if can_read_with_openpyxl(source_path):
                try:
                    _log(logger, "info", f"使用 openpyxl 只读读取源文件：{source_filename}")
                    if normalized_write_mode == "value":
                        source_values = read_sheet_values_with_openpyxl(
                            source_path,
                            target_sheet_actual_name,
                            matched_cells,
                        )
                        ignored_non_numeric_count += _accumulate_values_from_mapping(
                            source_values=source_values,
                            totals=totals,
                            source_path=source_path,
                            sheet_name=target_sheet_actual_name,
                            logger=logger,
                        )
                        _log(logger, "info", f"找到同名工作表：{target_sheet_actual_name}")
                        matched_sources.append((source_path, target_sheet_actual_name))
                    else:
                        actual_sheet_name = _get_openpyxl_sheet_name(source_path, target_sheet_actual_name)
                        if actual_sheet_name is None:
                            missing_sheet_file_count += 1
                            _log(
                                logger,
                                "info",
                                f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：缺少同名工作表",
                            )
                            continue
                        _log(logger, "info", f"找到同名工作表：{actual_sheet_name}")
                        matched_sources.append((source_path, actual_sheet_name))
                    continue
                except MissingSheetError:
                    missing_sheet_file_count += 1
                    _log(
                        logger,
                        "info",
                        f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：缺少同名工作表",
                    )
                    continue
                except Exception as e:
                    _log(logger, "error", f"openpyxl 读取失败，回退 COM：{source_filename}，原因：{e}")

            if source_excel is None:
                source_excel = win32com.client.DispatchEx("Excel.Application")
                source_excel_state = _set_excel_fast_mode(source_excel)
                source_excel.Visible = False

            com_result = _read_source_with_com(
                source_excel=source_excel,
                source_path=source_path,
                target_sheet_name=target_sheet_actual_name,
                matched_cells=matched_cells,
                write_mode=normalized_write_mode,
                logger=logger,
            )
            if com_result["status"] == "missing_sheet":
                missing_sheet_file_count += 1
                _log(
                    logger,
                    "info",
                    f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：缺少同名工作表",
                )
                continue
            if com_result["status"] == "failed":
                _log(
                    logger,
                    "error",
                    f"跳过源文件 {index}/{source_file_count}：{source_filename}，原因：读取失败：{com_result['error']}",
                )
                continue

            matched_sources.append((source_path, com_result["sheet_name"]))
            ignored_non_numeric_count += com_result["ignored_non_numeric_count"]
            for cell in matched_cells:
                totals[cell] += com_result["totals"][cell]

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


def _load_openpyxl_workbook(path: str):
    from openpyxl import load_workbook

    return load_workbook(path, read_only=True, data_only=True)


def _find_openpyxl_sheet_name(sheet_names: list[str], sheet_name: str) -> str | None:
    expected_name = _normalize_sheet_name(sheet_name)
    for actual_name in sheet_names:
        if _normalize_sheet_name(actual_name) == expected_name:
            return actual_name
    return None


def _get_openpyxl_sheet_name(path: str, sheet_name: str) -> str | None:
    workbook = _load_openpyxl_workbook(path)
    try:
        return _find_openpyxl_sheet_name(workbook.sheetnames, sheet_name)
    finally:
        workbook.close()


def _read_openpyxl_cells(sheet, cells: list[tuple[int, int]]) -> dict[tuple[int, int], object]:
    result = {cell: None for cell in cells}
    if not cells:
        return result

    min_row = min(row for row, _ in cells)
    max_row = max(row for row, _ in cells)
    min_col = min(col for _, col in cells)
    max_col = max(col for _, col in cells)
    target_cols_by_row: dict[int, set[int]] = {}
    for row, col in cells:
        target_cols_by_row.setdefault(row, set()).add(col)

    for row_offset, row_values in enumerate(
        sheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        )
    ):
        actual_row = min_row + row_offset
        target_cols = target_cols_by_row.get(actual_row)
        if not target_cols:
            continue
        for col in target_cols:
            value_index = col - min_col
            if value_index < len(row_values):
                result[(actual_row, col)] = row_values[value_index]

    return result


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


def _get_active_excel(win32com_client, pythoncom_module, logger=None):
    attempts: list[str] = []

    connectors = [
        (
            "win32com.client.GetActiveObject",
            lambda: [(win32com_client.GetActiveObject("Excel.Application"), None)],
        ),
        (
            "pythoncom.GetActiveObject",
            lambda: [(win32com_client.Dispatch(pythoncom_module.GetActiveObject("Excel.Application")), None)],
        ),
        (
            "运行对象表",
            lambda: _get_excel_candidates_from_rot(win32com_client, pythoncom_module),
        ),
        (
            "win32com.client.Dispatch",
            lambda: [(win32com_client.Dispatch("Excel.Application"), None)],
        ),
    ]

    for method_name, connect in connectors:
        try:
            candidates = connect()
        except Exception as e:
            detail = f"{method_name} 失败：{_format_exception(e)}"
            attempts.append(detail)
            _log(logger, "info", detail)
            continue

        if not candidates:
            detail = f"{method_name} 未找到可用的 Excel.Application"
            attempts.append(detail)
            _log(logger, "info", detail)
            continue

        for excel, source_name in candidates:
            has_workbook, workbook_detail = _excel_has_active_workbook(excel)
            source_detail = f"，来源：{source_name}" if source_name else ""
            detail = f"{method_name} 成功{source_detail}，{workbook_detail}"
            attempts.append(detail)
            _log(logger, "info", detail)
            if has_workbook:
                _log(logger, "info", f"当前连接到的 Excel 应用：成功，方式：{method_name}")
                return excel

            if method_name == "win32com.client.Dispatch":
                _quit_empty_hidden_excel(excel)

    raise RuntimeError(
        "无法连接到当前 Excel 实例或无法读取活动工作簿。"
        "请确认目标 Excel 已打开，未处于单元格编辑、弹窗或受保护视图状态，"
        "并确认工具和 Excel 以相同权限运行。"
        f"连接诊断：{'；'.join(attempts)}"
    )


def _get_excel_candidates_from_rot(win32com_client, pythoncom_module) -> list[tuple[object, str]]:
    candidates: list[tuple[object, str]] = []
    rot = pythoncom_module.GetRunningObjectTable()
    bind_context = pythoncom_module.CreateBindCtx(0)
    enum_moniker = rot.EnumRunning()

    while True:
        monikers = enum_moniker.Next(1)
        if not monikers:
            break

        moniker = monikers[0]
        try:
            display_name = moniker.GetDisplayName(bind_context, None)
        except Exception:
            display_name = ""

        if not _looks_like_excel_rot_entry(display_name):
            continue

        try:
            running_object = rot.GetObject(moniker)
        except Exception:
            try:
                running_object = moniker.BindToObject(bind_context, None, pythoncom_module.IID_IDispatch)
            except Exception:
                continue

        try:
            dispatch = win32com_client.Dispatch(running_object)
        except Exception:
            continue

        excel = _get_com_property(dispatch, "Application") or dispatch
        candidates.append((excel, display_name or "ROT"))

    return candidates


def _looks_like_excel_rot_entry(display_name: str) -> bool:
    lower_name = str(display_name).lower()
    return (
        "excel" in lower_name
        or lower_name.endswith(".xlsx")
        or lower_name.endswith(".xlsm")
        or lower_name.endswith(".xls")
    )


def _excel_has_active_workbook(excel) -> tuple[bool, str]:
    try:
        workbook = excel.ActiveWorkbook
    except Exception as e:
        return False, f"ActiveWorkbook 读取失败：{_format_exception(e)}"

    if workbook is None:
        return False, "ActiveWorkbook 为空"

    workbook_name = _get_com_object_name(workbook)
    return True, f"ActiveWorkbook.Name：{workbook_name}"


def _quit_empty_hidden_excel(excel) -> None:
    try:
        workbooks = excel.Workbooks
        workbook_count = int(workbooks.Count)
        visible = bool(excel.Visible)
    except Exception:
        return

    if workbook_count == 0 and not visible:
        try:
            excel.Quit()
        except Exception:
            pass


def _format_exception(error: Exception) -> str:
    return f"{type(error).__name__}: {error}"


def _get_active_target_context(excel, logger=None):
    target_workbook = _get_com_property(excel, "ActiveWorkbook")
    active_sheet = _get_com_property(excel, "ActiveSheet")
    active_cell = _get_selected_top_left_cell(excel)

    if active_sheet is None and active_cell is not None:
        active_sheet = _get_com_property(active_cell, "Worksheet")
    if target_workbook is None and active_sheet is not None:
        target_workbook = _get_com_property(active_sheet, "Parent")

    _log(logger, "info", f"ActiveWorkbook.Name：{_get_com_object_name(target_workbook)}")
    _log(logger, "info", f"ActiveSheet.Name：{_get_com_object_name(active_sheet)}")
    _log(logger, "info", f"ActiveCell.Address：{_get_cell_address_for_log(active_cell)}")

    if target_workbook is None or active_sheet is None or active_cell is None:
        raise RuntimeError("请先打开目标合并工作簿，并选中一个带颜色的单元格。")

    return target_workbook, active_sheet, active_cell


def _get_com_property(obj, attr: str):
    if obj is None:
        return None
    try:
        return getattr(obj, attr)
    except Exception:
        return None


def _get_com_object_name(obj) -> str:
    name = _get_com_property(obj, "Name")
    if name is None:
        return "无法获取"
    return str(name)


def _get_selected_top_left_cell(excel):
    selection = _get_com_property(excel, "Selection")
    cell = _get_first_cell_from_range(selection)
    if cell is not None:
        return cell

    return _get_com_property(excel, "ActiveCell")


def _get_first_cell_from_range(range_obj):
    if range_obj is None:
        return None

    cells = _get_com_property(range_obj, "Cells")
    if cells is None:
        return None

    try:
        return cells(1, 1)
    except Exception:
        pass

    item = _get_com_property(cells, "Item")
    if item is None:
        return None
    try:
        return item(1, 1)
    except Exception:
        return None


def _get_cell_address_for_log(cell) -> str:
    if cell is None:
        return "无法获取"

    try:
        return build_cell_address(int(cell.Row), int(cell.Column))
    except Exception:
        address = _get_com_property(cell, "Address")
        if address is None:
            return "无法获取"
        return str(address)


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
    fill_info = _get_cell_fill_color_info(cell)
    return fill_info["color"] if fill_info["has_fill"] else None


def _get_cell_fill_color_info(cell) -> dict:
    fill_info = {
        "color": None,
        "color_index": None,
        "pattern": None,
        "has_fill": False,
    }
    interior = _get_com_property(cell, "Interior")
    if interior is None:
        return fill_info

    fill_info["color"] = _get_com_property(interior, "Color")
    fill_info["color_index"] = _get_com_property(interior, "ColorIndex")
    fill_info["pattern"] = _get_com_property(interior, "Pattern")
    fill_info["has_fill"] = _has_effective_fill_color(
        fill_info["color"],
        fill_info["color_index"],
        fill_info["pattern"],
    )
    return fill_info


def _has_effective_fill_color(color, color_index, pattern) -> bool:
    if _same_excel_constant(color_index, XL_COLOR_INDEX_NONE):
        return False
    if color is not None:
        return True
    if _same_excel_constant(pattern, XL_PATTERN_NONE):
        return False
    return False


def _same_excel_constant(value, constant: int) -> bool:
    try:
        return int(value) == constant
    except (TypeError, ValueError):
        return False


def _log_active_cell_diagnostics(logger, target_workbook, active_sheet, active_cell, fill_info: dict) -> None:
    _log(logger, "info", f"当前活动工作簿：{_get_com_object_name(target_workbook)}")
    _log(logger, "info", f"当前活动工作表：{_get_com_object_name(active_sheet)}")
    _log(logger, "info", f"当前选中单元格：{_get_cell_address_for_log(active_cell)}")
    _log(logger, "info", f"当前选中单元格 Interior.Color：{fill_info['color']}")
    _log(logger, "info", f"当前选中单元格 Interior.ColorIndex：{fill_info['color_index']}")
    _log(logger, "info", f"当前选中单元格 Interior.Pattern：{fill_info['pattern']}")
    fill_status = "有填充色" if fill_info["has_fill"] else "无填充色"
    _log(logger, "info", f"当前选中单元格填充色判定：{fill_status}")


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


def _accumulate_values_from_mapping(
    source_values: dict[tuple[int, int], object],
    totals: dict[tuple[int, int], float],
    source_path: str,
    sheet_name: str,
    logger=None,
) -> int:
    ignored_non_numeric_count = 0
    for row, col in totals:
        value = source_values.get((row, col))
        number = to_number_or_none(value)
        if number is None:
            if not _is_blank_value(value):
                ignored_non_numeric_count += 1
                address = build_cell_address(row, col)
                _log(logger, "info", f"非数字已忽略：{os.path.basename(source_path)} {sheet_name}!{address} = {value}")
            continue
        totals[(row, col)] += number
    return ignored_non_numeric_count


def _read_source_with_com(
    source_excel,
    source_path: str,
    target_sheet_name: str,
    matched_cells: list[tuple[int, int]],
    write_mode: str,
    logger=None,
) -> dict:
    source_workbook = None
    source_filename = os.path.basename(source_path)
    totals = {cell: 0.0 for cell in matched_cells}

    try:
        _log(logger, "info", f"使用 COM 读取源文件：{source_filename}")
        source_workbook = source_excel.Workbooks.Open(
            source_path,
            ReadOnly=True,
            UpdateLinks=0,
        )
        source_sheet = _find_worksheet_by_name(source_workbook, target_sheet_name)
        if source_sheet is None:
            return {
                "status": "missing_sheet",
                "sheet_name": None,
                "totals": totals,
                "ignored_non_numeric_count": 0,
                "error": None,
            }

        _log(logger, "info", f"找到同名工作表：{source_sheet.Name}")
        ignored_non_numeric_count = 0
        if write_mode == "value":
            ignored_non_numeric_count = _accumulate_source_sheet_values_from_used_range(
                source_sheet=source_sheet,
                cells=matched_cells,
                totals=totals,
                source_path=source_path,
                logger=logger,
            )

        return {
            "status": "ok",
            "sheet_name": source_sheet.Name,
            "totals": totals,
            "ignored_non_numeric_count": ignored_non_numeric_count,
            "error": None,
        }
    except Exception as e:
        return {
            "status": "failed",
            "sheet_name": None,
            "totals": totals,
            "ignored_non_numeric_count": 0,
            "error": e,
        }
    finally:
        if source_workbook is not None:
            try:
                source_workbook.Close(SaveChanges=False)
            except Exception:
                pass


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
