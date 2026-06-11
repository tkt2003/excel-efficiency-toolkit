import os
import shutil
import tempfile
from datetime import datetime


EXCEL_RULE_TABLE_FORMAT_XLSX = 51
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
RULE_SHEET_NAME = "批量删除规则"
SOURCE_LIST_SHEET_NAME = "源文件清单"
RULE_TABLE_HEADERS = ["所有表格名", "保留表格名", "删除表格名", "排除文件名关键词"]


def is_excel_file(path: str) -> bool:
    filename = os.path.basename(path)
    if filename.startswith("~$"):
        return False

    _, ext = os.path.splitext(filename)
    return ext.lower() in EXCEL_EXTENSIONS


def normalize_sheet_name_for_compare(name: object) -> str:
    if name is None:
        return ""
    return str(name).strip().lower()


def get_worksheet_by_name(workbook, sheet_name: str):
    for sheet in workbook.Worksheets:
        if sheet.Name == sheet_name:
            return sheet
    raise ValueError(f"规则表缺少“{sheet_name}”sheet。")


def get_workbook_sheet_names(workbook) -> list[str]:
    return [sheet.Name for sheet in workbook.Worksheets]


def collect_unique_sheet_names(sheet_name_lists: list[list[str]]) -> list[str]:
    unique_names: list[str] = []
    seen: set[str] = set()

    for sheet_names in sheet_name_lists:
        for sheet_name in sheet_names:
            compare_name = normalize_sheet_name_for_compare(sheet_name)
            if not compare_name or compare_name in seen:
                continue
            seen.add(compare_name)
            unique_names.append(sheet_name)

    return unique_names


def build_rule_table_rows(unique_sheet_names: list[str]) -> list[list[str]]:
    rows = [RULE_TABLE_HEADERS]
    rows.extend([[sheet_name, "", "", ""] for sheet_name in unique_sheet_names])
    return rows


def build_source_path_rows(source_paths: list[str]) -> list[list[str]]:
    rows = [["源文件完整路径"]]
    rows.extend([[source_path] for source_path in source_paths])
    return rows


def get_unique_output_path(output_path: str) -> str:
    if not os.path.exists(output_path):
        return output_path

    base_path, ext = os.path.splitext(output_path)
    counter = 2
    while True:
        candidate = f"{base_path}_{counter}{ext}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def read_rule_values_from_rows(rows: list[list[object]]) -> dict:
    keep_names: list[str] = []
    delete_names: list[str] = []
    exclude_keywords: list[str] = []

    for row in rows[1:]:
        keep_name = _get_stripped_cell(row, 1)
        delete_name = _get_stripped_cell(row, 2)
        exclude_keyword = _get_stripped_cell(row, 3)

        if keep_name:
            keep_names.append(keep_name)
        if delete_name:
            delete_names.append(delete_name)
        if exclude_keyword:
            exclude_keywords.append(exclude_keyword)

    return {
        "keep_names": keep_names,
        "delete_names": delete_names,
        "exclude_keywords": exclude_keywords,
    }


def parse_source_paths_from_rows(rows: list[list[object]]) -> list[str]:
    source_paths: list[str] = []
    for row in rows[1:]:
        source_path = _get_stripped_cell(row, 0)
        if source_path:
            source_paths.append(source_path)
    return source_paths


def infer_delete_mode_from_rule_values(rule_values: dict) -> str | None:
    has_keep_names = any(normalize_sheet_name_for_compare(name) for name in rule_values.get("keep_names", []))
    has_delete_names = any(normalize_sheet_name_for_compare(name) for name in rule_values.get("delete_names", []))

    if has_keep_names and not has_delete_names:
        return "keep"
    if has_delete_names and not has_keep_names:
        return "delete"
    if has_keep_names and has_delete_names:
        return None
    raise ValueError("请在 B 列填写保留表格名，或在 C 列填写删除表格名。")


def should_exclude_file(filename: str, exclude_keywords: list[str]) -> bool:
    normalized_filename = str(filename).lower()
    for keyword in exclude_keywords:
        normalized_keyword = str(keyword).strip().lower()
        if normalized_keyword and normalized_keyword in normalized_filename:
            return True
    return False


def plan_sheets_to_delete(
    workbook_sheet_names: list[str],
    mode: str,
    keep_names: list[str],
    delete_names: list[str],
) -> list[str]:
    if not workbook_sheet_names:
        raise ValueError("工作簿没有工作表。")

    if mode == "keep":
        keep_keys = {normalize_sheet_name_for_compare(name) for name in keep_names if normalize_sheet_name_for_compare(name)}
        if not keep_keys:
            raise ValueError("保留模式下，规则表 B 列不能为空。")
        sheets_to_delete = [
            sheet_name
            for sheet_name in workbook_sheet_names
            if normalize_sheet_name_for_compare(sheet_name) not in keep_keys
        ]
    elif mode == "delete":
        delete_keys = {
            normalize_sheet_name_for_compare(name)
            for name in delete_names
            if normalize_sheet_name_for_compare(name)
        }
        if not delete_keys:
            raise ValueError("删除模式下，规则表 C 列不能为空。")
        sheets_to_delete = [
            sheet_name
            for sheet_name in workbook_sheet_names
            if normalize_sheet_name_for_compare(sheet_name) in delete_keys
        ]
    else:
        raise ValueError("删除模式无效，请选择保留模式或删除模式。")

    if len(sheets_to_delete) >= len(workbook_sheet_names):
        raise ValueError("删除后会导致工作簿没有任何工作表，已跳过。")

    return sheets_to_delete


def copy_source_to_output(source_path: str, output_dir: str) -> str:
    output_path = get_unique_output_path(os.path.join(output_dir, os.path.basename(source_path)))
    shutil.copy2(source_path, output_path)
    return output_path


def normalize_delete_mode(user_input: str) -> str:
    normalized = str(user_input).strip().lower()
    if normalized in {"1", "保留", "keep", "保留模式"}:
        return "keep"
    if normalized in {"2", "删除", "delete", "删除模式"}:
        return "delete"
    raise ValueError("删除模式无效，请输入 1/保留 或 2/删除。")


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


def _get_stripped_cell(row: list[object], index: int) -> str:
    if index >= len(row):
        return ""
    value = row[index]
    if value is None:
        return ""
    return str(value).strip()


def _write_rows_to_sheet(sheet, rows: list[list[str]]) -> None:
    sheet.Range(
        sheet.Cells(1, 1),
        sheet.Cells(len(rows), len(rows[0])),
    ).Value = tuple(tuple(row) for row in rows)


def _normalize_2d_range_values(values, row_count: int, col_count: int) -> list[list[object]]:
    if row_count == 1 and col_count == 1:
        return [[values]]
    if row_count == 1:
        return [list(values[0] if values and isinstance(values[0], tuple) else values)]
    if col_count == 1:
        return [[row[0] if isinstance(row, tuple) else row] for row in values]
    return [list(row) for row in values]


def _read_used_range_rows(sheet) -> list[list[object]]:
    used_range = sheet.UsedRange
    return _normalize_2d_range_values(
        used_range.Value,
        used_range.Rows.Count,
        used_range.Columns.Count,
    )


def _remove_file_quietly(path: str | None) -> None:
    if not path or not os.path.exists(path):
        return
    try:
        os.remove(path)
    except Exception:
        pass


def generate_delete_sheet_rule_table(
    source_paths: list[str],
    output_path: str,
    logger=None,
) -> dict:
    valid_source_paths = [path for path in source_paths if is_excel_file(path)]
    if not valid_source_paths:
        raise ValueError("请选择至少一个 Excel 文件。")
    if not output_path:
        raise ValueError("请选择规则表保存位置。")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    source_workbook = None
    rule_workbook = None
    sheet_name_lists: list[list[str]] = []
    read_success_count = 0
    actual_output_path = get_unique_output_path(os.path.abspath(output_path))

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        for source_path in valid_source_paths:
            source_workbook = None
            try:
                abs_source_path = os.path.abspath(source_path)
                _log(logger, "info", f"正在扫描工作簿：{abs_source_path}")
                source_workbook = excel.Workbooks.Open(
                    abs_source_path,
                    ReadOnly=True,
                    UpdateLinks=0,
                )
                sheet_names = [sheet.Name for sheet in source_workbook.Worksheets]
                sheet_name_lists.append(sheet_names)
                read_success_count += 1
            except Exception as e:
                _log(logger, "error", f"读取失败，已跳过：{source_path}。详细信息：{e}")
            finally:
                if source_workbook is not None:
                    try:
                        source_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    source_workbook = None

        if read_success_count == 0:
            raise RuntimeError("所有 Excel 文件都读取失败，未生成规则表。")

        unique_sheet_names = collect_unique_sheet_names(sheet_name_lists)
        rows = build_rule_table_rows(unique_sheet_names)

        rule_workbook = excel.Workbooks.Add()
        rule_sheet = rule_workbook.Worksheets(1)
        rule_sheet.Name = RULE_SHEET_NAME
        _write_rows_to_sheet(rule_sheet, rows)
        rule_workbook.SaveAs(actual_output_path, FileFormat=EXCEL_RULE_TABLE_FORMAT_XLSX)
        rule_workbook.Close(SaveChanges=False)
        rule_workbook = None

        return {
            "source_file_count": len(valid_source_paths),
            "read_success_count": read_success_count,
            "unique_sheet_count": len(unique_sheet_names),
            "output_path": actual_output_path,
        }

    finally:
        if source_workbook is not None:
            try:
                source_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if rule_workbook is not None:
            try:
                rule_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def generate_temporary_delete_sheet_rule_table(
    source_paths: list[str],
    logger=None,
) -> dict:
    valid_source_paths = [os.path.abspath(path) for path in source_paths if is_excel_file(path)]
    if not valid_source_paths:
        raise ValueError("请选择至少一个 Excel 文件。")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = get_unique_output_path(
        os.path.join(tempfile.gettempdir(), f"批量删除工作表_临时规则表_{timestamp}.xlsx")
    )

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    source_workbook = None
    rule_workbook = None
    sheet_name_lists: list[list[str]] = []
    read_success_count = 0
    old_sheet_count = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        old_sheet_count = excel.SheetsInNewWorkbook

        for source_path in valid_source_paths:
            source_workbook = None
            try:
                _log(logger, "info", f"正在扫描工作簿：{source_path}")
                source_workbook = excel.Workbooks.Open(
                    source_path,
                    ReadOnly=True,
                    UpdateLinks=0,
                )
                sheet_names = [sheet.Name for sheet in source_workbook.Worksheets]
                sheet_name_lists.append(sheet_names)
                read_success_count += 1
            except Exception as e:
                _log(logger, "error", f"读取失败，已跳过：{source_path}。详细信息：{e}")
            finally:
                if source_workbook is not None:
                    try:
                        source_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    source_workbook = None

        if read_success_count == 0:
            raise RuntimeError("所有 Excel 文件都读取失败，未生成规则表。")

        unique_sheet_names = collect_unique_sheet_names(sheet_name_lists)
        excel.SheetsInNewWorkbook = 2
        rule_workbook = excel.Workbooks.Add()
        excel.SheetsInNewWorkbook = old_sheet_count

        if rule_workbook.Worksheets.Count < 2:
            actual_names = ", ".join(get_workbook_sheet_names(rule_workbook))
            raise ValueError(f"临时规则表创建失败：需要 2 个 sheet，当前实际 sheet：{actual_names}")

        rule_sheet = rule_workbook.Worksheets(1)
        source_list_sheet = rule_workbook.Worksheets(2)
        rule_sheet.Name = RULE_SHEET_NAME
        source_list_sheet.Name = SOURCE_LIST_SHEET_NAME

        _write_rows_to_sheet(rule_sheet, build_rule_table_rows(unique_sheet_names))
        _write_rows_to_sheet(source_list_sheet, build_source_path_rows(valid_source_paths))

        try:
            rule_sheet = get_worksheet_by_name(rule_workbook, RULE_SHEET_NAME)
            get_worksheet_by_name(rule_workbook, SOURCE_LIST_SHEET_NAME)
        except ValueError as e:
            actual_names = ", ".join(get_workbook_sheet_names(rule_workbook))
            raise ValueError(f"{e} 当前实际 sheet：{actual_names}") from e

        try:
            rule_sheet.Activate()
            rule_sheet.Range("A1").Select()
        except Exception:
            try:
                rule_sheet.Activate()
            except Exception:
                pass

        rule_workbook.SaveAs(output_path, FileFormat=EXCEL_RULE_TABLE_FORMAT_XLSX)
        rule_workbook.Close(SaveChanges=False)
        rule_workbook = None

        return {
            "source_file_count": len(valid_source_paths),
            "read_success_count": read_success_count,
            "unique_sheet_count": len(unique_sheet_names),
            "output_path": output_path,
        }

    finally:
        if excel is not None and old_sheet_count is not None:
            try:
                excel.SheetsInNewWorkbook = old_sheet_count
            except Exception:
                pass
        if source_workbook is not None:
            try:
                source_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if rule_workbook is not None:
            try:
                rule_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def read_rule_values_from_rule_table(rule_table_path: str) -> dict:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    rule_workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        rule_workbook = excel.Workbooks.Open(
            os.path.abspath(rule_table_path),
            ReadOnly=True,
            UpdateLinks=0,
        )
        rule_sheet = get_worksheet_by_name(rule_workbook, RULE_SHEET_NAME)
        rule_values = read_rule_values_from_rows(_read_used_range_rows(rule_sheet))
        rule_workbook.Close(SaveChanges=False)
        rule_workbook = None
        return rule_values

    finally:
        if rule_workbook is not None:
            try:
                rule_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def execute_batch_delete_sheets(
    source_paths: list[str],
    rule_table_path: str,
    output_dir: str,
    mode: str,
    logger=None,
) -> dict:
    valid_source_paths = [path for path in source_paths if is_excel_file(path)]
    if not valid_source_paths:
        raise ValueError("请选择至少一个 Excel 文件。")
    if not is_excel_file(rule_table_path):
        raise ValueError("请选择有效的规则表 Excel 文件。")
    if not output_dir:
        raise ValueError("请选择输出文件夹。")

    os.makedirs(output_dir, exist_ok=True)
    normalized_mode = normalize_delete_mode(mode)

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    rule_workbook = None
    output_workbook = None

    copied_file_count = 0
    processed_file_count = 0
    skipped_file_count = 0
    failed_file_count = 0
    deleted_sheet_count = 0

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        rule_workbook = excel.Workbooks.Open(
            os.path.abspath(rule_table_path),
            ReadOnly=True,
            UpdateLinks=0,
        )
        try:
            rule_sheet = get_worksheet_by_name(rule_workbook, RULE_SHEET_NAME)
        except ValueError:
            rule_sheet = rule_workbook.Worksheets(1)
        rule_rows = _read_used_range_rows(rule_sheet)
        rule_values = read_rule_values_from_rows(rule_rows)
        rule_workbook.Close(SaveChanges=False)
        rule_workbook = None

        # 提前校验规则列，避免复制文件后才发现整批规则无效。
        if normalized_mode == "keep" and not [
            name for name in rule_values["keep_names"] if normalize_sheet_name_for_compare(name)
        ]:
            raise ValueError("保留模式下，规则表 B 列不能为空。")
        if normalized_mode == "delete" and not [
            name for name in rule_values["delete_names"] if normalize_sheet_name_for_compare(name)
        ]:
            raise ValueError("删除模式下，规则表 C 列不能为空。")

        for source_path in valid_source_paths:
            output_copy_path = None
            output_workbook = None
            filename = os.path.basename(source_path)

            if should_exclude_file(filename, rule_values["exclude_keywords"]):
                skipped_file_count += 1
                _log(logger, "info", f"文件名命中排除关键词，已跳过：{filename}")
                continue

            try:
                output_copy_path = copy_source_to_output(source_path, output_dir)
                copied_file_count += 1
                _log(logger, "info", f"已复制到输出副本：{output_copy_path}")

                output_workbook = excel.Workbooks.Open(
                    os.path.abspath(output_copy_path),
                    UpdateLinks=0,
                )
                workbook_sheet_names = [sheet.Name for sheet in output_workbook.Worksheets]
                sheets_to_delete = plan_sheets_to_delete(
                    workbook_sheet_names,
                    normalized_mode,
                    rule_values["keep_names"],
                    rule_values["delete_names"],
                )

                if not sheets_to_delete:
                    output_workbook.Save()
                    output_workbook.Close(SaveChanges=False)
                    output_workbook = None
                    processed_file_count += 1
                    _log(logger, "info", f"无需删除，已保存输出副本：{output_copy_path}")
                    continue

                for sheet_name in sheets_to_delete:
                    output_workbook.Worksheets(sheet_name).Delete()

                deleted_sheet_count += len(sheets_to_delete)
                output_workbook.Save()
                output_workbook.Close(SaveChanges=False)
                output_workbook = None
                processed_file_count += 1
                _log(logger, "info", f"处理完成：{filename}，删除 {len(sheets_to_delete)} 个工作表。")

            except ValueError as e:
                failed_file_count += 1
                if output_workbook is not None:
                    try:
                        output_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    output_workbook = None
                _remove_file_quietly(output_copy_path)
                if output_copy_path:
                    copied_file_count -= 1
                _log(logger, "error", f"处理失败，已跳过：{filename}。详细信息：{e}")
            except Exception as e:
                failed_file_count += 1
                if output_workbook is not None:
                    try:
                        output_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    output_workbook = None
                _remove_file_quietly(output_copy_path)
                if output_copy_path:
                    copied_file_count -= 1
                _log(logger, "error", f"处理失败，已跳过：{filename}。详细信息：{e}")

        return {
            "source_file_count": len(valid_source_paths),
            "copied_file_count": copied_file_count,
            "processed_file_count": processed_file_count,
            "skipped_file_count": skipped_file_count,
            "failed_file_count": failed_file_count,
            "deleted_sheet_count": deleted_sheet_count,
            "output_dir": os.path.abspath(output_dir),
        }

    finally:
        if rule_workbook is not None:
            try:
                rule_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if output_workbook is not None:
            try:
                output_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def execute_batch_delete_sheets_in_place(
    rule_table_path: str,
    mode: str | None = None,
    logger=None,
) -> dict:
    if not is_excel_file(rule_table_path):
        raise ValueError("请选择有效的规则表 Excel 文件。")

    normalized_mode = normalize_delete_mode(mode) if mode else None

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    rule_workbook = None
    source_workbook = None

    processed_file_count = 0
    skipped_file_count = 0
    failed_file_count = 0
    deleted_sheet_count = 0

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        rule_workbook = excel.Workbooks.Open(
            os.path.abspath(rule_table_path),
            ReadOnly=True,
            UpdateLinks=0,
        )
        rule_sheet = get_worksheet_by_name(rule_workbook, RULE_SHEET_NAME)
        source_list_sheet = get_worksheet_by_name(rule_workbook, SOURCE_LIST_SHEET_NAME)
        rule_values = read_rule_values_from_rows(_read_used_range_rows(rule_sheet))
        source_paths = parse_source_paths_from_rows(_read_used_range_rows(source_list_sheet))
        rule_workbook.Close(SaveChanges=False)
        rule_workbook = None

        if normalized_mode is None:
            inferred_mode = infer_delete_mode_from_rule_values(rule_values)
            if inferred_mode is None:
                raise ValueError("B 列和 C 列都填写了规则，请指定执行模式。")
            normalized_mode = inferred_mode

        valid_source_paths = [path for path in source_paths if is_excel_file(path)]
        if not valid_source_paths:
            raise ValueError("规则表源文件清单中没有有效 Excel 文件。")

        for source_path in valid_source_paths:
            source_workbook = None
            filename = os.path.basename(source_path)

            try:
                if not os.path.exists(source_path):
                    failed_file_count += 1
                    _log(logger, "error", f"源文件不存在，已跳过：{source_path}")
                    continue

                if should_exclude_file(filename, rule_values["exclude_keywords"]):
                    skipped_file_count += 1
                    _log(logger, "info", f"文件名命中排除关键词，已跳过：{filename}")
                    continue

                source_workbook = excel.Workbooks.Open(
                    os.path.abspath(source_path),
                    UpdateLinks=0,
                )
                workbook_sheet_names = [sheet.Name for sheet in source_workbook.Worksheets]
                sheets_to_delete = plan_sheets_to_delete(
                    workbook_sheet_names,
                    normalized_mode,
                    rule_values["keep_names"],
                    rule_values["delete_names"],
                )

                if not sheets_to_delete:
                    source_workbook.Close(SaveChanges=False)
                    source_workbook = None
                    processed_file_count += 1
                    _log(logger, "info", f"无需删除，未保存：{filename}")
                    continue

                for sheet_name in sheets_to_delete:
                    source_workbook.Worksheets(sheet_name).Delete()

                deleted_sheet_count += len(sheets_to_delete)
                source_workbook.Save()
                source_workbook.Close(SaveChanges=False)
                source_workbook = None
                processed_file_count += 1
                _log(logger, "info", f"处理完成：{filename}，删除 {len(sheets_to_delete)} 个工作表。")

            except Exception as e:
                failed_file_count += 1
                if source_workbook is not None:
                    try:
                        source_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    source_workbook = None
                _log(logger, "error", f"处理失败，已跳过：{filename}。详细信息：{e}")

        return {
            "source_file_count": len(valid_source_paths),
            "processed_file_count": processed_file_count,
            "skipped_file_count": skipped_file_count,
            "failed_file_count": failed_file_count,
            "deleted_sheet_count": deleted_sheet_count,
            "mode": normalized_mode,
        }

    finally:
        if rule_workbook is not None:
            try:
                rule_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if source_workbook is not None:
            try:
                source_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
