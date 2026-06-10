import os


EXCEL_RULE_TABLE_FORMAT_XLSX = 51
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
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


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


def _write_rows_to_sheet(sheet, rows: list[list[str]]) -> None:
    sheet.Range(
        sheet.Cells(1, 1),
        sheet.Cells(len(rows), len(rows[0])),
    ).Value = tuple(tuple(row) for row in rows)


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
        rule_sheet.Name = "批量删除规则"
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
