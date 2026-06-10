import os


INVALID_FILENAME_CHARS = '\\/:*?"<>|'
MAX_SHEET_FILENAME_LENGTH = 100
EXCEL_FILE_FORMAT_XLSX = 51


def clean_filename(name: str | None) -> str:
    if name is None:
        return ""

    cleaned = str(name).strip()
    for char in INVALID_FILENAME_CHARS:
        cleaned = cleaned.replace(char, "")
    return cleaned


def get_safe_sheet_filename(sheet_name: str | None, index: int) -> str:
    cleaned = clean_filename(sheet_name)
    if not cleaned:
        cleaned = f"Sheet_{index}"
    return cleaned[:MAX_SHEET_FILENAME_LENGTH]


def get_unique_name(base_name: str, used_names: set[str]) -> str:
    if base_name not in used_names:
        used_names.add(base_name)
        return base_name

    counter = 2
    while True:
        candidate = f"{base_name}_{counter}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def get_unique_filepath(directory: str, base_name: str, ext: str = ".xlsx") -> str:
    if not ext.startswith("."):
        ext = f".{ext}"

    counter = 1
    while True:
        suffix = "" if counter == 1 else f"_{counter}"
        filepath = os.path.join(directory, f"{base_name}{suffix}{ext}")
        if not os.path.exists(filepath):
            return filepath
        counter += 1


def export_workbook_sheets_to_files(source_path: str, output_dir: str, logger=None) -> list[str]:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    source_workbook = None
    exported_workbook = None
    exported_paths: list[str] = []

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        source_workbook = excel.Workbooks.Open(
            os.path.abspath(source_path),
            ReadOnly=True,
            UpdateLinks=0,
        )

        for index, sheet in enumerate(source_workbook.Worksheets, start=1):
            base_name = get_safe_sheet_filename(sheet.Name, index)
            output_path = get_unique_filepath(output_dir, base_name, ".xlsx")

            if logger:
                logger.info(f"正在导出工作表：{sheet.Name} -> {output_path}")

            sheet.Copy()
            exported_workbook = excel.ActiveWorkbook
            exported_workbook.SaveAs(os.path.abspath(output_path), FileFormat=EXCEL_FILE_FORMAT_XLSX)
            exported_workbook.Close(SaveChanges=False)
            exported_workbook = None
            exported_paths.append(output_path)

        return exported_paths

    except Exception:
        if exported_workbook is not None:
            try:
                exported_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        raise

    finally:
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
