import os

from .excel_com import get_active_excel


def _sheet_a1_sub_address(sheet_name: str) -> str:
    escaped_name = sheet_name.replace("'", "''")
    return f"'{escaped_name}'!A1"


def _sheet_names_for_index(all_sheet_names: list[str], index_sheet_name: str) -> list[str]:
    index_name = index_sheet_name.lower()
    return [name for name in all_sheet_names if name.lower() != index_name]


def _sheet_index_clear_end_row(sheet_count: int, used_last_row: int) -> int:
    return max(1, sheet_count + 1, used_last_row)


def _log(logger, level: str, message: str) -> None:
    if logger:
        getattr(logger, level)(message)


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


def _get_or_create_index_sheet(workbook, index_sheet_name: str):
    for sheet in workbook.Worksheets:
        if sheet.Name.lower() == index_sheet_name.lower():
            return sheet

    sheet = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
    sheet.Name = index_sheet_name
    return sheet

def generate_sheet_index_sheet_with_links(
    source_path: str,
    index_sheet_name: str = "工作表目录",
    logger=None,
) -> dict:
    """
    在所选工作簿中生成专门的工作表目录 sheet，并添加跳转链接。
    """
    workbook = _get_or_open_workbook(source_path, logger)
    index_sheet = _get_or_create_index_sheet(workbook, index_sheet_name)

    all_sheet_names = [sheet.Name for sheet in workbook.Worksheets]
    indexed_sheet_names = _sheet_names_for_index(all_sheet_names, index_sheet.Name)

    try:
        used_range = index_sheet.UsedRange
        used_last_row = used_range.Row + used_range.Rows.Count - 1
        clear_end_row = _sheet_index_clear_end_row(len(indexed_sheet_names), used_last_row)
        clear_range = index_sheet.Range(index_sheet.Cells(1, 1), index_sheet.Cells(clear_end_row, 1))
        clear_range.Hyperlinks.Delete()
        clear_range.ClearContents()
    except Exception as e:
        raise RuntimeError(f"清理目录 sheet 的 A 列旧目录失败：{e}") from e

    index_sheet.Cells(1, 1).Value = "工作表目录"

    try:
        for index, sheet_name in enumerate(indexed_sheet_names, start=2):
            cell = index_sheet.Cells(index, 1)
            cell.Value = sheet_name
            index_sheet.Hyperlinks.Add(
                Anchor=cell,
                Address="",
                SubAddress=_sheet_a1_sub_address(sheet_name),
                TextToDisplay=sheet_name,
            )
    except Exception as e:
        raise RuntimeError(f"添加工作表目录超链接失败：{e}") from e

    _log(logger, "info", f"目录生成成功，共收录 {len(indexed_sheet_names)} 个工作表。")
    return {
        "workbook_name": workbook.Name,
        "index_sheet_name": index_sheet.Name,
        "sheet_count": len(indexed_sheet_names),
    }
