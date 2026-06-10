from .excel_com import get_active_excel


def _sheet_a1_sub_address(sheet_name: str) -> str:
    escaped_name = sheet_name.replace("'", "''")
    return f"'{escaped_name}'!A1"


def _sheet_index_clear_end_row(sheet_count: int, used_last_row: int) -> int:
    return max(1, sheet_count + 1, used_last_row)

def list_sheet_names_to_active_sheet(logger):
    """
    将当前活动工作簿的所有工作表名称列出到活动工作表的 A 列。
    
    返回：
        bool: 执行成功返回 True，否则返回 False
    """
    logger.info("尝试连接当前运行的 Excel 实例...")
    excel = get_active_excel()
    if not excel:
        logger.error("操作失败：未检测到正在运行的 Excel。请先打开 Excel。")
        return False

    try:
        workbook = excel.ActiveWorkbook
        if not workbook:
            logger.error("操作失败：没有打开的工作簿。请先打开或新建一个 Excel 文件。")
            return False

        sheet = excel.ActiveSheet
        if not sheet:
            logger.error("操作失败：没有活动的工作表。")
            return False

        logger.info(f"成功连接工作簿：{workbook.Name}")
        logger.info("正在获取工作表名称...")

        sheet_names = [s.Name for s in workbook.Sheets]
        logger.info(f"共获取到 {len(sheet_names)} 个工作表。")

        logger.info("正在将名称写入活动工作表 A 列...")
        # 写入表头
        sheet.Cells(1, 1).Value = "工作表名称"
        # 写入数据
        for i, name in enumerate(sheet_names):
            sheet.Cells(i + 2, 1).Value = name

        logger.info("写入完成！请在 Excel 中查看。")
        return True

    except Exception as e:
        logger.error(f"操作 Excel 过程中发生错误：{str(e)}")
        return False


def generate_sheet_index_with_links(logger=None) -> dict:
    """
    在当前活动工作表 A 列生成带超链接的工作表目录。
    """
    if logger:
        logger.info("尝试连接当前运行的 Excel 实例...")

    excel = get_active_excel()
    if not excel:
        raise RuntimeError("未检测到正在运行的 Excel。请先打开 Excel。")

    workbook = excel.ActiveWorkbook
    if not workbook:
        raise RuntimeError("没有打开的工作簿。请先打开或新建一个 Excel 文件。")

    active_sheet = excel.ActiveSheet
    if not active_sheet:
        raise RuntimeError("没有活动的工作表。")

    sheets = [sheet for sheet in workbook.Worksheets]

    try:
        used_range = active_sheet.UsedRange
        used_last_row = used_range.Row + used_range.Rows.Count - 1
        clear_end_row = _sheet_index_clear_end_row(len(sheets), used_last_row)
        clear_range = active_sheet.Range(active_sheet.Cells(1, 1), active_sheet.Cells(clear_end_row, 1))
        clear_range.Hyperlinks.Delete()
        clear_range.ClearContents()
    except Exception as e:
        raise RuntimeError(f"清理 A 列旧目录失败：{e}") from e

    active_sheet.Cells(1, 1).Value = "工作表目录"

    try:
        for index, sheet in enumerate(sheets, start=2):
            cell = active_sheet.Cells(index, 1)
            cell.Value = sheet.Name
            active_sheet.Hyperlinks.Add(
                Anchor=cell,
                Address="",
                SubAddress=_sheet_a1_sub_address(sheet.Name),
                TextToDisplay=sheet.Name,
            )
    except Exception as e:
        raise RuntimeError(f"添加工作表目录超链接失败：{e}") from e

    if logger:
        logger.info(f"目录生成成功，共 {len(sheets)} 个工作表。")

    return {"sheet_count": len(sheets)}
