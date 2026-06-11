import re
from numbers import Number


ROUND_FORMULA_PATTERN = re.compile(r"^=\s*ROUND\s*\(", re.IGNORECASE)


def build_round_formula_from_value(value) -> str:
    return f"=ROUND({_format_number_for_formula(value)},2)"


def build_round_formula_from_formula(formula: str) -> str:
    text = str(formula or "").strip()
    if is_already_round_formula(text):
        return text
    inner_formula = text[1:] if text.startswith("=") else text
    return f"=ROUND({inner_formula},2)"


def is_already_round_formula(formula: str) -> bool:
    return bool(ROUND_FORMULA_PATTERN.match(str(formula or "").strip()))


def should_skip_cell_value(value) -> bool:
    if value is None or value == "":
        return True
    if isinstance(value, bool):
        return True
    if isinstance(value, str):
        return True
    return not isinstance(value, Number)


def round_selected_range_to_two_decimals(logger=None) -> dict:
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

    pythoncom.CoInitialize()
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception as e:
            raise RuntimeError("未检测到正在运行的 Excel，请先打开工作簿并选中区域后重试。") from e

        workbook = excel.ActiveWorkbook
        if workbook is None:
            raise RuntimeError("Excel 中没有活动工作簿，请先打开工作簿并选中区域后重试。")

        worksheet = excel.ActiveSheet
        selection = excel.Selection
        if selection is None:
            raise RuntimeError("Excel 中没有有效选区，请先选中需要处理的区域后重试。")

        context = {
            "workbook_name": str(workbook.Name),
            "sheet_name": str(worksheet.Name),
            "selection_address": _get_selection_address(selection),
            "success_count": 0,
            "skipped_count": 0,
        }

        for cell in selection.Cells:
            try:
                if _is_error_cell(cell):
                    context["skipped_count"] += 1
                    continue

                if bool(cell.HasFormula):
                    formula = str(cell.Formula)
                    if is_already_round_formula(formula):
                        context["skipped_count"] += 1
                        continue
                    cell.Formula = build_round_formula_from_formula(formula)
                    context["success_count"] += 1
                    continue

                value = cell.Value
                if should_skip_cell_value(value):
                    context["skipped_count"] += 1
                    continue
                cell.Formula = build_round_formula_from_value(value)
                context["success_count"] += 1
            except Exception as e:
                context["skipped_count"] += 1
                _log(logger, "error", f"单元格处理失败，已跳过：{e}")

        return context
    finally:
        pythoncom.CoUninitialize()


def _format_number_for_formula(value) -> str:
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return str(value)


def _get_selection_address(selection) -> str:
    try:
        return str(selection.Address(False, False))
    except Exception:
        return str(selection.Address).replace("$", "")


def _is_error_cell(cell) -> bool:
    try:
        text = str(cell.Text)
        return text.startswith("#")
    except Exception:
        return False


def _log(logger, level: str, message: str) -> None:
    if logger is None:
        return
    log_func = getattr(logger, level, None)
    if log_func:
        log_func(message)
