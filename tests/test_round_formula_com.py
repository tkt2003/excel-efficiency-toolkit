import sys
import types

import pytest

from src.excel_efficiency_toolkit.round_formula_ops import (
    round_selected_range_to_two_decimals,
)


class FakeCell:
    """模拟 Excel COM 单元格：记录写入的 Formula，可选注入错误/异常。"""

    def __init__(self, value=None, formula=None, has_formula=False, text=None, raise_on_write=False):
        self.Value = value
        # Excel COM 中即使普通数值单元格 Formula 也有值；这里用显式 formula 覆盖
        self._formula = formula if formula is not None else ("" if value is None else str(value))
        self.HasFormula = has_formula
        self.Text = text if text is not None else ("" if value is None else str(value))
        self._raise_on_write = raise_on_write
        self.written_formula = None

    @property
    def Formula(self):
        return self._formula

    @Formula.setter
    def Formula(self, new_formula):
        if self._raise_on_write:
            raise RuntimeError("模拟写入失败")
        self._formula = new_formula
        self.written_formula = new_formula


class FakeSelection:
    def __init__(self, cells, address="A1:A3"):
        self.Cells = cells
        self._address = address

    def Address(self, absolute_row, absolute_col):
        return self._address


class FakeWorksheet:
    def __init__(self, name="Sheet1"):
        self.Name = name


class FakeWorkbook:
    def __init__(self, name="测试.xlsx"):
        self.Name = name


class FakeExcel:
    def __init__(self, selection, workbook=None, worksheet=None):
        self.Selection = selection
        self.ActiveWorkbook = workbook if workbook is not None else FakeWorkbook()
        self.ActiveSheet = worksheet if worksheet is not None else FakeWorksheet()


def _install_fake_com(monkeypatch, excel):
    """把 fake pythoncom / win32com.client 注入 sys.modules，避免依赖真实 Excel。"""
    fake_pythoncom = types.ModuleType("pythoncom")
    fake_pythoncom.CoInitialize = lambda: None
    fake_pythoncom.CoUninitialize = lambda: None

    def _get_active_object(prog_id):
        assert prog_id == "Excel.Application"
        return excel

    fake_win32com = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")
    fake_client.GetActiveObject = _get_active_object
    fake_win32com.client = fake_client

    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)


def test_wraps_numeric_value_into_round_formula(monkeypatch):
    cell = FakeCell(value=123.456, has_formula=False)
    excel = FakeExcel(FakeSelection([cell]))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert cell.written_formula == "=ROUND(123.456,2)"
    assert result["success_count"] == 1
    assert result["skipped_count"] == 0
    assert result["workbook_name"] == "测试.xlsx"
    assert result["sheet_name"] == "Sheet1"


def test_wraps_existing_formula(monkeypatch):
    cell = FakeCell(has_formula=True, formula="=A1+B1", text="3")
    excel = FakeExcel(FakeSelection([cell]))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert cell.written_formula == "=ROUND(A1+B1,2)"
    assert result["success_count"] == 1
    assert result["skipped_count"] == 0


def test_skips_cell_already_round(monkeypatch):
    cell = FakeCell(has_formula=True, formula="=ROUND(A1+B1,2)", text="3")
    excel = FakeExcel(FakeSelection([cell]))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert cell.written_formula is None  # 未写入
    assert result["success_count"] == 0
    assert result["skipped_count"] == 1


def test_skips_error_cell(monkeypatch):
    cell = FakeCell(value=1, has_formula=False, text="#DIV/0!")
    excel = FakeExcel(FakeSelection([cell]))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert cell.written_formula is None
    assert result["success_count"] == 0
    assert result["skipped_count"] == 1


def test_skips_non_numeric_value(monkeypatch):
    text_cell = FakeCell(value="文本", has_formula=False, text="文本")
    bool_cell = FakeCell(value=True, has_formula=False, text="TRUE")
    empty_cell = FakeCell(value=None, has_formula=False, text="")
    excel = FakeExcel(FakeSelection([text_cell, bool_cell, empty_cell]))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert result["success_count"] == 0
    assert result["skipped_count"] == 3
    assert text_cell.written_formula is None
    assert bool_cell.written_formula is None
    assert empty_cell.written_formula is None


def test_single_cell_error_is_skipped_and_processing_continues(monkeypatch):
    bad_cell = FakeCell(value=1.0, has_formula=False, raise_on_write=True)
    good_cell = FakeCell(value=2.0, has_formula=False)
    excel = FakeExcel(FakeSelection([bad_cell, good_cell]))
    _install_fake_com(monkeypatch, excel)

    logs = []

    class RecordingLogger:
        def error(self, message):
            logs.append(message)

    result = round_selected_range_to_two_decimals(logger=RecordingLogger())

    # 写入失败的单元格计入跳过，后续单元格仍被处理
    assert good_cell.written_formula == "=ROUND(2.0,2)"
    assert result["success_count"] == 1
    assert result["skipped_count"] == 1
    assert len(logs) == 1


def test_mixed_selection_counts(monkeypatch):
    cells = [
        FakeCell(value=10.0, has_formula=False),            # 成功
        FakeCell(has_formula=True, formula="=A1", text="1"),  # 成功
        FakeCell(has_formula=True, formula="=ROUND(A1,2)", text="1"),  # 跳过（已 ROUND）
        FakeCell(value="abc", has_formula=False, text="abc"),  # 跳过（文本）
        FakeCell(value=5, has_formula=False, text="#N/A"),   # 跳过（错误单元格）
    ]
    excel = FakeExcel(FakeSelection(cells))
    _install_fake_com(monkeypatch, excel)

    result = round_selected_range_to_two_decimals()

    assert result["success_count"] == 2
    assert result["skipped_count"] == 3


def test_raises_when_no_active_workbook(monkeypatch):
    excel = FakeExcel(FakeSelection([FakeCell(value=1.0)]))
    excel.ActiveWorkbook = None
    _install_fake_com(monkeypatch, excel)

    with pytest.raises(RuntimeError, match="没有活动工作簿"):
        round_selected_range_to_two_decimals()


def test_raises_when_no_selection(monkeypatch):
    excel = FakeExcel(FakeSelection([FakeCell(value=1.0)]))
    excel.Selection = None
    _install_fake_com(monkeypatch, excel)

    with pytest.raises(RuntimeError, match="没有有效选区"):
        round_selected_range_to_two_decimals()
