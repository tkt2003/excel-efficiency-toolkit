from openpyxl import Workbook

from src.excel_efficiency_toolkit.workbook_drill_ops import (
    build_multi_file_result_headers,
    build_single_file_result_headers,
    build_unique_result_sheet_name,
    expand_range_addresses,
    range_value_to_address_map,
    should_skip_history_result_sheet,
    stringify_excel_error_value,
    write_multi_file_result_sheet,
    write_single_file_result_sheet,
    write_single_workbook_drill_result_to_com_sheet,
)


def test_expand_range_addresses_single_cell():
    assert expand_range_addresses("C5") == ["C5"]


def test_expand_range_addresses_area():
    assert expand_range_addresses("B2:C3") == ["B2", "C2", "B3", "C3"]


def test_build_unique_result_sheet_name_adds_sequence():
    assert (
        build_unique_result_sheet_name(["Sheet1", "数据穿透查询结果", "数据穿透查询结果_2"])
        == "数据穿透查询结果_3"
    )


def test_should_skip_history_result_sheet():
    assert should_skip_history_result_sheet("数据穿透查询结果")
    assert should_skip_history_result_sheet("数据穿透查询结果_2")
    assert not should_skip_history_result_sheet("数据穿透结果")


def test_stringify_excel_error_value_keeps_readable_error_text():
    assert stringify_excel_error_value("#VALUE!") == "#VALUE!"
    assert stringify_excel_error_value(2042) == "#N/A"
    assert stringify_excel_error_value("普通文本") == "普通文本"


def test_single_file_headers_and_sheet_write():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据穿透查询结果"

    write_single_file_result_sheet(
        sheet=sheet,
        base_sheet_name="货币资金",
        range_address="B2:C3",
        created_at_text="2026-06-12 10:11:12",
        records=[
            {
                "sheet_name": "Sheet1",
                "visible_status": "可见",
                "values_by_address": {"B2": 1, "C2": 2, "B3": 3, "C3": "#N/A"},
            }
        ],
    )

    assert sheet["A1"].value == "基准工作表：货币资金    基准选区：B2:C3    生成时间：2026-06-12 10:11:12"
    assert [sheet.cell(row=3, column=index).value for index in range(1, 7)] == [
        "工作表名",
        "工作表可见状态",
        "B2",
        "C2",
        "B3",
        "C3",
    ]
    assert [sheet.cell(row=4, column=index).value for index in range(1, 7)] == [
        "Sheet1",
        "可见",
        1,
        2,
        3,
        "#N/A",
    ]


def test_multi_file_headers_and_sheet_write():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据穿透查询结果"

    write_multi_file_result_sheet(
        sheet=sheet,
        records=[
            {
                "source_file_name": "a.xlsx",
                "source_file_path": r"D:\tmp\a.xlsx",
                "target_sheet_name": "货币资金",
                "values_by_address": {"B2": 1, "C2": 2, "B3": 3, "C3": 4},
                "status": "成功",
                "message": "读取成功",
            }
        ],
        range_address="B2:C3",
    )

    assert [cell.value for cell in sheet[1]] == [
        "源文件名",
        "源文件路径",
        "工作表名",
        "B2",
        "C2",
        "B3",
        "C3",
        "状态",
        "说明",
    ]
    assert [sheet.cell(row=2, column=index).value for index in range(1, 10)] == [
        "a.xlsx",
        r"D:\tmp\a.xlsx",
        "货币资金",
        1,
        2,
        3,
        4,
        "成功",
        "读取成功",
    ]


def test_single_workbook_result_can_write_to_com_sheet_like_object():
    sheet = _FakeComSheet()

    write_single_workbook_drill_result_to_com_sheet(
        sheet=sheet,
        base_sheet_name="Sheet1",
        range_address="C3:D4",
        created_at_text="2026-06-12 12:00:00",
        records=[
            {
                "sheet_name": "Sheet1",
                "visible_status": "可见",
                "values_by_address": {"C3": 1, "D3": 2, "C4": 3, "D4": 4},
            }
        ],
    )

    assert sheet.values[(1, 1)] == "基准工作表：Sheet1    基准选区：C3:D4    生成时间：2026-06-12 12:00:00"
    assert [sheet.values[(3, index)] for index in range(1, 7)] == [
        "工作表名",
        "工作表可见状态",
        "C3",
        "D3",
        "C4",
        "D4",
    ]
    assert [sheet.values[(4, index)] for index in range(1, 7)] == [
        "Sheet1",
        "可见",
        1,
        2,
        3,
        4,
    ]


def test_range_value_to_address_map_supports_single_and_area():
    assert range_value_to_address_map("C5", 8) == {"C5": 8}
    assert range_value_to_address_map("B2:C3", ((1, 2), (3, 4))) == {
        "B2": 1,
        "C2": 2,
        "B3": 3,
        "C3": 4,
    }


def test_build_headers_helpers():
    assert build_single_file_result_headers("C5") == ["工作表名", "工作表可见状态", "C5"]
    assert build_multi_file_result_headers("C5") == ["源文件名", "源文件路径", "工作表名", "C5", "状态", "说明"]


class _FakeComCell:
    def __init__(self, values, row_index, column_index):
        self._values = values
        self._key = (row_index, column_index)

    @property
    def Value(self):
        return self._values.get(self._key)

    @Value.setter
    def Value(self, value):
        self._values[self._key] = value


class _FakeComSheet:
    def __init__(self):
        self.values = {}

    def Cells(self, row_index, column_index):
        return _FakeComCell(self.values, row_index, column_index)


# ---------------------------------------------------------------------------
# Excel COM 活动会话读取（原 app.py 数据穿透逻辑）的假对象测试
# ---------------------------------------------------------------------------
import sys
import types

import pytest

from src.excel_efficiency_toolkit.workbook_drill_ops import (
    execute_single_workbook_drill,
    get_active_drill_context,
    get_selection_range_address,
    read_range_values_from_com_sheet,
)


class _FakeAreas:
    def __init__(self, count):
        self.Count = count


class _FakeSelectionCallableAddress:
    """Address 是方法的 COM 形态。"""

    def __init__(self, address, areas_count=1):
        self.Areas = _FakeAreas(areas_count)
        self._address = address

    def Address(self, absolute_row, absolute_col):
        return self._address


class _FakeSelectionPlainAddress:
    """Address 是属性的 COM 形态。"""

    def __init__(self, address, areas_count=1):
        self.Areas = _FakeAreas(areas_count)
        self.Address = address


class _FakeRange:
    def __init__(self, value):
        self.Value = value


class _FakeDrillSheet:
    def __init__(self, name, visible=-1, range_value=None, raise_on_read=False):
        self.Name = name
        self.Visible = visible
        self._range_value = range_value
        self._raise_on_read = raise_on_read
        self.values = {}

    def Range(self, address):
        if self._raise_on_read:
            raise RuntimeError("模拟读取失败")
        return _FakeRange(self._range_value)

    def Cells(self, row_index, column_index):
        return _FakeComCell(self.values, row_index, column_index)


class _FakeWorksheets:
    def __init__(self, sheets):
        self._sheets = list(sheets)

    def __iter__(self):
        return iter(self._sheets)

    def __call__(self, index):
        return self._sheets[index - 1]

    @property
    def Count(self):
        return len(self._sheets)

    def Add(self, After=None):
        new_sheet = _FakeDrillSheet("未命名")
        self._sheets.append(new_sheet)
        return new_sheet


class _FakeDrillWorkbook:
    def __init__(self, sheets, name="测试.xlsx", full_name=None, path="D:\\数据"):
        self.Worksheets = _FakeWorksheets(sheets)
        self.Name = name
        self.FullName = full_name if full_name is not None else f"{path}\\{name}"
        self.Path = path


class _FakeDrillExcel:
    def __init__(self, workbook, selection, active_sheet):
        self.ActiveWorkbook = workbook
        self.Selection = selection
        self.ActiveSheet = active_sheet


class _RecordingLogger:
    def __init__(self):
        self.infos = []
        self.errors = []

    def info(self, message):
        self.infos.append(message)

    def error(self, message):
        self.errors.append(message)


def _install_fake_com(monkeypatch, excel):
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


def test_get_selection_range_address_supports_callable_address():
    excel = types.SimpleNamespace(Selection=_FakeSelectionCallableAddress("$C$5"))
    assert get_selection_range_address(excel) == "C5"


def test_get_selection_range_address_supports_plain_address_attribute():
    excel = types.SimpleNamespace(Selection=_FakeSelectionPlainAddress("$B$2:$C$3"))
    assert get_selection_range_address(excel) == "B2:C3"


def test_get_selection_range_address_rejects_multi_area():
    excel = types.SimpleNamespace(Selection=_FakeSelectionCallableAddress("A1", areas_count=2))
    with pytest.raises(RuntimeError, match="多区域"):
        get_selection_range_address(excel)


def test_get_selection_range_address_rejects_missing_selection():
    excel = types.SimpleNamespace(Selection=None)
    with pytest.raises(RuntimeError, match="没有可识别的选区"):
        get_selection_range_address(excel)


def test_read_range_values_from_com_sheet_error_fills_message_per_address():
    sheet = _FakeDrillSheet("Sheet1", raise_on_read=True)
    logger = _RecordingLogger()

    values = read_range_values_from_com_sheet(sheet, "B2:C3", logger=logger)

    assert set(values) == {"B2", "C2", "B3", "C3"}
    assert all(value.startswith("读取异常：") for value in values.values())
    assert len(logger.errors) == 1


def test_execute_single_workbook_drill_skips_hidden_and_history_sheets():
    visible_sheet = _FakeDrillSheet("Sheet1", visible=-1, range_value=42)
    hidden_sheet = _FakeDrillSheet("Sheet2", visible=0, range_value=1)
    history_sheet = _FakeDrillSheet("数据穿透查询结果", visible=-1, range_value=2)
    workbook = _FakeDrillWorkbook([visible_sheet, hidden_sheet, history_sheet])
    excel = _FakeDrillExcel(workbook, _FakeSelectionCallableAddress("$C$5"), visible_sheet)

    result = execute_single_workbook_drill(excel)

    assert result["visible_sheet_count"] == 1
    assert result["range_address"] == "C5"
    assert result["sheet_name"] == "Sheet1"
    # 已有同名历史结果表，新表名追加序号
    assert result["result_sheet_name"] == "数据穿透查询结果_2"

    result_sheet = workbook.Worksheets(workbook.Worksheets.Count)
    assert result_sheet.Name == "数据穿透查询结果_2"
    assert [result_sheet.values[(3, index)] for index in range(1, 4)] == ["工作表名", "工作表可见状态", "C5"]
    assert [result_sheet.values[(4, index)] for index in range(1, 4)] == ["Sheet1", "可见", 42]


def test_execute_single_workbook_drill_raises_without_active_workbook():
    excel = _FakeDrillExcel(None, _FakeSelectionCallableAddress("A1"), None)
    with pytest.raises(RuntimeError, match="没有活动工作簿"):
        execute_single_workbook_drill(excel)


def test_get_active_drill_context_rejects_unsaved_workbook(monkeypatch):
    sheet = _FakeDrillSheet("Sheet1")
    workbook = _FakeDrillWorkbook([sheet], name="新工作簿", full_name="", path="")
    excel = _FakeDrillExcel(workbook, _FakeSelectionCallableAddress("A1"), sheet)
    _install_fake_com(monkeypatch, excel)

    with pytest.raises(RuntimeError, match="尚未保存"):
        get_active_drill_context(require_saved_workbook=True)


def test_get_active_drill_context_returns_full_context(monkeypatch):
    sheet = _FakeDrillSheet("汇总表")
    workbook = _FakeDrillWorkbook([sheet], name="底稿.xlsx", path="D:\\审计")
    excel = _FakeDrillExcel(workbook, _FakeSelectionCallableAddress("$B$2:$C$3"), sheet)
    _install_fake_com(monkeypatch, excel)

    context = get_active_drill_context(require_saved_workbook=True)

    assert context["workbook_name"] == "底稿.xlsx"
    assert context["workbook_path_text"] == "D:\\审计\\底稿.xlsx"
    assert context["output_dir"] == "D:\\审计"
    assert context["sheet_name"] == "汇总表"
    assert context["range_address"] == "B2:C3"
