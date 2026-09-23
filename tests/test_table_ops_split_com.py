from copy import deepcopy
from pathlib import Path

import pytest

import src.excel_efficiency_toolkit.table_ops as table_ops


class _FakeUsedRange:
    def __init__(self, sheet):
        self.Row = 1
        self.Column = 1
        self.Rows = type("Rows", (), {"Count": len(sheet.values)})()
        self.Columns = type(
            "Columns",
            (),
            {"Count": max((len(row) for row in sheet.values), default=1)},
        )()


class _FakeCell:
    def __init__(self, row, column):
        self.row = row
        self.column = column


class _FakeRange:
    def __init__(self, sheet, start_cell, end_cell):
        self.sheet = sheet
        self.start_cell = start_cell
        self.end_cell = end_cell

    @property
    def Value(self):
        rows = [
            tuple(
                self.sheet.values[row_index - 1][column_index - 1]
                for column_index in range(self.start_cell.column, self.end_cell.column + 1)
            )
            for row_index in range(self.start_cell.row, self.end_cell.row + 1)
        ]
        if len(rows) == 1 and len(rows[0]) == 1:
            return rows[0][0]
        return tuple(rows)


class _FakeDeletedRow:
    def __init__(self, sheet, row_number):
        self.sheet = sheet
        self.row_number = row_number

    def Delete(self):
        self.sheet.delete_history.append(self.row_number)
        del self.sheet.values[self.row_number - 1]


class _FakeWorksheet:
    def __init__(
        self,
        workbook,
        name,
        values,
        *,
        format_marker=None,
        page_setup=None,
        shapes=None,
        freeze_panes=None,
    ):
        self._workbook = workbook
        self._name = name
        self.values = [list(row) for row in values]
        self.format_marker = format_marker
        self.page_setup = page_setup
        self.shapes = shapes or []
        self.freeze_panes = freeze_panes
        self.delete_history = []
        self.copy_of = None

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, value):
        for sheet in self._workbook.Worksheets:
            if sheet is not self and sheet.Name.casefold() == value.casefold():
                raise ValueError(f"工作表名已存在：{value}")
        self._name = value

    @property
    def UsedRange(self):
        return _FakeUsedRange(self)

    def Cells(self, row, column):
        return _FakeCell(row, column)

    def Range(self, start_cell, end_cell):
        return _FakeRange(self, start_cell, end_cell)

    def Rows(self, row_number):
        return _FakeDeletedRow(self, row_number)

    def Copy(self, Before=None, After=None):
        target_after = After if After is not None else Before
        if target_after is None:
            return self._workbook.Application.copy_sheet_to_new_workbook(self)

        copied_sheet = _FakeWorksheet(
            self._workbook,
            f"{self.Name} (2)",
            deepcopy(self.values),
            format_marker=deepcopy(self.format_marker),
            page_setup=deepcopy(self.page_setup),
            shapes=deepcopy(self.shapes),
            freeze_panes=self.freeze_panes,
        )
        copied_sheet.copy_of = self
        self._workbook.copy_calls.append(
            {
                "source": self,
                "after": target_after,
                "copied": copied_sheet,
            }
        )
        self._workbook.Worksheets.insert_after(target_after, copied_sheet)


class _FakeWorksheets:
    def __init__(self, sheets):
        self._sheets = list(sheets)

    def __iter__(self):
        return iter(self._sheets)

    def __call__(self, index):
        if isinstance(index, str):
            for sheet in self._sheets:
                if sheet.Name == index:
                    return sheet
            raise KeyError(index)
        return self._sheets[index - 1]

    @property
    def Count(self):
        return len(self._sheets)

    def insert_after(self, after, sheet):
        self._sheets.insert(self._sheets.index(after) + 1, sheet)


class _FakeWorkbook:
    def __init__(self, sheet_specs, *, name="测试.xlsx", application=None, fail_save=False):
        self.Name = name
        self.Application = application
        self.copy_calls = []
        self.save_calls = []
        self.close_calls = []
        self.fail_save = fail_save
        self.Worksheets = _FakeWorksheets([])
        for spec in sheet_specs:
            self.Worksheets._sheets.append(
                _FakeWorksheet(self, spec["name"], spec["values"])
            )

    def SaveAs(self, path, FileFormat=None):
        self.save_calls.append({"path": path, "file_format": FileFormat})
        if self.fail_save:
            raise RuntimeError("模拟 SaveAs 失败")
        if Path(path).exists():
            raise AssertionError(f"不应覆盖已有文件：{path}")
        Path(path).write_text("fake workbook", encoding="utf-8")

    def Close(self, SaveChanges=False):
        self.close_calls.append(SaveChanges)
        if self.Application is not None:
            self.Application.ActiveWorkbook = self.Application.source_workbook


class _FakeExcel:
    def __init__(self, *, fail_save=False):
        self.ActiveWorkbook = None
        self.source_workbook = None
        self.fail_save = fail_save
        self.created_workbooks = []
        self.copy_calls = []
        self.quit_called = False
        self.Workbooks = _FakeWorksheets([])

    def Quit(self):
        self.quit_called = True

    def copy_sheet_to_new_workbook(self, source_sheet):
        output_workbook = _FakeWorkbook(
            [],
            name=f"输出{len(self.created_workbooks) + 1}.xlsx",
            application=self,
            fail_save=self.fail_save,
        )
        copied_sheet = _FakeWorksheet(
            output_workbook,
            f"{source_sheet.Name} (2)",
            deepcopy(source_sheet.values),
            format_marker=deepcopy(source_sheet.format_marker),
            page_setup=deepcopy(source_sheet.page_setup),
            shapes=deepcopy(source_sheet.shapes),
            freeze_panes=source_sheet.freeze_panes,
        )
        copied_sheet.copy_of = source_sheet
        output_workbook.Worksheets._sheets.append(copied_sheet)
        self.created_workbooks.append(output_workbook)
        self.copy_calls.append(
            {
                "source": source_sheet,
                "after": None,
                "copied": copied_sheet,
                "workbook": output_workbook,
            }
        )
        self.ActiveWorkbook = output_workbook
        return output_workbook


def _build_split_workbook():
    source_values = [
        ["姓名", "城市", "金额"],
        ["甲", "武汉", 100],
        ["乙", "深圳", 200],
        ["丙", "武汉", 300],
        ["丁", "", 400],
        ["戊", " A/B ", 500],
    ]
    workbook = _FakeWorkbook(
        [
            {"name": "源表", "values": source_values},
            {"name": "武汉", "values": [["已有工作表"]]},
        ]
    )
    source_sheet = workbook.Worksheets("源表")
    source_sheet.format_marker = {"font": "等线", "fill": "黄色", "border": "细线"}
    source_sheet.page_setup = {
        "orientation": "横向",
        "fit_to_pages_wide": 1,
        "header": "第 &[Page] 页",
        "footer": "内部资料",
    }
    source_sheet.shapes = ["公司 Logo"]
    source_sheet.freeze_panes = "B2"
    return workbook, source_sheet, source_values


def _build_file_split_workbook(*, fail_save=False):
    source_values = [
        ["姓名", "城市", "金额"],
        ["甲", "武汉", 100],
        ["乙", "深圳", 200],
        ["丙", "武汉", 300],
        ["丁", "", 400],
        ["戊", " A/B ", 500],
    ]
    excel = _FakeExcel(fail_save=fail_save)
    workbook = _FakeWorkbook(
        [{"name": "源表", "values": source_values}],
        application=excel,
    )
    excel.source_workbook = workbook
    excel.ActiveWorkbook = workbook
    source_sheet = workbook.Worksheets("源表")
    source_sheet.format_marker = {"font": "等线", "fill": "黄色", "border": "细线"}
    source_sheet.page_setup = {
        "orientation": "横向",
        "fit_to_pages_wide": 1,
        "header": "第 &[Page] 页",
        "footer": "内部资料",
    }
    source_sheet.shapes = ["公司 Logo"]
    source_sheet.freeze_panes = "B2"
    return excel, workbook, source_sheet, source_values


def test_get_split_rows_to_delete_returns_bottom_up_data_rows_only():
    assert table_ops._get_split_rows_to_delete(
        ["武汉", "深圳", "武汉", None, "A/B"],
        target="武汉",
        data_start_row=2,
    ) == [6, 5, 3]


def test_split_copies_source_sheet_and_deletes_non_target_rows_in_reverse_order(monkeypatch):
    workbook, source_sheet, source_values = _build_split_workbook()
    original_source_values = deepcopy(source_sheet.values)

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_column(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        column_input="B",
        header_row=1,
        data_start_row=2,
    )

    assert result["created_sheet_count"] == 4
    assert result["copied_row_count"] == 5
    assert result["source_sheet_name"] == "源表"

    copied_sheets = [sheet for sheet in workbook.Worksheets if sheet.copy_of is source_sheet]
    assert [sheet.Name for sheet in copied_sheets] == ["武汉_2", "深圳", "空白", "AB"]
    assert [
        (call["source"].Name, call["after"].Name)
        for call in workbook.copy_calls
    ] == [
        ("源表", "武汉"),
        ("源表", "武汉_2"),
        ("源表", "深圳"),
        ("源表", "空白"),
    ]

    expected_rows_by_sheet = {
        "武汉_2": [["甲", "武汉", 100], ["丙", "武汉", 300]],
        "深圳": [["乙", "深圳", 200]],
        "空白": [["丁", "", 400]],
        "AB": [["戊", " A/B ", 500]],
    }
    expected_delete_history = {
        "武汉_2": [6, 5, 3],
        "深圳": [6, 5, 4, 2],
        "空白": [6, 4, 3, 2],
        "AB": [5, 4, 3, 2],
    }

    for sheet in copied_sheets:
        assert sheet.values[0] == source_values[0]
        assert sheet.values[1:] == expected_rows_by_sheet[sheet.Name]
        assert sheet.delete_history == expected_delete_history[sheet.Name]
        assert all(row_number >= 2 for row_number in sheet.delete_history)
        assert sheet.delete_history == sorted(sheet.delete_history, reverse=True)
        assert sheet.format_marker == source_sheet.format_marker
        assert sheet.page_setup == source_sheet.page_setup
        assert sheet.shapes == source_sheet.shapes
        assert sheet.freeze_panes == source_sheet.freeze_panes

    assert source_sheet in list(workbook.Worksheets)
    assert source_sheet.values == original_source_values
    assert source_sheet.delete_history == []


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ('  A/B:C*D?E"F<G>H|I  ', "ABCDEFGHI"),
        ("  name.  ", "name"),
        ("CON", "_CON"),
        ("con.txt", "_con.txt"),
        (None, "空白"),
        ("<>:/\\|?*", "空白"),
        ("A" * 300, "A" * 250),
    ],
)
def test_get_safe_split_filename_handles_windows_filename_rules(value, expected):
    assert table_ops.get_safe_split_filename(value) == expected


def test_get_unique_split_filepath_uses_sequence_without_overwriting_existing_files(tmp_path):
    (tmp_path / "武汉.xlsx").write_text("existing", encoding="utf-8")
    (tmp_path / "武汉_2.xlsx").write_text("existing", encoding="utf-8")

    output_path = table_ops.get_unique_split_filepath(str(tmp_path), "武汉")

    assert output_path == str(tmp_path / "武汉_3.xlsx")
    assert (tmp_path / "武汉.xlsx").read_text(encoding="utf-8") == "existing"


def test_get_unique_split_filepath_reserves_paths_before_they_exist(tmp_path):
    reserved_paths = set()

    first_path = table_ops.get_unique_split_filepath(
        str(tmp_path),
        "同名",
        reserved_paths=reserved_paths,
    )
    second_path = table_ops.get_unique_split_filepath(
        str(tmp_path),
        "同名",
        reserved_paths=reserved_paths,
    )

    assert Path(first_path).name == "同名.xlsx"
    assert Path(second_path).name == "同名_2.xlsx"


def test_split_workbook_sheet_by_column_to_files_copies_source_sheet_and_closes_each_output(
    monkeypatch,
    tmp_path,
):
    excel, workbook, source_sheet, source_values = _build_file_split_workbook()
    original_source_values = deepcopy(source_sheet.values)
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_column_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        column_input="B",
        output_dir=tmp_path,
    )

    assert result == {
        "workbook_name": "测试.xlsx",
        "source_sheet_name": "源表",
        "created_file_count": 4,
        "copied_row_count": 5,
        "output_dir": str(tmp_path),
        "output_paths": [
            str(tmp_path / "武汉.xlsx"),
            str(tmp_path / "深圳.xlsx"),
            str(tmp_path / "空白.xlsx"),
            str(tmp_path / "AB.xlsx"),
        ],
    }

    expected_rows_by_file = {
        "武汉.xlsx": [["甲", "武汉", 100], ["丙", "武汉", 300]],
        "深圳.xlsx": [["乙", "深圳", 200]],
        "空白.xlsx": [["丁", "", 400]],
        "AB.xlsx": [["戊", " A/B ", 500]],
    }
    expected_delete_history = {
        "武汉.xlsx": [6, 5, 3],
        "深圳.xlsx": [6, 5, 4, 2],
        "空白.xlsx": [6, 4, 3, 2],
        "AB.xlsx": [5, 4, 3, 2],
    }

    assert [call["source"] for call in excel.copy_calls] == [source_sheet] * 4
    assert [call["after"] for call in excel.copy_calls] == [None] * 4
    for output_workbook in excel.created_workbooks:
        output_path = Path(output_workbook.save_calls[0]["path"])
        output_sheet = output_workbook.Worksheets(1)
        assert output_sheet.values[0] == source_values[0]
        assert output_sheet.values[1:] == expected_rows_by_file[output_path.name]
        assert output_sheet.delete_history == expected_delete_history[output_path.name]
        assert output_sheet.format_marker == source_sheet.format_marker
        assert output_sheet.page_setup == source_sheet.page_setup
        assert output_sheet.shapes == source_sheet.shapes
        assert output_sheet.freeze_panes == source_sheet.freeze_panes
        assert output_workbook.save_calls[0]["file_format"] == table_ops.EXCEL_FILE_FORMAT_XLSX
        assert output_workbook.close_calls == [False]

    assert [Path(path).name for path in result["output_paths"]] == [
        "武汉.xlsx",
        "深圳.xlsx",
        "空白.xlsx",
        "AB.xlsx",
    ]
    assert excel.ActiveWorkbook is workbook
    assert source_sheet in list(workbook.Worksheets)
    assert source_sheet.values == original_source_values
    assert source_sheet.delete_history == []
    assert workbook.close_calls == []


def test_split_workbook_sheet_by_column_to_files_avoids_existing_output_names(
    monkeypatch,
    tmp_path,
):
    excel, workbook, _, _ = _build_file_split_workbook()
    (tmp_path / "武汉.xlsx").write_text("existing", encoding="utf-8")
    (tmp_path / "武汉_2.xlsx").write_text("existing", encoding="utf-8")
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_column_to_files(
        "D:/测试/源文件.xlsx",
        "源表",
        "B",
        tmp_path,
    )

    assert Path(result["output_paths"][0]).name == "武汉_3.xlsx"
    assert (tmp_path / "武汉.xlsx").read_text(encoding="utf-8") == "existing"
    assert (tmp_path / "武汉_2.xlsx").read_text(encoding="utf-8") == "existing"
    assert excel.created_workbooks[0].close_calls == [False]


def test_split_workbook_sheet_by_column_to_files_closes_temporary_workbook_on_save_failure(
    monkeypatch,
    tmp_path,
):
    excel, workbook, source_sheet, original_values = _build_file_split_workbook(fail_save=True)
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    with pytest.raises(RuntimeError, match="SaveAs 失败"):
        table_ops.split_workbook_sheet_by_column_to_files(
            "D:/测试/源文件.xlsx",
            "源表",
            "B",
            tmp_path,
        )

    assert len(excel.created_workbooks) == 1
    assert excel.created_workbooks[0].close_calls == [False]
    assert source_sheet.values == original_values
    assert source_sheet.delete_history == []
    assert workbook.close_calls == []


@pytest.mark.parametrize(
    ("suffix", "message"),
    [
        (".xlsm", "VBA"),
        (".xls", "旧格式"),
    ],
)
def test_split_workbook_sheet_by_column_to_files_rejects_unsafe_source_formats(
    monkeypatch,
    tmp_path,
    suffix,
    message,
):
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda *args, **kwargs: pytest.fail("不应在格式校验前打开源工作簿"),
    )

    with pytest.raises(ValueError, match=message):
        table_ops.split_workbook_sheet_by_column_to_files(
            f"D:/测试/源文件{suffix}",
            "源表",
            "B",
            tmp_path,
        )


def test_split_workbook_sheet_by_column_to_files_rejects_invalid_output_directory(monkeypatch, tmp_path):
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda *args, **kwargs: pytest.fail("输出目录校验失败后不应打开源工作簿"),
    )

    for output_dir in (None, "   "):
        with pytest.raises(ValueError, match="输出目录不能为空"):
            table_ops.split_workbook_sheet_by_column_to_files(
                "D:/测试/源文件.xlsx",
                "源表",
                "B",
                output_dir,
            )

    with pytest.raises(FileNotFoundError, match="输出目录不存在"):
        table_ops.split_workbook_sheet_by_column_to_files(
            "D:/测试/源文件.xlsx",
            "源表",
            "B",
            tmp_path / "missing",
        )


def test_split_workbook_sheet_by_column_to_files_rejects_non_writable_output_directory(
    monkeypatch,
    tmp_path,
):
    monkeypatch.setattr(table_ops.os, "access", lambda path, mode: False)
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda *args, **kwargs: pytest.fail("不可写目录校验失败后不应打开源工作簿"),
    )

    with pytest.raises(PermissionError, match="输出目录不可写"):
        table_ops.split_workbook_sheet_by_column_to_files(
            "D:/测试/源文件.xlsx",
            "源表",
            "B",
            tmp_path,
        )


# ======================================================================
# 按行拆分底层测试
# ======================================================================


def test_validate_rows_per_part():
    assert table_ops.validate_rows_per_part(1) == 1
    assert table_ops.validate_rows_per_part(5) == 5

    for invalid in (0, -1, -10):
        with pytest.raises(ValueError, match="大于等于 1"):
            table_ops.validate_rows_per_part(invalid)

    for invalid_type in ("1", 1.5, None, True, False, [1]):
        with pytest.raises(ValueError, match="整数"):
            table_ops.validate_rows_per_part(invalid_type)


def test_build_row_split_plans_logic():
    # 5 行数据：2 ~ 6 行
    plans_1 = table_ops.build_row_split_plans(data_start_row=2, last_row=6, rows_per_part=1)
    assert len(plans_1) == 5
    assert [p.name for p in plans_1] == ["数据1", "数据2", "数据3", "数据4", "数据5"]
    assert [(p.keep_start_row, p.keep_end_row) for p in plans_1] == [
        (2, 2),
        (3, 3),
        (4, 4),
        (5, 5),
        (6, 6),
    ]

    plans_2 = table_ops.build_row_split_plans(data_start_row=2, last_row=6, rows_per_part=2)
    assert len(plans_2) == 3
    assert [p.name for p in plans_2] == ["第1批", "第2批", "第3批"]
    assert [(p.keep_start_row, p.keep_end_row) for p in plans_2] == [
        (2, 3),
        (4, 5),
        (6, 6),
    ]

    # rows_per_part 大于数据总行数
    plans_large = table_ops.build_row_split_plans(data_start_row=2, last_row=6, rows_per_part=10)
    assert len(plans_large) == 1
    assert plans_large[0].name == "第1批"
    assert (plans_large[0].keep_start_row, plans_large[0].keep_end_row) == (2, 6)

    # 无数据
    with pytest.raises(RuntimeError, match="没有可拆分的数据"):
        table_ops.build_row_split_plans(data_start_row=5, last_row=4, rows_per_part=1)


def test_get_row_split_rows_to_delete_descending_and_scope():
    # 表头 1~3 行，数据 4~10 行。保留 6~7 行。
    # 需删除：8~10 行从下往上 (10, 9, 8)，以及 4~5 行从下往上 (5, 4)
    rows_to_delete = table_ops._get_row_split_rows_to_delete(
        data_start_row=4,
        last_row=10,
        keep_start_row=6,
        keep_end_row=7,
    )
    assert rows_to_delete == [10, 9, 8, 5, 4]
    assert all(r >= 4 for r in rows_to_delete)
    assert rows_to_delete == sorted(rows_to_delete, reverse=True)


def test_split_workbook_sheet_by_rows_each_row(monkeypatch):
    workbook, source_sheet, source_values = _build_split_workbook()
    original_source_values = deepcopy(source_sheet.values)

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
    )

    assert result["created_sheet_count"] == 5
    assert result["copied_row_count"] == 5
    assert result["rows_per_part"] == 1
    assert result["source_sheet_name"] == "源表"

    copied_sheets = [sheet for sheet in workbook.Worksheets if sheet.copy_of is source_sheet]
    assert [sheet.Name for sheet in copied_sheets] == [
        "数据1",
        "数据2",
        "数据3",
        "数据4",
        "数据5",
    ]

    for index, sheet in enumerate(copied_sheets):
        # 表头保留
        assert sheet.values[0] == source_values[0]
        # 每份只有对应的一行数据
        assert sheet.values[1:] == [source_values[index + 1]]
        assert all(row_number >= 2 for row_number in sheet.delete_history)
        assert sheet.delete_history == sorted(sheet.delete_history, reverse=True)
        assert sheet.format_marker == source_sheet.format_marker
        assert sheet.page_setup == source_sheet.page_setup
        assert sheet.shapes == source_sheet.shapes
        assert sheet.freeze_panes == source_sheet.freeze_panes

    # 源表不被修改
    assert source_sheet in list(workbook.Worksheets)
    assert source_sheet.values == original_source_values
    assert source_sheet.delete_history == []


def test_split_workbook_sheet_by_rows_every_n_rows(monkeypatch):
    workbook, source_sheet, source_values = _build_split_workbook()
    original_source_values = deepcopy(source_sheet.values)

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        rows_per_part=2,
        header_row=1,
        data_start_row=2,
    )

    assert result["created_sheet_count"] == 3
    assert result["copied_row_count"] == 5
    assert result["rows_per_part"] == 2

    copied_sheets = [sheet for sheet in workbook.Worksheets if sheet.copy_of is source_sheet]
    assert [sheet.Name for sheet in copied_sheets] == ["第1批", "第2批", "第3批"]

    assert copied_sheets[0].values[1:] == [source_values[1], source_values[2]]
    assert copied_sheets[1].values[1:] == [source_values[3], source_values[4]]
    assert copied_sheets[2].values[1:] == [source_values[5]]

    for sheet in copied_sheets:
        assert sheet.values[0] == source_values[0]
        assert sheet.delete_history == sorted(sheet.delete_history, reverse=True)

    assert source_sheet.values == original_source_values
    assert source_sheet.delete_history == []


def test_split_workbook_sheet_by_rows_auto_deduplicates_existing_sheet_name(monkeypatch):
    workbook, source_sheet, _ = _build_split_workbook()
    # 预先添加一个同名工作表 "数据1"
    existing = _FakeWorksheet(workbook, "数据1", [["占位"]])
    workbook.Worksheets._sheets.append(existing)

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        rows_per_part=1,
    )

    copied_sheets = [sheet for sheet in workbook.Worksheets if sheet.copy_of is source_sheet]
    assert copied_sheets[0].Name == "数据1_2"


def test_split_workbook_sheet_by_rows_to_files_each_row(monkeypatch, tmp_path):
    excel, workbook, source_sheet, source_values = _build_file_split_workbook()
    original_source_values = deepcopy(source_sheet.values)

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=tmp_path,
        rows_per_part=1,
    )

    assert result == {
        "workbook_name": "测试.xlsx",
        "source_sheet_name": "源表",
        "created_file_count": 5,
        "copied_row_count": 5,
        "rows_per_part": 1,
        "output_dir": str(tmp_path),
        "output_paths": [
            str(tmp_path / "数据1.xlsx"),
            str(tmp_path / "数据2.xlsx"),
            str(tmp_path / "数据3.xlsx"),
            str(tmp_path / "数据4.xlsx"),
            str(tmp_path / "数据5.xlsx"),
        ],
    }

    assert [call["source"] for call in excel.copy_calls] == [source_sheet] * 5
    for index, output_workbook in enumerate(excel.created_workbooks):
        output_sheet = output_workbook.Worksheets(1)
        assert output_sheet.Name == f"数据{index + 1}"
        assert output_sheet.values[0] == source_values[0]
        assert output_sheet.values[1:] == [source_values[index + 1]]
        assert output_sheet.delete_history == sorted(output_sheet.delete_history, reverse=True)
        assert output_sheet.format_marker == source_sheet.format_marker
        assert output_sheet.page_setup == source_sheet.page_setup
        assert output_sheet.shapes == source_sheet.shapes
        assert output_sheet.freeze_panes == source_sheet.freeze_panes
        assert output_workbook.close_calls == [False]

    assert excel.ActiveWorkbook is workbook
    assert source_sheet.values == original_source_values
    assert source_sheet.delete_history == []


def test_split_workbook_sheet_by_rows_to_files_every_n_rows(monkeypatch, tmp_path):
    excel, workbook, source_sheet, source_values = _build_file_split_workbook()

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=tmp_path,
        rows_per_part=2,
    )

    assert result["created_file_count"] == 3
    assert result["rows_per_part"] == 2
    assert [Path(p).name for p in result["output_paths"]] == [
        "第1批.xlsx",
        "第2批.xlsx",
        "第3批.xlsx",
    ]

    for output_workbook in excel.created_workbooks:
        assert output_workbook.close_calls == [False]


def test_split_workbook_sheet_by_rows_to_files_avoids_existing_output_names(monkeypatch, tmp_path):
    excel, workbook, _, _ = _build_file_split_workbook()
    (tmp_path / "数据1.xlsx").write_text("existing", encoding="utf-8")

    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        "D:/测试/源文件.xlsx",
        "源表",
        tmp_path,
        rows_per_part=1,
    )

    assert Path(result["output_paths"][0]).name == "数据1_2.xlsx"
    assert (tmp_path / "数据1.xlsx").read_text(encoding="utf-8") == "existing"


def test_split_workbook_sheet_by_rows_to_files_closes_temporary_workbook_on_save_failure(
    monkeypatch,
    tmp_path,
):
    excel, workbook, source_sheet, original_values = _build_file_split_workbook(fail_save=True)
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    with pytest.raises(RuntimeError, match="SaveAs 失败"):
        table_ops.split_workbook_sheet_by_rows_to_files(
            "D:/测试/源文件.xlsx",
            "源表",
            tmp_path,
            rows_per_part=1,
        )

    assert len(excel.created_workbooks) == 1
    assert excel.created_workbooks[0].close_calls == [False]
    assert source_sheet.values == original_values
    assert source_sheet.delete_history == []


@pytest.mark.parametrize(
    ("suffix", "message"),
    [
        (".xlsm", "VBA"),
        (".xls", "旧格式"),
    ],
)
def test_split_workbook_sheet_by_rows_to_files_rejects_unsafe_source_formats(
    monkeypatch,
    tmp_path,
    suffix,
    message,
):
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda *args, **kwargs: pytest.fail("不应在格式校验前打开源工作簿"),
    )

    with pytest.raises(ValueError, match=message):
        table_ops.split_workbook_sheet_by_rows_to_files(
            f"D:/测试/源文件{suffix}",
            "源表",
            tmp_path,
            rows_per_part=1,
        )


def test_split_workbook_sheet_by_rows_rejects_no_data(monkeypatch):
    workbook, source_sheet, _ = _build_split_workbook()
    monkeypatch.setattr(
        table_ops,
        "_get_or_open_workbook",
        lambda source_path, logger=None: workbook,
    )

    # 设数据起始行大于最后行 (6)
    with pytest.raises(RuntimeError, match="没有可拆分的数据"):
        table_ops.split_workbook_sheet_by_rows(
            source_path="D:/测试/源文件.xlsx",
            source_sheet_name="源表",
            rows_per_part=1,
            header_row=1,
            data_start_row=10,
        )


def test_split_active_sheet_by_rows(monkeypatch):
    workbook, source_sheet, _ = _build_split_workbook()
    excel = _FakeExcel()
    excel.ActiveWorkbook = workbook
    excel.ActiveSheet = source_sheet

    monkeypatch.setattr(table_ops, "get_active_excel", lambda: excel)

    result = table_ops.split_active_sheet_by_rows(rows_per_part=2)
    assert result["created_sheet_count"] == 3


def test_split_by_rows_with_name_column_generates_named_files(monkeypatch, tmp_path):
    source_values = [
        ["序号", "姓名", "金额"],
        [1, "张三", 100],
        [2, "李四", 200],
    ]
    excel = _FakeExcel()
    workbook = _FakeWorkbook(
        [{"name": "源表", "values": source_values}],
        application=excel,
    )
    excel.source_workbook = workbook
    excel.ActiveWorkbook = workbook
    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=str(tmp_path),
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        name_column="B",
    )

    assert result["created_file_count"] == 2
    assert any(p.endswith("张三.xlsx") for p in result["output_paths"])
    assert any(p.endswith("李四.xlsx") for p in result["output_paths"])


def test_split_by_rows_chinese_filename(monkeypatch, tmp_path):
    source_values = [
        ["工号", "部门姓名", "绩效"],
        ["E001", "财务部/王主管", "A"],
        ["E002", "  市场运营部*李专员  ", "B"],
    ]
    excel = _FakeExcel()
    workbook = _FakeWorkbook(
        [{"name": "源表", "values": source_values}],
        application=excel,
    )
    excel.source_workbook = workbook
    excel.ActiveWorkbook = workbook

    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=str(tmp_path),
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        name_column="B",
    )

    assert result["created_file_count"] == 2
    # 非法字符 / 和 * 应被清理
    assert any("财务部王主管.xlsx" in p for p in result["output_paths"])
    assert any("市场运营部李专员.xlsx" in p for p in result["output_paths"])


def test_split_by_rows_duplicate_name_column_handling(monkeypatch, tmp_path):
    source_values = [
        ["工号", "姓名", "金额"],
        [1, "张三", 100],
        [2, "张三", 200],
        [3, "张三", 300],
    ]
    excel = _FakeExcel()
    workbook = _FakeWorkbook(
        [{"name": "源表", "values": source_values}],
        application=excel,
    )
    excel.source_workbook = workbook
    excel.ActiveWorkbook = workbook

    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=str(tmp_path),
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        name_column="B",
    )

    assert result["created_file_count"] == 3
    file_names = [Path(p).name for p in result["output_paths"]]
    assert file_names == ["张三.xlsx", "张三_2.xlsx", "张三_3.xlsx"]


def test_split_by_rows_cancellation_cleans_up_excel(monkeypatch, tmp_path):
    source_values = [
        ["工号", "姓名"],
        [1, "张三"],
        [2, "李四"],
        [3, "王五"],
    ]
    excel = _FakeExcel()
    workbook = _FakeWorkbook(
        [{"name": "源表", "values": source_values}],
        application=excel,
    )
    excel.source_workbook = workbook
    excel.ActiveWorkbook = workbook
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=workbook,
            owns_excel=True,
            owns_workbook=True,
            com_initialized=False,
        )
    )

    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    cancel_token = table_ops.CancellationToken()

    def progress_callback(current, total, name):
        if current >= 1:
            cancel_token.cancel()

    result = table_ops.split_workbook_sheet_by_rows_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        output_dir=str(tmp_path),
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        cancel_token=cancel_token,
        progress_callback=progress_callback,
    )

    assert result.get("cancelled") is True
    assert result["created_file_count"] == 1
    assert excel.quit_called is True


def test_split_by_column_cancellation_cleans_up_excel(monkeypatch, tmp_path):
    excel, workbook, source_sheet, _ = _build_file_split_workbook()
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=workbook,
            owns_excel=True,
            owns_workbook=True,
            com_initialized=False,
        )
    )
    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    cancel_token = table_ops.CancellationToken()

    def progress_callback(current, total, name):
        if current >= 1:
            cancel_token.cancel()

    result = table_ops.split_workbook_sheet_by_column_to_files(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        column_input="B",
        output_dir=str(tmp_path),
        header_row=1,
        data_start_row=2,
        cancel_token=cancel_token,
        progress_callback=progress_callback,
    )

    assert result.get("cancelled") is True
    assert result["created_file_count"] == 1
    assert excel.quit_called is True


def test_cleanup_excel_session_protects_existing_user_excel(monkeypatch):
    active_excel = _FakeExcel()
    active_excel.Hwnd = 123456
    workbook = _FakeWorkbook([], application=active_excel)

    monkeypatch.setattr(table_ops, "get_active_excel", lambda: active_excel)

    table_ops._cleanup_excel_session(active_excel, workbook)

    # 严禁退出用户已有的 Excel
    assert active_excel.quit_called is False
    assert workbook.close_calls == []


def test_cleanup_excel_session_quits_tool_created_excel():
    tool_excel = _FakeExcel()
    tool_excel.Hwnd = 999999
    excel_id = table_ops._get_excel_id(tool_excel)
    workbook = _FakeWorkbook([], application=tool_excel)
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=tool_excel,
            workbook=workbook,
            owns_excel=True,
            owns_workbook=True,
            com_initialized=False,
        )
    )

    table_ops._cleanup_excel_session(tool_excel, workbook)

    # 工具独立创建的 Excel 正常退出
    assert tool_excel.quit_called is True
    assert workbook.close_calls == [False]
    assert excel_id not in table_ops._TOOL_CREATED_EXCEL_IDS


def test_release_workbook_session_cleans_up_opened_file(monkeypatch):
    excel = _FakeExcel()
    excel.Hwnd = 888888
    excel_id = table_ops._get_excel_id(excel)
    table_ops._TOOL_CREATED_EXCEL_IDS.add(excel_id)

    workbook = _FakeWorkbook([], application=excel)
    fake_path = "D:/test_data/temp_source.xlsx"
    workbook.FullName = fake_path
    excel.Workbooks._sheets = [workbook]

    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=workbook,
            owns_excel=True,
            owns_workbook=True,
            com_initialized=False,
        )
    )

    table_ops.release_workbook_session(fake_path)

    assert workbook.close_calls == [False]
    assert excel.quit_called is True


def test_registered_user_excel_and_user_workbook_are_never_closed_or_quit():
    excel = _FakeExcel()
    workbook = _FakeWorkbook([], application=excel)
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=workbook,
            owns_excel=False,
            owns_workbook=False,
            com_initialized=False,
        )
    )

    table_ops._cleanup_excel_session(workbook=workbook)

    assert workbook.close_calls == []
    assert excel.quit_called is False


def test_registered_user_excel_and_tool_workbook_only_close_tool_workbook():
    excel = _FakeExcel()
    user_workbook = _FakeWorkbook([], name="用户.xlsx", application=excel)
    tool_workbook = _FakeWorkbook([], name="工具.xlsx", application=excel)
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=user_workbook,
            owns_excel=False,
            owns_workbook=False,
            com_initialized=False,
        )
    )
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=tool_workbook,
            owns_excel=False,
            owns_workbook=True,
            com_initialized=False,
        )
    )

    table_ops._cleanup_excel_session(workbook=tool_workbook)

    assert tool_workbook.close_calls == [False]
    assert user_workbook.close_calls == []
    assert excel.quit_called is False


def test_get_or_open_workbook_records_user_workbook_ownership_explicitly(monkeypatch, tmp_path):
    source_path = tmp_path / "source.xlsx"
    source_path.write_text("fixture", encoding="utf-8")
    excel = _FakeExcel()
    excel.Hwnd = 101
    workbook = _FakeWorkbook([], application=excel)
    workbook.FullName = str(source_path)
    excel.Workbooks._sheets = [workbook]
    monkeypatch.setattr(table_ops, "get_active_excel", lambda: excel)

    opened = table_ops._get_or_open_workbook(str(source_path))
    session = table_ops._find_workbook_session(opened)

    assert opened is workbook
    assert session is not None
    assert session.owns_excel is False
    assert session.owns_workbook is False

    table_ops._cleanup_excel_session(workbook=opened)
    assert workbook.close_calls == []
    assert excel.quit_called is False


def test_get_or_open_workbook_marks_only_tool_opened_workbook_owned(monkeypatch, tmp_path):
    source_path = tmp_path / "tool-opened.xlsx"
    source_path.write_text("fixture", encoding="utf-8")
    excel = _FakeExcel()
    excel.Hwnd = 202

    def open_workbook(path, UpdateLinks=0):
        workbook = _FakeWorkbook([], name="工具打开.xlsx", application=excel)
        workbook.FullName = str(path)
        excel.Workbooks._sheets.append(workbook)
        return workbook

    excel.Workbooks.Open = open_workbook
    monkeypatch.setattr(table_ops, "get_active_excel", lambda: excel)

    opened = table_ops._get_or_open_workbook(str(source_path))
    session = table_ops._find_workbook_session(opened)
    opened_again = table_ops._get_or_open_workbook(str(source_path))

    assert session is not None
    assert opened_again is opened
    assert session.owns_excel is False
    assert session.owns_workbook is True

    table_ops._cleanup_excel_session(workbook=opened)
    assert opened.close_calls == [False]
    assert excel.quit_called is False


def test_header_read_uses_one_bulk_range_and_resolves_merged_titles():
    class HeaderCell:
        def __init__(self, row, column, merged=False, top_left=None):
            self.Row = row
            self.Column = column
            self.row = row
            self.column = column
            self.MergeCells = merged
            self.MergeArea = top_left

    class HeaderMergeArea:
        def __init__(self, top_left):
            self.top_left = top_left

        def Cells(self, row, column):
            return self.top_left

    class HeaderSheet:
        Name = "源表"

        def __init__(self):
            self.values = [
                ["部门", None, "姓名"],
                ["编号", "金额", None],
                [1, 100, "张三"],
            ]
            self.range_calls = 0
            top_left = HeaderCell(1, 1)
            self.merged_cell = HeaderCell(1, 2, merged=True, top_left=HeaderMergeArea(top_left))

        @property
        def UsedRange(self):
            return _FakeUsedRange(self)

        def Cells(self, row, column):
            if (row, column) == (1, 2):
                return self.merged_cell
            return HeaderCell(row, column)

        def Range(self, start_cell, end_cell):
            self.range_calls += 1
            return _FakeRange(self, start_cell, end_cell)

    sheet = HeaderSheet()

    columns = table_ops._read_sheet_header_columns_from_sheet(sheet, header_row=2)

    assert columns == [("A", "部门 / 编号"), ("B", "部门 / 金额"), ("C", "姓名")]
    assert sheet.range_calls == 1


def test_sheet_split_keeps_user_workbook_and_excel_alive(monkeypatch):
    excel, workbook, _, _ = _build_file_split_workbook()
    table_ops._register_workbook_session(
        table_ops.WorkbookSession(
            excel=excel,
            workbook=workbook,
            owns_excel=False,
            owns_workbook=False,
            com_initialized=False,
        )
    )
    monkeypatch.setattr(table_ops, "_get_or_open_workbook", lambda *args, **kwargs: workbook)

    result = table_ops.split_workbook_sheet_by_column(
        source_path="D:/测试/源文件.xlsx",
        source_sheet_name="源表",
        column_input="B",
    )

    assert result["created_sheet_count"] == 4
    assert workbook.close_calls == []
    assert excel.quit_called is False

