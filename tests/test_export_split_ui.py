# -*- coding: utf-8 -*-
from types import SimpleNamespace
from unittest.mock import Mock

from src.excel_efficiency_toolkit.ui.features import export_split as export_split_ui
from src.excel_efficiency_toolkit.ui.features.export_split import ExportSplitController


def _controller(
    *,
    sheet_name="Sheet1",
    column="B",
    header_row=1,
    data_start_row=2,
    output_mode="sheets",
):
    dialogs = SimpleNamespace(
        ask_text=Mock(side_effect=[sheet_name, column]),
        ask_positive_int=Mock(side_effect=[header_row, data_start_row]),
        ask_choice=Mock(return_value=output_mode),
        show_info=Mock(),
    )
    logger = Mock()
    button = Mock()
    ctx = SimpleNamespace(
        logger=logger,
        root=None,
        dialogs=dialogs,
        button=Mock(return_value=button),
    )
    return ExportSplitController(ctx), dialogs, button, logger


def test_split_by_column_sheets_mode(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode="sheets")
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_sheet_count": 3,
            "copied_row_count": 15,
        }
    )
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column_to_files", mock_to_files)

    controller.run_split_sheet()

    dialogs.ask_choice.assert_called_once_with(
        "按列拆分",
        "请选择输出方式：",
        [
            ("拆分为工作表", "sheets"),
            ("拆分为多个文件", "files"),
        ],
    )
    mock_split.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        column_input="B",
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_to_files.assert_not_called()
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_column_files_mode(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode="files")
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value="D:/output_folder"),
    )
    mock_split = Mock()
    mock_to_files = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_file_count": 4,
            "copied_row_count": 20,
            "output_dir": "D:/output_folder",
            "output_paths": ["D:/output_folder/a.xlsx"],
        }
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column_to_files", mock_to_files)

    controller.run_split_sheet()

    mock_to_files.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        column_input="B",
        output_dir="D:/output_folder",
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_split.assert_not_called()
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_column_mode_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode=None)
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column_to_files", mock_to_files)

    controller.run_split_sheet()

    dialogs.ask_choice.assert_called_once()
    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_files_mode_directory_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode="files")
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value=""),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column_to_files", mock_to_files)

    controller.run_split_sheet()

    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_sheets_mode_exception(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode="sheets")
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(side_effect=RuntimeError("COM 打开失败"))
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column", mock_split)

    controller.run_split_sheet()

    logger.error.assert_called_with("按列拆分失败：COM 打开失败")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_column_files_mode_exception(monkeypatch):
    controller, dialogs, button, logger = _controller(output_mode="files")
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsm"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value="D:/output_folder"),
    )
    mock_to_files = Mock(
        side_effect=ValueError(
            "按列拆分为多个文件暂不支持 .xlsm：Worksheet.Copy 创建的新工作簿无法可靠保留 VBA 工程。"
        )
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_column_to_files", mock_to_files)

    controller.run_split_sheet()

    logger.error.assert_called_with(
        "按列拆分失败：按列拆分为多个文件暂不支持 .xlsm：Worksheet.Copy 创建的新工作簿无法可靠保留 VBA 工程。"
    )
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_column_source_file_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value=""),
    )

    controller.run_split_sheet()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_sheet_name_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = [None]

    controller.run_split_sheet()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_column_input_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = ["Sheet1", None]

    controller.run_split_sheet()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_header_row_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = ["Sheet1", "A"]
    dialogs.ask_positive_int.side_effect = [None]

    controller.run_split_sheet()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_data_start_row_cancelled(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = ["Sheet1", "A"]
    dialogs.ask_positive_int.side_effect = [1, None]

    controller.run_split_sheet()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_column_invalid_row_numbers(monkeypatch):
    controller, dialogs, button, logger = _controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = ["Sheet1", "A"]
    dialogs.ask_positive_int.side_effect = [5, 3]  # data_start_row <= header_row

    controller.run_split_sheet()

    logger.error.assert_called_with("输入无效：数据起始行号必须大于表头行号。")
    button.config.assert_not_called()


def _rows_controller(
    *,
    sheet_name="Sheet1",
    header_row=1,
    data_start_row=2,
    split_mode="each_row",
    rows_per_part=100,
    output_mode="sheets",
    choices=None,
    positive_ints=None,
):
    if choices is None:
        choices = [split_mode, output_mode]
    if positive_ints is None:
        if split_mode == "n_rows":
            positive_ints = [header_row, data_start_row, rows_per_part]
        else:
            positive_ints = [header_row, data_start_row]

    dialogs = SimpleNamespace(
        ask_text=Mock(side_effect=[sheet_name]),
        ask_positive_int=Mock(side_effect=positive_ints),
        ask_choice=Mock(side_effect=choices),
        show_info=Mock(),
    )
    logger = Mock()
    button = Mock()
    button_getter = Mock(return_value=button)
    ctx = SimpleNamespace(
        logger=logger,
        root=None,
        dialogs=dialogs,
        button=button_getter,
    )
    return ExportSplitController(ctx), dialogs, button, logger, button_getter


def test_split_by_rows_each_row_sheets_mode(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="each_row",
        output_mode="sheets",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_sheet_count": 5,
            "copied_row_count": 5,
            "rows_per_part": 1,
        }
    )
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_choice.call_count == 2
    assert dialogs.ask_positive_int.call_count == 2
    mock_split.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_to_files.assert_not_called()
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_n_rows_sheets_mode(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="n_rows",
        rows_per_part=50,
        output_mode="sheets",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_sheet_count": 3,
            "copied_row_count": 120,
            "rows_per_part": 50,
        }
    )
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_positive_int.call_count == 3
    mock_split.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        rows_per_part=50,
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_to_files.assert_not_called()
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_each_row_files_mode(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="each_row",
        output_mode="files",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value="D:/output_folder"),
    )
    mock_split = Mock()
    mock_to_files = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_file_count": 5,
            "copied_row_count": 5,
            "rows_per_part": 1,
            "output_dir": "D:/output_folder",
            "output_paths": ["D:/output_folder/a.xlsx"],
        }
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_positive_int.call_count == 2
    mock_to_files.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        output_dir="D:/output_folder",
        rows_per_part=1,
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_split.assert_not_called()
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_n_rows_files_mode(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="n_rows",
        rows_per_part=200,
        output_mode="files",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value="D:/output_folder"),
    )
    mock_split = Mock()
    mock_to_files = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_file_count": 2,
            "copied_row_count": 350,
            "rows_per_part": 200,
            "output_dir": "D:/output_folder",
            "output_paths": ["D:/output_folder/1.xlsx", "D:/output_folder/2.xlsx"],
        }
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_positive_int.call_count == 3
    mock_to_files.assert_called_once_with(
        source_path="D:/data/source.xlsx",
        source_sheet_name="Sheet1",
        output_dir="D:/output_folder",
        rows_per_part=200,
        header_row=1,
        data_start_row=2,
        logger=logger,
    )
    mock_split.assert_not_called()
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_split_mode_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        choices=[None],
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_choice.call_count == 1
    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_n_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        choices=["n_rows"],
        positive_ints=[1, 2, None],
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_output_mode_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        choices=["each_row", None],
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    assert dialogs.ask_choice.call_count == 2
    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_output_dir_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        split_mode="each_row",
        output_mode="files",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value=""),
    )
    mock_split = Mock()
    mock_to_files = Mock()
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    mock_split.assert_not_called()
    mock_to_files.assert_not_called()
    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_sheets_mode_exception(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="each_row",
        output_mode="sheets",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(side_effect=RuntimeError("COM 打开失败"))
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)

    controller.run_split_rows()

    logger.error.assert_called_with("按行拆分失败：COM 打开失败")
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_files_mode_exception(monkeypatch):
    controller, dialogs, button, logger, button_getter = _rows_controller(
        split_mode="each_row",
        output_mode="files",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsm"),
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askdirectory",
        Mock(return_value="D:/output_folder"),
    )
    mock_to_files = Mock(
        side_effect=ValueError(
            "按行拆分为多个文件暂不支持 .xlsm：Worksheet.Copy 创建的新工作簿无法可靠保留 VBA 工程。"
        )
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows_to_files", mock_to_files)

    controller.run_split_rows()

    logger.error.assert_called_with(
        "按行拆分失败：按行拆分为多个文件暂不支持 .xlsm：Worksheet.Copy 创建的新工作簿无法可靠保留 VBA 工程。"
    )
    button_getter.assert_called_with("btn_split_rows")
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_split_by_rows_invalid_row_numbers(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        positive_ints=[5, 3],
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )

    controller.run_split_rows()

    logger.error.assert_called_with("输入无效：数据起始行号必须大于表头行号。")
    button.config.assert_not_called()


def test_split_by_rows_each_row_does_not_ask_for_n(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        split_mode="each_row",
        output_mode="sheets",
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    mock_split = Mock(
        return_value={
            "workbook_name": "source.xlsx",
            "source_sheet_name": "Sheet1",
            "created_sheet_count": 1,
            "copied_row_count": 1,
            "rows_per_part": 1,
        }
    )
    monkeypatch.setattr(export_split_ui, "split_workbook_sheet_by_rows", mock_split)

    controller.run_split_rows()

    prompts = [call.args[1] for call in dialogs.ask_positive_int.call_args_list]
    assert "每份数据行数：" not in prompts
    assert dialogs.ask_positive_int.call_count == 2


def test_split_by_rows_source_file_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value=""),
    )

    controller.run_split_rows()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_sheet_name_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller()
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )
    dialogs.ask_text.side_effect = [None]

    controller.run_split_rows()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_header_row_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(positive_ints=[None])
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )

    controller.run_split_rows()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_data_start_row_cancelled(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(positive_ints=[1, None])
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )

    controller.run_split_rows()

    logger.info.assert_called_with("用户已取消操作")
    button.config.assert_not_called()


def test_split_by_rows_invalid_n(monkeypatch):
    controller, dialogs, button, logger, _ = _rows_controller(
        choices=["n_rows"],
        positive_ints=[1, 2, 0],
    )
    monkeypatch.setattr(
        export_split_ui.filedialog,
        "askopenfilename",
        Mock(return_value="D:/data/source.xlsx"),
    )

    controller.run_split_rows()

    logger.error.assert_called_with("输入无效：每份行数必须大于等于 1。")
    button.config.assert_not_called()
