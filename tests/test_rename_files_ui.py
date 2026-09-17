from types import SimpleNamespace
from unittest.mock import Mock

from src.excel_efficiency_toolkit.ui.features import rename_files as rename_files_ui
from src.excel_efficiency_toolkit.ui.features.rename_files import RenameFilesController


def _controller(selection):
    dialogs = SimpleNamespace(
        ask_choice=Mock(return_value=selection),
        show_info=Mock(),
    )
    button = Mock()
    ctx = SimpleNamespace(
        logger=Mock(),
        root=None,
        dialogs=dialogs,
        button=Mock(return_value=button),
    )
    return RenameFilesController(ctx), dialogs, button


def test_batch_rename_file_entry_uses_selected_files(monkeypatch, tmp_path):
    controller, dialogs, button = _controller("files")
    selected = (str(tmp_path / "甲.xlsx"), str(tmp_path / "乙.docx"))
    create_rules = Mock(return_value=str(tmp_path / "rename_rules.xlsx"))
    monkeypatch.setattr(rename_files_ui.filedialog, "askopenfilenames", Mock(return_value=selected))
    monkeypatch.setattr(rename_files_ui, "create_rename_rule_workbook", create_rules)
    monkeypatch.setattr(rename_files_ui.os, "startfile", Mock())
    monkeypatch.setattr(controller, "_confirm_rename_rule_ready", Mock(return_value=False))

    controller.run_batch_rename_files()

    dialogs.ask_choice.assert_called_once()
    assert dialogs.ask_choice.call_args.args[2] == [
        ("选择文件", "files"),
        ("选择文件夹", "folder"),
    ]
    create_rules.assert_called_once_with(list(selected))
    button.config.assert_any_call(state="disabled")
    button.config.assert_any_call(state="normal")


def test_batch_rename_folder_entry_includes_files_from_subfolders(monkeypatch, tmp_path):
    controller, dialogs, _ = _controller("folder")
    first = tmp_path / "甲.xlsx"
    second = tmp_path / "乙.docx"
    first.write_text("a", encoding="utf-8")
    second.write_text("b", encoding="utf-8")
    nested = tmp_path / "子目录"
    nested.mkdir()
    nested_file = nested / "应递归选择.xlsx"
    nested_file.write_text("c", encoding="utf-8")
    (tmp_path / "批量重命名规则_20260713_131047.xlsx").write_text("control", encoding="utf-8")
    (tmp_path / "~$批量重命名规则_20260713_131047.xlsx").write_text("lock", encoding="utf-8")
    create_rules = Mock(return_value=str(tmp_path / "rename_rules.xlsx"))
    monkeypatch.setattr(
        rename_files_ui.filedialog,
        "askopenfilenames",
        Mock(side_effect=AssertionError("文件夹入口不应打开文件选择器")),
    )
    monkeypatch.setattr(rename_files_ui.filedialog, "askdirectory", Mock(return_value=str(tmp_path)))
    monkeypatch.setattr(rename_files_ui, "create_rename_rule_workbook", create_rules)
    monkeypatch.setattr(rename_files_ui.os, "startfile", Mock())
    monkeypatch.setattr(controller, "_confirm_rename_rule_ready", Mock(return_value=False))

    controller.run_batch_rename_files()

    dialogs.ask_choice.assert_called_once()
    create_rules.assert_called_once()
    assert set(create_rules.call_args.args[0]) == {str(first), str(second), str(nested_file)}


def test_batch_rename_source_dialog_can_be_cancelled(monkeypatch):
    controller, dialogs, button = _controller(None)
    monkeypatch.setattr(
        rename_files_ui.filedialog,
        "askopenfilenames",
        Mock(side_effect=AssertionError("取消后不应打开文件选择器")),
    )
    monkeypatch.setattr(
        rename_files_ui.filedialog,
        "askdirectory",
        Mock(side_effect=AssertionError("取消后不应打开文件夹选择器")),
    )

    controller.run_batch_rename_files()

    dialogs.ask_choice.assert_called_once()
    button.config.assert_not_called()
