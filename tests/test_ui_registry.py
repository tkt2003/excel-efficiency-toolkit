from types import SimpleNamespace

from src.excel_efficiency_toolkit import __version__
from src.excel_efficiency_toolkit.ui.features.registry import build_feature_groups
from src.excel_efficiency_toolkit.ui.main_window import APP_VERSION
from src.excel_efficiency_toolkit.ui import theme


def _controller(**callbacks):
    return SimpleNamespace(**{name: (lambda: None) for name in callbacks})


def _controllers():
    return SimpleNamespace(
        merge=_controller(run_merge_sheets=None, run_merge_workbooks=None),
        export_split=_controller(
            run_export_sheets=None,
            run_split_sheet=None,
            run_split_rows=None,
        ),
        data_drill=_controller(run_data_drill=None),
        templates=_controller(run_template_generate=None),
        color_tools=_controller(run_color_sum=None, run_clear_by_color=None),
        workbook_misc=_controller(run_round_formula=None, run_sheet_index=None),
        rename_files=_controller(run_batch_rename_files=None),
        rename_sheets=_controller(run_batch_rename_sheets=None),
        link_replace=_controller(run_batch_link_replace=None),
        word_replace=_controller(run_word_batch_replace=None),
        delete_sheets=_controller(run_delete_sheets=None),
    )


def test_ui_version_uses_package_version():
    assert __version__ == "1.1.0"
    assert APP_VERSION == __version__


def test_log_area_uses_the_same_ui_font_family_as_the_application():
    assert theme.LOG_FONT[0] == theme.FONT_FAMILY


def test_feature_registry_has_complete_unique_navigation_metadata():
    groups = build_feature_groups(_controllers())

    assert len(groups) == 5
    assert [group.key for group in groups] == [
        "merge",
        "export",
        "maintenance",
        "templates",
        "workbook",
    ]
    assert len({group.key for group in groups}) == len(groups)
    assert all(group.title and group.icon and group.summary for group in groups)

    features = [feature for group in groups for feature in group.features]
    assert len(features) == 16
    assert len({feature.attr_name for feature in features}) == len(features)
    assert all(feature.title and feature.description for feature in features)
    assert all(callable(feature.command) for feature in features)


def test_export_split_sheet_feature_registration():
    groups = build_feature_groups(_controllers())
    export_group = next(group for group in groups if group.key == "export")
    split_feature = next(
        feature for feature in export_group.features if feature.attr_name == "btn_split_sheet"
    )
    assert split_feature.title == "按列拆分"
    assert split_feature.description == "按指定列内容拆分为工作表或多个 Excel 文件"


def test_export_split_rows_feature_registration():
    groups = build_feature_groups(_controllers())
    export_group = next(group for group in groups if group.key == "export")
    split_feature = next(
        feature for feature in export_group.features if feature.attr_name == "btn_split_rows"
    )
    assert split_feature.title == "按行拆分"
    assert split_feature.description == "按数据行拆分为工作表或多个 Excel 文件"
