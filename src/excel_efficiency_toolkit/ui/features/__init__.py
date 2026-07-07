# -*- coding: utf-8 -*-
"""功能控制器集合。

只允许静态 import（禁止 pkgutil/importlib 扫描），
保证 PyInstaller 依赖分析能收录全部控制器模块。
"""
from .color_tools import ColorToolsController
from .data_drill import DataDrillController
from .delete_sheets import DeleteSheetsController
from .export_split import ExportSplitController
from .link_replace import LinkReplaceController
from .merge import MergeController
from .rename_files import RenameFilesController
from .rename_sheets import RenameSheetsController
from .templates import TemplatesController
from .word_replace import WordReplaceController
from .workbook_misc import WorkbookMiscController


class Controllers:
    """按功能域组织的控制器集合。"""

    def __init__(self, ctx):
        self.export_split = ExportSplitController(ctx)
        self.merge = MergeController(ctx)
        self.data_drill = DataDrillController(ctx)
        self.templates = TemplatesController(ctx)
        self.color_tools = ColorToolsController(ctx)
        self.workbook_misc = WorkbookMiscController(ctx)
        self.rename_files = RenameFilesController(ctx)
        self.rename_sheets = RenameSheetsController(ctx)
        self.link_replace = LinkReplaceController(ctx)
        self.word_replace = WordReplaceController(ctx)
        self.delete_sheets = DeleteSheetsController(ctx)


def build_controllers(ctx):
    return Controllers(ctx)
