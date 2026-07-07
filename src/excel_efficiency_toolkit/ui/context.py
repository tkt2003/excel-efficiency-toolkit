# -*- coding: utf-8 -*-
"""UIContext：功能控制器共享的 GUI 服务上下文。"""
import tkinter as tk

from ..logging_utils import FlushingLogger
from .dialogs import DialogService


class UIContext:
    """聚合主窗口、日志、对话框服务和功能卡片句柄注册表。

    logger 在主窗口日志区创建完成后由组合根注入。
    """

    def __init__(self, root):
        self.root = root
        self.dialogs = DialogService(root)
        self.logger = None
        self.buttons = {}
        self.feature_cards = {}

    def button(self, name):
        """按注册名取功能卡片句柄；名字错误立即 KeyError。"""
        return self.buttons[name]

    def flush_ui(self):
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            pass

    def log_info(self, message):
        self.logger.info(message)
        self.flush_ui()

    def log_error(self, message):
        self.logger.error(message)
        self.flush_ui()

    def flushing_logger(self):
        return FlushingLogger(self.logger, self.flush_ui)
