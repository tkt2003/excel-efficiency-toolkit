# -*- coding: utf-8 -*-
"""功能控制器基类。"""


class FeatureController:
    """通过 UIContext 访问窗口、日志、对话框服务和功能卡片句柄。"""

    def __init__(self, ctx):
        self.ctx = ctx

    @property
    def root(self):
        return self.ctx.root

    @property
    def logger(self):
        return self.ctx.logger

    @property
    def dialogs(self):
        return self.ctx.dialogs

    def button(self, name):
        """按注册名取功能卡片句柄；名字错误立即 KeyError。"""
        return self.ctx.button(name)

    # ------------------------------------------------------------------
    # 日志快捷方式
    # ------------------------------------------------------------------
    def log_info(self, message):
        self.ctx.log_info(message)

    def log_error(self, message):
        self.ctx.log_error(message)

    def flushing_logger(self):
        return self.ctx.flushing_logger()

    def flush_ui(self):
        self.ctx.flush_ui()

    # ------------------------------------------------------------------
    # 对话框快捷方式
    # ------------------------------------------------------------------
    def ask_text(self, *args, **kwargs):
        return self.ctx.dialogs.ask_text(*args, **kwargs)

    def ask_choice(self, *args, **kwargs):
        return self.ctx.dialogs.ask_choice(*args, **kwargs)

    def show_info(self, *args, **kwargs):
        return self.ctx.dialogs.show_info(*args, **kwargs)

    def ask_positive_int(self, *args, **kwargs):
        return self.ctx.dialogs.ask_positive_int(*args, **kwargs)

    def ask_column(self, *args, **kwargs):
        return self.ctx.dialogs.ask_column(*args, **kwargs)

    def ask_split_preview(self, *args, **kwargs):
        return self.ctx.dialogs.ask_split_preview(*args, **kwargs)

    def create_progress_cancel_dialog(self, *args, **kwargs):
        return self.ctx.dialogs.create_progress_cancel_dialog(*args, **kwargs)
