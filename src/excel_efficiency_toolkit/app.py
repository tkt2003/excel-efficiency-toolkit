# -*- coding: utf-8 -*-
"""组合根：保持 run_app.py 与 PyInstaller 入口的 `app.main` 导入路径不变。

窗口与布局在 ui/main_window.py，功能交互在 ui/features/，
业务逻辑在各 *_ops.py 模块。
"""
import customtkinter as ctk

from .ui import theme
from .ui.main_window import APP_NAME, APP_TITLE, APP_VERSION, ExcelToolkitApp

theme.apply_appearance()


def main():
    root = ctk.CTk()
    ExcelToolkitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
