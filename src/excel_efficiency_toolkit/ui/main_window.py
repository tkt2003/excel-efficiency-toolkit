# -*- coding: utf-8 -*-
"""主窗口：装配头部、功能区、日志区，并提供 GUI 冒烟入口。"""
import os
import sys
import tkinter as tk

import customtkinter as ctk

from ..logging_utils import setup_logger
from . import theme
from .context import UIContext
from .features import build_controllers
from .features.registry import build_feature_groups
from .widgets import AppButton, create_section_panel

APP_NAME = "老头表格助手"
APP_VERSION = "1.0.0"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"


class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(theme.WINDOW_GEOMETRY)
        self.root.minsize(*theme.WINDOW_MIN_SIZE)
        self.root.configure(fg_color=theme.BG)

        self.ctx = UIContext(root)
        self.controllers = build_controllers(self.ctx)

        self.main_frame = ctk.CTkFrame(root, fg_color=theme.BG, corner_radius=0)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self._create_header(self.main_frame)
        self._create_feature_area(self.main_frame)
        self._create_log_area(self.main_frame)

        # 初始化自定义 logger（日志区创建后注入 ctx）
        self.logger = setup_logger(self.log_text)
        self.ctx.logger = self.logger
        self.logger.info(f"欢迎使用 {APP_TITLE}。程序已就绪。")
        if os.environ.get("EXCEL_TOOLKIT_GUI_SMOKE") == "1":
            self.root.after(300, self._run_gui_smoke_and_exit)

    # ------------------------------------------------------------------
    # 布局
    # ------------------------------------------------------------------
    def _create_header(self, parent):
        header = ctk.CTkFrame(
            parent,
            fg_color=theme.SURFACE,
            corner_radius=12,
            border_width=1,
            border_color=theme.BORDER_SOFT,
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)

        title_block = ctk.CTkFrame(header, fg_color="transparent")
        title_block.grid(row=0, column=0, sticky="ew", padx=18, pady=10)
        title_block.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_block,
            text=APP_TITLE,
            font=theme.HEADER_TITLE_FONT,
            text_color=theme.TEXT,
        ).grid(row=0, column=0, sticky="w")

        help_buttons = ctk.CTkFrame(title_block, fg_color="transparent")
        help_buttons.grid(row=0, column=1, rowspan=2, sticky="e")
        for index, (text, command) in enumerate((("使用说明", self.show_user_guide), ("关于", self.show_about))):
            AppButton(
                help_buttons,
                text=text,
                font=theme.HEADER_BUTTON_FONT,
                command=command,
                width=76,
                height=28,
                fg_color=theme.BUTTON_SECONDARY_BG,
                hover_color=theme.HEADER_BUTTON_HOVER,
                text_color=theme.TEXT,
                border_width=1,
                border_color=theme.BORDER_BUTTON,
                corner_radius=6,
            ).grid(row=0, column=index, padx=(8 if index else 0, 0))

        ctk.CTkLabel(
            title_block,
            text="Excel 效率工作台",
            font=theme.HEADER_SUBTITLE_FONT,
            text_color=theme.TEXT_SECONDARY,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

    def _create_feature_area(self, parent):
        feature_area = ctk.CTkScrollableFrame(
            parent,
            fg_color="transparent",
            scrollbar_button_color=theme.SCROLLBAR,
            scrollbar_button_hover_color=theme.SCROLLBAR_HOVER,
        )
        feature_area.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        feature_area.grid_columnconfigure(0, weight=1, uniform="feature_columns")
        feature_area.grid_columnconfigure(1, weight=1, uniform="feature_columns")

        groups = dict(build_feature_groups(self.controllers))
        left_sections = ("合并整理", "批量维护")
        right_sections = ("拆分导出", "模板生成 / 取数", "工作簿辅助")

        def register_card(attr_name, card, handle):
            self.ctx.buttons[attr_name] = handle
            self.ctx.feature_cards[attr_name] = card

        for column, section_names in enumerate((left_sections, right_sections)):
            column_frame = ctk.CTkFrame(feature_area, fg_color="transparent")
            column_frame.grid(row=0, column=column, sticky="new", padx=(0, 12) if column == 0 else (0, 0))
            column_frame.grid_columnconfigure(0, weight=1)

            for row, section_name in enumerate(section_names):
                panel = create_section_panel(column_frame, section_name, groups[section_name], register_card)
                panel.grid(row=row, column=0, sticky="ew", pady=(0, 8))

        feature_area.grid_rowconfigure(1, minsize=14)

    def _create_log_area(self, parent):
        self.frame_bottom = ctk.CTkFrame(
            parent,
            fg_color=theme.SURFACE,
            corner_radius=12,
            border_width=1,
            border_color=theme.BORDER_SOFT,
            height=theme.LOG_AREA_HEIGHT,
        )
        self.frame_bottom.grid(row=2, column=0, sticky="nsew")
        self.frame_bottom.grid_propagate(False)
        self.frame_bottom.grid_columnconfigure(0, weight=1)
        self.frame_bottom.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self.frame_bottom,
            text="运行日志",
            font=theme.LOG_TITLE_FONT,
            text_color=theme.TEXT,
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(8, 4))

        self.log_text = ctk.CTkTextbox(
            self.frame_bottom,
            state="disabled",
            height=theme.LOG_TEXT_HEIGHT,
            font=theme.LOG_FONT,
            fg_color=theme.SURFACE_SUNKEN,
            text_color=theme.TEXT_STRONG,
            border_width=1,
            border_color=theme.BORDER,
            corner_radius=8,
            wrap="word",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

    # ------------------------------------------------------------------
    # 帮助入口
    # ------------------------------------------------------------------
    def _get_project_root(self):
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))

    def _find_user_guide_path(self):
        candidates = []
        if getattr(sys, "frozen", False):
            candidates.append(os.path.dirname(sys.executable))
            candidates.append(self._get_project_root())
        else:
            candidates.append(self._get_project_root())
            candidates.append(os.getcwd())

        seen = set()
        for base_dir in candidates:
            guide_path = os.path.abspath(os.path.join(base_dir, "使用说明.txt"))
            if guide_path in seen:
                continue
            seen.add(guide_path)
            if os.path.isfile(guide_path):
                return guide_path
        return None

    def show_user_guide(self):
        guide_path = self._find_user_guide_path()
        if guide_path:
            try:
                os.startfile(guide_path)
                self.logger.info(f"已打开使用说明：{guide_path}")
                return
            except OSError as exc:
                self.logger.error(f"打开使用说明失败：{exc}")

        self.ctx.dialogs.show_info(
            "使用说明",
            (
                "老头表格助手用于处理常见 Excel 批量整理工作。\n"
                "批量修改类功能请先备份或先用副本试跑。\n"
                "详细说明请查看同目录下的 使用说明.txt。"
            ),
            dialog_width=460,
            dialog_height=280,
            wraplength=420,
        )

    def show_about(self):
        self.ctx.dialogs.show_info(
            "关于",
            (
                f"软件名：{APP_NAME}\n"
                f"版本号：v{APP_VERSION}\n"
                "说明：面向审计、财务、报表整理场景的 Excel 效率工具\n"
                "当前阶段：常用功能阶段性完成"
            ),
            dialog_width=460,
            dialog_height=300,
            wraplength=380,
        )

    # ------------------------------------------------------------------
    # GUI 冒烟
    # ------------------------------------------------------------------
    def _run_gui_smoke_and_exit(self):
        try:
            self.ctx.log_info("GUI 冒烟：开始检查关键入口。")
            for card in self.ctx.feature_cards.values():
                if not getattr(card, "_feature_enabled", True):
                    raise RuntimeError("GUI 冒烟：存在不可用功能卡片。")

            # 遍历切换全部卡片句柄，抓 stale attr_name 和状态回调异常
            for handle in self.ctx.buttons.values():
                handle.config(state="disabled")
                handle.config(state="normal")

            for dialog_callback in (
                self.show_about,
                self.controllers.data_drill.run_data_drill,
                self.controllers.templates.run_template_generate,
                self.controllers.color_tools.run_clear_by_color,
            ):
                dialog_callback()
                self.ctx.flush_ui()
            self.ctx.log_info("GUI 冒烟：完成。")
        finally:
            self.root.after(200, self.root.destroy)
