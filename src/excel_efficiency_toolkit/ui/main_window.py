# -*- coding: utf-8 -*-
"""v1.1 主窗口：场景化侧栏、独立功能页面与可折叠运行记录。"""
import os
import sys
import tkinter as tk

import customtkinter as ctk

from .. import __version__
from ..logging_utils import setup_logger
from . import theme
from .context import UIContext
from .features import build_controllers
from .features.registry import build_feature_groups
from .widgets import AppButton, create_category_card, create_feature_card

APP_NAME = "老头表格助手"
APP_VERSION = __version__
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
        self.feature_groups = build_feature_groups(self.controllers)
        self.pages = {}
        self.nav_buttons = {}
        self.current_page = None
        self.log_collapsed = False

        self.shell = ctk.CTkFrame(root, fg_color=theme.BG, corner_radius=0)
        self.shell.pack(fill=tk.BOTH, expand=True)
        self.shell.grid_columnconfigure(0, minsize=theme.SIDEBAR_WIDTH)
        self.shell.grid_columnconfigure(1, weight=1)
        self.shell.grid_rowconfigure(0, weight=1)

        self._create_sidebar(self.shell)
        self._create_workspace(self.shell)

        self.logger = setup_logger(self.log_text)
        self.ctx.logger = self.logger
        self.show_page("home")
        self.logger.info(f"欢迎使用 {APP_TITLE}。工作台已就绪。")
        if os.environ.get("EXCEL_TOOLKIT_GUI_SMOKE") == "1":
            self.root.after(300, self._run_gui_smoke_and_exit)

    def _create_sidebar(self, parent):
        sidebar = ctk.CTkFrame(
            parent, width=theme.SIDEBAR_WIDTH, fg_color=theme.SIDEBAR, corner_radius=0
        )
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_rowconfigure(3, weight=1)

        brand = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="ew", padx=20, pady=(24, 22))
        brand.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            brand, text="表", width=42, height=42, corner_radius=12,
            fg_color=theme.ACCENT, text_color=theme.TEXT_ON_ACCENT,
            font=theme.BRAND_MARK_FONT,
        ).grid(row=0, column=0, rowspan=2, padx=(0, 11))
        ctk.CTkLabel(
            brand, text=APP_NAME, font=theme.BRAND_FONT,
            text_color=theme.SIDEBAR_TEXT, anchor="w",
        ).grid(row=0, column=1, sticky="sw")
        ctk.CTkLabel(
            brand, text=f"效率工作台  ·  v{APP_VERSION}",
            font=theme.HEADER_SUBTITLE_FONT, text_color=theme.SIDEBAR_TEXT_MUTED, anchor="w",
        ).grid(row=1, column=1, sticky="nw", pady=(2, 0))

        ctk.CTkLabel(
            sidebar, text="工作空间", font=theme.SECTION_META_FONT,
            text_color=theme.SIDEBAR_TEXT_MUTED, anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 8))

        nav = ctk.CTkFrame(sidebar, fg_color="transparent")
        nav.grid(row=2, column=0, sticky="new", padx=12)
        nav.grid_columnconfigure(0, weight=1)
        self._add_nav_button(nav, 0, "home", "⌂", "总览")
        for row, group in enumerate(self.feature_groups, start=1):
            self._add_nav_button(nav, row, group.key, group.icon, group.title)

        privacy = ctk.CTkFrame(sidebar, fg_color=theme.SIDEBAR_RAISED, corner_radius=12)
        privacy.grid(row=4, column=0, sticky="sew", padx=16, pady=(16, 18))
        privacy.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            privacy, text="●", font=theme.SMALL_FONT, text_color=theme.SUCCESS,
        ).grid(row=0, column=0, padx=(12, 8), pady=(11, 0))
        ctk.CTkLabel(
            privacy, text="本地安全处理", font=theme.LABEL_BOLD_FONT,
            text_color=theme.SIDEBAR_TEXT, anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(11, 0))
        ctk.CTkLabel(
            privacy, text="文件仅在当前电脑上处理\n批量操作前请先备份",
            font=theme.SMALL_FONT, text_color=theme.SIDEBAR_TEXT_MUTED,
            justify="left", anchor="w",
        ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(5, 12))

    def _add_nav_button(self, parent, row, key, icon, title):
        button = AppButton(
            parent, text=f"  {icon}    {title}",
            command=lambda page_key=key: self.show_page(page_key),
            height=46, anchor="w", font=theme.NAV_FONT, fg_color="transparent",
            hover_color=theme.SIDEBAR_HOVER, text_color=theme.SIDEBAR_TEXT_MUTED,
            corner_radius=10,
        )
        button.grid(row=row, column=0, sticky="ew", pady=2)
        self.nav_buttons[key] = button

    def _create_workspace(self, parent):
        workspace = ctk.CTkFrame(parent, fg_color=theme.BG, corner_radius=0)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.grid_columnconfigure(0, weight=1)
        workspace.grid_rowconfigure(1, weight=1)
        self._create_header(workspace)
        self.page_host = ctk.CTkFrame(workspace, fg_color="transparent")
        self.page_host.grid(row=1, column=0, sticky="nsew", padx=22)
        self.page_host.grid_columnconfigure(0, weight=1)
        self.page_host.grid_rowconfigure(0, weight=1)
        self._create_pages(self.page_host)
        self._create_log_area(workspace)

    def _create_header(self, parent):
        header = ctk.CTkFrame(
            parent, height=theme.HEADER_HEIGHT, fg_color=theme.SURFACE,
            corner_radius=0, border_width=0,
        )
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        title_block = ctk.CTkFrame(header, fg_color="transparent")
        title_block.grid(row=0, column=0, sticky="w", padx=24, pady=16)
        self.page_title_label = ctk.CTkLabel(
            title_block, text="工作台总览", font=theme.HEADER_TITLE_FONT,
            text_color=theme.TEXT, anchor="w",
        )
        self.page_title_label.grid(row=0, column=0, sticky="w")
        self.page_subtitle_label = ctk.CTkLabel(
            title_block, text="选择一个工作场景，开始处理表格任务",
            font=theme.HEADER_SUBTITLE_FONT, text_color=theme.TEXT_SECONDARY, anchor="w",
        )
        self.page_subtitle_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        actions = ctk.CTkFrame(header, fg_color="transparent")
        actions.grid(row=0, column=1, sticky="e", padx=22)
        for index, (text, command) in enumerate((("使用说明", self.show_user_guide), ("关于", self.show_about))):
            AppButton(
                actions, text=text, font=theme.HEADER_BUTTON_FONT, command=command,
                width=84, height=34, fg_color=theme.BUTTON_SECONDARY_BG,
                hover_color=theme.HEADER_BUTTON_HOVER, text_color=theme.TEXT,
                border_width=1, border_color=theme.BORDER_BUTTON, corner_radius=9,
            ).grid(row=0, column=index, padx=(8 if index else 0, 0))

    def _create_pages(self, parent):
        self.pages["home"] = self._create_home_page(parent)
        for group in self.feature_groups:
            self.pages[group.key] = self._create_group_page(parent, group)
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew", pady=(20, 12))
            page.grid_remove()

    def _create_scroll_page(self, parent):
        page = ctk.CTkScrollableFrame(
            parent, fg_color="transparent", scrollbar_button_color=theme.SCROLLBAR,
            scrollbar_button_hover_color=theme.SCROLLBAR_HOVER,
        )
        page.grid_columnconfigure(0, weight=1, uniform="page_columns")
        page.grid_columnconfigure(1, weight=1, uniform="page_columns")
        return page

    def _create_home_page(self, parent):
        page = self._create_scroll_page(parent)
        hero = ctk.CTkFrame(page, fg_color=theme.ACCENT_DARK, corner_radius=18)
        hero.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        hero.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            hero, text="EXCEL PRODUCTIVITY SUITE", font=theme.HERO_KICKER_FONT,
            text_color=theme.SUCCESS, anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=26, pady=(22, 4))
        ctk.CTkLabel(
            hero, text="把重复的表格工作，交给工具。", font=theme.HERO_TITLE_FONT,
            text_color=theme.TEXT_ON_ACCENT, anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=26)
        ctk.CTkLabel(
            hero, text="从合并整理到模板生成，所有任务都按真实工作场景重新归类。",
            font=theme.HERO_BODY_FONT, text_color=theme.SIDEBAR_TEXT_MUTED, anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=26, pady=(8, 24))
        feature_count = sum(len(group.features) for group in self.feature_groups)
        ctk.CTkLabel(
            hero, text=f"{feature_count}\n项工具", font=theme.CARD_TITLE_FONT,
            text_color=theme.TEXT_ON_ACCENT, justify="center",
        ).grid(row=0, column=1, rowspan=3, padx=(18, 30))
        ctk.CTkLabel(
            page, text="选择工作场景", font=theme.SECTION_TITLE_FONT,
            text_color=theme.TEXT, anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(0, 4))
        ctk.CTkLabel(
            page, text="5 类场景 · 所有功能均在本地运行", font=theme.SECTION_META_FONT,
            text_color=theme.TEXT_MUTED, anchor="e",
        ).grid(row=1, column=1, sticky="e", pady=(0, 4))
        for index, group in enumerate(self.feature_groups):
            row = 2 + index // 2
            column = index % 2
            card = create_category_card(
                page, group.icon, group.title, group.summary, len(group.features),
                lambda page_key=group.key: self.show_page(page_key),
            )
            card.grid(
                row=row, column=column, sticky="ew",
                padx=(0, 8) if column == 0 else (8, 0), pady=(10, 6),
            )
        return page

    def _create_group_page(self, parent, group):
        page = self._create_scroll_page(parent)
        intro = ctk.CTkFrame(
            page, fg_color=theme.ACCENT_PALE, border_width=1,
            border_color=theme.BORDER_SOFT, corner_radius=16,
        )
        intro.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        intro.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            intro, text=group.icon, width=48, height=48, corner_radius=13,
            fg_color=theme.ACCENT, text_color=theme.TEXT_ON_ACCENT,
            font=theme.CARD_ICON_FONT,
        ).grid(row=0, column=0, rowspan=2, padx=(18, 14), pady=18)
        ctk.CTkLabel(
            intro, text=group.title, font=theme.SECTION_TITLE_FONT,
            text_color=theme.TEXT, anchor="w",
        ).grid(row=0, column=1, sticky="sw", pady=(18, 0))
        ctk.CTkLabel(
            intro, text=group.summary, font=theme.CARD_DESCRIPTION_FONT,
            text_color=theme.TEXT_SECONDARY, anchor="w",
        ).grid(row=1, column=1, sticky="nw", pady=(5, 18))
        ctk.CTkLabel(
            intro, text=f"{len(group.features)} 项工具", font=theme.SECTION_META_FONT,
            text_color=theme.ACCENT,
        ).grid(row=0, column=2, rowspan=2, padx=20)
        for index, feature in enumerate(group.features):
            row = 1 + index // 2
            column = index % 2
            card, handle = create_feature_card(
                page, group.icon, feature.title, feature.description, feature.command,
            )
            card.grid(
                row=row, column=column, sticky="ew",
                padx=(0, 8) if column == 0 else (8, 0), pady=(0, 14),
            )
            self.ctx.buttons[feature.attr_name] = handle
            self.ctx.feature_cards[feature.attr_name] = card
        return page

    def show_page(self, key):
        if key not in self.pages:
            raise KeyError(f"未知页面：{key}")
        if self.current_page == key:
            return
        if self.current_page is not None:
            self.pages[self.current_page].grid_remove()
        self.pages[key].grid()
        self.pages[key].tkraise()
        self.current_page = key
        group = next((item for item in self.feature_groups if item.key == key), None)
        title = group.title if group else "工作台总览"
        subtitle = group.summary if group else "选择一个工作场景，开始处理表格任务"
        self.page_title_label.configure(text=title)
        self.page_subtitle_label.configure(text=subtitle)
        for page_key, button in self.nav_buttons.items():
            active = page_key == key
            button.configure(
                fg_color=theme.ACCENT if active else "transparent",
                hover_color=theme.ACCENT_HOVER if active else theme.SIDEBAR_HOVER,
                text_color=theme.SIDEBAR_TEXT if active else theme.SIDEBAR_TEXT_MUTED,
                font=theme.NAV_FONT_ACTIVE if active else theme.NAV_FONT,
            )

    def _create_log_area(self, parent):
        self.frame_bottom = ctk.CTkFrame(
            parent, fg_color=theme.SURFACE, corner_radius=0,
            height=theme.LOG_AREA_HEIGHT, border_width=0,
        )
        self.frame_bottom.grid(row=2, column=0, sticky="ew")
        self.frame_bottom.grid_propagate(False)
        self.frame_bottom.grid_columnconfigure(0, weight=1)
        self.frame_bottom.grid_rowconfigure(1, weight=1)
        log_header = ctk.CTkFrame(self.frame_bottom, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=24, pady=(8, 5))
        log_header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            log_header, text="●", font=theme.SMALL_FONT, text_color=theme.SUCCESS,
        ).grid(row=0, column=0, padx=(0, 7))
        ctk.CTkLabel(
            log_header, text="运行记录", font=theme.LOG_TITLE_FONT,
            text_color=theme.TEXT, anchor="w",
        ).grid(row=0, column=1, sticky="w")
        self.log_toggle = AppButton(
            log_header, text="收起", command=self._toggle_log, width=54, height=26,
            font=theme.SMALL_FONT, fg_color="transparent",
            hover_color=theme.BUTTON_SECONDARY_HOVER, text_color=theme.TEXT_SECONDARY,
            corner_radius=7,
        )
        self.log_toggle.grid(row=0, column=2, sticky="e")
        self.log_text = ctk.CTkTextbox(
            self.frame_bottom, state="disabled", height=theme.LOG_TEXT_HEIGHT,
            font=theme.LOG_FONT, fg_color=theme.SURFACE_SUNKEN,
            text_color=theme.TEXT_STRONG, border_width=1,
            border_color=theme.BORDER_SOFT, corner_radius=9, wrap="word",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 10))

    def _toggle_log(self):
        self.log_collapsed = not self.log_collapsed
        if self.log_collapsed:
            self.log_text.grid_remove()
            self.frame_bottom.configure(height=theme.LOG_COLLAPSED_HEIGHT)
            self.log_toggle.configure(text="展开")
        else:
            self.log_text.grid()
            self.frame_bottom.configure(height=theme.LOG_AREA_HEIGHT)
            self.log_toggle.configure(text="收起")

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
            "老头表格助手用于处理常见 Excel 批量整理工作。\n批量修改类功能请先备份或先用副本试跑。\n详细说明请查看同目录下的 使用说明.txt。",
            dialog_width=460, dialog_height=280, wraplength=420,
        )

    def show_about(self):
        self.ctx.dialogs.show_info(
            "关于",
            f"软件名：{APP_NAME}\n版本号：v{APP_VERSION}\n说明：面向审计、财务、报表整理场景的 Excel 效率工具\n界面版本：场景化工作台",
            dialog_width=460, dialog_height=300, wraplength=380,
        )

    def _run_gui_smoke_and_exit(self):
        try:
            self.ctx.log_info("GUI 冒烟：开始检查关键入口。")
            for page_key in self.pages:
                self.show_page(page_key)
                self.ctx.flush_ui()
            self.show_page("home")
            for card in self.ctx.feature_cards.values():
                if not getattr(card, "_feature_enabled", True):
                    raise RuntimeError("GUI 冒烟：存在不可用功能卡片。")
            for handle in self.ctx.buttons.values():
                handle.config(state="disabled")
                handle.config(state="normal")
            self._toggle_log()
            self._toggle_log()
            for dialog_callback in (
                self.show_about,
                self.controllers.rename_files.run_batch_rename_files,
                self.controllers.data_drill.run_data_drill,
                self.controllers.templates.run_template_generate,
                self.controllers.color_tools.run_clear_by_color,
            ):
                dialog_callback()
                self.ctx.flush_ui()
            self.ctx.log_info("GUI 冒烟：完成。")
        finally:
            self.root.after(200, self.root.destroy)
