# -*- coding: utf-8 -*-
"""可复用 UI 控件：安全按钮、功能卡片句柄、功能卡片与分区面板工厂。"""
import tkinter as tk

import customtkinter as ctk

from . import theme


class AppButton(ctk.CTkButton):
    """在控件已销毁时安全忽略 config 调用的按钮（关闭窗口/冒烟场景）。"""

    def config(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        try:
            if not self.winfo_exists():
                return None
            return self.configure(**kwargs)
        except tk.TclError:
            return None

    configure = ctk.CTkButton.configure


class FeatureCardHandle:
    """功能卡片的轻量句柄，暴露 Tkinter 风格的 config(state=...) 接口。"""

    def __init__(self, set_state):
        self._set_state = set_state

    def config(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        return self.configure(**kwargs)

    def configure(self, **kwargs):
        state = kwargs.get("state")
        if state is not None:
            try:
                self._set_state(state)
            except tk.TclError:
                return None


def create_feature_card(parent, title, description, command):
    """创建一张可点击的功能卡片，返回 (card, handle)。"""
    has_description = bool(description)
    card = ctk.CTkFrame(
        parent,
        height=theme.CARD_HEIGHT_WITH_DESC if has_description else theme.CARD_HEIGHT,
        fg_color=theme.SURFACE_RAISED,
        border_width=1,
        border_color=theme.BORDER,
        corner_radius=10,
    )
    card.grid_propagate(False)
    card.grid_columnconfigure(0, weight=1)

    title_label = ctk.CTkLabel(
        card,
        text=title,
        font=theme.CARD_TITLE_FONT,
        text_color=theme.TEXT,
        anchor="w",
    )
    title_label.grid(
        row=0,
        column=0,
        sticky="ew",
        padx=(14, 8),
        pady=(10, 2) if has_description else (0, 0),
    )
    widgets = [card, title_label]
    desc_label = None
    if has_description:
        desc_label = ctk.CTkLabel(
            card,
            text=description,
            font=theme.CARD_DESCRIPTION_FONT,
            text_color=theme.TEXT_SECONDARY,
            anchor="w",
            justify="left",
        )
        desc_label.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(2, 12))
        widgets.append(desc_label)
    else:
        card.grid_rowconfigure(0, weight=1)

    def set_state(state):
        is_disabled = state == "disabled"
        card.configure(
            fg_color=theme.SURFACE_DISABLED if is_disabled else theme.SURFACE_RAISED,
            border_color=theme.BORDER_DISABLED if is_disabled else theme.BORDER,
        )
        title_label.configure(text_color=theme.TEXT_DISABLED if is_disabled else theme.TEXT)
        if desc_label is not None:
            desc_label.configure(text_color=theme.TEXT_DISABLED if is_disabled else theme.TEXT_SECONDARY)
        card._feature_enabled = not is_disabled

    card._feature_enabled = True
    handle = FeatureCardHandle(set_state)
    _bind_card_action(card, command, widgets)
    return card, handle


def _bind_card_action(card, command, widgets):
    def is_enabled():
        return getattr(card, "_feature_enabled", True)

    def run_action(event=None):
        if is_enabled():
            command()

    def on_enter(event=None):
        if is_enabled():
            card.configure(fg_color=theme.SURFACE_HOVER, border_color=theme.BORDER_HOVER)

    def on_leave(event=None):
        if is_enabled():
            card.configure(fg_color=theme.SURFACE_RAISED, border_color=theme.BORDER)

    for widget in widgets:
        widget.bind("<Button-1>", run_action)
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
        try:
            widget.configure(cursor="hand2")
        except tk.TclError:
            pass


def create_section_panel(parent, title, features, register_card):
    """创建一个功能分区面板。

    features: (feature_title, description, attr_name, command) 元组列表。
    register_card(attr_name, card, handle) 由调用方注册卡片句柄。
    """
    panel = ctk.CTkFrame(
        parent,
        fg_color=theme.SURFACE,
        corner_radius=12,
        border_width=1,
        border_color=theme.BORDER_SOFT,
    )
    panel.grid_columnconfigure(0, weight=1)

    title_row = ctk.CTkFrame(panel, fg_color="transparent")
    title_row.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
    ctk.CTkFrame(
        title_row,
        width=theme.SECTION_BAR_WIDTH,
        height=theme.SECTION_BAR_HEIGHT,
        fg_color=theme.ACCENT,
        corner_radius=2,
    ).pack(side="left", padx=(0, 8))
    ctk.CTkLabel(
        title_row,
        text=title,
        font=theme.SECTION_TITLE_FONT,
        text_color=theme.TEXT,
    ).pack(side="left")

    for row, (feature_title, description, attr_name, command) in enumerate(features, start=1):
        card, handle = create_feature_card(panel, feature_title, description, command)
        card.grid(row=row, column=0, sticky="ew", padx=10, pady=(0, 9))
        register_card(attr_name, card, handle)

    return panel
