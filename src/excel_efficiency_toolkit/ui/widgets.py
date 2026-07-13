# -*- coding: utf-8 -*-
"""v1.1 工作台使用的可复用交互组件。"""
import tkinter as tk

import customtkinter as ctk

from . import theme


class AppButton(ctk.CTkButton):
    """窗口关闭后仍可安全接收控制器的 config 调用。"""

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
    """向业务控制器暴露 Tkinter 风格的启用/禁用接口。"""

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


def _bind_clickable(surface, widgets, command, enter, leave):
    def run_action(event=None):
        if surface._feature_enabled:
            command()

    def on_enter(event=None):
        if surface._feature_enabled:
            enter()

    def on_leave(event=None):
        if surface._feature_enabled:
            leave()

    for widget in widgets:
        widget.bind("<Button-1>", run_action)
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
        widget.configure(cursor="hand2")


def create_feature_card(parent, icon, title, description, command):
    """创建带图标、说明和明确进入暗示的功能卡片。"""
    card = ctk.CTkFrame(
        parent,
        height=theme.FEATURE_CARD_HEIGHT,
        fg_color=theme.SURFACE,
        border_width=1,
        border_color=theme.BORDER_SOFT,
        corner_radius=14,
    )
    card.grid_propagate(False)
    card.grid_columnconfigure(1, weight=1)
    card.grid_rowconfigure(0, weight=1)

    icon_box = ctk.CTkFrame(
        card,
        width=42,
        height=42,
        fg_color=theme.ACCENT_PALE,
        corner_radius=11,
    )
    icon_box.grid(row=0, column=0, padx=(16, 12), pady=18)
    icon_box.grid_propagate(False)
    icon_label = ctk.CTkLabel(
        icon_box,
        text=icon,
        font=theme.CARD_ICON_FONT,
        text_color=theme.ACCENT_DARK,
    )
    icon_label.place(relx=0.5, rely=0.5, anchor="center")

    copy = ctk.CTkFrame(card, fg_color="transparent")
    copy.grid(row=0, column=1, sticky="ew", pady=15)
    copy.grid_columnconfigure(0, weight=1)
    title_label = ctk.CTkLabel(
        copy,
        text=title,
        font=theme.CARD_TITLE_FONT,
        text_color=theme.TEXT,
        anchor="w",
    )
    title_label.grid(row=0, column=0, sticky="ew")
    description_label = ctk.CTkLabel(
        copy,
        text=description,
        font=theme.CARD_DESCRIPTION_FONT,
        text_color=theme.TEXT_SECONDARY,
        anchor="w",
        justify="left",
        wraplength=360,
    )
    description_label.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    action_label = ctk.CTkLabel(
        card,
        text="›",
        width=28,
        font=theme.CARD_ACTION_FONT,
        text_color=theme.TEXT_MUTED,
    )
    action_label.grid(row=0, column=2, padx=(6, 14))

    widgets = (card, icon_box, icon_label, copy, title_label, description_label, action_label)

    def set_state(state):
        disabled = state == "disabled"
        card._feature_enabled = not disabled
        card.configure(
            fg_color=theme.SURFACE_DISABLED if disabled else theme.SURFACE,
            border_color=theme.BORDER_DISABLED if disabled else theme.BORDER_SOFT,
        )
        icon_box.configure(fg_color=theme.SURFACE_DISABLED if disabled else theme.ACCENT_PALE)
        title_label.configure(text_color=theme.TEXT_DISABLED if disabled else theme.TEXT)
        description_label.configure(text_color=theme.TEXT_DISABLED if disabled else theme.TEXT_SECONDARY)
        action_label.configure(text_color=theme.TEXT_DISABLED if disabled else theme.TEXT_MUTED)

    def enter():
        card.configure(fg_color=theme.SURFACE_HOVER, border_color=theme.BORDER_HOVER)
        action_label.configure(text_color=theme.ACCENT)

    def leave():
        card.configure(fg_color=theme.SURFACE, border_color=theme.BORDER_SOFT)
        action_label.configure(text_color=theme.TEXT_MUTED)

    card._feature_enabled = True
    _bind_clickable(card, widgets, command, enter, leave)
    return card, FeatureCardHandle(set_state)


def create_category_card(parent, icon, title, summary, count, command):
    """创建首页场景入口卡片。"""
    card = ctk.CTkFrame(
        parent,
        height=theme.CATEGORY_CARD_HEIGHT,
        fg_color=theme.SURFACE,
        border_width=1,
        border_color=theme.BORDER_SOFT,
        corner_radius=16,
    )
    card.grid_propagate(False)
    card.grid_columnconfigure(1, weight=1)
    card._feature_enabled = True

    badge = ctk.CTkLabel(
        card,
        text=icon,
        width=42,
        height=42,
        corner_radius=12,
        fg_color=theme.ACCENT_SOFT,
        text_color=theme.ACCENT_DARK,
        font=theme.CARD_ICON_FONT,
    )
    badge.grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=(18, 18), sticky="n")

    title_label = ctk.CTkLabel(
        card,
        text=title,
        font=theme.CARD_TITLE_FONT,
        text_color=theme.TEXT,
        anchor="w",
    )
    title_label.grid(row=0, column=1, sticky="sew", pady=(18, 0))
    count_label = ctk.CTkLabel(
        card,
        text=f"{count} 项工具  ›",
        font=theme.SECTION_META_FONT,
        text_color=theme.ACCENT,
        anchor="e",
    )
    count_label.grid(row=0, column=2, sticky="se", padx=(8, 16), pady=(18, 0))
    summary_label = ctk.CTkLabel(
        card,
        text=summary,
        font=theme.CARD_DESCRIPTION_FONT,
        text_color=theme.TEXT_SECONDARY,
        anchor="nw",
        justify="left",
        wraplength=360,
    )
    summary_label.grid(row=1, column=1, columnspan=2, sticky="new", padx=(0, 16), pady=(6, 16))

    widgets = (card, badge, title_label, count_label, summary_label)
    _bind_clickable(
        card,
        widgets,
        command,
        lambda: card.configure(fg_color=theme.SURFACE_HOVER, border_color=theme.BORDER_HOVER),
        lambda: card.configure(fg_color=theme.SURFACE, border_color=theme.BORDER_SOFT),
    )
    return card
