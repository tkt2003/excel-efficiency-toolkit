# -*- coding: utf-8 -*-
"""通用 no-grab 对话框服务。

对话框不使用 grab_set 独占输入，通过 wait_variable + transient + 短暂 topmost
组合保证前置显示；这些语义是刻意的，请勿"清理"。
EXCEL_TOOLKIT_GUI_SMOKE=1 时对话框会自动关闭，供 GUI 冒烟使用。
"""
import os
import tkinter as tk

import customtkinter as ctk

from . import theme
from .widgets import AppButton


class DialogService:
    """基于给定主窗口 root 的对话框工具集。"""

    def __init__(self, root):
        self.root = root

    # ------------------------------------------------------------------
    # 基础构件
    # ------------------------------------------------------------------
    def create_dialog_card(self, title, resizable=False):
        dialog = ctk.CTkToplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(resizable, resizable)
        dialog.transient(self.root)
        dialog.configure(fg_color=theme.BG)
        dialog.minsize(1, 1)

        card = ctk.CTkFrame(
            dialog,
            fg_color=theme.SURFACE,
            corner_radius=10,
            border_width=1,
            border_color=theme.BORDER_SOFT,
        )
        outer_pad = 16
        card.pack(fill=tk.BOTH, expand=True, padx=outer_pad, pady=outer_pad)
        dialog._dialog_card = card

        ctk.CTkLabel(
            card,
            text=title,
            font=theme.DIALOG_TITLE_FONT,
            text_color=theme.TEXT,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        return dialog, card

    def add_message(self, parent, message, wraplength=360, pady=(0, 12)):
        label = ctk.CTkLabel(
            parent,
            text=message,
            font=theme.DIALOG_BODY_FONT,
            wraplength=wraplength,
            justify="left",
            anchor="w",
            text_color=theme.TEXT,
        )
        label.pack(anchor="w", fill=tk.X, padx=16, pady=pady)
        return label

    def create_button_bar(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(padx=12, pady=(6, 14), fill=tk.X)
        return button_frame

    def add_button(self, parent, text, command, primary=False, width=None):
        button_width = width or max(86, min(260, len(text) * 14 + 28))
        return AppButton(
            parent,
            text=text,
            command=command,
            font=theme.DIALOG_BUTTON_FONT,
            width=button_width,
            height=theme.DIALOG_BUTTON_HEIGHT,
            fg_color=theme.ACCENT if primary else theme.BUTTON_SECONDARY_BG,
            hover_color=theme.ACCENT_HOVER if primary else theme.BUTTON_SECONDARY_HOVER,
            text_color=theme.TEXT_ON_ACCENT if primary else theme.TEXT,
            border_width=0 if primary else 1,
            border_color=theme.BORDER,
            corner_radius=7,
        )

    # ------------------------------------------------------------------
    # 显示与关闭
    # ------------------------------------------------------------------
    def center_window(self, window, width=None, height=None):
        window.update_idletasks()
        width = width or window.winfo_reqwidth()
        height = height or window.winfo_reqheight()

        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = int((screen_width - width) / 2)
        y = int((screen_height - height) / 2)
        screen_margin = 16
        x = min(max(screen_margin, x), max(screen_margin, screen_width - width - screen_margin))
        y = min(max(screen_margin, y), max(screen_margin, screen_height - height - screen_margin))
        window.geometry(f"{width}x{height}+{x}+{y}")

    def show_no_grab(
        self,
        dialog,
        focus_widget=None,
        min_width=380,
        min_height=None,
        width=None,
        height=None,
    ):
        final_width = width if width is not None else max(dialog.winfo_reqwidth(), min_width)
        dialog.update_idletasks()
        final_height = height if height is not None else max(dialog.winfo_reqheight(), min_height or 300)
        self.center_window(dialog, width=final_width, height=final_height)
        dialog.deiconify()
        dialog.lift()
        try:
            dialog.attributes("-topmost", True)
            dialog.after(200, lambda: dialog.winfo_exists() and dialog.attributes("-topmost", False))
        except tk.TclError:
            pass
        try:
            dialog.focus_force()
        except tk.TclError:
            pass
        if focus_widget is not None:
            try:
                focus_widget.focus_set()
            except tk.TclError:
                pass

    def schedule_smoke_close(self, dialog, close_callback=None, delay_ms=250):
        if os.environ.get("EXCEL_TOOLKIT_GUI_SMOKE") != "1":
            return

        def close():
            if dialog.winfo_exists():
                try:
                    if close_callback is None:
                        dialog.destroy()
                    else:
                        close_callback()
                except tk.TclError:
                    pass

        dialog.after(delay_ms, close)

    # ------------------------------------------------------------------
    # 常用对话框
    # ------------------------------------------------------------------
    def ask_text(
        self,
        title,
        prompt,
        default="",
        entry_width=30,
        dialog_width=520,
        dialog_height=300,
        wraplength=460,
    ):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.create_dialog_card(title)

        self.add_message(card, prompt, wraplength=wraplength, pady=(0, 10))

        entry = ctk.CTkEntry(
            card,
            font=theme.ENTRY_FONT,
            width=entry_width * 10,
            height=theme.ENTRY_HEIGHT,
            fg_color=theme.SURFACE_RAISED,
            border_color=theme.BORDER,
            text_color=theme.TEXT,
        )
        entry.pack(padx=16, pady=(0, 10), fill=tk.X)
        entry.insert(0, default)
        entry.select_range(0, tk.END)

        button_frame = self.create_button_bar(card)

        def confirm():
            result["value"] = entry.get()
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.add_button(button_frame, "确定", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Return>", lambda event: confirm())
        dialog.bind("<Escape>", lambda event: cancel())
        self.show_no_grab(dialog, focus_widget=entry, width=dialog_width, height=dialog_height)
        self.schedule_smoke_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["value"]

    def ask_choice(
        self,
        title,
        prompt,
        choices,
        dialog_width=760,
        dialog_height=380,
        wraplength=680,
    ):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.create_dialog_card(title)

        self.add_message(card, prompt, wraplength=wraplength, pady=(0, 12))

        button_frame = self.create_button_bar(card)

        def choose(value):
            result["value"] = value
            done.set(True)
            dialog.destroy()

        self.add_button(button_frame, "取消", lambda: choose(None), width=70).pack(side=tk.RIGHT)

        for label, value in reversed(choices):
            self.add_button(
                button_frame,
                label,
                lambda selected=value: choose(selected),
                primary=True,
                width=min(max(128, len(label) * 10 + 36), 180),
            ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        dialog.bind("<Escape>", lambda event: choose(None))
        self.show_no_grab(dialog, width=dialog_width, height=dialog_height)
        self.schedule_smoke_close(dialog, lambda: choose(None))
        self.root.wait_variable(done)
        return result["value"]

    def show_info(self, title, message, dialog_width=620, dialog_height=380, wraplength=500):
        dialog, card = self.create_dialog_card(title)

        self.add_message(card, message, wraplength=wraplength, pady=(0, 12))

        button_frame = self.create_button_bar(card)
        self.add_button(button_frame, "确定", dialog.destroy, primary=True).pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.show_no_grab(dialog, width=dialog_width, height=dialog_height)
        self.schedule_smoke_close(dialog)

    def ask_positive_int(self, title, prompt, initialvalue):
        value = self.ask_text(title, prompt, default=str(initialvalue))
        if value is None:
            return None
        try:
            return int(value.strip())
        except ValueError:
            raise ValueError(f"{prompt}必须是整数。")
