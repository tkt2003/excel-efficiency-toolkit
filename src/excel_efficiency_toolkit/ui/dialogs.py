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

    def ask_column(
        self,
        title,
        prompt,
        columns: list[tuple[str, str]],
        default="A",
        allow_empty=False,
        dialog_width=520,
        dialog_height=480,
    ):
        """通用列选择弹窗。

        展示列表如 'A | 物料编码'，支持点击选择，也支持手动输入列字母。
        返回选中的列字母（如 'B'），若取消返回 None，若允许空且用户跳过返回空字符串 ''。
        """
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.create_dialog_card(title)

        self.add_message(card, prompt, wraplength=460, pady=(0, 8))

        scroll_frame = ctk.CTkScrollableFrame(
            card,
            height=200,
            fg_color=theme.SURFACE_SUNKEN,
            border_width=1,
            border_color=theme.BORDER_SOFT,
            corner_radius=8,
        )
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 10))

        input_frame = ctk.CTkFrame(card, fg_color="transparent")
        input_frame.pack(fill=tk.X, padx=16, pady=(0, 10))

        ctk.CTkLabel(
            input_frame,
            text="选择或输入列字母：",
            font=theme.LABEL_FONT,
            text_color=theme.TEXT,
        ).pack(side=tk.LEFT, padx=(0, 8))

        entry = ctk.CTkEntry(
            input_frame,
            font=theme.ENTRY_FONT,
            width=120,
            height=theme.ENTRY_HEIGHT,
            fg_color=theme.SURFACE_RAISED,
            border_color=theme.BORDER,
            text_color=theme.TEXT,
        )
        entry.pack(side=tk.LEFT)
        entry.insert(0, default or "")

        item_buttons = []

        def select_col(col_letter):
            entry.delete(0, tk.END)
            entry.insert(0, col_letter)
            for b_col, b_widget in item_buttons:
                if b_col == col_letter:
                    b_widget.configure(
                        fg_color=theme.ACCENT,
                        text_color=theme.TEXT_ON_ACCENT,
                    )
                else:
                    b_widget.configure(
                        fg_color="transparent",
                        text_color=theme.TEXT,
                    )

        for col, col_title in columns:
            display_text = f"{col} | {col_title}" if col_title else f"{col} | (空白)"
            btn = AppButton(
                scroll_frame,
                text=display_text,
                command=lambda c=col: select_col(c),
                anchor="w",
                font=theme.DIALOG_BODY_FONT,
                height=30,
                fg_color=theme.ACCENT if col == default else "transparent",
                text_color=theme.TEXT_ON_ACCENT if col == default else theme.TEXT,
                hover_color=theme.BUTTON_SECONDARY_HOVER,
                corner_radius=6,
            )
            btn.pack(fill=tk.X, padx=4, pady=2)
            btn.bind("<Double-Button-1>", lambda event, c=col: (select_col(c), confirm()))
            item_buttons.append((col, btn))

        button_frame = self.create_button_bar(card)

        def confirm():
            val = entry.get().strip().upper()
            if not val:
                if allow_empty:
                    result["value"] = ""
                    done.set(True)
                    dialog.destroy()
                    return
                return
            result["value"] = val
            done.set(True)
            dialog.destroy()

        def skip():
            result["value"] = ""
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        if allow_empty:
            self.add_button(button_frame, "不使用 / 跳过", skip).pack(side=tk.RIGHT, padx=(0, 8))
        self.add_button(button_frame, "确定", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Return>", lambda event: confirm())
        dialog.bind("<Escape>", lambda event: cancel())

        self.show_no_grab(dialog, focus_widget=entry, width=dialog_width, height=dialog_height)
        self.schedule_smoke_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["value"]

    def ask_split_preview(
        self,
        title: str,
        preview_items: list[tuple[str, str]],
        prompt: str = "请确认以下拆分参数是否正确，确认后将开始执行：",
        dialog_width=520,
        dialog_height=420,
    ) -> bool:
        """拆分执行前预览确认弹窗。返回 True（确认执行）或 False（取消）。"""
        result = {"confirmed": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.create_dialog_card(title)

        self.add_message(card, prompt, wraplength=460, pady=(0, 10))

        info_frame = ctk.CTkFrame(
            card,
            fg_color=theme.SURFACE_SUNKEN,
            border_width=1,
            border_color=theme.BORDER_SOFT,
            corner_radius=8,
        )
        info_frame.pack(fill=tk.X, padx=16, pady=(0, 14))

        for row_idx, (label, val) in enumerate(preview_items):
            ctk.CTkLabel(
                info_frame,
                text=f"{label}：",
                font=theme.LABEL_BOLD_FONT,
                text_color=theme.TEXT_SECONDARY,
                anchor="e",
                width=110,
            ).grid(row=row_idx, column=0, sticky="e", padx=(14, 6), pady=4)

            ctk.CTkLabel(
                info_frame,
                text=str(val),
                font=theme.DIALOG_BODY_FONT,
                text_color=theme.TEXT_STRONG,
                anchor="w",
                justify="left",
                wraplength=320,
            ).grid(row=row_idx, column=1, sticky="w", padx=(0, 14), pady=4)

        button_frame = self.create_button_bar(card)

        def confirm():
            result["confirmed"] = True
            done.set(True)
            dialog.destroy()

        def cancel():
            result["confirmed"] = False
            done.set(True)
            dialog.destroy()

        self.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.add_button(button_frame, "开始执行", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Return>", lambda event: confirm())
        dialog.bind("<Escape>", lambda event: cancel())

        self.show_no_grab(dialog, width=dialog_width, height=dialog_height)
        self.schedule_smoke_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["confirmed"]

    def create_progress_cancel_dialog(
        self,
        title: str,
        initial_message: str,
        on_cancel=None,
        dialog_width=420,
        dialog_height=200,
    ):
        """创建执行过程中的取消控制弹窗，返回控制器对象（包含 update_message 与 close 方法）。"""
        dialog, card = self.create_dialog_card(title)

        status_label = self.add_message(card, initial_message, wraplength=360, pady=(0, 14))
        button_frame = self.create_button_bar(card)

        cancelled = [False]

        def cancel():
            if not cancelled[0]:
                cancelled[0] = True
                cancel_btn.configure(text="正在取消...", state="disabled")
                if on_cancel:
                    on_cancel()

        cancel_btn = self.add_button(button_frame, "取消执行", cancel, primary=False)
        cancel_btn.pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda event: cancel())

        self.show_no_grab(dialog, width=dialog_width, height=dialog_height)

        class ProgressController:
            def update_message(self_, msg):
                try:
                    if dialog.winfo_exists():
                        status_label.configure(text=msg)
                except tk.TclError:
                    pass

            def close(self_):
                try:
                    if dialog.winfo_exists():
                        dialog.destroy()
                except tk.TclError:
                    pass

        return ProgressController()
