# -*- coding: utf-8 -*-
"""Word 批量替换：统一替换 / 分别替换，预览确认后执行。"""
import os
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

from ...text_display import format_display_row
from ...word_replace_ops import (
    MODE_COMMON,
    MODE_PER_FILE,
    build_common_replace_rules_from_entries,
    create_per_file_replace_rule_template,
    execute_word_replace,
    has_empty_replacement,
    load_common_replace_rules,
    load_per_file_replace_rules,
    preview_word_replace,
)
from .. import theme
from .base import FeatureController


class WordReplaceController(FeatureController):
    def run_word_batch_replace(self):
        """按钮回调函数，批量替换 Word 文件文字。"""
        self.logger.info("Word 批量替换：开始操作。")
        mode = self.ask_choice(
            "Word 批量替换",
            "请选择替换方式：\n\n统一替换：多个 Word 文件替换成同样内容\n分别替换：不同 Word 文件替换成不同内容",
            [
                ("统一替换", MODE_COMMON),
                ("分别替换", MODE_PER_FILE),
            ],
            dialog_width=640,
            dialog_height=300,
            wraplength=560,
        )
        if mode is None:
            self.logger.info("用户已取消操作")
            return

        word_paths = filedialog.askopenfilenames(
            title="请选择要批量替换的 Word 文件",
            filetypes=[
                ("Word 文件", "*.docx *.docm *.doc"),
                ("所有文件", "*.*"),
            ],
        )
        if not word_paths:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_word_replace").config(state="disabled")
        try:
            if mode == MODE_COMMON:
                self._run_common_word_replace_flow(list(word_paths))
            else:
                self._run_per_file_word_replace_flow(list(word_paths))
        except Exception as e:
            self.logger.error(f"Word 批量替换失败：{e}")
            self.show_info(
                "Word 批量替换",
                f"Word 批量替换失败：{e}",
                dialog_width=680,
                wraplength=620,
            )
        finally:
            self.button("btn_word_replace").config(state="normal")

    def _run_common_word_replace_flow(self, word_paths):
        rules = self._ask_common_word_replace_rules()
        if rules is None:
            self.logger.info("用户已取消操作")
            return
        self._preview_and_confirm_word_replace(MODE_COMMON, word_paths, rules)

    def _run_per_file_word_replace_flow(self, word_paths):
        try:
            rule_workbook_path = create_per_file_replace_rule_template(word_paths)
            self.logger.info(f"已自动生成 Word 分别替换规则表：{rule_workbook_path}")
        except Exception as e:
            self.logger.error(f"生成 Word 分别替换规则表失败：{e}")
            self.show_info(
                "分别替换",
                f"生成规则表失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
            return

        try:
            os.startfile(rule_workbook_path)
            self.logger.info("规则表已自动打开。请填写并保存后，回到工具开始预览。")
        except Exception as e:
            self.logger.error(f"规则表已生成，但自动打开失败：{e}")
            self.show_info(
                "分别替换",
                f"规则表已生成，但自动打开失败：{e}",
                dialog_width=680,
                wraplength=620,
            )
            return

        rule_workbook_path = self._ask_per_file_word_rule_workbook(rule_workbook_path)
        if not rule_workbook_path:
            self.logger.info("用户已取消操作")
            return
        try:
            rules = load_per_file_replace_rules(rule_workbook_path)
        except Exception as e:
            self.logger.error(f"读取 Word 分别替换规则表失败：{e}")
            self.show_info(
                "分别替换",
                "请先保存规则表；如仍失败，请关闭 Excel 后重试。",
                dialog_width=620,
                wraplength=560,
            )
            return
        self._preview_and_confirm_word_replace(MODE_PER_FILE, word_paths, rules)

    def _preview_and_confirm_word_replace(self, mode, word_paths, rules):
        rule_count = self._count_word_replace_rules(rules)
        if rule_count == 0:
            self.logger.info("没有可执行的 Word 替换规则。")
            self.show_info(
                "Word 批量替换",
                "没有可执行的替换规则，请检查“查找内容”和“是否启用”。",
                dialog_width=620,
                wraplength=560,
            )
            return

        self.logger.info(f"已选择 Word 文件数量：{len(word_paths)}")
        self.logger.info(f"已读取 Word 替换规则数量：{rule_count}")
        if has_empty_replacement(rules) and not self._confirm_word_empty_replacement():
            self.logger.info("用户已取消操作")
            return

        self.logger.info("正在预览 Word 替换计划。")
        preview_result = self._preview_word_replace(mode, word_paths, rules)
        self._log_word_replace_preview(preview_result)

        if preview_result["summary"]["pending_count"] == 0:
            self.show_info(
                "Word 批量替换",
                "预览完成，没有需要执行的替换。\n请检查规则和所选 Word 文件。",
                dialog_width=680,
                wraplength=620,
            )
            return

        if not self._confirm_word_replace_preview_ready(preview_result):
            self.logger.info("用户已取消操作")

    def _ask_common_word_replace_rules(self):
        result = {"rules": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("统一替换", resizable=True)

        self.dialogs.add_message(
            card,
            "多个 Word 文件会使用下面相同规则替换。请填写“查找内容”和“替换为”，空行会自动忽略。",
            wraplength=780,
            pady=(0, 10),
        )

        header_frame = ctk.CTkFrame(card, fg_color="transparent")
        header_frame.pack(fill=tk.X, padx=16, pady=(0, 4))
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            header_frame,
            text="查找内容",
            font=theme.LABEL_BOLD_FONT,
            text_color=theme.TEXT,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(
            header_frame,
            text="替换为",
            font=theme.LABEL_BOLD_FONT,
            text_color=theme.TEXT,
        ).grid(row=0, column=1, sticky="w", padx=(8, 8))

        rows_frame = ctk.CTkScrollableFrame(
            card,
            height=250,
            fg_color=theme.SURFACE_RAISED,
            border_width=1,
            border_color=theme.BORDER,
            scrollbar_button_color=theme.SCROLLBAR,
            scrollbar_button_hover_color=theme.SCROLLBAR_HOVER,
        )
        rows_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 10))
        rows_frame.grid_columnconfigure(0, weight=1)
        rows_frame.grid_columnconfigure(1, weight=1)
        rule_rows = []

        def refresh_rows():
            for index, row in enumerate(rule_rows):
                if row["frame"].winfo_exists():
                    row["frame"].grid(row=index, column=0, columnspan=3, sticky="ew", padx=6, pady=(6, 0))

        def remove_row(row):
            if row in rule_rows:
                rule_rows.remove(row)
            if row["frame"].winfo_exists():
                row["frame"].destroy()
            if not rule_rows:
                add_rule_row()
            refresh_rows()

        def add_rule_row(find_text="", replace_text=""):
            row_frame = ctk.CTkFrame(rows_frame, fg_color="transparent")
            row_frame.grid_columnconfigure(0, weight=1)
            row_frame.grid_columnconfigure(1, weight=1)
            find_entry = ctk.CTkEntry(
                row_frame,
                font=theme.LABEL_FONT,
                height=theme.ENTRY_HEIGHT,
                fg_color=theme.SURFACE,
                border_color=theme.BORDER,
                text_color=theme.TEXT,
            )
            find_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            find_entry.insert(0, find_text)
            replace_entry = ctk.CTkEntry(
                row_frame,
                font=theme.LABEL_FONT,
                height=theme.ENTRY_HEIGHT,
                fg_color=theme.SURFACE,
                border_color=theme.BORDER,
                text_color=theme.TEXT,
            )
            replace_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            replace_entry.insert(0, replace_text)
            row = {"frame": row_frame, "find": find_entry, "replace": replace_entry}
            delete_button = self.dialogs.add_button(row_frame, "删除", lambda: remove_row(row), width=62)
            delete_button.grid(row=0, column=2, sticky="e")
            rule_rows.append(row)
            refresh_rows()
            find_entry.focus_set()

        def import_from_excel():
            rule_workbook_path = filedialog.askopenfilename(
                title="请选择统一替换规则 Excel",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                    ("所有文件", "*.*"),
                ],
            )
            if not rule_workbook_path:
                return
            try:
                imported_rules = load_common_replace_rules(rule_workbook_path)
            except Exception as e:
                self.show_info(
                    "统一替换",
                    f"导入规则失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
                return
            for row in list(rule_rows):
                row["frame"].destroy()
            rule_rows.clear()
            for rule in imported_rules:
                add_rule_row(rule.find_text, rule.replace_text)
            self.logger.info(f"已从 Excel 导入统一替换规则：{len(imported_rules)} 条。")

        def preview():
            entries = [(row["find"].get(), row["replace"].get()) for row in rule_rows]
            try:
                result["rules"] = build_common_replace_rules_from_entries(entries)
            except ValueError as e:
                self.show_info(
                    "统一替换",
                    str(e),
                    dialog_width=600,
                    wraplength=540,
                )
                return
            done.set(True)
            dialog.destroy()

        def cancel():
            done.set(True)
            dialog.destroy()

        add_rule_row()

        button_frame = self.dialogs.create_button_bar(card)
        self.dialogs.add_button(button_frame, "从 Excel 导入规则", import_from_excel, width=160).pack(side=tk.LEFT)
        self.dialogs.add_button(button_frame, "增加一行", add_rule_row, width=105).pack(side=tk.LEFT, padx=(8, 0))
        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.dialogs.add_button(button_frame, "预览", preview, primary=True, width=96).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=860, height=520)
        self.root.wait_variable(done)
        return result["rules"]

    def _ask_per_file_word_rule_workbook(self, rule_workbook_path):
        result = {"path": rule_workbook_path}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("分别替换")

        self.dialogs.add_message(
            card,
            (
                "不同 Word 文件替换成不同内容。\n\n"
                "规则表已自动生成并打开。\n"
                "请在 Excel 中填写“查找内容”“替换为”，保存后回到这里点击“我已填写并保存，开始预览”。"
            ),
            wraplength=580,
            pady=(0, 10),
        )

        path_label = ctk.CTkLabel(
            card,
            text=f"当前规则表：\n{rule_workbook_path}",
            font=theme.LABEL_FONT,
            wraplength=560,
            justify="left",
            anchor="w",
            text_color=theme.TEXT_MUTED,
        )
        path_label.pack(anchor="w", fill=tk.X, padx=16, pady=(0, 10))

        def start_preview():
            rule_workbook_path = result["path"]
            if not rule_workbook_path:
                self.show_info(
                    "分别替换",
                    "规则表未生成，请重新进入分别替换。",
                    dialog_width=620,
                    wraplength=560,
                )
                return
            if not os.path.exists(rule_workbook_path):
                self.show_info(
                    "分别替换",
                    "规则表不存在，请重新进入分别替换。",
                    dialog_width=620,
                    wraplength=560,
                )
                return
            done.set(True)
            dialog.destroy()

        def cancel():
            done.set(True)
            dialog.destroy()

        button_frame = self.dialogs.create_button_bar(card)
        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.dialogs.add_button(
            button_frame,
            "我已填写并保存，开始预览",
            start_preview,
            primary=True,
            width=220,
        ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=700, height=320)
        self.root.wait_variable(done)
        return result["path"]

    def _preview_word_replace(self, mode, word_paths, rules):
        if mode == MODE_COMMON:
            return preview_word_replace(
                word_paths=word_paths,
                mode=mode,
                common_rules=rules,
                logger=self.flushing_logger(),
            )
        return preview_word_replace(
            word_paths=word_paths,
            mode=mode,
            per_file_rules=rules,
            logger=self.flushing_logger(),
        )

    def _count_word_replace_rules(self, rules):
        if isinstance(rules, dict):
            return sum(len(file_rules) for file_rules in rules.values())
        return len(rules)

    def _confirm_word_empty_replacement(self):
        choice = self.ask_choice(
            "Word 批量替换",
            "规则中存在“替换为”为空的行。执行后对应内容会被清空。\n是否继续预览？",
            [
                ("继续预览", "continue"),
                ("取消", "cancel"),
            ],
            dialog_width=620,
            dialog_height=260,
            wraplength=560,
        )
        return choice == "continue"

    def _confirm_word_replace_preview_ready(self, preview_result):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("Word 批量替换预览", resizable=True)

        summary = preview_result["summary"]
        self.dialogs.add_message(
            card,
            (
                "请检查替换计划，确认无误后执行。\n"
                f"待执行规则：{summary['pending_count']} 条；"
                f"预计替换：{summary['estimated_replace_count']} 处；"
                f"跳过：{summary['skipped_count']} 条；"
                f"失败：{summary['failed_count']} 条。"
            ),
            wraplength=820,
            pady=(0, 10),
        )

        preview_text = ctk.CTkTextbox(
            card,
            font=theme.PREVIEW_FONT,
            height=300,
            fg_color=theme.SURFACE_RAISED,
            border_width=1,
            border_color=theme.BORDER,
            text_color=theme.TEXT,
            wrap="none",
        )
        preview_text.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 10))
        preview_text.insert("1.0", self._format_word_replace_preview(preview_result))
        preview_text.configure(state="disabled")

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info("正在执行 Word 批量替换。")
                execute_result = execute_word_replace(
                    preview_result["actions"],
                    logger=self.flushing_logger(),
                )
                self._log_word_replace_execution(execute_result)
                result_message = self._format_word_replace_execution_message(execute_result)
                self.show_info(
                    "Word 批量替换",
                    result_message,
                    dialog_width=760,
                    wraplength=700,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"Word 批量替换执行失败：{e}")
                self.show_info(
                    "Word 批量替换",
                    f"Word 批量替换执行失败：{e}",
                    dialog_width=680,
                    wraplength=620,
                )
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self.dialogs.add_button(
            button_frame,
            "确认执行替换",
            execute,
            primary=True,
            width=160,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=920, height=520)
        self.root.wait_variable(done)
        return result["execute"]

    def _format_word_replace_preview(self, preview_result):
        columns = [
            ("文件路径", 68),
            ("查找内容", 18),
            ("替换为", 18),
            ("预计替换次数", 12),
            ("状态", 8),
            ("说明", 28),
        ]
        lines = [
            format_display_row([title for title, _ in columns], columns),
            format_display_row(["-" * width for _, width in columns], columns),
        ]
        for item in preview_result["items"]:
            lines.append(
                format_display_row(
                    [
                        item["file_path"],
                        item["find_text"],
                        item["replace_text"],
                        str(item["estimated_count"]),
                        item["status"],
                        item["message"],
                    ],
                    columns,
                )
            )
        return "\n".join(lines)

    def _log_word_replace_preview(self, preview_result):
        summary = preview_result["summary"]
        self.logger.info(
            "Word 批量替换预览完成："
            f"文件 {summary['file_count']} 个；"
            f"待执行规则 {summary['pending_count']} 条；"
            f"预计替换 {summary['estimated_replace_count']} 处；"
            f"跳过 {summary['skipped_count']} 条；"
            f"失败 {summary['failed_count']} 条。"
        )
        self.logger.info("页眉页脚将尽量按 Word COM StoryRanges 处理；脚注、尾注、批注、文本框不作为 P0 范围。")

    def _log_word_replace_execution(self, execute_result):
        summary = execute_result["summary"]
        for file_log in execute_result["file_logs"]:
            message = (
                f"Word 批量替换日志：{file_log['file_name']}；"
                f"规则数 {file_log['rule_count']}；"
                f"替换次数 {file_log['replace_count']}；"
                f"状态 {file_log['status']}；"
                f"备份状态 {file_log['backup_status']}"
            )
            if file_log["backup_path"]:
                message += f"；备份路径 {file_log['backup_path']}"
            if file_log["message"]:
                message += f"；说明 {file_log['message']}"
            if file_log["status"] == "失败":
                self.logger.error(message)
            else:
                self.logger.info(message)
        self.logger.info(
            "Word 批量替换完成："
            f"成功文件 {summary['success_file_count']} 个；"
            f"跳过文件 {summary['skipped_file_count']} 个；"
            f"失败文件 {summary['failed_file_count']} 个；"
            f"实际替换 {summary['replace_count']} 处。"
        )

    def _format_word_replace_execution_message(self, execute_result):
        summary = execute_result["summary"]
        return (
            "Word 批量替换完成。\n"
            f"成功文件：{summary['success_file_count']} 个\n"
            f"跳过文件：{summary['skipped_file_count']} 个\n"
            f"失败文件：{summary['failed_file_count']} 个\n"
            f"实际替换：{summary['replace_count']} 处\n\n"
            "执行前已为实际处理的文件自动备份。\n"
            "详细备份路径和失败原因见下方日志。"
        )
