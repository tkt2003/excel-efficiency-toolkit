# -*- coding: utf-8 -*-
"""按模板批量生成 Excel：替换一个链接 / 替换多个链接两种模式。"""
import os
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

from ...template_multi_link_ops import (
    create_template_multi_link_rule_workbook,
    execute_template_multi_link_generation_from_rule_workbook,
)
from ...template_tb_report_ops import (
    DEFAULT_OUTPUT_NAME_SUFFIX,
    generate_reports_from_template_and_tb_files,
    read_template_external_links,
)
from .. import theme
from .base import FeatureController


class TemplatesController(FeatureController):
    def run_template_generate(self):
        """按钮回调函数，先在弹窗中让用户选择生成模式，再调用对应实现。"""
        self.log_info("按模板批量生成 Excel：开始选择生成模式。")
        mode = self.ask_choice(
            "按模板批量生成 Excel",
            "请选择生成方式：\n\n"
            "替换一个链接生成\n"
            "适合：模板里只需要更换一个 TB / 附注 / 底稿链接。\n\n"
            "替换多个链接生成\n"
            "适合：模板里有合并 TB、单体 TB 等多个链接，需要用规则表逐项替换。",
            [
                ("替换一个链接生成", "single"),
                ("替换多个链接生成", "multi"),
            ],
            dialog_width=575,
            dialog_height=310,
            wraplength=520,
        )
        if mode is None:
            self.log_info("用户已取消操作。")
            return
        if mode == "single":
            self.log_info("已选择：替换一个链接生成。")
            self.run_template_tb_report()
            return
        if mode == "multi":
            self.log_info("已选择：替换多个链接生成。")
            self.run_template_multi_link()

    def run_template_tb_report(self):
        """扫描模板外部链接、多选 TB 文件、按 TB 生成多份报表（替换一个链接生成）。"""
        self.log_info("按模板批量生成 Excel（替换一个链接生成）：开始操作。")
        template_path = filedialog.askopenfilename(
            title="请选择模板 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not template_path:
            self.log_info("用户已取消操作。")
            return

        self.button("btn_template_tb_report").config(state="disabled")
        try:
            try:
                self.log_info("正在打开模板并扫描外部链接，请稍候……")
                template_links = read_template_external_links(template_path)
            except Exception as e:
                self.log_error(f"扫描模板外部链接失败：{e}")
                self.show_info(
                    "按模板批量生成 Excel",
                    f"扫描模板外部链接失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            if not template_links:
                self.log_error("模板没有外部链接，无法执行换链接生成报表。")
                self.show_info(
                    "按模板批量生成 Excel",
                    "模板没有外部链接，无法执行换链接生成报表。\n请确认模板含有引用其他工作簿的公式后重试。",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            self.log_info(f"模板外部链接数量：{len(template_links)}")
            for link_path in template_links:
                self.log_info(f"  - {link_path}")

            if len(template_links) == 1:
                old_link_path = template_links[0]
                self.log_info(f"模板只有 1 个外部链接，已自动选择：{old_link_path}")
            else:
                old_link_path = self._ask_old_link_choice_no_grab(
                    template_links,
                    title="选择要替换的旧链接",
                    prompt="模板存在多个外部链接，请选择要替换的旧链接：",
                )
                if not old_link_path:
                    self.log_info("用户已取消操作。")
                    return
                self.log_info(f"用户选择要替换的旧链接：{old_link_path}")

            tb_paths = filedialog.askopenfilenames(
                title="请多选 TB 文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*"),
                ],
            )
            if not tb_paths:
                self.log_info("用户已取消操作。")
                return

            output_dir = filedialog.askdirectory(title="请选择输出目录")
            if not output_dir:
                self.log_info("用户已取消操作。")
                return

            suffix_input = self.ask_text(
                "按模板批量生成 Excel",
                "请输入输出文件名后缀，留空则直接使用 TB 文件名主体。\n"
                "常用：_批量生成、_附注、_财务报表、_底稿、_报表",
                default=DEFAULT_OUTPUT_NAME_SUFFIX,
                entry_width=24,
                dialog_width=560,
                wraplength=500,
            )
            if suffix_input is None:
                self.log_info("用户已取消操作。")
                return
            output_name_suffix = suffix_input.strip()

            confirmed = self.ask_choice(
                "按模板批量生成 Excel",
                "请确认生成口径：\n"
                f"模板路径：{template_path}\n"
                f"被替换的旧链接：{old_link_path}\n"
                f"TB 文件数量：{len(tb_paths)}\n"
                f"输出目录：{output_dir}\n"
                f"输出后缀：{output_name_suffix or '（留空：直接使用 TB 文件名主体）'}\n\n"
                "程序会逐个 TB 复制完整模板，并将选中的旧链接替换为该 TB 路径；\n"
                "其他外部链接保持不变；原模板不会被修改。",
                [("开始生成", "run")],
                dialog_width=820,
                wraplength=740,
            )
            if confirmed != "run":
                self.log_info("用户已取消操作。")
                return

            self.log_info(f"TB 文件数量：{len(tb_paths)}")
            self.log_info(f"输出目录：{output_dir}")
            self.log_info(f"输出后缀：{output_name_suffix or '（留空）'}")
            result = generate_reports_from_template_and_tb_files(
                template_path=template_path,
                tb_paths=list(tb_paths),
                output_dir=output_dir,
                old_link_path=old_link_path,
                output_name_suffix=output_name_suffix,
                logger=self.flushing_logger(),
            )
            for record in result["records"]:
                self.log_info(
                    f"第 {record.index} 个：{record.tb_name} -> {record.output_name or '-'}；"
                    f"{record.status}；{record.message}"
                )
            self.log_info(
                "按模板批量生成 Excel完成："
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个；"
                f"失败 {result['failed_count']} 个；"
                f"日志：{result['log_path']}"
            )
            self.show_info(
                "按模板批量生成 Excel",
                "生成完成。\n"
                f"TB 文件数量：{result['tb_file_count']}\n"
                f"成功数量：{result['success_count']}\n"
                f"跳过数量：{result['skipped_count']}\n"
                f"失败数量：{result['failed_count']}\n"
                f"输出目录：{result['output_dir']}\n"
                f"日志文件：{result['log_path']}\n\n"
                "原模板工作簿未被修改。",
                dialog_width=720,
                wraplength=650,
            )
        except Exception as e:
            self.log_error(f"按模板批量生成 Excel失败：{type(e).__name__}: {e}")
            self.show_info(
                "按模板批量生成 Excel",
                f"按模板批量生成 Excel失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.button("btn_template_tb_report").config(state="normal")

    def _ask_old_link_choice_no_grab(self, link_paths, title, prompt):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card(title)

        self.dialogs.add_message(card, prompt, wraplength=520, pady=(0, 12))

        selected_index = tk.IntVar(master=dialog, value=0)
        radio_frame = ctk.CTkFrame(
            card,
            fg_color=theme.SURFACE_RAISED,
            corner_radius=8,
            border_width=1,
            border_color=theme.BORDER,
        )
        radio_frame.pack(fill=tk.X, padx=18, pady=(0, 12))
        for index, link_path in enumerate(link_paths):
            row_frame = ctk.CTkFrame(radio_frame, fg_color="transparent")
            row_frame.pack(fill=tk.X, padx=12, pady=5)
            row_frame.grid_columnconfigure(1, weight=1)
            ctk.CTkRadioButton(
                row_frame,
                text="",
                variable=selected_index,
                value=index,
                width=24,
                fg_color=theme.ACCENT,
                hover_color=theme.ACCENT_HOVER,
                border_color=theme.BORDER,
            ).grid(row=0, column=0, sticky="n", pady=1)
            link_label = ctk.CTkLabel(
                row_frame,
                text=link_path,
                font=theme.SMALL_FONT,
                wraplength=520,
                justify="left",
                text_color=theme.TEXT,
            )
            link_label.grid(row=0, column=1, sticky="w", padx=(8, 0))
            link_label.bind("<Button-1>", lambda event, selected=index: selected_index.set(selected))

        button_frame = self.dialogs.create_button_bar(card)

        def confirm():
            try:
                result["value"] = link_paths[selected_index.get()]
            except (IndexError, tk.TclError):
                result["value"] = None
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.dialogs.add_button(button_frame, "确定", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda event: cancel())
        self.dialogs.show_no_grab(dialog, width=600, height=340)
        self.root.wait_variable(done)
        return result["value"]

    def run_template_multi_link(self):
        """扫描模板外部链接、生成规则表后批量按多链接替换（替换多个链接生成）。"""
        self.log_info("按模板批量生成 Excel（替换多个链接生成）：开始操作。")
        template_path = filedialog.askopenfilename(
            title="请选择模板 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not template_path:
            self.log_info("用户已取消操作。")
            return

        self.button("btn_template_tb_report").config(state="disabled")
        try:
            try:
                self.log_info("正在打开模板并扫描外部链接，请稍候……")
                template_links = read_template_external_links(template_path)
            except Exception as e:
                self.log_error(f"扫描模板外部链接失败：{e}")
                self.show_info(
                    "按模板多链接生成 Excel",
                    f"扫描模板外部链接失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            if not template_links:
                self.log_error("模板没有外部链接，无法生成多链接规则表。")
                self.show_info(
                    "按模板多链接生成 Excel",
                    "模板没有外部链接，无法生成多链接规则表。\n请确认模板含有引用其他工作簿的公式后重试。",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            self.log_info(f"模板外部链接数量：{len(template_links)}")
            for link_path in template_links:
                self.log_info(f"  - {link_path}")

            if len(template_links) == 1:
                main_old_link = template_links[0]
                self.log_info(f"模板只有 1 个外部链接，已作为主链接：{main_old_link}")
            else:
                main_old_link = self._ask_old_link_choice_no_grab(
                    template_links,
                    title="选择主链接",
                    prompt=(
                        "模板存在多个外部链接，请选择主链接。\n\n"
                        "主链接对应的新源文件数量，将决定生成多少份 Excel。\n"
                        "其他链接可在后续规则表中按行填写；新链接留空则保持原链接不变。"
                    ),
                )
                if not main_old_link:
                    self.log_info("用户已取消操作。")
                    return
                self.log_info(f"用户选择主链接：{main_old_link}")

            main_source_paths = filedialog.askopenfilenames(
                title="请多选主链接对应的新源文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*"),
                ],
            )
            if not main_source_paths:
                self.log_info("用户已取消操作。")
                return

            suffix_input = self.ask_text(
                "按模板多链接生成 Excel",
                "请输入输出文件名后缀，留空则直接使用主源文件名主体。\n"
                "常用：_批量生成、_附注、_财务报表、_底稿、_报表",
                default=DEFAULT_OUTPUT_NAME_SUFFIX,
                entry_width=24,
                dialog_width=560,
                wraplength=500,
            )
            if suffix_input is None:
                self.log_info("用户已取消操作。")
                return
            output_name_suffix = suffix_input.strip()

            output_dir = filedialog.askdirectory(title="请选择输出目录")
            if not output_dir:
                self.log_info("用户已取消操作。")
                return

            self.log_info(f"主源文件数量：{len(main_source_paths)}")
            self.log_info(f"输出目录：{output_dir}")
            self.log_info(f"输出后缀：{output_name_suffix or '（留空）'}")
            rule_path = create_template_multi_link_rule_workbook(
                template_path=template_path,
                output_dir=output_dir,
                old_links=list(template_links),
                main_old_link=main_old_link,
                main_source_paths=list(main_source_paths),
                output_name_suffix=output_name_suffix,
            )
            self.log_info(f"多链接生成规则表路径：{rule_path}")

            try:
                os.startfile(rule_path)
                self.log_info("规则表已自动打开。请检查并填写其他新链接列，保存并关闭规则表后点击执行。")
            except Exception as open_error:
                self.log_error(f"规则表已生成，但自动打开失败：{open_error}")
                self.log_info(f"请手动打开规则表：{rule_path}")

            if not self._confirm_template_multi_link_rule_ready(rule_path):
                self.log_info("用户已取消操作。")
        except Exception as e:
            self.log_error(f"按模板多链接生成 Excel失败：{type(e).__name__}: {e}")
            self.show_info(
                "按模板多链接生成 Excel",
                f"按模板多链接生成 Excel失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.button("btn_template_tb_report").config(state="normal")

    def _confirm_template_multi_link_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("按模板多链接生成 Excel")

        self.dialogs.add_message(
            card,
            (
                "规则表已打开。\n"
                "请检查【生成清单】，按需填写其他旧链接对应的新链接列；\n"
                "新链接留空表示不替换该旧链接。\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=460,
            pady=(0, 12),
        )

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.log_info(f"正在执行多链接生成规则表：{rule_workbook_path}")
                try:
                    summary = execute_template_multi_link_generation_from_rule_workbook(
                        rule_workbook_path,
                        logger=self.flushing_logger(),
                    )
                except (PermissionError, OSError) as e:
                    self.log_error(f"规则表无法读取或回写：{e}")
                    self._show_template_multi_link_rule_workbook_busy_message()
                    return

                result_message = (
                    "按模板多链接生成 Excel 完成。\n"
                    f"规则行数：{summary['rule_count']}\n"
                    f"成功：{summary['success_count']} 个\n"
                    f"跳过：{summary['skipped_count']} 个\n"
                    f"失败：{summary['failed_count']} 个\n\n"
                    "规则表已更新状态和处理日志。\n"
                    "原模板工作簿未被修改。"
                )
                self.log_info(
                    "按模板多链接生成 Excel完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self.show_info(
                    "按模板多链接生成 Excel",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.log_error(f"按模板多链接生成 Excel失败：{e}")
                self.show_info(
                    "按模板多链接生成 Excel",
                    f"按模板多链接生成 Excel失败：{e}",
                    dialog_width=620,
                    wraplength=560,
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
            "我已检查规则，开始生成",
            execute,
            primary=True,
            width=210,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=550, height=280)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_template_multi_link_rule_workbook_busy_message(self):
        self.log_error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self.show_info(
            "按模板多链接生成 Excel",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行。",
            dialog_width=620,
            wraplength=560,
        )
