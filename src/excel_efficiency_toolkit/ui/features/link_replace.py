# -*- coding: utf-8 -*-
"""批量更换 Excel 外部链接：扫描链接生成规则表 → 用户填写 → 确认执行。"""
import os
import tkinter as tk
from tkinter import filedialog

from ...link_replace_ops import (
    execute_link_replacement_from_rule_workbook,
    generate_temporary_link_replace_rule_workbook,
)
from .base import FeatureController


class LinkReplaceController(FeatureController):
    def run_batch_link_replace(self):
        """按钮回调函数，扫描外部链接并按规则批量 ChangeLink"""
        self.logger.info("批量更换多文件链接：开始操作。")
        source_paths = filedialog.askopenfilenames(
            title="请选择需要批量更换多文件链接的 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_link_replace").config(state="disabled")
        try:
            result = generate_temporary_link_replace_rule_workbook(
                source_paths=list(source_paths),
                logger=self.flushing_logger(),
            )
            self.logger.info(f"选择文件数量：{result['source_file_count']}")
            self.logger.info(f"读取成功文件数：{result['read_success_count']}")
            self.logger.info(f"扫描到外部链接数量：{result['link_count']}")
            self.logger.info(f"临时规则表路径：{result['output_path']}")

            try:
                os.startfile(result["output_path"])
                self.logger.info("规则表已自动打开。请填写 D 列，保存并关闭规则表后，再点击弹窗中的执行按钮。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")
                self.logger.info(f"请手动打开规则表：{result['output_path']}")

            if not self._confirm_link_replace_rule_ready(result["output_path"]):
                self.logger.info("用户已取消操作")
        except Exception as e:
            self.logger.error(f"批量更换多文件链接失败：{e}")
            self.show_info(
                "批量更换多文件链接",
                f"批量更换多文件链接失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.button("btn_link_replace").config(state="normal")

    def _confirm_link_replace_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("批量更换多文件链接")

        self.dialogs.add_message(
            card,
            (
                "规则表已打开。\n"
                "请在【链接清单】D列填写新链接路径，E列可填写是否执行，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=420,
            pady=(0, 12),
        )

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info(f"正在执行批量更换多文件链接规则表：{rule_workbook_path}")
                try:
                    summary = execute_link_replacement_from_rule_workbook(
                        rule_workbook_path,
                        logger=self.flushing_logger(),
                    )
                except (PermissionError, OSError) as e:
                    self.logger.error(f"规则表无法读取或回写：{e}")
                    self._show_link_replace_rule_workbook_busy_message()
                    return

                result_message = (
                    "批量更换多文件链接完成。\n"
                    f"规则行数：{summary['rule_count']}\n"
                    f"成功：{summary['success_count']} 个\n"
                    f"跳过：{summary['skipped_count']} 个\n"
                    f"失败：{summary['failed_count']} 个\n\n"
                    "规则表已更新状态和处理日志。\n"
                    "成功处理的目标工作簿已按参数保存。"
                )
                self.logger.info(
                    "批量更换多文件链接完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self.show_info(
                    "批量更换多文件链接",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量更换多文件链接失败：{e}")
                self.show_info(
                    "批量更换多文件链接",
                    f"批量更换多文件链接失败：{e}",
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
            "我已填好规则，执行批量更换多文件链接",
            execute,
            primary=True,
            width=320,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=550, height=280)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_link_replace_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self.show_info(
            "批量更换多文件链接",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行批量更换多文件链接。",
            dialog_width=620,
            wraplength=560,
        )
