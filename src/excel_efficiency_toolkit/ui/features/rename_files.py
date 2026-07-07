# -*- coding: utf-8 -*-
"""批量重命名文件：生成规则表 → 用户填写 → 确认执行。"""
import os
import tkinter as tk
from tkinter import filedialog

from ...rename_file_ops import (
    build_rename_plan,
    create_rename_rule_workbook,
    execute_rename_plan,
    read_rename_rules,
    read_rename_settings,
    write_rename_results_to_workbook,
)
from .base import FeatureController


class RenameFilesController(FeatureController):
    def run_batch_rename_files(self):
        """按钮回调函数，生成临时规则表后确认执行批量重命名文件"""
        self.logger.info("批量重命名文件：开始操作。")
        source_paths = filedialog.askopenfilenames(
            title="请选择需要批量重命名的文件",
            filetypes=[("所有文件", "*.*")],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_rename_files").config(state="disabled")
        try:
            self.logger.info(f"已选择文件数量：{len(source_paths)}")
            rule_path = create_rename_rule_workbook(list(source_paths))
            self.logger.info(f"已生成重命名规则表：{rule_path}")
            try:
                os.startfile(rule_path)
                self.logger.info("规则表已打开，请填写 C/D 列，保存并关闭规则表后点击执行。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")
                self.logger.info(f"请手动打开规则表：{rule_path}")

            if not self._confirm_rename_rule_ready(rule_path):
                self.logger.info("用户已取消操作")
        except Exception as e:
            self.logger.error(f"批量重命名文件失败：{e}")
            self.show_info(
                "批量重命名文件",
                f"批量重命名文件失败：{e}",
            )
        finally:
            self.button("btn_rename_files").config(state="normal")

    def _confirm_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("批量重命名文件")

        self.dialogs.add_message(
            card,
            (
                "规则表已打开。\n"
                "请在【重命名清单】中填写新文件名和后缀名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=400,
            pady=(0, 12),
        )

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info(f"正在读取重命名规则表：{rule_workbook_path}")
                try:
                    write_rename_results_to_workbook(rule_workbook_path, [])
                except PermissionError:
                    self._show_rule_workbook_busy_message()
                    return
                except OSError as e:
                    self.logger.error(f"规则表无法回写：{e}")
                    self._show_rule_workbook_busy_message()
                    return

                settings = read_rename_settings(rule_workbook_path)
                for warning in settings.warnings or []:
                    self.logger.info(warning)

                rules = read_rename_rules(rule_workbook_path)
                actions = build_rename_plan(rules, settings)
                valid_actions = [action for action in actions if action.status == "成功"]
                skipped_actions = [action for action in actions if action.status == "跳过"]
                self.logger.info(f"规则行数：{len(rules)}")
                self.logger.info(f"有效重命名任务：{len(valid_actions)}")
                self.logger.info(f"跳过任务：{len(skipped_actions)}")

                for index, action in enumerate(valid_actions, start=1):
                    self.logger.info(
                        f"正在重命名 {index}/{len(valid_actions)}："
                        f"{os.path.basename(action.original_path)} -> {action.final_name}"
                    )

                summary = execute_rename_plan(actions)
                try:
                    write_rename_results_to_workbook(rule_workbook_path, actions)
                    self.logger.info(f"规则表已更新：{rule_workbook_path}")
                    result_message = (
                        "重命名完成。\n"
                        f"成功：{summary['success_count']} 个\n"
                        f"跳过：{summary['skipped_count']} 个\n"
                        f"失败：{summary['failed_count']} 个\n\n"
                        "规则表已更新状态/备注，请查看确认。"
                    )
                except PermissionError:
                    self.logger.error("规则表被占用，无法回写状态、说明和处理日志。")
                    result_message = (
                        "重命名已执行，但规则表无法回写。\n"
                        "请确认规则表已保存并关闭后，再查看文件结果。"
                    )
                except OSError as e:
                    self.logger.error(f"规则表无法回写：{e}")
                    result_message = (
                        "重命名已执行，但规则表无法回写。\n"
                        f"原因：{e}"
                    )

                self.logger.info(
                    "批量重命名完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self.show_info(
                    "批量重命名文件",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名文件失败：{e}")
                self.show_info(
                    "批量重命名文件",
                    f"批量重命名文件失败：{e}",
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
            "我已填好规则，执行批量重命名",
            execute,
            primary=True,
            width=260,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=510, height=255)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self.show_info(
            "批量重命名文件",
            "规则表仍被 Excel 占用，无法回写。\n"
            "请先保存并关闭规则表，再点击执行批量重命名。",
            dialog_width=620,
            wraplength=560,
        )
