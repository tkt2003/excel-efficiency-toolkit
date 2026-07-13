# -*- coding: utf-8 -*-
"""批量重命名工作表：读取工作表清单生成规则表 → 用户填写 → 确认执行。"""
import os
import tkinter as tk
from tkinter import filedialog

from ...rename_sheet_ops import (
    build_skipped_sheet_rename_actions,
    build_sheet_rename_plan,
    create_sheet_rename_rule_workbook,
    execute_sheet_rename_plan,
    group_sheet_rename_rules_by_workbook_path,
    is_excel_workbook_file,
    is_office_temp_file,
    read_sheet_infos_from_workbook_files,
    read_sheet_rename_rules,
    read_sheet_rename_settings,
    summarize_sheet_rename_actions,
    write_sheet_rename_results_to_workbook,
)
from .base import FeatureController


class RenameSheetsController(FeatureController):
    def run_batch_rename_sheets(self):
        """按钮回调函数，生成临时规则表后确认批量重命名工作表"""
        self.logger.info("批量重命名工作表：开始操作。")
        source_paths = filedialog.askopenfilenames(
            title="请选择需要批量重命名工作表的 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_rename_sheets").config(state="disabled")
        try:
            self.logger.info(f"已选择文件数量：{len(source_paths)}")
            workbook_infos = read_sheet_infos_from_workbook_files(list(source_paths), logger=self.logger)
            readable_workbook_count = len([info for info in workbook_infos if not info.get("error")])
            sheet_count = sum(len(info.get("sheet_infos", [])) for info in workbook_infos)
            self.logger.info(f"读取成功工作簿数量：{readable_workbook_count}")
            self.logger.info(f"读取工作表数量：{sheet_count}")

            rule_path = create_sheet_rename_rule_workbook(workbook_infos=workbook_infos)
            self.logger.info(f"已生成工作表重命名规则表：{rule_path}")
            try:
                os.startfile(rule_path)
                self.logger.info("规则表已打开，请填写 D 列，保存并关闭规则表后点击执行。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")
                self.logger.info(f"请手动打开规则表：{rule_path}")

            if not self._confirm_sheet_rename_rule_ready(rule_path):
                self.logger.info("用户已取消操作")
        except Exception as e:
            self.logger.error(f"批量重命名工作表失败：{e}")
            self.show_info(
                "批量重命名工作表",
                f"批量重命名工作表失败：{e}",
            )
        finally:
            self.button("btn_rename_sheets").config(state="normal")

    def _confirm_sheet_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("批量重命名工作表")

        self.dialogs.add_message(
            card,
            (
                "规则表已打开。\n"
                "请在【重命名清单】D列填写新工作表名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=400,
            pady=(0, 12),
        )

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info(f"正在读取工作表重命名规则表：{rule_workbook_path}")
                try:
                    write_sheet_rename_results_to_workbook(rule_workbook_path, [])
                except (PermissionError, OSError) as e:
                    self.logger.error(f"规则表无法读取或回写：{e}")
                    self._show_sheet_rule_workbook_busy_message()
                    return

                try:
                    settings = read_sheet_rename_settings(rule_workbook_path)
                    rules = read_sheet_rename_rules(rule_workbook_path)
                except (PermissionError, OSError) as e:
                    self.logger.error(f"规则表无法读取：{e}")
                    self._show_sheet_rule_workbook_busy_message()
                    return

                for warning in settings.warnings or []:
                    self.logger.info(warning)

                self.logger.info(f"规则行数：{len(rules)}")
                actions = self._execute_sheet_rename_rules_by_workbook(rules, settings)
                summary = summarize_sheet_rename_actions(actions)
                try:
                    write_sheet_rename_results_to_workbook(rule_workbook_path, actions)
                    self.logger.info(f"规则表已更新：{rule_workbook_path}")
                    result_message = (
                        "重命名完成。\n"
                        f"成功：{summary['success_count']} 个\n"
                        f"跳过：{summary['skipped_count']} 个\n"
                        f"失败：{summary['failed_count']} 个\n\n"
                        "规则表已更新状态/备注，请查看确认。\n"
                        "已保存成功处理的目标工作簿。"
                    )
                except (PermissionError, OSError) as e:
                    self.logger.error(f"规则表无法回写：{e}")
                    result_message = (
                        "重命名已执行，但规则表无法回写。\n"
                        "请确认规则表已保存并关闭后，再查看工作簿结果。"
                    )

                self.logger.info(
                    "批量重命名工作表完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self.show_info(
                    "批量重命名工作表",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名工作表失败：{e}")
                self.show_info(
                    "批量重命名工作表",
                    f"批量重命名工作表失败：{e}",
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
            "我已填好规则，执行批量重命名工作表",
            execute,
            primary=True,
            width=320,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=510, height=255)
        self.root.wait_variable(done)
        return result["execute"]

    def _execute_sheet_rename_rules_by_workbook(self, rules, settings):
        import pythoncom
        import win32com.client

        grouped_rules = group_sheet_rename_rules_by_workbook_path(rules)
        all_actions = []
        pythoncom.CoInitialize()
        excel = None
        workbook = None

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for rules_for_workbook in grouped_rules.values():
                workbook = None
                workbook_path = rules_for_workbook[0].workbook_path
                if settings.skip_temp_files and is_office_temp_file(workbook_path):
                    all_actions.extend(build_skipped_sheet_rename_actions(rules_for_workbook, "临时文件已跳过"))
                    continue
                if not os.path.exists(workbook_path):
                    all_actions.extend(build_skipped_sheet_rename_actions(rules_for_workbook, "原文件不存在"))
                    continue
                if not is_excel_workbook_file(workbook_path):
                    all_actions.extend(build_skipped_sheet_rename_actions(rules_for_workbook, "不是支持的 Excel 文件"))
                    continue

                try:
                    self.logger.info(f"正在打开目标工作簿：{workbook_path}")
                    workbook = excel.Workbooks.Open(
                        os.path.abspath(workbook_path),
                        UpdateLinks=0,
                        ReadOnly=False,
                    )
                    existing_sheet_names = [sheet.Name for sheet in workbook.Worksheets]
                    actions = build_sheet_rename_plan(rules_for_workbook, existing_sheet_names, settings)
                    valid_actions = [action for action in actions if action.status == "成功"]
                    skipped_actions = [action for action in actions if action.status == "跳过"]
                    self.logger.info(
                        f"{os.path.basename(workbook_path)}："
                        f"有效重命名任务 {len(valid_actions)} 个；跳过任务 {len(skipped_actions)} 个。"
                    )

                    for index, action in enumerate(valid_actions, start=1):
                        self.logger.info(
                            f"正在重命名 {index}/{len(valid_actions)}："
                            f"{action.original_sheet_name} -> {action.target_sheet_name}"
                        )

                    execute_sheet_rename_plan(workbook, actions)
                    if any(action.status == "成功" for action in actions):
                        try:
                            workbook.Save()
                            self.logger.info(f"目标工作簿已保存：{workbook_path}")
                        except Exception as save_error:
                            for action in actions:
                                if action.status == "成功":
                                    action.status = "失败"
                                    action.message = f"保存工作簿失败：{save_error}"
                            self.logger.error(f"保存工作簿失败：{workbook_path}。详细信息：{save_error}")

                    all_actions.extend(actions)
                except Exception as e:
                    self.logger.error(f"处理工作簿失败：{workbook_path}。详细信息：{e}")
                    all_actions.extend(build_skipped_sheet_rename_actions(rules_for_workbook, f"文件打开或处理失败：{e}"))
                finally:
                    if workbook is not None:
                        try:
                            workbook.Close(SaveChanges=False)
                        except Exception:
                            pass
                        workbook = None

            return all_actions

        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def _show_sheet_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self.show_info(
            "批量重命名工作表",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行批量重命名工作表。",
            dialog_width=620,
            wraplength=560,
        )
