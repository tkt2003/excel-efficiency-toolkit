# -*- coding: utf-8 -*-
"""颜色工具：按颜色汇总求和、按颜色清空内容。"""
import tkinter as tk
from tkinter import filedialog

import customtkinter as ctk

from ...clear_by_color_ops import (
    clear_multiple_workbooks_by_color,
    execute_clear_active_workbook_plan,
    plan_clear_active_workbook_by_color,
)
from ...color_sum_ops import sum_current_sheet_by_fill_color, sum_matching_sheets_by_fill_color
from .. import theme
from .base import FeatureController


class ColorToolsController(FeatureController):
    # ------------------------------------------------------------------
    # 按颜色汇总求和
    # ------------------------------------------------------------------
    def run_color_sum(self):
        """按钮回调函数，按当前选中单元格填充色汇总指定 sheet 的同地址单元格"""
        self.log_info("按颜色汇总求和：开始操作。")
        try:
            sum_scope = self.ask_choice(
                "按颜色汇总求和",
                "请选择汇总范围：\n1 仅汇总一个 sheet\n2 汇总所有匹配 sheet",
                [
                    ("仅汇总一个 sheet", "single"),
                    ("汇总所有匹配 sheet", "all"),
                ],
                dialog_width=510,
                dialog_height=255,
                wraplength=460,
            )
            if sum_scope is None:
                self.log_info("用户已取消操作。")
                return

            target_sheet_name = None
            if sum_scope == "single":
                target_sheet_name = self.ask_text(
                    "按颜色汇总求和",
                    "请输入要汇总的 sheet 名；留空则使用当前活动 sheet。",
                    entry_width=26,
                    dialog_width=520,
                    wraplength=460,
                )
                if target_sheet_name is None:
                    self.log_info("用户已取消操作。")
                    return

            source_paths = filedialog.askopenfilenames(
                title="请选择源 Excel 文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                    ("所有文件", "*.*"),
                ],
            )
            if not source_paths:
                self.log_info("用户已取消操作。")
                return

            write_mode = self.ask_choice(
                "按颜色汇总求和",
                "请选择写入方式：\n1 写入求和公式\n2 只写入汇总数值",
                [
                    ("写入求和公式", "formula"),
                    ("只写入汇总数值", "value"),
                ],
                dialog_width=510,
                dialog_height=255,
                wraplength=460,
            )
            if write_mode is None:
                self.log_info("用户已取消操作。")
                return

            self.button("btn_color_sum").config(state="disabled")
            if sum_scope == "single":
                result = sum_current_sheet_by_fill_color(
                    source_paths=list(source_paths),
                    write_mode=write_mode,
                    target_sheet_name=target_sheet_name,
                    logger=self.flushing_logger(),
                )
                self._log_single_color_sum_result(result)
                self._show_single_color_sum_result(result)
            else:
                result = sum_matching_sheets_by_fill_color(
                    source_paths=list(source_paths),
                    write_mode=write_mode,
                    logger=self.flushing_logger(),
                )
                self._log_all_color_sum_result(result)
                self._show_all_color_sum_result(result)
        except Exception as e:
            self.log_error(f"按颜色汇总求和失败：{type(e).__name__}: {e}")
        finally:
            self.button("btn_color_sum").config(state="normal")

    def _log_single_color_sum_result(self, result):
        self.log_info(
            "按颜色汇总求和完成："
            f"写入 {result['written_cell_count']} 个单元格；"
            f"找到同名工作表源文件 {result['matched_source_file_count']} 个；"
            f"缺少同名工作表 {result['missing_sheet_file_count']} 个；"
            f"忽略非数字 {result['ignored_non_numeric_count']} 个。"
        )
        self.log_info("目标工作簿未自动保存，请检查后自行保存。")

    def _show_single_color_sum_result(self, result):
        self.show_info(
            "按颜色汇总求和",
            "汇总完成。\n"
            f"目标工作表：{result['target_sheet_name']}\n"
            f"写入单元格：{result['written_cell_count']}\n"
            f"参与源文件：{result['matched_source_file_count']}\n"
            f"缺少同名工作表：{result['missing_sheet_file_count']}\n"
            f"忽略非数字：{result['ignored_non_numeric_count']}\n\n"
            "目标工作簿未自动保存，请检查后自行保存。",
            dialog_width=720,
            wraplength=650,
        )

    def _log_all_color_sum_result(self, result):
        self.log_info(
            "按颜色汇总求和完成："
            f"目标计划匹配 {result['matched_sheet_count']} 个工作表；"
            f"实际写入 {result['written_sheet_count']} 个工作表、{result['written_cell_count']} 个单元格；"
            f"参与源文件 {result['matched_source_file_count']} 个；"
            f"匹配源工作表 {result['matched_source_sheet_count']} 个；"
            f"缺少同名工作表 {result['missing_sheet_file_count']} 个；"
            f"忽略非数字 {result['ignored_non_numeric_count']} 个。"
        )
        self.log_info("目标工作簿未自动保存，请检查后自行保存。")

    def _show_all_color_sum_result(self, result):
        self.show_info(
            "按颜色汇总求和",
            "汇总完成。\n"
            f"目标计划匹配工作表：{result['matched_sheet_count']}\n"
            f"实际写入工作表：{result['written_sheet_count']}\n"
            f"写入单元格：{result['written_cell_count']}\n"
            f"参与源文件：{result['matched_source_file_count']}\n"
            f"匹配源工作表：{result['matched_source_sheet_count']}\n"
            f"缺少同名工作表：{result['missing_sheet_file_count']}\n"
            f"忽略非数字：{result['ignored_non_numeric_count']}\n\n"
            "目标工作簿未自动保存，请检查后自行保存。",
            dialog_width=720,
            wraplength=650,
        )

    # ------------------------------------------------------------------
    # 按颜色清空内容
    # ------------------------------------------------------------------
    def run_clear_by_color(self):
        self.log_info("按颜色清空内容：开始操作。")
        try:
            mode = self.ask_choice(
                "按颜色清空内容",
                "请选择清空范围：\n\n"
                "当前活动工作簿\n"
                "清空当前打开的活动工作簿中所有可见 sheet 里，同颜色单元格的内容。\n\n"
                "多个工作簿\n"
                "批量清空所选工作簿中所有可见 sheet 里，同颜色单元格的内容。",
                [
                    ("当前活动工作簿", "active"),
                    ("多个工作簿", "multi"),
                ],
                dialog_width=560,
                dialog_height=310,
                wraplength=500,
            )
            if mode is None:
                self.log_info("用户已取消操作。")
                return

            self.button("btn_clear_by_color").config(state="disabled")
            if mode == "active":
                self._run_clear_current_workbook_by_color()
            else:
                self._run_clear_multiple_workbooks_by_color()
        except Exception as e:
            self.log_error(f"按颜色清空内容失败：{type(e).__name__}: {e}")
            self.show_info(
                "按颜色清空内容",
                f"按颜色清空内容失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.button("btn_clear_by_color").config(state="normal")

    def _run_clear_current_workbook_by_color(self):
        plan = plan_clear_active_workbook_by_color(logger=self.flushing_logger())
        if plan["matched_cell_count"] == 0:
            self.log_info("未找到匹配单元格，已取消清空。")
            self.show_info(
                "按颜色清空内容",
                "未找到匹配单元格，不执行清空。",
            )
            return

        result = execute_clear_active_workbook_plan(plan, logger=self.flushing_logger())
        self.show_info(
            "按颜色清空内容",
            "清空完成。\n"
            f"当前颜色：{result['current_color_text']}\n"
            f"匹配工作表：{result['matched_sheet_count']}\n"
            f"清空工作表：{result['cleared_sheet_count']}\n"
            f"清空单元格：{result['cleared_cell_count']}\n"
            f"批量区域：{result['range_group_count']} 组\n\n"
            "当前工作簿未自动保存，请检查后自行保存。",
            dialog_width=520,
            wraplength=460,
        )

    def _run_clear_multiple_workbooks_by_color(self):
        target_paths = filedialog.askopenfilenames(
            title="请选择需要按颜色清空内容的 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not target_paths:
            self.log_info("用户已取消操作。")
            return

        skip_backup = self._ask_clear_multi_backup_option_no_grab()
        if skip_backup is None:
            self.log_info("用户已取消操作。")
            return

        result = clear_multiple_workbooks_by_color(
            list(target_paths),
            skip_backup=skip_backup,
            logger=self.flushing_logger(),
        )
        backup_message = (
            f"批次备份目录：{result['batch_backup_dir']}"
            if result["batch_backup_dir"]
            else "备份方式：用户选择跳过备份"
            if skip_backup
            else "备份方式：无需备份"
        )
        self.show_info(
            "按颜色清空内容",
            "批量清空完成。\n"
            f"当前颜色：{result['current_color_text']}\n"
            f"成功处理文件：{result['processed_file_count']}\n"
            f"跳过文件：{result['skipped_file_count']}\n"
            f"失败文件：{result['failed_file_count']}\n"
            f"清空单元格：{result['cleared_cell_count']}\n"
            f"{backup_message}\n"
            f"处理日志：{result['log_path']}",
            dialog_width=720,
            wraplength=650,
        )

    def _ask_clear_multi_backup_option_no_grab(self):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self.dialogs.create_dialog_card("按颜色清空内容")

        self.dialogs.add_message(
            card,
            (
                "多个工作簿模式默认会先扫描，再把实际会被修改的文件集中备份到一个批次备份文件夹。\n"
                "无匹配、跳过或失败且未修改的文件不会备份。"
            ),
            wraplength=480,
            pady=(0, 12),
        )

        skip_backup_var = tk.BooleanVar(master=self.root, value=False)
        ctk.CTkCheckBox(
            card,
            text="已自行备份，本次不再生成备份文件",
            variable=skip_backup_var,
            font=theme.SMALL_FONT,
            text_color=theme.TEXT,
            fg_color=theme.ACCENT,
            hover_color=theme.ACCENT_HOVER,
            border_color=theme.BORDER,
        ).pack(anchor="w", fill=tk.X, padx=18, pady=(0, 12))

        button_frame = self.dialogs.create_button_bar(card)

        def confirm():
            result["value"] = skip_backup_var.get()
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self.dialogs.add_button(button_frame, "开始执行", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda event: cancel())
        self.dialogs.show_no_grab(dialog, width=560, height=300)
        self.dialogs.schedule_smoke_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["value"]
