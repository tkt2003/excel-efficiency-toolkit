# -*- coding: utf-8 -*-
"""合并整理：多表合并到一张表、多个工作簿合并到一个文件。"""
from tkinter import filedialog

from ...table_ops import merge_workbook_sheets_to_new_sheet, validate_row_numbers
from ...workbook_merge_ops import merge_workbooks_to_existing_workbook
from .base import FeatureController


class MergeController(FeatureController):
    def run_merge_sheets(self):
        """按钮回调函数，将所选工作簿中的可见工作表合并到一个新工作表"""
        source_path = filedialog.askopenfilename(
            title="请选择源 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_path:
            self.logger.info("用户已取消操作")
            return

        result_sheet_name = self.ask_text(
            "多表合并",
            "结果 sheet 名：",
            default="合并结果",
        )
        if result_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        try:
            header_row = self.ask_positive_int("多表合并", "表头行号：", 1)
            if header_row is None:
                self.logger.info("用户已取消操作")
                return
            data_start_row = self.ask_positive_int("多表合并", "数据起始行号：", 2)
            if data_start_row is None:
                self.logger.info("用户已取消操作")
                return
            validate_row_numbers(header_row, data_start_row)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return

        self.button("btn_merge_sheets").config(state="disabled")
        try:
            self.logger.info(f"源工作簿：{source_path}")
            result = merge_workbook_sheets_to_new_sheet(
                source_path=source_path,
                header_row=header_row,
                data_start_row=data_start_row,
                result_sheet_name=result_sheet_name,
                logger=self.logger,
            )
            self.logger.info(f"工作簿名：{result['workbook_name']}")
            self.logger.info(f"结果 sheet 名：{result['result_sheet_name']}")
            self.logger.info(f"合并 sheet 数：{result['source_sheet_count']}")
            self.logger.info(f"追加行数：{result['appended_row_count']}")
        except Exception as e:
            self.logger.error(f"多表合并失败：{e}")
        finally:
            self.button("btn_merge_sheets").config(state="normal")

    def run_merge_workbooks(self):
        """按钮回调函数，将多个源工作簿中的指定工作表导入到一个已有目标工作簿"""
        self.log_info("多簿到一簿：开始操作。")
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

        target_path = filedialog.askopenfilename(
            title="请选择已有目标 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not target_path:
            self.log_info("用户已取消操作。")
            return

        requested_sheet_name = self.ask_text(
            "多簿到一簿",
            "请输入源 sheet 名；留空则导入每个源文件的第一个可见 sheet。",
            entry_width=28,
            dialog_width=440,
            wraplength=380,
        )
        if requested_sheet_name is None:
            self.log_info("用户已取消操作。")
            return

        values_only = self.ask_choice(
            "多簿到一簿",
            "请选择复制方式：\n默认建议选择“否，保留公式和基础格式”。",
            [
                ("否，保留公式和基础格式", False),
                ("是，只复制值", True),
            ],
            dialog_width=500,
            wraplength=440,
        )
        if values_only is None:
            self.log_info("用户已取消操作。")
            return

        confirmed = self.ask_choice(
            "多簿到一簿",
            (
                "执行前请确认：\n"
                "目标工作簿会被修改并保存；\n"
                "请先关闭目标工作簿；\n"
                "程序会自动生成备份；\n"
                "源文件不会被修改。"
            ),
            [
                ("我已关闭目标，继续", True),
            ],
            dialog_width=500,
            wraplength=440,
        )
        if confirmed is not True:
            self.log_info("用户已取消操作。")
            return

        self.button("btn_merge_workbooks").config(state="disabled")
        try:
            self.log_info(f"源文件数量：{len(source_paths)}")
            self.log_info(f"目标工作簿：{target_path}")
            result = merge_workbooks_to_existing_workbook(
                source_paths=list(source_paths),
                target_path=target_path,
                requested_sheet_name=requested_sheet_name,
                values_only=values_only,
                logger=self.flushing_logger(),
            )
            self.log_info(
                "多簿到一簿完成："
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个；"
                f"失败 {result['failed_count']} 个。"
            )
            self.show_info(
                "多簿到一簿",
                "导入完成。\n"
                f"成功导入数量：{result['success_count']}\n"
                f"跳过数量：{result['skipped_count']}\n"
                f"失败数量：{result['failed_count']}\n"
                f"备份路径：{result['backup_path']}\n"
                f"目标工作簿路径：{result['target_path']}\n\n"
                "已生成多簿汇总目录和多簿汇总日志。",
                dialog_width=620,
                wraplength=560,
            )
        except Exception as e:
            self.log_error(f"多簿到一簿失败：{e}")
            self.show_info(
                "多簿到一簿",
                f"多簿到一簿失败：{e}",
                dialog_width=560,
                wraplength=500,
            )
        finally:
            self.button("btn_merge_workbooks").config(state="normal")
