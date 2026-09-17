# -*- coding: utf-8 -*-
"""拆分导出：按工作表拆分文件、按列拆分、按行拆分。"""
from tkinter import filedialog

from ...export_ops import export_workbook_sheets_to_files
from ...table_ops import (
    parse_column_index,
    split_workbook_sheet_by_column,
    split_workbook_sheet_by_column_to_files,
    split_workbook_sheet_by_rows,
    split_workbook_sheet_by_rows_to_files,
    validate_row_numbers,
    validate_rows_per_part,
)
from .base import FeatureController


class ExportSplitController(FeatureController):
    def run_export_sheets(self):
        """按钮回调函数，将一个工作簿按工作表拆分为多个文件"""
        source_path = filedialog.askopenfilename(
            title="请选择源 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_path:
            self.logger.info("用户已取消操作")
            return

        output_dir = filedialog.askdirectory(title="请选择输出目录")
        if not output_dir:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_export_sheets").config(state="disabled")
        try:
            self.logger.info(f"源文件：{source_path}")
            self.logger.info(f"输出目录：{output_dir}")
            exported_paths = export_workbook_sheets_to_files(source_path, output_dir, self.logger)
            self.logger.info(f"成功导出 {len(exported_paths)} 个文件。")
        except Exception as e:
            self.logger.error(f"导出失败：请确认文件未损坏、已安装 Microsoft Excel，并且输出目录可写。详细信息：{e}")
        finally:
            self.button("btn_export_sheets").config(state="normal")

    def run_split_sheet(self):
        """按钮回调函数，按指定列内容拆分为工作表或多个 Excel 文件"""
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

        source_sheet_name = self.ask_text(
            "按列拆分",
            "源 sheet 名（只有一个 sheet 时可留空）：",
        )
        if source_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        column_input = self.ask_text(
            "按列拆分",
            "拆分列，例如 A、B、C 或 1、2、3：",
        )
        if column_input is None:
            self.logger.info("用户已取消操作")
            return

        try:
            parse_column_index(column_input)
            header_row = self.ask_positive_int("按列拆分", "表头行号：", 1)
            if header_row is None:
                self.logger.info("用户已取消操作")
                return
            data_start_row = self.ask_positive_int("按列拆分", "数据起始行号：", 2)
            if data_start_row is None:
                self.logger.info("用户已取消操作")
                return
            validate_row_numbers(header_row, data_start_row)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return

        output_mode = self.ask_choice(
            "按列拆分",
            "请选择输出方式：",
            [
                ("拆分为工作表", "sheets"),
                ("拆分为多个文件", "files"),
            ],
        )
        if output_mode is None:
            self.logger.info("用户已取消操作")
            return

        if output_mode == "files":
            output_dir = filedialog.askdirectory(title="请选择拆分文件的输出目录")
            if not output_dir:
                self.logger.info("用户已取消操作")
                return

        self.button("btn_split_sheet").config(state="disabled")
        try:
            self.logger.info(f"源工作簿：{source_path}")
            if output_mode == "files":
                self.logger.info(f"输出目录：{output_dir}")
                result = split_workbook_sheet_by_column_to_files(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    column_input=column_input,
                    output_dir=output_dir,
                    header_row=header_row,
                    data_start_row=data_start_row,
                    logger=self.logger,
                )
                self.logger.info(f"工作簿名：{result['workbook_name']}")
                self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
                self.logger.info(f"已生成 {result['created_file_count']} 个 Excel 文件。")
                self.logger.info(f"输出目录：{result['output_dir']}")
            else:
                result = split_workbook_sheet_by_column(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    column_input=column_input,
                    header_row=header_row,
                    data_start_row=data_start_row,
                    logger=self.logger,
                )
                self.logger.info(f"工作簿名：{result['workbook_name']}")
                self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
                self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
                self.logger.info(f"复制行数：{result['copied_row_count']}")
        except Exception as e:
            self.logger.error(f"按列拆分失败：{e}")
        finally:
            self.button("btn_split_sheet").config(state="normal")

    def run_split_rows(self):
        """按钮回调函数，按数据行拆分为工作表或多个 Excel 文件"""
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

        source_sheet_name = self.ask_text(
            "按行拆分",
            "源 sheet 名（只有一个 sheet 时可留空）：",
        )
        if source_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        try:
            header_row = self.ask_positive_int("按行拆分", "表头行号：", 1)
            if header_row is None:
                self.logger.info("用户已取消操作")
                return
            data_start_row = self.ask_positive_int("按行拆分", "数据起始行号：", 2)
            if data_start_row is None:
                self.logger.info("用户已取消操作")
                return
            validate_row_numbers(header_row, data_start_row)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return

        split_mode = self.ask_choice(
            "按行拆分",
            "请选择拆分方式：",
            [
                ("每行一份", "each_row"),
                ("每 N 行一份", "n_rows"),
            ],
        )
        if split_mode is None:
            self.logger.info("用户已取消操作")
            return

        if split_mode == "each_row":
            rows_per_part = 1
        else:
            try:
                rows_per_part = self.ask_positive_int("按行拆分", "每份数据行数：", 100)
                if rows_per_part is None:
                    self.logger.info("用户已取消操作")
                    return
                validate_rows_per_part(rows_per_part)
            except ValueError as e:
                self.logger.error(f"输入无效：{e}")
                return

        output_mode = self.ask_choice(
            "按行拆分",
            "请选择输出方式：",
            [
                ("拆分为工作表", "sheets"),
                ("拆分为多个文件", "files"),
            ],
        )
        if output_mode is None:
            self.logger.info("用户已取消操作")
            return

        if output_mode == "files":
            output_dir = filedialog.askdirectory(title="请选择拆分文件的输出目录")
            if not output_dir:
                self.logger.info("用户已取消操作")
                return

        self.button("btn_split_rows").config(state="disabled")
        try:
            self.logger.info(f"源工作簿：{source_path}")
            split_mode_desc = "每行一份" if rows_per_part == 1 else f"每 {rows_per_part} 行一份"
            if output_mode == "files":
                self.logger.info(f"输出目录：{output_dir}")
                result = split_workbook_sheet_by_rows_to_files(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    output_dir=output_dir,
                    rows_per_part=rows_per_part,
                    header_row=header_row,
                    data_start_row=data_start_row,
                    logger=self.logger,
                )
                self.logger.info(f"工作簿名：{result['workbook_name']}")
                self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
                self.logger.info(f"拆分方式：{split_mode_desc}")
                self.logger.info(f"成功生成 {result['created_file_count']} 个 Excel 文件。")
                self.logger.info(f"输出目录：{result['output_dir']}")
            else:
                result = split_workbook_sheet_by_rows(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    rows_per_part=rows_per_part,
                    header_row=header_row,
                    data_start_row=data_start_row,
                    logger=self.logger,
                )
                self.logger.info(f"工作簿名：{result['workbook_name']}")
                self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
                self.logger.info(f"拆分方式：{split_mode_desc}")
                self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
                self.logger.info(f"复制行数：{result['copied_row_count']}")
        except Exception as e:
            self.logger.error(f"按行拆分失败：{e}")
        finally:
            self.button("btn_split_rows").config(state="normal")
