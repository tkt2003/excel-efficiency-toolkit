# -*- coding: utf-8 -*-
"""工作簿辅助：生成工作表目录、选区 ROUND 保留两位。"""
from tkinter import filedialog

from ...round_formula_ops import round_selected_range_to_two_decimals
from ...sheet_ops import generate_sheet_index_sheet_with_links
from .base import FeatureController


class WorkbookMiscController(FeatureController):
    def run_sheet_index(self):
        """按钮回调函数，在所选工作簿中生成带超链接的工作表目录"""
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

        self.button("btn_sheet_index").config(state="disabled")
        try:
            result = generate_sheet_index_sheet_with_links(source_path=source_path, logger=self.logger)
            self.logger.info(f"源工作簿路径：{source_path}")
            self.logger.info(f"工作簿名：{result['workbook_name']}")
            self.logger.info(f"目录 sheet 名：{result['index_sheet_name']}")
            self.logger.info(f"收录 sheet 数量：{result['sheet_count']}")
        except Exception as e:
            self.logger.error(f"生成带链接的工作表目录失败：{e}")
        finally:
            self.button("btn_sheet_index").config(state="normal")

    def run_round_formula(self):
        """按钮回调函数，将当前 Excel 选区内数值/公式直接包裹 ROUND(...,2)"""
        self.log_info("选区 ROUND 保留两位：开始操作。")
        self.button("btn_round_formula").config(state="disabled")
        try:
            result = round_selected_range_to_two_decimals(logger=self.flushing_logger())
            self.log_info(
                "选区 ROUND 保留两位完成："
                f"工作簿 {result['workbook_name']}；"
                f"Sheet {result['sheet_name']}；"
                f"选区 {result['selection_address']}；"
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个。"
            )
            self.log_info("当前工作簿未自动保存，请检查后自行保存。")
            self.show_info(
                "选区 ROUND 保留两位",
                "处理完成。\n"
                f"当前工作簿名：{result['workbook_name']}\n"
                f"当前 sheet 名：{result['sheet_name']}\n"
                f"当前选区地址：{result['selection_address']}\n"
                f"成功处理数量：{result['success_count']}\n"
                f"跳过数量：{result['skipped_count']}\n\n"
                "当前工作簿未自动保存，请检查后自行保存。",
                dialog_width=460,
                dialog_height=340,
                wraplength=400,
            )
        except Exception as e:
            self.log_error(f"选区 ROUND 保留两位失败：{type(e).__name__}: {e}")
            self.show_info(
                "选区 ROUND 保留两位",
                f"选区 ROUND 保留两位失败：{e}",
                dialog_width=420,
                dialog_height=240,
                wraplength=380,
            )
        finally:
            self.button("btn_round_formula").config(state="normal")
