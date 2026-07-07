# -*- coding: utf-8 -*-
"""数据穿透查询：单文件 / 多文件模式。COM 逻辑在 ops 层，这里只做交互编排。"""
import os
from tkinter import filedialog

from ...data_drill_ops import (
    build_data_drill_range_records,
    build_unique_output_path,
    summarize_data_drill_records,
    write_data_drill_result_workbook,
)
from ...excel_com import run_in_excel_com_session
from ...workbook_drill_ops import (
    RESULT_SHEET_BASE_NAME,
    execute_single_workbook_drill,
    get_active_drill_context,
)
from .base import FeatureController


class DataDrillController(FeatureController):
    def run_data_drill(self):
        """统一入口：按当前活动工作表和选区执行单文件或多文件数据穿透查询"""
        self.log_info("数据穿透查询：开始选择查询方式。")
        try:
            mode = self.ask_choice(
                "数据穿透查询",
                "请选择数据穿透查询方式：\n\n"
                "单文件数据穿透查询\n"
                "适合：在当前工作簿内，汇总所有 sheet 同一单元格或同一区域的值。\n\n"
                "多文件数据穿透查询\n"
                "适合：从多个工作簿中，读取同名 sheet 同一单元格或同一区域的值。",
                [
                    ("单文件数据穿透查询", "single"),
                    ("多文件数据穿透查询", "multi"),
                ],
                dialog_width=575,
                dialog_height=320,
                wraplength=520,
            )
            if mode is None:
                self.log_info("用户已取消操作。")
                return

            self.button("btn_data_drill").config(state="disabled")
            if mode == "single":
                self._run_single_workbook_data_drill()
            elif mode == "multi":
                self._run_multi_workbook_data_drill()
        except Exception as e:
            self.log_error(f"数据穿透查询失败：{type(e).__name__}: {e}")
            self.show_info(
                "数据穿透查询",
                f"数据穿透查询失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.button("btn_data_drill").config(state="normal")

    def _confirm_data_drill_context(self, context):
        choice = self.ask_choice(
            "数据穿透查询",
            "请确认当前取数点：\n"
            f"当前工作簿名：{context['workbook_name']}\n"
            f"当前工作簿路径：{context['workbook_path_text']}\n"
            f"当前 Sheet 名：{context['sheet_name']}\n"
            f"当前选区：{context['range_address']}\n"
            f"即将读取各源文件中的：{context['sheet_name']}!{context['range_address']}\n\n"
            "结果文件将生成到当前工作簿同目录。\n"
            "当前工作簿不会被修改或保存。",
            [("继续选择源文件", "continue")],
            dialog_width=760,
            wraplength=680,
        )
        return choice == "continue"

    def _run_single_workbook_data_drill(self):
        self.log_info("数据穿透查询（单文件）：开始操作。")
        result = run_in_excel_com_session(
            lambda excel: execute_single_workbook_drill(excel, logger=self.flushing_logger())
        )
        self.show_info(
            "数据穿透查询（单文件）",
            "数据穿透查询（单文件）完成。\n"
            f"当前工作簿：{result['workbook_path_text']}\n"
            f"当前工作表：{result['sheet_name']}\n"
            f"当前选区：{result['range_address']}\n"
            f"读取工作表数量：{result['visible_sheet_count']}\n"
            f"结果工作表：{result['result_sheet_name']}\n\n"
            "当前活动工作簿未自动保存。",
            dialog_width=620,
            wraplength=560,
        )

    def _run_multi_workbook_data_drill(self):
        self.log_info("数据穿透查询（多文件）：开始操作。")
        context = get_active_drill_context(require_saved_workbook=True)
        if not self._confirm_data_drill_context(context):
            self.log_info("用户已取消操作。")
            return

        source_paths = filedialog.askopenfilenames(
            title="请选择源 Excel 文件",
            initialdir=context["output_dir"],
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_paths:
            self.log_info("用户已取消操作。")
            return

        output_path = build_unique_output_path(context["output_dir"], base_name=RESULT_SHEET_BASE_NAME)
        self.log_info(f"当前工作表：{context['sheet_name']}")
        self.log_info(f"当前选区：{context['range_address']}")
        self.log_info(f"源文件数量：{len(source_paths)}")
        records = build_data_drill_range_records(
            source_paths=list(source_paths),
            sheet_name=context["sheet_name"],
            range_address=context["range_address"],
            logger=self.flushing_logger(),
        )
        write_data_drill_result_workbook(
            records=records,
            output_path=output_path,
            source_sheet_name=context["sheet_name"],
            cell_address=context["range_address"],
        )
        summary = summarize_data_drill_records(records)

        try:
            os.startfile(output_path)
            self.log_info(f"结果文件已生成并尝试打开：{output_path}")
        except Exception as open_error:
            self.log_error(f"结果文件已生成，但自动打开失败：{open_error}")
            self.log_info(f"请手动打开结果文件：{output_path}")

        self.log_info(f"输出文件：{output_path}")
        self.log_info("数据穿透查询（多文件）完成。")
        self.log_info(
            "数据穿透查询（多文件）统计："
            f"成功 {summary['success_count']} 个；"
            f"跳过 {summary['skipped_count']} 个；"
            f"失败 {summary['failed_count']} 个。"
        )
        self.log_info("当前活动工作簿和源文件均未被修改、未被自动保存。")
        self.show_info(
            "数据穿透查询（多文件）",
            "数据穿透查询（多文件）完成。\n"
            f"当前工作表：{context['sheet_name']}\n"
            f"当前选区：{context['range_address']}\n"
            f"源文件数量：{len(source_paths)}\n"
            f"成功数量：{summary['success_count']}\n"
            f"跳过数量：{summary['skipped_count']}\n"
            f"失败数量：{summary['failed_count']}\n"
            f"输出文件：{output_path}\n\n"
            "当前活动工作簿和源文件均未被修改、未被自动保存。",
            dialog_width=640,
            wraplength=580,
        )
