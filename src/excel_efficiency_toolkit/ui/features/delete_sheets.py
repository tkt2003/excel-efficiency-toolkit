# -*- coding: utf-8 -*-
"""批量删除工作表：生成规则表 → 用户填写 → 确认执行。"""
import tkinter as tk
import os
from tkinter import filedialog

from ...delete_sheet_ops import (
    execute_batch_delete_sheets_in_place,
    generate_temporary_delete_sheet_rule_table,
    resolve_delete_mode_from_rule_table,
)
from .base import FeatureController


class DeleteSheetsController(FeatureController):
    def run_delete_sheets(self):
        """按钮回调函数，生成临时规则表后确认执行批量删除"""
        source_paths = filedialog.askopenfilenames(
            title="请选择要批量删除工作表的 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.button("btn_delete_sheets").config(state="disabled")
        try:
            result = generate_temporary_delete_sheet_rule_table(
                source_paths=list(source_paths),
                logger=self.logger,
            )
            self.logger.info(f"选择文件数量：{result['source_file_count']}")
            self.logger.info(f"读取成功文件数：{result['read_success_count']}")
            self.logger.info(f"唯一表格名数量：{result['unique_sheet_count']}")
            self.logger.info(f"临时规则表路径：{result['output_path']}")
            try:
                os.startfile(result["output_path"])
                self.logger.info("规则表已自动打开。请填写并保存规则表后，再点击弹窗中的执行按钮。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")

            if not self._confirm_delete_rule_ready(result["output_path"]):
                self.logger.info("用户已取消操作")
                return
        except Exception as e:
            self.logger.error(f"批量删除工作表失败：{e}")
        finally:
            self.button("btn_delete_sheets").config(state="normal")

    def _confirm_delete_rule_ready(self, rule_table_path):
        result = {"execute": False}
        done = tk.BooleanVar(value=False)
        dialog, card = self.dialogs.create_dialog_card("批量删除工作表")

        self.dialogs.add_message(
            card,
            "规则表已打开。请在 B/C/D 列填写规则并保存规则表后，再点击执行。",
            wraplength=360,
            pady=(0, 12),
        )

        button_frame = self.dialogs.create_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                mode = self._resolve_delete_mode_from_rule_table(rule_table_path)
                if mode is None:
                    return

                execute_result = execute_batch_delete_sheets_in_place(
                    rule_table_path=rule_table_path,
                    mode=mode,
                    logger=self.logger,
                )
                self.logger.info(f"有效源文件数：{execute_result['source_file_count']}")
                self.logger.info(f"成功处理文件数：{execute_result['processed_file_count']}")
                self.logger.info(f"跳过文件数：{execute_result['skipped_file_count']}")
                self.logger.info(f"失败文件数：{execute_result['failed_file_count']}")
                self.logger.info(f"删除工作表总数：{execute_result['deleted_sheet_count']}")
                self.logger.info(f"实际模式：{execute_result['mode']}")
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量删除工作表失败：{e}")
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self.dialogs.add_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self.dialogs.add_button(
            button_frame,
            "我已填好规则，执行批量删除",
            execute,
            primary=True,
            width=230,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.dialogs.show_no_grab(dialog, width=510, height=255)
        self.root.wait_variable(done)
        return result["execute"]

    def _resolve_delete_mode_from_rule_table(self, rule_table_path):
        inferred_mode = resolve_delete_mode_from_rule_table(rule_table_path)
        if inferred_mode is not None:
            return inferred_mode

        mode = self.ask_choice(
            "批量删除工作表",
            "B列和C列都填写了规则，请选择执行模式：\n1 保留模式：只保留 B 列表格名，删除其他表格\n2 删除模式：删除 C 列表格名",
            [
                ("保留模式", "keep"),
                ("删除模式", "delete"),
            ],
            dialog_width=760,
            wraplength=680,
        )
        if mode is None:
            self.logger.info("用户已取消操作")
            return None

        return mode
