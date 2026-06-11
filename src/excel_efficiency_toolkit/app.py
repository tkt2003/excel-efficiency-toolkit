import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
from .color_sum_ops import sum_current_sheet_by_fill_color
from .delete_sheet_ops import (
    execute_batch_delete_sheets_in_place,
    generate_temporary_delete_sheet_rule_table,
    infer_delete_mode_from_rule_values,
    read_rule_values_from_rule_table,
)
from .export_ops import export_workbook_sheets_to_files
from .logging_utils import setup_logger
from .sheet_ops import generate_sheet_index_sheet_with_links
from .table_ops import (
    merge_workbook_sheets_to_new_sheet,
    parse_column_index,
    split_workbook_sheet_by_column,
    validate_row_numbers,
)

class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 效率工具台")
        self.root.geometry("760x620")

        # 顶部按钮区域
        self.frame_top = tk.Frame(root)
        self.frame_top.pack(pady=20, fill=tk.X)

        self.btn_export_sheets = tk.Button(
            self.frame_top,
            text="一簿按工作表拆分为多个文件",
            font=("Microsoft YaHei", 12),
            command=self.run_export_sheets,
            bg="#f0f0f0"
        )
        self.btn_export_sheets.pack(pady=(10, 0))

        self.btn_merge_sheets = tk.Button(
            self.frame_top,
            text="多表合并到一个新表",
            font=("Microsoft YaHei", 12),
            command=self.run_merge_sheets,
            bg="#f0f0f0"
        )
        self.btn_merge_sheets.pack(pady=(10, 0))

        self.btn_split_sheet = tk.Button(
            self.frame_top,
            text="按指定列拆分工作表",
            font=("Microsoft YaHei", 12),
            command=self.run_split_sheet,
            bg="#f0f0f0"
        )
        self.btn_split_sheet.pack(pady=(10, 0))

        self.btn_sheet_index = tk.Button(
            self.frame_top,
            text="生成带链接的工作表目录",
            font=("Microsoft YaHei", 12),
            command=self.run_sheet_index,
            bg="#f0f0f0"
        )
        self.btn_sheet_index.pack(pady=(10, 0))

        self.btn_delete_sheets = tk.Button(
            self.frame_top,
            text="批量删除工作表",
            font=("Microsoft YaHei", 12),
            command=self.run_delete_sheets,
            bg="#f0f0f0"
        )
        self.btn_delete_sheets.pack(pady=(10, 0))

        self.btn_color_sum = tk.Button(
            self.frame_top,
            text="按颜色汇总求和",
            font=("Microsoft YaHei", 12),
            command=self.run_color_sum,
            bg="#f0f0f0"
        )
        self.btn_color_sum.pack(pady=(10, 0))

        # 底部日志区域
        self.frame_bottom = tk.Frame(root)
        self.frame_bottom.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        # 日志标签
        tk.Label(self.frame_bottom, text="运行日志：").pack(anchor="w")

        # 日志输出文本框
        self.log_text = scrolledtext.ScrolledText(
            self.frame_bottom, 
            state='disabled', 
            height=15, 
            font=("Consolas", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 初始化自定义 logger
        self.logger = setup_logger(self.log_text)
        self.logger.info("欢迎使用 Excel 效率工具台。程序已就绪。")

    def _center_window(self, window, width=None, height=None):
        window.update_idletasks()
        width = width or window.winfo_reqwidth()
        height = height or window.winfo_reqheight()

        try:
            parent_x = self.root.winfo_rootx()
            parent_y = self.root.winfo_rooty()
            parent_w = self.root.winfo_width()
            parent_h = self.root.winfo_height()
        except tk.TclError:
            parent_w = 0
            parent_h = 0

        if parent_w > 1 and parent_h > 1:
            x = parent_x + (parent_w - width) // 2
            y = parent_y + (parent_h - height) // 2
        else:
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = int((screen_width - width) / 2)
            y = int((screen_height - height) / 2)

        x = max(0, x)
        y = max(0, y)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def _flush_ui(self):
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            pass

    def _log_info(self, message):
        self.logger.info(message)
        self._flush_ui()

    def _log_error(self, message):
        self.logger.error(message)
        self._flush_ui()

    def _flushing_logger(self):
        return _FlushingLogger(self.logger, self._flush_ui)

    def _show_dialog_no_grab(self, dialog, focus_widget=None, min_width=380, min_height=None):
        dialog.update_idletasks()
        width = max(dialog.winfo_reqwidth(), min_width)
        height = max(dialog.winfo_reqheight(), min_height) if min_height else None
        self._center_window(dialog, width=width, height=height)
        dialog.deiconify()
        dialog.lift()
        try:
            dialog.attributes("-topmost", True)
            dialog.after(200, lambda: dialog.winfo_exists() and dialog.attributes("-topmost", False))
        except tk.TclError:
            pass
        try:
            dialog.focus_force()
        except tk.TclError:
            pass
        if focus_widget is not None:
            try:
                focus_widget.focus_set()
            except tk.TclError:
                pass

    def _ask_text_no_grab(self, title, prompt, default="", entry_width=30, dialog_width=380, wraplength=340):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text=prompt,
            font=("Microsoft YaHei", 11),
            wraplength=wraplength,
            justify="left",
            padx=16,
            pady=8,
        ).pack(anchor="w")

        entry = tk.Entry(dialog, font=("Microsoft YaHei", 11), width=entry_width)
        entry.pack(padx=16, pady=(0, 12), fill=tk.X)
        entry.insert(0, default)
        entry.select_range(0, tk.END)

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)

        def confirm():
            result["value"] = entry.get()
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        tk.Button(button_frame, text="取消", command=cancel, font=("Microsoft YaHei", 10), width=10).pack(
            side=tk.RIGHT,
        )
        tk.Button(button_frame, text="确定", command=confirm, font=("Microsoft YaHei", 10), width=10).pack(
            side=tk.RIGHT,
            padx=(0, 8),
        )

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Return>", lambda event: confirm())
        dialog.bind("<Escape>", lambda event: cancel())
        self._show_dialog_no_grab(dialog, focus_widget=entry, min_width=dialog_width)
        self.root.wait_variable(done)
        return result["value"]

    def _ask_choice_no_grab(self, title, prompt, choices, dialog_width=420, wraplength=360):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text=prompt,
            font=("Microsoft YaHei", 11),
            wraplength=wraplength,
            justify="left",
            padx=16,
            pady=12,
        ).pack(anchor="w")

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)

        def choose(value):
            result["value"] = value
            done.set(True)
            dialog.destroy()

        tk.Button(
            button_frame,
            text="取消",
            command=lambda: choose(None),
            font=("Microsoft YaHei", 10),
            width=10,
        ).pack(side=tk.RIGHT)

        for label, value in reversed(choices):
            tk.Button(
                button_frame,
                text=label,
                command=lambda selected=value: choose(selected),
                font=("Microsoft YaHei", 10),
                width=14,
            ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        dialog.bind("<Escape>", lambda event: choose(None))
        self._show_dialog_no_grab(dialog, min_width=dialog_width)
        self.root.wait_variable(done)
        return result["value"]

    def _show_info_no_grab(self, title, message, dialog_width=400, wraplength=360):
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text=message,
            font=("Microsoft YaHei", 11),
            wraplength=wraplength,
            justify="left",
            padx=16,
            pady=12,
        ).pack(anchor="w")

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)
        tk.Button(
            button_frame,
            text="确定",
            command=dialog.destroy,
            font=("Microsoft YaHei", 10),
            width=10,
        ).pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self._show_dialog_no_grab(dialog, min_width=dialog_width)

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

        self.btn_export_sheets.config(state="disabled")
        try:
            self.logger.info(f"源文件：{source_path}")
            self.logger.info(f"输出目录：{output_dir}")
            exported_paths = export_workbook_sheets_to_files(source_path, output_dir, self.logger)
            self.logger.info(f"成功导出 {len(exported_paths)} 个文件。")
        except Exception as e:
            self.logger.error(f"导出失败：请确认文件未损坏、已安装 Microsoft Excel，并且输出目录可写。详细信息：{e}")
        finally:
            self.btn_export_sheets.config(state="normal")

    def _ask_positive_int(self, title, prompt, initialvalue):
        value = self._ask_text_no_grab(title, prompt, default=str(initialvalue))
        if value is None:
            return None
        try:
            return int(value.strip())
        except ValueError:
            raise ValueError(f"{prompt}必须是整数。")

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

        result_sheet_name = self._ask_text_no_grab(
            "多表合并",
            "结果 sheet 名：",
            default="合并结果",
        )
        if result_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        try:
            header_row = self._ask_positive_int("多表合并", "表头行号：", 1)
            if header_row is None:
                self.logger.info("用户已取消操作")
                return
            data_start_row = self._ask_positive_int("多表合并", "数据起始行号：", 2)
            if data_start_row is None:
                self.logger.info("用户已取消操作")
                return
            validate_row_numbers(header_row, data_start_row)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return

        self.btn_merge_sheets.config(state="disabled")
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
            self.btn_merge_sheets.config(state="normal")

    def run_split_sheet(self):
        """按钮回调函数，按指定列拆分所选工作簿中的指定工作表"""
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

        source_sheet_name = self._ask_text_no_grab(
            "按列拆分",
            "源 sheet 名（只有一个 sheet 时可留空）：",
        )
        if source_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        column_input = self._ask_text_no_grab(
            "按列拆分",
            "拆分列，例如 A、B、C 或 1、2、3：",
        )
        if column_input is None:
            self.logger.info("用户已取消操作")
            return

        try:
            parse_column_index(column_input)
            header_row = self._ask_positive_int("按列拆分", "表头行号：", 1)
            if header_row is None:
                self.logger.info("用户已取消操作")
                return
            data_start_row = self._ask_positive_int("按列拆分", "数据起始行号：", 2)
            if data_start_row is None:
                self.logger.info("用户已取消操作")
                return
            validate_row_numbers(header_row, data_start_row)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return

        self.btn_split_sheet.config(state="disabled")
        try:
            self.logger.info(f"源工作簿：{source_path}")
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
            self.btn_split_sheet.config(state="normal")

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

        self.btn_sheet_index.config(state="disabled")
        try:
            result = generate_sheet_index_sheet_with_links(source_path=source_path, logger=self.logger)
            self.logger.info(f"源工作簿路径：{source_path}")
            self.logger.info(f"工作簿名：{result['workbook_name']}")
            self.logger.info(f"目录 sheet 名：{result['index_sheet_name']}")
            self.logger.info(f"收录 sheet 数量：{result['sheet_count']}")
        except Exception as e:
            self.logger.error(f"生成带链接的工作表目录失败：{e}")
        finally:
            self.btn_sheet_index.config(state="normal")

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

        self.btn_delete_sheets.config(state="disabled")
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
            self.btn_delete_sheets.config(state="normal")

    def run_color_sum(self):
        """按钮回调函数，按当前选中单元格填充色汇总指定 sheet 的同地址单元格"""
        self._log_info("按颜色汇总求和：开始操作。")
        try:
            target_sheet_name = self._ask_text_no_grab(
                "按颜色汇总求和",
                "请输入要汇总的 sheet 名；留空则使用当前活动 sheet。",
                entry_width=26,
                dialog_width=360,
                wraplength=320,
            )
            if target_sheet_name is None:
                self._log_info("用户已取消操作。")
                return

            source_paths = filedialog.askopenfilenames(
                title="请选择源 Excel 文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                    ("所有文件", "*.*"),
                ],
            )
            if not source_paths:
                self._log_info("用户已取消操作。")
                return

            write_mode = self._ask_choice_no_grab(
                "按颜色汇总求和",
                "请选择写入方式：\n1 写入求和公式\n2 只写入汇总数值",
                [
                    ("写入求和公式", "formula"),
                    ("只写入汇总数值", "value"),
                ],
            )
            if write_mode is None:
                self._log_info("用户已取消操作。")
                return

            self.btn_color_sum.config(state="disabled")
            result = sum_current_sheet_by_fill_color(
                source_paths=list(source_paths),
                write_mode=write_mode,
                target_sheet_name=target_sheet_name,
                logger=self._flushing_logger(),
            )
            self._log_info(
                "按颜色汇总求和完成："
                f"写入 {result['written_cell_count']} 个单元格；"
                f"找到同名工作表源文件 {result['matched_source_file_count']} 个；"
                f"缺少同名工作表 {result['missing_sheet_file_count']} 个；"
                f"忽略非数字 {result['ignored_non_numeric_count']} 个。"
            )
            self._log_info("目标工作簿未自动保存，请检查后自行保存。")
            self._show_info_no_grab(
                "按颜色汇总求和",
                "汇总完成。\n"
                f"目标工作表：{result['target_sheet_name']}\n"
                f"写入单元格：{result['written_cell_count']}\n"
                f"参与源文件：{result['matched_source_file_count']}\n"
                f"缺少同名工作表：{result['missing_sheet_file_count']}\n"
                f"忽略非数字：{result['ignored_non_numeric_count']}\n\n"
                "目标工作簿未自动保存，请检查后自行保存。",
            )
        except Exception as e:
            self._log_error(f"按颜色汇总求和失败：{type(e).__name__}: {e}")
        finally:
            self.btn_color_sum.config(state="normal")

    def _confirm_delete_rule_ready(self, rule_table_path):
        result = {"execute": False}
        done = tk.BooleanVar(value=False)
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title("批量删除工作表")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text="规则表已打开。请在 B/C/D 列填写规则并保存规则表后，再点击执行。",
            font=("Microsoft YaHei", 11),
            wraplength=360,
            justify="left",
            padx=16,
            pady=12,
        ).pack()

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)

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

        tk.Button(
            button_frame,
            text="取消",
            command=cancel,
            font=("Microsoft YaHei", 10),
            width=10,
        ).pack(side=tk.RIGHT)
        execute_button = tk.Button(
            button_frame,
            text="我已填好规则，执行批量删除",
            command=execute,
            font=("Microsoft YaHei", 10),
            width=22,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, min_width=400, min_height=160)
        self.root.wait_variable(done)
        return result["execute"]

    def _resolve_delete_mode_from_rule_table(self, rule_table_path):
        rule_values = read_rule_values_from_rule_table(rule_table_path)
        inferred_mode = infer_delete_mode_from_rule_values(rule_values)
        if inferred_mode is not None:
            return inferred_mode

        mode = self._ask_choice_no_grab(
            "批量删除工作表",
            "B列和C列都填写了规则，请选择执行模式：\n1 保留模式：只保留 B 列表格名，删除其他表格\n2 删除模式：删除 C 列表格名",
            [
                ("保留模式", "keep"),
                ("删除模式", "delete"),
            ],
        )
        if mode is None:
            self.logger.info("用户已取消操作")
            return None

        return mode


class _FlushingLogger:
    def __init__(self, logger, flush):
        self._logger = logger
        self._flush = flush

    def info(self, message):
        self._logger.info(message)
        self._flush()

    def error(self, message):
        self._logger.error(message)
        self._flush()

    def __getattr__(self, name):
        return getattr(self._logger, name)

def main():
    root = tk.Tk()
    app = ExcelToolkitApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
