import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, simpledialog
from .delete_sheet_ops import (
    execute_batch_delete_sheets_in_place,
    generate_temporary_delete_sheet_rule_table,
    infer_delete_mode_from_rule_values,
    normalize_delete_mode,
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
        value = simpledialog.askstring(title, prompt, initialvalue=str(initialvalue), parent=self.root)
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

        result_sheet_name = simpledialog.askstring(
            "多表合并",
            "结果 sheet 名：",
            initialvalue="合并结果",
            parent=self.root,
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

        source_sheet_name = simpledialog.askstring(
            "按列拆分",
            "源 sheet 名（只有一个 sheet 时可留空）：",
            parent=self.root,
        )
        if source_sheet_name is None:
            self.logger.info("用户已取消操作")
            return

        column_input = simpledialog.askstring(
            "按列拆分",
            "拆分列，例如 A、B、C 或 1、2、3：",
            parent=self.root,
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

    def _confirm_delete_rule_ready(self, rule_table_path):
        result = {"execute": False}
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title("批量删除工作表")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        try:
            dialog.attributes("-toolwindow", True)
        except tk.TclError:
            pass

        tk.Label(
            dialog,
            text="规则表已打开。请在 B/C/D 列填写规则并保存规则表后，再点击执行。",
            font=("Microsoft YaHei", 11),
            wraplength=460,
            justify="left",
            padx=20,
            pady=16,
        ).pack()

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=20, pady=(0, 16), fill=tk.X)

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
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量删除工作表失败：{e}")
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            dialog.destroy()

        execute_button = tk.Button(
            button_frame,
            text="我已填好规则，执行批量删除",
            command=execute,
            font=("Microsoft YaHei", 10),
        )
        execute_button.pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(
            button_frame,
            text="取消",
            command=cancel,
            font=("Microsoft YaHei", 10),
        ).pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.update_idletasks()
        width = max(dialog.winfo_reqwidth(), 430)
        height = max(dialog.winfo_reqheight(), 160)
        self._center_window(dialog, width=width, height=height)
        dialog.deiconify()
        dialog.lift()
        dialog.focus_force()
        dialog.after(
            50,
            lambda: dialog.winfo_exists() and self._center_window(dialog, width=width, height=height),
        )
        self.root.wait_window(dialog)
        return result["execute"]

    def _resolve_delete_mode_from_rule_table(self, rule_table_path):
        rule_values = read_rule_values_from_rule_table(rule_table_path)
        inferred_mode = infer_delete_mode_from_rule_values(rule_values)
        if inferred_mode is not None:
            return inferred_mode

        mode_input = simpledialog.askstring(
            "批量删除工作表",
            "B列和C列都填写了规则，请选择执行模式：\n1 保留模式：只保留 B 列表格名，删除其他表格\n2 删除模式：删除 C 列表格名",
            parent=self.root,
        )
        if mode_input is None:
            self.logger.info("用户已取消操作")
            return None

        try:
            return normalize_delete_mode(mode_input)
        except ValueError as e:
            self.logger.error(f"输入无效：{e}")
            return None

def main():
    root = tk.Tk()
    app = ExcelToolkitApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
