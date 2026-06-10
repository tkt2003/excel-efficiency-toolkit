import tkinter as tk
from tkinter import filedialog, scrolledtext, simpledialog
from .export_ops import export_workbook_sheets_to_files
from .logging_utils import setup_logger
from .sheet_ops import generate_sheet_index_with_links, list_sheet_names_to_active_sheet
from .table_ops import (
    merge_visible_sheets_to_new_sheet,
    parse_column_index,
    split_active_sheet_by_column,
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

        self.btn_list_sheets = tk.Button(
            self.frame_top,
            text="列出当前工作簿所有工作表",
            font=("Microsoft YaHei", 12),
            command=self.run_list_sheets,
            bg="#f0f0f0"
        )
        self.btn_list_sheets.pack()

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
            text="按指定列拆分当前表为多个工作表",
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

    def run_list_sheets(self):
        """按钮回调函数，执行获取工作表的操作"""
        self.btn_list_sheets.config(state="disabled")
        try:
            list_sheet_names_to_active_sheet(self.logger)
        finally:
            self.btn_list_sheets.config(state="normal")

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
        """按钮回调函数，将可见工作表合并到一个新工作表"""
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
            result = merge_visible_sheets_to_new_sheet(
                header_row=header_row,
                data_start_row=data_start_row,
                result_sheet_name=result_sheet_name,
                logger=self.logger,
            )
            self.logger.info(f"结果 sheet 名：{result['result_sheet_name']}")
            self.logger.info(f"合并 sheet 数：{result['source_sheet_count']}")
            self.logger.info(f"追加行数：{result['appended_row_count']}")
        except Exception as e:
            self.logger.error(f"多表合并失败：{e}")
        finally:
            self.btn_merge_sheets.config(state="normal")

    def run_split_sheet(self):
        """按钮回调函数，按指定列拆分当前活动工作表"""
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
            result = split_active_sheet_by_column(
                column_input=column_input,
                header_row=header_row,
                data_start_row=data_start_row,
                logger=self.logger,
            )
            self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
            self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
            self.logger.info(f"复制行数：{result['copied_row_count']}")
        except Exception as e:
            self.logger.error(f"按列拆分失败：{e}")
        finally:
            self.btn_split_sheet.config(state="normal")

    def run_sheet_index(self):
        """按钮回调函数，生成带超链接的工作表目录"""
        self.btn_sheet_index.config(state="disabled")
        try:
            result = generate_sheet_index_with_links(self.logger)
            self.logger.info(f"工作表数量：{result['sheet_count']}")
        except Exception as e:
            self.logger.error(f"生成带链接的工作表目录失败：{e}")
        finally:
            self.btn_sheet_index.config(state="normal")

def main():
    root = tk.Tk()
    app = ExcelToolkitApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
