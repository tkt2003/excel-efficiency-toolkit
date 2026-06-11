import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
from .color_sum_ops import sum_current_sheet_by_fill_color, sum_matching_sheets_by_fill_color
from .data_drill_ops import (
    build_data_drill_records,
    build_unique_output_path,
    summarize_data_drill_records,
    write_data_drill_result_workbook,
)
from .delete_sheet_ops import (
    execute_batch_delete_sheets_in_place,
    generate_temporary_delete_sheet_rule_table,
    infer_delete_mode_from_rule_values,
    read_rule_values_from_rule_table,
)
from .export_ops import export_workbook_sheets_to_files
from .logging_utils import setup_logger
from .rename_file_ops import (
    build_rename_plan,
    create_rename_rule_workbook,
    execute_rename_plan,
    read_rename_rules,
    read_rename_settings,
    write_rename_results_to_workbook,
)
from .rename_sheet_ops import (
    build_skipped_sheet_rename_actions,
    build_sheet_rename_plan,
    create_sheet_rename_rule_workbook,
    execute_sheet_rename_plan,
    group_sheet_rename_rules_by_workbook_path,
    is_excel_workbook_file,
    is_office_temp_file,
    is_sheet_hidden_by_visible_value,
    read_sheet_rename_rules,
    read_sheet_rename_settings,
    summarize_sheet_rename_actions,
    write_sheet_rename_results_to_workbook,
)
from .sheet_ops import generate_sheet_index_sheet_with_links
from .table_ops import (
    merge_workbook_sheets_to_new_sheet,
    parse_column_index,
    split_workbook_sheet_by_column,
    validate_row_numbers,
)
from .workbook_merge_ops import merge_workbooks_to_existing_workbook

class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 效率工具台")
        self.root.geometry("860x760")
        self.root.minsize(780, 720)
        self.bg_color = "#f5f6f8"
        self.card_color = "#ffffff"
        self.border_color = "#d7dce2"
        self.text_color = "#1f2937"
        self.root.configure(bg=self.bg_color)

        self.main_frame = tk.Frame(root, bg=self.bg_color)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        self._create_header(self.main_frame)
        self._create_feature_area(self.main_frame)
        self._create_log_area(self.main_frame)

        # 初始化自定义 logger
        self.logger = setup_logger(self.log_text)
        self.logger.info("欢迎使用 Excel 效率工具台。程序已就绪。")

    def _create_header(self, parent):
        header = tk.Frame(parent, bg=self.bg_color)
        header.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            header,
            text="Excel 效率工具台",
            font=("Microsoft YaHei", 18, "bold"),
            fg=self.text_color,
            bg=self.bg_color,
        ).pack(anchor="w")
        tk.Label(
            header,
            text="请选择需要执行的功能。多数功能不自动保存；会保存目标文件的功能会在执行前提示。",
            font=("Microsoft YaHei", 10),
            fg="#5f6b7a",
            bg=self.bg_color,
        ).pack(anchor="w", pady=(4, 0))

    def _create_feature_area(self, parent):
        feature_area = tk.Frame(parent, bg=self.bg_color)
        feature_area.pack(fill=tk.X, pady=(0, 14))
        feature_area.columnconfigure(0, weight=1, uniform="feature")
        feature_area.columnconfigure(1, weight=1, uniform="feature")

        left_column = tk.Frame(feature_area, bg=self.bg_color)
        right_column = tk.Frame(feature_area, bg=self.bg_color)
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right_column.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        self._create_feature_group(
            left_column,
            "拆分导出",
            [
                ("一簿按工作表拆分为多个文件", "btn_export_sheets", self.run_export_sheets),
                ("按指定列拆分工作表", "btn_split_sheet", self.run_split_sheet),
            ],
        )
        self._create_feature_group(
            left_column,
            "合并整理",
            [
                ("多表合并到一个新表", "btn_merge_sheets", self.run_merge_sheets),
                ("多簿到一簿", "btn_merge_workbooks", self.run_merge_workbooks),
            ],
        )
        self._create_feature_group(
            left_column,
            "模板填报 / 附注汇总",
            [
                ("按颜色汇总求和", "btn_color_sum", self.run_color_sum),
                ("数据穿透取数", "btn_data_drill", self.run_data_drill),
            ],
        )

        self._create_feature_group(
            right_column,
            "目录与检查",
            [
                ("生成带链接的工作表目录", "btn_sheet_index", self.run_sheet_index),
            ],
        )
        self._create_feature_group(
            right_column,
            "批量维护",
            [
                ("批量删除工作表", "btn_delete_sheets", self.run_delete_sheets),
                ("批量重命名文件", "btn_rename_files", self.run_batch_rename_files),
                ("批量重命名工作表", "btn_rename_sheets", self.run_batch_rename_sheets),
            ],
        )

    def _create_feature_group(self, parent, title, items):
        group = tk.LabelFrame(
            parent,
            text=title,
            font=("Microsoft YaHei", 11, "bold"),
            fg=self.text_color,
            bg=self.card_color,
            bd=1,
            relief="solid",
            padx=12,
            pady=10,
            labelanchor="nw",
        )
        group.pack(fill=tk.X, pady=(0, 10))

        for index, (text, attr_name, command) in enumerate(items):
            button = tk.Button(
                group,
                text=text,
                font=("Microsoft YaHei", 10),
                command=command,
                width=30,
                height=2,
                bg="#f8fafc",
                fg=self.text_color,
                activebackground="#e9eef5",
                activeforeground=self.text_color,
                relief="groove",
                bd=1,
                anchor="w",
                padx=10,
            )
            button.pack(fill=tk.X, pady=(0 if index == 0 else 8, 0))
            setattr(self, attr_name, button)

    def _create_log_area(self, parent):
        self.frame_bottom = tk.Frame(parent, bg=self.bg_color)
        self.frame_bottom.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            self.frame_bottom,
            text="运行日志",
            font=("Microsoft YaHei", 11, "bold"),
            fg=self.text_color,
            bg=self.bg_color,
        ).pack(anchor="w", pady=(0, 6))

        log_frame = tk.Frame(
            self.frame_bottom,
            bg=self.card_color,
            highlightbackground=self.border_color,
            highlightthickness=1,
        )
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            state="disabled",
            height=14,
            font=("Consolas", 10),
            bg="#ffffff",
            fg="#111827",
            insertbackground="#111827",
            relief="flat",
            bd=0,
            padx=8,
            pady=8,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

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

    def run_merge_workbooks(self):
        """按钮回调函数，将多个源工作簿中的指定工作表导入到一个已有目标工作簿"""
        self._log_info("多簿到一簿：开始操作。")
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

        target_path = filedialog.askopenfilename(
            title="请选择已有目标 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not target_path:
            self._log_info("用户已取消操作。")
            return

        requested_sheet_name = self._ask_text_no_grab(
            "多簿到一簿",
            "请输入源 sheet 名；留空则导入每个源文件的第一个可见 sheet。",
            entry_width=28,
            dialog_width=440,
            wraplength=380,
        )
        if requested_sheet_name is None:
            self._log_info("用户已取消操作。")
            return

        values_only = self._ask_choice_no_grab(
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
            self._log_info("用户已取消操作。")
            return

        confirmed = self._ask_choice_no_grab(
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
            self._log_info("用户已取消操作。")
            return

        self.btn_merge_workbooks.config(state="disabled")
        try:
            self._log_info(f"源文件数量：{len(source_paths)}")
            self._log_info(f"目标工作簿：{target_path}")
            result = merge_workbooks_to_existing_workbook(
                source_paths=list(source_paths),
                target_path=target_path,
                requested_sheet_name=requested_sheet_name,
                values_only=values_only,
                logger=self._flushing_logger(),
            )
            self._log_info(
                "多簿到一簿完成："
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个；"
                f"失败 {result['failed_count']} 个。"
            )
            self._show_info_no_grab(
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
            self._log_error(f"多簿到一簿失败：{e}")
            self._show_info_no_grab(
                "多簿到一簿",
                f"多簿到一簿失败：{e}",
                dialog_width=560,
                wraplength=500,
            )
        finally:
            self.btn_merge_workbooks.config(state="normal")

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

        self.btn_rename_sheets.config(state="disabled")
        try:
            self.logger.info(f"已选择文件数量：{len(source_paths)}")
            workbook_infos = self._read_sheet_infos_from_workbook_files(list(source_paths))
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
            self._show_info_no_grab(
                "批量重命名工作表",
                f"批量重命名工作表失败：{e}",
            )
        finally:
            self.btn_rename_sheets.config(state="normal")

    def _read_sheet_infos_from_workbook_files(self, source_paths):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        excel = None
        workbook = None
        workbook_infos = []

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for source_path in source_paths:
                workbook = None
                abs_path = os.path.abspath(source_path)
                workbook_info = {
                    "workbook_path": abs_path,
                    "workbook_name": os.path.basename(abs_path),
                    "sheet_infos": [],
                }
                if is_office_temp_file(abs_path):
                    workbook_info["error"] = "临时文件已跳过"
                    workbook_infos.append(workbook_info)
                    continue
                if not is_excel_workbook_file(abs_path):
                    workbook_info["error"] = "不是支持的 Excel 文件"
                    workbook_infos.append(workbook_info)
                    continue
                if not os.path.exists(abs_path):
                    workbook_info["error"] = "原文件不存在"
                    workbook_infos.append(workbook_info)
                    continue

                try:
                    self.logger.info(f"正在读取工作表清单：{abs_path}")
                    workbook = excel.Workbooks.Open(
                        abs_path,
                        ReadOnly=True,
                        UpdateLinks=0,
                    )
                    workbook_info["workbook_name"] = workbook.Name
                    workbook_info["sheet_infos"] = self._read_workbook_sheet_infos(workbook)
                except Exception as e:
                    workbook_info["error"] = f"文件打开失败：{e}"
                finally:
                    if workbook is not None:
                        try:
                            workbook.Close(SaveChanges=False)
                        except Exception:
                            pass
                        workbook = None

                workbook_infos.append(workbook_info)

            return workbook_infos

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

    def _read_workbook_sheet_infos(self, workbook):
        sheet_infos = []
        for index, sheet in enumerate(workbook.Worksheets, start=1):
            sheet_infos.append(
                {
                    "name": sheet.Name,
                    "is_hidden": is_sheet_hidden_by_visible_value(sheet.Visible),
                    "order": index,
                }
            )
        return sheet_infos

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

    def run_batch_rename_files(self):
        """按钮回调函数，生成临时规则表后确认执行批量重命名文件"""
        self.logger.info("批量重命名文件：开始操作。")
        source_paths = filedialog.askopenfilenames(
            title="请选择需要批量重命名的文件",
            filetypes=[("所有文件", "*.*")],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.btn_rename_files.config(state="disabled")
        try:
            self.logger.info(f"已选择文件数量：{len(source_paths)}")
            rule_path = create_rename_rule_workbook(list(source_paths))
            self.logger.info(f"已生成重命名规则表：{rule_path}")
            try:
                os.startfile(rule_path)
                self.logger.info("规则表已打开，请填写 C/D 列，保存并关闭规则表后点击执行。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")
                self.logger.info(f"请手动打开规则表：{rule_path}")

            if not self._confirm_rename_rule_ready(rule_path):
                self.logger.info("用户已取消操作")
        except Exception as e:
            self.logger.error(f"批量重命名文件失败：{e}")
            self._show_info_no_grab(
                "批量重命名文件",
                f"批量重命名文件失败：{e}",
            )
        finally:
            self.btn_rename_files.config(state="normal")

    def run_color_sum(self):
        """按钮回调函数，按当前选中单元格填充色汇总指定 sheet 的同地址单元格"""
        self._log_info("按颜色汇总求和：开始操作。")
        try:
            sum_scope = self._ask_choice_no_grab(
                "按颜色汇总求和",
                "请选择汇总范围：\n1 仅汇总一个 sheet\n2 汇总所有匹配 sheet",
                [
                    ("仅汇总一个 sheet", "single"),
                    ("汇总所有匹配 sheet", "all"),
                ],
                dialog_width=440,
                wraplength=380,
            )
            if sum_scope is None:
                self._log_info("用户已取消操作。")
                return

            target_sheet_name = None
            if sum_scope == "single":
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
            if sum_scope == "single":
                result = sum_current_sheet_by_fill_color(
                    source_paths=list(source_paths),
                    write_mode=write_mode,
                    target_sheet_name=target_sheet_name,
                    logger=self._flushing_logger(),
                )
                self._log_single_color_sum_result(result)
                self._show_single_color_sum_result(result)
            else:
                result = sum_matching_sheets_by_fill_color(
                    source_paths=list(source_paths),
                    write_mode=write_mode,
                    logger=self._flushing_logger(),
                )
                self._log_all_color_sum_result(result)
                self._show_all_color_sum_result(result)
        except Exception as e:
            self._log_error(f"按颜色汇总求和失败：{type(e).__name__}: {e}")
        finally:
            self.btn_color_sum.config(state="normal")

    def run_data_drill(self):
        """按钮回调函数，按当前活动 sheet 和选中单元格批量穿透读取源工作簿取值"""
        self._log_info("数据穿透取数：开始操作。")
        try:
            context = self._get_active_excel_drill_context()
            if not self._confirm_data_drill_context(context):
                self._log_info("用户已取消操作。")
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
                self._log_info("用户已取消操作。")
                return

            self.btn_data_drill.config(state="disabled")
            output_path = build_unique_output_path(context["output_dir"])
            self._log_info(
                f"当前取数点：{context['sheet_name']}!{context['cell_address']}；"
                f"源文件数量：{len(source_paths)}。"
            )
            records = build_data_drill_records(
                source_paths=list(source_paths),
                sheet_name=context["sheet_name"],
                cell_address=context["cell_address"],
                logger=self._flushing_logger(),
            )
            write_data_drill_result_workbook(
                records=records,
                output_path=output_path,
                source_sheet_name=context["sheet_name"],
                cell_address=context["cell_address"],
            )
            summary = summarize_data_drill_records(records)

            try:
                os.startfile(output_path)
                self._log_info(f"结果文件已生成并尝试打开：{output_path}")
            except Exception as open_error:
                self._log_error(f"结果文件已生成，但自动打开失败：{open_error}")
                self._log_info(f"请手动打开结果文件：{output_path}")

            self._log_info(
                "数据穿透取数完成："
                f"成功 {summary['success_count']} 个；"
                f"跳过 {summary['skipped_count']} 个；"
                f"失败 {summary['failed_count']} 个。"
            )
            self._log_info("当前活动工作簿未被修改、未被保存。")
            self._show_info_no_grab(
                "数据穿透取数",
                "数据穿透取数完成。\n"
                f"成功数量：{summary['success_count']}\n"
                f"跳过数量：{summary['skipped_count']}\n"
                f"失败数量：{summary['failed_count']}\n"
                f"结果文件路径：{output_path}\n\n"
                "当前活动工作簿未被修改、未被保存。",
                dialog_width=560,
                wraplength=500,
            )
        except Exception as e:
            self._log_error(f"数据穿透取数失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "数据穿透取数",
                f"数据穿透取数失败：{e}",
                dialog_width=520,
                wraplength=460,
            )
        finally:
            self.btn_data_drill.config(state="normal")

    def _confirm_data_drill_context(self, context):
        choice = self._ask_choice_no_grab(
            "数据穿透取数",
            "请确认当前取数点：\n"
            f"当前工作簿名：{context['workbook_name']}\n"
            f"当前工作簿路径：{context['workbook_path']}\n"
            f"当前 Sheet 名：{context['sheet_name']}\n"
            f"当前单元格地址：{context['cell_address']}\n"
            f"即将读取各源文件中的：{context['sheet_name']}!{context['cell_address']}\n\n"
            "结果文件将生成到当前工作簿同目录。\n"
            "当前工作簿不会被修改或保存。",
            [("继续选择源文件", "continue")],
            dialog_width=560,
            wraplength=500,
        )
        return choice == "continue"

    def _get_current_excel_context_for_data_drill(self):
        return self._get_active_excel_drill_context()

    def _get_active_excel_drill_context(self):
        try:
            import pythoncom
            import win32com.client
        except Exception as e:
            raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

        pythoncom.CoInitialize()
        steps = []
        try:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception as e:
                raise RuntimeError(
                    "未检测到正在运行的 Excel，请先打开目标/合并工作簿后重试。"
                    "当前步骤：尚未连接 Excel。"
                ) from e
            steps.append("已连接 Excel")

            try:
                workbook = excel.ActiveWorkbook
            except Exception as e:
                raise RuntimeError(
                    "Excel 中没有活动工作簿，请先打开目标/合并工作簿后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                ) from e
            if workbook is None:
                raise RuntimeError(
                    "Excel 中没有活动工作簿，请先打开目标/合并工作簿后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                )
            steps.append("已取得 ActiveWorkbook")

            workbook_name = str(workbook.Name)
            workbook_dir = str(workbook.Path or "").strip()
            if not workbook_dir:
                raise RuntimeError(
                    "当前活动工作簿尚未保存，无法确定结果文件输出目录。请先保存当前工作簿后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                )
            workbook_path = str(workbook.FullName or "").strip()
            if not os.path.dirname(workbook_path):
                workbook_path = os.path.join(workbook_dir, workbook_name)
            steps.append("已取得工作簿保存路径")

            try:
                active_sheet = excel.ActiveSheet
                sheet_name = str(active_sheet.Name)
            except Exception as e:
                raise RuntimeError(
                    "无法读取当前活动 Sheet，请先切换到目标工作表后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                ) from e
            if not sheet_name:
                raise RuntimeError(
                    "无法读取当前活动 Sheet，请先切换到目标工作表后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                )
            steps.append("已取得 ActiveSheet")

            cell_address = self._get_active_cell_address_for_data_drill(excel)
            if not cell_address:
                raise RuntimeError(
                    "没有有效选区，请先选中一个需要追查的单元格后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}，但未取得 ActiveCell 或可用 Selection。"
                )
            steps.append("已取得 ActiveCell")

            return {
                "workbook_name": workbook_name,
                "workbook_path": workbook_path,
                "output_dir": workbook_dir,
                "workbook_dir": workbook_dir,
                "sheet_name": sheet_name,
                "cell_address": cell_address,
            }
        finally:
            pythoncom.CoUninitialize()

    def _get_active_cell_address_for_data_drill(self, excel):
        active_cell_error = None
        try:
            active_cell = excel.ActiveCell
            cell_address = self._get_cell_address_for_data_drill(active_cell)
            if cell_address:
                return cell_address
        except Exception as e:
            active_cell_error = e

        try:
            selection = excel.Selection
        except Exception as e:
            detail = f"已连接 Excel，但未取得 ActiveCell，也无法读取 Selection：{e}"
            if active_cell_error is not None:
                detail = f"已连接 Excel，但未取得 ActiveCell：{active_cell_error}；无法读取 Selection：{e}"
            raise RuntimeError(detail) from e

        cell_address = self._get_selection_cell_address_for_data_drill(selection)
        if cell_address:
            return cell_address

        detail = "已连接 Excel，但未取得 ActiveCell；Selection 不是可识别的单元格区域。"
        if active_cell_error is not None:
            detail = f"已连接 Excel，但未取得 ActiveCell：{active_cell_error}；Selection 不是可识别的单元格区域。"
        raise RuntimeError(detail)

    def _get_selection_cell_address_for_data_drill(self, selection):
        if selection is None:
            return None

        cell_candidates = [
            lambda: selection.Cells(1, 1),
            lambda: selection.Cells.Item(1, 1),
            lambda: selection.Areas(1).Cells(1, 1),
            lambda: selection.Item(1, 1),
            lambda: selection,
        ]
        for get_cell in cell_candidates:
            try:
                cell_address = self._get_cell_address_for_data_drill(get_cell())
                if cell_address:
                    return self._left_top_cell_address(cell_address)
            except Exception:
                continue
        return None

    def _get_cell_address_for_data_drill(self, cell):
        if cell is None:
            return None
        try:
            address_member = cell.Address
            if callable(address_member):
                address = address_member(False, False)
            else:
                address = address_member
        except Exception:
            address = cell.Address(False, False)
        return self._left_top_cell_address(address)

    def _left_top_cell_address(self, address):
        first_area = str(address or "").split(",")[0].strip()
        return first_area.split(":")[0].replace("$", "").strip()

    def _format_data_drill_context_steps(self, steps):
        return "、".join(steps) if steps else "尚未连接 Excel"

    def _log_single_color_sum_result(self, result):
        self._log_info(
            "按颜色汇总求和完成："
            f"写入 {result['written_cell_count']} 个单元格；"
            f"找到同名工作表源文件 {result['matched_source_file_count']} 个；"
            f"缺少同名工作表 {result['missing_sheet_file_count']} 个；"
            f"忽略非数字 {result['ignored_non_numeric_count']} 个。"
        )
        self._log_info("目标工作簿未自动保存，请检查后自行保存。")

    def _show_single_color_sum_result(self, result):
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

    def _log_all_color_sum_result(self, result):
        self._log_info(
            "按颜色汇总求和完成："
            f"目标计划匹配 {result['matched_sheet_count']} 个工作表；"
            f"实际写入 {result['written_sheet_count']} 个工作表、{result['written_cell_count']} 个单元格；"
            f"参与源文件 {result['matched_source_file_count']} 个；"
            f"匹配源工作表 {result['matched_source_sheet_count']} 个；"
            f"缺少同名工作表 {result['missing_sheet_file_count']} 个；"
            f"忽略非数字 {result['ignored_non_numeric_count']} 个。"
        )
        self._log_info("目标工作簿未自动保存，请检查后自行保存。")

    def _show_all_color_sum_result(self, result):
        self._show_info_no_grab(
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
        )

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

    def _confirm_sheet_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title("批量重命名工作表")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text=(
                "规则表已打开。\n"
                "请在【重命名清单】D列填写新工作表名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            font=("Microsoft YaHei", 11),
            wraplength=400,
            justify="left",
            padx=16,
            pady=12,
        ).pack(anchor="w")

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)

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
                self._show_info_no_grab("批量重命名工作表", result_message)
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名工作表失败：{e}")
                self._show_info_no_grab(
                    "批量重命名工作表",
                    f"批量重命名工作表失败：{e}",
                )
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
            text="我已填好规则，执行批量重命名工作表",
            command=execute,
            font=("Microsoft YaHei", 10),
            width=32,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, min_width=500, min_height=170)
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
        self._show_info_no_grab(
            "批量重命名工作表",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行批量重命名工作表。",
        )

    def _confirm_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title("批量重命名文件")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        tk.Label(
            dialog,
            text=(
                "规则表已打开。\n"
                "请在【重命名清单】中填写新文件名和后缀名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            font=("Microsoft YaHei", 11),
            wraplength=400,
            justify="left",
            padx=16,
            pady=12,
        ).pack(anchor="w")

        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 14), fill=tk.X)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info(f"正在读取重命名规则表：{rule_workbook_path}")
                try:
                    write_rename_results_to_workbook(rule_workbook_path, [])
                except PermissionError:
                    self._show_rule_workbook_busy_message()
                    return
                except OSError as e:
                    self.logger.error(f"规则表无法回写：{e}")
                    self._show_rule_workbook_busy_message()
                    return

                settings = read_rename_settings(rule_workbook_path)
                for warning in settings.warnings or []:
                    self.logger.info(warning)

                rules = read_rename_rules(rule_workbook_path)
                actions = build_rename_plan(rules, settings)
                valid_actions = [action for action in actions if action.status == "成功"]
                skipped_actions = [action for action in actions if action.status == "跳过"]
                self.logger.info(f"规则行数：{len(rules)}")
                self.logger.info(f"有效重命名任务：{len(valid_actions)}")
                self.logger.info(f"跳过任务：{len(skipped_actions)}")

                for index, action in enumerate(valid_actions, start=1):
                    self.logger.info(
                        f"正在重命名 {index}/{len(valid_actions)}："
                        f"{os.path.basename(action.original_path)} -> {action.final_name}"
                    )

                summary = execute_rename_plan(actions)
                try:
                    write_rename_results_to_workbook(rule_workbook_path, actions)
                    self.logger.info(f"规则表已更新：{rule_workbook_path}")
                    result_message = (
                        "重命名完成。\n"
                        f"成功：{summary['success_count']} 个\n"
                        f"跳过：{summary['skipped_count']} 个\n"
                        f"失败：{summary['failed_count']} 个\n\n"
                        "规则表已更新状态/备注，请查看确认。"
                    )
                except PermissionError:
                    self.logger.error("规则表被占用，无法回写状态、说明和处理日志。")
                    result_message = (
                        "重命名已执行，但规则表无法回写。\n"
                        "请确认规则表已保存并关闭后，再查看文件结果。"
                    )
                except OSError as e:
                    self.logger.error(f"规则表无法回写：{e}")
                    result_message = (
                        "重命名已执行，但规则表无法回写。\n"
                        f"原因：{e}"
                    )

                self.logger.info(
                    "批量重命名完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self._show_info_no_grab("批量重命名文件", result_message)
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名文件失败：{e}")
                self._show_info_no_grab(
                    "批量重命名文件",
                    f"批量重命名文件失败：{e}",
                )
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
            text="我已填好规则，执行批量重命名",
            command=execute,
            font=("Microsoft YaHei", 10),
            width=28,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, min_width=460, min_height=170)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self._show_info_no_grab(
            "批量重命名文件",
            "规则表仍被 Excel 占用，无法回写。\n"
            "请先保存并关闭规则表，再点击执行批量重命名。",
        )


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
