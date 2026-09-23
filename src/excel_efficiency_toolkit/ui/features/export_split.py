# -*- coding: utf-8 -*-
"""拆分导出：按工作表拆分文件、按列拆分、按行拆分。"""
import os
import threading
from tkinter import filedialog

from ...export_ops import export_workbook_sheets_to_files
from ...table_ops import (
    CancellationToken,
    _get_last_used_row_and_col,
    _get_or_open_workbook,
    _read_range_values,
    build_split_targets,
    get_split_preview_info,
    parse_column_index,
    read_sheet_header_columns,
    release_workbook_session,
    resolve_source_sheet_name,
    split_workbook_sheet_by_column,
    split_workbook_sheet_by_column_to_files,
    split_workbook_sheet_by_rows,
    split_workbook_sheet_by_rows_to_files,
    validate_row_numbers,
    validate_rows_per_part,
)
from .base import FeatureController



class ExportSplitController(FeatureController):
    def _run_worker(self, operation, on_success, on_error, on_finally):
        """在真实 UI 中后台执行 COM；无 root 的测试上下文保持同步。"""
        if self.root is None:
            try:
                on_success(operation())
            except Exception as exc:
                on_error(exc)
            finally:
                on_finally()
            return

        def worker():
            pythoncom_module = None
            result = None
            error = None
            try:
                import pythoncom as pythoncom_module
                pythoncom_module.CoInitialize()
                result = operation()
            except Exception as exc:
                error = exc
            finally:
                if pythoncom_module is not None:
                    try:
                        pythoncom_module.CoUninitialize()
                    except Exception as exc:
                        if error is None:
                            error = exc
            if error is None:
                self.root.after(0, lambda result=result: on_success(result))
            else:
                self.root.after(0, lambda error=error: on_error(error))
            self.root.after(0, on_finally)

        threading.Thread(target=worker, name="excel-split-worker", daemon=True).start()

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
        def operation():
            self.logger.info(f"源文件：{source_path}")
            self.logger.info(f"输出目录：{output_dir}")
            return export_workbook_sheets_to_files(source_path, output_dir, self.logger)

        def on_success(exported_paths):
            self.logger.info(f"成功导出 {len(exported_paths)} 个文件。")

        def on_error(e):
            self.logger.error(f"导出失败：请确认文件未损坏、已安装 Microsoft Excel，并且输出目录可写。详细信息：{e}")

        self._run_worker(
            operation,
            on_success,
            on_error,
            lambda: self.button("btn_export_sheets").config(state="normal"),
        )

    def _start_column_split_execution(
        self,
        source_path,
        source_sheet_name,
        column_input,
        output_mode,
        output_dir,
        header_row,
        data_start_row,
    ):
        """在 UI 线程只创建控件，实际 COM 拆分在 worker 中执行。"""
        self.button("btn_split_sheet").config(state="disabled")
        cancel_token = CancellationToken()
        progress_dialog = None
        if hasattr(self.dialogs, "create_progress_cancel_dialog"):
            progress_dialog = self.create_progress_cancel_dialog(
                title="拆分进行中",
                initial_message="正在准备拆分任务...",
                on_cancel=cancel_token.cancel,
            )

        def progress_cb(current, total, name):
            if progress_dialog:
                message = f"正在处理 ({current}/{total})：{name}"
                self.root.after(0, lambda: progress_dialog.update_message(message))

        def operation():
            self.logger.info(f"源工作簿：{source_path}")
            if output_mode == "files":
                self.logger.info(f"输出目录：{output_dir}")
                return split_workbook_sheet_by_column_to_files(
                    source_path=source_path,
                    source_sheet_name=source_sheet_name,
                    column_input=column_input,
                    output_dir=output_dir,
                    header_row=header_row,
                    data_start_row=data_start_row,
                    logger=self.logger,
                    cancel_token=cancel_token,
                    progress_callback=progress_cb,
                )
            return split_workbook_sheet_by_column(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                column_input=column_input,
                header_row=header_row,
                data_start_row=data_start_row,
                logger=self.logger,
                cancel_token=cancel_token,
                progress_callback=progress_cb,
            )

        def on_success(result):
            if result.get("cancelled"):
                self.logger.info("拆分任务已由用户取消。")
                return
            self.logger.info(f"工作簿名：{result['workbook_name']}")
            self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
            if output_mode == "files":
                self.logger.info(f"已生成 {result['created_file_count']} 个 Excel 文件。")
                self.logger.info(f"输出目录：{result['output_dir']}")
            else:
                self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
                self.logger.info(f"复制行数：{result['copied_row_count']}")

        def on_error(exc):
            self.logger.error(f"按列拆分失败：{exc}")

        self._run_worker(
            operation,
            on_success,
            on_error,
            lambda: self._finish_split_ui(progress_dialog, "btn_split_sheet"),
        )

    def _continue_column_advanced(self, source_path, source_sheet_name, header_row, data_start_row, columns):
        column_input = self.ask_column(
            "按列拆分 - 选择拆分列",
            "请选择或输入要作为拆分依据的列：",
            columns=columns,
            default="",
            allow_empty=False,
        )
        if column_input is None:
            self.logger.info("用户已取消操作")
            return
        try:
            parse_column_index(column_input)
        except ValueError as exc:
            self.logger.error(f"输入无效：{exc}")
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

        output_dir = None
        if output_mode == "files":
            output_dir = filedialog.askdirectory(title="请选择拆分文件的输出目录")
            if not output_dir:
                self.logger.info("用户已取消操作")
                return

        def show_preview(preview_info):
            preview_items = [
                ("源 Sheet 名", preview_info["source_sheet_name"]),
                ("数据行数量", preview_info["data_row_count"]),
                ("拆分方式", preview_info["split_mode_desc"]),
                ("命名依据", preview_info["name_column_desc"]),
                ("预计生成", preview_info["estimated_count_desc"]),
            ]
            if not self.ask_split_preview("按列拆分 - 执行前确认", preview_items):
                self.logger.info("用户已取消操作")
                return
            self._start_column_split_execution(
                source_path,
                source_sheet_name,
                column_input,
                output_mode,
                output_dir,
                header_row,
                data_start_row,
            )

        self._run_worker(
            lambda: get_split_preview_info(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                header_row=header_row,
                data_start_row=data_start_row,
                split_mode="column",
                split_arg=column_input,
                output_mode=output_mode,
                logger=self.logger,
            ),
            show_preview,
            lambda exc: self.logger.error(f"读取拆分预览失败：{exc}"),
            lambda: None,
        )

    def _start_column_advanced(self, source_path, source_sheet_name, header_row, data_start_row):
        self._run_worker(
            lambda: read_sheet_header_columns(
                sheet_or_path=source_path,
                header_row=header_row,
                source_sheet_name=source_sheet_name,
                logger=self.logger,
            ),
            lambda columns: self._continue_column_advanced(
                source_path, source_sheet_name, header_row, data_start_row, columns
            ),
            lambda exc: (
                self.logger.error(f"读取表头列信息失败：{exc}"),
                self._continue_column_advanced(
                    source_path, source_sheet_name, header_row, data_start_row, []
                ),
            ),
            lambda: None,
        )

    def _continue_rows_advanced(
        self, source_path, source_sheet_name, output_mode, output_dir,
        rows_per_part, header_row, data_start_row, columns,
    ):
        name_column = None
        if output_mode == "files":
            name_column_input = self.ask_column(
                "按行拆分 - 文件命名列",
                "请选择文件名来源列（可选，留空或跳过则使用默认编号命名）：",
                columns=columns,
                default="",
                allow_empty=True,
            )
            if name_column_input is None:
                self.logger.info("用户已取消操作")
                return
            name_column = name_column_input.strip().upper() if name_column_input.strip() else None
            if name_column:
                try:
                    parse_column_index(name_column)
                except ValueError as exc:
                    self.logger.error(f"输入无效：{exc}")
                    return

        def show_preview(preview_info):
            preview_items = [
                ("源 Sheet 名", preview_info["source_sheet_name"]),
                ("数据行数量", preview_info["data_row_count"]),
                ("拆分方式", preview_info["split_mode_desc"]),
                ("命名列", preview_info["name_column_desc"]),
                ("预计生成", preview_info["estimated_count_desc"]),
            ]
            if not self.ask_split_preview("按行拆分 - 执行前确认", preview_items):
                self.logger.info("用户已取消操作")
                return
            cancel_token = CancellationToken()
            progress_dialog = self.create_progress_cancel_dialog(
                title="拆分进行中",
                initial_message="正在准备拆分任务...",
                on_cancel=cancel_token.cancel,
            ) if hasattr(self.dialogs, "create_progress_cancel_dialog") else None
            self._run_rows_worker(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                output_mode=output_mode,
                output_dir=output_dir,
                rows_per_part=rows_per_part,
                header_row=header_row,
                data_start_row=data_start_row,
                name_column=name_column,
                cancel_token=cancel_token,
                progress_dialog=progress_dialog,
            )

        # 预览 COM 也在后台运行；真正执行阶段由 _run_rows_worker 建立取消对话框。
        self._run_worker(
            lambda: get_split_preview_info(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                header_row=header_row,
                data_start_row=data_start_row,
                split_mode="rows",
                split_arg=rows_per_part,
                output_mode=output_mode,
                name_column=name_column,
                logger=self.logger,
            ),
            show_preview,
            lambda exc: self.logger.error(f"读取拆分预览失败：{exc}"),
            lambda: None,
        )

    def _start_rows_advanced(
        self, source_path, source_sheet_name, output_mode, output_dir,
        rows_per_part, header_row, data_start_row,
    ):
        def continue_with_columns(columns):
            self._continue_rows_advanced(
                source_path,
                source_sheet_name,
                output_mode,
                output_dir,
                rows_per_part,
                header_row,
                data_start_row,
                columns,
            )

        if output_mode != "files":
            continue_with_columns([])
            return

        self._run_worker(
            lambda: read_sheet_header_columns(
                sheet_or_path=source_path,
                header_row=header_row,
                source_sheet_name=source_sheet_name,
                logger=self.logger,
            ),
            continue_with_columns,
            lambda exc: (
                self.logger.error(f"读取表头列信息失败：{exc}"),
                continue_with_columns([]),
            ),
            lambda: None,
        )

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

        use_advanced_ui = (
            hasattr(self.dialogs, "ask_column")
            and hasattr(self.dialogs, "ask_split_preview")
        )

        if not use_advanced_ui:
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
        else:
            try:
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

            if self.root is not None:
                self._start_column_advanced(
                    source_path, source_sheet_name, header_row, data_start_row
                )
                return

            try:
                columns = read_sheet_header_columns(
                    sheet_or_path=source_path,
                    header_row=header_row,
                    source_sheet_name=source_sheet_name,
                    logger=self.logger,
                )
            except Exception as e:
                self.logger.error(f"读取表头列信息失败：{e}")
                columns = []

            column_input = self.ask_column(
                "按列拆分 - 选择拆分列",
                "请选择或输入要作为拆分依据的列：",
                columns=columns,
                default="",
                allow_empty=False,
            )
            if column_input is None:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

            try:
                parse_column_index(column_input)
            except ValueError as e:
                self.logger.error(f"输入无效：{e}")
                release_workbook_session(source_path, self.logger)
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
            release_workbook_session(source_path, self.logger)
            return

        if output_mode == "files":
            output_dir = filedialog.askdirectory(title="请选择拆分文件的输出目录")
            if not output_dir:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

        if use_advanced_ui:
            preview_info = get_split_preview_info(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                header_row=header_row,
                data_start_row=data_start_row,
                split_mode="column",
                split_arg=column_input,
                output_mode=output_mode,
                logger=self.logger,
            )
            preview_items = [
                ("源 Sheet 名", preview_info["source_sheet_name"]),
                ("数据行数量", preview_info["data_row_count"]),
                ("拆分方式", preview_info["split_mode_desc"]),
                ("命名依据", preview_info["name_column_desc"]),
                ("预计生成", preview_info["estimated_count_desc"]),
            ]
            confirmed = self.ask_split_preview(
                "按列拆分 - 执行前确认",
                preview_items,
            )
            if not confirmed:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

        self.button("btn_split_sheet").config(state="disabled")
        cancel_token = CancellationToken() if use_advanced_ui else None
        progress_dialog = None
        if use_advanced_ui and hasattr(self.dialogs, "create_progress_cancel_dialog"):
            progress_dialog = self.create_progress_cancel_dialog(
                title="拆分进行中",
                initial_message="正在准备拆分任务...",
                on_cancel=cancel_token.cancel,
            )

        def progress_cb(current, total, name):
            if progress_dialog:
                message = f"正在处理 ({current}/{total})：{name}"
                if self.root is None:
                    progress_dialog.update_message(message)
                else:
                    self.root.after(0, lambda: progress_dialog.update_message(message))

        def operation():
            self.logger.info(f"源工作簿：{source_path}")
            if output_mode == "files":
                self.logger.info(f"输出目录：{output_dir}")
                kwargs = {
                    "source_path": source_path,
                    "source_sheet_name": source_sheet_name,
                    "column_input": column_input,
                    "output_dir": output_dir,
                    "header_row": header_row,
                    "data_start_row": data_start_row,
                    "logger": self.logger,
                }
                if use_advanced_ui:
                    kwargs["cancel_token"] = cancel_token
                    kwargs["progress_callback"] = progress_cb

                return split_workbook_sheet_by_column_to_files(**kwargs)
            else:
                kwargs = {
                    "source_path": source_path,
                    "source_sheet_name": source_sheet_name,
                    "column_input": column_input,
                    "header_row": header_row,
                    "data_start_row": data_start_row,
                    "logger": self.logger,
                }
                if use_advanced_ui:
                    kwargs["cancel_token"] = cancel_token
                    kwargs["progress_callback"] = progress_cb

                return split_workbook_sheet_by_column(**kwargs)

        def on_success(result):
            if result.get("cancelled"):
                self.logger.info("拆分任务已由用户取消。")
                return
            self.logger.info(f"工作簿名：{result['workbook_name']}")
            self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
            if output_mode == "files":
                self.logger.info(f"已生成 {result['created_file_count']} 个 Excel 文件。")
                self.logger.info(f"输出目录：{result['output_dir']}")
            else:
                self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
                self.logger.info(f"复制行数：{result['copied_row_count']}")

        def on_error(e):
            self.logger.error(f"按列拆分失败：{e}")

        self._run_worker(operation, on_success, on_error, lambda: self._finish_split_ui(progress_dialog, "btn_split_sheet"))

    def _finish_split_ui(self, progress_dialog, button_name):
        if progress_dialog:
            progress_dialog.close()
        self.button(button_name).config(state="normal")

    def _run_rows_operation(self, source_path, source_sheet_name, output_mode, output_dir,
                            rows_per_part, header_row, data_start_row, name_column,
                            cancel_token, progress_cb):
        kwargs = {
            "source_path": source_path,
            "source_sheet_name": source_sheet_name,
            "rows_per_part": rows_per_part,
            "header_row": header_row,
            "data_start_row": data_start_row,
            "logger": self.logger,
        }
        if output_mode == "files":
            kwargs["output_dir"] = output_dir
            if name_column is not None:
                kwargs["name_column"] = name_column
            target = split_workbook_sheet_by_rows_to_files
        else:
            if name_column is not None:
                kwargs["name_column"] = name_column
            target = split_workbook_sheet_by_rows
        if cancel_token is not None:
            kwargs["cancel_token"] = cancel_token
            kwargs["progress_callback"] = progress_cb
        return target(**kwargs)

    def _log_rows_success(self, result, output_mode, split_mode_desc):
        if result.get("cancelled"):
            self.logger.info("拆分任务已由用户取消。")
            return
        self.logger.info(f"工作簿名：{result['workbook_name']}")
        self.logger.info(f"源 sheet 名：{result['source_sheet_name']}")
        self.logger.info(f"拆分方式：{split_mode_desc}")
        if output_mode == "files":
            self.logger.info(f"成功生成 {result['created_file_count']} 个 Excel 文件。")
            self.logger.info(f"输出目录：{result['output_dir']}")
        else:
            self.logger.info(f"生成 sheet 数：{result['created_sheet_count']}")
            self.logger.info(f"复制行数：{result['copied_row_count']}")

    def _run_rows_worker(self, source_path, source_sheet_name, output_mode, output_dir,
                         rows_per_part, header_row, data_start_row, name_column,
                         cancel_token, progress_dialog):
        def progress_cb(current, total, name):
            if progress_dialog:
                message = f"正在处理 ({current}/{total})：{name}"
                if self.root is None:
                    progress_dialog.update_message(message)
                else:
                    self.root.after(0, lambda: progress_dialog.update_message(message))

        operation = lambda: self._run_rows_operation(
            source_path, source_sheet_name, output_mode, output_dir,
            rows_per_part, header_row, data_start_row, name_column,
            cancel_token, progress_cb,
        )
        on_success = lambda result: self._log_rows_success(
            result, output_mode,
            "每行一份" if rows_per_part == 1 else f"每 {rows_per_part} 行一份",
        )
        def on_error(exc):
            self.logger.error(f"按行拆分失败：{exc}")
        self._run_worker(operation, on_success, on_error,
                         lambda: self._finish_split_ui(progress_dialog, "btn_split_rows"))

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

        use_advanced_ui = (
            hasattr(self.dialogs, "ask_column")
            and hasattr(self.dialogs, "ask_split_preview")
        )

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
            release_workbook_session(source_path, self.logger)
            return

        if output_mode == "files":
            output_dir = filedialog.askdirectory(title="请选择拆分文件的输出目录")
            if not output_dir:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

        name_column = None
        if use_advanced_ui and self.root is not None:
            self._start_rows_advanced(
                source_path,
                source_sheet_name,
                output_mode,
                output_dir if output_mode == "files" else None,
                rows_per_part,
                header_row,
                data_start_row,
            )
            return

        if use_advanced_ui and output_mode == "files":
            try:
                columns = read_sheet_header_columns(
                    sheet_or_path=source_path,
                    header_row=header_row,
                    source_sheet_name=source_sheet_name,
                    logger=self.logger,
                )
            except Exception as e:
                self.logger.error(f"读取表头列信息失败：{e}")
                columns = []

            name_column_input = self.ask_column(
                "按行拆分 - 文件命名列",
                "请选择文件名来源列（可选，留空或跳过则使用默认编号命名）：",
                columns=columns,
                default="",
                allow_empty=True,
            )
            if name_column_input is None:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

            name_column = name_column_input.strip().upper() if name_column_input.strip() else None
            if name_column:
                try:
                    parse_column_index(name_column)
                except ValueError as e:
                    self.logger.error(f"输入无效：{e}")
                    release_workbook_session(source_path, self.logger)
                    return

        if use_advanced_ui:
            preview_info = get_split_preview_info(
                source_path=source_path,
                source_sheet_name=source_sheet_name,
                header_row=header_row,
                data_start_row=data_start_row,
                split_mode="rows",
                split_arg=rows_per_part,
                output_mode=output_mode,
                name_column=name_column,
                logger=self.logger,
            )
            preview_items = [
                ("源 Sheet 名", preview_info["source_sheet_name"]),
                ("数据行数量", preview_info["data_row_count"]),
                ("拆分方式", preview_info["split_mode_desc"]),
                ("命名列", preview_info["name_column_desc"]),
                ("预计生成", preview_info["estimated_count_desc"]),
            ]
            confirmed = self.ask_split_preview(
                "按行拆分 - 执行前确认",
                preview_items,
            )
            if not confirmed:
                self.logger.info("用户已取消操作")
                release_workbook_session(source_path, self.logger)
                return

        self.button("btn_split_rows").config(state="disabled")
        cancel_token = CancellationToken() if use_advanced_ui else None
        progress_dialog = None
        if use_advanced_ui and hasattr(self.dialogs, "create_progress_cancel_dialog"):
            progress_dialog = self.create_progress_cancel_dialog(
                title="拆分进行中",
                initial_message="正在准备拆分任务...",
                on_cancel=cancel_token.cancel,
            )

        self._run_rows_worker(
            source_path=source_path,
            source_sheet_name=source_sheet_name,
            output_mode=output_mode,
            output_dir=output_dir if output_mode == "files" else None,
            rows_per_part=rows_per_part,
            header_row=header_row,
            data_start_row=data_start_row,
            name_column=name_column if output_mode == "files" else None,
            cancel_token=cancel_token if use_advanced_ui else None,
            progress_dialog=progress_dialog,
        )
