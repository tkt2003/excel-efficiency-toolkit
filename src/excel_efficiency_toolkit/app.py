import os
import sys
import tkinter as tk
import customtkinter as ctk
from datetime import datetime
from tkinter import filedialog
from .clear_by_color_ops import (
    clear_multiple_workbooks_by_color,
    execute_clear_active_workbook_plan,
    plan_clear_active_workbook_by_color,
)
from .color_sum_ops import sum_current_sheet_by_fill_color, sum_matching_sheets_by_fill_color
from .data_drill_ops import (
    build_data_drill_range_records,
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
from .link_replace_ops import (
    execute_link_replacement_from_rule_workbook,
    generate_temporary_link_replace_rule_workbook,
)
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
from .report_generate_ops import (
    DEFAULT_CHECKLIST_SHEET_NAME,
    generate_reports_from_checklist,
)
from .round_formula_ops import round_selected_range_to_two_decimals
from .template_multi_link_ops import (
    create_template_multi_link_rule_workbook,
    execute_template_multi_link_generation_from_rule_workbook,
)
from .template_tb_report_ops import (
    DEFAULT_OUTPUT_NAME_SUFFIX,
    generate_reports_from_template_and_tb_files,
    read_template_external_links,
)
from .sheet_ops import generate_sheet_index_sheet_with_links
from .table_ops import (
    merge_workbook_sheets_to_new_sheet,
    parse_column_index,
    split_workbook_sheet_by_column,
    validate_row_numbers,
)
from .workbook_drill_ops import (
    RESULT_SHEET_BASE_NAME,
    build_unique_result_sheet_name,
    is_multi_area_range,
    normalize_range_address,
    should_skip_history_result_sheet,
    write_single_workbook_drill_result_to_com_sheet,
    write_single_file_result_sheet,
)
from .workbook_merge_ops import merge_workbooks_to_existing_workbook

APP_NAME = "老头表格助手"
APP_VERSION = "0.1.0"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class AppButton(ctk.CTkButton):
    def config(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        try:
            if not self.winfo_exists():
                return None
            return self.configure(**kwargs)
        except tk.TclError:
            return None

    configure = ctk.CTkButton.configure


class FeatureCardHandle:
    def __init__(self, set_state):
        self._set_state = set_state

    def config(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        return self.configure(**kwargs)

    def configure(self, **kwargs):
        state = kwargs.get("state")
        if state is not None:
            try:
                self._set_state(state)
            except tk.TclError:
                return None


class ExcelToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.root.minsize(1120, 700)
        self.bg_color = "#f3f6fa"
        self.card_color = "#ffffff"
        self.panel_color = "#eaf0f7"
        self.border_color = "#d8e0ea"
        self.text_color = "#1f2937"
        self.muted_text_color = "#64748b"
        self.subtle_text_color = "#94a3b8"
        self.accent_color = "#3b6ea8"
        self.accent_soft_color = "#dbeafe"
        self.card_hover_color = "#f1f6fb"
        self.card_disabled_color = "#f8fafc"
        self.header_subtitle_font = ("Microsoft YaHei", 11)
        self.header_button_font = ("Microsoft YaHei", 12)
        self.section_title_font = ("Microsoft YaHei", 15, "bold")
        self.card_title_font = ("Microsoft YaHei", 12, "bold")
        self.card_description_font = ("Microsoft YaHei", 11)
        self.log_font = ("Consolas", 12)
        self.feature_buttons = {}
        self.feature_cards = {}
        self.root.configure(fg_color=self.bg_color)

        self.main_frame = ctk.CTkFrame(root, fg_color=self.bg_color, corner_radius=0)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self._create_header(self.main_frame)
        self._create_feature_area(self.main_frame)
        self._create_log_area(self.main_frame)

        # 初始化自定义 logger
        self.logger = setup_logger(self.log_text)
        self.logger.info(f"欢迎使用 {APP_TITLE}。程序已就绪。")
        if os.environ.get("EXCEL_TOOLKIT_GUI_SMOKE") == "1":
            self.root.after(300, self._run_gui_smoke_and_exit)

    def _feature_groups(self):
        return [
            (
                "拆分导出",
                [
                    (
                        "按工作表拆分文件",
                        "把一个 Excel 文件按工作表拆成多个文件",
                        "btn_export_sheets",
                        self.run_export_sheets,
                    ),
                    (
                        "按指定列拆分工作表",
                        "按指定列的值拆分当前工作表",
                        "btn_split_sheet",
                        self.run_split_sheet,
                    ),
                ],
            ),
            (
                "合并整理",
                [
                    (
                        "多个工作表合并到一张表",
                        "把多个工作表追加合并到一个结果表",
                        "btn_merge_sheets",
                        self.run_merge_sheets,
                    ),
                    (
                        "多个 Excel 文件合并到一个文件",
                        "把多个 Excel 文件导入到一个目标文件",
                        "btn_merge_workbooks",
                        self.run_merge_workbooks,
                    ),
                ],
            ),
            (
                "模板生成 / 取数",
                [
                    (
                        "按颜色汇总求和",
                        "按填充色位置汇总多个文件中的数值",
                        "btn_color_sum",
                        self.run_color_sum,
                    ),
                    (
                        "按颜色清空内容",
                        "清空指定填充色单元格内容，保留格式",
                        "btn_clear_by_color",
                        self.run_clear_by_color,
                    ),
                    (
                        "数据穿透查询",
                        "按单元格或区域查询多个表或文件中的数据",
                        "btn_data_drill",
                        self.run_data_drill,
                    ),
                    (
                        "按模板批量生成 Excel",
                        "替换链接并生成多个 Excel 文件",
                        "btn_template_tb_report",
                        self.run_template_generate,
                    ),
                    (
                        "选区 ROUND 保留两位",
                        "套用 ROUND 公式保留两位小数",
                        "btn_round_formula",
                        self.run_round_formula,
                    ),
                ],
            ),
            (
                "目录与链接",
                [
                    (
                        "批量更换 Excel 链接",
                        "批量替换多个 Excel 文件中的外部链接",
                        "btn_link_replace",
                        self.run_batch_link_replace,
                    ),
                    (
                        "生成工作表目录",
                        "生成可点击跳转的工作表目录",
                        "btn_sheet_index",
                        self.run_sheet_index,
                    ),
                ],
            ),
            (
                "批量维护",
                [
                    ("批量删除工作表", "按规则表批量删除工作表", "btn_delete_sheets", self.run_delete_sheets),
                    ("批量重命名文件", "按规则表批量修改文件名", "btn_rename_files", self.run_batch_rename_files),
                    (
                        "批量重命名工作表",
                        "按规则表批量修改工作表名",
                        "btn_rename_sheets",
                        self.run_batch_rename_sheets,
                    ),
                ],
            ),
        ]

    def _create_header(self, parent):
        header = ctk.CTkFrame(
            parent,
            fg_color=self.card_color,
            corner_radius=12,
            border_width=1,
            border_color="#e5edf5",
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.grid_columnconfigure(0, weight=1)

        title_block = ctk.CTkFrame(header, fg_color="transparent")
        title_block.grid(row=0, column=0, sticky="ew", padx=18, pady=10)
        title_block.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_block,
            text=APP_TITLE,
            font=("Microsoft YaHei", 19, "bold"),
            text_color=self.text_color,
        ).grid(row=0, column=0, sticky="w")

        help_buttons = ctk.CTkFrame(title_block, fg_color="transparent")
        help_buttons.grid(row=0, column=1, rowspan=2, sticky="e")
        for index, (text, command) in enumerate((("使用说明", self.show_user_guide), ("关于", self.show_about))):
            AppButton(
                help_buttons,
                text=text,
                font=self.header_button_font,
                command=command,
                width=76,
                height=28,
                fg_color="#f8fafc",
                hover_color="#eef4fb",
                text_color=self.text_color,
                border_width=1,
                border_color="#dfe7f1",
                corner_radius=6,
            ).grid(row=0, column=index, padx=(8 if index else 0, 0))

        ctk.CTkLabel(
            title_block,
            text="Excel 效率工作台",
            font=self.header_subtitle_font,
            text_color="#475569",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

    def _create_feature_area(self, parent):
        feature_area = ctk.CTkScrollableFrame(
            parent,
            fg_color="transparent",
            scrollbar_button_color="#cbd5e1",
            scrollbar_button_hover_color="#94a3b8",
        )
        feature_area.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        feature_area.grid_columnconfigure(0, weight=1, uniform="feature_columns")
        feature_area.grid_columnconfigure(1, weight=1, uniform="feature_columns")

        groups = dict(self._feature_groups())
        left_sections = ("拆分导出", "合并整理", "批量维护")
        right_sections = ("模板生成 / 取数", "目录与链接")

        for column, section_names in enumerate((left_sections, right_sections)):
            column_frame = ctk.CTkFrame(feature_area, fg_color="transparent")
            column_frame.grid(row=0, column=column, sticky="new", padx=(0, 12) if column == 0 else (0, 0))
            column_frame.grid_columnconfigure(0, weight=1)

            for row, section_name in enumerate(section_names):
                panel = self._create_section_panel(column_frame, section_name, groups[section_name])
                panel.grid(row=row, column=0, sticky="ew", pady=(0, 8))

        feature_area.grid_rowconfigure(1, minsize=14)

    def _create_section_panel(self, parent, title, features):
        panel = ctk.CTkFrame(
            parent,
            fg_color=self.card_color,
            corner_radius=12,
            border_width=1,
            border_color="#e5edf5",
        )
        panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            panel,
            text=title,
            font=self.section_title_font,
            text_color=self.text_color,
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(9, 6))

        for row, (feature_title, description, attr_name, command) in enumerate(features, start=1):
            card = self._create_feature_card(panel, feature_title, description, attr_name, command)
            card.grid(row=row, column=0, sticky="ew", padx=10, pady=(0, 9))

        return panel

    def _create_feature_card(self, parent, title, description, attr_name, command):
        has_description = bool(description)
        card = ctk.CTkFrame(
            parent,
            height=70 if has_description else 50,
            fg_color="#fbfdff",
            border_width=1,
            border_color=self.border_color,
            corner_radius=10,
        )
        card.grid_propagate(False)
        card.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            card,
            text=title,
            font=self.card_title_font,
            text_color=self.text_color,
            anchor="w",
        )
        title_label.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(12, 8),
            pady=(8, 1) if has_description else (0, 0),
        )
        widgets = [card, title_label]
        desc_label = None
        if has_description:
            desc_label = ctk.CTkLabel(
                card,
                text=description,
                font=self.card_description_font,
                text_color="#475569",
                anchor="w",
                justify="left",
            )
            desc_label.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=(1, 10))
            widgets.append(desc_label)
        else:
            card.grid_rowconfigure(0, weight=1)

        def set_state(state):
            is_disabled = state == "disabled"
            card.configure(
                fg_color=self.card_disabled_color if is_disabled else "#fbfdff",
                border_color="#e5eaf0" if is_disabled else self.border_color,
            )
            title_label.configure(text_color=self.subtle_text_color if is_disabled else self.text_color)
            if desc_label is not None:
                desc_label.configure(text_color=self.subtle_text_color if is_disabled else "#475569")
            if is_disabled:
                card._feature_enabled = False
            else:
                card._feature_enabled = True

        card._feature_enabled = True
        handle = FeatureCardHandle(set_state)
        setattr(self, attr_name, handle)
        self.feature_buttons[attr_name] = handle
        self.feature_cards[attr_name] = card
        self._bind_card_action(card, command, widgets)
        return card

    def _bind_card_action(self, card, command, widgets):
        def is_enabled():
            return getattr(card, "_feature_enabled", True)

        def run_action(event=None):
            if is_enabled():
                command()

        def on_enter(event=None):
            if is_enabled():
                card.configure(fg_color=self.card_hover_color, border_color="#c9d9ea")

        def on_leave(event=None):
            if is_enabled():
                card.configure(fg_color="#fbfdff", border_color=self.border_color)

        for widget in widgets:
            widget.bind("<Button-1>", run_action)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            try:
                widget.configure(cursor="hand2")
            except tk.TclError:
                pass

    def _run_gui_smoke_and_exit(self):
        try:
            self._log_info("GUI 冒烟：开始检查关键入口。")
            for card in self.feature_cards.values():
                if not getattr(card, "_feature_enabled", True):
                    raise RuntimeError("GUI 冒烟：存在不可用功能卡片。")

            for dialog_callback in (self.show_about, self.run_data_drill, self.run_template_generate, self.run_clear_by_color):
                dialog_callback()
                self._flush_ui()
            self._log_info("GUI 冒烟：完成。")
        finally:
            self.root.after(200, self.root.destroy)

    def _create_log_area(self, parent):
        self.frame_bottom = ctk.CTkFrame(
            parent,
            fg_color=self.card_color,
            corner_radius=12,
            border_width=1,
            border_color="#e5edf5",
            height=116,
        )
        self.frame_bottom.grid(row=2, column=0, sticky="nsew")
        self.frame_bottom.grid_propagate(False)
        self.frame_bottom.grid_columnconfigure(0, weight=1)
        self.frame_bottom.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self.frame_bottom,
            text="运行日志",
            font=("Microsoft YaHei", 12, "bold"),
            text_color=self.text_color,
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(8, 4))

        self.log_text = ctk.CTkTextbox(
            self.frame_bottom,
            state="disabled",
            height=66,
            font=self.log_font,
            fg_color="#f8fafc",
            text_color="#111827",
            border_width=1,
            border_color=self.border_color,
            corner_radius=8,
            wrap="word",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

    def _get_project_root(self):
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

    def _find_user_guide_path(self):
        candidates = []
        if getattr(sys, "frozen", False):
            candidates.append(os.path.dirname(sys.executable))
            candidates.append(self._get_project_root())
        else:
            candidates.append(self._get_project_root())
            candidates.append(os.getcwd())

        seen = set()
        for base_dir in candidates:
            guide_path = os.path.abspath(os.path.join(base_dir, "使用说明.txt"))
            if guide_path in seen:
                continue
            seen.add(guide_path)
            if os.path.isfile(guide_path):
                return guide_path
        return None

    def show_user_guide(self):
        guide_path = self._find_user_guide_path()
        if guide_path:
            try:
                os.startfile(guide_path)
                self.logger.info(f"已打开使用说明：{guide_path}")
                return
            except OSError as exc:
                self.logger.error(f"打开使用说明失败：{exc}")

        self._show_info_no_grab(
            "使用说明",
            (
                "老头表格助手用于处理常见 Excel 批量整理工作。\n"
                "批量修改类功能请先备份或先用副本试跑。\n"
                "详细说明请查看同目录下的 使用说明.txt。"
            ),
            dialog_width=460,
            dialog_height=280,
            wraplength=420,
        )

    def show_about(self):
        self._show_info_no_grab(
            "关于",
            (
                f"软件名：{APP_NAME}\n"
                f"版本号：v{APP_VERSION}\n"
                "说明：面向审计、财务、报表整理场景的 Excel 效率工具\n"
                "当前阶段：常用功能阶段性完成"
            ),
            dialog_width=460,
            dialog_height=300,
            wraplength=380,
        )

    def _center_window(self, window, width=None, height=None):
        window.update_idletasks()
        width = width or window.winfo_reqwidth()
        height = height or window.winfo_reqheight()

        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = int((screen_width - width) / 2)
        y = int((screen_height - height) / 2)
        screen_margin = 16
        x = min(max(screen_margin, x), max(screen_margin, screen_width - width - screen_margin))
        y = min(max(screen_margin, y), max(screen_margin, screen_height - height - screen_margin))
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

    def _create_dialog_card(self, title, resizable=False):
        dialog = ctk.CTkToplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.resizable(resizable, resizable)
        dialog.transient(self.root)
        dialog.configure(fg_color=self.bg_color)
        dialog.minsize(1, 1)

        card = ctk.CTkFrame(
            dialog,
            fg_color=self.card_color,
            corner_radius=10,
            border_width=1,
            border_color="#e5edf5",
        )
        outer_pad = 16
        card.pack(fill=tk.BOTH, expand=True, padx=outer_pad, pady=outer_pad)
        dialog._dialog_card = card

        ctk.CTkLabel(
            card,
            text=title,
            font=("Microsoft YaHei", 22, "bold"),
            text_color=self.text_color,
        ).pack(anchor="w", padx=16, pady=(14, 8))

        return dialog, card

    def _add_dialog_message(self, parent, message, wraplength=360, pady=(0, 12)):
        label = ctk.CTkLabel(
            parent,
            text=message,
            font=("Microsoft YaHei", 16),
            wraplength=wraplength,
            justify="left",
            anchor="w",
            text_color=self.text_color,
        )
        label.pack(anchor="w", fill=tk.X, padx=16, pady=pady)
        return label

    def _create_dialog_button_bar(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(padx=12, pady=(6, 14), fill=tk.X)
        return button_frame

    def _add_dialog_button(self, parent, text, command, primary=False, width=None):
        button_width = width or max(86, min(260, len(text) * 14 + 28))
        return AppButton(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei", 15),
            width=button_width,
            height=36,
            fg_color=self.accent_color if primary else "#f8fafc",
            hover_color="#1d4ed8" if primary else "#e9eef5",
            text_color="#ffffff" if primary else self.text_color,
            border_width=0 if primary else 1,
            border_color=self.border_color,
            corner_radius=7,
        )

    def _show_dialog_no_grab(
        self,
        dialog,
        focus_widget=None,
        min_width=380,
        min_height=None,
        width=None,
        height=None,
    ):
        final_width = width if width is not None else max(dialog.winfo_reqwidth(), min_width)
        dialog.update_idletasks()
        final_height = height if height is not None else max(dialog.winfo_reqheight(), min_height or 300)
        self._center_window(dialog, width=final_width, height=final_height)
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

    def _schedule_smoke_dialog_close(self, dialog, close_callback=None, delay_ms=250):
        if os.environ.get("EXCEL_TOOLKIT_GUI_SMOKE") != "1":
            return

        def close():
            if dialog.winfo_exists():
                try:
                    if close_callback is None:
                        dialog.destroy()
                    else:
                        close_callback()
                except tk.TclError:
                    pass

        dialog.after(delay_ms, close)

    def _ask_text_no_grab(
        self,
        title,
        prompt,
        default="",
        entry_width=30,
        dialog_width=520,
        dialog_height=300,
        wraplength=460,
    ):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card(title)

        self._add_dialog_message(card, prompt, wraplength=wraplength, pady=(0, 10))

        entry = ctk.CTkEntry(
            card,
            font=("Microsoft YaHei", 12),
            width=entry_width * 10,
            height=34,
            fg_color="#fbfdff",
            border_color=self.border_color,
            text_color=self.text_color,
        )
        entry.pack(padx=16, pady=(0, 10), fill=tk.X)
        entry.insert(0, default)
        entry.select_range(0, tk.END)

        button_frame = self._create_dialog_button_bar(card)

        def confirm():
            result["value"] = entry.get()
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self._add_dialog_button(button_frame, "确定", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Return>", lambda event: confirm())
        dialog.bind("<Escape>", lambda event: cancel())
        self._show_dialog_no_grab(dialog, focus_widget=entry, width=dialog_width, height=dialog_height)
        self._schedule_smoke_dialog_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["value"]

    def _ask_choice_no_grab(
        self,
        title,
        prompt,
        choices,
        dialog_width=760,
        dialog_height=380,
        wraplength=680,
    ):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card(title)

        self._add_dialog_message(card, prompt, wraplength=wraplength, pady=(0, 12))

        button_frame = self._create_dialog_button_bar(card)

        def choose(value):
            result["value"] = value
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", lambda: choose(None), width=70).pack(side=tk.RIGHT)

        for label, value in reversed(choices):
            self._add_dialog_button(
                button_frame,
                label,
                lambda selected=value: choose(selected),
                primary=True,
                width=min(max(128, len(label) * 10 + 36), 180),
            ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        dialog.bind("<Escape>", lambda event: choose(None))
        self._show_dialog_no_grab(dialog, width=dialog_width, height=dialog_height)
        self._schedule_smoke_dialog_close(dialog, lambda: choose(None))
        self.root.wait_variable(done)
        return result["value"]

    def _show_info_no_grab(self, title, message, dialog_width=620, dialog_height=380, wraplength=500):
        dialog, card = self._create_dialog_card(title)

        self._add_dialog_message(card, message, wraplength=wraplength, pady=(0, 12))

        button_frame = self._create_dialog_button_bar(card)
        self._add_dialog_button(button_frame, "确定", dialog.destroy, primary=True).pack(side=tk.RIGHT)

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self._show_dialog_no_grab(dialog, width=dialog_width, height=dialog_height)
        self._schedule_smoke_dialog_close(dialog)

    def _ask_clear_multi_backup_option_no_grab(self):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card("按颜色清空内容")

        self._add_dialog_message(
            card,
            (
                "多个工作簿模式默认会先扫描，再把实际会被修改的文件集中备份到一个批次备份文件夹。\n"
                "无匹配、跳过或失败且未修改的文件不会备份。"
            ),
            wraplength=480,
            pady=(0, 12),
        )

        skip_backup_var = tk.BooleanVar(master=self.root, value=False)
        ctk.CTkCheckBox(
            card,
            text="已自行备份，本次不再生成备份文件",
            variable=skip_backup_var,
            font=("Microsoft YaHei", 11),
            text_color=self.text_color,
            fg_color=self.accent_color,
            hover_color="#1d4ed8",
            border_color=self.border_color,
        ).pack(anchor="w", fill=tk.X, padx=18, pady=(0, 12))

        button_frame = self._create_dialog_button_bar(card)

        def confirm():
            result["value"] = skip_backup_var.get()
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self._add_dialog_button(button_frame, "开始执行", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda event: cancel())
        self._show_dialog_no_grab(dialog, width=560, height=300)
        self._schedule_smoke_dialog_close(dialog, cancel)
        self.root.wait_variable(done)
        return result["value"]

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
                dialog_width=510,
                dialog_height=255,
                wraplength=460,
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
                    dialog_width=520,
                    wraplength=460,
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
                dialog_width=510,
                dialog_height=255,
                wraplength=460,
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

    def run_clear_by_color(self):
        self._log_info("按颜色清空内容：开始操作。")
        try:
            mode = self._ask_choice_no_grab(
                "按颜色清空内容",
                "请选择清空范围：\n\n"
                "当前活动工作簿\n"
                "清空当前打开的活动工作簿中所有可见 sheet 里，同颜色单元格的内容。\n\n"
                "多个工作簿\n"
                "批量清空所选工作簿中所有可见 sheet 里，同颜色单元格的内容。",
                [
                    ("当前活动工作簿", "active"),
                    ("多个工作簿", "multi"),
                ],
                dialog_width=560,
                dialog_height=310,
                wraplength=500,
            )
            if mode is None:
                self._log_info("用户已取消操作。")
                return

            self.btn_clear_by_color.config(state="disabled")
            if mode == "active":
                self._run_clear_current_workbook_by_color()
            else:
                self._run_clear_multiple_workbooks_by_color()
        except Exception as e:
            self._log_error(f"按颜色清空内容失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "按颜色清空内容",
                f"按颜色清空内容失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.btn_clear_by_color.config(state="normal")

    def _run_clear_current_workbook_by_color(self):
        plan = plan_clear_active_workbook_by_color(logger=self._flushing_logger())
        if plan["matched_cell_count"] == 0:
            self._log_info("未找到匹配单元格，已取消清空。")
            self._show_info_no_grab(
                "按颜色清空内容",
                "未找到匹配单元格，不执行清空。",
            )
            return

        result = execute_clear_active_workbook_plan(plan, logger=self._flushing_logger())
        self._show_info_no_grab(
            "按颜色清空内容",
            "清空完成。\n"
            f"当前颜色：{result['current_color_text']}\n"
            f"匹配工作表：{result['matched_sheet_count']}\n"
            f"清空工作表：{result['cleared_sheet_count']}\n"
            f"清空单元格：{result['cleared_cell_count']}\n"
            f"批量区域：{result['range_group_count']} 组\n\n"
            "当前工作簿未自动保存，请检查后自行保存。",
            dialog_width=520,
            wraplength=460,
        )

    def _run_clear_multiple_workbooks_by_color(self):
        target_paths = filedialog.askopenfilenames(
            title="请选择需要按颜色清空内容的 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not target_paths:
            self._log_info("用户已取消操作。")
            return

        skip_backup = self._ask_clear_multi_backup_option_no_grab()
        if skip_backup is None:
            self._log_info("用户已取消操作。")
            return

        result = clear_multiple_workbooks_by_color(
            list(target_paths),
            skip_backup=skip_backup,
            logger=self._flushing_logger(),
        )
        backup_message = (
            f"批次备份目录：{result['batch_backup_dir']}"
            if result["batch_backup_dir"]
            else "备份方式：用户选择跳过备份"
            if skip_backup
            else "备份方式：无需备份"
        )
        self._show_info_no_grab(
            "按颜色清空内容",
            "批量清空完成。\n"
            f"当前颜色：{result['current_color_text']}\n"
            f"成功处理文件：{result['processed_file_count']}\n"
            f"跳过文件：{result['skipped_file_count']}\n"
            f"失败文件：{result['failed_file_count']}\n"
            f"清空单元格：{result['cleared_cell_count']}\n"
            f"{backup_message}\n"
            f"处理日志：{result['log_path']}",
            dialog_width=720,
            wraplength=650,
        )

    def run_data_drill(self):
        """统一入口：按当前活动工作表和选区执行单文件或多文件数据穿透查询"""
        self._log_info("数据穿透查询：开始选择查询方式。")
        try:
            mode = self._ask_choice_no_grab(
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
                self._log_info("用户已取消操作。")
                return

            self.btn_data_drill.config(state="disabled")
            if mode == "single":
                self._run_single_workbook_data_drill()
            elif mode == "multi":
                self._run_multi_workbook_data_drill()
        except Exception as e:
            self._log_error(f"数据穿透查询失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "数据穿透查询",
                f"数据穿透查询失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.btn_data_drill.config(state="normal")

    def run_report_generate(self):
        """按钮回调函数，按模板清单逐行复制完整工作簿并替换外部链接"""
        self._log_info("按清单生成报表：开始操作。")
        template_path = filedialog.askopenfilename(
            title="请选择模板 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not template_path:
            self._log_info("用户已取消操作。")
            return

        sheet_name = self._ask_text_no_grab(
            "按清单生成报表",
            "请输入清单 sheet 名：",
            default=DEFAULT_CHECKLIST_SHEET_NAME,
            entry_width=24,
            dialog_width=360,
        )
        if sheet_name is None:
            self._log_info("用户已取消操作。")
            return
        sheet_name = sheet_name.strip() or DEFAULT_CHECKLIST_SHEET_NAME

        output_dir = filedialog.askdirectory(title="请选择输出目录")
        if not output_dir:
            self._log_info("用户已取消操作。")
            return

        choice = self._ask_choice_no_grab(
            "按清单生成报表",
            "请确认生成口径：\n"
            f"模板路径：{template_path}\n"
            f"清单 sheet：{sheet_name}\n"
            f"输出目录：{output_dir}\n\n"
            "清单字段固定为：A列新 TB / 新链接文件路径，B列输出文件名主体，"
            "C列公司名称，D列报表类型。\n"
            "程序会逐行复制完整模板工作簿，并把模板外部链接替换为 A 列路径。\n"
            "原模板工作簿不会被修改。",
            [("开始生成", "run")],
            dialog_width=680,
            wraplength=620,
        )
        if choice != "run":
            self._log_info("用户已取消操作。")
            return

        self.btn_template_tb_report.config(state="disabled")
        try:
            result = generate_reports_from_checklist(
                template_path=template_path,
                sheet_name=sheet_name,
                output_dir=output_dir,
                logger=self._flushing_logger(),
            )
            self._log_info(
                "按清单生成报表完成："
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个；"
                f"失败 {result['failed_count']} 个；"
                f"模板外部链接 {result['template_link_count']} 个。"
            )
            for record in result["records"]:
                self._log_info(
                    f"第 {record['row_number']} 行：{record['status']}，"
                    f"{record['message']}，输出：{record.get('output_path') or '-'}"
                )
            self._show_info_no_grab(
                "按清单生成报表",
                "生成完成。\n"
                f"成功数量：{result['success_count']}\n"
                f"跳过数量：{result['skipped_count']}\n"
                f"失败数量：{result['failed_count']}\n"
                f"输出目录：{result['output_dir']}\n\n"
                "原模板工作簿未被修改。",
                dialog_width=560,
                wraplength=500,
            )
        except Exception as e:
            self._log_error(f"按清单生成报表失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "按清单生成报表",
                f"按清单生成报表失败：{e}",
                dialog_width=560,
                wraplength=500,
            )
        finally:
            self.btn_template_tb_report.config(state="normal")

    def run_template_generate(self):
        """按钮回调函数，先在弹窗中让用户选择生成模式，再调用对应实现。"""
        self._log_info("按模板批量生成 Excel：开始选择生成模式。")
        mode = self._ask_choice_no_grab(
            "按模板批量生成 Excel",
            "请选择生成方式：\n\n"
            "替换一个链接生成\n"
            "适合：模板里只需要更换一个 TB / 附注 / 底稿链接。\n\n"
            "替换多个链接生成\n"
            "适合：模板里有合并 TB、单体 TB 等多个链接，需要用规则表逐项替换。",
            [
                ("替换一个链接生成", "single"),
                ("替换多个链接生成", "multi"),
            ],
            dialog_width=575,
            dialog_height=310,
            wraplength=520,
        )
        if mode is None:
            self._log_info("用户已取消操作。")
            return
        if mode == "single":
            self._log_info("已选择：替换一个链接生成。")
            self.run_template_tb_report()
            return
        if mode == "multi":
            self._log_info("已选择：替换多个链接生成。")
            self.run_template_multi_link()

    def run_template_tb_report(self):
        """扫描模板外部链接、多选 TB 文件、按 TB 生成多份报表（替换一个链接生成）。"""
        self._log_info("按模板批量生成 Excel（替换一个链接生成）：开始操作。")
        template_path = filedialog.askopenfilename(
            title="请选择模板 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not template_path:
            self._log_info("用户已取消操作。")
            return

        self.btn_template_tb_report.config(state="disabled")
        try:
            try:
                self._log_info("正在打开模板并扫描外部链接，请稍候……")
                template_links = read_template_external_links(template_path)
            except Exception as e:
                self._log_error(f"扫描模板外部链接失败：{e}")
                self._show_info_no_grab(
                    "按模板批量生成 Excel",
                    f"扫描模板外部链接失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            if not template_links:
                self._log_error("模板没有外部链接，无法执行换链接生成报表。")
                self._show_info_no_grab(
                    "按模板批量生成 Excel",
                    "模板没有外部链接，无法执行换链接生成报表。\n请确认模板含有引用其他工作簿的公式后重试。",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            self._log_info(f"模板外部链接数量：{len(template_links)}")
            for link_path in template_links:
                self._log_info(f"  - {link_path}")

            if len(template_links) == 1:
                old_link_path = template_links[0]
                self._log_info(f"模板只有 1 个外部链接，已自动选择：{old_link_path}")
            else:
                old_link_path = self._ask_old_link_choice_no_grab(
                    template_links,
                    title="选择要替换的旧链接",
                    prompt="模板存在多个外部链接，请选择要替换的旧链接：",
                )
                if not old_link_path:
                    self._log_info("用户已取消操作。")
                    return
                self._log_info(f"用户选择要替换的旧链接：{old_link_path}")

            tb_paths = filedialog.askopenfilenames(
                title="请多选 TB 文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*"),
                ],
            )
            if not tb_paths:
                self._log_info("用户已取消操作。")
                return

            output_dir = filedialog.askdirectory(title="请选择输出目录")
            if not output_dir:
                self._log_info("用户已取消操作。")
                return

            suffix_input = self._ask_text_no_grab(
                "按模板批量生成 Excel",
                "请输入输出文件名后缀，留空则直接使用 TB 文件名主体。\n"
                "常用：_批量生成、_附注、_财务报表、_底稿、_报表",
                default=DEFAULT_OUTPUT_NAME_SUFFIX,
                entry_width=24,
                dialog_width=560,
                wraplength=500,
            )
            if suffix_input is None:
                self._log_info("用户已取消操作。")
                return
            output_name_suffix = suffix_input.strip()

            confirmed = self._ask_choice_no_grab(
                "按模板批量生成 Excel",
                "请确认生成口径：\n"
                f"模板路径：{template_path}\n"
                f"被替换的旧链接：{old_link_path}\n"
                f"TB 文件数量：{len(tb_paths)}\n"
                f"输出目录：{output_dir}\n"
                f"输出后缀：{output_name_suffix or '（留空：直接使用 TB 文件名主体）'}\n\n"
                "程序会逐个 TB 复制完整模板，并将选中的旧链接替换为该 TB 路径；\n"
                "其他外部链接保持不变；原模板不会被修改。",
                [("开始生成", "run")],
                dialog_width=820,
                wraplength=740,
            )
            if confirmed != "run":
                self._log_info("用户已取消操作。")
                return

            self._log_info(f"TB 文件数量：{len(tb_paths)}")
            self._log_info(f"输出目录：{output_dir}")
            self._log_info(f"输出后缀：{output_name_suffix or '（留空）'}")
            result = generate_reports_from_template_and_tb_files(
                template_path=template_path,
                tb_paths=list(tb_paths),
                output_dir=output_dir,
                old_link_path=old_link_path,
                output_name_suffix=output_name_suffix,
                logger=self._flushing_logger(),
            )
            for record in result["records"]:
                self._log_info(
                    f"第 {record.index} 个：{record.tb_name} -> {record.output_name or '-'}；"
                    f"{record.status}；{record.message}"
                )
            self._log_info(
                "按模板批量生成 Excel完成："
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个；"
                f"失败 {result['failed_count']} 个；"
                f"日志：{result['log_path']}"
            )
            self._show_info_no_grab(
                "按模板批量生成 Excel",
                "生成完成。\n"
                f"TB 文件数量：{result['tb_file_count']}\n"
                f"成功数量：{result['success_count']}\n"
                f"跳过数量：{result['skipped_count']}\n"
                f"失败数量：{result['failed_count']}\n"
                f"输出目录：{result['output_dir']}\n"
                f"日志文件：{result['log_path']}\n\n"
                "原模板工作簿未被修改。",
                dialog_width=720,
                wraplength=650,
            )
        except Exception as e:
            self._log_error(f"按模板批量生成 Excel失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "按模板批量生成 Excel",
                f"按模板批量生成 Excel失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.btn_template_tb_report.config(state="normal")

    def _ask_old_link_choice_no_grab(self, link_paths, title, prompt):
        result = {"value": None}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card(title)

        self._add_dialog_message(card, prompt, wraplength=520, pady=(0, 12))

        selected_index = tk.IntVar(master=dialog, value=0)
        radio_frame = ctk.CTkFrame(
            card,
            fg_color="#fbfdff",
            corner_radius=8,
            border_width=1,
            border_color=self.border_color,
        )
        radio_frame.pack(fill=tk.X, padx=18, pady=(0, 12))
        for index, link_path in enumerate(link_paths):
            row_frame = ctk.CTkFrame(radio_frame, fg_color="transparent")
            row_frame.pack(fill=tk.X, padx=12, pady=5)
            row_frame.grid_columnconfigure(1, weight=1)
            ctk.CTkRadioButton(
                row_frame,
                text="",
                variable=selected_index,
                value=index,
                width=24,
                fg_color=self.accent_color,
                hover_color="#1d4ed8",
                border_color=self.border_color,
            ).grid(row=0, column=0, sticky="n", pady=1)
            link_label = ctk.CTkLabel(
                row_frame,
                text=link_path,
                font=("Microsoft YaHei", 11),
                wraplength=520,
                justify="left",
                text_color=self.text_color,
            )
            link_label.grid(row=0, column=1, sticky="w", padx=(8, 0))
            link_label.bind("<Button-1>", lambda event, selected=index: selected_index.set(selected))

        button_frame = self._create_dialog_button_bar(card)

        def confirm():
            try:
                result["value"] = link_paths[selected_index.get()]
            except (IndexError, tk.TclError):
                result["value"] = None
            done.set(True)
            dialog.destroy()

        def cancel():
            result["value"] = None
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        self._add_dialog_button(button_frame, "确定", confirm, primary=True).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda event: cancel())
        self._show_dialog_no_grab(dialog, width=600, height=340)
        self.root.wait_variable(done)
        return result["value"]

    def run_template_multi_link(self):
        """扫描模板外部链接、生成规则表后批量按多链接替换（替换多个链接生成）。"""
        self._log_info("按模板批量生成 Excel（替换多个链接生成）：开始操作。")
        template_path = filedialog.askopenfilename(
            title="请选择模板 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not template_path:
            self._log_info("用户已取消操作。")
            return

        self.btn_template_tb_report.config(state="disabled")
        try:
            try:
                self._log_info("正在打开模板并扫描外部链接，请稍候……")
                template_links = read_template_external_links(template_path)
            except Exception as e:
                self._log_error(f"扫描模板外部链接失败：{e}")
                self._show_info_no_grab(
                    "按模板多链接生成 Excel",
                    f"扫描模板外部链接失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            if not template_links:
                self._log_error("模板没有外部链接，无法生成多链接规则表。")
                self._show_info_no_grab(
                    "按模板多链接生成 Excel",
                    "模板没有外部链接，无法生成多链接规则表。\n请确认模板含有引用其他工作簿的公式后重试。",
                    dialog_width=620,
                    wraplength=560,
                )
                return

            self._log_info(f"模板外部链接数量：{len(template_links)}")
            for link_path in template_links:
                self._log_info(f"  - {link_path}")

            if len(template_links) == 1:
                main_old_link = template_links[0]
                self._log_info(f"模板只有 1 个外部链接，已作为主链接：{main_old_link}")
            else:
                main_old_link = self._ask_old_link_choice_no_grab(
                    template_links,
                    title="选择主链接",
                    prompt=(
                        "模板存在多个外部链接，请选择主链接。\n\n"
                        "主链接对应的新源文件数量，将决定生成多少份 Excel。\n"
                        "其他链接可在后续规则表中按行填写；新链接留空则保持原链接不变。"
                    ),
                )
                if not main_old_link:
                    self._log_info("用户已取消操作。")
                    return
                self._log_info(f"用户选择主链接：{main_old_link}")

            main_source_paths = filedialog.askopenfilenames(
                title="请多选主链接对应的新源文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*"),
                ],
            )
            if not main_source_paths:
                self._log_info("用户已取消操作。")
                return

            suffix_input = self._ask_text_no_grab(
                "按模板多链接生成 Excel",
                "请输入输出文件名后缀，留空则直接使用主源文件名主体。\n"
                "常用：_批量生成、_附注、_财务报表、_底稿、_报表",
                default=DEFAULT_OUTPUT_NAME_SUFFIX,
                entry_width=24,
                dialog_width=560,
                wraplength=500,
            )
            if suffix_input is None:
                self._log_info("用户已取消操作。")
                return
            output_name_suffix = suffix_input.strip()

            output_dir = filedialog.askdirectory(title="请选择输出目录")
            if not output_dir:
                self._log_info("用户已取消操作。")
                return

            self._log_info(f"主源文件数量：{len(main_source_paths)}")
            self._log_info(f"输出目录：{output_dir}")
            self._log_info(f"输出后缀：{output_name_suffix or '（留空）'}")
            rule_path = create_template_multi_link_rule_workbook(
                template_path=template_path,
                output_dir=output_dir,
                old_links=list(template_links),
                main_old_link=main_old_link,
                main_source_paths=list(main_source_paths),
                output_name_suffix=output_name_suffix,
            )
            self._log_info(f"多链接生成规则表路径：{rule_path}")

            try:
                os.startfile(rule_path)
                self._log_info("规则表已自动打开。请检查并填写其他新链接列，保存并关闭规则表后点击执行。")
            except Exception as open_error:
                self._log_error(f"规则表已生成，但自动打开失败：{open_error}")
                self._log_info(f"请手动打开规则表：{rule_path}")

            if not self._confirm_template_multi_link_rule_ready(rule_path):
                self._log_info("用户已取消操作。")
        except Exception as e:
            self._log_error(f"按模板多链接生成 Excel失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "按模板多链接生成 Excel",
                f"按模板多链接生成 Excel失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.btn_template_tb_report.config(state="normal")

    def _confirm_template_multi_link_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card("按模板多链接生成 Excel")

        self._add_dialog_message(
            card,
            (
                "规则表已打开。\n"
                "请检查【生成清单】，按需填写其他旧链接对应的新链接列；\n"
                "新链接留空表示不替换该旧链接。\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=460,
            pady=(0, 12),
        )

        button_frame = self._create_dialog_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self._log_info(f"正在执行多链接生成规则表：{rule_workbook_path}")
                try:
                    summary = execute_template_multi_link_generation_from_rule_workbook(
                        rule_workbook_path,
                        logger=self._flushing_logger(),
                    )
                except (PermissionError, OSError) as e:
                    self._log_error(f"规则表无法读取或回写：{e}")
                    self._show_template_multi_link_rule_workbook_busy_message()
                    return

                result_message = (
                    "按模板多链接生成 Excel 完成。\n"
                    f"规则行数：{summary['rule_count']}\n"
                    f"成功：{summary['success_count']} 个\n"
                    f"跳过：{summary['skipped_count']} 个\n"
                    f"失败：{summary['failed_count']} 个\n\n"
                    "规则表已更新状态和处理日志。\n"
                    "原模板工作簿未被修改。"
                )
                self._log_info(
                    "按模板多链接生成 Excel完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self._show_info_no_grab(
                    "按模板多链接生成 Excel",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self._log_error(f"按模板多链接生成 Excel失败：{e}")
                self._show_info_no_grab(
                    "按模板多链接生成 Excel",
                    f"按模板多链接生成 Excel失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self._add_dialog_button(
            button_frame,
            "我已检查规则，开始生成",
            execute,
            primary=True,
            width=210,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, width=550, height=280)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_template_multi_link_rule_workbook_busy_message(self):
        self._log_error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self._show_info_no_grab(
            "按模板多链接生成 Excel",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行。",
            dialog_width=620,
            wraplength=560,
        )

    def run_round_formula(self):
        """按钮回调函数，将当前 Excel 选区内数值/公式直接包裹 ROUND(...,2)"""
        self._log_info("选区 ROUND 保留两位：开始操作。")
        self.btn_round_formula.config(state="disabled")
        try:
            result = round_selected_range_to_two_decimals(logger=self._flushing_logger())
            self._log_info(
                "选区 ROUND 保留两位完成："
                f"工作簿 {result['workbook_name']}；"
                f"Sheet {result['sheet_name']}；"
                f"选区 {result['selection_address']}；"
                f"成功 {result['success_count']} 个；"
                f"跳过 {result['skipped_count']} 个。"
            )
            self._log_info("当前工作簿未自动保存，请检查后自行保存。")
            self._show_info_no_grab(
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
            self._log_error(f"选区 ROUND 保留两位失败：{type(e).__name__}: {e}")
            self._show_info_no_grab(
                "选区 ROUND 保留两位",
                f"选区 ROUND 保留两位失败：{e}",
                dialog_width=420,
                dialog_height=240,
                wraplength=380,
            )
        finally:
            self.btn_round_formula.config(state="normal")

    def _confirm_data_drill_context(self, context):
        choice = self._ask_choice_no_grab(
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

    def _get_current_excel_context_for_data_drill(self):
        return self._get_active_excel_drill_context(require_saved_workbook=True)

    def _run_single_workbook_data_drill(self):
        self._log_info("数据穿透查询（单文件）：开始操作。")
        self._run_excel_data_drill_in_com_session(self._execute_single_workbook_data_drill)

    def _run_multi_workbook_data_drill(self):
        self._log_info("数据穿透查询（多文件）：开始操作。")
        context = self._get_active_excel_drill_context(require_saved_workbook=True)
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

        output_path = build_unique_output_path(context["output_dir"], base_name=RESULT_SHEET_BASE_NAME)
        self._log_info(f"当前工作表：{context['sheet_name']}")
        self._log_info(f"当前选区：{context['range_address']}")
        self._log_info(f"源文件数量：{len(source_paths)}")
        records = build_data_drill_range_records(
            source_paths=list(source_paths),
            sheet_name=context["sheet_name"],
            range_address=context["range_address"],
            logger=self._flushing_logger(),
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
            self._log_info(f"结果文件已生成并尝试打开：{output_path}")
        except Exception as open_error:
            self._log_error(f"结果文件已生成，但自动打开失败：{open_error}")
            self._log_info(f"请手动打开结果文件：{output_path}")

        self._log_info(f"输出文件：{output_path}")
        self._log_info("数据穿透查询（多文件）完成。")
        self._log_info(
            "数据穿透查询（多文件）统计："
            f"成功 {summary['success_count']} 个；"
            f"跳过 {summary['skipped_count']} 个；"
            f"失败 {summary['failed_count']} 个。"
        )
        self._log_info("当前活动工作簿和源文件均未被修改、未被自动保存。")
        self._show_info_no_grab(
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

    def _get_active_excel_drill_context(self, require_saved_workbook: bool):
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
            if require_saved_workbook and not workbook_dir:
                raise RuntimeError(
                    "当前活动工作簿尚未保存，无法确定结果文件输出目录。请先保存当前工作簿后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}。"
                )
            workbook_path = str(workbook.FullName or "").strip()
            if workbook_path and not os.path.dirname(workbook_path) and workbook_dir:
                workbook_path = os.path.join(workbook_dir, workbook_name)
            workbook_path_text = workbook_path or workbook_name
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

            range_address = self._get_active_range_address_for_data_drill(excel)
            if not range_address:
                raise RuntimeError(
                    "没有有效选区，请先选中一个需要追查的单元格或区域后重试。"
                    f"当前步骤：{self._format_data_drill_context_steps(steps)}，但未取得 ActiveCell 或可用 Selection。"
                )
            steps.append("已取得 Selection")

            return {
                "workbook_name": workbook_name,
                "workbook_path": workbook_path,
                "workbook_path_text": workbook_path_text,
                "output_dir": workbook_dir,
                "workbook_dir": workbook_dir,
                "sheet_name": sheet_name,
                "range_address": range_address,
            }
        finally:
            pythoncom.CoUninitialize()

    def _get_active_range_address_for_data_drill(self, excel):
        try:
            selection = excel.Selection
        except Exception as e:
            raise RuntimeError(f"已连接 Excel，但无法读取当前 Selection：{e}") from e

        if selection is None:
            raise RuntimeError("已连接 Excel，但当前没有可识别的选区。")
        try:
            areas = selection.Areas.Count
        except Exception:
            areas = 1
        if areas and int(areas) > 1:
            raise RuntimeError("不支持多区域选区，请只选择一个连续区域后重试。")

        try:
            address_member = selection.Address
            if callable(address_member):
                address = address_member(False, False)
            else:
                address = address_member
        except Exception:
            address = selection.Address(False, False)
        normalized_address = normalize_range_address(address)
        if is_multi_area_range(normalized_address):
            raise RuntimeError("不支持多区域选区，请只选择一个连续区域后重试。")
        return normalized_address

    def _run_excel_data_drill_in_com_session(self, runner):
        try:
            import pythoncom
            import win32com.client
        except Exception as e:
            raise RuntimeError("无法加载 Excel COM 组件，请确认已安装 pywin32 并在 Windows + Excel 环境运行。") from e

        pythoncom.CoInitialize()
        try:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception as e:
                raise RuntimeError("未检测到正在运行的 Excel，请先打开目标/合并工作簿后重试。") from e
            runner(excel)
        finally:
            pythoncom.CoUninitialize()

    def _execute_single_workbook_data_drill(self, excel):
        workbook = excel.ActiveWorkbook
        if workbook is None:
            raise RuntimeError("Excel 中没有活动工作簿，请先打开目标工作簿后重试。")

        active_sheet = excel.ActiveSheet
        if active_sheet is None or not str(active_sheet.Name):
            raise RuntimeError("无法读取当前活动 Sheet，请先切换到目标工作表后重试。")

        range_address = self._get_active_range_address_for_data_drill(excel)
        workbook_name = str(workbook.Name)
        workbook_path = str(workbook.FullName or "").strip()
        workbook_path_text = workbook_path or workbook_name
        sheet_name = str(active_sheet.Name)

        self._log_info(f"当前工作簿：{workbook_path_text}")
        self._log_info(f"当前工作表：{sheet_name}")
        self._log_info(f"当前选区：{range_address}")

        records = []
        visible_sheet_count = 0
        for sheet in workbook.Worksheets:
            visible_status = "可见" if not is_sheet_hidden_by_visible_value(sheet.Visible) else "隐藏"
            if visible_status != "可见":
                continue
            if should_skip_history_result_sheet(str(sheet.Name), RESULT_SHEET_BASE_NAME):
                continue

            visible_sheet_count += 1
            values_by_address = self._read_excel_range_values_from_com_sheet(sheet, range_address)
            records.append(
                {
                    "sheet_name": str(sheet.Name),
                    "visible_status": visible_status,
                    "values_by_address": values_by_address,
                }
            )

        result_sheet_name = build_unique_result_sheet_name([str(sheet.Name) for sheet in workbook.Worksheets])
        result_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        result_sheet.Name = result_sheet_name
        write_single_workbook_drill_result_to_com_sheet(
            sheet=result_sheet,
            base_sheet_name=sheet_name,
            range_address=range_address,
            created_at_text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            records=records,
        )

        self._log_info(f"读取工作表数量：{visible_sheet_count}")
        self._log_info(f"结果工作表：{result_sheet_name}")
        self._log_info("数据穿透查询（单文件）完成。")
        self._log_info("当前活动工作簿已新增结果工作表，但未自动保存。")
        self._show_info_no_grab(
            "数据穿透查询（单文件）",
            "数据穿透查询（单文件）完成。\n"
            f"当前工作簿：{workbook_path_text}\n"
            f"当前工作表：{sheet_name}\n"
            f"当前选区：{range_address}\n"
            f"读取工作表数量：{visible_sheet_count}\n"
            f"结果工作表：{result_sheet_name}\n\n"
            "当前活动工作簿未自动保存。",
            dialog_width=620,
            wraplength=560,
        )

    def _read_excel_range_values_from_com_sheet(self, sheet, range_address: str) -> dict:
        try:
            raw_value = sheet.Range(range_address).Value
            from .workbook_drill_ops import range_value_to_address_map

            return range_value_to_address_map(range_address, raw_value)
        except Exception as e:
            message = f"读取异常：{e}"
            self._log_error(f"{sheet.Name}!{range_address} 读取失败：{message}")
            from .workbook_drill_ops import expand_range_addresses

            return {address: message for address in expand_range_addresses(range_address)}

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
            dialog_width=720,
            wraplength=650,
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
            dialog_width=720,
            wraplength=650,
        )

    def _confirm_delete_rule_ready(self, rule_table_path):
        result = {"execute": False}
        done = tk.BooleanVar(value=False)
        dialog, card = self._create_dialog_card("批量删除工作表")

        self._add_dialog_message(
            card,
            "规则表已打开。请在 B/C/D 列填写规则并保存规则表后，再点击执行。",
            wraplength=360,
            pady=(0, 12),
        )

        button_frame = self._create_dialog_button_bar(card)

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

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self._add_dialog_button(
            button_frame,
            "我已填好规则，执行批量删除",
            execute,
            primary=True,
            width=230,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, width=510, height=255)
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
            dialog_width=760,
            wraplength=680,
        )
        if mode is None:
            self.logger.info("用户已取消操作")
            return None

        return mode

    def run_batch_link_replace(self):
        """按钮回调函数，扫描外部链接并按规则批量 ChangeLink"""
        self.logger.info("批量更换多文件链接：开始操作。")
        source_paths = filedialog.askopenfilenames(
            title="请选择需要批量更换多文件链接的 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not source_paths:
            self.logger.info("用户已取消操作")
            return

        self.btn_link_replace.config(state="disabled")
        try:
            result = generate_temporary_link_replace_rule_workbook(
                source_paths=list(source_paths),
                logger=self._flushing_logger(),
            )
            self.logger.info(f"选择文件数量：{result['source_file_count']}")
            self.logger.info(f"读取成功文件数：{result['read_success_count']}")
            self.logger.info(f"扫描到外部链接数量：{result['link_count']}")
            self.logger.info(f"临时规则表路径：{result['output_path']}")

            try:
                os.startfile(result["output_path"])
                self.logger.info("规则表已自动打开。请填写 D 列，保存并关闭规则表后，再点击弹窗中的执行按钮。")
            except Exception as open_error:
                self.logger.error(f"规则表已生成，但自动打开失败：{open_error}")
                self.logger.info(f"请手动打开规则表：{result['output_path']}")

            if not self._confirm_link_replace_rule_ready(result["output_path"]):
                self.logger.info("用户已取消操作")
        except Exception as e:
            self.logger.error(f"批量更换多文件链接失败：{e}")
            self._show_info_no_grab(
                "批量更换多文件链接",
                f"批量更换多文件链接失败：{e}",
                dialog_width=620,
                wraplength=560,
            )
        finally:
            self.btn_link_replace.config(state="normal")

    def _confirm_link_replace_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card("批量更换多文件链接")

        self._add_dialog_message(
            card,
            (
                "规则表已打开。\n"
                "请在【链接清单】D列填写新链接路径，E列可填写是否执行，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=420,
            pady=(0, 12),
        )

        button_frame = self._create_dialog_button_bar(card)

        def execute():
            execute_button.config(state="disabled")
            try:
                self.logger.info(f"正在执行批量更换多文件链接规则表：{rule_workbook_path}")
                try:
                    summary = execute_link_replacement_from_rule_workbook(
                        rule_workbook_path,
                        logger=self._flushing_logger(),
                    )
                except (PermissionError, OSError) as e:
                    self.logger.error(f"规则表无法读取或回写：{e}")
                    self._show_link_replace_rule_workbook_busy_message()
                    return

                result_message = (
                    "批量更换多文件链接完成。\n"
                    f"规则行数：{summary['rule_count']}\n"
                    f"成功：{summary['success_count']} 个\n"
                    f"跳过：{summary['skipped_count']} 个\n"
                    f"失败：{summary['failed_count']} 个\n\n"
                    "规则表已更新状态和处理日志。\n"
                    "成功处理的目标工作簿已按参数保存。"
                )
                self.logger.info(
                    "批量更换多文件链接完成："
                    f"成功 {summary['success_count']} 个；"
                    f"跳过 {summary['skipped_count']} 个；"
                    f"失败 {summary['failed_count']} 个。"
                )
                self._show_info_no_grab(
                    "批量更换多文件链接",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量更换多文件链接失败：{e}")
                self._show_info_no_grab(
                    "批量更换多文件链接",
                    f"批量更换多文件链接失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self._add_dialog_button(
            button_frame,
            "我已填好规则，执行批量更换多文件链接",
            execute,
            primary=True,
            width=320,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, width=550, height=280)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_link_replace_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self._show_info_no_grab(
            "批量更换多文件链接",
            "规则表仍被 Excel 占用，无法读取或回写。\n"
            "请先保存并关闭规则表，再点击执行批量更换多文件链接。",
            dialog_width=620,
            wraplength=560,
        )

    def _confirm_sheet_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card("批量重命名工作表")

        self._add_dialog_message(
            card,
            (
                "规则表已打开。\n"
                "请在【重命名清单】D列填写新工作表名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=400,
            pady=(0, 12),
        )

        button_frame = self._create_dialog_button_bar(card)

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
                self._show_info_no_grab(
                    "批量重命名工作表",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名工作表失败：{e}")
                self._show_info_no_grab(
                    "批量重命名工作表",
                    f"批量重命名工作表失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self._add_dialog_button(
            button_frame,
            "我已填好规则，执行批量重命名工作表",
            execute,
            primary=True,
            width=320,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, width=510, height=255)
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
            dialog_width=620,
            wraplength=560,
        )

    def _confirm_rename_rule_ready(self, rule_workbook_path):
        result = {"execute": False}
        done = tk.BooleanVar(master=self.root, value=False)
        dialog, card = self._create_dialog_card("批量重命名文件")

        self._add_dialog_message(
            card,
            (
                "规则表已打开。\n"
                "请在【重命名清单】中填写新文件名和后缀名，\n"
                "保存并关闭/释放规则表后，再点击执行。"
            ),
            wraplength=400,
            pady=(0, 12),
        )

        button_frame = self._create_dialog_button_bar(card)

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
                self._show_info_no_grab(
                    "批量重命名文件",
                    result_message,
                    dialog_width=720,
                    wraplength=650,
                )
                result["execute"] = True
                done.set(True)
                dialog.destroy()
            except Exception as e:
                self.logger.error(f"批量重命名文件失败：{e}")
                self._show_info_no_grab(
                    "批量重命名文件",
                    f"批量重命名文件失败：{e}",
                    dialog_width=620,
                    wraplength=560,
                )
            finally:
                if dialog.winfo_exists():
                    execute_button.config(state="normal")

        def cancel():
            done.set(True)
            dialog.destroy()

        self._add_dialog_button(button_frame, "取消", cancel).pack(side=tk.RIGHT)
        execute_button = self._add_dialog_button(
            button_frame,
            "我已填好规则，执行批量重命名",
            execute,
            primary=True,
            width=260,
        )
        execute_button.pack(side=tk.RIGHT, padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._show_dialog_no_grab(dialog, width=510, height=255)
        self.root.wait_variable(done)
        return result["execute"]

    def _show_rule_workbook_busy_message(self):
        self.logger.error("规则表仍被 Excel 占用，无法继续执行。请保存并关闭规则表后重试。")
        self._show_info_no_grab(
            "批量重命名文件",
            "规则表仍被 Excel 占用，无法回写。\n"
            "请先保存并关闭规则表，再点击执行批量重命名。",
            dialog_width=620,
            wraplength=560,
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
    root = ctk.CTk()
    app = ExcelToolkitApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
