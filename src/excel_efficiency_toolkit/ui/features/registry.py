# -*- coding: utf-8 -*-
"""声明式功能注册表：5 个分区 × (标题, 描述, attr_name, 回调)。

attr_name 是功能卡片句柄在 ctx.buttons 中的注册名，
控制器通过 self.button(attr_name) 控制启用/禁用。
"""


def build_feature_groups(controllers):
    return [
        (
            "拆分导出",
            [
                (
                    "按工作表拆分文件",
                    "",
                    "btn_export_sheets",
                    controllers.export_split.run_export_sheets,
                ),
                (
                    "按指定列拆分工作表",
                    "",
                    "btn_split_sheet",
                    controllers.export_split.run_split_sheet,
                ),
            ],
        ),
        (
            "合并整理",
            [
                (
                    "多个工作表合并到一张表",
                    "",
                    "btn_merge_sheets",
                    controllers.merge.run_merge_sheets,
                ),
                (
                    "多个 Excel 文件合并到一个文件",
                    "",
                    "btn_merge_workbooks",
                    controllers.merge.run_merge_workbooks,
                ),
            ],
        ),
        (
            "模板生成 / 取数",
            [
                (
                    "数据穿透查询",
                    "查询多文件/多工作表中的单元格或区域数值",
                    "btn_data_drill",
                    controllers.data_drill.run_data_drill,
                ),
                (
                    "按模板批量生成 Excel",
                    "替换链接并生成多个 Excel 文件",
                    "btn_template_tb_report",
                    controllers.templates.run_template_generate,
                ),
                (
                    "按颜色汇总求和",
                    "按填充色位置对多文件数值求和",
                    "btn_color_sum",
                    controllers.color_tools.run_color_sum,
                ),
                (
                    "按颜色清空内容",
                    "",
                    "btn_clear_by_color",
                    controllers.color_tools.run_clear_by_color,
                ),
                (
                    "选区 ROUND 保留两位",
                    "",
                    "btn_round_formula",
                    controllers.workbook_misc.run_round_formula,
                ),
            ],
        ),
        (
            "工作簿辅助",
            [
                (
                    "生成工作表目录",
                    "",
                    "btn_sheet_index",
                    controllers.workbook_misc.run_sheet_index,
                ),
            ],
        ),
        (
            "批量维护",
            [
                ("批量重命名文件", "", "btn_rename_files", controllers.rename_files.run_batch_rename_files),
                (
                    "批量重命名工作表",
                    "",
                    "btn_rename_sheets",
                    controllers.rename_sheets.run_batch_rename_sheets,
                ),
                (
                    "批量更换 Excel 链接",
                    "批量替换多个文件的外部链接",
                    "btn_link_replace",
                    controllers.link_replace.run_batch_link_replace,
                ),
                (
                    "Word 批量替换",
                    "批量替换多个 Word 文件中的指定文字",
                    "btn_word_replace",
                    controllers.word_replace.run_word_batch_replace,
                ),
                ("批量删除工作表", "", "btn_delete_sheets", controllers.delete_sheets.run_delete_sheets),
            ],
        ),
    ]
