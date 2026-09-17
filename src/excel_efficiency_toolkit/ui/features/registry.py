# -*- coding: utf-8 -*-
"""声明式功能目录：页面分组、说明文字和业务回调的唯一来源。"""
from dataclasses import dataclass
from typing import Callable, Tuple


@dataclass(frozen=True)
class FeatureSpec:
    title: str
    description: str
    attr_name: str
    command: Callable


@dataclass(frozen=True)
class FeatureGroup:
    key: str
    title: str
    icon: str
    summary: str
    features: Tuple[FeatureSpec, ...]


def build_feature_groups(controllers):
    """构建按实际工作场景组织的功能目录。"""
    return (
        FeatureGroup(
            "merge",
            "合并整理",
            "合",
            "聚合分散表格，快速形成可继续分析的数据底稿。",
            (
                FeatureSpec(
                    "多个工作表合并到一张表",
                    "汇总同一工作簿或多个来源中的表格数据",
                    "btn_merge_sheets",
                    controllers.merge.run_merge_sheets,
                ),
                FeatureSpec(
                    "多个 Excel 文件合并到一个文件",
                    "保留工作表结构，将多个文件集中到单一工作簿",
                    "btn_merge_workbooks",
                    controllers.merge.run_merge_workbooks,
                ),
            ),
        ),
        FeatureGroup(
            "export",
            "拆分导出",
            "拆",
            "按工作表或字段批量拆分，减少重复另存与筛选操作。",
            (
                FeatureSpec(
                    "按工作表拆分文件",
                    "将每张工作表分别导出为独立 Excel 文件",
                    "btn_export_sheets",
                    controllers.export_split.run_export_sheets,
                ),
                FeatureSpec(
                    "按列拆分",
                    "按指定列内容拆分为工作表或多个 Excel 文件",
                    "btn_split_sheet",
                    controllers.export_split.run_split_sheet,
                ),
                FeatureSpec(
                    "按行拆分",
                    "按数据行拆分为工作表或多个 Excel 文件",
                    "btn_split_rows",
                    controllers.export_split.run_split_rows,
                ),
            ),
        ),
        FeatureGroup(
            "maintenance",
            "批量维护",
            "维",
            "集中处理命名、链接、文本替换和工作表清理。",
            (
                FeatureSpec(
                    "批量重命名文件",
                    "通过规则预览并批量更新文件名称",
                    "btn_rename_files",
                    controllers.rename_files.run_batch_rename_files,
                ),
                FeatureSpec(
                    "批量重命名工作表",
                    "统一修改多个工作簿中的工作表名称",
                    "btn_rename_sheets",
                    controllers.rename_sheets.run_batch_rename_sheets,
                ),
                FeatureSpec(
                    "批量更换 Excel 链接",
                    "查找并替换多个文件中的外部链接地址",
                    "btn_link_replace",
                    controllers.link_replace.run_batch_link_replace,
                ),
                FeatureSpec(
                    "Word 批量替换",
                    "批量替换多个 Word 文件中的指定文字",
                    "btn_word_replace",
                    controllers.word_replace.run_word_batch_replace,
                ),
                FeatureSpec(
                    "批量删除工作表",
                    "按名称规则从多个工作簿中移除目标工作表",
                    "btn_delete_sheets",
                    controllers.delete_sheets.run_delete_sheets,
                ),
            ),
        ),
        FeatureGroup(
            "templates",
            "模板与取数",
            "模",
            "围绕模板、颜色和公式完成批量取数与报表生成。",
            (
                FeatureSpec(
                    "数据穿透查询",
                    "查询多文件、多工作表中的单元格或区域数值",
                    "btn_data_drill",
                    controllers.data_drill.run_data_drill,
                ),
                FeatureSpec(
                    "按模板批量生成 Excel",
                    "替换模板链接并一次生成多个目标 Excel 文件",
                    "btn_template_tb_report",
                    controllers.templates.run_template_generate,
                ),
                FeatureSpec(
                    "按颜色汇总求和",
                    "识别填充色位置并对多文件中的数值进行汇总",
                    "btn_color_sum",
                    controllers.color_tools.run_color_sum,
                ),
                FeatureSpec(
                    "按颜色清空内容",
                    "按指定填充色批量定位并清空单元格内容",
                    "btn_clear_by_color",
                    controllers.color_tools.run_clear_by_color,
                ),
                FeatureSpec(
                    "选区 ROUND 保留两位",
                    "为当前 Excel 选区公式统一增加两位小数处理",
                    "btn_round_formula",
                    controllers.workbook_misc.run_round_formula,
                ),
            ),
        ),
        FeatureGroup(
            "workbook",
            "工作簿辅助",
            "簿",
            "补齐常用的工作簿导航与结构整理能力。",
            (
                FeatureSpec(
                    "生成工作表目录",
                    "为当前工作簿生成带跳转链接的工作表索引",
                    "btn_sheet_index",
                    controllers.workbook_misc.run_sheet_index,
                ),
            ),
        ),
    )
