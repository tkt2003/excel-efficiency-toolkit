# -*- coding: utf-8 -*-
"""UI 设计 token 唯一来源：颜色、字体、窗口尺寸。

设计方向："冷静高对比的软蓝工作台" + 适老化可读性：
白色表面浮在冷灰蓝画布上，单一强调蓝贯穿主按钮/选中态，
正文对比度提高，整体字号加大一档（字号变化必须配对高度变化）。
除本模块外，ui/ 下不允许出现 hex 颜色或字体字面量。
"""
import customtkinter as ctk

# ---------------------------------------------------------------------------
# 颜色
# ---------------------------------------------------------------------------
BG = "#f3f6fa"                  # 窗口画布
SURFACE = "#ffffff"             # 卡片 / 面板 / 对话框表面
SURFACE_RAISED = "#fbfdff"      # 功能卡片、输入框、预览区表面
SURFACE_SUNKEN = "#f8fafc"      # 日志文本框背景
SURFACE_HOVER = "#eaf3fc"       # 功能卡片悬停
SURFACE_DISABLED = "#f6f8fa"    # 功能卡片禁用

BORDER = "#d8e0ea"              # 通用边框
BORDER_SOFT = "#e5edf5"         # 面板 / 对话框边框
BORDER_HOVER = "#a9c4e0"        # 功能卡片悬停边框
BORDER_DISABLED = "#e5eaf0"     # 功能卡片禁用边框
BORDER_BUTTON = "#dfe7f1"       # 头部次要按钮边框

TEXT = "#1b2a3b"                # 正文（白底约 13:1 对比度）
TEXT_STRONG = "#111827"         # 日志正文
TEXT_SECONDARY = "#44576d"      # 副标题 / 卡片描述（白底 ≥7:1）
TEXT_MUTED = "#64748b"          # 弱化说明文字
TEXT_DISABLED = "#9aa8b8"       # 禁用态文字（只用于禁用，不作正文）
TEXT_ON_ACCENT = "#ffffff"      # 主按钮文字

ACCENT = "#2f6bb0"              # 主强调色（主按钮、单选、复选、分区标题条）
ACCENT_HOVER = "#275b98"        # 主强调悬停（同色系加深，替换原饱和 #1d4ed8）
ACCENT_SOFT = "#dbeafe"         # 强调浅底

BUTTON_SECONDARY_BG = "#f8fafc"
BUTTON_SECONDARY_HOVER = "#e9eef5"
HEADER_BUTTON_HOVER = "#eef4fb"

SCROLLBAR = "#cbd5e1"
SCROLLBAR_HOVER = "#94a3b8"

# ---------------------------------------------------------------------------
# 字体（整体加大一档，配套高度见"尺寸"段）
# ---------------------------------------------------------------------------
FONT_FAMILY = "Microsoft YaHei UI"
FONT_MONO = "Consolas"

HEADER_TITLE_FONT = (FONT_FAMILY, 20, "bold")
HEADER_SUBTITLE_FONT = (FONT_FAMILY, 12)
HEADER_BUTTON_FONT = (FONT_FAMILY, 13)
SECTION_TITLE_FONT = (FONT_FAMILY, 17, "bold")
CARD_TITLE_FONT = (FONT_FAMILY, 14, "bold")
CARD_DESCRIPTION_FONT = (FONT_FAMILY, 13)
LOG_TITLE_FONT = (FONT_FAMILY, 13, "bold")
LOG_FONT = (FONT_MONO, 13)

DIALOG_TITLE_FONT = (FONT_FAMILY, 20, "bold")
DIALOG_BODY_FONT = (FONT_FAMILY, 16)
DIALOG_BUTTON_FONT = (FONT_FAMILY, 15)
ENTRY_FONT = (FONT_FAMILY, 13)
LABEL_FONT = (FONT_FAMILY, 13)
LABEL_BOLD_FONT = (FONT_FAMILY, 13, "bold")
SMALL_FONT = (FONT_FAMILY, 12)
PREVIEW_FONT = (FONT_MONO, 12)

# ---------------------------------------------------------------------------
# 尺寸（与字号配对：卡片/按钮加高避免中文被裁）
# ---------------------------------------------------------------------------
WINDOW_GEOMETRY = "1200x800"
WINDOW_MIN_SIZE = (1140, 720)

CARD_HEIGHT = 58                # 无描述功能卡片高度
CARD_HEIGHT_WITH_DESC = 80      # 带描述功能卡片高度
DIALOG_BUTTON_HEIGHT = 40
ENTRY_HEIGHT = 36
LOG_AREA_HEIGHT = 150
LOG_TEXT_HEIGHT = 96
SECTION_BAR_WIDTH = 3           # 分区标题左侧强调条宽度
SECTION_BAR_HEIGHT = 20


def apply_appearance():
    """设置 CustomTkinter 全局外观，必须在创建任何窗口前调用。"""
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
