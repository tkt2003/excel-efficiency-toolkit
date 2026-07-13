# -*- coding: utf-8 -*-
"""终端/日志表格显示宽度工具：按东亚字符宽度对齐、截断和填充单元格文本。"""
import unicodedata


def char_display_width(char) -> int:
    return 2 if unicodedata.east_asian_width(char) in {"F", "W"} else 1


def display_width(text) -> int:
    return sum(char_display_width(char) for char in str(text))


def take_display_prefix(text, width) -> str:
    result = []
    current_width = 0
    for char in text:
        char_width = char_display_width(char)
        if current_width + char_width > width:
            break
        result.append(char)
        current_width += char_width
    return "".join(result)


def take_display_suffix(text, width) -> str:
    result = []
    current_width = 0
    for char in reversed(text):
        char_width = char_display_width(char)
        if current_width + char_width > width:
            break
        result.append(char)
        current_width += char_width
    return "".join(reversed(result))


def middle_truncate_display(text, width) -> str:
    if width <= 3:
        return "." * width
    target = width - 3
    left_width = target // 2
    right_width = target - left_width
    left = take_display_prefix(text, left_width)
    right = take_display_suffix(text, right_width)
    return f"{left}...{right}"


def compact_log_cell(value) -> str:
    """把换行压成空格并截断到 80 字符，用于单行日志。"""
    return str(value).replace("\r", " ").replace("\n", " ")[:80]


def fit_display_cell(value, width) -> str:
    text = compact_log_cell(value)
    if display_width(text) > width:
        text = middle_truncate_display(text, width)
    padding = max(0, width - display_width(text))
    return text + (" " * padding)


def format_display_row(values, columns) -> str:
    """按 columns=[(标题, 宽度), ...] 对齐一行值。"""
    cells = []
    for value, (_, width) in zip(values, columns):
        cells.append(fit_display_cell(value, width))
    return "  ".join(cells)
