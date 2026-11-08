"""Excel 读写辅助：文本归一化、日期识别、颜色识别与合并单元格取值。"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

EXCEL_SUFFIXES = (".xlsx", ".xlsm")


def load_readonly_workbook(path: str | Path):
    """以只读语义打开工作簿（不落盘、不修改源文件）。

    颜色与合并单元格区域在 openpyxl 的 read_only 模式下不可用，
    因此这里使用普通加载 + data_only，全程不调用 save。
    """
    return load_workbook(filename=str(path), data_only=True, keep_links=False)


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_code(value: Any) -> str:
    """产品代号比较用的归一化文本：去空白、统一大写、去掉整数小数尾。"""
    text = cell_text(value)
    if not text:
        return ""
    text = re.sub(r"\s+", "", text).upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", cell_text(value)).replace("\u3000", "")


def parse_date_value(value: Any, reference: date | None = None) -> date | None:
    """尽可能识别表头中的日期，支持 datetime、Excel 序列号及常见中文格式。"""
    reference = reference or date.today()

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)) and 20_000 <= float(value) <= 80_000:
        try:
            return from_excel(float(value)).date()
        except (ValueError, OverflowError):
            return None

    if not isinstance(value, str):
        return None

    text = value.strip()
    if not text:
        return None

    text = re.sub(r"[（(].*?[）)]", "", text).strip()

    full_patterns = (
        (r"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?", "ymd"),
        (r"^\s*(\d{4})(\d{2})(\d{2})\s*$", "ymd"),
        (r"^\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?\s*$", "ymd"),
    )
    for pattern, kind in full_patterns:
        match = re.search(pattern, text)
        if match:
            try:
                year, month, day = (int(part) for part in match.groups())
                return date(year, month, day)
            except ValueError:
                continue

    # 月/日格式只用于纯日期文本，避免把“3B901010-2”之类的产品代号误判成日期
    if re.search(r"[A-Za-z]", text):
        return None

    month_day_patterns = (
        r"(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?",
        r"^(\d{1,2})\s*月\s*(\d{1,2})\s*$",
        r"^(\d{1,2})\s*日$",
    )
    for pattern in month_day_patterns:
        match = re.search(pattern, text)
        if not match:
            continue
        try:
            if len(match.groups()) == 1:
                month, day = reference.month, int(match.group(1))
            else:
                month, day = int(match.group(1)), int(match.group(2))
        except ValueError:
            continue
        candidates = []
        for delta_year in (-1, 0, 1):
            try:
                candidates.append(date(reference.year + delta_year, month, day))
            except ValueError:
                continue
        if candidates:
            return min(candidates, key=lambda item: abs((item - reference).days))
        return None

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def color_key(cell) -> str | None:
    """提取单元格填充色；白色与无填充视为空。"""
    fill = getattr(cell, "fill", None)
    if fill is None or fill.fill_type != "solid":
        return None
    foreground = getattr(fill, "fgColor", None)
    if foreground is None:
        return None
    color_type = getattr(foreground, "type", None)
    if color_type == "rgb":
        value = foreground.rgb
        if not isinstance(value, str):
            return None
        hex_value = value.upper()
        rgb = hex_value[-6:]
        if rgb == "FFFFFF" or value in ("00000000", "00FFFFFF", "FFFFFFFF"):
            return None
        return rgb
    if color_type == "theme":
        theme = getattr(foreground, "theme", None)
        if theme is None:
            return None
        tint = round(float(getattr(foreground, "tint", 0.0) or 0.0), 4)
        return f"theme:{theme}:{tint}"
    if color_type == "indexed":
        indexed = getattr(foreground, "indexed", None)
        if indexed is None:
            return None
        return f"indexed:{indexed}"
    return None


def find_merged_range(merged_ranges: Iterable, row: int, col: int):
    for merged in merged_ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= col <= merged.max_col
        ):
            return merged
    return None


def merged_value(ws, merged_ranges: Iterable, row: int, col: int):
    """读取单元格值；若位于合并区域内则返回合并区域左上角的值。"""
    merged = find_merged_range(merged_ranges, row, col)
    if merged is not None:
        return ws.cell(merged.min_row, merged.min_col).value
    return ws.cell(row, col).value
