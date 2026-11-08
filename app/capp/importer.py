"""控制计划 Excel 导入：多级表头拼字段名、定位数据区。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import get_column_letter

from app.utils.excel_utils import (
    EXCEL_SUFFIXES,
    cell_text,
    find_merged_range,
    load_readonly_workbook,
    merged_value,
)

KEY_HEADER = "零件/过程编号"
EMPTY_ROW_LIMIT = 20
MAX_DATA_ROWS = 5000

TARGET_KEYWORDS: dict[str, tuple[str, ...]] = {
    "过程编号（工序号）": ("零件/过程编号", "过程编号", "工序号"),
    "过程名称": ("过程名称", "工序名称"),
    "机器、装置夹具、工装": ("机器、装置夹具、工装", "机器", "装置", "夹具", "工装"),
    "特性编号": ("特性_编号", "特性编号", "特性号"),
    "产品特性": ("特性_产品", "产品特性"),
    "过程特性": ("特性_过程", "过程特性"),
    "特殊特性分类（特殊特性）": ("特殊特性分类", "特殊特性"),
    "规范/公差": ("规范/公差", "规范", "公差"),
    "评价测量技术（评价测量技术）": ("评价测量技术", "测量技术"),
    "样本数量": ("样本_容量", "样本容量", "样本数量", "容量"),
    "样本频率": ("样本_频率", "样本频率", "频率"),
    "控制方法": ("控制方法", "方法"),
    "反应计划": ("反应计划",),
}


class ControlPlanError(Exception):
    pass


@dataclass(frozen=True)
class SourceColumn:
    index: int
    name: str
    letter: str

    @property
    def display(self) -> str:
        return f"{self.letter} · {self.name}"


@dataclass
class ControlPlanImport:
    file_path: str
    sheet_name: str
    columns: list[SourceColumn]
    rows: list[list[str]]
    header_start: int
    header_end: int
    warnings: list[str] = field(default_factory=list)


def _normalize(value) -> str:
    text = cell_text(value)
    return re.sub(r"\s+", "", text)


def _clean_header(value) -> str:
    text = cell_text(value)
    text = re.sub(r"[\r\n\t]+", " ", text)
    return re.sub(r"\s{2,}", " ", text).strip()


def _find_key_cell(worksheet):
    max_row = min(worksheet.max_row, 40)
    max_col = min(worksheet.max_column, 80)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            text = _normalize(worksheet.cell(row, col).value)
            if not text:
                continue
            if KEY_HEADER.replace(" ", "") in text:
                return row, col
            if "零件" in text and "过程编号" in text:
                return row, col
    return None


def _build_columns(worksheet, merged_ranges, start_row: int, end_row: int, max_col: int):
    columns: list[SourceColumn] = []
    seen_names: set[str] = set()
    for col in range(1, max_col + 1):
        parts: list[str] = []
        for row in range(start_row, end_row + 1):
            text = _clean_header(merged_value(worksheet, merged_ranges, row, col))
            if not text:
                continue
            if parts and parts[-1] == text:
                continue
            if text in parts:
                continue
            parts.append(text)
        if not parts:
            continue
        name = "_".join(parts)
        # 控制计划里同一个表头可能在多列重复出现，映射时只保留首次出现的一列
        key = re.sub(r"\s+", "", name).lower()
        if key in seen_names:
            continue
        seen_names.add(key)
        columns.append(SourceColumn(index=col, name=name, letter=get_column_letter(col)))
    if not columns:
        raise ControlPlanError("未能识别任何表头列，请检查控制计划表头。")
    return columns


def _value_text(worksheet, merged_ranges, row: int, col: int) -> str:
    """读取数据单元格文本；百分比格式的数值按 Excel 显示值输出（1 → 100%）。"""
    merged = find_merged_range(merged_ranges, row, col)
    cell = (
        worksheet.cell(merged.min_row, merged.min_col)
        if merged is not None
        else worksheet.cell(row, col)
    )
    value = cell.value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number_format = (cell.number_format or "").replace(" ", "")
        if "%" in number_format:
            text = f"{value * 100:.10f}".rstrip("0").rstrip(".")
            return f"{text}%"
    return cell_text(value)


def _row_has_data(values: list[str]) -> bool:
    return any(value.strip() for value in values)


def import_control_plan(file_path: str | Path) -> ControlPlanImport:
    path = Path(file_path)
    if not path.is_file():
        raise ControlPlanError(f"文件不存在：{path}")
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise ControlPlanError("仅支持 .xlsx / .xlsm 控制计划文件。")

    try:
        workbook = load_readonly_workbook(path)
    except Exception as exc:  # noqa: BLE001
        raise ControlPlanError(f"无法打开控制计划文件：{exc}") from exc

    try:
        worksheet = None
        key_cell = None
        for candidate in workbook.worksheets:
            key_cell = _find_key_cell(candidate)
            if key_cell is not None:
                worksheet = candidate
                break
        if worksheet is None or key_cell is None:
            raise ControlPlanError(
                "未找到“零件/过程编号”列，请确认文件为控制计划格式。"
            )

        merged_ranges = list(worksheet.merged_cells.ranges)
        key_row, key_col = key_cell
        merged = find_merged_range(merged_ranges, key_row, key_col)
        if merged is not None:
            header_start, header_end = merged.min_row, merged.max_row
        else:
            header_start = header_end = key_row

        max_col = worksheet.max_column
        columns = _build_columns(
            worksheet, merged_ranges, header_start, header_end, max_col
        )

        rows: list[list[str]] = []
        empty_run = 0
        for row in range(header_end + 1, worksheet.max_row + 1):
            values = [
                _value_text(worksheet, merged_ranges, row, column.index)
                for column in columns
            ]
            if not _row_has_data(values):
                empty_run += 1
                if empty_run >= EMPTY_ROW_LIMIT:
                    break
                continue
            empty_run = 0
            rows.append(values)
            if len(rows) >= MAX_DATA_ROWS:
                break

        if not rows:
            raise ControlPlanError("表头下方未读取到数据行，请检查控制计划内容。")

        warnings: list[str] = []
        if merged is None:
            warnings.append(
                "“零件/过程编号”单元格不是合并单元格，已按单行表头解析。"
            )
        return ControlPlanImport(
            file_path=str(path),
            sheet_name=worksheet.title,
            columns=columns,
            rows=rows,
            header_start=header_start,
            header_end=header_end,
            warnings=warnings,
        )
    finally:
        workbook.close()


def auto_mapping(columns: list[SourceColumn]) -> dict[str, int | None]:
    """按关键词自动匹配目标字段 → 源列 index，匹配不上为 None。"""
    mapping: dict[str, int | None] = {}
    used: set[int] = set()
    for target, keywords in TARGET_KEYWORDS.items():
        best_index: int | None = None
        best_score = 0
        for column in columns:
            if column.index in used:
                continue
            name = _normalize(column.name)
            score = 0
            for keyword in keywords:
                normalized_keyword = _normalize(keyword)
                if not normalized_keyword:
                    continue
                if name == normalized_keyword:
                    score = max(score, 120)
                elif normalized_keyword in name:
                    score = max(score, 100 - len(name) // 10)
            if score > best_score:
                best_score = score
                best_index = column.index
        if best_index is not None:
            mapping[target] = best_index
            used.add(best_index)
        else:
            mapping[target] = None
    return mapping
