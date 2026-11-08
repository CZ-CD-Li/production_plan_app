"""SPC 计划解析：定位序号行，识别数据范围、单位与量具精度。"""

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

HEADER_TOKEN = "序号"
EMPTY_ROW_LIMIT = 3
MAX_ITEMS = 300

_NUM = r"(\d+(?:\.\d+)?)"
_PM_RE = re.compile(_NUM + r"\s*±\s*" + _NUM)
_RANGE_RE = re.compile(_NUM + r"\s*[-~～至]\s*" + _NUM)
_PLAIN_RE = re.compile(r"^" + _NUM + r"$")
_COMPARATORS = ("≥", ">=", "＞=", ">", "＞", "≤", "<=", "＜=", "<", "＜")

_UNIT_TAIL_RE = re.compile(r"([A-Za-zμ℃°%/·件]+)\s*$")


class SpcPlanError(Exception):
    pass


@dataclass
class ParsedSpec:
    lower: float | None = None
    target: float | None = None
    upper: float | None = None
    unit: str = ""
    recognized: bool = False
    note: str = ""
    kind: str = ""  # pm=中心值±公差，range=上下限，lower=只有下限，upper=只有上限


@dataclass
class SpcColumn:
    """标题栏中的一列，供用户选择公差对应栏。"""

    index: int
    letter: str
    title: str

    @property
    def label(self) -> str:
        return f"{self.letter} 列 · {self.title}"


@dataclass
class SpcItem:
    row: int
    index: str
    description: str
    gauge_name: str
    gauge_spec: str
    precision_text: str
    precision: float
    lower: float | None = None
    target: float | None = None
    upper: float | None = None
    unit: str = ""
    note: str = ""
    spec_kind: str = ""
    chart_type: str = ""

    @property
    def ready(self) -> bool:
        return self.problem == ""

    @property
    def problem(self) -> str:
        if self.lower is None or self.target is None or self.upper is None:
            return "数据范围未识别，请人工填写下限/目标值/上限"
        if not (self.lower < self.target < self.upper):
            return "需满足 下限 < 目标值 < 上限"
        if not self.unit.strip():
            return "单位未识别，请人工填写单位"
        if not (self.precision > 0):
            return "量具精度需为大于 0 的数字"
        return ""

    @property
    def spec_text(self) -> str:
        if self.lower is None or self.upper is None:
            return ""
        unit = self.unit.strip()
        return f"{_fmt(self.lower)}~{_fmt(self.upper)}{unit}"


@dataclass
class SpcPlanImport:
    file_path: str
    sheet_name: str
    items: list[SpcItem] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    part_name: str = ""
    part_number: str = ""
    analysis_col: int | None = None
    conclusion_col: int | None = None
    columns: list[SpcColumn] = field(default_factory=list)
    tolerance_col: int | None = None


def _fmt(value: float) -> str:
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return text or "0"


def _normalize_unit(token: str) -> str:
    """保留源表写法（KN / Kg / N 等），仅去掉空白。"""
    return token.strip().replace(" ", "")


def parse_spec(text: str) -> ParsedSpec:
    """解析特性描述末尾的范围与单位，如 ≥40KN、356±56N、≥1200-1500Kg。"""
    raw = cell_text(text)
    if not raw:
        return ParsedSpec(note="特性描述为空")

    # 只取最后一个冒号后的数值部分，避免把 “0.1m/s” 这类工况写成单位
    body = re.split(r"[：:]", raw)[-1].replace("\n", " ").strip()
    if not body:
        body = raw.replace("\n", " ").strip()

    unit = ""
    match = _UNIT_TAIL_RE.search(body)
    if match:
        unit = _normalize_unit(match.group(1))
        if unit:
            body = body[: match.start()].strip()

    body = body.replace(" ", "").replace("，", "").replace(",", "")
    qualifier = ""
    for token in _COMPARATORS:
        if body.startswith(token):
            qualifier = ">=" if token in ("≥", ">=", "＞=", ">", "＞") else "<="
            body = body[len(token):].strip()
            break

    pm = _PM_RE.search(body)
    if pm:
        center = float(pm.group(1))
        tolerance = float(pm.group(2))
        return ParsedSpec(
            lower=round(center - tolerance, 6),
            target=center,
            upper=round(center + tolerance, 6),
            unit=unit,
            recognized=bool(unit),
            note="" if unit else "未识别到单位",
            kind="pm",
        )

    rng = _RANGE_RE.search(body)
    if rng:
        low, high = sorted((float(rng.group(1)), float(rng.group(2))))
        return ParsedSpec(
            lower=low,
            target=round((low + high) / 2, 6),
            upper=high,
            unit=unit,
            recognized=bool(unit),
            note="" if unit else "未识别到单位",
            kind="range",
        )

    plain = _PLAIN_RE.match(body)
    if plain and qualifier:
        value = float(plain.group(1))
        # 单边公差默认目标值 ±5、公差 ±5：≥X → X~X+10，≤X → X-10~X
        if qualifier == ">=":
            lower, target, upper = value, value + 5, value + 10
            kind = "lower"
        else:
            lower, target, upper = value - 10, value - 5, value
            kind = "upper"
        return ParsedSpec(
            lower=lower,
            target=target,
            upper=upper,
            unit=unit,
            recognized=bool(unit),
            note="" if unit else "未识别到单位",
            kind=kind,
        )

    hint = "未识别到数据范围"
    if not unit:
        hint += "及单位"
    return ParsedSpec(unit=unit, recognized=False, note=hint)


def _extract_precision(text: str) -> tuple[float, str]:
    raw = cell_text(text)
    match = re.search(r"(\d+(?:\.\d+)?)", raw)
    if not match:
        return 0.1, raw
    try:
        value = float(match.group(1))
    except ValueError:
        return 0.1, raw
    return (value if value > 0 else 0.1), raw


def _find_sequence_start(worksheet) -> tuple[int, int] | None:
    """定位内容为 1 且下一行内容为 2 的单元格，返回 (行, 列)。"""
    limit_row = min(worksheet.max_row, 200)
    limit_col = min(worksheet.max_column, 20)
    for row in range(1, limit_row):
        for col in range(1, limit_col + 1):
            if cell_text(worksheet.cell(row, col).value) != "1":
                continue
            if cell_text(worksheet.cell(row + 1, col).value) == "2":
                return row, col
    return None


def _header_title(
    worksheet, merged_ranges, header_row: int, col: int
) -> str:
    """读取标题栏单元格文本；本行无内容时向上找一级表头。"""
    for row in (header_row, header_row - 1):
        if row < 1:
            break
        text = re.sub(
            r"\s+",
            "",
            cell_text(merged_value(worksheet, merged_ranges, row, col)),
        )
        if text:
            return text
    return ""


def _find_header_columns(
    worksheet, merged_ranges, header_row: int, index_col: int
) -> list[SpcColumn]:
    """列出序号列右侧带标题的列，供用户选择公差对应栏。"""
    limit_col = min(worksheet.max_column, 30)
    columns: list[SpcColumn] = []
    seen: set[int] = set()
    for col in range(index_col + 1, limit_col + 1):
        title = _header_title(worksheet, merged_ranges, header_row, col)
        if not title:
            continue
        merged = find_merged_range(merged_ranges, header_row, col)
        anchor = merged.min_col if merged is not None else col
        if anchor == index_col or anchor in seen:
            continue
        seen.add(anchor)
        columns.append(
            SpcColumn(
                index=anchor,
                letter=get_column_letter(anchor),
                title=title,
            )
        )
    return columns


def _find_gauge_columns(
    worksheet, merged_ranges, header_row: int, index_col: int
) -> tuple[int, int, int]:
    """按标题栏名称定位量具的名称/规格/精度列，找不到时按模板默认列回退。"""
    default = (index_col + 3, index_col + 4, index_col + 5)
    found: dict[str, int] = {}
    limit_col = min(worksheet.max_column, index_col + 6)
    for col in range(index_col + 1, limit_col + 1):
        text = re.sub(
            r"\s+",
            "",
            cell_text(merged_value(worksheet, merged_ranges, header_row, col)),
        )
        if not text:
            continue
        if "精度" in text and "precision" not in found:
            found["precision"] = col
        elif "规格" in text and "spec" not in found:
            found["spec"] = col
        elif "名称" in text and "name" not in found:
            found["name"] = col
    name_col = found.get("name", default[0])
    spec_col = found.get("spec", default[1])
    precision_col = found.get("precision", default[2])
    return name_col, spec_col, precision_col


def _right_neighbor_text(worksheet, merged_ranges, row: int, col: int) -> str:
    """读取标签单元格右侧的值（标签或值本身是合并单元格时按合并区域处理）。"""
    merged = find_merged_range(merged_ranges, row, col)
    start_col = merged.max_col + 1 if merged is not None else col + 1
    for offset in range(3):
        target_col = start_col + offset
        if target_col > worksheet.max_column:
            break
        text = cell_text(
            merged_value(worksheet, merged_ranges, row, target_col)
        ).replace("\n", " ").strip()
        if text:
            return text
    return ""


def _find_part_info(
    worksheet, merged_ranges, header_row: int
) -> tuple[str, str]:
    """在表头上方查找零部件名称/件号标签，取标签右侧单元格内容。"""
    part_name = ""
    part_number = ""
    limit_col = min(worksheet.max_column, 12)
    for row in range(1, header_row):
        for col in range(1, limit_col + 1):
            text = re.sub(
                r"\s+",
                "",
                cell_text(merged_value(worksheet, merged_ranges, row, col)),
            )
            if not text:
                continue
            if not part_name and any(
                label in text for label in ("零部件名称", "零件名称")
            ):
                part_name = _right_neighbor_text(
                    worksheet, merged_ranges, row, col
                )
            elif not part_number and any(
                label in text
                for label in ("零部件件号", "零件件号", "零件号", "零部件号")
            ):
                part_number = _right_neighbor_text(
                    worksheet, merged_ranges, row, col
                )
        if part_name and part_number:
            break
    return part_name, part_number


def _find_header_column(
    worksheet,
    merged_ranges,
    header_row: int,
    keyword: str,
    exclude_col: int | None = None,
) -> int | None:
    """按表头关键字定位列，找不到时返回 None。"""
    limit_col = min(worksheet.max_column, 20)
    for col in range(1, limit_col + 1):
        if exclude_col is not None and col == exclude_col:
            continue
        text = re.sub(
            r"\s+",
            "",
            cell_text(merged_value(worksheet, merged_ranges, header_row, col)),
        )
        if keyword not in text:
            continue
        merged = find_merged_range(merged_ranges, header_row, col)
        return merged.min_col if merged is not None else col
    return None


def _find_analysis_column(
    worksheet, merged_ranges, header_row: int
) -> int | None:
    """定位“分析结果”列，找不到时返回 None。"""
    return _find_header_column(
        worksheet, merged_ranges, header_row, "分析结果"
    )


def import_spc_plan(
    file_path: str | Path, tolerance_col: int | None = None
) -> SpcPlanImport:
    """解析 SPC 计划：先定位序号列与标题栏，选定公差栏后再解析各项数据。"""
    path = Path(file_path)
    if not path.is_file():
        raise SpcPlanError(f"文件不存在：{path}")
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise SpcPlanError("仅支持 .xlsx / .xlsm SPC 计划文件。")

    try:
        workbook = load_readonly_workbook(path)
    except Exception as exc:  # noqa: BLE001
        raise SpcPlanError(f"无法打开 SPC 计划文件：{exc}") from exc

    try:
        candidates = [
            sheet for sheet in workbook.worksheets if "模板" in sheet.title
        ] + [sheet for sheet in workbook.worksheets if "模板" not in sheet.title]
        worksheet = None
        anchor = None
        for sheet in candidates:
            found = _find_sequence_start(sheet)
            if found is not None:
                worksheet = sheet
                anchor = found
                break
        if worksheet is None or anchor is None:
            raise SpcPlanError(
                "未找到内容为 1 且下一行为 2 的序号单元格，"
                "请确认文件为 SPC 计划格式。"
            )

        seq_row, index_col = anchor
        header_row = seq_row - 1
        if header_row < 1:
            raise SpcPlanError("序号 1 上方没有标题栏，请检查 SPC 计划内容。")
        merged_ranges = list(worksheet.merged_cells.ranges)
        columns = _find_header_columns(
            worksheet, merged_ranges, header_row, index_col
        )
        if not columns:
            raise SpcPlanError("未识别到标题栏内容，请检查 SPC 计划表头。")
        if tolerance_col is not None and tolerance_col not in {
            column.index for column in columns
        }:
            raise SpcPlanError("所选公差栏不在标题栏范围内，请重新选择。")

        name_col, spec_col, precision_col = _find_gauge_columns(
            worksheet, merged_ranges, header_row, index_col
        )
        part_name, part_number = _find_part_info(
            worksheet, merged_ranges, header_row
        )
        analysis_col = _find_analysis_column(
            worksheet, merged_ranges, header_row
        )
        conclusion_col = _find_header_column(
            worksheet, merged_ranges, header_row, "结论", exclude_col=analysis_col
        )
        items: list[SpcItem] = []
        if tolerance_col is not None:
            items = _read_items(
                worksheet,
                merged_ranges,
                seq_row,
                index_col,
                tolerance_col,
                name_col,
                spec_col,
                precision_col,
            )
            if not items:
                raise SpcPlanError("未读取到任何序号行数据，请检查 SPC 计划内容。")

        warnings = [
            f"序号 {item.index} 未识别：{item.note}"
            for item in items
            if not item.ready and item.note
        ]
        return SpcPlanImport(
            file_path=str(path),
            sheet_name=worksheet.title,
            items=items,
            warnings=warnings,
            part_name=part_name,
            part_number=part_number,
            analysis_col=analysis_col,
            conclusion_col=conclusion_col,
            columns=columns,
            tolerance_col=tolerance_col,
        )
    finally:
        workbook.close()


def _read_items(
    worksheet,
    merged_ranges,
    seq_row: int,
    index_col: int,
    tolerance_col: int,
    name_col: int,
    spec_col: int,
    precision_col: int,
) -> list[SpcItem]:
    """按用户选定的公差栏逐行解析数据范围、单位与量具精度。"""
    items: list[SpcItem] = []
    empty_streak = 0
    current_index = ""
    max_row = min(worksheet.max_row, seq_row + MAX_ITEMS + 10)
    for row in range(seq_row, max_row + 1):
        index_text = cell_text(
            merged_value(worksheet, merged_ranges, row, index_col)
        ).replace("\n", " ").strip()
        description = cell_text(
            merged_value(worksheet, merged_ranges, row, tolerance_col)
        ).replace("\n", " ").strip()
        if not index_text and not description:
            empty_streak += 1
            if empty_streak >= EMPTY_ROW_LIMIT:
                break
            continue
        empty_streak = 0
        if re.match(r"^\d+(?:\.\d+)?$", index_text):
            current_index = index_text
        elif not index_text and current_index:
            # 序号列合并时，后续行沿用合并区域内的序号
            pass
        else:
            # 表头分隔行、页脚（供应商提交人、备注等）不是数据项
            continue
        if not description:
            continue
        gauge_name = cell_text(
            merged_value(worksheet, merged_ranges, row, name_col)
        ).replace("\n", " ").strip()
        gauge_spec = cell_text(
            merged_value(worksheet, merged_ranges, row, spec_col)
        ).replace("\n", " ").strip()
        precision_text = cell_text(
            merged_value(worksheet, merged_ranges, row, precision_col)
        ).replace("\n", " ").strip()
        precision, _ = _extract_precision(precision_text)
        parsed = parse_spec(description)
        items.append(
            SpcItem(
                row=row,
                index=index_text or current_index,
                description=description,
                gauge_name=gauge_name,
                gauge_spec=gauge_spec,
                precision_text=precision_text,
                precision=precision,
                lower=parsed.lower,
                target=parsed.target,
                upper=parsed.upper,
                unit=parsed.unit,
                note=parsed.note,
                spec_kind=parsed.kind,
            )
        )
    return items
