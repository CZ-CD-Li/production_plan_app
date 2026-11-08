"""作业指导书解析结果 → 主机厂控制计划模板写出。

模板 ``控制计划格式样板-3.xlsx`` 每个 sheet 对应一个主机厂格式：
前若干行是主机厂表头，``**起始行`` 标定数据位置，第 27/28 行是特殊特性
符号，第 30 行是 VBA 使用的 18 个标准列名，第 31 行起是数据区。

写出时按 VBA 的做法：识别第 30 行表头 → 从数据区原型行取样式 → 写入数据
并合并单元格、设置行高 → 删除 ``**起始行`` 到第 30 行 → 调整打印区域和页脚。
为了保留模板中的 logo 图片、复选框、OLE 对象和页面设置，这里直接修改
xlsx 包内目标 sheet 的 XML 字符串，而不是用 openpyxl 另存工作簿。
"""

from __future__ import annotations

import math
import posixpath
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl.utils import get_column_letter

from app.capp.work_instruction import (
    PLAN_HEADERS,
    SymbolValue,
    WorkInstructionParse,
)
from app.config import DATA_DIR, ROOT_DIR

TEMPLATE_FILENAME = "控制计划格式样板-3.xlsx"
TEMPLATE_DIR = DATA_DIR / "templates"
TEMPLATE_PATH = TEMPLATE_DIR / TEMPLATE_FILENAME

HEADER_ROW = 30
DATA_START_ROW = 31
MPR_TEXT = "**起始行"

SINGLE_LINE_HEIGHT = 11
HEIGHT_PADDING = 4
CHAR_WIDTH = 1.75
MERGE_RESERVE = 1.0
MERGE_OVER_RESERVE = 1.4
MERGE_THRESHOLD = 5.0
MAX_HEIGHT_LINES = (400 - HEIGHT_PADDING) / SINGLE_LINE_HEIGHT

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_CORE_COLUMNS = PLAN_HEADERS[:13]

_ROW_PATTERN = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
_ROW_NUMBER_PATTERN = re.compile(r'<row\b[^>]*?\br="(\d+)"')
_REL_PATTERN = re.compile(r"<Relationship\b[^>]*?(?:/>|>.*?</Relationship>)", re.S)
_OVERRIDE_PATTERN = re.compile(r"<Override\b[^>]*/>")


def _q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


class ControlPlanGenerateError(Exception):
    """无法按模板生成控制计划。"""


@dataclass(frozen=True)
class TemplateColumn:
    target: int
    column: int
    span: int


@dataclass
class TemplateLayout:
    sheet_name: str
    display_name: str
    marker_row: int
    columns: dict[int, TemplateColumn]
    symbols: dict[int, str]
    supported: bool
    reason: str = ""
    missing_names: tuple[str, ...] = ()


@dataclass
class SheetInfo:
    name: str
    path: str
    index: int
    layout: TemplateLayout
    max_col: int
    max_row: int
    footer_start: int | None
    widths: dict[int, float]
    default_width: float
    styles: dict[int, str]
    merge_refs: list[str]
    footer_merges: list[str]
    print_area: str | None
    print_titles: str | None


_INFO_CACHE: dict[str, tuple[int, dict[str, SheetInfo], dict[str, bytes]]] = {}


# ---------------------------------------------------------------- 模板定位


def _template_sources() -> list[Path]:
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass) / "data" / "templates" / TEMPLATE_FILENAME)
        candidates.append(Path(meipass) / "templates" / TEMPLATE_FILENAME)
        candidates.append(Path(meipass) / TEMPLATE_FILENAME)
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / TEMPLATE_FILENAME)
    candidates.append(ROOT_DIR / TEMPLATE_FILENAME)
    candidates.append(Path(__file__).resolve().parent / TEMPLATE_FILENAME)
    return candidates


def ensure_template_file() -> Path:
    if TEMPLATE_PATH.is_file():
        return TEMPLATE_PATH
    for source in _template_sources():
        if source.is_file():
            TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, TEMPLATE_PATH)
            return TEMPLATE_PATH
    raise ControlPlanGenerateError(
        f"找不到控制计划模板“{TEMPLATE_FILENAME}”，"
        "请将模板放到程序目录或 data/templates 下。"
    )


# ---------------------------------------------------------------- 包解析


def _col_index(ref: str) -> int:
    match = re.match(r"\$?([A-Z]+)", ref or "")
    if not match:
        return 0
    result = 0
    for char in match.group(1):
        result = result * 26 + (ord(char) - 64)
    return result


def _row_index(ref: str) -> int:
    match = re.search(r"\$?(\d+)$", ref or "")
    return int(match.group(1)) if match else 0


def _parse_range(ref: str) -> tuple[int, int, int, int] | None:
    if not ref:
        return None
    parts = ref.replace("$", "").split(":")
    if len(parts) == 1:
        row, col = _row_index(parts[0]), _col_index(parts[0])
        return (row, col, row, col) if row and col else None
    if len(parts) != 2:
        return None
    min_row, min_col = _row_index(parts[0]), _col_index(parts[0])
    max_row, max_col = _row_index(parts[1]), _col_index(parts[1])
    if not (min_row and min_col and max_row and max_col):
        return None
    return min_row, min_col, max_row, max_col


def _range_ref(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    start = f"{get_column_letter(min_col)}{min_row}"
    if (min_row, min_col) == (max_row, max_col):
        return start
    return f"{start}:{get_column_letter(max_col)}{max_row}"


def _resolve_target(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _rels_path(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", name + ".rels")


def _read_rels(parts: dict[str, bytes], part: str) -> dict[str, tuple[str, str]]:
    path = _rels_path(part)
    if path not in parts:
        return {}
    root = ET.fromstring(parts[path])
    result: dict[str, tuple[str, str]] = {}
    for rel in root:
        rel_id = rel.get("Id", "")
        rel_type = rel.get("Type", "")
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            result[rel_id] = (rel_type, target)
            continue
        result[rel_id] = (rel_type, _resolve_target(part, target))
    return result


def _read_shared_strings(parts: dict[str, bytes]) -> list[str]:
    raw = parts.get("xl/sharedStrings.xml")
    if not raw:
        return []
    root = ET.fromstring(raw)
    return [
        "".join(node.text or "" for node in item.iter(_q("t")))
        for item in root.findall(_q("si"))
    ]


def _workbook_sheets(parts: dict[str, bytes]) -> list[tuple[str, str, str, int]]:
    root = ET.fromstring(parts["xl/workbook.xml"])
    rels = _read_rels(parts, "xl/workbook.xml")
    result: list[tuple[str, str, str, int]] = []
    sheets = root.find(_q("sheets"))
    if sheets is None:
        return result
    for index, sheet in enumerate(sheets.findall(_q("sheet"))):
        name = sheet.get("name", "")
        rel_id = sheet.get(f"{{{REL_NS}}}id", "")
        info = rels.get(rel_id)
        if info is None:
            continue
        try:
            sheet_id = int(sheet.get("sheetId", index + 1))
        except ValueError:
            sheet_id = index + 1
        result.append((name, info[1], rel_id, sheet_id))
    return result


def _defined_names(parts: dict[str, bytes]) -> dict[str, str]:
    root = ET.fromstring(parts["xl/workbook.xml"])
    names = root.find(_q("definedNames"))
    result: dict[str, str] = {}
    if names is None:
        return result
    for node in names.findall(_q("definedName")):
        key = f"{node.get('localSheetId', '')}:{node.get('name', '')}"
        result[key] = (node.text or "").strip()
    return result


def _cell_value(cell: ET.Element, shared: list[str]) -> Any:
    kind = cell.get("t")
    if kind == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(_q("t")))
    value = cell.find(_q("v"))
    if kind == "s":
        if value is None or value.text is None:
            return ""
        try:
            return shared[int(value.text)]
        except (ValueError, IndexError):
            return ""
    if value is None or value.text is None:
        return None
    if kind == "b":
        return value.text == "1"
    if kind in ("str", "e", "d"):
        return value.text
    text = value.text
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    try:
        return float(text)
    except ValueError:
        return text


def _normalize_header(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", "", value).lstrip("*")


# ---------------------------------------------------------------- 模板分析


def _analyze_sheet(
    parts: dict[str, bytes],
    shared: list[str],
    name: str,
    path: str,
    index: int,
) -> SheetInfo:
    root = ET.fromstring(parts[path])

    schema = root.find(_q("sheetFormatPr"))
    default_width = 8.43
    if schema is not None:
        try:
            default_width = float(schema.get("defaultColWidth", "8.43"))
        except ValueError:
            default_width = 8.43

    widths: dict[int, float] = {}
    cols = root.find(_q("cols"))
    if cols is not None:
        for col in cols.findall(_q("col")):
            try:
                start = int(col.get("min", "1"))
                end = int(col.get("max", str(start)))
                width = float(col.get("width", default_width))
            except ValueError:
                continue
            for column in range(start, end + 1):
                widths[column] = width

    values: dict[tuple[int, int], Any] = {}
    styles: dict[tuple[int, int], str] = {}
    for row in root.iter(_q("row")):
        try:
            row_number = int(row.get("r", "0"))
        except ValueError:
            continue
        for cell in row.findall(_q("c")):
            col_number = _col_index(cell.get("r", ""))
            if not col_number:
                continue
            values[(row_number, col_number)] = _cell_value(cell, shared)
            style = cell.get("s")
            if style:
                styles[(row_number, col_number)] = style

    actual_max_row = max((row for row, _ in values), default=0)
    actual_max_col = max((col for _, col in values), default=0)
    dimension = root.find(_q("dimension"))
    if dimension is not None and dimension.get("ref"):
        last = dimension.get("ref", "").split(":")[-1]
        actual_max_row = max(actual_max_row, _row_index(last))
        actual_max_col = max(actual_max_col, _col_index(last))

    marker_row = 0
    for (row, _col), value in sorted(values.items()):
        if isinstance(value, str) and MPR_TEXT in value:
            marker_row = row
            break

    footer_start: int | None = None
    for row in range(DATA_START_ROW, actual_max_row + 1):
        for col in range(1, actual_max_col + 1):
            value = values.get((row, col))
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            footer_start = row
            break
        if footer_start is not None:
            break

    merge_refs: list[str] = []
    header_spans: dict[int, int] = {}
    footer_merges: list[str] = []
    merges = root.find(_q("mergeCells"))
    if merges is not None:
        for merge in merges.findall(_q("mergeCell")):
            ref = merge.get("ref", "")
            parsed = _parse_range(ref)
            if parsed is None:
                continue
            min_row, min_col, max_row, max_col = parsed
            merge_refs.append(ref)
            if min_row == HEADER_ROW == max_row:
                header_spans[min_col] = max(
                    header_spans.get(min_col, 1), max_col - min_col + 1
                )
            if footer_start is not None and min_row >= footer_start:
                footer_merges.append(ref)

    headers: dict[int, tuple[int, str]] = {}
    header_edge_col = 0
    for col in range(1, actual_max_col + 1):
        value = values.get((HEADER_ROW, col))
        header = _normalize_header(value)
        if not header:
            continue
        header_edge_col = max(header_edge_col, col + header_spans.get(col, 1) - 1)
        if header not in PLAN_HEADERS:
            continue
        target = PLAN_HEADERS.index(header)
        if target not in headers:
            headers[target] = (col, header)

    columns: dict[int, TemplateColumn] = {}
    for target, (col, _header) in headers.items():
        columns[target] = TemplateColumn(
            target=target, column=col, span=header_spans.get(col, 1)
        )

    symbols: dict[int, str] = {}
    for symbol_col in range(1, 5):
        value = values.get((28, symbol_col))
        symbols[symbol_col] = str(value).strip() if value is not None else ""

    header_names = {
        _normalize_header(values.get((HEADER_ROW, col)))
        for col in range(1, actual_max_col + 1)
    }
    dual_layout = bool(header_names & {"过程特性产品特性", "产品项目"})
    missing_core = [
        plan_name
        for target, plan_name in enumerate(_CORE_COLUMNS)
        if target not in columns
    ]
    missing_extra = [
        plan_name
        for target, plan_name in enumerate(PLAN_HEADERS)
        if target >= 13 and target not in columns
    ]

    reason = ""
    if dual_layout:
        reason = "双栏（产品/过程分列）格式，VBS 逻辑不支持"
    elif marker_row == 0:
        reason = "未找到 **起始行 标定"
    elif missing_core:
        reason = "缺少标准列：" + "、".join(missing_core)

    display_name = name.replace("没完成", "").strip() or name
    layout = TemplateLayout(
        sheet_name=name,
        display_name=display_name,
        marker_row=marker_row,
        columns=columns,
        symbols=symbols,
        supported=not reason,
        reason=reason,
        missing_names=tuple(missing_extra),
    )

    data_styles: dict[int, str] = {}
    for col in range(1, actual_max_col + 1):
        for row in range(DATA_START_ROW, DATA_START_ROW + 3):
            style = styles.get((row, col))
            if style:
                data_styles[col] = style
                break

    names = _defined_names(parts)
    print_area = names.get(f"{index}:_xlnm.Print_Area")
    print_titles = names.get(f"{index}:_xlnm.Print_Titles")
    if print_area in ("#REF!", ""):
        print_area = None
    if print_titles in ("#REF!", ""):
        print_titles = None

    print_col = 0
    if print_area and "!" in print_area:
        print_col = _col_index(print_area.rsplit(":", 1)[-1])
    max_col = max(header_edge_col, print_col)
    if max_col <= 0:
        max_col = actual_max_col

    return SheetInfo(
        name=name,
        path=path,
        index=index,
        layout=layout,
        max_col=max_col,
        max_row=max(actual_max_row, 1),
        footer_start=footer_start,
        widths=widths,
        default_width=default_width,
        styles=data_styles,
        merge_refs=merge_refs,
        footer_merges=footer_merges,
        print_area=print_area,
        print_titles=print_titles,
    )


def _load_infos(
    path: Path | None = None, force: bool = False
) -> tuple[dict[str, SheetInfo], dict[str, bytes]]:
    source = path or ensure_template_file()
    key = str(source)
    stamp = source.stat().st_mtime_ns
    cached = _INFO_CACHE.get(key)
    if cached and cached[0] == stamp and not force:
        return cached[1], cached[2]

    with zipfile.ZipFile(source) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    shared = _read_shared_strings(parts)
    infos: dict[str, SheetInfo] = {}
    for index, (name, sheet_path, _rel_id, _sheet_id) in enumerate(
        _workbook_sheets(parts)
    ):
        if name == "辅助":
            continue
        infos[name] = _analyze_sheet(parts, shared, name, sheet_path, index)
    _INFO_CACHE[key] = (stamp, infos, parts)
    return infos, parts


def list_templates(force: bool = False) -> list[TemplateLayout]:
    infos, _parts = _load_infos(force=force)
    return [info.layout for info in infos.values()]


def get_layout(sheet_name: str) -> TemplateLayout:
    infos, _parts = _load_infos()
    if sheet_name not in infos:
        raise ControlPlanGenerateError(f"模板中不存在工作表“{sheet_name}”。")
    return infos[sheet_name].layout


# ---------------------------------------------------------------- 值处理


def resolve_value(value: Any, symbols: dict[int, str]) -> Any:
    if isinstance(value, SymbolValue):
        text = value.text
        symbol = symbols.get(value.column, "")
        return f"{symbol}/{text}" if symbol else text
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if value is None:
        return ""
    if isinstance(value, str):
        # VBA 在写完后把“本工艺卡片”“工艺卡片”替换为空
        return value.replace("本工艺卡片", "").replace("工艺卡片", "")
    return value


def preview_text(value: Any, symbols: dict[int, str]) -> str:
    resolved = resolve_value(value, symbols)
    if isinstance(resolved, float) and resolved.is_integer():
        resolved = int(resolved)
    return str(resolved).strip()


def _text_length(value: Any, symbols: dict[int, str]) -> int:
    return len(preview_text(value, symbols))


# ---------------------------------------------------------------- 行高


def _group_widths(info: SheetInfo) -> dict[int, float]:
    widths: dict[int, float] = {}
    for target, column in info.layout.columns.items():
        total = 0.0
        for offset in range(column.span):
            total += info.widths.get(column.column + offset, info.default_width)
        widths[target] = max(total, 0.1)
    return widths


def _calculate_heights(
    rows: list,
    layout: TemplateLayout,
    widths: dict[int, float],
    symbols: dict[int, str],
) -> list[float]:
    count = len(rows)
    if not count:
        return []
    lines = [0] * count
    merge_ends: dict[tuple[int, int], tuple[int, int]] = {}

    for index, row in enumerate(rows):
        best = 0.0
        for target, column in layout.columns.items():
            value = row.values[target]
            span = row.spans[target]
            if span > 1:
                merge_ends[(index + span - 1, target)] = (
                    index,
                    _text_length(value, symbols),
                )
            if not value or span > 1:
                continue
            width = widths.get(target, 0)
            if width <= 0:
                continue
            ratio = _text_length(value, symbols) * CHAR_WIDTH / width
            # VBA：zdz 与「未乘余量的行数」比较，命中后把乘过余量的值存回 zdz
            if best < ratio:
                best = ratio * (
                    MERGE_OVER_RESERVE if ratio >= MERGE_THRESHOLD else MERGE_RESERVE
                )
        if best > MAX_HEIGHT_LINES:
            best = MAX_HEIGHT_LINES
        lines[index] = int(math.ceil(best))

    for index in range(count):
        for target in sorted(layout.columns):
            info = merge_ends.get((index, target))
            if info is None:
                continue
            start, length = info
            span = index - start + 1
            width = widths.get(target, 0)
            if width <= 0:
                continue
            ratio = length * CHAR_WIDTH / width
            if ratio * MERGE_RESERVE <= span:
                continue
            total = sum(lines[start + offset] for offset in range(span))
            if total >= ratio:
                continue
            factor = (
                MERGE_OVER_RESERVE if ratio >= MERGE_THRESHOLD else MERGE_RESERVE
            )
            height = min(math.ceil(ratio) * factor / span, MAX_HEIGHT_LINES)
            height = max(int(math.ceil(height)), 1)
            for offset in range(span):
                lines[start + offset] = height

    return [line * SINGLE_LINE_HEIGHT + HEIGHT_PADDING for line in lines]


# ---------------------------------------------------------------- 写出辅助


def _escape(text: str) -> str:
    text = "".join(
        char for char in text if char in "\t\n\r" or ord(char) >= 0x20
    )
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )


def _cell_xml(row: int, col: int, style: str | None, value: Any) -> str:
    ref = f"{get_column_letter(col)}{row}"
    style_attr = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{style_attr} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            return f'<c r="{ref}"{style_attr}/>'
        number = int(value) if isinstance(value, float) and value.is_integer() else value
        return f'<c r="{ref}"{style_attr}><v>{number}</v></c>'
    text = _escape(str(value))
    return (
        f'<c r="{ref}"{style_attr} t="inlineStr">'
        f'<is><t xml:space="preserve">{text}</t></is></c>'
    )


def _row_xml(
    row_number: int,
    values: list[Any],
    styles: dict[int, str | None],
    max_col: int,
    height: float | None,
) -> str:
    attributes = f' r="{row_number}" spans="1:{max_col}"'
    if height is not None and height > 0:
        attributes += f' ht="{height:g}" customHeight="1"'
    cells = "".join(
        _cell_xml(row_number, col, styles.get(col), values[col - 1])
        for col in range(1, max_col + 1)
    )
    return f"<row{attributes}>{cells}</row>"


def _column_styles(info: SheetInfo) -> dict[int, str | None]:
    positions = sorted(info.styles)
    result: dict[int, str | None] = {}
    for col in range(1, info.max_col + 1):
        result[col] = info.styles.get(col)
    for _target, column in info.layout.columns.items():
        for offset in range(column.span):
            col = column.column + offset
            if result.get(col):
                continue
            if positions:
                nearest = min(positions, key=lambda item: abs(item - col))
                result[col] = info.styles[nearest]
    if positions:
        for col in range(1, info.max_col + 1):
            if result.get(col):
                continue
            nearest = min(positions, key=lambda item: abs(item - col))
            result[col] = info.styles[nearest]
    return result


def _should_band(row, layout: TemplateLayout) -> tuple[int, int] | None:
    if not row.band_merge:
        return None
    if 7 not in layout.columns or 12 not in layout.columns:
        return None
    start = layout.columns[7].column
    end = layout.columns[12].column - 1
    return (start, end) if end > start else None


def _data_merge_refs(rows: list, info: SheetInfo, data_start: int) -> list[str]:
    refs: list[str] = []
    seen: set[str] = set()
    covered: set[tuple[int, int]] = set()
    bands: dict[int, tuple[int, int]] = {}

    def add(min_row: int, min_col: int, max_row: int, max_col: int) -> None:
        if (min_row, min_col) == (max_row, max_col):
            return
        ref = _range_ref(min_row, min_col, max_row, max_col)
        if ref in seen:
            return
        seen.add(ref)
        refs.append(ref)
        for row_index in range(min_row - data_start, max_row - data_start + 1):
            for col in range(min_col, max_col + 1):
                covered.add((row_index, col))

    for index, row in enumerate(rows):
        target_row = data_start + index
        band = _should_band(row, info.layout)
        if band:
            bands[index] = band
        for target in sorted(info.layout.columns):
            column = info.layout.columns[target]
            first_col = column.column
            last_col = first_col + column.span - 1
            if band and band[0] <= first_col <= band[1]:
                continue
            if (index, first_col) in covered:
                continue
            span = row.spans[target]
            if span <= 0:
                # VBA 只对 sckzjhl > 0 的单元格做合并
                continue
            if span > 1:
                add(target_row, first_col, target_row + span - 1, last_col)
            elif last_col > first_col:
                add(target_row, first_col, target_row, last_col)

    for index in sorted(bands):
        start, end = bands[index]
        target_row = data_start + index
        run_start: int | None = None
        for col in range(start, end + 1):
            if (index, col) in covered:
                if run_start is not None:
                    add(target_row, run_start, target_row, col - 1)
                    run_start = None
                continue
            if run_start is None:
                run_start = col
        if run_start is not None:
            add(target_row, run_start, target_row, end)
    return refs


def _shift_row_xml(row_xml: str, delta: int) -> str:
    if delta == 0:
        return row_xml

    def shift_row(match: re.Match) -> str:
        return f'{match.group(1)}{int(match.group(2)) + delta}{match.group(3)}'

    row_xml = re.sub(r'(<row\b[^>]*?\br=")(\d+)(")', shift_row, row_xml, count=1)

    def shift_cell(match: re.Match) -> str:
        return (
            f"{match.group(1)}{match.group(2)}"
            f"{int(match.group(3)) + delta}{match.group(4)}"
        )

    return re.sub(r'(\br=")([A-Z]+)(\d+)(")', shift_cell, row_xml)


def _shift_ref(ref: str, delta: int) -> str | None:
    parsed = _parse_range(ref)
    if parsed is None:
        return None
    min_row, min_col, max_row, max_col = parsed
    return _range_ref(min_row + delta, min_col, max_row + delta, max_col)


def _sheet_body(sheet_xml: str) -> tuple[str, str, str]:
    """返回 (sheetData 起始标签及其前文本, sheetData 内容, sheetData 结束标签及后文本)。"""
    start = sheet_xml.find("<sheetData")
    if start < 0:
        raise ControlPlanGenerateError("模板工作表缺少 sheetData 节点。")
    open_end = sheet_xml.find(">", start) + 1
    tag = sheet_xml[start:open_end]
    if tag.endswith("/>"):
        return sheet_xml[:start] + "<sheetData>", "", sheet_xml[open_end:]
    close = sheet_xml.find("</sheetData>", open_end)
    if close < 0:
        raise ControlPlanGenerateError("模板工作表 sheetData 节点不完整。")
    return (
        sheet_xml[:open_end],
        sheet_xml[open_end:close],
        sheet_xml[close:],
    )


def _split_rows(body: str) -> list[tuple[int, str]]:
    result: list[tuple[int, str]] = []
    for match in _ROW_PATTERN.finditer(body):
        text = match.group(0)
        number = _ROW_NUMBER_PATTERN.search(text)
        if number is None:
            continue
        result.append((int(number.group(1)), text))
    return result


def _extract_merges(sheet_xml: str) -> tuple[str, list[str]]:
    pattern = re.compile(r"<mergeCells\b[^>]*?(?:/>|>.*?</mergeCells>)", re.S)
    match = pattern.search(sheet_xml)
    if match is None:
        return sheet_xml, []
    refs = re.findall(r'<mergeCell\b[^>]*?\bref="([^"]+)"', match.group(0))
    return sheet_xml[: match.start()] + sheet_xml[match.end() :], refs


_MERGE_SUCCESSORS = (
    "<phoneticPr",
    "<conditionalFormatting",
    "<dataValidations",
    "<hyperlinks",
    "<printOptions",
    "<pageMargins",
    "<pageSetup",
    "<headerFooter",
    "<rowBreaks",
    "<colBreaks",
    "<customProperties",
    "<cellWatches",
    "<ignoredErrors",
    "<smartTags",
    "<drawing",
    "<legacyDrawing",
    "<oleObjects",
    "<controls",
    "<extLst",
    "</worksheet>",
)


def _insert_merges(sheet_xml: str, refs: list[str]) -> str:
    if not refs:
        return sheet_xml
    block = (
        f'<mergeCells count="{len(refs)}">'
        + "".join(f'<mergeCell ref="{ref}"/>' for ref in refs)
        + "</mergeCells>"
    )
    index = sheet_xml.find("</sheetData>")
    if index < 0:
        raise ControlPlanGenerateError("模板工作表缺少 sheetData 节点。")
    index += len("</sheetData>")
    # mergeCells 必须位于 sheetData 之后、phoneticPr/conditionalFormatting 等元素之前
    positions = [
        sheet_xml.find(token, index)
        for token in _MERGE_SUCCESSORS
    ]
    positions = [p for p in positions if p >= 0]
    if positions:
        index = min(positions)
    return sheet_xml[:index] + block + sheet_xml[index:]


def _replace_dimension(sheet_xml: str, max_col: int, last_row: int) -> str:
    ref = f"A1:{get_column_letter(max_col)}{last_row}"
    if re.search(r"<dimension\b[^>]*/>", sheet_xml):
        return re.sub(r"<dimension\b[^>]*/>", f'<dimension ref="{ref}"/>', sheet_xml, count=1)
    return sheet_xml


def _remove_row_breaks(sheet_xml: str) -> str:
    return re.sub(
        r"<rowBreaks\b[^>]*?(?:/>|>.*?</rowBreaks>)", "", sheet_xml, flags=re.S
    )


# ---------------------------------------------------------------- 工作簿包写出


def _defined_name_text(value: str) -> tuple[str, str] | None:
    if "!" not in value or "#REF!" in value:
        return None
    prefix, ref = value.split("!", 1)
    if not ref:
        return None
    return prefix, ref


def _sheet_reference(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def _rewrite_print_area(value: str, sheet_name: str, last_row: int) -> str | None:
    parsed = _defined_name_text(value)
    if parsed is None:
        return None
    _prefix, ref = parsed
    if ":" in ref:
        start, end = ref.split(":", 1)
        end = re.sub(r"(\$?[A-Z]+\$?)(\d+)$", lambda m: f"{m.group(1)}{last_row}", end)
        ref = f"{start}:{end}"
    return f"{_sheet_reference(sheet_name)}!{ref}"


def _rewrite_print_titles(value: str, sheet_name: str) -> str | None:
    parsed = _defined_name_text(value)
    if parsed is None:
        return None
    _prefix, ref = parsed
    return f"{_sheet_reference(sheet_name)}!{ref}"


def _rewrite_workbook_xml(
    xml_text: str,
    sheet_id: int,
    rel_id: str,
    sheet_name: str,
    print_area: str | None,
    print_titles: str | None,
    last_row: int,
) -> str:
    sheets_block = (
        f'<sheets><sheet name="{sheet_name}" sheetId="{sheet_id}" '
        f'r:id="{rel_id}"/></sheets>'
    )
    if re.search(r"<sheets>.*?</sheets>", xml_text, re.S):
        xml_text = re.sub(r"<sheets>.*?</sheets>", sheets_block, xml_text, flags=re.S)
    xml_text = re.sub(r"<externalReferences>.*?</externalReferences>", "", xml_text, flags=re.S)
    xml_text = re.sub(r"<externalReferences\s*/>", "", xml_text)
    xml_text = re.sub(r'(<workbookView[^>]*activeTab=")\d+(")', r"\g<1>0\2", xml_text)
    xml_text = re.sub(r'(<calcPr[^>]*calcMode=")[^"]+(")', r"\g<1>auto\2", xml_text)

    names: list[str] = []
    if print_area:
        updated = _rewrite_print_area(print_area, sheet_name, last_row)
        if updated:
            names.append(
                '<definedName name="_xlnm.Print_Area" localSheetId="0">'
                + _escape(updated)
                + "</definedName>"
            )
    if print_titles:
        updated = _rewrite_print_titles(print_titles, sheet_name)
        if updated:
            names.append(
                '<definedName name="_xlnm.Print_Titles" localSheetId="0">'
                + _escape(updated)
                + "</definedName>"
            )
    block = f"<definedNames>{''.join(names)}</definedNames>" if names else ""
    if re.search(r"<definedNames>.*?</definedNames>", xml_text, re.S):
        xml_text = re.sub(r"<definedNames>.*?</definedNames>", block, xml_text, flags=re.S)
    elif block:
        insert_at = xml_text.find("<calcPr")
        if insert_at < 0:
            insert_at = xml_text.find("</workbook>")
        xml_text = xml_text[:insert_at] + block + xml_text[insert_at:]
    return xml_text


def _keep_part(
    parts: dict[str, bytes],
    sheet: SheetInfo,
    keep_rel_id: str,
) -> set[str]:
    keep: set[str] = {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "docProps/core.xml",
        sheet.path,
    }
    if "docProps/custom.xml" in parts:
        keep.add("docProps/custom.xml")
    if _rels_path(sheet.path) in parts:
        keep.add(_rels_path(sheet.path))

    queue = [sheet.path, "xl/workbook.xml"]
    while queue:
        part = queue.pop()
        for rel_id, (rel_type, target) in _read_rels(parts, part).items():
            if target.startswith(("http:", "https:", "file:", "\\\\")):
                continue
            if target not in parts:
                continue
            if rel_type.endswith("/worksheet") and target != sheet.path:
                continue
            if rel_type.endswith("/chartsheet") or rel_type.endswith("/calcChain"):
                continue
            if rel_type.endswith("/externalLink"):
                continue
            if target not in keep:
                keep.add(target)
                queue.append(target)
            rels = _rels_path(target)
            if rels in parts and rels not in keep:
                keep.add(rels)
                queue.append(rels)
    return keep


def _rewrite_workbook_rels(
    parts: dict[str, bytes], keep_rel_id: str, keep: set[str]
) -> bytes:
    text = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    pieces: list[str] = []
    for match in _REL_PATTERN.finditer(text):
        element = match.group(0)
        rel_id = re.search(r'\bId="([^"]+)"', element)
        target = re.search(r'\bTarget="([^"]*)"', element)
        rel_type = re.search(r'\bType="([^"]*)"', element)
        mode = re.search(r'\bTargetMode="([^"]*)"', element)
        if not rel_id or not target or not rel_type:
            continue
        if rel_id.group(1) == keep_rel_id:
            pieces.append(element)
            continue
        if mode and mode.group(1) == "External":
            continue
        if rel_type.group(1).endswith(("/chartsheet", "/calcChain", "/worksheet")):
            continue
        if rel_type.group(1).endswith("/externalLink"):
            continue
        resolved = _resolve_target("xl/workbook.xml", target.group(1))
        if resolved in keep:
            pieces.append(element)
    body = "".join(pieces)
    header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (
        header
        + f'<Relationships xmlns="{PKG_REL_NS}">'
        + body
        + "</Relationships>"
    ).encode("utf-8")


def _rewrite_root_rels(parts: dict[str, bytes], keep: set[str]) -> bytes:
    text = parts["_rels/.rels"].decode("utf-8")
    pieces: list[str] = []
    for match in _REL_PATTERN.finditer(text):
        element = match.group(0)
        target = re.search(r'\bTarget="([^"]*)"', element)
        mode = re.search(r'\bTargetMode="([^"]*)"', element)
        if not target:
            continue
        if mode and mode.group(1) == "External":
            continue
        resolved = _resolve_target("", target.group(1))
        if resolved in keep:
            pieces.append(element)
    header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (
        header
        + f'<Relationships xmlns="{PKG_REL_NS}">'
        + "".join(pieces)
        + "</Relationships>"
    ).encode("utf-8")


_DEFAULT_CONTENT_TYPES = {
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
    "vml": "application/vnd.openxmlformats-officedocument.vmlDrawing",
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "svg": "image/svg+xml",
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
}


def _part_content_type(part: str) -> str | None:
    if part == "xl/workbook.xml":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    if part.startswith("xl/worksheets/sheet") and part.endswith(".xml"):
        return (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        )
    if part == "xl/styles.xml":
        return (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
        )
    if part == "xl/sharedStrings.xml":
        return (
            "application/vnd.openxmlformats-officedocument.spreadsheetml."
            "sharedStrings+xml"
        )
    if part.startswith("xl/theme/"):
        return "application/vnd.openxmlformats-officedocument.theme+xml"
    if part.startswith("xl/drawings/") and part.endswith(".xml"):
        return "application/vnd.openxmlformats-officedocument.drawing+xml"
    if part.startswith("xl/drawings/") and part.endswith(".vml"):
        return "application/vnd.openxmlformats-officedocument.vmlDrawing"
    if part.startswith("xl/charts/") and part.endswith(".xml"):
        return "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    if part.startswith("xl/ctrlProps/"):
        return "application/vnd.ms-excel.controlproperties+xml"
    if part.startswith("xl/activeX/"):
        return "application/vnd.ms-office.activeX+xml"
    if part == "docProps/core.xml":
        return "application/vnd.openxmlformats-package.core-properties+xml"
    if part == "docProps/custom.xml":
        return (
            "application/vnd.openxmlformats-officedocument.custom-properties+xml"
        )
    return None


def _rewrite_content_types(
    parts: dict[str, bytes], keep: set[str]
) -> bytes:
    original = parts["[Content_Types].xml"].decode("utf-8")
    defaults: dict[str, str] = dict(_DEFAULT_CONTENT_TYPES)
    for match in re.finditer(r'<Default\b[^>]*?Extension="([^"]+)"[^>]*?ContentType="([^"]+)"[^>]*/>', original):
        defaults[match.group(1).lower()] = match.group(2)

    overrides: dict[str, str] = {}
    for match in _OVERRIDE_PATTERN.finditer(original):
        element = match.group(0)
        part = re.search(r'\bPartName="/([^"]+)"', element)
        content_type = re.search(r'\bContentType="([^"]+)"', element)
        if part and content_type:
            overrides[part.group(1)] = content_type.group(1)

    for part in keep:
        if part in overrides or part == "[Content_Types].xml":
            continue
        if part.endswith(".rels"):
            continue
        extension = part.rsplit(".", 1)[-1].lower()
        if extension != "xml" and extension in defaults:
            continue
        content_type = _part_content_type(part)
        if content_type:
            overrides[part] = content_type
        elif part.startswith("docProps/") and part.endswith(".xml"):
            overrides[part] = "application/xml"

    pieces = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
        f'<Types xmlns="{CT_NS}">',
    ]
    for extension, content_type in sorted(defaults.items()):
        if extension:
            pieces.append(
                f'<Default Extension="{extension}" ContentType="{content_type}"/>'
            )
    for part, content_type in sorted(overrides.items()):
        if part in keep:
            pieces.append(
                f'<Override PartName="/{part}" ContentType="{content_type}"/>'
            )
    pieces.append("</Types>")
    return "".join(pieces).encode("utf-8")


def _write_package(
    parts: dict[str, bytes],
    keep: set[str],
    output: Path,
    sheet_path: str,
    sheet_xml: bytes,
    workbook_xml: bytes,
    workbook_rels: bytes,
    root_rels: bytes,
    content_types: bytes,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix(output.suffix + ".tmp")
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
            target.writestr("[Content_Types].xml", content_types)
            target.writestr("_rels/.rels", root_rels)
            for name in sorted(keep):
                if name in ("[Content_Types].xml", "_rels/.rels"):
                    continue
                if name == sheet_path:
                    target.writestr(name, sheet_xml)
                elif name == "xl/workbook.xml":
                    target.writestr(name, workbook_xml)
                elif name == "xl/_rels/workbook.xml.rels":
                    target.writestr(name, workbook_rels)
                else:
                    target.writestr(name, parts[name])
        if output.exists():
            output.unlink()
        temp.replace(output)
    except Exception as exc:  # noqa: BLE001
        if temp.exists():
            try:
                temp.unlink()
            except OSError:
                pass
        raise ControlPlanGenerateError(f"保存控制计划失败：{exc}") from exc


# ---------------------------------------------------------------- 生成入口


def generate_control_plan(
    parsed: WorkInstructionParse,
    sheet_name: str,
    output_path: str | Path,
) -> Path:
    if not parsed.rows:
        raise ControlPlanGenerateError("没有可写入的控制计划数据。")
    infos, parts = _load_infos()
    if sheet_name not in infos:
        raise ControlPlanGenerateError(f"模板中不存在工作表“{sheet_name}”。")
    info = infos[sheet_name]
    layout = info.layout
    if not layout.supported:
        raise ControlPlanGenerateError(
            f"模板“{layout.display_name}”暂不支持：{layout.reason}"
        )

    sheet_id = 0
    rel_id = ""
    for name, _path, rid, sid in _workbook_sheets(parts):
        if name == sheet_name:
            sheet_id = sid
            rel_id = rid
            break

    sheet_xml = parts[info.path].decode("utf-8")
    sheet_xml, existing_merges = _extract_merges(sheet_xml)
    prefix, body, suffix = _sheet_body(sheet_xml)
    rows = _split_rows(body)

    marker = layout.marker_row
    header_rows = [text for number, text in rows if number < marker]
    footer_rows = [
        text
        for number, text in rows
        if info.footer_start is not None and number >= info.footer_start
    ]

    data_start = marker
    data_count = len(parsed.rows)
    data_end = data_start + data_count - 1
    # VBA 先删 marker..30，再删「数据后第 2 行 .. 第 428 行」，
    # 因此原第 429 行起的页脚整体上移 (data_end - 427) 行。
    footer_delta = data_end - 427
    blank_rows = max(1, (info.footer_start or 431) - 428)
    blank_start = data_end + 1
    footer_start_new = (
        info.footer_start + footer_delta if info.footer_start is not None else None
    )

    if data_count > 396:
        raise ControlPlanGenerateError(
            f"控制计划数据 {data_count} 行，超过模板可容纳的行数（396 行）。"
        )

    styles = _column_styles(info)
    widths = _group_widths(info)
    heights = _calculate_heights(parsed.rows, layout, widths, layout.symbols)

    data_xml: list[str] = []
    empty_values: list[Any] = [""] * info.max_col
    for index, row in enumerate(parsed.rows):
        values = list(empty_values)
        for target, column in layout.columns.items():
            value = resolve_value(row.values[target], layout.symbols)
            if value != "":
                values[column.column - 1] = value
        data_xml.append(
            _row_xml(
                data_start + index,
                values,
                styles,
                info.max_col,
                heights[index],
            )
        )

    for offset in range(blank_rows):
        data_xml.append(
            _row_xml(
                blank_start + offset,
                empty_values,
                styles,
                info.max_col,
                SINGLE_LINE_HEIGHT + HEIGHT_PADDING,
            )
        )

    shifted_footer = [
        _shift_row_xml(text, footer_delta) for text in footer_rows
    ]

    new_body = "".join(header_rows + data_xml + shifted_footer)
    sheet_xml = prefix + new_body + suffix

    merge_refs: list[str] = []
    seen: set[str] = set()
    for ref in existing_merges:
        parsed_ref = _parse_range(ref)
        if parsed_ref is None or parsed_ref[2] >= marker:
            continue
        if ref not in seen:
            merge_refs.append(ref)
            seen.add(ref)
    for ref in _data_merge_refs(parsed.rows, info, data_start):
        if ref not in seen:
            merge_refs.append(ref)
            seen.add(ref)
    if info.footer_start is not None:
        for ref in info.footer_merges:
            shifted = _shift_ref(ref, footer_delta)
            if shifted and shifted not in seen:
                merge_refs.append(shifted)
                seen.add(shifted)

    sheet_xml = _insert_merges(sheet_xml, merge_refs)
    print_end = 0
    if info.print_area and ":" in info.print_area:
        print_end = _row_index(info.print_area.rsplit(":", 1)[-1])
    if print_end:
        last_data_row = print_end + footer_delta
    elif footer_rows:
        last_data_row = footer_start_new + len(footer_rows) - 1
    else:
        last_data_row = blank_start + blank_rows - 1
    if footer_rows and footer_start_new is not None:
        last_data_row = max(
            last_data_row, footer_start_new + len(footer_rows) - 1
        )
    sheet_xml = _replace_dimension(sheet_xml, info.max_col, last_data_row)
    sheet_xml = _remove_row_breaks(sheet_xml)

    workbook_xml = _rewrite_workbook_xml(
        parts["xl/workbook.xml"].decode("utf-8"),
        sheet_id,
        rel_id,
        "控制计划",
        info.print_area,
        info.print_titles,
        last_data_row,
    )

    keep = _keep_part(parts, info, rel_id)
    workbook_rels = _rewrite_workbook_rels(parts, rel_id, keep)
    root_rels = _rewrite_root_rels(parts, keep)
    content_types = _rewrite_content_types(parts, keep)

    output = Path(output_path)
    _write_package(
        parts,
        keep,
        output,
        info.path,
        sheet_xml.encode("utf-8"),
        workbook_xml.encode("utf-8"),
        workbook_rels,
        root_rels,
        content_types,
    )
    return output
