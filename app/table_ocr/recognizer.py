"""表格识图：从 PDF（文字层或扫描图片）中识别表格并整理为可复制数据。"""

from __future__ import annotations

import math
import re
import threading
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

TARGET_COLUMNS = ("序号", "代号", "版本", "名称", "数量")

_COLUMN_KEYWORDS: dict[str, tuple[str, ...]] = {
    "序号": ("序号", "序 号", "项次", "件号"),
    "版本": ("版本", "版次", "版本号", "version", "rev", "revision"),
    "名称": (
        "名称",
        "品名",
        "零件名称",
        "物料名称",
        "部品名称",
        "description",
        "partname",
    ),
    "数量": ("数量", "用量", "单件数量", "qty", "quantity"),
    "代号": (
        "代号",
        "代码",
        "图号",
        "零件号",
        "料号",
        "编码",
        "物料号",
        "partno",
        "partnumber",
        "drawingno",
        "itemcode",
    ),
}
# 英文短表头易与其他单词混淆，改为整词精确匹配
_EXACT_COLUMNS: dict[str, str] = {
    "no": "序号",
    "rev": "版本",
    "qty": "数量",
    "q'ty": "数量",
}
# 按固定顺序匹配，避免“零件名称”被误判为代号
_MATCH_ORDER = ("序号", "版本", "名称", "数量", "代号")

PDF_SUFFIXES = (".pdf",)
OCR_DPI = 240
OCR_TILE_PIXELS = 2400
OCR_PREVIEW_DPI = 110
OCR_REGION_DPI = 240
VECTOR_MIN_LINE = 40.0
REGION_LIMIT = 3


class TableOcrError(Exception):
    pass


@dataclass
class TableRow:
    index: str = ""
    code: str = ""
    version: str = ""
    name: str = ""
    quantity: str = ""
    order: int = 0

    def value(self, column: str) -> str:
        return {
            "序号": self.index,
            "代号": self.code,
            "版本": self.version,
            "名称": self.name,
            "数量": self.quantity,
        }.get(column, "")


@dataclass
class TextItem:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    score: float = 1.0

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2


@dataclass
class TableRecognition:
    rows: list[TableRow] = field(default_factory=list)
    page_count: int = 0
    text_pages: int = 0
    ocr_pages: int = 0
    warnings: list[str] = field(default_factory=list)


def normalize_text(value) -> str:
    if value is None:
        return ""
    # 全角数字/字母（如“０”）统一成半角，避免序号、代号出现异形字符
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", text).strip()


def _compact(value) -> str:
    return re.sub(r"\s+", "", normalize_text(value)).lower()


def _column_for(value) -> str:
    text = _compact(value)
    if not text:
        return ""
    exact = text.strip(" .:：、,，")
    if exact in _EXACT_COLUMNS:
        return _EXACT_COLUMNS[exact]
    for column in _MATCH_ORDER:
        if any(keyword in text for keyword in _COLUMN_KEYWORDS[column]):
            return column
    return ""


def _index_number(value: str) -> int | None:
    match = re.search(r"\d+", value or "")
    return int(match.group(0)) if match else None


def _header_mapping(cells) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col, cell in enumerate(cells):
        column = _column_for(cell)
        if column and column not in mapping:
            mapping[column] = col
    return mapping


def _is_header(mapping: dict[str, int]) -> bool:
    return "序号" in mapping and len(mapping) >= 3


# ---------- 文字层表格 ----------


def _find_tables(page) -> tuple[list, str]:
    """先按表格线识别，再尝试无边框的文字型表格。"""
    for strategy in ("lines_strict", "text"):
        try:
            finder = page.find_tables(strategy=strategy)
        except Exception:  # noqa: BLE001
            continue
        tables = getattr(finder, "tables", None)
        if tables is None:
            tables = list(finder)
        if tables:
            return list(tables), strategy
    return [], ""


def _words_items(page, bbox=None) -> list[TextItem]:
    """把文字层的单词转换为带坐标的文本块。"""
    items: list[TextItem] = []
    for word in page.get_text("words"):
        text = str(word[4]).strip()
        if not text:
            continue
        x0, y0, x1, y1 = (float(value) for value in word[:4])
        if bbox is not None and not (
            bbox[0] <= x0 <= bbox[2] and bbox[1] <= (y0 + y1) / 2 <= bbox[3]
        ):
            continue
        items.append(TextItem(text, x0, y0, x1, y1))
    return items


def _rows_quality(rows: list[TableRow]) -> int:
    """统计内容较完整的行数，用于判断哪种解析方式更可靠。"""
    return sum(
        1
        for row in rows
        if sum(1 for value in (row.code, row.name, row.quantity) if value) >= 2
    )


def _rows_from_image_page(page, engine) -> list[TableRow]:
    """图片型/图纸型页面：先按矢量线条定位表格，再只识别表格区域。"""
    for rect, hint in _vector_regions(page):
        rows = _rows_from_region(page, rect, engine, hint)
        if rows:
            return rows
    rows = _rows_from_ocr_items(_ocr_items(page, engine))
    if rows:
        return rows
    # 整页 OCR 没找到时，按 90°/270° 再试一次，兼容竖排明细表
    for angle in (270, 90):
        items = _ocr_rotated_page(page, engine, angle)
        rows = _rows_from_ocr_items(items)
        if rows:
            return rows
    return []


def _ocr_rotated_page(page, engine, angle: int) -> list[TextItem]:
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(OCR_PREVIEW_DPI / 72, OCR_PREVIEW_DPI / 72).prerotate(
            angle
        )
    )
    import cv2
    import numpy as np

    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, pixmap.n
    )
    if pixmap.n == 3:
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    return _ocr_image(image, engine)


def _rows_from_text_page(page) -> list[TableRow]:
    rows: list[TableRow] = []
    tables, strategy = _find_tables(page)
    for table in tables:
        if strategy == "lines_strict":
            data = None
            try:
                data = _grid_data(page, table)
            except Exception:  # noqa: BLE001
                data = None
            if not data:
                try:
                    data = table.extract()
                except Exception:  # noqa: BLE001
                    continue
            rows.extend(_rows_from_table_data(data, len(rows)))
        else:
            # 无边框表格的列边界不可靠，改用表头文字位置划列
            rows.extend(
                _rows_from_ocr_items(_words_items(page, table.bbox), len(rows))
            )
    # 图纸类 PDF 的几何线条会干扰表格识别，整页文字定位作为兜底
    page_rows = _rows_from_ocr_items(_words_items(page), len(rows))
    if _rows_quality(rows) < _rows_quality(page_rows):
        return page_rows
    return rows


def _grid_data(page, table) -> list[list[str]] | None:
    """按表格网格重排文字，避免单元格过窄时文字被截断。"""
    row_boxes = [row.bbox for row in getattr(table, "rows", [])]
    if not row_boxes:
        return None
    x_edges = {round(table.bbox[0], 1), round(table.bbox[2], 1)}
    for row in table.rows:
        for cell in row.cells or []:
            if cell:
                x_edges.add(round(cell[0], 1))
    edges = sorted(x_edges)
    if len(edges) < 3:
        return None
    column_count = len(edges) - 1
    matrix = [["" for _ in range(column_count)] for _ in row_boxes]

    page_x0, page_y0, page_x1, page_y1 = table.bbox
    words = [
        word
        for word in page.get_text("words")
        if page_x0 <= word[0] <= page_x1
        and page_y0 <= (word[1] + word[3]) / 2 <= page_y1
        and word[4].strip()
    ]
    words.sort(key=lambda word: (word[1], word[0]))

    for word in words:
        center_y = (word[1] + word[3]) / 2
        row_index = None
        for index, box in enumerate(row_boxes):
            if box[1] <= center_y <= box[3]:
                row_index = index
                break
        if row_index is None:
            continue
        column_index = 0
        for index in range(column_count):
            if word[0] <= edges[index + 1]:
                column_index = index
                break
        else:
            column_index = column_count - 1
        text = word[4].strip()
        current = matrix[row_index][column_index]
        matrix[row_index][column_index] = f"{current} {text}".strip()
    return matrix


def _is_data_row(values: dict[str, str]) -> bool:
    """序号为数字即为数据行；序号缺失时要求至少两列有内容，避免误收页脚。"""
    index_text = values.get("序号", "")
    if index_text and _index_number(index_text) is not None:
        return True
    return sum(
        1 for column in ("代号", "版本", "名称", "数量") if values.get(column)
    ) >= 2


def _build_row(values: dict[str, str], order: int) -> TableRow:
    return TableRow(
        index=values.get("序号", ""),
        code=values.get("代号", ""),
        version=values.get("版本", ""),
        name=values.get("名称", ""),
        quantity=values.get("数量", ""),
        order=order,
    )


def _rows_from_table_data(data, start_order: int) -> list[TableRow]:
    rows: list[TableRow] = []
    mapping: dict[str, int] = {}
    previous: TableRow | None = None
    order = start_order
    for raw in data:
        cells = [normalize_text(cell).replace("\n", " ") for cell in raw]
        found = _header_mapping(cells)
        if _is_header(found):
            mapping = found
            previous = None
            continue
        if not mapping:
            continue
        values = {
            column: normalize_text(cells[index]) if index < len(cells) else ""
            for column, index in mapping.items()
        }
        index_text = values.get("序号", "")
        if not index_text:
            if previous is not None:
                for column, text in values.items():
                    if column != "序号" and text:
                        current = previous.value(column)
                        setattr(previous, _attr(column), f"{current} {text}".strip())
            continue
        if not _is_data_row(values):
            continue
        order += 1
        previous = _build_row(values, order)
        rows.append(previous)
    return rows


def _attr(column: str) -> str:
    return {
        "序号": "index",
        "代号": "code",
        "版本": "version",
        "名称": "name",
        "数量": "quantity",
    }[column]


# ---------- 图片 OCR ----------


_ENGINE = None
_ENGINE_LOCK = threading.Lock()


def _get_ocr_engine():
    global _ENGINE
    with _ENGINE_LOCK:
        if _ENGINE is not None:
            return _ENGINE
        try:
            from rapidocr_onnxruntime import RapidOCR
        except Exception as exc:  # noqa: BLE001
            raise TableOcrError(
                "未安装本地 OCR 引擎 RapidOCR，请先执行 "
                "pip install rapidocr_onnxruntime 后重试。"
            ) from exc
        _ENGINE = RapidOCR()
        return _ENGINE


def warm_up() -> None:
    """预加载 OCR 引擎并空跑一次，减少首次识别的等待。"""
    import numpy as np

    engine = _get_ocr_engine()
    engine(np.full((64, 320, 3), 255, dtype=np.uint8))


def _ocr_items(page, engine) -> list[TextItem]:
    """按块识别大图纸，避免整页缩放后小字识别失败。"""
    scale = OCR_DPI / 72.0
    rect = page.rect
    columns = max(1, math.ceil(rect.width * scale / OCR_TILE_PIXELS))
    rows_count = max(1, math.ceil(rect.height * scale / OCR_TILE_PIXELS))
    step_x = rect.width / columns
    step_y = rect.height / rows_count
    overlap_x = step_x * 0.15
    overlap_y = step_y * 0.15

    items: list[TextItem] = []
    for row_index in range(rows_count):
        for column_index in range(columns):
            clip = pymupdf.Rect(
                max(rect.x0, rect.x0 + column_index * step_x - overlap_x),
                max(rect.y0, rect.y0 + row_index * step_y - overlap_y),
                min(rect.x1, rect.x0 + (column_index + 1) * step_x + overlap_x),
                min(rect.y1, rect.y0 + (row_index + 1) * step_y + overlap_y),
            )
            if clip.is_empty:
                continue
            try:
                pixmap = page.get_pixmap(dpi=OCR_DPI, clip=clip)
                result, _ = engine(pixmap.tobytes("png"))
            except Exception as exc:  # noqa: BLE001
                raise TableOcrError(f"OCR 识别失败：{exc}") from exc
            scale_x = clip.width / max(pixmap.width, 1)
            scale_y = clip.height / max(pixmap.height, 1)
            for entry in result or []:
                if len(entry) < 2:
                    continue
                box, text = entry[0], entry[1]
                score = float(entry[2]) if len(entry) > 2 else 1.0
                text = normalize_text(text)
                if not text or not box:
                    continue
                xs = [clip.x0 + float(point[0]) * scale_x for point in box]
                ys = [clip.y0 + float(point[1]) * scale_y for point in box]
                item = TextItem(
                    text, min(xs), min(ys), max(xs), max(ys), score
                )
                if _absorb_duplicate(items, item):
                    continue
                items.append(item)
    return items


def _absorb_duplicate(items: list[TextItem], item: TextItem) -> bool:
    """合并分块识别在拼接处产生的重复或截断文字。"""
    item_height = max(item.y1 - item.y0, 1.0)
    for existing in items:
        height = max(existing.y1 - existing.y0, item_height)
        same_line = abs(
            (existing.y0 + existing.y1) - (item.y0 + item.y1)
        ) / 2 <= height * 0.7
        if not same_line:
            continue
        overlap = min(existing.x1, item.x1) - max(existing.x0, item.x0)
        shorter = min(existing.x1 - existing.x0, item.x1 - item.x0)
        if shorter <= 0 or overlap < shorter * 0.5:
            continue
        if len(item.text) > len(existing.text):
            existing.text = item.text
            existing.x0 = min(existing.x0, item.x0)
            existing.x1 = max(existing.x1, item.x1)
            existing.y0 = min(existing.y0, item.y0)
            existing.y1 = max(existing.y1, item.y1)
            existing.score = max(existing.score, item.score)
        return True
    return False


def _cluster_lines(items: list[TextItem]) -> list[list[TextItem]]:
    if not items:
        return []
    items = _merge_stacked_labels(items)
    heights = sorted(item.y1 - item.y0 for item in items)
    median_height = heights[len(heights) // 2] or 10.0
    tolerance = max(8.0, median_height * 0.8)
    ordered = sorted(items, key=lambda item: (item.y0 + item.y1) / 2)
    clusters: list[list[TextItem]] = []
    for item in ordered:
        center = (item.y0 + item.y1) / 2
        if clusters:
            last_center = sum(
                (i.y0 + i.y1) / 2 for i in clusters[-1]
            ) / len(clusters[-1])
            if abs(center - last_center) <= tolerance:
                clusters[-1].append(item)
                continue
        clusters.append([item])
    return [sorted(cluster, key=lambda item: item.x0) for cluster in clusters]


def _merge_stacked_labels(items: list[TextItem]) -> list[TextItem]:
    """合并竖排表头（如“序/号”上下两行），便于按列名匹配。"""
    items = list(items)
    single_han = re.compile(r"[\u4e00-\u9fff]")
    index = 0
    while index < len(items):
        current = items[index]
        if not single_han.fullmatch(current.text):
            index += 1
            continue
        height = max(current.y1 - current.y0, 1.0)
        merged = False
        for other_index in range(index + 1, len(items)):
            other = items[other_index]
            if not single_han.fullmatch(other.text):
                continue
            overlap = min(current.x1, other.x1) - max(current.x0, other.x0)
            narrower = min(current.x1 - current.x0, other.x1 - other.x0)
            if narrower <= 0 or overlap < narrower * 0.8:
                continue
            gap = max(other.y0 - current.y1, current.y0 - other.y1)
            if gap > height * 0.8:
                continue
            upper, lower = (
                (current, other) if current.y0 <= other.y0 else (other, current)
            )
            upper.text = f"{upper.text}{lower.text}"
            upper.x0 = min(upper.x0, lower.x0)
            upper.x1 = max(upper.x1, lower.x1)
            upper.y0 = min(upper.y0, lower.y0)
            upper.y1 = max(upper.y1, lower.y1)
            items.pop(other_index)
            merged = True
            break
        if not merged:
            index += 1
    return items


def _merge_line_cells(cluster: list[TextItem]) -> list[TextItem]:
    """把同一单元格内被拆散的文字（如竖排“数 量”）合并成一个文本块。"""
    if not cluster:
        return []
    height = sorted(item.y1 - item.y0 for item in cluster)[len(cluster) // 2]
    gap_limit = max(3.0, height * 0.6)
    merged: list[TextItem] = []
    current: TextItem | None = None
    for item in cluster:
        if current is not None and item.x0 - current.x1 <= gap_limit:
            current.text = f"{current.text}{item.text}"
            current.x0 = min(current.x0, item.x0)
            current.y0 = min(current.y0, item.y0)
            current.x1 = max(current.x1, item.x1)
            current.y1 = max(current.y1, item.y1)
            current.score = min(current.score, item.score)
            continue
        if current is not None:
            merged.append(current)
        current = TextItem(item.text, item.x0, item.y0, item.x1, item.y1, item.score)
    if current is not None:
        merged.append(current)
    return merged


def _header_columns(cluster: list[TextItem]) -> dict[str, TextItem]:
    mapping: dict[str, TextItem] = {}
    for item in _merge_line_cells(cluster):
        column = _column_for(item.text)
        if column and column not in mapping:
            mapping[column] = item
    return mapping


def _zones_from_header(
    mapping: dict[str, TextItem]
) -> list[tuple[float, float, str]]:
    ordered = sorted(mapping.items(), key=lambda pair: pair[1].center_x)
    centers = [item.center_x for _, item in ordered]
    boundaries = [
        (centers[index] + centers[index + 1]) / 2
        for index in range(len(centers) - 1)
    ]
    zones: list[tuple[float, float, str]] = []
    for index, (column, item) in enumerate(ordered):
        left = -1e9 if index == 0 else boundaries[index - 1]
        right = 1e9 if index == len(ordered) - 1 else boundaries[index]
        zones.append((left, right, column))
    return zones


def _zone_column(zones, value: float) -> str:
    for left, right, column in zones:
        if left <= value <= right:
            return column
    return ""


def _rows_from_ocr_items(
    items: list[TextItem], start_order: int = 0
) -> list[TableRow]:
    rows: list[TableRow] = []
    lines = _cluster_lines(items)
    heights = sorted(item.y1 - item.y0 for item in items)
    median_height = heights[len(heights) // 2] if heights else 10.0
    header_indexes = [
        index
        for index, line in enumerate(lines)
        if _is_header(_header_columns(line))
    ]
    order = start_order
    for position, header_index in enumerate(header_indexes):
        mapping = _header_columns(lines[header_index])
        zones = _zones_from_header(mapping)
        header_y = max(item.y1 for item in mapping.values())
        end = (
            header_indexes[position + 1]
            if position + 1 < len(header_indexes)
            else len(lines)
        )
        previous: TableRow | None = None
        previous_top = 0.0
        for line in lines[header_index + 1 : end]:
            line_top = min(item.y0 for item in line)
            if line_top < header_y:
                continue
            values: dict[str, str] = {}
            for item in line:
                column = _zone_column(zones, item.x0)
                if not column:
                    continue
                values[column] = (
                    f"{values[column]} {item.text}".strip()
                    if column in values
                    else item.text
                )
            index_text = values.get("序号", "")
            if not index_text:
                # 仅把紧贴上一行的换行内容并入上一行，且只补到代号/名称列
                tight = (
                    previous is not None
                    and line_top - previous_top <= median_height * 1.2
                    and len(values) <= 2
                )
                if tight:
                    for column, text in values.items():
                        if column in ("代号", "名称") and text:
                            current = previous.value(column)
                            setattr(
                                previous,
                                _attr(column),
                                f"{current} {text}".strip(),
                            )
                    continue
            if not _is_data_row(values):
                continue
            order += 1
            previous = _build_row(values, order)
            previous_top = line_top
            rows.append(previous)
    return rows


def _grid_rows(items: list[TextItem], image) -> list[TableRow] | None:
    """按图片中的网格线切分单元格，比按表头位置划列更准确。"""
    try:
        import cv2
        import numpy as np
    except Exception:  # noqa: BLE001
        return None
    if image is None:
        return None
    gray = image
    if len(gray.shape) == 3:
        gray = cv2.cvtColor(gray, cv2.COLOR_BGR2GRAY)
    height, width = gray.shape
    binary = cv2.adaptiveThreshold(
        ~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -3
    )
    vkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(12, height // 2)))
    hkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(12, width // 2), 1))
    vlines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vkernel)
    hlines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, hkernel)
    xs = _boundaries(vlines.sum(axis=0))
    ys = _boundaries(hlines.sum(axis=1))
    if len(xs) < 3 or len(ys) < 3:
        return None
    cells = [["" for _ in range(len(xs) - 1)] for _ in range(len(ys) - 1)]
    for item in items:
        column = _band_index(xs, item.center_x)
        row = _band_index(ys, (item.y0 + item.y1) / 2)
        if column is None or row is None:
            continue
        current = cells[row][column]
        cells[row][column] = f"{current} {item.text}".strip()

    header_row = None
    mapping: dict[str, int] = {}
    top = list(range(min(4, len(cells))))
    bottom = list(range(max(0, len(cells) - 4), len(cells)))
    for row_index in dict.fromkeys(top + bottom):
        found = {}
        for column, text in enumerate(cells[row_index]):
            name = _column_for(text)
            if name and name not in found:
                found[name] = column
        if _is_header(found):
            header_row = row_index
            mapping = found
            break
    if header_row is None:
        return None

    rows: list[TableRow] = []
    order = 0
    for row_index, row in enumerate(cells):
        if row_index == header_row:
            continue
        values = {
            name: row[column] if column < len(row) else ""
            for name, column in mapping.items()
        }
        if not _is_data_row(values):
            continue
        order += 1
        rows.append(_build_row(values, order))
    return rows


def _boundaries(profile) -> list[float]:
    """把投影中连续的线条像素聚成一条边界线。"""
    bands: list[float] = []
    start = None
    for index, value in enumerate(profile):
        if value > 0:
            if start is None:
                start = index
        elif start is not None:
            bands.append((start + index - 1) / 2)
            start = None
    if start is not None:
        bands.append((start + len(profile) - 1) / 2)
    return bands


def _band_index(boundaries: list[float], value: float) -> int | None:
    for index in range(len(boundaries) - 1):
        if boundaries[index] <= value <= boundaries[index + 1]:
            return index
    return None


def _overlap_ratio(a0, a1, b0, b1) -> float:
    inter = min(a1, b1) - max(a0, b0)
    if inter <= 0:
        return 0.0
    shorter = min(a1 - a0, b1 - b0)
    return inter / shorter if shorter > 0 else 0.0


def _vector_segments(page):
    """把矢量图形拆成水平/垂直线段：(垂直坐标, 起点, 终点)。"""
    horizontal: list[tuple[float, float, float]] = []
    vertical: list[tuple[float, float, float]] = []
    try:
        drawings = page.get_drawings()
    except Exception:  # noqa: BLE001
        return horizontal, vertical
    for drawing in drawings:
        for item in drawing["items"]:
            kind = item[0]
            if kind == "l":
                p1, p2 = item[1], item[2]
                dx, dy = abs(p1.x - p2.x), abs(p1.y - p2.y)
                if dy <= 1.5 and dx >= VECTOR_MIN_LINE:
                    horizontal.append(
                        ((p1.y + p2.y) / 2, min(p1.x, p2.x), max(p1.x, p2.x))
                    )
                elif dx <= 1.5 and dy >= VECTOR_MIN_LINE:
                    vertical.append(
                        ((p1.x + p2.x) / 2, min(p1.y, p2.y), max(p1.y, p2.y))
                    )
            elif kind == "re":
                rect = item[1]
                if rect.width >= VECTOR_MIN_LINE:
                    horizontal.append((rect.y0, rect.x0, rect.x1))
                    horizontal.append((rect.y1, rect.x0, rect.x1))
                if rect.height >= VECTOR_MIN_LINE:
                    vertical.append((rect.x0, rect.y0, rect.y1))
                    vertical.append((rect.x1, rect.y0, rect.y1))
    return horizontal, vertical


def _collinear_levels(segments, tolerance: float = 2.0) -> list[dict]:
    """把同一直线上的重复线段（重绘、虚线）合并成一条。"""
    levels: list[dict] = []
    for center, start, end in sorted(segments, key=lambda seg: seg[0]):
        for level in levels:
            if abs(level["center"] - center) <= tolerance and _overlap_ratio(
                level["start"], level["end"], start, end
            ) >= 0.3:
                level["count"] += 1
                level["start"] = min(level["start"], start)
                level["end"] = max(level["end"], end)
                break
        else:
            levels.append(
                {"center": center, "start": start, "end": end, "count": 1}
            )
    return levels


def _span_groups(levels, min_gap: float = 2.0, max_gap: float = 120.0) -> list[list[dict]]:
    """把长度对齐、间距接近的平行线归为一组（表格的行线/列线）。"""
    buckets: dict[tuple[int, int], list[dict]] = {}
    for level in levels:
        key = (round(level["start"] / 25), round(level["end"] / 25))
        buckets.setdefault(key, []).append(level)
    groups: list[list[dict]] = []
    for bucket in buckets.values():
        current: list[dict] = []
        for level in sorted(bucket, key=lambda lv: lv["center"]):
            if current and min_gap <= level["center"] - current[-1]["center"] <= max_gap:
                current.append(level)
            else:
                if len(current) >= 5:
                    groups.append(current)
                current = [level]
        if len(current) >= 5:
            groups.append(current)
    return groups


def _spacing_uniform(group: list[dict]) -> bool:
    gaps = [
        group[index + 1]["center"] - group[index]["center"]
        for index in range(len(group) - 1)
    ]
    if not gaps:
        return False
    mean = sum(gaps) / len(gaps)
    if mean <= 0:
        return False
    deviation = (sum((gap - mean) ** 2 for gap in gaps) / len(gaps)) ** 0.5
    return deviation / mean <= 0.4


def _strip_regions(levels, cross_levels, vertical_parallel: bool) -> list[pymupdf.Rect]:
    """等距平行线组视为表格，另一方向的贯穿线用于确认。"""
    regions: list[pymupdf.Rect] = []
    for group in _span_groups(levels):
        if not _spacing_uniform(group):
            continue
        cross0, cross1 = group[0]["center"], group[-1]["center"]
        along0 = min(level["start"] for level in group)
        along1 = max(level["end"] for level in group)
        if vertical_parallel:
            rect = pymupdf.Rect(cross0, along0, cross1, along1)
        else:
            rect = pymupdf.Rect(along0, cross0, along1, cross1)
        if rect.width < 40 or rect.height < 40:
            continue
        supporting = 0
        for level in cross_levels:
            if vertical_parallel:
                if not rect.y0 - 6 <= level["center"] <= rect.y1 + 6:
                    continue
                if _overlap_ratio(level["start"], level["end"], rect.x0, rect.x1) < 0.5:
                    continue
            else:
                if not rect.x0 - 6 <= level["center"] <= rect.x1 + 6:
                    continue
                if _overlap_ratio(level["start"], level["end"], rect.y0, rect.y1) < 0.5:
                    continue
            supporting += 1
        if supporting < 2:
            continue
        regions.append(pymupdf.Rect(rect.x0 - 4, rect.y0 - 4, rect.x1 + 4, rect.y1 + 4))
    return regions


def _vector_regions(page) -> list[tuple[pymupdf.Rect, int]]:
    """用矢量线条定位表格区域，并给出文字方向提示，避免整页 OCR。"""
    horizontal, vertical = _vector_segments(page)
    if len(horizontal) < 4 or len(vertical) < 4:
        return []
    horizontal_levels = _collinear_levels(horizontal)
    vertical_levels = _collinear_levels(vertical)
    regions = [
        (rect, 0) for rect in _strip_regions(horizontal_levels, vertical_levels, False)
    ]
    # 竖排图纸：行线在纵向，文字随之旋转 90°
    regions.extend(
        (rect, 90)
        for rect in _strip_regions(vertical_levels, horizontal_levels, True)
    )
    regions.sort(key=lambda item: -(item[0].width * item[0].height))
    picked: list[tuple[pymupdf.Rect, int]] = []
    for rect, hint in regions:
        if any(
            _overlap_ratio(rect.x0, rect.x1, other.x0, other.x1) >= 0.8
            and _overlap_ratio(rect.y0, rect.y1, other.y0, other.y1) >= 0.8
            for other, _ in picked
        ):
            continue
        picked.append((rect, hint))
        if len(picked) >= REGION_LIMIT:
            break
    return picked


def _render_crop(page, rect: pymupdf.Rect, dpi: float, angle: int):
    """渲染指定区域，可按 90° 倍数旋转，返回 (图片, 缩放比例)。"""
    import cv2
    import numpy as np

    pixmap = page.get_pixmap(dpi=dpi, clip=rect)
    image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height, pixmap.width, pixmap.n
    )
    if pixmap.n == 3:
        image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    if angle == 90:
        image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    elif angle == 270:
        image = cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)
    elif angle == 180:
        image = cv2.rotate(image, cv2.ROTATE_180)
    return image


def _ocr_image(image, engine) -> list[TextItem]:
    import cv2

    ok, buffer = cv2.imencode(".png", image)
    if not ok:
        return []
    result, _ = engine(buffer.tobytes())
    items: list[TextItem] = []
    for entry in result or []:
        if len(entry) < 2:
            continue
        box, text = entry[0], entry[1]
        score = float(entry[2]) if len(entry) > 2 else 1.0
        text = normalize_text(text)
        if not text or not box:
            continue
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        items.append(TextItem(text, min(xs), min(ys), max(xs), max(ys), score))
    return items


def _rows_from_region(page, rect, engine, angle_hint: int = 0) -> list[TableRow]:
    """按矢量线条给出的方向直接高清识别表格区域，失败时才逐向试探。"""

    def recognize(angle: int, dpi: float) -> list[TableRow]:
        image = _render_crop(page, rect, dpi, angle)
        items = _ocr_image(image, engine)
        if not items:
            return []
        rows = _grid_rows(items, image)
        if rows is None:
            rows = _rows_from_ocr_items(items)
        if rows:
            return rows
        # 表头位于区域下方时，翻转纵坐标后再按表头划列
        height = float(image.shape[0])
        flipped = [
            TextItem(item.text, item.x0, height - item.y1, item.x1, height - item.y0,
                     item.score)
            for item in items
        ]
        return _rows_from_ocr_items(flipped)

    rows = recognize(angle_hint, OCR_REGION_DPI)
    if rows:
        return rows
    for candidate in (0, 90, 270, 180):
        if candidate == angle_hint:
            continue
        preview = _render_crop(page, rect, OCR_PREVIEW_DPI, candidate)
        preview_items = _ocr_image(preview, engine)
        if not any(
            _is_header(_header_columns(line))
            for line in _cluster_lines(preview_items)
        ):
            continue
        rows = recognize(candidate, OCR_REGION_DPI)
        if rows:
            return rows
    return []


def _looks_like_data(row: TableRow) -> bool:
    """判断缺少序号的行是否像明细数据，用于排除图纸标题栏文字。"""
    if re.search(r"[A-Za-z0-9]", row.code or ""):
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?", (row.quantity or "").strip()):
        return True
    return False


def _drop_label_rows(rows: list[TableRow]) -> list[TableRow]:
    """丢弃没有序号且不像明细数据的行（图纸标题栏、页脚等）。"""
    return [
        row
        for row in rows
        if _index_number(row.index) is not None or _looks_like_data(row)
    ]


def _fill_missing_indexes(rows: list[TableRow]) -> None:
    """OCR 漏掉个别序号时，按上下相邻序号补齐。"""
    for index, row in enumerate(rows):
        if _index_number(row.index) is not None:
            continue
        previous = _index_number(rows[index - 1].index) if index > 0 else None
        following = (
            _index_number(rows[index + 1].index)
            if index + 1 < len(rows)
            else None
        )
        if (
            previous is not None
            and following is not None
            and following == previous + 1
        ):
            row.index = str(previous + 1)


# ---------- 对外入口 ----------


def recognize_pdf(file_path: str | Path, engine=None) -> TableRecognition:
    path = Path(file_path)
    if not path.is_file():
        raise TableOcrError(f"文件不存在：{path}")
    if path.suffix.lower() not in PDF_SUFFIXES:
        raise TableOcrError("仅支持 .pdf 文件。")

    try:
        document = pymupdf.open(path)
    except Exception as exc:  # noqa: BLE001
        raise TableOcrError(f"无法打开 PDF：{exc}") from exc

    result = TableRecognition()
    try:
        result.page_count = document.page_count
        for page_number, page in enumerate(document, 1):
            rows = _rows_from_text_page(page)
            if rows:
                result.text_pages += 1
                result.rows.extend(rows)
                continue
            if engine is None:
                engine = _get_ocr_engine()
            try:
                rows = _rows_from_image_page(page, engine)
            except TableOcrError:
                raise
            except Exception as exc:  # noqa: BLE001
                result.warnings.append(f"第 {page_number} 页识别失败：{exc}")
                continue
            if rows:
                result.ocr_pages += 1
                result.rows.extend(rows)
            else:
                result.warnings.append(
                    f"第 {page_number} 页未识别到含序号/代号/版本/名称/数量的明细表"
                )
    finally:
        document.close()

    if not result.rows:
        raise TableOcrError(
            "未在 PDF 中识别到明细表。请确认图纸或表格中包含序号、代号、"
            "版本、名称、数量这类列名；若为扫描图纸，建议使用 300dpi 以上的清晰文件。"
        )
    _fill_missing_indexes(result.rows)
    result.rows = _drop_label_rows(result.rows)
    if not result.rows:
        raise TableOcrError(
            "未在 PDF 中识别到明细表。请确认图纸或表格中包含序号、代号、"
            "版本、名称、数量这类列名；若为扫描图纸，建议使用 300dpi 以上的清晰文件。"
        )
    result.rows.sort(
        key=lambda row: (
            (0, _index_number(row.index), row.order)
            if _index_number(row.index) is not None
            else (1, 0, row.order)
        )
    )
    abnormal = [row for row in result.rows if _index_number(row.index) is None]
    if abnormal:
        result.warnings.append(
            f"有 {len(abnormal)} 行序号未识别为数字（{abnormal[0].index} 等），"
            "已按识别顺序排在最后，请人工核对。"
        )
    return result
