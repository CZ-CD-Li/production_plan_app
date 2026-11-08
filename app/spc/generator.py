"""SPC 数据生成与 SPC sheet 写出。"""

from __future__ import annotations

import math
import os
import posixpath
import random
import re
import statistics
import sys
import time
import zipfile
from copy import copy, deepcopy
from dataclasses import dataclass, field
from datetime import date
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape, unescape

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage

from app.config import DATA_DIR, ROOT_DIR
from app.models import User
from app.spc.spec_parser import SpcItem, SpcPlanImport
from app.utils.debug_log import info, warn

GROUPS = 25
ROW_START = 9
COL_START = 3  # C 列
MAX_PER_GROUP = 5
TEMPLATE_CHANGAN = "changan"
TEMPLATE_WULING = "wuling"
TEMPLATE_FILES = {
    TEMPLATE_CHANGAN: "SPCsheet示例.xlsx",
    TEMPLATE_WULING: "五菱SPC模板.xlsx",
}
TEMPLATE_LABELS = {TEMPLATE_CHANGAN: "长安", TEMPLATE_WULING: "五菱"}
TEMPLATE_NAME = TEMPLATE_FILES[TEMPLATE_CHANGAN]
XMR_TEMPLATE_NAME = "SPC单值极差图模板.xlsx"
WULING_TEMPLATE_NAME = TEMPLATE_FILES[TEMPLATE_WULING]
WULING_PER_GROUP = 5
WULING_SAMPLE_COUNT = GROUPS * WULING_PER_GROUP
# 五菱模板固定 25 组 × 5 件，按 10 + 10 + 5 组横向排列
WULING_BLOCKS = ((10, 3, 10), (18, 3, 10), (26, 3, 5))

CHART_XBAR = "xbar"  # 均值极差图（Xbar-R）
CHART_XMR = "xmr"  # 单值极差图（X-MR）
XMR_POINTS = GROUPS  # 单值极差图固定 25 个单值（模板 C..AA 列）
XMR_D2 = 1.128
XMR_D4 = 3.267

D2 = {1: XMR_D2, 2: 1.128, 3: 1.693, 4: 2.059, 5: 2.326}
D3 = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.0}
D4 = {1: XMR_D4, 2: 3.267, 3: 2.574, 4: 2.282, 5: 2.114}

CPK_MIN, CPK_MAX = 1.67, 1.92
PPK_MIN, PPK_MAX = 1.65, 1.92
CPK_TARGET = 1.80
MAX_ATTEMPTS = 400


class SpcGenerateError(Exception):
    pass


def cpk_grade(cpk: float) -> str:
    """按 CPK 值给出分析结果等级。"""
    if cpk < 1:
        return "无法接受"
    if cpk < 1.33:
        return "合格"
    if cpk <= 1.67:
        return "良好"
    return "优秀"


@dataclass
class ItemResult:
    index: str
    description: str
    sheet_name: str
    gauge_name: str
    lower: float
    target: float
    upper: float
    unit: str
    per_group: int
    cpk: float
    ppk: float
    mean: float
    sigma_within: float
    sigma_overall: float
    attempts: int
    row: int = 0
    values: list[list[float]] = field(default_factory=list)
    chart_type: str = CHART_XBAR
    spec_kind: str = ""
    mr_bar: float = 0.0

    @property
    def cpk_text(self) -> str:
        return f"CPK={self.cpk:.2f}"

    @property
    def grade(self) -> str:
        return cpk_grade(self.cpk)

    @property
    def analysis_text(self) -> str:
        return f"CPK={self.cpk:.2f}，PPK={self.ppk:.2f}"

    @property
    def chart_label(self) -> str:
        return "单值极差图" if self.chart_type == CHART_XMR else "均值极差图"

    @property
    def total_points(self) -> int:
        return GROUPS * self.per_group

    @property
    def rbar(self) -> float:
        if self.chart_type == CHART_XMR:
            return self.mr_bar
        return self.sigma_within * D2[self.per_group]

    @property
    def max_range(self) -> float:
        if self.chart_type == CHART_XMR:
            flat = [group[0] for group in self.values]
            diffs = [
                abs(current - previous)
                for previous, current in zip(flat, flat[1:])
            ]
            return max(diffs) if diffs else 0.0
        ranges = [
            max(group) - min(group) for group in self.values if len(group) > 1
        ]
        return max(ranges) if ranges else 0.0

    @property
    def r_upper(self) -> float:
        """R / MR 管制图最大值：控制上限与实际最大极差取大。"""
        factor = XMR_D4 if self.chart_type == CHART_XMR else D4[self.per_group]
        bar = self.mr_bar if self.chart_type == CHART_XMR else self.rbar
        return max(factor * bar, self.max_range)


@dataclass
class SpcGenerationResult:
    output_path: str
    sample_count: int
    per_group: int
    results: list[ItemResult] = field(default_factory=list)
    template_key: str = TEMPLATE_CHANGAN

    @property
    def template_label(self) -> str:
        return TEMPLATE_LABELS.get(self.template_key, self.template_key)


def template_path(name: str = TEMPLATE_NAME) -> Path:
    candidates = []
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass) / name)
    candidates.append(ROOT_DIR / name)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise SpcGenerateError(
        f"未找到 SPC 模板文件 {name}，请将模板放在程序目录下。"
    )


def validate_sample_count(sample_count: int) -> int:
    if sample_count % GROUPS != 0:
        raise SpcGenerateError(f"数据数量必须是 {GROUPS} 的倍数。")
    per_group = sample_count // GROUPS
    if per_group < 2:
        raise SpcGenerateError("每组至少需要 2 个样本，数据数量不得少于 50。")
    if per_group > MAX_PER_GROUP:
        raise SpcGenerateError(
            f"SPC 模板最多 5 行样本，每组最多 {MAX_PER_GROUP} 个，"
            f"数据数量不得超过 {GROUPS * MAX_PER_GROUP}。"
        )
    return per_group


def _step_decimals(step: float) -> int:
    text = f"{step:.10f}".rstrip("0")
    if "." not in text:
        return 0
    return len(text.split(".")[1])


def _ceil_to_step(value: float, step: float, decimals: int) -> float:
    """按量具精度向上取整（等价 Excel CEILING）。"""
    units = math.ceil(round(value / step, 9))
    return round(units * step, decimals)


def _has_seven_run(values: list[float]) -> bool:
    if len(values) < 7:
        return False
    streak = 1
    direction = 0
    for previous, current in zip(values, values[1:]):
        if current > previous:
            step = 1
        elif current < previous:
            step = -1
        else:
            step = 0
        if step and step == direction:
            streak += 1
        else:
            direction = step
            streak = 1 if step else 0
        if streak >= 7:
            return True
    return False


def _metrics(
    values: list[list[float]],
    lower: float,
    upper: float,
    per_group: int,
) -> dict:
    flat = [value for group in values for value in group]
    mean = statistics.fmean(flat)
    group_means = [statistics.fmean(group) for group in values]
    ranges = [max(group) - min(group) for group in values]
    rbar = statistics.fmean(ranges) if ranges else 0.0
    sigma_within = rbar / D2[per_group] if rbar else 0.0
    sigma_overall = statistics.stdev(flat) if len(flat) > 1 else 0.0
    cpk = (
        min((upper - mean) / (3 * sigma_within), (mean - lower) / (3 * sigma_within))
        if sigma_within > 0
        else 0.0
    )
    ppk = (
        min((upper - mean) / (3 * sigma_overall), (mean - lower) / (3 * sigma_overall))
        if sigma_overall > 0
        else 0.0
    )
    return {
        "flat": flat,
        "mean": mean,
        "group_means": group_means,
        "ranges": ranges,
        "rbar": rbar,
        "sigma_within": sigma_within,
        "sigma_overall": sigma_overall,
        "cpk": cpk,
        "ppk": ppk,
    }


def _in_control(metrics: dict, per_group: int) -> bool:
    mean = metrics["mean"]
    rbar = metrics["rbar"]
    if rbar <= 0:
        return False
    lower_x = mean - rbar
    upper_x = mean + rbar
    if any(value < lower_x - 1e-9 or value > upper_x + 1e-9 for value in metrics["group_means"]):
        return False
    upper_r = D4[per_group] * rbar
    lower_r = D3[per_group] * rbar
    if any(value < lower_r - 1e-9 or value > upper_r + 1e-9 for value in metrics["ranges"]):
        return False
    return True


def _clamp_group_means(
    values: list[list[float]],
    step: float,
    decimals: int,
    lower: float,
    upper: float,
) -> None:
    """把偏离中心过多的组整体平移到控制限内（平移不改变组内极差）。"""
    for _ in range(3):
        flat = [value for group in values for value in group]
        mean = statistics.fmean(flat)
        ranges = [max(group) - min(group) for group in values]
        rbar = statistics.fmean(ranges) if ranges else 0.0
        if rbar <= 0:
            return
        limit = 0.9 * rbar
        changed = False
        for index, group in enumerate(values):
            difference = statistics.fmean(group) - mean
            if abs(difference) <= limit:
                continue
            excess = abs(difference) - limit
            steps = math.ceil(round(excess / step, 9))
            delta = -math.copysign(steps * step, difference)
            shifted = [round(value + delta, decimals) for value in group]
            if all(lower - 1e-9 <= value <= upper + 1e-9 for value in shifted):
                values[index] = shifted
                changed = True
        if not changed:
            return


def _mr_ok(flat: list[float]) -> bool:
    """IX-MR 图：移动极差不得超出 D4×MRbar。"""
    diffs = [abs(current - previous) for previous, current in zip(flat, flat[1:])]
    if not diffs:
        return False
    mrbar = statistics.fmean(diffs)
    if mrbar <= 0:
        return False
    return all(diff <= D4[2] * mrbar + 1e-9 for diff in diffs)


def _generate_values(
    item: SpcItem, per_group: int
) -> tuple[list[list[float]], dict, int]:
    lower = float(item.lower)
    target = float(item.target)
    upper = float(item.upper)
    step = float(item.precision)
    decimals = _step_decimals(step)
    sigma_spec = (upper - lower) / 6 / 1.8
    scale = 1.0
    last: dict | None = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        sigma_within = sigma_spec * scale
        sigma_between = sigma_spec * scale * 0.15
        values: list[list[float]] = []
        for _ in range(GROUPS):
            group_mean = random.gauss(target, sigma_between)
            group = [
                _ceil_to_step(random.gauss(group_mean, sigma_within), step, decimals)
                for _ in range(per_group)
            ]
            values.append(group)

        flat = [value for group in values for value in group]
        if not all(lower - 1e-9 <= value <= upper + 1e-9 for value in flat):
            continue

        _clamp_group_means(values, step, decimals, lower, upper)
        metrics = _metrics(values, lower, upper, per_group)
        last = metrics
        flat = metrics["flat"]
        in_spec = all(lower - 1e-9 <= value <= upper + 1e-9 for value in flat)
        varied = len({round(value, 9) for value in flat}) > 1
        cpk = metrics["cpk"]
        ppk = metrics["ppk"]
        ok = (
            in_spec
            and varied
            and CPK_MIN <= cpk <= CPK_MAX
            and PPK_MIN <= ppk <= PPK_MAX
            and _in_control(metrics, per_group)
            and _mr_ok(flat)
            and all(
                abs(value - metrics["mean"]) <= 3 * metrics["sigma_overall"] + 1e-9
                for value in flat
            )
            and not _has_seven_run(metrics["group_means"])
            and not _has_seven_run(flat)
        )
        if ok:
            info(
                "[SPC-生成]",
                f"序号={item.index} attempts={attempt} cpk={cpk:.3f} ppk={ppk:.3f}",
            )
            return values, metrics, attempt

        if cpk > CPK_MAX:
            scale = min(1.8, scale * cpk / CPK_TARGET)
        elif cpk < CPK_MIN:
            scale = max(0.45, scale * max(cpk, 0.6) / CPK_TARGET)
        else:
            scale = min(1.8, max(0.45, scale * random.uniform(0.98, 1.02)))

    detail = ""
    if last is not None:
        detail = (
            f"（最后一次 CPK={last['cpk']:.2f}、PPK={last['ppk']:.2f}）"
        )
    raise SpcGenerateError(
        f"序号 {item.index} 在 {MAX_ATTEMPTS} 次内未生成合格数据{detail}，"
        "请检查上下限与量具精度是否合理。"
    )


def _xmr_metrics(
    flat: list[float], lower: float, upper: float, kind: str = ""
) -> dict:
    """单值极差图指标：σ 用移动极差 / d2 估计，MRbar 即平均移动极差。

    单边公差只按真实存在的界限计算 CPK/PPK（只有下限时为 CPL，只有上限时为 CPU）。
    """
    moving = [
        abs(current - previous) for previous, current in zip(flat, flat[1:])
    ]
    mean = statistics.fmean(flat)
    mrbar = statistics.fmean(moving) if moving else 0.0
    sigma_within = mrbar / XMR_D2 if mrbar else 0.0
    sigma_overall = statistics.stdev(flat) if len(flat) > 1 else 0.0
    if kind == "lower":
        cpk = (mean - lower) / (3 * sigma_within) if sigma_within > 0 else 0.0
        ppk = (mean - lower) / (3 * sigma_overall) if sigma_overall > 0 else 0.0
    elif kind == "upper":
        cpk = (upper - mean) / (3 * sigma_within) if sigma_within > 0 else 0.0
        ppk = (upper - mean) / (3 * sigma_overall) if sigma_overall > 0 else 0.0
    else:
        cpk = (
            min(
                (upper - mean) / (3 * sigma_within),
                (mean - lower) / (3 * sigma_within),
            )
            if sigma_within > 0
            else 0.0
        )
        ppk = (
            min(
                (upper - mean) / (3 * sigma_overall),
                (mean - lower) / (3 * sigma_overall),
            )
            if sigma_overall > 0
            else 0.0
        )
    return {
        "flat": flat,
        "mean": mean,
        "group_means": flat,
        "ranges": moving,
        "moving": moving,
        "mr_bar": mrbar,
        "rbar": mrbar,
        "sigma_within": sigma_within,
        "sigma_overall": sigma_overall,
        "cpk": cpk,
        "ppk": ppk,
    }


def _xmr_in_control(metrics: dict) -> bool:
    """X 图：单值落在 Xbar±2.66MRbar；MR 图：移动极差不超过 3.267MRbar。"""
    mrbar = metrics["mr_bar"]
    if mrbar <= 0:
        return False
    mean = metrics["mean"]
    lower_x = mean - 2.66 * mrbar
    upper_x = mean + 2.66 * mrbar
    if any(
        value < lower_x - 1e-9 or value > upper_x + 1e-9
        for value in metrics["flat"]
    ):
        return False
    return all(
        value <= XMR_D4 * mrbar + 1e-9 for value in metrics["moving"]
    )


def _generate_xmr_values(
    item: SpcItem,
) -> tuple[list[list[float]], dict, int]:
    """生成 25 个单值，满足单值极差图的判异规则。"""
    lower = float(item.lower)
    target = float(item.target)
    upper = float(item.upper)
    step = float(item.precision)
    decimals = _step_decimals(step)
    sigma_spec = (upper - lower) / 6 / 1.8
    scale = 1.0
    last: dict | None = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        sigma = sigma_spec * scale
        flat = [
            _ceil_to_step(random.gauss(target, sigma), step, decimals)
            for _ in range(XMR_POINTS)
        ]
        if not all(lower - 1e-9 <= value <= upper + 1e-9 for value in flat):
            continue

        metrics = _xmr_metrics(flat, lower, upper, item.spec_kind)
        last = metrics
        cpk = metrics["cpk"]
        ppk = metrics["ppk"]
        ok = (
            len({round(value, 9) for value in flat}) > 1
            and CPK_MIN <= cpk <= CPK_MAX
            and PPK_MIN <= ppk <= PPK_MAX
            and _xmr_in_control(metrics)
            and not _has_seven_run(flat)
        )
        if ok:
            info(
                "[SPC-生成]",
                f"序号={item.index} 单值极差 attempts={attempt} "
                f"cpk={cpk:.3f} ppk={ppk:.3f}",
            )
            return [[value] for value in flat], metrics, attempt

        if cpk > CPK_MAX:
            scale = min(1.8, scale * cpk / CPK_TARGET)
        elif cpk < CPK_MIN:
            scale = max(0.45, scale * max(cpk, 0.6) / CPK_TARGET)
        else:
            scale = min(1.8, max(0.45, scale * random.uniform(0.98, 1.02)))

    detail = ""
    if last is not None:
        detail = (
            f"（最后一次 CPK={last['cpk']:.2f}、PPK={last['ppk']:.2f}）"
        )
    raise SpcGenerateError(
        f"序号 {item.index} 在 {MAX_ATTEMPTS} 次内未生成合格数据{detail}，"
        "请检查上下限与量具精度是否合理。"
    )


def _sanitize_sheet_title(index: str, description: str, used: set[str]) -> str:
    raw = f"{index}.{description}"
    cleaned = re.sub(r"[\\/*?:\[\]：]", "", raw)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip("'")
    if not cleaned:
        cleaned = f"{index}.SPC"
    cleaned = cleaned[:31]
    candidate = cleaned
    suffix = 2
    while candidate.lower() in used:
        tail = f"({suffix})"
        candidate = cleaned[: 31 - len(tail)] + tail
        suffix += 1
    used.add(candidate.lower())
    return candidate


def _retarget_ref(holder, old_title: str, new_title: str) -> None:
    if holder is None:
        return
    for attr in ("numRef", "strRef"):
        ref = getattr(holder, attr, None)
        if ref is None or not getattr(ref, "f", None):
            continue
        ref.f = ref.f.replace(f"'{old_title}'!", f"'{new_title}'!")
        ref.f = ref.f.replace(f"{old_title}!", f"{new_title}!")


def _add_charts(templates, target, old_title: str, new_title: str) -> None:
    for chart in templates:
        clone = deepcopy(chart)
        for series in clone.series:
            _retarget_ref(series.val, old_title, new_title)
            _retarget_ref(series.cat, old_title, new_title)
            _retarget_ref(series.tx, old_title, new_title)
        target.add_chart(clone)


@dataclass
class ImageSpec:
    payload: bytes
    anchor: object
    width: int
    height: int


def _read_images(source: Path, worksheet) -> list[ImageSpec]:
    """从 xlsx 压缩包里直接读取图片二进制，避免复用已关闭的文件句柄。"""
    images = list(worksheet._images)
    if not images:
        return []
    try:
        with zipfile.ZipFile(source) as archive:
            media = [
                name
                for name in archive.namelist()
                if name.startswith("xl/media/")
                and Path(name).suffix.lower() in (".png", ".jpg", ".jpeg", ".gif", ".bmp")
            ]
            payloads = [archive.read(name) for name in media]
    except (OSError, zipfile.BadZipFile):
        return []
    specs: list[ImageSpec] = []
    for index, image in enumerate(images):
        if not payloads:
            break
        payload = payloads[index % len(payloads)]
        specs.append(
            ImageSpec(
                payload=payload,
                anchor=deepcopy(image.anchor),
                width=image.width,
                height=image.height,
            )
        )
    return specs


def _add_images(specs: list[ImageSpec], worksheet) -> None:
    for spec in specs:
        clone = XLImage(BytesIO(spec.payload))
        clone.width = spec.width
        clone.height = spec.height
        clone.anchor = deepcopy(spec.anchor)
        worksheet.add_image(clone)


PLAN_SHEET_TITLE = "SPC计划"


def _copy_print_area(source, target) -> None:
    """复制打印区域；openpyxl 读取到的是带原表名的字符串，需只取单元格范围。"""
    areas = source.print_area
    if not areas:
        return
    if isinstance(areas, str):
        areas = [areas]
    coords = [str(area).split("!")[-1] for area in areas]
    target.print_area = coords


def _copy_plan_sheet(imported: SpcPlanImport, workbook):
    """把导入的 SPC 计划表复制为生成文件的第一页（保留样式、合并与图片）。"""
    source_path = Path(imported.file_path)
    try:
        source_book = load_workbook(source_path, data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise SpcGenerateError(f"无法读取 SPC 计划原表：{exc}") from exc
    try:
        source = source_book[imported.sheet_name]
    except KeyError as exc:
        raise SpcGenerateError(
            f"原表中找不到工作表“{imported.sheet_name}”。"
        ) from exc

    target = workbook.create_sheet(PLAN_SHEET_TITLE, 0)
    for row in source.iter_rows():
        for cell in row:
            clone = target.cell(cell.row, cell.column)
            if cell.value is not None:
                clone.value = cell.value
            if cell.has_style:
                clone.font = copy(cell.font)
                clone.border = copy(cell.border)
                clone.fill = copy(cell.fill)
                clone.number_format = cell.number_format
                clone.protection = copy(cell.protection)
                clone.alignment = copy(cell.alignment)
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))
    for key, dimension in source.column_dimensions.items():
        clone = target.column_dimensions[key]
        clone.width = dimension.width
        clone.hidden = dimension.hidden
        clone.bestFit = dimension.bestFit
    for key, dimension in source.row_dimensions.items():
        clone = target.row_dimensions[key]
        clone.height = dimension.height
        clone.hidden = dimension.hidden
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    target.sheet_view.zoomScale = source.sheet_view.zoomScale
    target.sheet_format = deepcopy(source.sheet_format)
    target.page_margins = deepcopy(source.page_margins)
    target.page_setup = deepcopy(source.page_setup)
    target.print_options = deepcopy(source.print_options)
    _copy_print_area(source, target)
    if source.sheet_properties.tabColor is not None:
        target.sheet_properties.tabColor = copy(source.sheet_properties.tabColor)

    image_specs = _read_images(source_path, source)
    if image_specs:
        _add_images(image_specs, target)
    source_book.close()
    return target


def _copy_template_sheet(source, target_book, title: str):
    """把模板页复制到主工作簿：公式、样式、合并、行高列宽一并带走。"""
    target = target_book.create_sheet(title)
    for row in source.iter_rows():
        for cell in row:
            clone = target.cell(cell.row, cell.column)
            if cell.value is not None:
                clone.value = cell.value
            if cell.has_style:
                clone.font = copy(cell.font)
                clone.border = copy(cell.border)
                clone.fill = copy(cell.fill)
                clone.number_format = cell.number_format
                clone.protection = copy(cell.protection)
                clone.alignment = copy(cell.alignment)
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))
    for key, dimension in source.column_dimensions.items():
        clone = target.column_dimensions[key]
        clone.width = dimension.width
        clone.hidden = dimension.hidden
        clone.bestFit = dimension.bestFit
    for key, dimension in source.row_dimensions.items():
        clone = target.row_dimensions[key]
        clone.height = dimension.height
        clone.hidden = dimension.hidden
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    target.sheet_view.zoomScale = source.sheet_view.zoomScale
    target.sheet_format = deepcopy(source.sheet_format)
    target.page_margins = deepcopy(source.page_margins)
    target.page_setup = deepcopy(source.page_setup)
    target.print_options = deepcopy(source.print_options)
    _copy_print_area(source, target)
    if source.sheet_properties.tabColor is not None:
        target.sheet_properties.tabColor = copy(source.sheet_properties.tabColor)
    return target


def _fill_xmr_sheet(
    worksheet,
    item: SpcItem,
    result: ItemResult,
    user_name: str,
    today: date,
    part_name: str,
    part_number: str,
) -> None:
    """单值极差图：只写计划里真正给出的规格界限，没有中心值就留空 K5。"""
    worksheet["M4"] = 1
    kind = item.spec_kind
    if kind == "lower":
        worksheet["K4"] = None
        worksheet["K5"] = None
        worksheet["K6"] = result.lower
    elif kind == "upper":
        worksheet["K4"] = result.upper
        worksheet["K5"] = None
        worksheet["K6"] = None
    elif kind == "range":
        worksheet["K4"] = result.upper
        worksheet["K5"] = None
        worksheet["K6"] = result.lower
    else:
        worksheet["K4"] = result.upper
        worksheet["K5"] = result.target
        worksheet["K6"] = result.lower
    worksheet["W3"] = result.gauge_name
    worksheet["W4"] = result.unit
    worksheet["W5"] = result.description
    worksheet["W6"] = user_name
    worksheet["AC6"] = today.strftime("%Y.%m.%d")
    if part_name:
        worksheet["C5"] = part_name
    if part_number:
        worksheet["C6"] = part_number

    for col in range(GROUPS):
        worksheet.cell(ROW_START, COL_START + col).value = result.values[col][0]


def _num_text(value: float) -> str:
    text = f"{value:.6f}".rstrip("0").rstrip(".")
    return text or "0"


def _wuling_spec_text(item: SpcItem, result: ItemResult) -> str:
    text = _num_text
    if item.spec_kind == "lower":
        return f"≥{text(result.lower)}{result.unit}"
    if item.spec_kind == "upper":
        return f"≤{text(result.upper)}{result.unit}"
    if item.spec_kind == "range":
        return f"{text(result.lower)}~{text(result.upper)}{result.unit}"
    return f"{text(result.target)}±{text(result.target - result.lower)}{result.unit}"


PLAN_MAPPING_FIELDS = (
    "特性描述",
    "下限",
    "目标值",
    "上限",
    "单位",
    "量具名称",
    "量具精度",
)

_PLAN_VALUE_GETTERS = {
    "特性描述": lambda item: item.description,
    "下限": lambda item: item.lower,
    "目标值": lambda item: item.target,
    "上限": lambda item: item.upper,
    "单位": lambda item: item.unit,
    "量具名称": lambda item: item.gauge_name,
    "量具精度": lambda item: item.precision_text,
}


def _fill_plan_values(
    worksheet, items: list[SpcItem], mapping: dict[str, int] | None
) -> None:
    """把确认表中填写的字段回填到计划页对应的标题栏列。"""
    if not mapping:
        return
    for item in items:
        for field, column in mapping.items():
            getter = _PLAN_VALUE_GETTERS.get(field)
            if getter is None:
                continue
            value = getter(item)
            if value is None or value == "":
                continue
            worksheet.cell(row=int(item.row), column=int(column)).value = value


def _fill_wuling_sheet(
    worksheet,
    item: SpcItem,
    result: ItemResult,
    today: date,
    part_name: str,
    part_number: str,
) -> None:
    """按长安模板的填写口径，写五菱模板的标题栏与规格限。"""
    worksheet["C6"] = part_number or None
    worksheet["H6"] = part_name or None
    worksheet["C7"] = today.strftime("%Y.%m.%d")
    worksheet["C8"] = result.description
    worksheet["H8"] = _wuling_spec_text(item, result)
    worksheet["D33"] = result.upper
    worksheet["I33"] = result.lower

    group_index = 0
    for start_row, start_col, count in WULING_BLOCKS:
        for offset in range(count):
            group = result.values[group_index]
            for sample_index, value in enumerate(group):
                worksheet.cell(start_row + sample_index, start_col + offset).value = value
            group_index += 1


def _sheet_parts_from(data: dict[str, bytes]) -> dict[str, str]:
    """工作表标题 → 工作表部件路径。"""
    workbook_xml = data["xl/workbook.xml"].decode("utf-8")
    rels = _rels(data["xl/_rels/workbook.xml.rels"].decode("utf-8"))
    result: dict[str, str] = {}
    for element in re.findall(r"<sheet\b[^>]*/>", workbook_xml):
        name_match = re.search(r'name="([^"]*)"', element)
        rid_match = re.search(r'r:id="([^"]+)"', element)
        if not name_match or not rid_match:
            continue
        name = unescape(name_match.group(1))
        part = _part_path("xl/workbook.xml", rels.get(rid_match.group(1), ""))
        if part and part in data:
            result[name] = part
    return result


def _load_xmr_chart_xml(template: Path) -> tuple[str, str]:
    """按系列引用区分单值图（第 9 行）与移动极差图（第 17 行）。"""
    try:
        with zipfile.ZipFile(template) as archive:
            parts = sorted(
                (
                    name
                    for name in archive.namelist()
                    if re.match(r"xl/charts/chart\d+\.xml$", name)
                ),
                key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),
            )
            payloads = {
                name: archive.read(name).decode("utf-8") for name in parts
            }
    except (OSError, zipfile.BadZipFile) as exc:
        raise SpcGenerateError(f"无法读取单值极差图模板图表：{exc}") from exc

    x_chart = ""
    mr_chart = ""
    for name in parts:
        xml = payloads[name]
        if "$C$9:$AA$9" in xml and not x_chart:
            x_chart = xml
        elif "$C$17:$AA$17" in xml and not mr_chart:
            mr_chart = xml
    if not x_chart or not mr_chart:
        raise SpcGenerateError("单值极差图模板中未找到 X 图 / MR 图。")
    return x_chart, mr_chart


def _load_xmr_drawing(template: Path) -> str:
    """读取模板的绘图部件（含两个图表锚点与文字标注）。"""
    try:
        with zipfile.ZipFile(template) as archive:
            candidates = sorted(
                (
                    name
                    for name in archive.namelist()
                    if re.match(r"xl/drawings/drawing\d+\.xml$", name)
                ),
                key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),
            )
            if not candidates:
                raise SpcGenerateError("单值极差图模板中未找到图表绘图部件。")
            return archive.read(candidates[0]).decode("utf-8")
    except zipfile.BadZipFile as exc:
        raise SpcGenerateError(f"无法读取单值极差图模板绘图：{exc}") from exc


DRAWING_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
CHART_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
IMAGE_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
DRAWING_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawing+xml"
)
CHART_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)


def _xmr_axis_bounds(result: ItemResult) -> tuple[float, float, float]:
    """X 图轴范围覆盖规格界限、控制限与数据；MR 图从 0 到控制上限。"""
    values = [group[0] for group in result.values]
    limits = list(values)
    if result.mr_bar > 0:
        limits.extend(
            [
                result.mean + 2.66 * result.mr_bar,
                result.mean - 2.66 * result.mr_bar,
            ]
        )
    kind = result.spec_kind
    if kind in ("pm", "range"):
        low, high = result.lower, result.upper
    elif kind == "lower":
        low, high = result.lower, max(limits)
    elif kind == "upper":
        low, high = min(limits), result.upper
    else:
        low, high = result.lower, result.upper

    margin = max(high - low, result.upper - result.lower) * 0.05
    low -= margin
    high += margin
    step = _nice_step(high - low)
    low = math.floor(low / step) * step
    high = math.ceil(high / step) * step
    if high <= low:
        high = low + step

    mr_max = max(XMR_D4 * result.mr_bar, result.max_range)
    if mr_max > 0:
        mr_step = _nice_step(mr_max)
        mr_max = math.ceil(mr_max / mr_step) * mr_step
    else:
        mr_max = 1.0
    return round(low, 10), round(high, 10), round(mr_max, 10)


_SHEET_ELEMENTS_AFTER_DRAWING = (
    "<legacyDrawing",
    "<legacyDrawingHF",
    "<picture",
    "<oleObjects",
    "<controls",
    "<webPublishItems",
    "<tableParts",
    "<extLst",
)


def _append_sheet_element(xml: str, snippet: str) -> str:
    """把 drawing 等元素插到 worksheet 结尾前（保持在 schema 要求的顺序）。"""
    for tag in _SHEET_ELEMENTS_AFTER_DRAWING:
        position = xml.find(tag)
        if position != -1:
            return xml[:position] + snippet + xml[position:]
    return xml.replace("</worksheet>", snippet + "</worksheet>")


def _ensure_relationship_namespace(xml: str) -> str:
    """openpyxl 在无关系引用时不声明 xmlns:r，插入 drawing 后需要补上。"""
    if "xmlns:r=" in xml:
        return xml
    start = xml.find("<worksheet")
    if start == -1:
        return xml
    closing = xml.find(">", start)
    if closing == -1:
        return xml
    return (
        xml[:closing]
        + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        + xml[closing:]
    )


def _attach_xmr_charts(
    output: Path,
    template: Path,
    old_title: str,
    axes: dict[str, tuple[float, float, float]],
) -> None:
    """把单值极差图模板里的 X 图 / MR 图挂到生成文件的新工作表中。"""
    if not axes:
        return
    x_chart, mr_chart = _load_xmr_chart_xml(template)
    drawing = _load_xmr_drawing(template)

    with zipfile.ZipFile(output) as archive:
        entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
    data = {item.filename: payload for item, payload in entries}
    sheet_parts = _sheet_parts_from(data)

    drawing_index = max(
        [
            int(match.group(1))
            for name in data
            if (match := re.match(r"xl/drawings/drawing(\d+)\.xml$", name))
        ]
        or [0]
    )
    chart_index = max(
        [
            int(match.group(1))
            for name in data
            if (match := re.match(r"xl/charts/chart(\d+)\.xml$", name))
        ]
        or [0]
    )
    content_types = data["[Content_Types].xml"].decode("utf-8")
    overrides: list[str] = []

    for title, (minimum, maximum, mr_upper) in axes.items():
        part = sheet_parts.get(title)
        if not part:
            continue
        drawing_index += 1
        chart_x = chart_index + 1
        chart_mr = chart_index + 2
        chart_index += 2

        drawing_part = f"xl/drawings/drawing{drawing_index}.xml"
        drawing_rels_part = (
            f"xl/drawings/_rels/drawing{drawing_index}.xml.rels"
        )
        data[drawing_part] = drawing.encode("utf-8")
        data[drawing_rels_part] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships">'
            f'<Relationship Id="rId1" Type="{CHART_REL}" '
            f'Target="../charts/chart{chart_x}.xml"/>'
            f'<Relationship Id="rId2" Type="{CHART_REL}" '
            f'Target="../charts/chart{chart_mr}.xml"/>'
            "</Relationships>"
        ).encode("utf-8")

        x_xml = _patch_chart_title(x_chart, old_title, title)
        x_xml = _patch_val_axis(
            x_xml,
            minimum=minimum,
            maximum=maximum,
            major_unit=_nice_step(maximum - minimum),
        )
        data[f"xl/charts/chart{chart_x}.xml"] = x_xml.encode("utf-8")
        mr_xml = _patch_chart_title(mr_chart, old_title, title)
        mr_xml = _patch_val_axis(mr_xml, minimum=0, maximum=mr_upper)
        data[f"xl/charts/chart{chart_mr}.xml"] = mr_xml.encode("utf-8")

        rels_part = posixpath.join(
            posixpath.dirname(part),
            "_rels",
            posixpath.basename(part) + ".rels",
        )
        index = 1
        if rels_part in data:
            rels_xml = data[rels_part].decode("utf-8")
            used = set(re.findall(r'Id="([^"]+)"', rels_xml))
            while f"rId{index}" in used:
                index += 1
            addition = (
                f'<Relationship Id="rId{index}" Type="{DRAWING_REL}" '
                f'Target="../drawings/drawing{drawing_index}.xml"/>'
            )
            rels_xml = rels_xml.replace("</Relationships>", addition + "</Relationships>")
        else:
            rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                'relationships">'
                f'<Relationship Id="rId{index}" Type="{DRAWING_REL}" '
                f'Target="../drawings/drawing{drawing_index}.xml"/>'
                "</Relationships>"
            )
        data[rels_part] = rels_xml.encode("utf-8")

        sheet_xml = data[part].decode("utf-8")
        if "<drawing " not in sheet_xml:
            sheet_xml = _append_sheet_element(
                sheet_xml, f'<drawing r:id="rId{index}"/>'
            )
            sheet_xml = _ensure_relationship_namespace(sheet_xml)
            data[part] = sheet_xml.encode("utf-8")

        for candidate in (
            (drawing_part, DRAWING_TYPE),
            (f"xl/charts/chart{chart_x}.xml", CHART_TYPE),
            (f"xl/charts/chart{chart_mr}.xml", CHART_TYPE),
        ):
            if f'PartName="/{candidate[0]}"' in content_types:
                continue
            overrides.append(
                f'<Override PartName="/{candidate[0]}" ContentType="{candidate[1]}"/>'
            )

    if overrides:
        content_types = content_types.replace(
            "</Types>", "".join(overrides) + "</Types>"
        )
        data["[Content_Types].xml"] = content_types.encode("utf-8")

    temp = output.with_name(output.name + ".tmp")
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as writer:
        for info, _ in entries:
            writer.writestr(info.filename, data[info.filename])
        for name, payload in data.items():
            if name not in {info.filename for info, _ in entries}:
                writer.writestr(name, payload)
    _replace_file(temp, output)


_IMAGE_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
}


def _load_wuling_assets(template: Path) -> dict:
    """读取五菱模板的绘图、关系、图表与图片，生成时按原样复用。"""
    try:
        with zipfile.ZipFile(template) as archive:
            names = set(archive.namelist())
            drawing_parts = sorted(
                (
                    name
                    for name in names
                    if re.match(r"xl/drawings/drawing\d+\.xml$", name)
                ),
                key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),
            )
            if not drawing_parts:
                raise SpcGenerateError("五菱 SPC 模板中未找到图表绘图部件。")
            drawing_part = drawing_parts[0]
            rels_part = posixpath.join(
                posixpath.dirname(drawing_part),
                "_rels",
                posixpath.basename(drawing_part) + ".rels",
            )
            rels = (
                _rels(archive.read(rels_part).decode("utf-8"))
                if rels_part in names
                else {}
            )
            charts: dict[str, str] = {}
            media: dict[str, tuple[str, bytes]] = {}
            for rid, target in rels.items():
                part = _part_path(drawing_part, target)
                if part not in names:
                    continue
                if part.startswith("xl/charts/"):
                    charts[rid] = archive.read(part).decode("utf-8")
                elif part.startswith("xl/media/"):
                    media[rid] = (Path(part).suffix.lower(), archive.read(part))
            if not charts:
                raise SpcGenerateError("五菱 SPC 模板中未找到 X 图 / R 图。")
            return {
                "drawing": archive.read(drawing_part).decode("utf-8"),
                "rels": rels,
                "charts": charts,
                "media": media,
            }
    except (OSError, zipfile.BadZipFile) as exc:
        raise SpcGenerateError(f"无法读取五菱 SPC 模板图表：{exc}") from exc


def _patch_wuling_chart(xml: str, old_title: str, title: str) -> str:
    """五菱图表只改工作表引用，坐标轴保持模板的自动缩放。"""
    return _patch_chart_title(xml, old_title, title)


def _attach_wuling_charts(
    output: Path,
    template: Path,
    old_title: str,
    sheet_titles: list[str],
) -> None:
    """把五菱模板的绘图（X 图、R 图、图片与文字标注）挂到生成工作表。"""
    if not sheet_titles:
        return
    assets = _load_wuling_assets(template)
    blueprint = assets["drawing"]
    rels = assets["rels"]
    charts = assets["charts"]
    media = assets["media"]

    with zipfile.ZipFile(output) as archive:
        entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
    data = {item.filename: payload for item, payload in entries}
    sheet_parts = _sheet_parts_from(data)

    drawing_index = max(
        [
            int(match.group(1))
            for name in data
            if (match := re.match(r"xl/drawings/drawing(\d+)\.xml$", name))
        ]
        or [0]
    )
    chart_index = max(
        [
            int(match.group(1))
            for name in data
            if (match := re.match(r"xl/charts/chart(\d+)\.xml$", name))
        ]
        or [0]
    )
    media_index = max(
        [
            int(match.group(1))
            for name in data
            if (match := re.match(r"xl/media/wuling_spc(\d+)\.[A-Za-z]+$", name))
        ]
        or [0]
    )
    content_types = data["[Content_Types].xml"].decode("utf-8")
    defaults: list[str] = []
    overrides: list[str] = []

    media_targets: dict[str, str] = {}
    for rid, (suffix, payload) in media.items():
        media_index += 1
        part = f"xl/media/wuling_spc{media_index}{suffix}"
        while part in data:
            media_index += 1
            part = f"xl/media/wuling_spc{media_index}{suffix}"
        data[part] = payload
        media_targets[rid] = "../media/" + posixpath.basename(part)
        extension = suffix.lstrip(".")
        if not re.search(
            f'<Default Extension="{re.escape(extension)}"', content_types, re.I
        ):
            content_type = _IMAGE_TYPES.get(suffix, "application/octet-stream")
            defaults.append(
                f'<Default Extension="{extension}" ContentType="{content_type}"/>'
            )

    for title in sheet_titles:
        part = sheet_parts.get(title)
        if not part:
            continue
        drawing_index += 1
        drawing_part = f"xl/drawings/drawing{drawing_index}.xml"
        drawing_rels_part = (
            f"xl/drawings/_rels/drawing{drawing_index}.xml.rels"
        )
        rel_parts: list[str] = []
        for rid in rels:
            if rid in charts:
                chart_index += 1
                chart_part = f"xl/charts/chart{chart_index}.xml"
                data[chart_part] = _patch_wuling_chart(
                    charts[rid], old_title, title
                ).encode("utf-8")
                rel_parts.append(
                    f'<Relationship Id="{rid}" Type="{CHART_REL}" '
                    f'Target="../charts/chart{chart_index}.xml"/>'
                )
                if f'PartName="/{chart_part}"' not in content_types:
                    overrides.append(
                        f'<Override PartName="/{chart_part}" '
                        f'ContentType="{CHART_TYPE}"/>'
                    )
            elif rid in media_targets:
                rel_parts.append(
                    f'<Relationship Id="{rid}" Type="{IMAGE_REL}" '
                    f'Target="{media_targets[rid]}"/>'
                )
        data[drawing_part] = blueprint.encode("utf-8")
        data[drawing_rels_part] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships">' + "".join(rel_parts) + "</Relationships>"
        ).encode("utf-8")
        if f'PartName="/{drawing_part}"' not in content_types:
            overrides.append(
                f'<Override PartName="/{drawing_part}" '
                f'ContentType="{DRAWING_TYPE}"/>'
            )

        rels_part = posixpath.join(
            posixpath.dirname(part),
            "_rels",
            posixpath.basename(part) + ".rels",
        )
        index = 1
        if rels_part in data:
            rels_xml = data[rels_part].decode("utf-8")
            used = set(re.findall(r'Id="([^"]+)"', rels_xml))
            while f"rId{index}" in used:
                index += 1
            addition = (
                f'<Relationship Id="rId{index}" Type="{DRAWING_REL}" '
                f'Target="../drawings/drawing{drawing_index}.xml"/>'
            )
            rels_xml = rels_xml.replace("</Relationships>", addition + "</Relationships>")
        else:
            rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                'relationships">'
                f'<Relationship Id="rId{index}" Type="{DRAWING_REL}" '
                f'Target="../drawings/drawing{drawing_index}.xml"/>'
                "</Relationships>"
            )
        data[rels_part] = rels_xml.encode("utf-8")

        sheet_xml = data[part].decode("utf-8")
        if "<drawing " not in sheet_xml:
            sheet_xml = _append_sheet_element(
                sheet_xml, f'<drawing r:id="rId{index}"/>'
            )
            sheet_xml = _ensure_relationship_namespace(sheet_xml)
            data[part] = sheet_xml.encode("utf-8")

    if defaults or overrides:
        content_types = content_types.replace(
            "</Types>", "".join(defaults + overrides) + "</Types>"
        )
        data["[Content_Types].xml"] = content_types.encode("utf-8")

    temp = output.with_name(output.name + ".tmp")
    existing = {info.filename for info, _ in entries}
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as writer:
        for info, _ in entries:
            writer.writestr(info.filename, data[info.filename])
        for name, payload in data.items():
            if name not in existing:
                writer.writestr(name, payload)
    _replace_file(temp, output)


def _fill_analysis(worksheet, imported: SpcPlanImport, results: list[ItemResult]) -> None:
    """把 CPK/PPK 写回“分析结果”列，CPK 评价等级写入“结论”列。"""
    for result in results:
        if not result.row:
            continue
        if imported.analysis_col is not None:
            worksheet.cell(result.row, imported.analysis_col).value = (
                result.analysis_text
            )
        if imported.conclusion_col is not None:
            worksheet.cell(result.row, imported.conclusion_col).value = result.grade


def _load_template_charts(source: Path) -> tuple[str, str]:
    """读取模板里 X 管制图与 R 管制图的原始 XML（用于保留样式）。"""
    parts: list[str] = []
    payloads: dict[str, str] = {}
    try:
        with zipfile.ZipFile(source) as archive:
            parts = sorted(
                (
                    name
                    for name in archive.namelist()
                    if re.match(r"xl/charts/chart\d+\.xml$", name)
                ),
                key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),
            )
            payloads = {
                name: archive.read(name).decode("utf-8") for name in parts
            }
    except (OSError, zipfile.BadZipFile) as exc:
        raise SpcGenerateError(f"无法读取 SPC 模板图表：{exc}") from exc

    x_chart = ""
    r_chart = ""
    for name in parts:
        xml = payloads[name]
        if "$C$16:" in xml and not x_chart:
            x_chart = xml
        elif "$C$17:" in xml and not r_chart:
            r_chart = xml
    if not x_chart or not r_chart:
        raise SpcGenerateError("SPC 模板中未找到 X 管制图 / R 管制图。")
    return x_chart, r_chart


def _patch_chart_title(xml: str, old_title: str, new_title: str) -> str:
    quoted = "'" + new_title.replace("'", "''") + "'"
    xml = xml.replace(f"'{old_title}'!", escape(quoted) + "!")
    # 中文表名在公式里可以不带引号，模板即为此种写法
    return xml.replace(f"{old_title}!", escape(quoted) + "!")


def _nice_step(span: float) -> float:
    """把区间除以 5 后取 1 / 2 / 2.5 / 5 系列的整齐刻度间隔。"""
    if span <= 0:
        return 1.0
    target = span / 5
    exponent = math.floor(math.log10(target))
    base = 10.0 ** exponent
    candidates = [base * factor for factor in (1, 2, 2.5, 5, 10)]
    return min(candidates, key=lambda value: (abs(value - target), value))


def _patch_val_axis(
    xml: str,
    minimum: float | str | None = None,
    maximum: float | str | None = None,
    major_unit: float | None = None,
) -> str:
    """改写第一个数值轴的边界与主要刻度单位。"""

    def patch(match: re.Match) -> str:
        block = match.group(0)
        scaling_match = re.search(r"<c:scaling>.*?</c:scaling>", block, re.S)
        if scaling_match is not None and (minimum is not None or maximum is not None):
            scaling = scaling_match.group(0)
            if maximum is not None:
                text = str(int(maximum)) if isinstance(maximum, int) else str(maximum)
                if re.search(r'<c:max val="[^"]*"\s*/>', scaling):
                    scaling = re.sub(
                        r'<c:max val="[^"]*"\s*/>',
                        f'<c:max val="{text}"/>',
                        scaling,
                        count=1,
                    )
                else:
                    scaling = scaling.replace(
                        "</c:scaling>", f'<c:max val="{text}"/></c:scaling>'
                    )
            if minimum is not None:
                text = str(int(minimum)) if isinstance(minimum, int) else str(minimum)
                if re.search(r'<c:min val="[^"]*"\s*/>', scaling):
                    scaling = re.sub(
                        r'<c:min val="[^"]*"\s*/>',
                        f'<c:min val="{text}"/>',
                        scaling,
                        count=1,
                    )
                else:
                    scaling = scaling.replace(
                        "</c:scaling>", f'<c:min val="{text}"/></c:scaling>'
                    )
            block = block.replace(scaling_match.group(0), scaling)
        if major_unit is not None:
            text = str(int(major_unit)) if major_unit == int(major_unit) else str(major_unit)
            if re.search(r'<c:majorUnit val="[^"]*"\s*/>', block):
                block = re.sub(
                    r'<c:majorUnit val="[^"]*"\s*/>',
                    f'<c:majorUnit val="{text}"/>',
                    block,
                    count=1,
                )
            else:
                block = block.replace(
                    "</c:valAx>", f'<c:majorUnit val="{text}"/></c:valAx>'
                )
        return block

    return re.sub(
        r"<c:valAx>.*?</c:valAx>", patch, xml, count=1, flags=re.S
    )


def _rels(xml: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for element in re.findall(r"<Relationship\b[^>]*/>", xml):
        id_match = re.search(r'Id="([^"]+)"', element)
        target_match = re.search(r'Target="([^"]+)"', element)
        if id_match and target_match:
            result[id_match.group(1)] = target_match.group(1)
    return result


def _part_path(base_part: str, target: str) -> str:
    if not target:
        return ""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target)
    )


def _sheet_chart_parts(archive: zipfile.ZipFile) -> dict[str, list[str]]:
    """从 zip 中解析每个工作表对应的图表部件（按绘图中的先后顺序）。"""
    names = set(archive.namelist())
    workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    workbook_rels = _rels(
        archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    )
    result: dict[str, list[str]] = {}
    for element in re.findall(r"<sheet\b[^>]*/>", workbook_xml):
        name_match = re.search(r'name="([^"]*)"', element)
        rid_match = re.search(r'r:id="([^"]+)"', element)
        if not name_match or not rid_match:
            continue
        name = unescape(name_match.group(1))
        sheet_part = _part_path(
            "xl/workbook.xml", workbook_rels.get(rid_match.group(1), "")
        )
        if not sheet_part or sheet_part not in names:
            continue
        rels_part = posixpath.join(
            posixpath.dirname(sheet_part),
            "_rels",
            posixpath.basename(sheet_part) + ".rels",
        )
        if rels_part not in names:
            continue
        sheet_rels = _rels(archive.read(rels_part).decode("utf-8"))
        drawing_part = ""
        for target in sheet_rels.values():
            part = _part_path(sheet_part, target)
            if part.startswith("xl/drawings/"):
                drawing_part = part
                break
        if not drawing_part or drawing_part not in names:
            continue
        drawing_rels_part = posixpath.join(
            posixpath.dirname(drawing_part),
            "_rels",
            posixpath.basename(drawing_part) + ".rels",
        )
        if drawing_rels_part not in names:
            continue
        drawing_rels = _rels(
            archive.read(drawing_rels_part).decode("utf-8")
        )
        drawing_xml = archive.read(drawing_part).decode("utf-8")
        charts: list[str] = []
        for rid in re.findall(
            r'<[A-Za-z0-9]+:chart\b[^>]*r:id="([^"]+)"', drawing_xml
        ):
            part = _part_path(drawing_part, drawing_rels.get(rid, ""))
            if part.startswith("xl/charts/"):
                charts.append(part)
        if charts:
            result[name] = charts
    return result


def _replace_file(source: Path, target: Path) -> None:
    """替换目标文件；刚生成的 xlsx 偶尔被杀毒软件短暂占用，重试等待。"""
    for attempt in range(10):
        try:
            os.replace(source, target)
            return
        except PermissionError:
            if attempt == 9:
                raise
            time.sleep(0.2)


def _restore_chart_styles(
    output: Path,
    template: Path,
    old_title: str,
    axes: dict[str, tuple[float, float, float]],
) -> None:
    """用模板原始图表 XML 替换 openpyxl 重写后的图表，保留原图样式。"""
    x_chart, r_chart = _load_template_charts(template)
    with zipfile.ZipFile(output) as archive:
        mapping = _sheet_chart_parts(archive)
        replacements: dict[str, bytes] = {}
        for title, (lower, upper, r_upper) in axes.items():
            charts = mapping.get(title, [])
            if len(charts) < 2:
                continue
            x_xml = _patch_chart_title(x_chart, old_title, title)
            x_xml = _patch_val_axis(
                x_xml,
                minimum=math.floor(lower),
                maximum=math.ceil(upper),
                major_unit=_nice_step(upper - lower),
            )
            replacements[charts[0]] = x_xml.encode("utf-8")
            r_xml = _patch_chart_title(r_chart, old_title, title)
            r_xml = _patch_val_axis(r_xml, maximum=math.ceil(r_upper))
            replacements[charts[1]] = r_xml.encode("utf-8")
        if not replacements:
            return
        temp = output.with_name(output.name + ".tmp")
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as writer:
            for item in archive.infolist():
                data = replacements.get(item.filename)
                writer.writestr(
                    item,
                    data if data is not None else archive.read(item.filename),
                )
    _replace_file(temp, output)


def _fill_sheet(
    worksheet,
    item: SpcItem,
    result: ItemResult,
    user_name: str,
    today: date,
    part_name: str,
    part_number: str,
) -> None:
    worksheet["K4"] = result.upper
    worksheet["K5"] = result.target
    worksheet["K6"] = result.lower
    worksheet["M4"] = result.per_group
    worksheet["W3"] = result.gauge_name
    worksheet["W4"] = result.unit
    worksheet["W5"] = result.description
    worksheet["W6"] = user_name
    if part_name:
        worksheet["C5"] = part_name
    if part_number:
        worksheet["C6"] = part_number
    worksheet["AC6"] = today.strftime("%Y.%m.%d")

    last_row = ROW_START + result.per_group - 1
    # 中心值统计范围跟随实际样本行数（原模板固定到第 11 行，5 件样本会算错）
    worksheet["AD16"] = f'=IF(C9="","",AVERAGE(C9:AB{last_row}))'

    for col in range(GROUPS):
        for row in range(MAX_PER_GROUP):
            cell = worksheet.cell(ROW_START + row, COL_START + col)
            if row < result.per_group:
                cell.value = result.values[col][row]
            else:
                cell.value = None


def _generate_wuling(
    imported: SpcPlanImport,
    items: list[SpcItem],
    output_path: Path | None = None,
    plan_column_map: dict[str, int] | None = None,
) -> SpcGenerationResult:
    """按五菱模板生成：固定 5 件/组 × 25 组，逐项复制模板页并填写标题栏。"""
    template = template_path(WULING_TEMPLATE_NAME)
    try:
        source_book = load_workbook(template, data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise SpcGenerateError(f"无法打开五菱 SPC 模板：{exc}") from exc
    source = source_book.worksheets[0]
    source_title = source.title

    workbook = Workbook()
    default_sheet = workbook.active
    plan_sheet = _copy_plan_sheet(imported, workbook)
    workbook.remove(default_sheet)
    used_titles: set[str] = {plan_sheet.title.lower()}

    results: list[ItemResult] = []
    today = date.today()
    for item in items:
        values, metrics, attempts = _generate_values(item, WULING_PER_GROUP)
        sheet_title = _sanitize_sheet_title(item.index, item.description, used_titles)
        result = ItemResult(
            index=item.index,
            description=item.description,
            sheet_name=sheet_title,
            gauge_name=item.gauge_name,
            lower=float(item.lower),
            target=float(item.target),
            upper=float(item.upper),
            unit=item.unit.strip(),
            per_group=WULING_PER_GROUP,
            cpk=metrics["cpk"],
            ppk=metrics["ppk"],
            mean=metrics["mean"],
            sigma_within=metrics["sigma_within"],
            sigma_overall=metrics["sigma_overall"],
            attempts=attempts,
            row=item.row,
            values=values,
            chart_type=CHART_XBAR,
            spec_kind=item.spec_kind,
            mr_bar=0.0,
        )
        worksheet = _copy_template_sheet(source, workbook, sheet_title)
        _fill_wuling_sheet(
            worksheet,
            item,
            result,
            today,
            imported.part_name,
            imported.part_number,
        )
        results.append(result)

    _fill_analysis(plan_sheet, imported, results)
    _fill_plan_values(plan_sheet, items, plan_column_map)
    source_book.close()

    if output_path is None:
        stem = Path(imported.file_path).stem
        output_path = DATA_DIR / "spc_output" / f"{stem}-SPC-五菱.xlsx"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if output_path.exists():
            output_path.unlink()
        workbook.save(output_path)
    except PermissionError as exc:
        raise SpcGenerateError(
            f"无法写入 {output_path}，请关闭已打开的同名文件后重试。"
        ) from exc
    finally:
        workbook.close()

    _attach_wuling_charts(
        output_path,
        template,
        source_title,
        [result.sheet_name for result in results],
    )

    warn(
        "[SPC-完成]",
        f"template=wuling items={len(results)} x={WULING_SAMPLE_COUNT} "
        f"per_group={WULING_PER_GROUP} output={output_path}",
    )
    return SpcGenerationResult(
        output_path=str(output_path),
        sample_count=WULING_SAMPLE_COUNT,
        per_group=WULING_PER_GROUP,
        results=results,
        template_key=TEMPLATE_WULING,
    )


def generate_spc(
    imported: SpcPlanImport,
    items: list[SpcItem],
    sample_count: int,
    user: User,
    output_path: Path | None = None,
    template_key: str = TEMPLATE_CHANGAN,
    plan_column_map: dict[str, int] | None = None,
) -> SpcGenerationResult:
    if not items:
        raise SpcGenerateError("没有可生成的数据项。")
    for item in items:
        if not item.ready:
            raise SpcGenerateError(f"序号 {item.index}：{item.problem}")
    if template_key == TEMPLATE_WULING:
        return _generate_wuling(imported, items, output_path, plan_column_map)
    if template_key not in TEMPLATE_FILES:
        raise SpcGenerateError(f"未知的 SPC 模板：{template_key}")

    xmr_items = [item for item in items if item.chart_type == CHART_XMR]
    xbar_items = [item for item in items if item.chart_type != CHART_XMR]
    # 只有存在均值极差图时才校验样本数量；全是单值极差图时固定 25 个单值
    per_group = validate_sample_count(sample_count) if xbar_items else 1

    template = template_path()
    try:
        workbook = load_workbook(template, data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise SpcGenerateError(f"无法打开 SPC 模板：{exc}") from exc

    base_sheet = workbook.worksheets[0]
    base_title = base_sheet.title
    chart_templates = [deepcopy(chart) for chart in base_sheet._charts]
    base_sheet._charts = []
    image_specs = _read_images(template, base_sheet)
    if image_specs:
        # 模板里加载的图片句柄在保存时会失效，换成内存中的副本
        base_sheet._images = []
        _add_images(image_specs, base_sheet)
    plan_sheet = _copy_plan_sheet(imported, workbook)
    used_titles: set[str] = {plan_sheet.title.lower()}
    if not xbar_items:
        # 全部使用单值极差图时，模板自带的均值极差图页用不上
        workbook.remove(base_sheet)

    xmr_source = None
    xmr_book = None
    xmr_template = None
    xmr_title = ""
    if xmr_items:
        xmr_template = template_path(XMR_TEMPLATE_NAME)
        try:
            xmr_book = load_workbook(xmr_template, data_only=False)
        except Exception as exc:  # noqa: BLE001
            workbook.close()
            raise SpcGenerateError(
                f"无法打开单值极差图模板 {XMR_TEMPLATE_NAME}：{exc}"
            ) from exc
        xmr_source = xmr_book.worksheets[0]
        xmr_title = xmr_source.title

    results: list[ItemResult] = []
    today = date.today()
    first_xbar = True

    for item in items:
        is_xmr = item.chart_type == CHART_XMR
        if is_xmr:
            values, metrics, attempts = _generate_xmr_values(item)
        else:
            values, metrics, attempts = _generate_values(item, per_group)
        sheet_title = _sanitize_sheet_title(item.index, item.description, used_titles)
        result = ItemResult(
            index=item.index,
            description=item.description,
            sheet_name=sheet_title,
            gauge_name=item.gauge_name,
            lower=float(item.lower),
            target=float(item.target),
            upper=float(item.upper),
            unit=item.unit.strip(),
            per_group=1 if is_xmr else per_group,
            cpk=metrics["cpk"],
            ppk=metrics["ppk"],
            mean=metrics["mean"],
            sigma_within=metrics["sigma_within"],
            sigma_overall=metrics["sigma_overall"],
            attempts=attempts,
            row=item.row,
            values=values,
            chart_type=CHART_XMR if is_xmr else CHART_XBAR,
            spec_kind=item.spec_kind,
            mr_bar=metrics.get("mr_bar", 0.0),
        )
        if is_xmr:
            worksheet = _copy_template_sheet(xmr_source, workbook, sheet_title)
            _fill_xmr_sheet(
                worksheet,
                item,
                result,
                user.name,
                today,
                imported.part_name,
                imported.part_number,
            )
        elif first_xbar:
            first_xbar = False
            worksheet = base_sheet
            worksheet.title = sheet_title
            _add_charts(chart_templates, worksheet, base_title, sheet_title)
            _fill_sheet(
                worksheet,
                item,
                result,
                user.name,
                today,
                imported.part_name,
                imported.part_number,
            )
        else:
            worksheet = workbook.copy_worksheet(base_sheet)
            worksheet.title = sheet_title
            _add_charts(chart_templates, worksheet, base_title, sheet_title)
            _add_images(image_specs, worksheet)
            _fill_sheet(
                worksheet,
                item,
                result,
                user.name,
                today,
                imported.part_name,
                imported.part_number,
            )
        results.append(result)

    _fill_analysis(plan_sheet, imported, results)
    if xmr_book is not None:
        xmr_book.close()

    if output_path is None:
        stem = Path(imported.file_path).stem
        output_path = DATA_DIR / "spc_output" / f"{stem}-SPC.xlsx"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if output_path.exists():
            output_path.unlink()
        workbook.save(output_path)
    except PermissionError as exc:
        raise SpcGenerateError(
            f"无法写入 {output_path}，请关闭已打开的同名文件后重试。"
        ) from exc
    finally:
        workbook.close()

    xbar_results = [result for result in results if result.chart_type == CHART_XBAR]
    if xbar_results:
        # openpyxl 重新序列化会丢掉图表样式，这里用模板原始图表还原并设置轴边界
        _restore_chart_styles(
            output_path,
            template,
            base_title,
            {
                result.sheet_name: (result.lower, result.upper, result.r_upper)
                for result in xbar_results
            },
        )

    xmr_results = [result for result in results if result.chart_type == CHART_XMR]
    if xmr_results and xmr_template is not None:
        _attach_xmr_charts(
            output_path,
            xmr_template,
            xmr_title,
            {
                result.sheet_name: _xmr_axis_bounds(result)
                for result in xmr_results
            },
        )

    warn(
        "[SPC-完成]",
        f"items={len(results)} xbar={len(xbar_results)} xmr={len(xmr_results)} "
        f"x={sample_count} per_group={per_group} output={output_path}",
    )
    return SpcGenerationResult(
        output_path=str(output_path),
        sample_count=sample_count if xbar_items else XMR_POINTS,
        per_group=per_group,
        results=results,
        template_key=TEMPLATE_CHANGAN,
    )
