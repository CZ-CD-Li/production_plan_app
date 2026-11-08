"""作业指导书 → 控制计划数据解析。

按照 ``控制计划VBS.md`` 中的 VBA 逻辑，逐 sheet 读取作业指导书中的
“制造条件确认”与“品质特性确认”两段内容，整理成控制计划标准 18 列数据，
并记录每个单元格在模板中需要纵向合并的行数。
"""

from __future__ import annotations

import sys
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.config import ROOT_DIR

PLAN_HEADERS = (
    "工序编号",
    "工序名称",
    "机器工装量具辅材",
    "序号",
    "产品特性",
    "过程特性",
    "特殊特性分类",
    "产品过程要求",
    "评价测量技术",
    "数量",
    "频率",
    "控制方法",
    "反应计划",
    "分类",
    "担当",
    "维护保养",
    "防错管理",
    "记录表单编号",
)

SOURCE_HEADERS = (
    "分类",
    "过程特性",
    "项目",
    "要求",
    "频次",
    "方法",
    "记录表格",
    "反应计划",
)

RECORD_KEYWORDS = (
    ("过程质量", "8731101"),
    ("制造条件", "8722010"),
    ("检验", "8731102"),
    ("照明", "8731102"),
    ("拉压力", "871022102"),
    ("防错", "8722201"),
)

MAINTENANCE_PROCESS_KEYS = (
    "活塞",
    "底阀",
    "固定环",
    "拧紧",
    "工作缸",
    "清洗",
    "注油",
    "力值",
    "充气",
    "口",
    "端盖",
    "装后",
    "定位垫",
    "套管",
    "打码",
)

POKA_EQUIPMENT_KEYS = (
    ".",
    "数片机",
    "拧紧",
    "工作缸",
    "示功",
    "端盖",
    "装簧",
)

GREASE_TEMPLATE_NAME = "涂润滑脂标准作业指导书.xlsx"
GREASE_PROCESS_NAME = "涂润滑脂"


class WorkInstructionError(Exception):
    """作业指导书结构不符合控制计划提取要求。"""


@dataclass(frozen=True)
class SymbolValue:
    """特殊特性符号占位；第 1~4 列对应模板第 28 行的四个符号。"""

    column: int
    text: str

    @property
    def preview(self) -> str:
        mark = "Z" if self.column in (1, 3) else "G"
        return f"{mark}/{self.text}"


@dataclass
class PlanRow:
    values: list[Any] = field(default_factory=lambda: [""] * 18)
    spans: list[int] = field(default_factory=lambda: [0] * 18)
    band_merge: bool = False


@dataclass
class ProcessBlock:
    sheet_name: str
    process_no: str
    process_name: str
    start: int
    count: int


@dataclass
class WorkInstructionParse:
    file_path: str
    product_name: str
    product_code: str
    rows: list[PlanRow] = field(default_factory=list)
    blocks: list[ProcessBlock] = field(default_factory=list)
    skipped_sheets: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def display_value(value: Any) -> str:
    if isinstance(value, SymbolValue):
        return value.preview
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _text(value: Any) -> str:
    if isinstance(value, SymbolValue):
        return value.text
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_empty(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _find_cell(ws: Worksheet, text: str):
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if isinstance(value, str):
                if value.strip() == text:
                    return cell
            elif value == text:
                return cell
    return None


def _right_value(ws: Worksheet, label) -> Any:
    """取标签右侧的值；若标签本身是横向合并，则跳过整个合并区。"""
    target_col = label.column + 1
    for merged in ws.merged_cells.ranges:
        if label.coordinate in merged:
            target_col = merged.max_col + 1
            break
    if target_col > ws.max_column:
        return ""
    return ws.cell(label.row, target_col).value


def _merge_index(ws: Worksheet) -> dict[tuple[int, int], tuple[int, int]]:
    """记录每个合并区左上角单元格 -> (纵向行数, 最右列)。"""
    result: dict[tuple[int, int], tuple[int, int]] = {}
    for merged in ws.merged_cells.ranges:
        result[(merged.min_row, merged.min_col)] = (
            merged.max_row - merged.min_row + 1,
            merged.max_col,
        )
    return result


def _merge_span(
    index: dict[tuple[int, int], tuple[int, int]], row: int, col: int
) -> tuple[int, int]:
    return index.get((row, col), (1, col))


def _adjusted_rows(span: int, start_row: int, empty_rows: dict[int, bool]) -> int:
    original = span
    for offset in range(2, original + 1):
        if empty_rows.get(start_row + offset - 1, False):
            span -= 1
    return max(span, 1)


def _row_is_empty(ws: Worksheet, row: int, cols: dict[str, int]) -> bool:
    for name in SOURCE_HEADERS[2:]:
        if not _is_empty(ws.cell(row, cols[name]).value):
            return False
    return True


def _split_frequency(value: Any) -> tuple[Any, Any]:
    text = _text(value)
    if "/" in text:
        parts = text.split("/")
        return parts[0], parts[1] if len(parts) > 1 else ""
    return value, ""


def _record_number(value: Any) -> str:
    text = _text(value)
    result = "——"
    for keyword, number in RECORD_KEYWORDS:
        if keyword in text:
            result = number
    return result


def _apply_common_fields(
    ws: Worksheet,
    row: int,
    cols: dict[str, int],
    plan: PlanRow,
    index: dict[tuple[int, int], tuple[int, int]],
    empty_rows: dict[int, bool],
    process_name: str,
) -> None:
    """写入要求、方法、记录表格、反应计划、频次（两段共用的列）。"""
    requirement = ws.cell(row, cols["要求"]).value
    plan.values[7] = "" if requirement is None else requirement
    if not _is_empty(requirement):
        plan.spans[7] = 1
    span, _ = _merge_span(index, row, cols["要求"])
    if span > 1:
        plan.spans[7] = _adjusted_rows(span, row, empty_rows)

    method = ws.cell(row, cols["方法"]).value
    plan.values[8] = "" if method is None else method
    if not _is_empty(method):
        plan.spans[8] = 1
    span, _ = _merge_span(index, row, cols["方法"])
    if span > 1:
        plan.spans[8] = _adjusted_rows(span, row, empty_rows)

    record = ws.cell(row, cols["记录表格"]).value
    if not _is_empty(record):
        plan.values[11] = record
        plan.values[14] = "检验员" if "检验" in process_name else "操作者"
        plan.values[17] = _record_number(record)
        plan.spans[11] = 1
        plan.spans[14] = 1
        plan.spans[17] = 1
        span, _ = _merge_span(index, row, cols["记录表格"])
        if span > 1:
            adjusted = _adjusted_rows(span, row, empty_rows)
            plan.spans[11] = adjusted
            plan.spans[14] = adjusted
            plan.spans[17] = adjusted

    reaction = ws.cell(row, cols["反应计划"]).value
    plan.values[12] = "" if reaction is None else reaction
    if not _is_empty(reaction):
        plan.spans[12] = 1
    span, _ = _merge_span(index, row, cols["反应计划"])
    if span > 1:
        plan.spans[12] = _adjusted_rows(span, row, empty_rows)

    frequency = ws.cell(row, cols["频次"]).value
    if not _is_empty(frequency):
        plan.spans[9] = 1
        plan.spans[10] = 1
        quantity, rate = _split_frequency(frequency)
        plan.values[9] = quantity
        plan.values[10] = rate
    span, _ = _merge_span(index, row, cols["频次"])
    if span > 1:
        adjusted = _adjusted_rows(span, row, empty_rows)
        plan.spans[9] = adjusted
        plan.spans[10] = adjusted


class _BlockParser:
    def __init__(self, ws: Worksheet, parse: WorkInstructionParse) -> None:
        self.ws = ws
        self.result = parse
        self.rows = parse.rows
        self.start = len(parse.rows)
        self.process_name = ""
        self.process_no = ""
        self.equipment_name = ""
        self.equipment_model = ""
        self.next_seq = 1
        self.param_rows = 0
        self.poka_inserted = 0

    def parse(self) -> None:
        ws = self.ws
        quality = _find_cell(ws, "品质特性确认")
        if quality is None:
            self.result.skipped_sheets.append(ws.title)
            return

        process_name_cell = _find_cell(ws, "工序名称")
        process_no_cell = _find_cell(ws, "工序编号")
        equipment_cell = _find_cell(ws, "设备名称")
        model_cell = _find_cell(ws, "设备型号")
        if process_name_cell is None or process_no_cell is None:
            raise WorkInstructionError(
                f"工作表“{ws.title}”缺少“工序名称”或“工序编号”标定。"
            )
        self.process_name = _text(_right_value(ws, process_name_cell))
        process_no = _right_value(ws, process_no_cell)
        self.process_no = _text(process_no)
        self.equipment_name = (
            _text(_right_value(ws, equipment_cell))
            if equipment_cell is not None
            else ""
        )
        self.equipment_model = (
            _text(_right_value(ws, model_cell)) if model_cell is not None else ""
        )
        if self.equipment_model in ("——", "--"):
            self.equipment_model = ""

        manufacturing = _find_cell(ws, "制造条件确认")
        if manufacturing is None:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”缺少“制造条件确认”标定，"
                "需要“制造条件确认”及下一行表头。"
            )
        if quality.row < manufacturing.row:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”的“品质特性确认”位于"
                "“制造条件确认”之前，请调整。"
            )
        if quality.column != manufacturing.column:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”的“品质特性确认”与"
                "“制造条件确认”不在同一列，请调整。"
            )

        quality_width = 1
        for merged in ws.merged_cells.ranges:
            if quality.coordinate in merged:
                quality_width = merged.max_col - merged.min_col + 1
                break

        header_row = manufacturing.row + 1
        cols: dict[str, int] = {}
        for name in SOURCE_HEADERS:
            for col in range(quality.column, quality.column + quality_width):
                value = ws.cell(header_row, col).value
                if isinstance(value, str) and value.strip() == name:
                    cols[name] = col
                    break
        missing = [name for name in SOURCE_HEADERS if name not in cols]
        if missing:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”的表头缺少："
                + "、".join(missing)
            )
        if cols["反应计划"] < cols["项目"]:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”的“反应计划”列位置异常。"
            )

        safety = _find_cell(ws, "安全要求")
        if safety is None:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”缺少“安全要求”标定。"
            )
        max_row = safety.row
        if max_row < quality.row:
            raise WorkInstructionError(
                f"工序“{self.process_name or ws.title}”的“安全要求”必须位于"
                "品质特性确认下方。"
            )

        index = _merge_index(ws)
        empty_rows = {
            row: _row_is_empty(ws, row, cols)
            for row in range(manufacturing.row + 2, max_row + 1)
        }
        category = self._parse_manufacturing(
            manufacturing.row, quality.row, cols, index, empty_rows
        )
        quality_start = len(self.rows)
        self._parse_quality(quality.row, max_row, cols, index, empty_rows, category)

        count = len(self.rows) - self.start
        if count == 0:
            self.result.warnings.append(
                f"工序“{self.process_name or ws.title}”未读取到数据行，已跳过。"
            )
            return

        first = self.rows[self.start]
        first.values[0] = process_no
        first.values[1] = self.process_name
        first.spans[0] = count
        first.spans[1] = count

        quality_count = len(self.rows) - quality_start
        if quality_count:
            first_quality = self.rows[quality_start]
            first_quality.values[15] = "——"
            first_quality.values[16] = "——"
            first_quality.spans[2] = quality_count
            first_quality.spans[15] = quality_count
            first_quality.spans[16] = quality_count

        self.result.blocks.append(
            ProcessBlock(
                sheet_name=ws.title,
                process_no=self.process_no,
                process_name=self.process_name,
                start=self.start,
                count=count,
            )
        )

    def _parse_manufacturing(
        self,
        manufacturing_row: int,
        quality_row: int,
        cols: dict[str, int],
        index: dict[tuple[int, int], tuple[int, int]],
        empty_rows: dict[int, bool],
    ) -> str:
        ws = self.ws
        block_start = self.start
        category = ""
        tooling_total = 0
        gauge_total = 0
        material_total = 0

        for row in range(manufacturing_row + 2, quality_row):
            if not _is_empty(ws.cell(row, cols["分类"]).value):
                category = _text(ws.cell(row, cols["分类"]).value)
            if empty_rows.get(row, True):
                continue

            plan = PlanRow()
            plan.values[13] = category
            plan.spans[13] = 1

            self._write_process_characteristic(plan, row, cols)

            project = ws.cell(row, cols["项目"]).value
            if not _is_empty(project):
                plan.values[3] = self.next_seq
                plan.spans[3] = 1
                plan.spans[4] = 1
                plan.spans[5] = 1
                plan.spans[6] = 1
                self.next_seq += 1

            if category == "工艺参数":
                if self.param_rows == 0:
                    plan.values[2] = self.equipment_name + self.equipment_model
                plan.values[5] = "" if project is None else project
                if not _is_empty(project):
                    span, _ = _merge_span(index, row, cols["项目"])
                    if span > 1:
                        adjusted = _adjusted_rows(span, row, empty_rows)
                        plan.spans[3] = adjusted
                        plan.spans[4] = adjusted
                        plan.spans[5] = adjusted
                        plan.spans[6] = adjusted
                self.param_rows += 1
            else:
                previous_is_param = bool(self.rows) and (
                    self.rows[-1].values[13] == "工艺参数"
                )
                if previous_is_param:
                    if self._needs_poka():
                        extra = PlanRow()
                        extra.values[2] = self.equipment_name
                        extra.values[5] = "防错"
                        extra.values[7] = "按《防错项目管理清单》执行"
                        extra.values[12] = "反应计划(5)"
                        extra.values[13] = "防错"
                        extra.values[14] = "操作者"
                        extra.values[15] = "——"
                        extra.values[16] = "——"
                        extra.values[17] = "8722201"
                        for target in (2, 7, 12, 13, 14, 15, 16, 17):
                            extra.spans[target] = 1
                        # VBA：防错行占用当前序号，本行顺延取下一个序号
                        if not _is_empty(project):
                            extra.values[3] = plan.values[3]
                            extra.spans[3] = 1
                        extra.band_merge = True
                        self.rows.append(extra)
                        self.rows[block_start].values[16] = "《防错项目管理清单》"
                        plan.values[3] = self.next_seq
                        plan.spans[3] = 1
                        self.next_seq += 1
                        self.poka_inserted = 1
                    else:
                        self.rows[block_start].values[16] = "——"

                if not _is_empty(project):
                    plan.spans[2] = 1

                if category == "工装":
                    plan.values[5] = "外观"
                    plan.values[13] = "工装"
                    plan.spans[13] = 1
                    tooling_total += 1
                elif category == "量具":
                    plan.values[5] = "有效性"
                    gauge_total += 1
                elif category == "辅材":
                    plan.values[5] = "型号规格"
                    material_total += 1

                plan.values[2] = "" if project is None else project
                span, _ = _merge_span(index, row, cols["项目"])
                if span > 1:
                    adjusted = _adjusted_rows(span, row, empty_rows)
                    plan.spans[2] = adjusted
                    plan.spans[3] = adjusted
                    plan.spans[4] = adjusted
                    plan.spans[5] = adjusted
                    plan.spans[6] = adjusted

            _apply_common_fields(
                ws, row, cols, plan, index, empty_rows, self.process_name
            )
            self.rows.append(plan)

        if block_start < len(self.rows):
            if self.param_rows:
                self.rows[block_start].spans[2] = self.param_rows
                self.rows[block_start].spans[16] = self.param_rows

            offset = block_start + self.param_rows + self.poka_inserted
            for count, _target in (
                (tooling_total, "tooling"),
                (gauge_total, "gauge"),
                (material_total, "material"),
            ):
                if not count:
                    continue
                item = self.rows[offset]
                for column in (15, 16):
                    item.spans[column] = count
                    item.values[column] = "——"
                offset += count

            if self.param_rows and self._needs_maintenance():
                self.rows[block_start].spans[15] = self.param_rows
                self.rows[block_start].values[15] = "《设备点检、润滑基准书》"

        return category

    def _needs_poka(self) -> bool:
        return any(
            key in self.equipment_name for key in POKA_EQUIPMENT_KEYS
        ) or "阻尼力" in self.process_name

    def _needs_maintenance(self) -> bool:
        if any(key in self.process_name for key in MAINTENANCE_PROCESS_KEYS):
            return True
        return "装簧" in self.equipment_name

    def _write_process_characteristic(
        self, plan: PlanRow, row: int, cols: dict[str, int]
    ) -> None:
        value = self.ws.cell(row, cols["过程特性"]).value
        text = _text(value)
        if not text:
            return
        if "/" in text:
            plan.values[6] = value
        elif "Z" in text.upper():
            plan.values[6] = SymbolValue(1, text)
        elif "G" in text.upper():
            plan.values[6] = SymbolValue(2, text)

    def _parse_quality(
        self,
        quality_row: int,
        max_row: int,
        cols: dict[str, int],
        index: dict[tuple[int, int], tuple[int, int]],
        empty_rows: dict[int, bool],
        category: str,
    ) -> None:
        ws = self.ws
        if not category:
            category = "品质确认"
        quality_rows = 0

        for row in range(quality_row + 2, max_row + 1):
            if not _is_empty(ws.cell(row, cols["分类"]).value):
                category = "品质确认"
            if empty_rows.get(row, True):
                continue

            plan = PlanRow()
            self.rows.append(plan)
            quality_rows += 1
            plan.values[13] = category
            plan.spans[13] = 1
            self._write_quality_characteristic(plan, row, cols)

            project = ws.cell(row, cols["项目"]).value
            if not _is_empty(project):
                plan.values[3] = self.next_seq
                plan.spans[3] = 1
                plan.spans[4] = 1
                plan.spans[5] = 1
                plan.spans[6] = 1
                self.next_seq += 1

            plan.values[4] = "" if project is None else project
            span, _ = _merge_span(index, row, cols["项目"])
            if span > 1:
                adjusted = _adjusted_rows(span, row, empty_rows)
                plan.spans[3] = adjusted
                plan.spans[4] = adjusted
                plan.spans[5] = adjusted
                plan.spans[6] = adjusted

            _apply_common_fields(
                ws, row, cols, plan, index, empty_rows, self.process_name
            )

            if row == max_row and "检验" in self.process_name and "总成" in self.process_name:
                self._append_audit_rows(category)
                quality_rows += 4

        if quality_rows == 0:
            self.result.warnings.append(
                f"工序“{self.process_name or ws.title}”未读取到品质特性数据行。"
            )

    def _write_quality_characteristic(
        self, plan: PlanRow, row: int, cols: dict[str, int]
    ) -> None:
        value = self.ws.cell(row, cols["过程特性"]).value
        text = _text(value)
        if not text:
            return
        if "/" in text:
            plan.values[6] = value
        elif "Z" in text.upper():
            plan.values[6] = SymbolValue(3, text)
        elif "G" in text.upper():
            plan.values[6] = SymbolValue(4, text)

    def _append_audit_rows(self, category: str) -> None:
        definitions = (
            ("型式试验", "型式试验计划及报告", "参照《型式试验、测试失败或检查不通过时的应对》"),
            ("产品审核", "产品审核计划", ""),
            ("全尺寸检测", "全尺寸检测计划及报告", ""),
            ("MSA分析", "MSA分析计划及报告", ""),
        )
        for position, (name, plan_text, reaction) in enumerate(definitions):
            plan = PlanRow()
            plan.values[3] = self.next_seq
            self.next_seq += 1
            plan.values[4] = name
            plan.values[7] = plan_text
            plan.values[13] = category
            plan.values[14] = "检验员"
            plan.values[17] = "——"
            plan.spans[3] = 1
            plan.spans[4] = 1
            plan.spans[7] = 1
            plan.spans[13] = 1
            plan.spans[14] = 1
            plan.spans[17] = 1
            if position == 0:
                plan.values[12] = reaction
                plan.spans[12] = 4
            plan.band_merge = True
            self.rows.append(plan)


def parse_work_instruction(file_path: str | Path) -> WorkInstructionParse:
    path = Path(file_path)
    if not path.is_file():
        raise WorkInstructionError(f"文件不存在：{path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise WorkInstructionError("仅支持 .xlsx / .xlsm 作业指导书。")

    try:
        workbook = load_workbook(path, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise WorkInstructionError(f"无法打开作业指导书：{exc}") from exc

    parse = WorkInstructionParse(
        file_path=str(path), product_name="", product_code=""
    )
    try:
        process_sheets = 0
        for worksheet in workbook.worksheets:
            quality = _find_cell(worksheet, "品质特性确认")
            if quality is None:
                parse.skipped_sheets.append(worksheet.title)
                continue
            process_sheets += 1
            if not parse.product_name:
                name_label = _find_cell(worksheet, "产品名称")
                if name_label is not None:
                    parse.product_name = _text(_right_value(worksheet, name_label))
            if not parse.product_code:
                code_label = _find_cell(worksheet, "产品代号")
                if code_label is not None:
                    parse.product_code = _text(_right_value(worksheet, code_label))
            _BlockParser(worksheet, parse).parse()
        if process_sheets == 0:
            raise WorkInstructionError(
                "未找到包含“品质特性确认”的工序工作表，请确认文件是否为作业指导书。"
            )
        if not parse.rows:
            raise WorkInstructionError("未读取到可转换的控制计划数据。")
    finally:
        workbook.close()
    return parse


def is_grease_process(name: str) -> bool:
    text = str(name).replace(" ", "")
    return GREASE_PROCESS_NAME in text or "油封涂润滑脂" in text


def grease_template_path() -> Path:
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(
            Path(meipass) / "data" / "templates" / GREASE_TEMPLATE_NAME
        )
    candidates.append(ROOT_DIR / "data" / "templates" / GREASE_TEMPLATE_NAME)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise WorkInstructionError(
        f"未找到涂润滑脂作业指导书 {GREASE_TEMPLATE_NAME}，"
        "请确认 data/templates 是否随程序打包。"
    )


def load_grease_plan_rows() -> list[PlanRow]:
    """读取涂润滑脂标准作业指导书，返回其控制计划数据行（工序号固定为 1）。"""
    parsed = parse_work_instruction(grease_template_path())
    for block in parsed.blocks:
        if is_grease_process(block.process_name):
            rows = deepcopy(parsed.rows[block.start : block.start + block.count])
            for row in rows:
                row.values[0] = "1"
            return rows
    raise WorkInstructionError("涂润滑脂作业指导书中未识别到涂润滑脂工序。")


def _next_process_no(value: Any) -> str:
    text = display_value(value)
    try:
        return str(int(float(text)) + 1)
    except ValueError:
        return text


def ensure_grease_process(parse: WorkInstructionParse) -> bool:
    """缺少油封涂润滑脂工序时，将其作为序号 1 插入，其余工序序号后移。"""
    if any(is_grease_process(block.process_name) for block in parse.blocks):
        return False
    grease_rows = load_grease_plan_rows()
    count = len(grease_rows)
    for block in parse.blocks:
        for row in parse.rows[block.start : block.start + block.count]:
            row.values[0] = _next_process_no(row.values[0])
        block.process_no = _next_process_no(block.process_no)
        block.start += count
    parse.rows = grease_rows + parse.rows
    parse.blocks.insert(
        0,
        ProcessBlock(
            sheet_name=GREASE_PROCESS_NAME,
            process_no="1",
            process_name=GREASE_PROCESS_NAME,
            start=0,
            count=count,
        ),
    )
    return True
