"""控制计划 → CAPP 质量策划 Excel 生成。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from app.capp.importer import ControlPlanImport
from app.capp.rules import (
    KZJH_COLUMNS,
    TARGET_FIELDS,
    build_quality_rows,
    source_dedup_key,
)
from app.capp.work_instruction import (
    display_value,
    is_grease_process,
    load_grease_plan_rows,
)
from app.config import DATA_DIR
from app.database import Database
from app.models import User

OUTPUT_DIR = DATA_DIR / "capp_output"


class CappGenerateError(Exception):
    pass


@dataclass
class GeneratedResult:
    product_code: str
    output_path: str
    row_count: int
    source_count: int
    removed_source: int
    warnings: list[str] = field(default_factory=list)


def extract_product_code(file_path: str | Path) -> str:
    stem = Path(file_path).stem
    match = re.search(r"[A-Za-z0-9]+", stem)
    return match.group() if match else stem


def build_source_rows(
    imported: ControlPlanImport, mapping: dict[str, int | None]
) -> tuple[list[dict[str, str]], int]:
    positions = {column.index: index for index, column in enumerate(imported.columns)}
    sources: list[dict[str, str]] = []
    seen: set[tuple[str, ...]] = set()
    removed = 0
    for values in imported.rows:
        source: dict[str, str] = {}
        for target in TARGET_FIELDS:
            column_index = mapping.get(target)
            position = positions.get(column_index) if column_index is not None else None
            source[target] = values[position] if position is not None else ""
        if not any(source.values()):
            continue
        key = source_dedup_key(source)
        if key in seen:
            removed += 1
            continue
        seen.add(key)
        sources.append(source)
    return sources, removed


# 作业指导书 18 列 → 控制计划目标字段
_PLAN_TARGET_COLUMNS: tuple[tuple[int, str], ...] = (
    (0, "过程编号（工序号）"),
    (1, "过程名称"),
    (2, "机器、装置夹具、工装"),
    (3, "特性编号"),
    (4, "产品特性"),
    (5, "过程特性"),
    (6, "特殊特性分类（特殊特性）"),
    (7, "规范/公差"),
    (8, "评价测量技术（评价测量技术）"),
    (9, "样本数量"),
    (10, "样本频率"),
    (11, "控制方法"),
    (12, "反应计划"),
)


def grease_source_rows() -> list[dict[str, str]]:
    """把涂润滑脂作业指导书的工序数据转换成质量策划源行。"""
    rows: list[dict[str, str]] = []
    carry: dict[str, str] = {}
    for plan_row in load_grease_plan_rows():
        source = {target: "" for target in TARGET_FIELDS}
        for index, target in _PLAN_TARGET_COLUMNS:
            text = display_value(plan_row.values[index])
            if index in (0, 1, 2):
                if not text:
                    text = carry.get(target, "")
                carry[target] = text
            source[target] = text
        source["过程编号（工序号）"] = "1"
        rows.append(source)
    return rows


def _increment_process_no(value: str) -> str:
    text = value.strip()
    try:
        return str(int(float(text)) + 1)
    except ValueError:
        return text


def ensure_grease_sources(
    sources: list[dict[str, str]],
) -> tuple[list[dict[str, str]], bool]:
    """控制计划缺少涂润滑脂工序时，将其作为序号 1 插入并后移其余工序。"""
    if any(is_grease_process(source.get("过程名称", "")) for source in sources):
        return sources, False
    grease_rows = grease_source_rows()
    for source in sources:
        source["过程编号（工序号）"] = _increment_process_no(
            source.get("过程编号（工序号）", "")
        )
    return grease_rows + sources, True


def write_workbook(rows: list[dict[str, str]], output_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "质量策划"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D9E1F2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    worksheet.append(list(KZJH_COLUMNS))
    for col_index, _ in enumerate(KZJH_COLUMNS, start=1):
        cell = worksheet.cell(1, col_index)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    worksheet.row_dimensions[1].height = 34

    for row in rows:
        worksheet.append([row.get(column, "") for column in KZJH_COLUMNS])

    widths = {
        "工序号": 12,
        "特性类型": 12,
        "特征描述（产品特性+范围公差）": 46,
        "特殊特性": 10,
        "判定类型": 14,
        "实验室检测": 12,
        "自动采集": 10,
        "目标值": 12,
        "上限": 10,
        "下限": 10,
        "单位": 9,
        "检验方式": 22,
        "频次1（时间/）": 14,
        "频次2（数量/件）": 15,
        "抽样数": 10,
        "评价测量技术": 20,
        "责任人": 10,
        "数据保存方法": 13,
        "检查单位": 10,
        "检验顺序": 10,
        "对应物料": 12,
        "备注": 16,
    }
    for col_index, column in enumerate(KZJH_COLUMNS, start=1):
        worksheet.column_dimensions[worksheet.cell(1, col_index).column_letter].width = (
            widths.get(column, 14)
        )

    for row_index in range(2, worksheet.max_row + 1):
        for col_index in range(1, len(KZJH_COLUMNS) + 1):
            cell = worksheet.cell(row_index, col_index)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

    worksheet.freeze_panes = "A2"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()
    workbook.save(output_path)


def generate(
    imported: ControlPlanImport,
    mapping: dict[str, int | None],
    user: User,
    db: Database,
) -> GeneratedResult:
    missing = [target for target in TARGET_FIELDS if mapping.get(target) is None]
    sources, removed_source = build_source_rows(imported, mapping)
    if not sources:
        raise CappGenerateError("未读取到可转换的数据行，请检查字段映射。")
    if user.role.strip() == "装配":
        sources, grease_added = ensure_grease_sources(sources)
    else:
        grease_added = False

    quality_rows = build_quality_rows(sources)
    if not quality_rows:
        raise CappGenerateError(
            "转换后没有有效记录（检验方式为空的行会被过滤），请检查样本数量与样本频率字段。"
        )

    product_code = extract_product_code(imported.file_path)
    output_path = OUTPUT_DIR / f"{Path(imported.file_path).stem}-CAPP质量策划.xlsx"
    write_workbook(quality_rows, output_path)

    created_at = datetime.now().isoformat(timespec="seconds")
    with db.transaction() as conn:
        conn.execute(
            """
            INSERT INTO "CAPP质量策划记录"
                (product_code, source_file, output_path, employee_no,
                 employee_name, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                product_code,
                imported.file_path,
                str(output_path),
                user.employee_no,
                user.name,
                created_at,
            ),
        )

    warnings = list(imported.warnings)
    if grease_added:
        warnings.append("已自动补充“涂润滑脂”工序（序号 1）")
    if missing:
        warnings.append(
            f"未映射字段 {len(missing)} 个：{'、'.join(missing)}"
        )
    if removed_source:
        warnings.append(f"源行级去重删除 {removed_source} 行")
    return GeneratedResult(
        product_code=product_code,
        output_path=str(output_path),
        row_count=len(quality_rows),
        source_count=len(sources),
        removed_source=removed_source,
        warnings=warnings,
    )


def list_history(db: Database, user: User, limit: int = 100) -> list[dict]:
    rows = db.fetchall(
        """
        SELECT id, product_code, source_file, output_path, created_at
        FROM "CAPP质量策划记录"
        WHERE employee_no = ?
        ORDER BY created_at DESC, id DESC
        LIMIT ?
        """,
        (user.employee_no, limit),
    )
    return [dict(row) for row in rows]
