"""“新产品项目任务分布清单”解析与 Project Lead 表维护。"""

from __future__ import annotations

import re
import sqlite3
import time
import unicodedata
from collections import Counter
from pathlib import Path

from openpyxl.utils import get_column_letter

from app.config import PROJECT_LEAD_SHEET
from app.database import Database
from app.models import ProjectLeadImportResult, ProjectLeadRecord
from app.utils.debug_log import detail, error, info, warn
from app.utils.excel_utils import (
    cell_text,
    color_key,
    load_readonly_workbook,
    normalize_code,
    normalize_header,
)


class ProjectLeadError(Exception):
    pass


HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "id": ("序号",),
    "host": ("主机", "主机厂"),
    "product_code": ("公司代号", "产品代号", "产品编号"),
    "production_line": ("生产线", "线别"),
    "welding": ("焊接",),
    "assembly": ("装配",),
}

REQUIRED_FIELDS = ("host", "product_code")
EMPTY_ROW_LIMIT = 3


def split_product_codes(value) -> list[str]:
    """把公司代号单元格拆成单个产品代号。

    支持 / 分隔、括号内的备用代号，以及中文括号说明：
    A936/B088、B115（B424）、B083（昭和）→ A936 / B088 / B115 / B424 / B083。
    """
    text = unicodedata.normalize("NFKC", cell_text(value))
    if not text:
        return []
    text = re.sub(r"[（(【\[]", "/", text)
    text = re.sub(r"[）)】\]]", "/", text)
    parts = re.split(r"[/／、,，;；\s]+", text)
    codes: list[str] = []
    for part in parts:
        code = normalize_code(part)
        if not code:
            continue
        if not re.fullmatch(r"[A-Z0-9][A-Z0-9._\-]*", code):
            continue
        codes.append(code)
    return codes


class ProjectLeadService:
    """读取项目分布清单，更新 SQLite 中的 ProjectLead 表。"""

    def import_from_file(self, path: str | Path, db: Database) -> ProjectLeadImportResult:
        file_path = Path(path)
        info("[项目清单-1 文件]", f"path={file_path}")
        if not file_path.exists():
            error("[项目清单-1 文件]", f"文件不存在：{file_path}")
            raise ProjectLeadError(f"文件不存在：{file_path}")
        if file_path.suffix.lower() not in (".xlsx", ".xlsm"):
            error("[项目清单-1 文件]", f"不支持的扩展名：{file_path.suffix}")
            raise ProjectLeadError("项目分布清单仅支持 .xlsx / .xlsm 文件。")
        try:
            size = file_path.stat().st_size
        except OSError:
            size = 0
        info("[项目清单-1 文件]", f"size={size} bytes suffix={file_path.suffix.lower()}")

        try:
            workbook = load_readonly_workbook(file_path)
        except Exception as exc:  # noqa: BLE001
            error("[项目清单-1 文件]", f"打开失败：{type(exc).__name__}: {exc}")
            raise ProjectLeadError(f"无法打开文件：{exc}") from exc

        try:
            info("[项目清单-2 工作表]", f"工作簿内工作表={workbook.sheetnames}")
            worksheet = self._find_sheet(workbook)
            header_row, mapping = self._find_header_row(worksheet)
            info(
                "[项目清单-4 表头结果]",
                f"header_row={header_row} "
                f"columns={ {field: get_column_letter(col) for field, col in mapping.items()} } "
                f"max_row={worksheet.max_row} max_column={worksheet.max_column}",
            )
            band = self._read_mass_band(worksheet)
            color_to_mass: dict[str, int] = {}
            for item in band:
                # 从左向右比较，首个同色单元格即对应的 mass
                color_to_mass.setdefault(item["color"], item["mass"])
            warnings: list[str] = []

            info(
                "[项目清单-5 颜色映射]",
                f"band_count={len(band)} mapping={color_to_mass}",
            )
            if not band:
                warnings.append("未在 N1 起识别到有效颜色块，所有产品 mass 记为 0。")
                warn("[项目清单-5 颜色映射]", "N1 起未读到有效填充色，mass 将全部为 0")
            info(
                "[项目清单-6 行读取]",
                f"从第 {header_row + 1} 行开始扫描，共 {worksheet.max_row} 行",
            )

            records: list[ProjectLeadRecord] = []
            mass_counts: Counter[int] = Counter()
            seen_codes: set[str] = set()
            duplicate_codes: list[str] = []
            empty_run = 0
            logged_rows = 0
            stop_reason = "已扫描到最后一行"
            for row in range(header_row + 1, worksheet.max_row + 1):
                code_cell = worksheet.cell(row, mapping["product_code"])
                raw_code = cell_text(code_cell.value)
                if not raw_code.strip():
                    empty_run += 1
                    if empty_run >= EMPTY_ROW_LIMIT:
                        stop_reason = (
                            f"第 {row} 行起连续 {empty_run} 行公司代号为空，停止读取"
                        )
                        break
                    continue
                empty_run = 0
                codes = split_product_codes(raw_code)
                if not codes:
                    warn(
                        "[项目清单-6 行读取]",
                        f"row={row} 公司代号={raw_code!r} 无法拆分出有效产品代号",
                    )
                    continue
                code_color = color_key(code_cell) or ""
                mass = color_to_mass.get(code_color, 0)
                host = self._field_text(worksheet, row, mapping, "host")
                production_line = self._field_text(
                    worksheet, row, mapping, "production_line"
                )
                welding = self._field_text(worksheet, row, mapping, "welding")
                assembly = self._field_text(worksheet, row, mapping, "assembly")
                record_id = self._field_text(worksheet, row, mapping, "id")
                added = 0
                row_duplicates: list[str] = []
                for code in codes:
                    if code in seen_codes:
                        duplicate_codes.append(f"{row}:{code}")
                        row_duplicates.append(code)
                        continue
                    seen_codes.add(code)
                    records.append(
                        ProjectLeadRecord(
                            id=record_id,
                            host=host,
                            product_code=code,
                            production_line=production_line,
                            welding=welding,
                            assembly=assembly,
                            mass=mass,
                        )
                    )
                    mass_counts[mass] += 1
                    added += 1
                if logged_rows < 300:
                    detail(
                        "[项目清单-6 行读取]",
                        f"row={row} id={record_id!r} 原值={raw_code!r} "
                        f"拆分={codes} fill={code_color or '-'} mass={mass} "
                        f"host={host!r} 新增={added} 重复={row_duplicates}",
                    )
                    logged_rows += 1

            if logged_rows >= 300:
                warn("[项目清单-6 行读取]", "数据行超过 300，后续行日志已省略")
            info("[项目清单-6 读取结束]", stop_reason)
            info(
                "[项目清单-6 行读取结果]",
                f"拆分后记录={len(records)} 去重删除={len(duplicate_codes)} "
                f"mass_counts={dict(mass_counts)}",
            )
            if duplicate_codes:
                info(
                    "[项目清单-6 重复代号]",
                    f"保留首次出现，删除 {len(duplicate_codes)} 个：{duplicate_codes[:30]}",
                )
            if not records:
                error("[项目清单-6 行读取结果]", "未读取到任何产品代号，请检查表头列映射")
                raise ProjectLeadError("未读取到有效产品记录，请检查表头与数据区域。")

            self._replace_rows(db, records)
            if not any(0 <= mass <= 4 for mass in mass_counts):
                warnings.append("没有产品被识别为 mass 0-4，生产计划将无法匹配到条目。")
                warn("[项目清单-6 行读取结果]", "没有任何产品 mass 在 0-4 范围")
            return ProjectLeadImportResult(
                total=len(records),
                mass_counts=dict(mass_counts),
                color_band=band,
                warnings=warnings,
            )
        finally:
            workbook.close()

    def _find_sheet(self, workbook):
        for name in workbook.sheetnames:
            if name.strip() == PROJECT_LEAD_SHEET:
                info("[项目清单-3 工作表]", f"已匹配工作表={name!r}")
                return workbook[name]
        error(
            "[项目清单-3 工作表]",
            f"未找到期望工作表={PROJECT_LEAD_SHEET!r}，实际={workbook.sheetnames}",
        )
        raise ProjectLeadError(f'缺少工作表“{PROJECT_LEAD_SHEET}”。')

    def _find_header_row(self, worksheet) -> tuple[int, dict[str, int]]:
        best_row = 0
        best_mapping: dict[str, int] = {}
        best_score = 0
        max_scan_row = min(max(worksheet.max_row, 2), 10)

        for row in range(1, max_scan_row + 1):
            mapping: dict[str, int] = {}
            for col in range(1, worksheet.max_column + 1):
                header = normalize_header(worksheet.cell(row, col).value)
                if not header:
                    continue
                for field, aliases in HEADER_ALIASES.items():
                    if field in mapping:
                        continue
                    if any(header == alias for alias in aliases):
                        mapping[field] = col
                        break
            score = len(mapping)
            required_hits = sum(
                1 for field in REQUIRED_FIELDS if field in mapping
            )
            detail(
                "[项目清单-4 表头扫描]",
                f"row={row} matched={required_hits} fields={list(mapping)}",
            )
            if required_hits >= 2 and score > best_score:
                best_row, best_mapping, best_score = row, mapping, score

        if not best_row or "product_code" not in best_mapping:
            # 按约定第 2 行为表头，尝试兜底解析
            warn("[项目清单-4 表头扫描]", "常规扫描未找到表头，回退按第 2 行解析")
            mapping = {}
            for col in range(1, worksheet.max_column + 1):
                header = normalize_header(worksheet.cell(2, col).value)
                if not header:
                    continue
                for field, aliases in HEADER_ALIASES.items():
                    if field in mapping:
                        continue
                    if any(alias in header for alias in aliases):
                        mapping[field] = col
                        break
            if "product_code" not in mapping:
                error(
                    "[项目清单-4 表头结果]",
                    f"第 2 行未找到公司代号/产品代号列，第 2 行内容="
                    f"{[cell_text(worksheet.cell(2, c).value) for c in range(1, min(worksheet.max_column, 20) + 1)]}",
                )
                raise ProjectLeadError("未找到“公司代号”列，请检查第 2 行表头。")
            best_row, best_mapping = 2, mapping

        info(
            "[项目清单-4 表头扫描结果]",
            f"header_row={best_row} fields={list(best_mapping)}",
        )
        return best_row, best_mapping

    @staticmethod
    def _field_text(worksheet, row: int, mapping: dict[str, int], field: str) -> str:
        col = mapping.get(field)
        if not col:
            return ""
        return cell_text(worksheet.cell(row, col).value)

    def _read_mass_band(self, worksheet) -> list[dict]:
        band: list[dict] = []
        last_col = max(worksheet.max_column, 14)
        stop_reason = "已扫描到最后一列"
        for col in range(14, last_col + 1):
            cell = worksheet.cell(1, col)
            fill = cell.fill
            foreground = getattr(fill, "fgColor", None)
            raw_fill = {
                "type": getattr(fill, "fill_type", None),
                "rgb": getattr(foreground, "rgb", None),
                "theme": getattr(foreground, "theme", None),
                "tint": getattr(foreground, "tint", None),
                "indexed": getattr(foreground, "indexed", None),
            }
            key = color_key(cell)
            detail(
                "[项目清单-5 颜色块扫描]",
                f"cell={get_column_letter(col)}1 value={cell_text(cell.value)!r} "
                f"fill={raw_fill} key={key or '-'}",
            )
            if key is None:
                stop_reason = f"{get_column_letter(col)}1 为白色或无填充，停止"
                break
            band.append({"column": get_column_letter(col), "color": key})
        total = len(band)
        for index, item in enumerate(band):
            item["mass"] = total - index
        info(
            "[项目清单-5 颜色块结果]",
            f"count={total} stop={stop_reason} "
            f"band={[(item['column'], item['color'], item['mass']) for item in band]}",
        )
        return band

    def _replace_rows(self, db: Database, records: list[ProjectLeadRecord]) -> None:
        info("[项目清单-7 写入数据库]", f"先清空 Project Lead，再写入 {len(records)} 行")
        payload = [
            (
                record.id,
                record.host,
                record.product_code,
                record.production_line,
                record.welding,
                record.assembly,
                record.mass,
            )
            for record in records
        ]
        last_error: Exception | None = None
        for attempt in range(1, 4):
            try:
                # 本机库单机写入，先取写锁避免与其他本机实例升级锁
                with db.lead_transaction(immediate=True) as conn:
                    conn.execute('DELETE FROM "Project Lead"')
                    conn.executemany(
                        """
                        INSERT INTO "Project Lead"
                            (id, "主机", "产品代号", "生产线", "焊接", "装配", mass)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        payload,
                    )
                last_error = None
                break
            except sqlite3.OperationalError as exc:
                last_error = exc
                if "locked" not in str(exc).lower() and "busy" not in str(exc).lower():
                    raise
                warn(
                    "[项目清单-7 写入数据库]",
                    f"其他主机正在写入，第 {attempt} 次重试：{exc}",
                )
                time.sleep(1.5 * attempt)
        if last_error is not None:
            error("[项目清单-7 写入数据库]", f"重试后仍无法写入：{last_error}")
            raise ProjectLeadError(
                "数据库正被其他主机占用，请稍后再同步项目清单。"
            ) from last_error
        info("[项目清单-7 写入数据库]", f"完成，Project Lead 现有 {len(records)} 行")
