"""生产计划 Excel 定位、日期列识别与 5 天数据提取。"""

from __future__ import annotations

import os
from datetime import date, timedelta
from pathlib import Path

from openpyxl.utils import get_column_letter

from app.config import DATE_OFFSETS, MODES, PLAN_FILE_KEYWORD
from app.models import PlanLoadResult, PlanRow, ProjectLeadRecord
from app.utils.debug_log import detail, error, info, warn
from app.utils.excel_utils import (
    EXCEL_SUFFIXES,
    cell_text,
    load_readonly_workbook,
    merged_value,
    normalize_code,
    parse_date_value,
)


class PlanDataError(Exception):
    pass


class ProductionPlanService:
    """按业务规则读取 301 生产计划，并分组为 装配/焊接 × 5 天。"""

    def load(
        self,
        plan_root: str | Path,
        today: date,
        project_map: dict[str, ProjectLeadRecord],
    ) -> PlanLoadResult:
        root = Path(plan_root)
        info("[PLAN-1 开始加载]", f"root={root} today={today.isoformat()}")
        if not root.exists() or not root.is_dir():
            error("[PLAN-1 开始加载]", f"目录不存在或不是文件夹：{root}")
            raise PlanDataError(f"生产计划目录不存在：{root}")

        file_path, file_notices = self._resolve_plan_file(root, today)
        info("[PLAN-1 文件定位结果]", f"file={file_path} notices={file_notices}")
        try:
            workbook = load_readonly_workbook(file_path)
        except Exception as exc:  # noqa: BLE001
            error("[PLAN-2 打开文件]", f"{type(exc).__name__}: {exc}")
            raise PlanDataError(f"无法打开生产计划文件：{exc}") from exc

        result = PlanLoadResult(file_path=str(file_path))
        result.notices.extend(file_notices)
        try:
            info("[PLAN-2 打开文件]", f"file={file_path} sheets={workbook.sheetnames}")
            worksheet, used_date, sheet_notice = self._select_sheet(workbook, today)
            result.sheet_name = worksheet.title
            result.used_sheet_date = used_date
            info(
                "[PLAN-2 工作表选择结果]",
                f"sheet={worksheet.title!r} used_sheet_date={used_date.isoformat()} "
                f"notice={sheet_notice or '-'}",
            )
            if sheet_notice:
                result.notices.append(sheet_notice)

            target_dates = [today + timedelta(days=offset) for offset in DATE_OFFSETS]
            info(
                "[PLAN-3 目标日期]",
                f"{[d.isoformat() for d in target_dates]}",
            )
            date_blocks = self._find_date_blocks(worksheet, target_dates, today)
            result.missing_dates = [d for d in target_dates if d not in date_blocks]

            rows_by_mode_date: dict[str, dict[date, list[PlanRow]]] = {
                mode: {day: [] for day in target_dates} for mode in MODES
            }
            if not date_blocks:
                info(
                    "[PLAN-3 日期列结果]",
                    "未识别到任何目标日期列，输出表头区域原始内容",
                )
                self._log_header_snapshot(worksheet, today)
                result.error = (
                    f"工作表“{worksheet.title}”中未找到真实当前日期 "
                    f"{today.isoformat()} 附近五天的日期列。"
                )
                error("[PLAN-3 日期列结果]", result.error)
            else:
                merged_ranges = list(worksheet.merged_cells.ranges)
                info("[PLAN-4 开始提取]", f"merged_ranges={len(merged_ranges)}")
                self._extract_rows(
                    worksheet,
                    merged_ranges,
                    date_blocks,
                    target_dates,
                    project_map,
                    rows_by_mode_date,
                )
                if result.missing_dates:
                    missing_text = "、".join(
                        d.strftime("%m月%d日") for d in result.missing_dates
                    )
                    result.notices.append(
                        f"工作表“{worksheet.title}”缺少日期列：{missing_text}，对应日期显示为空。"
                    )
                if today not in date_blocks:
                    result.error = (
                        f"工作表“{worksheet.title}”中未找到真实当前日期 "
                        f"{today.isoformat()} 对应的列，请检查计划表表头。"
                    )
                    error("[PLAN-3 日期列结果]", result.error)

            counts = {
                mode: {
                    day.isoformat(): len(rows)
                    for day, rows in day_rows.items()
                }
                for mode, day_rows in rows_by_mode_date.items()
            }
            total = sum(
                len(rows)
                for day_rows in rows_by_mode_date.values()
                for rows in day_rows.values()
            )
            info("[PLAN-5 提取结果]", f"total={total} counts={counts}")
            info(
                "[PLAN-5 结果状态]",
                f"notices={result.notices} errors={[result.error] if result.error else []}",
            )
            result.rows_by_mode_date = rows_by_mode_date
            return result
        finally:
            workbook.close()

    def _resolve_plan_file(self, root: Path, today: date) -> tuple[Path, list[str]]:
        notices: list[str] = []
        year_dir = self._find_directory(root, str(today.year))
        if year_dir is None:
            year_dir = root
            notices.append(f"未找到包含“{today.year}年”的文件夹，已在根目录检索计划文件。")
            warn("[PLAN-1 年份目录]", f"未找到包含 {today.year} 的文件夹，回退根目录 {root}")
        else:
            info("[PLAN-1 年份目录]", f"命中 {year_dir}")

        if today.month == 12:
            next_year, next_month = today.year + 1, 1
        else:
            next_year, next_month = today.year, today.month + 1
        info(
            "[PLAN-1 月份目录]",
            f"优先次月={next_year}年{next_month}月，回退当月={today.year}年{today.month}月",
        )

        next_search_root = year_dir
        if next_year != today.year:
            next_search_root = self._find_directory(root, str(next_year)) or year_dir

        file_path: Path | None = None
        next_month_dir = self._find_directory(next_search_root, f"{next_month}月")
        if next_month_dir is not None:
            info("[PLAN-1 月份目录]", f"次月目录={next_month_dir}")
            file_path = self._find_301_file(next_month_dir)
        else:
            info("[PLAN-1 月份目录]", f"未找到次月目录“{next_month}月”")

        if file_path is None:
            current_month_dir = self._find_directory(year_dir, f"{today.month}月")
            if current_month_dir is not None:
                info("[PLAN-1 月份目录]", f"当月目录={current_month_dir}")
                file_path = self._find_301_file(current_month_dir)
            else:
                info("[PLAN-1 月份目录]", f"未找到当月目录“{today.month}月”")

        if file_path is None:
            warn("[PLAN-1 月份目录]", "当月/次月均无可用 301 文件，尝试年度目录")
            file_path = self._find_301_file(year_dir)
            if file_path is not None:
                notices.append("未在当月/次月目录中定位到 301 文件，已使用年度目录下最新的 301 文件。")

        if file_path is None:
            error("[PLAN-1 文件定位结果]", f"在 {root} 下未找到文件名包含 {PLAN_FILE_KEYWORD} 的 xlsx/xlsm")
            raise PlanDataError(
                f"在“{root}”下未找到文件名包含“{PLAN_FILE_KEYWORD}”的 Excel 计划文件。"
            )
        info("[PLAN-1 文件定位结果]", f"选中={file_path}")
        return file_path, notices

    def _find_directory(self, root: Path, token: str, max_depth: int = 3) -> Path | None:
        if not root.exists():
            detail("[PLAN-1 目录检索]", f"root={root} 不存在，token={token}")
            return None
        token = token.lower()
        candidates: list[tuple[int, str, Path]] = []
        root_depth = len(root.parts)
        for dirpath, dirnames, _ in os.walk(root):
            current = Path(dirpath)
            depth = len(current.parts) - root_depth
            if depth >= max_depth:
                dirnames[:] = []
            for name in list(dirnames):
                if name.startswith(("~$", ".", "__")):
                    dirnames.remove(name)
                    continue
                if token in name.lower():
                    candidates.append((depth, name, current / name))
        if not candidates:
            detail("[PLAN-1 目录检索]", f"root={root} token={token!r} 未命中")
            return None
        candidates.sort(key=lambda item: (item[0], item[1]))
        detail(
            "[PLAN-1 目录检索]",
            f"root={root} token={token!r} 命中={[str(item[2]) for item in candidates[:5]]}",
        )
        return candidates[0][2]

    def _find_301_file(self, folder: Path) -> Path | None:
        if not folder.exists():
            detail("[PLAN-1 文件检索]", f"目录不存在：{folder}")
            return None
        candidates: list[tuple[float, Path]] = []
        seen_files: list[str] = []
        unsupported: list[str] = []
        root_depth = len(folder.parts)
        for dirpath, dirnames, filenames in os.walk(folder):
            current = Path(dirpath)
            depth = len(current.parts) - root_depth
            if depth >= 2:
                dirnames[:] = []
            for filename in filenames:
                path = current / filename
                if filename.startswith("~$"):
                    continue
                if PLAN_FILE_KEYWORD not in path.stem:
                    continue
                suffix = path.suffix.lower()
                seen_files.append(f"{filename}({suffix})")
                if suffix not in EXCEL_SUFFIXES:
                    unsupported.append(f"{filename}({suffix})")
                    continue
                try:
                    mtime = path.stat().st_mtime
                except OSError:
                    mtime = 0.0
                candidates.append((mtime, path))
        detail(
            "[PLAN-1 文件检索]",
            f"folder={folder} 名称含301的文件={seen_files or '无'}",
        )
        if unsupported:
            warn(
                "[PLAN-1 文件检索]",
                f"以下文件不是 .xlsx/.xlsm，无法读取：{unsupported}",
            )
        if not candidates:
            detail("[PLAN-1 文件检索]", f"folder={folder} 无可用 301 文件")
            return None
        candidates.sort(key=lambda item: item[0], reverse=True)
        chosen = candidates[0][1]
        detail(
            "[PLAN-1 文件检索]",
            f"folder={folder} 候选={[(p.name, round(m, 1)) for m, p in candidates]} 选中={chosen.name}",
        )
        return chosen

    def _select_sheet(self, workbook, today: date):
        sheet_names = list(workbook.sheetnames)
        info("[PLAN-2 工作表选择]", f"可用工作表={sheet_names}")
        for offset in (0, -1, -2):
            target = today + timedelta(days=offset)
            token = target.strftime("%Y%m%d")
            detail(
                "[PLAN-2 工作表选择]",
                f"尝试 offset={offset} date={target.isoformat()} sheet名={token}",
            )
            for name in sheet_names:
                if name.strip() == token:
                    info("[PLAN-2 工作表选择]", f"精确匹配={name!r} offset={offset}")
                    notice = ""
                    if offset != 0:
                        notice = f"未找到今日工作表，已回退使用工作表“{name}”。"
                    return workbook[name], target, notice
            for name in sheet_names:
                if token in name:
                    info("[PLAN-2 工作表选择]", f"包含匹配={name!r} token={token} offset={offset}")
                    notice = (
                        f"未找到今日工作表，已回退使用工作表“{name}”（名称包含 {token}）。"
                    )
                    return workbook[name], target, notice
        tried = "、".join(
            (today + timedelta(days=offset)).strftime("%Y%m%d") for offset in (0, -1, -2)
        )
        error("[PLAN-2 工作表选择]", f"未找到 {tried}，实际={sheet_names}")
        raise PlanDataError(f"工作簿中未找到工作表：{tried}。")

    def _find_date_blocks(
        self, worksheet, target_dates: list[date], reference: date
    ) -> dict[date, dict]:
        merged_ranges = list(worksheet.merged_cells.ranges)
        blocks: dict[date, dict] = {}
        scores: dict[date, tuple] = {}
        parsed_hits = 0

        max_row = min(max(worksheet.max_row, 1), 20)
        max_col = min(max(worksheet.max_column, 1), 1000)
        info(
            "[PLAN-3 日期列扫描]",
            f"扫描范围 rows=1..{max_row} cols=1..{max_col} "
            f"merged_ranges={len(merged_ranges)} target={[d.isoformat() for d in target_dates]}",
        )
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = worksheet.cell(row, col)
                value = cell.value
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    # 数量等纯数字不应被当作 Excel 日期序列号
                    continue
                parsed = parse_date_value(value, reference)
                if parsed is None:
                    continue
                parsed_hits += 1
                detail(
                    "[PLAN-3 日期单元格]",
                    f"{get_column_letter(col)}{row} raw={value!r} -> {parsed.isoformat()} "
                    f"in_target={parsed in target_dates}",
                )
                if parsed not in target_dates:
                    continue

                merged = None
                for merged_range in merged_ranges:
                    if (
                        merged_range.min_row <= row <= merged_range.max_row
                        and merged_range.min_col <= col <= merged_range.max_col
                    ):
                        merged = merged_range
                        break

                if merged is not None:
                    if row != merged.min_row or col != merged.min_col:
                        continue
                    start_col = merged.min_col
                    end_col = merged.max_col
                    header_row = merged.min_row
                    data_start = merged.max_row + 1
                else:
                    start_col = end_col = col
                    header_row = row
                    data_start = row + 1

                width = end_col - start_col + 1
                reasonable_width = 2 <= width <= 10
                score = (
                    1 if reasonable_width else 0,
                    -header_row,
                    -abs(width - 4),
                    -start_col,
                )
                detail(
                    "[PLAN-3 日期列候选]",
                    f"date={parsed.isoformat()} cell={get_column_letter(col)}{row} "
                    f"merged={str(merged) if merged else '-'} range="
                    f"{get_column_letter(start_col)}{header_row}:{get_column_letter(end_col)}{header_row} "
                    f"data_start={data_start} width={width} score={score}",
                )
                if parsed not in scores or score > scores[parsed]:
                    scores[parsed] = score
                    blocks[parsed] = {
                        "start_col": start_col,
                        "end_col": end_col,
                        "header_row": header_row,
                        "data_start_row": data_start,
                    }
                    detail("[PLAN-3 日期列候选]", f"日期 {parsed.isoformat()} 采用该候选")

        if parsed_hits == 0:
            warn("[PLAN-3 日期列扫描]", "扫描范围内没有解析出任何日期值")
        for day in target_dates:
            block = blocks.get(day)
            if block is None:
                warn("[PLAN-3 日期列结果]", f"缺少日期 {day.isoformat()} 的列")
                continue
            info(
                "[PLAN-3 日期列结果]",
                f"date={day.isoformat()} "
                f"col={get_column_letter(block['start_col'])}..{get_column_letter(block['end_col'])} "
                f"header_row={block['header_row']} data_start_row={block['data_start_row']}",
            )
        return blocks

    def _log_header_snapshot(self, worksheet, reference: date) -> None:
        """日期列识别失败时，输出表头区域的原始单元格内容，便于核对格式。"""
        max_row = min(max(worksheet.max_row, 1), 15)
        max_col = min(max(worksheet.max_column, 1), 40)
        info(
            "[PLAN-3 表头快照]",
            f"rows=1..{max_row} cols=1..{max_col}",
        )
        for row in range(1, max_row + 1):
            values: list[str] = []
            for col in range(1, max_col + 1):
                value = worksheet.cell(row, col).value
                if value is None or value == "":
                    continue
                parsed = parse_date_value(value, reference)
                text = cell_text(value).replace("\n", "\\n")
                suffix = f"→{parsed.isoformat()}" if parsed else ""
                values.append(f"{get_column_letter(col)}{row}={text!r}{suffix}")
            if values:
                detail("[PLAN-3 表头快照]", " | ".join(values))
        ranges = list(worksheet.merged_cells.ranges)
        if ranges:
            detail(
                "[PLAN-3 合并区域]",
                f"前 30 个={[str(item) for item in ranges[:30]]}",
            )

    def _extract_rows(
        self,
        worksheet,
        merged_ranges: list,
        date_blocks: dict[date, dict],
        target_dates: list[date],
        project_map: dict[str, ProjectLeadRecord],
        rows_by_mode_date: dict[str, dict[date, list[PlanRow]]],
    ) -> None:
        seen: set[tuple] = set()
        max_col = worksheet.max_column
        max_row = worksheet.max_row
        qualified_codes = sorted(
            (
                candidate
                for candidate, record in project_map.items()
                if record.mass in (0, 1, 2, 3, 4)
            ),
            key=len,
            reverse=True,
        )
        info(
            "[PLAN-4 提取参数]",
            f"max_row={max_row} max_col={max_col} "
            f"project_map={len(project_map)} qualified_projects={len(qualified_codes)}",
        )

        for day in target_dates:
            block = date_blocks.get(day)
            if block is None:
                warn("[PLAN-4 日期提取]", f"date={day.isoformat()} 无日期列，跳过")
                continue
            start_col = block["start_col"]
            data_start = block["data_start_row"]
            empty_run = 0
            stats = {
                "scanned": 0,
                "empty_code": 0,
                "code_sample": [],
                "unmatched": 0,
                "unmatched_sample": [],
                "contained": 0,
                "contained_sample": [],
                "mass_filtered": 0,
                "mass_sample": [],
                "mode_filtered": 0,
                "mode_sample": [],
                "duplicated": 0,
                "added": 0,
            }
            info(
                "[PLAN-4 日期提取]",
                f"date={day.isoformat()} "
                f"code_col={get_column_letter(start_col)} "
                f"range={get_column_letter(block['end_col'])} "
                f"quantity_col={get_column_letter(start_col + 1)} "
                f"remark_col={get_column_letter(start_col + 3)} "
                f"data_start_row={data_start}",
            )

            for row in range(data_start, max_row + 1):
                stats["scanned"] += 1
                code_value = merged_value(worksheet, merged_ranges, row, start_col)
                code = normalize_code(code_value)
                if not code:
                    empty_run += 1
                    stats["empty_code"] += 1
                    if empty_run >= 40:
                        detail(
                            "[PLAN-4 日期提取]",
                            f"date={day.isoformat()} 第 {row} 行起连续 {empty_run} 行空代号，停止扫描",
                        )
                        break
                    continue
                empty_run = 0
                if len(stats["code_sample"]) < 30:
                    stats["code_sample"].append(
                        f"{row}:{cell_text(code_value)!r}"
                    )

                project = project_map.get(code)
                match_note = "精确"
                if project is None:
                    # 计划表代号可能带后缀，只要包含清单中的产品代号即可匹配（最长优先）
                    for candidate in qualified_codes:
                        if candidate in code:
                            project = project_map[candidate]
                            match_note = f"包含:{candidate}"
                            stats["contained"] += 1
                            if len(stats["contained_sample"]) < 15:
                                stats["contained_sample"].append(
                                    f"{row}:{code}⊇{candidate}"
                                )
                            detail(
                                "[PLAN-4 包含匹配]",
                                f"date={day.isoformat()} row={row} code={code!r} "
                                f"包含清单代号={candidate!r} mass={project.mass}",
                            )
                            break
                if project is None:
                    stats["unmatched"] += 1
                    if len(stats["unmatched_sample"]) < 15:
                        stats["unmatched_sample"].append(f"{row}:{code}")
                    detail(
                        "[PLAN-4 行匹配]",
                        f"date={day.isoformat()} row={row} code={code!r} -> 未在项目清单中找到",
                    )
                    continue
                if project.mass not in (0, 1, 2, 3, 4):
                    stats["mass_filtered"] += 1
                    if len(stats["mass_sample"]) < 15:
                        stats["mass_sample"].append(f"{row}:{code}=mass{project.mass}")
                    detail(
                        "[PLAN-4 行匹配]",
                        f"date={day.isoformat()} row={row} code={code!r} "
                        f"mass={project.mass} -> mass 不在 0-4，过滤",
                    )
                    continue

                line_text = cell_text(merged_value(worksheet, merged_ranges, row, 1))
                mode = self._classify_mode(line_text)
                if mode is None:
                    stats["mode_filtered"] += 1
                    if len(stats["mode_sample"]) < 15:
                        stats["mode_sample"].append(f"{row}:{line_text!r}")
                    detail(
                        "[PLAN-4 行匹配]",
                        f"date={day.isoformat()} row={row} A列={line_text!r} "
                        f"-> 不含 C/B，无法判定工序",
                    )
                    continue

                quantity = cell_text(
                    merged_value(worksheet, merged_ranges, row, start_col + 1)
                ) if start_col + 1 <= max_col else ""
                remark = cell_text(
                    merged_value(worksheet, merged_ranges, row, start_col + 3)
                ) if start_col + 3 <= max_col else ""
                owner = project.assembly if mode == "装配" else project.welding

                row_key = (day, mode, line_text, code, quantity, remark)
                if row_key in seen:
                    stats["duplicated"] += 1
                    detail(
                        "[PLAN-4 行匹配]",
                        f"date={day.isoformat()} row={row} code={code!r} -> 重复行，跳过",
                    )
                    continue
                seen.add(row_key)

                rows_by_mode_date[mode][day].append(
                    PlanRow(
                        production_line=line_text,
                        host=project.host,
                        product_code=cell_text(code_value),
                        quantity=quantity,
                        owner=owner,
                        remark=remark,
                    )
                )
                stats["added"] += 1
                detail(
                    "[PLAN-4 行入表]",
                    f"date={day.isoformat()} row={row} 工序={mode} "
                    f"A列={line_text!r} code={code!r} 匹配={match_note} mass={project.mass} "
                    f"数量={quantity!r} 备注={remark!r} 负责人={owner!r} 主机={project.host!r}",
                )

            info(
                "[PLAN-4 日期提取统计]",
                f"date={day.isoformat()} 扫描={stats['scanned']} 空代号={stats['empty_code']} "
                f"未匹配={stats['unmatched']} 包含匹配={stats['contained']} "
                f"mass过滤={stats['mass_filtered']} "
                f"工序过滤={stats['mode_filtered']} 重复={stats['duplicated']} "
                f"写入={stats['added']} "
                f"装配={len(rows_by_mode_date['装配'][day])} "
                f"焊接={len(rows_by_mode_date['焊接'][day])}",
            )
            if stats["empty_code"] == stats["scanned"]:
                warn(
                    "[PLAN-4 日期提取]",
                    f"date={day.isoformat()} 所有行代号列 {get_column_letter(start_col)} 均为空，"
                    f"请确认该日期列的起始列是否正确",
                )
            if stats["unmatched"]:
                warn(
                    "[PLAN-4 代号未匹配样例]",
                    f"date={day.isoformat()} 样例={stats['unmatched_sample']}",
                )
            if stats["mass_filtered"]:
                warn(
                    "[PLAN-4 mass过滤样例]",
                    f"date={day.isoformat()} 样例={stats['mass_sample']}",
                )
            if stats["mode_filtered"]:
                warn(
                    "[PLAN-4 工序过滤样例]",
                    f"date={day.isoformat()} A列样例={stats['mode_sample']}",
                )
            if stats["contained"]:
                info(
                    "[PLAN-4 包含匹配样例]",
                    f"date={day.isoformat()} 样例={stats['contained_sample']}",
                )
            detail(
                "[PLAN-4 代号样例]",
                f"date={day.isoformat()} 代号列样例={stats['code_sample']}",
            )

    @staticmethod
    def _classify_mode(line_text: str) -> str | None:
        normalized = line_text.upper()
        if "C" in normalized:
            return "装配"
        if "B" in normalized:
            return "焊接"
        return None
