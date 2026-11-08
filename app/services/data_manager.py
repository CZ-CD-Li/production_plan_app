"""统一编排项目分布清单与生产计划的加载。"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

from app.config import DATE_OFFSETS, MODES
from app.database import Database
from app.models import LoadOutcome, PlanLoadResult, PlanRow, ProjectLeadRecord, User
from app.services.plan_share import PlanShareService, fetch_latest_remote
from app.services.production_plan_service import (
    PlanDataError,
    ProductionPlanService,
)
from app.services.project_lead_service import ProjectLeadError, ProjectLeadService
from app.services.settings_service import SettingsService
from app.utils.debug_log import detail, error, info, warn
from app.utils.excel_utils import normalize_code


class DataManager:
    def __init__(self, db: Database, user: User | None = None) -> None:
        self.db = db
        self.user = user
        self.settings_service = SettingsService(db)
        self.project_service = ProjectLeadService()
        self.plan_service = ProductionPlanService()
        self.share_service = PlanShareService()
        self.today = date.today()
        self.project_map: dict[str, ProjectLeadRecord] = {}
        self.rows_by_mode_date: dict[str, dict[date, list]] = {
            mode: {} for mode in MODES
        }
        self.plan_result = PlanLoadResult()
        self.outcome = LoadOutcome()
        self.plan_source = ""  # 数据来源说明，如 "远程 192.168.1.101"

    def reload(self) -> LoadOutcome:
        self.today = date.today()
        employee_no = self.user.employee_no if self.user else None
        settings = self.settings_service.get_all(employee_no)
        errors: list[str] = []
        notices: list[str] = []

        info("[DATA-1 开始同步]", f"today={self.today.isoformat()} user={employee_no}")
        info(
            "[DATA-1 数据源配置]",
            f"项目清单={settings.get('project_lead_path', '')!r} "
            f"计划目录={settings.get('production_plan_dir', '')!r}",
        )

        self.rows_by_mode_date = {
            mode: {self.today + timedelta(days=offset): [] for offset in DATE_OFFSETS}
            for mode in MODES
        }
        self.plan_result = PlanLoadResult()
        self.plan_source = ""

        # 先尝试从其他终端的共享接口直接取用最新计划，成功则跳过本机 Excel 计算
        remote = self._try_remote_plan(settings, notices)
        if remote:
            self.outcome = self._build_outcome(errors, notices)
            return self.outcome

        project_count = self._sync_project_lead(
            settings.get("project_lead_path", "").strip(), errors, notices
        )
        self.project_map = self._load_project_map()

        qualified = sum(
            1 for record in self.project_map.values() if 0 <= record.mass <= 4
        )
        info(
            "[DATA-2 项目清单结果]",
            f"数据库行数={project_count} 有效代号={len(self.project_map)} "
            f"mass0-4={qualified}",
        )
        if self.project_map:
            sample = [
                f"{code}=mass{record.mass}"
                for code, record in list(self.project_map.items())[:20]
            ]
            detail("[DATA-2 项目代号样例]", f"{sample}")
        if not self.project_map:
            warn("[DATA-2 项目清单结果]", "有效代号为 0，生产计划无法匹配，将跳过计划加载")
            notices.append("项目分布清单暂无有效数据，生产计划无法匹配产品。")
        else:
            self._sync_plan(
                settings.get("production_plan_dir", "").strip(),
                errors,
                notices,
            )

        if not self.plan_result.rows_by_mode_date:
            for mode in MODES:
                self.rows_by_mode_date[mode] = {
                    self.today + timedelta(days=offset): [] for offset in DATE_OFFSETS
                }
        else:
            for mode in MODES:
                for day, rows in self.plan_result.rows_by_mode_date.get(mode, {}).items():
                    self.rows_by_mode_date[mode][day] = rows
                for offset in DATE_OFFSETS:
                    self.rows_by_mode_date[mode].setdefault(
                        self.today + timedelta(days=offset), []
                    )

        plan_count = sum(
            len(rows)
            for mode_rows in self.rows_by_mode_date.values()
            for rows in mode_rows.values()
        )
        for mode in MODES:
            counts = {
                day.isoformat(): len(rows)
                for day, rows in sorted(self.rows_by_mode_date[mode].items())
            }
            info("[DATA-3 最终分组]", f"{mode} 各日期行数={counts}")
        info(
            "[DATA-3 同步结果]",
            f"计划总行数={plan_count} plan_file={self.plan_result.file_path!r} "
            f"sheet={self.plan_result.sheet_name!r} "
            f"errors={errors} notices={notices}",
        )
        if plan_count == 0:
            warn(
                "[DATA-3 同步结果]",
                "最终计划行数为 0，请根据上方 PLAN / DATA 日志定位失败环节",
            )
        self._publish_snapshot()
        self.outcome = self._build_outcome(errors, notices)
        return self.outcome

    def _build_outcome(self, errors: list[str], notices: list[str]) -> LoadOutcome:
        plan_count = sum(
            len(rows)
            for mode_rows in self.rows_by_mode_date.values()
            for rows in mode_rows.values()
        )
        project_count = self._project_row_count()
        return LoadOutcome(
            errors=errors,
            notices=notices,
            project_count=project_count,
            matched_project_count=sum(
                1 for record in self.project_map.values() if 0 <= record.mass <= 4
            ),
            plan_count=plan_count,
            plan_file=Path(self.plan_result.file_path).name
            if self.plan_result.file_path
            else "",
            sheet_name=self.plan_result.sheet_name,
            synced_at=datetime.now(),
        )

    def _try_remote_plan(self, settings: dict[str, str], notices: list[str]) -> bool:
        """从设置的终端 IP 中获取最新的计划快照，成功返回 True。"""
        ips = [settings.get("ip1", ""), settings.get("ip2", ""), settings.get("ip3", "")]
        if not any(ip.strip() for ip in ips):
            return False
        try:
            found = fetch_latest_remote(ips, self.today)
        except Exception as exc:  # noqa: BLE001
            warn("[DATA-1 计划共享]", f"远程查询异常：{type(exc).__name__}: {exc}")
            return False
        if not found:
            info("[DATA-1 计划共享]", "各终端暂无可用计划数据，转为本机计算")
            return False
        payload, source_ip = found
        rows_payload = payload.get("rows", {})
        adopted: dict[str, dict[date, list]] = {}
        for mode in MODES:
            mode_rows = rows_payload.get(mode, {})
            if not isinstance(mode_rows, dict):
                continue
            day_map: dict[date, list] = {}
            for day_text, items in mode_rows.items():
                try:
                    day = date.fromisoformat(str(day_text))
                except ValueError:
                    continue
                day_map[day] = [
                    PlanRow(
                        production_line=str(item.get("production_line", "")),
                        host=str(item.get("host", "")),
                        product_code=str(item.get("product_code", "")),
                        quantity=str(item.get("quantity", "")),
                        owner=str(item.get("owner", "")),
                        remark=str(item.get("remark", "")),
                    )
                    for item in items
                    if isinstance(item, dict)
                ]
            adopted[mode] = day_map
        for mode in MODES:
            for day, rows in adopted.get(mode, {}).items():
                self.rows_by_mode_date[mode][day] = rows
        self.plan_result = PlanLoadResult(
            rows_by_mode_date=adopted,
            file_path=str(payload.get("plan_file", "")),
            sheet_name=str(payload.get("sheet_name", "")),
        )
        self.project_map = self._load_project_map()
        self.plan_source = f"远程 {source_ip}"
        # 取到的快照继续在本机发布，让数据在网内流转
        self.share_service.publish(payload)
        synced_at = str(payload.get("synced_at", ""))
        notices.append(f"已直接使用 {source_ip} 计算的计划数据（{synced_at}），本机未重复计算。")
        info(
            "[DATA-1 计划共享]",
            f"采用 {source_ip} 的快照 synced_at={synced_at} "
            f"行数={self._build_outcome([], []).plan_count}",
        )
        return True

    def _publish_snapshot(self) -> None:
        """把本机计算结果发布到共享接口，供其他终端直接取用。"""
        plan_count = sum(
            len(rows)
            for mode_rows in self.rows_by_mode_date.values()
            for rows in mode_rows.values()
        )
        if plan_count == 0:
            return
        rows_payload: dict[str, dict[str, list]] = {}
        for mode, day_map in self.rows_by_mode_date.items():
            rows_payload[mode] = {
                day.isoformat(): [
                    {
                        "production_line": row.production_line,
                        "host": row.host,
                        "product_code": row.product_code,
                        "quantity": row.quantity,
                        "owner": row.owner,
                        "remark": row.remark,
                    }
                    for row in rows
                ]
                for day, rows in day_map.items()
            }
        self.share_service.publish(
            {
                "today": self.today.isoformat(),
                "synced_at": datetime.now().isoformat(timespec="seconds"),
                "plan_file": self.plan_result.file_path,
                "sheet_name": self.plan_result.sheet_name,
                "rows": rows_payload,
            }
        )
        info("[计划共享]", f"已发布本机计划快照，共 {plan_count} 行")

    def _sync_project_lead(
        self, path_text: str, errors: list[str], notices: list[str]
    ) -> int:
        if not path_text:
            error("[DATA-2 项目清单]", "未配置项目分布清单文件")
            errors.append("未配置项目分布清单文件，请在系统设置中选择 Excel 文件。")
            return self._project_row_count()
        if not Path(path_text).exists():
            error("[DATA-2 项目清单]", f"文件不存在：{path_text}")
            errors.append(f"项目分布清单文件不存在：{path_text}")
            return self._project_row_count()

        info("[DATA-2 项目清单]", f"开始导入文件={path_text}")
        try:
            result = self.project_service.import_from_file(path_text, self.db)
            self.settings_service.set_many(
                {
                    "project_lead_synced_at": datetime.now().isoformat(
                        timespec="seconds"
                    )
                }
            )
            notices.extend(result.warnings)
            info(
                "[DATA-2 项目清单]",
                f"导入完成 total={result.total} warnings={result.warnings}",
            )
            return result.total
        except ProjectLeadError as exc:
            error("[DATA-2 项目清单]", f"读取失败：{exc}")
            errors.append(f"项目分布清单读取失败：{exc}")
        except Exception as exc:  # noqa: BLE001
            error("[DATA-2 项目清单]", f"读取异常：{type(exc).__name__}: {exc}")
            errors.append(f"项目分布清单读取异常：{exc}")
        return self._project_row_count()

    def _sync_plan(
        self, directory_text: str, errors: list[str], notices: list[str]
    ) -> None:
        if not directory_text:
            error("[DATA-3 生产计划]", "未配置生产计划目录")
            errors.append("未配置生产计划目录，请在系统设置中选择计划文件所在文件夹。")
            return
        if not Path(directory_text).exists():
            error("[DATA-3 生产计划]", f"目录不存在：{directory_text}")
            errors.append(f"生产计划目录不存在：{directory_text}")
            return

        info("[DATA-3 生产计划]", f"开始加载目录={directory_text}")
        try:
            result = self.plan_service.load(directory_text, self.today, self.project_map)
        except PlanDataError as exc:
            error("[DATA-3 生产计划]", f"读取失败：{exc}")
            errors.append(f"生产计划读取失败：{exc}")
            return
        except Exception as exc:  # noqa: BLE001
            error("[DATA-3 生产计划]", f"读取异常：{type(exc).__name__}: {exc}")
            errors.append(f"生产计划读取异常：{exc}")
            return

        self.plan_result = result
        notices.extend(result.notices)
        if result.error:
            errors.append(result.error)
        info(
            "[DATA-3 生产计划]",
            f"加载完成 file={result.file_path!r} sheet={result.sheet_name!r} "
            f"used_sheet_date={result.used_sheet_date} "
            f"missing_dates={[d.isoformat() for d in result.missing_dates]}",
        )

    def _load_project_map(self) -> dict[str, ProjectLeadRecord]:
        rows = self.db.lead_fetchall(
            """
            SELECT id, "主机", "产品代号", "生产线", "焊接", "装配", mass
            FROM "Project Lead"
            """
        )
        mapping: dict[str, ProjectLeadRecord] = {}
        for row in rows:
            code = normalize_code(row["产品代号"])
            if not code:
                continue
            record = ProjectLeadRecord(
                id=row["id"] or "",
                host=row["主机"] or "",
                product_code=row["产品代号"] or "",
                production_line=row["生产线"] or "",
                welding=row["焊接"] or "",
                assembly=row["装配"] or "",
                mass=int(row["mass"] or 0),
            )
            existing = mapping.get(code)
            if existing is None or existing.mass < record.mass:
                mapping[code] = record
        mass_counts: dict[int, int] = {}
        for record in mapping.values():
            mass_counts[record.mass] = mass_counts.get(record.mass, 0) + 1
        info(
            "[DATA-2 Project Lead 映射]",
            f"数据库行数={len(rows)} 去重代号={len(mapping)} "
            f"mass分布={dict(sorted(mass_counts.items()))}",
        )
        if not mapping:
            warn("[DATA-2 Project Lead 映射]", "数据库 Project Lead 表为空")
        return mapping

    def _project_row_count(self) -> int:
        row = self.db.lead_fetchone('SELECT COUNT(*) AS count FROM "Project Lead"')
        return int(row["count"]) if row is not None else 0

    def rows_for(self, mode: str, day: date) -> list:
        return self.rows_by_mode_date.get(mode, {}).get(day, [])

    def dates(self) -> list[date]:
        return [self.today + timedelta(days=offset) for offset in DATE_OFFSETS]
