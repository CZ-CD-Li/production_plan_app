"""跨模块共享的数据模型。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime


@dataclass(frozen=True)
class User:
    id: int
    employee_no: str
    name: str
    role: str = "生产计划员"


@dataclass(frozen=True)
class ProjectLeadRecord:
    """“新产品项目任务分布清单”中的一个产品代号。"""

    id: str
    host: str
    product_code: str
    production_line: str
    welding: str
    assembly: str
    mass: int


@dataclass(frozen=True)
class PlanRow:
    """最终展示在 5 天计划表格中的一条记录。"""

    production_line: str
    host: str
    product_code: str
    quantity: str
    owner: str
    remark: str


@dataclass
class PlanLoadResult:
    rows_by_mode_date: dict[str, dict[date, list[PlanRow]]] = field(default_factory=dict)
    file_path: str = ""
    sheet_name: str = ""
    used_sheet_date: date | None = None
    missing_dates: list[date] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    error: str = ""


@dataclass
class ProjectLeadImportResult:
    total: int
    mass_counts: dict[int, int]
    color_band: list[dict]
    warnings: list[str] = field(default_factory=list)


@dataclass
class LoadOutcome:
    errors: list[str] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    project_count: int = 0
    matched_project_count: int = 0
    plan_count: int = 0
    plan_file: str = ""
    sheet_name: str = ""
    synced_at: datetime | None = None

    @property
    def ok(self) -> bool:
        return not self.errors
