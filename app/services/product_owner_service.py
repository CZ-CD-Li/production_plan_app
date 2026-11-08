"""产品责任人查询：按产品代号、主机、生产线组合筛选 Project Lead。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.services.line_access import line_codes_in_text

FIELDS = ("产品代号", "主机", "生产线", "装配", "焊接")


@dataclass(frozen=True)
class ProductOwnerRow:
    product_code: str
    host: str
    production_line: str
    assembly: str
    welding: str


class ProductOwnerService:
    def __init__(self, db) -> None:
        self.db = db

    def query(
        self, product_code: str = "", host: str = "", line: str = ""
    ) -> list[ProductOwnerRow]:
        product = product_code.strip().upper()
        host_text = host.strip().upper()
        line_code = line.strip().upper()
        rows = self.db.lead_fetchall(
            """
            SELECT DISTINCT "产品代号", "主机", "生产线", "装配", "焊接"
            FROM "Project Lead"
            ORDER BY "产品代号", "生产线"
            """
        )
        results: list[ProductOwnerRow] = []
        for row in rows:
            values = [_text(row[field]) for field in FIELDS]
            if product and product not in values[0].upper():
                continue
            if host_text and host_text not in values[1].upper():
                continue
            if line_code and line_code not in line_codes_in_text(values[2]):
                continue
            results.append(ProductOwnerRow(*values))
        return results


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()
