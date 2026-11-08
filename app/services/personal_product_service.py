"""当前用户负责的产品查询。"""

from __future__ import annotations

import re
from dataclasses import dataclass

from app.database import Database

_SPLIT_NAMES = re.compile(r"[\s,，、;；/\\|]+")


def _names_in_field(field: str, user_name: str) -> bool:
    """判断负责人字段中是否包含当前用户姓名（按分隔符精确比对）。"""
    text = (field or "").strip()
    if not text:
        return False
    if text == user_name:
        return True
    return user_name in _SPLIT_NAMES.split(text)


@dataclass(frozen=True)
class PersonalProduct:
    product_code: str
    host: str
    production_line: str


class PersonalProductService:
    def __init__(self, db: Database) -> None:
        self.db = db

    def products_for(self, user_name: str) -> list[PersonalProduct]:
        if not user_name:
            return []
        rows = self.db.lead_fetchall(
            """
            SELECT DISTINCT "产品代号" AS code, "主机" AS host, "生产线" AS line,
                   "焊接" AS welding, "装配" AS assembly
            FROM "Project Lead"
            WHERE instr("装配", ?) > 0 OR instr("焊接", ?) > 0
            ORDER BY "生产线", "产品代号"
            """,
            (user_name, user_name),
        )
        products: list[PersonalProduct] = []
        for row in rows:
            code = (row["code"] or "").strip()
            if not code:
                continue
            if not (
                _names_in_field(row["assembly"], user_name)
                or _names_in_field(row["welding"], user_name)
            ):
                continue
            products.append(
                PersonalProduct(
                    product_code=code,
                    host=(row["host"] or "").strip(),
                    production_line=(row["line"] or "").strip(),
                )
            )
        return products

    def lines_for(self, user_name: str) -> list[str]:
        lines = {item.production_line for item in self.products_for(user_name)}
        return sorted(line for line in lines if line)

    def products_by_line(self, user_name: str, line: str) -> list[PersonalProduct]:
        return [
            item
            for item in self.products_for(user_name)
            if item.production_line == line
        ]
