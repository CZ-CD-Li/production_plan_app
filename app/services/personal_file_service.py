"""个人工装 / 工艺文档 / 跟线记录的登记、查询与文件保存。"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.config import ENFORCE_UDC_ADDRESS
from app.database import Database
from app.models import User

CATEGORY_CONFIG: dict[str, dict[str, str]] = {
    "tooling": {
        "table": "工装管理",
        "label": "工装",
        "upload_title": "工装上传",
        "name_label": "工装名称",
        "list_title": "个人工装列表",
    },
    "process": {
        "table": "工艺文档管理",
        "label": "工艺文档",
        "upload_title": "工艺文档上传",
        "name_label": "工艺文档名称",
        "list_title": "个人工艺文档列表",
    },
    "tracking": {
        "table": "跟线记录",
        "label": "跟线记录",
        "upload_title": "跟线记录上传",
        "name_label": "跟线记录名称",
        "list_title": "个人跟线记录列表",
    },
}

_INVALID_FILENAME = re.compile(r'[\\/:*?"<>|]')


class PersonalFileError(Exception):
    pass


class UdcPathError(PersonalFileError):
    pass


@dataclass(frozen=True)
class PersonalFileRecord:
    id: int
    product_code: str
    production_line: str
    process: str
    name: str
    employee_no: str
    employee_name: str
    save_dir: str
    file_path: str
    file_name: str
    uploaded_at: str


def is_udc_path(value: str) -> bool:
    """判断保存地址是否为生产网 UDC 地址。

    接受 UNC 网络路径（\\\\服务器\\共享）以及显式 udc: 路径。
    """
    text = value.strip()
    if not text:
        return False
    normalized = text.replace("/", "\\")
    if normalized.lower().startswith("udc:"):
        return True
    if normalized.startswith("\\\\"):
        parts = [part for part in normalized.strip("\\").split("\\") if part]
        return len(parts) >= 2
    return "\\udc\\" in normalized.lower()


def build_file_name(
    product_code: str,
    production_line: str,
    process: str,
    name: str,
    suffix: str,
) -> str:
    parts = [
        product_code.strip(),
        production_line.strip(),
        process.strip(),
        name.strip(),
    ]
    safe_parts = [_INVALID_FILENAME.sub("_", part) for part in parts]
    stem = "-".join(part for part in safe_parts if part)
    if not stem:
        raise PersonalFileError("产品代号、生产线、工序和名称不能同时为空。")
    return stem + suffix


class PersonalFileService:
    def __init__(self, db: Database) -> None:
        self.db = db

    @staticmethod
    def config(category: str) -> dict[str, str]:
        config = CATEGORY_CONFIG.get(category)
        if config is None:
            raise PersonalFileError(f"未知文件分类：{category}")
        return config

    def _table(self, category: str) -> str:
        return self.config(category)["table"]

    def list_lines(self, category: str, employee_no: str) -> list[str]:
        rows = self.db.fetchall(
            f'SELECT DISTINCT production_line FROM "{self._table(category)}" '
            "WHERE employee_no = ? ORDER BY production_line",
            (employee_no,),
        )
        return [(row["production_line"] or "").strip() for row in rows]

    def list_products(self, category: str, employee_no: str, line: str) -> list[str]:
        rows = self.db.fetchall(
            f'SELECT DISTINCT product_code FROM "{self._table(category)}" '
            "WHERE employee_no = ? AND production_line = ? ORDER BY product_code",
            (employee_no, line),
        )
        return [(row["product_code"] or "").strip() for row in rows]

    def list_files(
        self, category: str, employee_no: str, line: str, product_code: str
    ) -> list[PersonalFileRecord]:
        rows = self.db.fetchall(
            f'SELECT * FROM "{self._table(category)}" '
            "WHERE employee_no = ? AND production_line = ? AND product_code = ? "
            "ORDER BY uploaded_at DESC, id DESC",
            (employee_no, line, product_code),
        )
        return [self._to_record(row) for row in rows]

    def list_uploaders(self, category: str) -> list[str]:
        """所有上传过该分类文件的人员姓名（编制）。"""
        rows = self.db.fetchall(
            f'SELECT DISTINCT employee_name FROM "{self._table(category)}" '
            'WHERE employee_name <> "" ORDER BY employee_name'
        )
        return [(row["employee_name"] or "").strip() for row in rows]

    def query_files(
        self,
        category: str,
        product_code: str = "",
        line: str = "",
        author: str = "",
    ) -> list[PersonalFileRecord]:
        """跨用户查询：产品代号模糊匹配，生产线 / 编制精确或模糊匹配。"""
        clauses: list[str] = []
        params: list[str] = []
        if product_code.strip():
            clauses.append("product_code LIKE ?")
            params.append(f"%{product_code.strip()}%")
        if line.strip():
            clauses.append("production_line = ? COLLATE NOCASE")
            params.append(line.strip())
        if author.strip():
            clauses.append("(employee_name LIKE ? OR employee_no LIKE ?)")
            params.extend([f"%{author.strip()}%", f"%{author.strip()}%"])
        where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
        rows = self.db.fetchall(
            f'SELECT * FROM "{self._table(category)}"{where} '
            "ORDER BY production_line, product_code, uploaded_at DESC, id DESC",
            tuple(params),
        )
        return [self._to_record(row) for row in rows]

    def get(self, category: str, record_id: int) -> PersonalFileRecord | None:
        row = self.db.fetchone(
            f'SELECT * FROM "{self._table(category)}" WHERE id = ?', (record_id,)
        )
        return self._to_record(row) if row is not None else None

    def save_upload(
        self,
        category: str,
        user: User,
        source_file: str | Path,
        save_dir: str | Path,
        product_code: str,
        production_line: str,
        process: str,
        name: str,
    ) -> PersonalFileRecord:
        table = self._table(category)
        source = Path(source_file)
        if not source.is_file():
            raise PersonalFileError("请选择要上传的文件。")
        if not product_code.strip():
            raise PersonalFileError("请选择产品代号。")
        if not production_line.strip():
            raise PersonalFileError("请输入生产线。")
        if category != "process" and not process.strip():
            raise PersonalFileError("请选择工序。")
        if not name.strip():
            raise PersonalFileError("请输入名称。")

        save_dir_text = str(save_dir).strip()
        if ENFORCE_UDC_ADDRESS and not is_udc_path(save_dir_text):
            raise UdcPathError("只能将文件保存在生产网上。")

        target_dir = Path(save_dir_text)
        try:
            exists = target_dir.is_dir()
        except OSError as exc:
            raise PersonalFileError(f"无法访问保存地址：{exc}") from exc
        if not exists:
            raise PersonalFileError(f"保存地址不存在：{target_dir}")

        file_name = build_file_name(
            product_code, production_line, process, name, source.suffix
        )
        target = target_dir / file_name
        if target.exists():
            raise PersonalFileError(f"目标位置已存在同名文件：{file_name}")

        try:
            shutil.copy2(source, target)
        except OSError as exc:
            raise PersonalFileError(f"文件保存失败：{exc}") from exc

        uploaded_at = datetime.now().isoformat(timespec="seconds")
        with self.db.transaction() as conn:
            cursor = conn.execute(
                f"""
                INSERT INTO "{table}"
                    (product_code, production_line, process, name,
                     employee_no, employee_name, save_dir, file_path,
                     file_name, uploaded_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    product_code.strip(),
                    production_line.strip(),
                    process.strip(),
                    name.strip(),
                    user.employee_no,
                    user.name,
                    str(target_dir),
                    str(target),
                    file_name,
                    uploaded_at,
                ),
            )
            record_id = int(cursor.lastrowid)
        record = self.get(category, record_id)
        if record is None:
            raise PersonalFileError("文件已保存，但登记记录写入失败。")
        return record

    def delete(self, category: str, record_id: int, remove_file: bool = True) -> None:
        record = self.get(category, record_id)
        if record is None:
            return
        with self.db.transaction() as conn:
            conn.execute(
                f'DELETE FROM "{self._table(category)}" WHERE id = ?', (record_id,)
            )
        if remove_file:
            try:
                path = Path(record.file_path)
                if path.is_file():
                    path.unlink()
            except OSError:
                # 网络文件暂时不可达时，登记记录已删除，不阻断用户操作
                pass

    @staticmethod
    def _to_record(row) -> PersonalFileRecord:
        return PersonalFileRecord(
            id=int(row["id"]),
            product_code=row["product_code"] or "",
            production_line=row["production_line"] or "",
            process=row["process"] or "",
            name=row["name"] or "",
            employee_no=row["employee_no"] or "",
            employee_name=row["employee_name"] or "",
            save_dir=row["save_dir"] or "",
            file_path=row["file_path"] or "",
            file_name=row["file_name"] or "",
            uploaded_at=row["uploaded_at"] or "",
        )
