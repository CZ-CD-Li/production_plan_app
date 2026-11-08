"""SQLite 数据访问层：共享库（账户/设置/个人文件/CAPP）+ 本机 Project Lead 库。"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from app.config import (
    ensure_runtime_dirs,
    is_network_path,
    resolve_db_path,
    resolve_lead_db_path,
)
from app.utils.security import hash_password


class Database:
    """共享库（账户、设置、个人文件、CAPP 记录）+ 本机 Project Lead 库。"""

    def __init__(
        self,
        path: str | Path | None = None,
        lead_path: str | Path | None = None,
    ) -> None:
        ensure_runtime_dirs()
        self.path = Path(path) if path is not None else resolve_db_path()
        self.is_network = is_network_path(self.path)
        if lead_path is not None:
            self.lead_path = Path(lead_path)
        elif path is not None:
            # 显式传入共享库路径（测试/多实例）时，本机库放在同一目录
            self.lead_path = self.path.with_name("project_lead.db")
        else:
            self.lead_path = resolve_lead_db_path()
        self.lead_is_network = is_network_path(self.lead_path)
        if not self.path.parent.exists():
            self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.lead_path.parent.exists():
            self.lead_path.parent.mkdir(parents=True, exist_ok=True)
        lead_was_new = not self.lead_path.exists()
        # 网络盘上多人同时写入时，等待锁而不是立即失败
        self.conn = sqlite3.connect(
            str(self.path), timeout=15, check_same_thread=False
        )
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA busy_timeout=15000")
        # WAL 依赖共享内存，在网络共享盘上不可靠；共享库统一退回 DELETE 日志
        mode = "DELETE" if self.is_network else "WAL"
        try:
            self.conn.execute(f"PRAGMA journal_mode={mode}")
        except sqlite3.OperationalError as exc:
            # 切换到 DELETE 需要独占访问，其他主机正在写入时会短暂失败
            raise sqlite3.OperationalError(
                f"无法将数据库日志模式切换为 {mode}：{exc}。"
                "请稍后重启，或确认共享盘上没有其他主机正在运行。"
            ) from exc
        self.conn.execute("PRAGMA synchronous=NORMAL")
        self._create_schema()
        self._seed_accounts()
        self._open_lead_database(lead_was_new)

    # ---- 本机 Project Lead 库 ----

    def _open_lead_database(self, lead_was_new: bool) -> None:
        self.lead_conn = sqlite3.connect(
            str(self.lead_path), timeout=15, check_same_thread=False
        )
        self.lead_conn.row_factory = sqlite3.Row
        self.lead_conn.execute("PRAGMA busy_timeout=15000")
        lead_mode = "DELETE" if self.lead_is_network else "WAL"
        self.lead_conn.execute(f"PRAGMA journal_mode={lead_mode}")
        self.lead_conn.execute("PRAGMA synchronous=NORMAL")
        self._ensure_project_lead_table(self.lead_conn)
        if lead_was_new:
            self._copy_lead_from_shared()

    def _copy_lead_from_shared(self) -> None:
        """首次启用本机库时，把共享库里已有的 Project Lead 数据复制过来。"""
        exists = self.conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='Project Lead'"
        ).fetchone()
        if not exists:
            return
        rows = self.conn.execute(
            """
            SELECT id, "主机", "产品代号", "生产线", "焊接", "装配", mass
            FROM "Project Lead"
            """
        ).fetchall()
        if not rows:
            return
        with self.lead_conn:
            self.lead_conn.execute('DELETE FROM "Project Lead"')
            self.lead_conn.executemany(
                """
                INSERT INTO "Project Lead"
                    (id, "主机", "产品代号", "生产线", "焊接", "装配", mass)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [tuple(row) for row in rows],
            )

    @contextmanager
    def lead_transaction(self, immediate: bool = False):
        try:
            if immediate and not self.lead_conn.in_transaction:
                self.lead_conn.execute("BEGIN IMMEDIATE")
            yield self.lead_conn
            self.lead_conn.commit()
        except Exception:
            self.lead_conn.rollback()
            raise

    def lead_fetchall(
        self, sql: str, params: Iterable[Any] = ()
    ) -> list[sqlite3.Row]:
        return list(self.lead_conn.execute(sql, tuple(params)).fetchall())

    def lead_fetchone(
        self, sql: str, params: Iterable[Any] = ()
    ) -> sqlite3.Row | None:
        return self.lead_conn.execute(sql, tuple(params)).fetchone()

    def _create_schema(self) -> None:
        with self.conn:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_no TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    name TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT '生产计划员',
                    created_at TEXT NOT NULL
                )
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            self._ensure_personal_file_tables()
            self._ensure_capp_table()

    PERSONAL_FILE_TABLES = ("工装管理", "工艺文档管理", "跟线记录")

    def _ensure_personal_file_tables(self) -> None:
        """个人工装、工艺文档与跟线记录共用的文件登记表。"""
        for table in self.PERSONAL_FILE_TABLES:
            self.conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS "{table}" (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    product_code TEXT NOT NULL,
                    production_line TEXT NOT NULL,
                    process TEXT NOT NULL,
                    name TEXT NOT NULL,
                    employee_no TEXT NOT NULL,
                    employee_name TEXT NOT NULL,
                    save_dir TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    file_name TEXT NOT NULL,
                    uploaded_at TEXT NOT NULL
                )
                """
            )
            self.conn.execute(
                f'CREATE INDEX IF NOT EXISTS "idx_{table}_owner" '
                f'ON "{table}" (employee_no, production_line, product_code)'
            )

    def _ensure_capp_table(self) -> None:
        """CAPP 质量策划生成记录。"""
        self.conn.execute(
            """
            CREATE TABLE IF NOT EXISTS "CAPP质量策划记录" (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_code TEXT NOT NULL,
                source_file TEXT NOT NULL,
                output_path TEXT NOT NULL,
                employee_no TEXT,
                employee_name TEXT,
                created_at TEXT NOT NULL
            )
            """
        )

    def _ensure_project_lead_table(self, conn: sqlite3.Connection) -> None:
        """按最新字段结构维护本机 Project Lead 表，旧结构直接丢弃并重建。"""
        expected = ["id", "主机", "产品代号", "生产线", "焊接", "装配", "mass"]
        existing = [
            row[1] for row in conn.execute('PRAGMA table_info("Project Lead")')
        ]
        if existing and existing != expected:
            conn.execute('DROP TABLE "Project Lead"')
            existing = []
        if not existing:
            conn.execute(
                """
                CREATE TABLE "Project Lead" (
                    id TEXT,
                    "主机" TEXT,
                    "产品代号" TEXT NOT NULL,
                    "生产线" TEXT,
                    "焊接" TEXT,
                    "装配" TEXT,
                    mass INTEGER NOT NULL DEFAULT 0
                )
                """
            )
        conn.execute(
            'CREATE INDEX IF NOT EXISTS idx_project_lead_code ON "Project Lead" ("产品代号")'
        )

    def _seed_accounts(self) -> None:
        count = self.conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if count:
            return
        now = datetime.now().isoformat(timespec="seconds")
        with self.conn:
            self.conn.executemany(
                """
                INSERT INTO users (employee_no, password_hash, name, role, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                [
                    ("1001", hash_password("123456"), "张伟", "生产计划员", now),
                    ("1002", hash_password("123456"), "李娜", "工艺工程师", now),
                ],
            )

    @contextmanager
    def transaction(self, immediate: bool = False):
        """事务上下文；immediate=True 时先取得写锁，避免多机同时写时互相等待。"""
        try:
            if immediate and not self.conn.in_transaction:
                self.conn.execute("BEGIN IMMEDIATE")
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def execute(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Cursor:
        with self.conn:
            return self.conn.execute(sql, tuple(params))

    def executemany(self, sql: str, rows: Iterable[Iterable[Any]]) -> None:
        with self.conn:
            self.conn.executemany(sql, rows)

    def fetchone(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Row | None:
        return self.conn.execute(sql, tuple(params)).fetchone()

    def fetchall(self, sql: str, params: Iterable[Any] = ()) -> list[sqlite3.Row]:
        return list(self.conn.execute(sql, tuple(params)).fetchall())

    def close(self) -> None:
        self.conn.close()
        self.lead_conn.close()
