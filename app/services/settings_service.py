"""系统设置读写：路径与终端 IP 按登录用户独立保存，互不影响。"""

from __future__ import annotations

from app.config import DEMO_DIR
from app.database import Database

DEMO_PROJECT_FILE = DEMO_DIR / "新产品项目任务分布清单.xlsx"

# 以下设置项按用户（工号）独立存储；其余设置项（如同步时间戳）全局共享
USER_SCOPED_KEYS = (
    "ip1",
    "ip2",
    "ip3",
    "production_plan_dir",
    "project_lead_path",
)


class SettingsService:
    DEFAULT_IPS = ("192.168.1.101", "192.168.1.102", "192.168.1.103")

    def __init__(self, db: Database) -> None:
        self.db = db
        self._ensure_defaults()

    def _default_values(self) -> dict[str, str]:
        return {
            "ip1": self.DEFAULT_IPS[0],
            "ip2": self.DEFAULT_IPS[1],
            "ip3": self.DEFAULT_IPS[2],
            "production_plan_dir": str(DEMO_DIR) if DEMO_DIR.exists() else "",
            "project_lead_path": str(DEMO_PROJECT_FILE) if DEMO_PROJECT_FILE.exists() else "",
            "project_lead_synced_at": "",
        }

    @staticmethod
    def _scoped_key(key: str, employee_no: str | None) -> str:
        if employee_no and key in USER_SCOPED_KEYS:
            return f"{employee_no}:{key}"
        return key

    def _ensure_defaults(self) -> None:
        # 只保证全局默认项存在；用户级设置在首次保存时按需写入
        existing = {row["key"] for row in self.db.fetchall("SELECT key FROM settings")}
        missing = {
            key: value
            for key, value in self._default_values().items()
            if key not in existing and key not in USER_SCOPED_KEYS and value
        }
        if missing:
            with self.db.transaction() as conn:
                conn.executemany(
                    "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                    list(missing.items()),
                )

    def get_all(self, employee_no: str | None = None) -> dict[str, str]:
        raw = {
            row["key"]: row["value"]
            for row in self.db.fetchall("SELECT key, value FROM settings")
        }
        values: dict[str, str] = {}
        for key, default in self._default_values().items():
            if key in USER_SCOPED_KEYS and employee_no:
                # 用户级优先（含用户主动清空的空值）；无用户级记录时回退全局旧值
                scoped = self._scoped_key(key, employee_no)
                if scoped in raw:
                    values[key] = raw[scoped]
                elif key in raw:
                    values[key] = raw[key]
                else:
                    values[key] = default
            else:
                values[key] = raw.get(key, default)
        return values

    def get(self, key: str, default: str = "", employee_no: str | None = None) -> str:
        scoped = self._scoped_key(key, employee_no)
        row = self.db.fetchone("SELECT value FROM settings WHERE key = ?", (scoped,))
        if row is not None:
            return row["value"]
        if scoped != key:
            row = self.db.fetchone("SELECT value FROM settings WHERE key = ?", (key,))
            if row is not None:
                return row["value"]
        return default

    def set_many(self, values: dict[str, str], employee_no: str | None = None) -> None:
        if not values:
            return
        pairs = [
            (self._scoped_key(key, employee_no), str(value))
            for key, value in values.items()
        ]
        with self.db.transaction() as conn:
            conn.executemany(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                pairs,
            )
