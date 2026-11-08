"""登录认证与密码修改。"""

from __future__ import annotations

from app.database import Database
from app.models import User
from app.utils.security import hash_password, verify_password


class AuthService:
    def __init__(self, db: Database) -> None:
        self.db = db

    def authenticate(self, employee_no: str, password: str) -> User | None:
        employee_no = employee_no.strip()
        if not employee_no or not password:
            return None
        row = self.db.fetchone(
            "SELECT id, employee_no, password_hash, name, role FROM users WHERE employee_no = ?",
            (employee_no,),
        )
        if row is None or not verify_password(password, row["password_hash"]):
            return None
        return User(
            id=int(row["id"]),
            employee_no=row["employee_no"],
            name=row["name"],
            role=row["role"],
        )

    def change_password(
        self, user: User, old_password: str, new_password: str, confirm_password: str
    ) -> tuple[bool, str]:
        if not old_password or not new_password:
            return False, "请完整填写原密码与新密码。"
        if len(new_password) < 6:
            return False, "新密码长度不能少于 6 位。"
        if new_password != confirm_password:
            return False, "两次输入的新密码不一致。"
        row = self.db.fetchone(
            "SELECT password_hash FROM users WHERE id = ?", (user.id,)
        )
        if row is None or not verify_password(old_password, row["password_hash"]):
            return False, "原密码不正确。"
        self.db.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (hash_password(new_password), user.id),
        )
        return True, "密码修改成功，请使用新密码登录。"
