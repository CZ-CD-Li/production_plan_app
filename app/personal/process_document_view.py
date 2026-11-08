"""个人-工艺管理页。"""

from __future__ import annotations

from app.models import User
from app.personal.file_management_view import FileManagementView


class PersonalProcessDocView(FileManagementView):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(db, user, "process", parent)
