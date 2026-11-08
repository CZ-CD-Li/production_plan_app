"""个人-工装管理页。"""

from __future__ import annotations

from app.models import User
from app.personal.file_management_view import FileManagementView


class PersonalToolingView(FileManagementView):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(db, user, "tooling", parent)
