"""个人-跟线记录页。"""

from __future__ import annotations

from app.models import User
from app.personal.file_management_view import FileManagementView


class PersonalTrackingView(FileManagementView):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(db, user, "tracking", parent)
