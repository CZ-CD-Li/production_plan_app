"""个人工装 / 工艺文档 / 跟线记录通用管理页。"""

from __future__ import annotations

from PyQt5.QtWidgets import QVBoxLayout, QWidget

from app.models import User
from app.personal.file_list_panel import PersonalFileListPanel
from app.personal.file_upload_panel import FileUploadPanel


class FileManagementView(QWidget):
    def __init__(self, db, user: User, category: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("personalFilePage")
        self.category = category

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(14)

        self.upload_panel = FileUploadPanel(db, user, category)
        self.list_panel = PersonalFileListPanel(db, user, category)
        self.upload_panel.uploaded.connect(self.list_panel.refresh)
        self.upload_panel.uploaded.connect(self.upload_panel.on_activated)

        root.addWidget(self.upload_panel, 0)
        root.addWidget(self.list_panel, 1)

    def on_activated(self) -> None:
        self.upload_panel.on_activated()
        self.list_panel.on_activated()

    def select_product(self, product_code: str, line: str = "") -> None:
        self.upload_panel.select_product(product_code, line)
        self.list_panel.select_context(product_code, line)
