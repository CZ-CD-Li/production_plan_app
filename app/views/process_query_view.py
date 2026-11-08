"""工艺查询：按产品代号、生产线、编制检索并下载工艺文件。"""

from __future__ import annotations

from app.views.library_query_view import LibraryQueryView


class ProcessQueryView(LibraryQueryView):
    def __init__(self, db, user, parent=None) -> None:
        super().__init__(db, user, "process", parent)
