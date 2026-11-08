"""支持点击选择与拖拽的文件投放区域。"""

from __future__ import annotations

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from PyQt5.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
)

from app.utils.icons import line_icon


class FileDropArea(QFrame):
    file_selected = pyqtSignal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("dropArea")
        self.setAcceptDrops(True)
        self.setMinimumHeight(88)
        self._file_path = ""

        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(14)

        icon = QLabel()
        icon.setPixmap(line_icon("file", "#6B6B6B", 48).pixmap(28, 28))
        icon.setStyleSheet("background:transparent;")
        layout.addWidget(icon, 0, Qt.AlignVCenter)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(3)
        self.title_label = QLabel("拖入文件，或点击右侧按钮选择")
        self.title_label.setStyleSheet(
            "background:transparent;color:#333333;font-size:13px;font-weight:600;"
        )
        self.file_label = QLabel("支持任意格式，上传前会自动按规则重命名")
        self.file_label.setObjectName("fieldHint")
        self.file_label.setStyleSheet("background:transparent;")
        text_layout.addWidget(self.title_label)
        text_layout.addWidget(self.file_label)
        layout.addLayout(text_layout, 1)

        self.choose_button = QPushButton("选择文件")
        self.choose_button.setObjectName("secondaryButton")
        self.choose_button.setCursor(Qt.PointingHandCursor)
        self.choose_button.clicked.connect(self._choose_file)
        layout.addWidget(self.choose_button, 0, Qt.AlignVCenter)

    def file_path(self) -> str:
        return self._file_path

    def set_file(self, path: str) -> None:
        self._file_path = path
        if path:
            name = path.replace("\\", "/").split("/")[-1]
            self.file_label.setText(name)
            self.file_label.setToolTip(path)
        else:
            self.file_label.setText("支持任意格式，上传前会自动按规则重命名")
            self.file_label.setToolTip("")
        self._set_dragging(False)
        self.file_selected.emit(path)

    def _choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "所有文件 (*.*)")
        if path:
            self.set_file(path)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_dragging(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self._set_dragging(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        self._set_dragging(False)
        for url in event.mimeData().urls():
            if url.isLocalFile():
                self.set_file(url.toLocalFile())
                event.acceptProposedAction()
                return
        event.ignore()

    def _set_dragging(self, dragging: bool) -> None:
        self.setProperty("dragging", "true" if dragging else "false")
        self.style().unpolish(self)
        self.style().polish(self)
