"""头像与可点击容器组件。"""

from __future__ import annotations

from PyQt5.QtCore import QRectF, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PyQt5.QtWidgets import QFrame, QWidget

_GRADIENTS = (
    ("#5C5C5C", "#1C1C1C"),
    ("#7A7A7A", "#2A2A2A"),
    ("#4A4A4A", "#111111"),
    ("#6E6E6E", "#242424"),
    ("#8C8C8C", "#3A3A3A"),
)


class Avatar(QWidget):
    def __init__(self, name: str, size: int = 44, parent=None) -> None:
        super().__init__(parent)
        self._name = name or ""
        self._size = size
        self.setFixedSize(size, size)

    def set_name(self, name: str) -> None:
        self._name = name
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(0.5, 0.5, self._size - 1, self._size - 1)

        seed = sum(ord(char) for char in self._name) % len(_GRADIENTS)
        start, end = _GRADIENTS[seed]
        gradient = QLinearGradient(0, 0, self._size, self._size)
        gradient.setColorAt(0, QColor(start))
        gradient.setColorAt(1, QColor(end))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        painter.drawEllipse(rect)

        painter.setPen(QPen(QColor(255, 255, 255, 50), 1))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(rect.adjusted(1.5, 1.5, -1.5, -1.5))

        initial = (self._name.strip()[:1] or "U").upper()
        font = QFont("Microsoft YaHei UI", int(self._size * 0.34), QFont.DemiBold)
        font.setPixelSize(max(13, int(self._size * 0.36)))
        painter.setFont(font)
        painter.setPen(QColor("#FFFFFF"))
        painter.drawText(rect, Qt.AlignCenter, initial)
        painter.end()


class ClickableFrame(QFrame):
    clicked = pyqtSignal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setCursor(Qt.PointingHandCursor)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)
