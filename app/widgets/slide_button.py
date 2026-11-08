"""悬停滑动按钮：鼠标经过时高亮块从左侧滑入，用于导航与分段切换。"""

from __future__ import annotations

from PyQt5.QtCore import QEasingCurve, QRectF, Qt, QVariantAnimation
from PyQt5.QtGui import QColor, QPainter, QPen
from PyQt5.QtWidgets import QPushButton


class SlideButton(QPushButton):
    def __init__(self, text: str = "", variant: str = "nav", parent=None) -> None:
        super().__init__(text, parent)
        self._variant = variant
        self._progress = 0.0
        self._animation = QVariantAnimation(self)
        self._animation.setDuration(190)
        self._animation.setEasingCurve(QEasingCurve.OutCubic)
        self._animation.valueChanged.connect(self._on_value)
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.NoFocus)
        self.setAttribute(Qt.WA_Hover, True)

    def _animate_to(self, value: float) -> None:
        self._animation.stop()
        self._animation.setStartValue(self._progress)
        self._animation.setEndValue(value)
        self._animation.start()

    def _on_value(self, value) -> None:
        self._progress = float(value)
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        if self.isEnabled():
            self._animate_to(1.0)
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._animate_to(0.0)
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(0.5, 1.5, -0.5, -1.5)

        if self._variant == "nav":
            self._paint_nav(painter, rect)
        else:
            self._paint_segment(painter, rect)
        painter.end()
        super().paintEvent(event)

    def _paint_nav(self, painter: QPainter, rect: QRectF) -> None:
        if self.isChecked():
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(255, 255, 255, 30))
            painter.drawRoundedRect(rect, 8, 8)
            painter.setBrush(QColor("#FFFFFF"))
            bar = QRectF(rect.left(), rect.top() + 11, 3, rect.height() - 22)
            painter.drawRoundedRect(bar, 1.5, 1.5)
            return
        if self._progress <= 0.0:
            return
        painter.save()
        painter.setClipRect(
            QRectF(rect.left(), rect.top(), rect.width() * self._progress, rect.height())
        )
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(255, 255, 255, 26))
        painter.drawRoundedRect(rect, 8, 8)
        painter.restore()

    def _paint_segment(self, painter: QPainter, rect: QRectF) -> None:
        if self.isChecked():
            painter.setPen(QPen(QColor("#E2E2E2"), 1))
            painter.setBrush(QColor("#FFFFFF"))
            painter.drawRoundedRect(rect, 8, 8)
            return
        if self._progress <= 0.0:
            return
        painter.save()
        slider = QRectF(rect.left() + 2, rect.top(), rect.width() - 4, rect.height())
        painter.setClipRect(
            QRectF(
                slider.left(),
                slider.top(),
                slider.width() * self._progress,
                slider.height(),
            )
        )
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor("#E1E1E1"))
        painter.drawRoundedRect(slider, 7, 7)
        painter.restore()
