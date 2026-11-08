"""使用 QPainter 绘制界面图标，避免依赖外部图片资源。"""

from __future__ import annotations

from PyQt5.QtCore import QPointF, QRectF, Qt
from PyQt5.QtGui import (
    QColor,
    QIcon,
    QLinearGradient,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
)


def _draw(kind: str, painter: QPainter) -> None:
    if kind == "plan":
        painter.drawRoundedRect(QRectF(4, 5, 16, 15), 2.5, 2.5)
        painter.drawLine(QPointF(4, 9), QPointF(20, 9))
        painter.drawLine(QPointF(8, 3), QPointF(8, 6.5))
        painter.drawLine(QPointF(16, 3), QPointF(16, 6.5))
        painter.drawEllipse(QRectF(7.4, 12.3, 1.2, 1.2))
        painter.drawEllipse(QRectF(11.4, 12.3, 1.2, 1.2))
        painter.drawEllipse(QRectF(15.4, 12.3, 1.2, 1.2))
        painter.drawEllipse(QRectF(7.4, 15.8, 1.2, 1.2))
        painter.drawEllipse(QRectF(11.4, 15.8, 1.2, 1.2))
    elif kind == "search":
        painter.drawEllipse(QRectF(4, 4, 11, 11))
        painter.drawLine(QPointF(14.8, 14.8), QPointF(20, 20))
    elif kind == "personal":
        painter.drawRoundedRect(QRectF(5, 3, 14, 18), 2.2, 2.2)
        painter.drawLine(QPointF(8.5, 8), QPointF(15.5, 8))
        painter.drawLine(QPointF(8.5, 12), QPointF(15.5, 12))
        painter.drawLine(QPointF(8.5, 16), QPointF(12.5, 16))
        path = QPainterPath(QPointF(14, 17.8))
        path.lineTo(15.8, 19.5)
        path.lineTo(19, 15.8)
        painter.drawPath(path)
    elif kind == "custom":
        painter.drawLine(QPointF(4, 7), QPointF(20, 7))
        painter.drawLine(QPointF(4, 12), QPointF(20, 12))
        painter.drawLine(QPointF(4, 17), QPointF(20, 17))
        painter.drawEllipse(QRectF(7.2, 4.8, 4.2, 4.2))
        painter.drawEllipse(QRectF(14.2, 9.8, 4.2, 4.2))
        painter.drawEllipse(QRectF(9.2, 14.8, 4.2, 4.2))
    elif kind == "gear":
        painter.drawEllipse(QRectF(9, 9, 6, 6))
        for angle in range(0, 360, 45):
            painter.save()
            painter.translate(12, 12)
            painter.rotate(angle)
            painter.drawLine(QPointF(0, -5), QPointF(0, -8))
            painter.restore()
    elif kind == "refresh":
        painter.drawArc(QRectF(4, 4, 16, 16), 45 * 16, 285 * 16)
        painter.drawLine(QPointF(17.2, 2.8), QPointF(20.4, 5.2))
        painter.drawLine(QPointF(20.4, 5.2), QPointF(17.4, 7.6))
    elif kind == "arrow-left":
        painter.drawLine(QPointF(19.5, 12), QPointF(4.8, 12))
        painter.drawLine(QPointF(10.5, 6.8), QPointF(4.8, 12))
        painter.drawLine(QPointF(10.5, 17.2), QPointF(4.8, 12))
    elif kind == "user":
        painter.drawEllipse(QRectF(9.2, 4.5, 5.6, 5.6))
        path = QPainterPath(QPointF(5.5, 19.5))
        path.cubicTo(6.2, 14.8, 17.8, 14.8, 18.5, 19.5)
        painter.drawPath(path)
    elif kind == "lock":
        painter.drawRoundedRect(QRectF(5, 10.5, 14, 9.5), 2, 2)
        painter.drawArc(QRectF(8, 4.5, 8, 10), 0, 180 * 16)
        painter.drawLine(QPointF(9, 9.5), QPointF(9, 10.5))
        painter.drawLine(QPointF(15, 9.5), QPointF(15, 10.5))
    elif kind == "folder":
        path = QPainterPath(QPointF(3.5, 7))
        path.lineTo(9.2, 7)
        path.lineTo(11.2, 9.2)
        path.lineTo(20.5, 9.2)
        path.lineTo(20.5, 19)
        path.lineTo(3.5, 19)
        path.closeSubpath()
        painter.drawPath(path)
    elif kind == "file":
        path = QPainterPath(QPointF(6, 3))
        path.lineTo(14, 3)
        path.lineTo(18.5, 7.5)
        path.lineTo(18.5, 21)
        path.lineTo(6, 21)
        path.closeSubpath()
        painter.drawPath(path)
        painter.drawLine(QPointF(14, 3), QPointF(14, 8))
        painter.drawLine(QPointF(14, 8), QPointF(18.5, 8))
        painter.drawLine(QPointF(9, 13), QPointF(15.5, 13))
        painter.drawLine(QPointF(9, 17), QPointF(15.5, 17))
    elif kind == "spark":
        path = QPainterPath(QPointF(12, 3))
        path.lineTo(14.2, 9.8)
        path.lineTo(21, 12)
        path.lineTo(14.2, 14.2)
        path.lineTo(12, 21)
        path.lineTo(9.8, 14.2)
        path.lineTo(3, 12)
        path.lineTo(9.8, 9.8)
        path.closeSubpath()
        painter.drawPath(path)


def _pixmap(kind: str, color: str, size: int) -> QPixmap:
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.scale(size / 24.0, size / 24.0)
    pen = QPen(QColor(color))
    pen.setWidthF(1.7)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    _draw(kind, painter)
    painter.end()
    return pixmap


def nav_icon(kind: str) -> QIcon:
    icon = QIcon()
    icon.addPixmap(_pixmap(kind, "#9C9C9C", 44), QIcon.Normal, QIcon.Off)
    icon.addPixmap(_pixmap(kind, "#FFFFFF", 44), QIcon.Normal, QIcon.On)
    icon.addPixmap(_pixmap(kind, "#FFFFFF", 44), QIcon.Active, QIcon.On)
    return icon


def line_icon(kind: str, color: str = "#333333", size: int = 36) -> QIcon:
    return QIcon(_pixmap(kind, color, size))


def app_icon() -> QIcon:
    icon = QIcon()
    for size in (16, 24, 32, 48, 64, 128, 256):
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)
        gradient = QLinearGradient(0, 0, size, size)
        gradient.setColorAt(0, QColor("#4A4A4A"))
        gradient.setColorAt(1, QColor("#0D0D0D"))
        painter.setPen(Qt.NoPen)
        painter.setBrush(gradient)
        radius = size * 0.24
        painter.drawRoundedRect(QRectF(1, 1, size - 2, size - 2), radius, radius)

        pen = QPen(QColor("#FFFFFF"))
        pen.setWidthF(max(1.2, size * 0.045))
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        left = size * 0.26
        top = size * 0.24
        width = size * 0.48
        height = size * 0.5
        painter.drawRoundedRect(
            QRectF(left, top, width, height), size * 0.05, size * 0.05
        )
        painter.drawLine(
            QPointF(left, top + height * 0.28),
            QPointF(left + width, top + height * 0.28),
        )
        painter.drawLine(
            QPointF(left + width * 0.3, top - size * 0.045),
            QPointF(left + width * 0.3, top + height * 0.1),
        )
        painter.drawLine(
            QPointF(left + width * 0.7, top - size * 0.045),
            QPointF(left + width * 0.7, top + height * 0.1),
        )
        check = QPainterPath(QPointF(left + width * 0.2, top + height * 0.62))
        check.lineTo(left + width * 0.42, top + height * 0.82)
        check.lineTo(left + width * 0.82, top + height * 0.4)
        painter.drawPath(check)
        painter.end()
        icon.addPixmap(pixmap)
    return icon
