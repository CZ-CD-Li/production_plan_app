"""自绘窗口标题栏：品牌信息 + 最小化/最大化/关闭。"""

from __future__ import annotations

from PyQt5.QtCore import QPoint, QRectF, Qt
from PyQt5.QtGui import QColor, QPainter, QPen
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QWidget,
)

from app.utils.icons import app_icon

_BUTTON_SIZE = (44, 30)


class _WindowButton(QPushButton):
    def __init__(self, kind: str, theme: str = "light", parent=None) -> None:
        super().__init__(parent)
        self._kind = kind
        self._theme = theme
        self._hovered = False
        self._maximized = False
        self.setFixedSize(*_BUTTON_SIZE)
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.NoFocus)
        self.setToolTip(
            {"min": "最小化", "max": "最大化", "close": "关闭"}.get(kind, "")
        )

    def set_maximized(self, value: bool) -> None:
        self._maximized = value
        self.setToolTip("还原" if value else "最大化")
        self.update()

    def enterEvent(self, event) -> None:  # noqa: N802
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hovered = False
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(2, 2, -2, -2)
        dark = self._theme == "dark"

        if self._hovered:
            if self._kind == "close":
                fill = QColor("#1A1A1A") if not dark else QColor("#F2F2F2")
            else:
                fill = QColor("#F0F0F0") if not dark else QColor("#2B2B2B")
            painter.setPen(Qt.NoPen)
            painter.setBrush(fill)
            painter.drawRoundedRect(rect, 6, 6)

        if self._kind == "close" and self._hovered:
            color = QColor("#FFFFFF") if not dark else QColor("#141414")
        else:
            color = QColor("#333333") if not dark else QColor("#CFCFCF")

        painter.setPen(QPen(color, 1.3))
        center = rect.center()
        cx, cy = center.x(), center.y()
        if self._kind == "min":
            painter.drawLine(QPoint(int(cx - 5), int(cy + 1)), QPoint(int(cx + 5), int(cy + 1)))
        elif self._kind == "close":
            painter.drawLine(QPoint(int(cx - 4), int(cy - 4)), QPoint(int(cx + 4), int(cy + 4)))
            painter.drawLine(QPoint(int(cx - 4), int(cy + 4)), QPoint(int(cx + 4), int(cy - 4)))
        elif self._maximized:
            painter.drawRect(QRectF(cx - 4, cy - 5, 8, 8))
            painter.drawLine(QPoint(int(cx - 2), int(cy - 5)), QPoint(int(cx + 4), int(cy - 5)))
            painter.drawLine(QPoint(int(cx + 4), int(cy - 5)), QPoint(int(cx + 4), int(cy + 1)))
        else:
            painter.drawRect(QRectF(cx - 4.5, cy - 4.5, 9, 9))
        painter.end()


class WindowControls(QWidget):
    """窗口按钮组：最小化 / 最大化 / 关闭，可嵌入页面任意位置。"""

    def __init__(
        self,
        theme: str = "light",
        *,
        minimizable: bool = True,
        maximizable: bool = True,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._maximizable = maximizable
        self.setAttribute(Qt.WA_StyledBackground, False)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        self.min_button = _WindowButton("min", theme)
        self.min_button.setVisible(minimizable)
        self.min_button.clicked.connect(self._minimize)
        layout.addWidget(self.min_button, 0, Qt.AlignVCenter)

        self.max_button = _WindowButton("max", theme)
        self.max_button.setVisible(maximizable)
        self.max_button.clicked.connect(self.toggle_maximize)
        layout.addWidget(self.max_button, 0, Qt.AlignVCenter)

        self.close_button = _WindowButton("close", theme)
        self.close_button.clicked.connect(self._close)
        layout.addWidget(self.close_button, 0, Qt.AlignVCenter)

    def _minimize(self) -> None:
        self.window().showMinimized()

    def _close(self) -> None:
        self.window().close()

    def toggle_maximize(self) -> None:
        if not self._maximizable:
            return
        window = self.window()
        if window.isMaximized():
            window.showNormal()
        else:
            window.showMaximized()
        self.sync_window_state(window.isMaximized())

    def sync_window_state(self, maximized: bool) -> None:
        self.max_button.set_maximized(maximized)


class WindowTitleBar(QWidget):
    """标题栏内建窗口操作，双击空白区域切换最大化。"""

    def __init__(
        self,
        title: str,
        subtitle: str = "",
        *,
        theme: str = "light",
        minimizable: bool = True,
        maximizable: bool = True,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("titleBar")
        self.setProperty("theme", theme)
        self.setFixedHeight(40)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._maximizable = maximizable
        self._drag_offset: QPoint | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 0, 8, 0)
        layout.setSpacing(9)

        logo = QLabel()
        logo.setPixmap(app_icon().pixmap(18, 18))
        logo.setFixedSize(18, 18)
        layout.addWidget(logo, 0, Qt.AlignVCenter)

        title_label = QLabel(title)
        title_label.setObjectName("titleBarTitle")
        layout.addWidget(title_label, 0, Qt.AlignVCenter)

        if subtitle:
            dot = QLabel("·")
            dot.setObjectName("titleBarSub")
            layout.addWidget(dot, 0, Qt.AlignVCenter)
            sub_label = QLabel(subtitle)
            sub_label.setObjectName("titleBarSub")
            layout.addWidget(sub_label, 0, Qt.AlignVCenter)

        layout.addStretch(1)

        self.controls = WindowControls(
            theme, minimizable=minimizable, maximizable=maximizable
        )
        layout.addWidget(self.controls, 0, Qt.AlignVCenter)

    # ---- 窗口操作 ----

    def toggle_maximize(self) -> None:
        self.controls.toggle_maximize()

    def sync_window_state(self, maximized: bool) -> None:
        self.controls.sync_window_state(maximized)

    # ---- 拖动 ----

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() != Qt.LeftButton:
            super().mousePressEvent(event)
            return
        window = self.window()
        if window.isMaximized():
            ratio = event.x() / max(1, self.width())
            window.showNormal()
            self.sync_window_state(False)
            window.move(
                event.globalPos().x() - int(window.width() * ratio),
                max(0, event.globalPos().y() - 20),
            )
        moved = (
            window.start_system_move()
            if hasattr(window, "start_system_move")
            else False
        )
        self._drag_offset = (
            None if moved else event.globalPos() - window.frameGeometry().topLeft()
        )
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._drag_offset is not None and event.buttons() & Qt.LeftButton:
            self.window().move(event.globalPos() - self._drag_offset)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        self._drag_offset = None
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton and self._maximizable:
            self.toggle_maximize()
            return
        super().mouseDoubleClickEvent(event)
