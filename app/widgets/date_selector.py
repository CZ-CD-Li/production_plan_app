"""表格下方的 5 天日期切换按钮。"""

from __future__ import annotations

from datetime import date

from PyQt5.QtCore import QRect, QSize, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QPainter, QPen
from PyQt5.QtWidgets import QAbstractButton, QButtonGroup, QHBoxLayout, QWidget

from app.config import DATE_LABELS, DATE_OFFSETS


class DateButton(QAbstractButton):
    def __init__(self, label: str, day: date, parent=None) -> None:
        super().__init__(parent)
        self.label = label
        self.day = day
        self.setText(f"{label} {day.strftime('%Y-%m-%d')}")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setAttribute(Qt.WA_Hover, True)
        self.setFixedHeight(62)
        self.setMinimumWidth(112)
        self.setMaximumWidth(190)

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(122, 62)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = self.rect().adjusted(1, 1, -1, -1)

        checked = self.isChecked()
        hover = self.underMouse()

        if checked:
            background = QColor("#1A1A1A")
            border = QColor("#1A1A1A")
            title_color = QColor("#FFFFFF")
            date_color = QColor(255, 255, 255, 175)
        elif hover:
            background = QColor("#F2F2F2")
            border = QColor("#C9C9C9")
            title_color = QColor("#111111")
            date_color = QColor("#6B6B6B")
        else:
            background = QColor("#FFFFFF")
            border = QColor("#E3E3E3")
            title_color = QColor("#333333")
            date_color = QColor("#8A8A8A")

        painter.setPen(QPen(border, 1))
        painter.setBrush(background)
        painter.drawRoundedRect(rect, 10, 10)

        title_font = QFont("Microsoft YaHei UI")
        title_font.setPixelSize(13)
        title_font.setWeight(QFont.DemiBold)
        painter.setFont(title_font)
        painter.setPen(title_color)
        painter.drawText(QRect(6, 11, self.width() - 12, 18), Qt.AlignCenter, self.label)

        date_font = QFont("Microsoft YaHei UI")
        date_font.setPixelSize(10)
        painter.setFont(date_font)
        painter.setPen(date_color)
        painter.drawText(
            QRect(6, 34, self.width() - 12, 16),
            Qt.AlignCenter,
            self.day.strftime("%Y-%m-%d"),
        )
        painter.end()

    def enterEvent(self, event) -> None:  # noqa: N802
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self.update()
        super().leaveEvent(event)


class DateSelector(QWidget):
    date_changed = pyqtSignal(object)

    def __init__(self, dates: list[date], parent=None) -> None:
        super().__init__(parent)
        self._buttons: dict[date, DateButton] = {}
        self.current_date: date | None = None

        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(10)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._build_buttons(dates)

    def _build_buttons(self, dates: list[date]) -> None:
        for offset, day in zip(DATE_OFFSETS, dates):
            button = DateButton(DATE_LABELS[offset], day)
            button.clicked.connect(lambda _checked=False, b=button: self._on_selected(b))
            self._group.addButton(button)
            self._layout.addWidget(button, 1)
            self._buttons[day] = button

        if dates:
            today_index = DATE_OFFSETS.index(0) if 0 in DATE_OFFSETS else 0
            self.select(dates[today_index] if len(dates) == len(DATE_OFFSETS) else dates[0])

    def set_dates(self, dates: list[date]) -> None:
        if dates == self.dates():
            return
        self._group.setExclusive(False)
        for button in self._buttons.values():
            self._group.removeButton(button)
            self._layout.removeWidget(button)
            button.setParent(None)
            button.deleteLater()
        self._buttons.clear()
        self._build_buttons(dates)
        self._group.setExclusive(True)

    def dates(self) -> list[date]:
        return list(self._buttons)

    def _on_selected(self, button: DateButton) -> None:
        self.current_date = button.day
        self.date_changed.emit(button.day)

    def select(self, day: date) -> None:
        button = self._buttons.get(day)
        if button is None:
            return
        button.setChecked(True)
        self.current_date = day

    def selected_date(self) -> date | None:
        return self.current_date
