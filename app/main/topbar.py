"""内容区顶部栏。"""

from __future__ import annotations

from datetime import date

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app.config import WEEKDAY_LABELS
from app.models import LoadOutcome
from app.utils.icons import line_icon


class TopBar(QWidget):
    refresh_requested = pyqtSignal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("topbar")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setFixedHeight(82)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(28, 0, 24, 0)
        layout.setSpacing(16)

        titles = QVBoxLayout()
        titles.setContentsMargins(0, 0, 0, 0)
        titles.setSpacing(3)
        self.title_label = QLabel("5天新品计划")
        self.title_label.setObjectName("pageTitle")
        self.subtitle_label = QLabel("")
        self.subtitle_label.setObjectName("pageSubtitle")
        titles.addWidget(self.title_label)
        titles.addWidget(self.subtitle_label)
        layout.addLayout(titles)
        layout.addStretch(1)

        self.date_label = QLabel()
        self.date_label.setObjectName("dateChip")
        layout.addWidget(self.date_label, 0, Qt.AlignVCenter)
        self.set_today(date.today())

        status_wrap = QHBoxLayout()
        status_wrap.setContentsMargins(0, 0, 0, 0)
        status_wrap.setSpacing(7)
        self.status_dot = QFrame()
        self.status_dot.setFixedSize(8, 8)
        self.status_dot.setStyleSheet("background:#CFCFCF;border-radius:4px;")
        self.status_label = QLabel("等待数据同步")
        self.status_label.setObjectName("statusText")
        status_wrap.addWidget(self.status_dot)
        status_wrap.addWidget(self.status_label)
        layout.addLayout(status_wrap)

        self.refresh_button = QPushButton("刷新数据")
        self.refresh_button.setObjectName("secondaryButton")
        self.refresh_button.setCursor(Qt.PointingHandCursor)
        self.refresh_button.setIcon(line_icon("refresh", "#333333", 34))
        self.refresh_button.setIconSize(QSize(16, 16))
        self.refresh_button.clicked.connect(self.refresh_requested.emit)
        layout.addWidget(self.refresh_button)

    def set_page(self, title: str, subtitle: str) -> None:
        self.title_label.setText(title)
        self.subtitle_label.setText(subtitle)

    def set_today(self, today: date) -> None:
        self.date_label.setText(
            f"{today.year}年{today.month}月{today.day}日 "
            f"{WEEKDAY_LABELS[today.weekday()]}"
        )

    def set_refresh_visible(self, visible: bool) -> None:
        self.refresh_button.setVisible(visible)

    def set_syncing(self) -> None:
        self.status_dot.setStyleSheet("background:#8C8C8C;border-radius:4px;")
        self.status_label.setText("正在同步数据…")
        self.refresh_button.setEnabled(False)

    def set_outcome(self, outcome: LoadOutcome) -> None:
        self.refresh_button.setEnabled(True)
        if not outcome.errors:
            color = "#1F1F1F"
            text = "数据已同步"
        else:
            color = "#000000"
            text = "数据同步异常"
        if outcome.synced_at is not None:
            text = f"{text} {outcome.synced_at.strftime('%H:%M')}"
        self.status_dot.setStyleSheet(f"background:{color};border-radius:4px;")
        self.status_label.setText(text)
        self.status_label.setStyleSheet(
            "color:#111111;font-weight:600;" if outcome.errors else "color:#5C5C5C;"
        )
        tooltip = "\n".join(outcome.errors or outcome.notices) or text
        self.status_label.setToolTip(tooltip)
        self.status_dot.setToolTip(tooltip)
