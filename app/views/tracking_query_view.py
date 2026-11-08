"""跟线记录查询（开发中占位页）。"""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget


class TrackingQueryView(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)

        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(32, 40, 32, 40)
        layout.setSpacing(12)

        title = QLabel("跟线记录查询")
        title.setObjectName("sectionTitle")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        hint = QLabel("正在开发中，敬请期待。")
        hint.setObjectName("sectionHint")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("font-size:16px;")
        layout.addWidget(hint)

        root.addWidget(card)
        root.addStretch(1)
