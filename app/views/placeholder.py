"""功能分区占位页，供尚未配置数据的模块展示。"""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QVBoxLayout,
    QWidget,
)

from app.utils.icons import line_icon


class PlaceholderPage(QWidget):
    def __init__(
        self,
        title: str,
        description: str,
        features: tuple[str, ...] | list[str],
        icon_kind: str,
        parent=None,
    ) -> None:
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(0)

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(44, 40, 44, 40)
        card_layout.setSpacing(0)
        card_layout.addStretch(1)

        icon_wrap = QHBoxLayout()
        icon_wrap.addStretch(1)
        icon_frame = QFrame()
        icon_frame.setFixedSize(76, 76)
        icon_frame.setStyleSheet("background:#F0F0F0;border-radius:38px;")
        icon_layout = QVBoxLayout(icon_frame)
        icon_layout.setContentsMargins(0, 0, 0, 0)
        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setPixmap(
            line_icon(icon_kind, "#4A4A4A", 96).pixmap(42, 42)
        )
        icon_label.setStyleSheet("background:transparent;")
        icon_layout.addWidget(icon_label)
        icon_wrap.addWidget(icon_frame)
        icon_wrap.addStretch(1)
        card_layout.addLayout(icon_wrap)
        card_layout.addSpacing(22)

        tag_wrap = QHBoxLayout()
        tag_wrap.addStretch(1)
        tag = QLabel("模块规划中")
        tag.setObjectName("placeholderTag")
        tag_wrap.addWidget(tag)
        tag_wrap.addStretch(1)
        card_layout.addLayout(tag_wrap)
        card_layout.addSpacing(12)

        title_label = QLabel(title)
        title_label.setObjectName("placeholderTitle")
        title_label.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(title_label)
        card_layout.addSpacing(8)

        description_label = QLabel(description)
        description_label.setObjectName("placeholderText")
        description_label.setWordWrap(True)
        description_label.setAlignment(Qt.AlignCenter)
        description_label.setMaximumWidth(560)
        description_row = QHBoxLayout()
        description_row.addStretch(1)
        description_row.addWidget(description_label)
        description_row.addStretch(1)
        card_layout.addLayout(description_row)
        card_layout.addSpacing(28)

        feature_holder = QWidget()
        feature_layout = QVBoxLayout(feature_holder)
        feature_layout.setContentsMargins(0, 0, 0, 0)
        feature_layout.setSpacing(12)
        for feature in features:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(10)
            dot = QFrame()
            dot.setFixedSize(5, 5)
            dot.setStyleSheet("background:#C4C4C4;border-radius:2px;")
            row.addWidget(dot, 0, Qt.AlignTop)
            label = QLabel(feature)
            label.setObjectName("placeholderFeature")
            row.addWidget(label, 1)
            feature_layout.addLayout(row)
        feature_row = QHBoxLayout()
        feature_row.addStretch(1)
        feature_row.addWidget(feature_holder)
        feature_row.addStretch(1)
        card_layout.addLayout(feature_row)
        card_layout.addStretch(1)

        root.addWidget(card, 1)
