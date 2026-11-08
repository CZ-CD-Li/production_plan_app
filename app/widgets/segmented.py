"""装配 / 焊接分段切换控件。"""

from __future__ import annotations

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QButtonGroup, QFrame, QHBoxLayout, QPushButton

from app.widgets.slide_button import SlideButton


class ModeSelector(QFrame):
    mode_changed = pyqtSignal(str)

    def __init__(self, modes: tuple[str, ...] | list[str], parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("segmented")
        self.setFixedHeight(42)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(2)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._buttons: dict[str, QPushButton] = {}

        for mode in modes:
            button = SlideButton(mode, "segment")
            button.setObjectName("slideSegmentButton")
            button.setCheckable(True)
            button.setFixedHeight(34)
            button.clicked.connect(
                lambda _checked=False, m=mode: self._on_clicked(m)
            )
            self._group.addButton(button)
            layout.addWidget(button)
            self._buttons[mode] = button

        if modes:
            self._buttons[modes[0]].setChecked(True)

    def _on_clicked(self, mode: str) -> None:
        self.mode_changed.emit(mode)

    def set_mode(self, mode: str) -> None:
        button = self._buttons.get(mode)
        if button is not None:
            button.setChecked(True)

    def current_mode(self) -> str:
        for mode, button in self._buttons.items():
            if button.isChecked():
                return mode
        return next(iter(self._buttons), "")
