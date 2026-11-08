"""只允许选择合法产线代号的联想输入框。"""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QComboBox, QCompleter

from app.services.line_access import normalise_line


class LineCombo(QComboBox):
    """可输入联想，但最终值必须落在允许的产线代号内。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._allowed: tuple[str, ...] = ()
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.lineEdit().setPlaceholderText("输入产线代号，如 C1")
        self.lineEdit().editingFinished.connect(self._canonicalise)

    def set_allowed(self, lines: tuple[str, ...], current: str = "") -> None:
        self._allowed = tuple(lines)
        self.blockSignals(True)
        self.clear()
        self.addItems(self._allowed)
        self.setEditText(self._current_or_empty(current))
        self.blockSignals(False)
        completer = self.completer()
        if completer is not None:
            completer.setCompletionMode(QCompleter.PopupCompletion)
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)

    def set_line(self, value: str) -> bool:
        text = self._current_or_empty(value)
        self.setEditText(text)
        return bool(text)

    def current_line(self) -> str | None:
        """返回规范产线代号；输入为空或不匹配时返回 None。"""
        code = normalise_line(self.currentText())
        if code is None or code not in self._allowed:
            return None
        return code

    def _canonicalise(self) -> None:
        code = self.current_line()
        if code is not None and self.currentText().strip() != code:
            self.setEditText(code)

    def _current_or_empty(self, value: str) -> str:
        code = normalise_line(value)
        if code is not None and code in self._allowed:
            return code
        return ""
