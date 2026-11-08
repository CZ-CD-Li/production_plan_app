"""无边框窗口支持：标题栏拖动、窗口边缘缩放。"""

from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes

from PyQt5.QtCore import QPoint, Qt

WM_NCHITTEST = 0x0084
HTLEFT = 10
HTRIGHT = 11
HTTOP = 12
HTTOPLEFT = 13
HTTOPRIGHT = 14
HTBOTTOM = 15
HTBOTTOMLEFT = 16
HTBOTTOMRIGHT = 17

RESIZE_BORDER = 7


class _MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", ctypes.c_size_t),
        ("lParam", ctypes.c_ssize_t),
        ("time", wintypes.DWORD),
        ("pt", wintypes.POINT),
    ]


class FramelessMixin:
    """混入位置必须放在窗口类之前：class X(FramelessMixin, QMainWindow)。"""

    _resizable = True

    def enable_frameless(self, *, resizable: bool = True) -> None:
        self._resizable = resizable
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setAttribute(Qt.WA_Hover, True)

    def start_system_move(self) -> bool:
        """交给系统处理拖动，避免自绘标题栏拖动时卡顿。"""
        handle = self.windowHandle()
        if handle is None:
            return False
        try:
            return bool(handle.startSystemMove())
        except AttributeError:
            return False

    def nativeEvent(self, event_type, message):  # noqa: N802
        if (
            self._resizable
            and sys.platform == "win32"
            and event_type == b"windows_generic_MSG"
            and not self.isMaximized()
            and not self.isFullScreen()
        ):
            try:
                msg = _MSG.from_address(int(message))
            except (TypeError, ValueError):
                msg = None
            if msg is not None and msg.message == WM_NCHITTEST:
                x = ctypes.c_short(msg.lParam & 0xFFFF).value
                y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
                pos = self.mapFromGlobal(QPoint(x, y))
                width, height = self.width(), self.height()
                left = pos.x() < RESIZE_BORDER
                right = pos.x() > width - RESIZE_BORDER
                top = pos.y() < RESIZE_BORDER
                bottom = pos.y() > height - RESIZE_BORDER
                if top and left:
                    return True, HTTOPLEFT
                if top and right:
                    return True, HTTOPRIGHT
                if bottom and left:
                    return True, HTBOTTOMLEFT
                if bottom and right:
                    return True, HTBOTTOMRIGHT
                if left:
                    return True, HTLEFT
                if right:
                    return True, HTRIGHT
                if top:
                    return True, HTTOP
                if bottom:
                    return True, HTBOTTOM
        return super().nativeEvent(event_type, message)
