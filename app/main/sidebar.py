"""左侧导航栏。"""

from __future__ import annotations

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app.config import APP_VERSION
from app.models import User
from app.utils.icons import line_icon, nav_icon
from app.widgets.avatar import Avatar, ClickableFrame
from app.widgets.slide_button import SlideButton

NAV_ITEMS = (
    ("plan", "5天新品计划", "plan"),
    ("process", "工艺查询", "search"),
    ("tooling_query", "工装查询", "folder"),
    ("owner_query", "产品责任人查询", "user"),
    ("tracking_query", "跟线记录查询", "personal"),
    ("personal", "个人工艺管理", "personal"),
    ("custom", "个性化功能定制", "custom"),
)

PERSONAL_SUB_ITEMS = (
    ("personal_products", "个人-产品列表"),
    ("personal_tooling", "个人-工装管理"),
    ("personal_process", "个人-工艺管理"),
    ("personal_tracking", "个人-跟线记录"),
)

CUSTOM_SUB_ITEMS = (
    ("capp_generate", "转CAPP质量策划"),
    ("control_plan", "作业指导书转控制计划"),
    ("spc_generate", "SPC生成"),
    ("table_ocr", "表格识图"),
)

SUB_ITEMS = {
    "personal": PERSONAL_SUB_ITEMS,
    "custom": CUSTOM_SUB_ITEMS,
}


class Sidebar(QWidget):
    nav_selected = pyqtSignal(str)
    settings_requested = pyqtSignal()

    def __init__(self, user: User, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setFixedWidth(236)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 20, 14, 16)
        layout.setSpacing(0)

        profile = ClickableFrame()
        profile.setObjectName("avatarCard")
        profile.setFixedHeight(72)
        profile.setToolTip("点击进入系统设置")
        profile_layout = QHBoxLayout(profile)
        profile_layout.setContentsMargins(12, 10, 12, 10)
        profile_layout.setSpacing(10)

        avatar = Avatar(user.name, 42)
        profile_layout.addWidget(avatar)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(3)
        name_label = QLabel(user.name)
        name_label.setObjectName("userName")
        meta_label = QLabel(f"{user.role} · {user.employee_no}")
        meta_label.setObjectName("userMeta")
        text_layout.addWidget(name_label)
        text_layout.addWidget(meta_label)
        profile_layout.addLayout(text_layout, 1)

        gear_badge = QFrame()
        gear_badge.setFixedSize(26, 26)
        gear_badge.setStyleSheet("background:#1F1F1F;border-radius:8px;")
        gear_layout = QVBoxLayout(gear_badge)
        gear_layout.setContentsMargins(0, 0, 0, 0)
        gear = QLabel()
        gear.setAlignment(Qt.AlignCenter)
        gear.setPixmap(line_icon("gear", "#B5B5B5", 32).pixmap(15, 15))
        gear.setStyleSheet("background:transparent;")
        gear_layout.addWidget(gear)
        profile_layout.addWidget(gear_badge, 0, Qt.AlignVCenter)
        profile.clicked.connect(self.settings_requested.emit)
        layout.addWidget(profile)
        layout.addSpacing(24)

        section = QLabel("工作台")
        section.setObjectName("navSection")
        section.setContentsMargins(10, 0, 0, 0)
        layout.addWidget(section)
        layout.addSpacing(10)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._buttons: dict[str, QPushButton] = {}
        self._holders: dict[str, QWidget] = {}
        for key, label, icon_kind in NAV_ITEMS:
            button = SlideButton(label, "nav")
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setIcon(nav_icon(icon_kind))
            button.setIconSize(QSize(20, 20))
            target_key = (
                SUB_ITEMS[key][0][0] if key in SUB_ITEMS else key
            )
            button.clicked.connect(
                lambda _checked=False, nav_key=key, first_key=target_key: (
                    self._handle_nav_click(nav_key, first_key)
                )
            )
            self._group.addButton(button)
            layout.addWidget(button)
            layout.addSpacing(4)
            self._buttons[key] = button

            if key in SUB_ITEMS:
                holder = QWidget()
                holder.setObjectName("navSubHolder")
                holder.setAttribute(Qt.WA_StyledBackground, True)
                holder_layout = QVBoxLayout(holder)
                holder_layout.setContentsMargins(0, 0, 0, 0)
                holder_layout.setSpacing(2)
                for sub_key, sub_label in SUB_ITEMS[key]:
                    sub_button = SlideButton(sub_label, "nav")
                    sub_button.setObjectName("navSubButton")
                    sub_button.setCheckable(True)
                    sub_button.clicked.connect(
                        lambda _checked=False, nav_key=sub_key: self.nav_selected.emit(
                            nav_key
                        )
                    )
                    self._group.addButton(sub_button)
                    holder_layout.addWidget(sub_button)
                    self._buttons[sub_key] = sub_button
                holder.hide()
                layout.addWidget(holder)
                self._holders[key] = holder
                layout.addSpacing(4)

        layout.addStretch(1)

        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet("background: rgba(255,255,255,0.07);")
        layout.addWidget(divider)
        layout.addSpacing(14)

        footer = QLabel("计划数据 · 每日自动同步")
        footer.setObjectName("sidebarFooter")
        layout.addWidget(footer)
        layout.addSpacing(6)
        version = QLabel(f"V{APP_VERSION}")
        version.setObjectName("sidebarVersion")
        layout.addWidget(version)

    def _handle_nav_click(self, key: str, target_key: str) -> None:
        """点击一级标题：展开自身二级目录并收起其他目录。"""
        self._expand_group(key)
        self.nav_selected.emit(target_key)

    def _expand_group(self, key: str | None) -> None:
        for group_key, holder in self._holders.items():
            holder.setVisible(group_key == key)

    def set_active(self, key: str | None) -> None:
        if key is None:
            self._expand_group(None)
            self._group.setExclusive(False)
            for button in self._buttons.values():
                button.setChecked(False)
            self._group.setExclusive(True)
            return
        group = next(
            (
                group_key
                for group_key, items in SUB_ITEMS.items()
                if any(sub_key == key for sub_key, _label in items)
            ),
            None,
        )
        self._expand_group(group)
        button = self._buttons.get(key)
        if button is not None:
            button.setChecked(True)
