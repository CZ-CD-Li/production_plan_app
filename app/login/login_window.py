"""登录窗口。"""

from __future__ import annotations

from PyQt5.QtCore import QPoint, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app.config import APP_NAME, APP_VERSION
from app.database import Database
from app.models import User
from app.services.auth_service import AuthService
from app.utils.debug_log import info, warn
from app.utils.icons import app_icon, line_icon
from app.widgets.frameless import FramelessMixin
from app.widgets.title_bar import WindowControls


class _DragPanel(QFrame):
    """品牌面板：按住空白区域可拖动整个无边框窗口。"""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._drag_offset: QPoint | None = None

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() != Qt.LeftButton:
            super().mousePressEvent(event)
            return
        window = self.window()
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


class LoginWindow(FramelessMixin, QWidget):
    login_success = pyqtSignal(object)

    def __init__(self, db: Database) -> None:
        super().__init__()
        self.enable_frameless(resizable=False)
        self.auth = AuthService(db)
        self.setObjectName("loginFrame")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setWindowTitle(f"{APP_NAME} · 登录")
        self.setWindowIcon(app_icon())
        self.setFixedSize(940, 560)

        shell = QVBoxLayout(self)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        card = QFrame()
        card.setObjectName("loginCard")
        card_layout = QHBoxLayout(card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)

        card_layout.addWidget(self._build_brand_panel())
        card_layout.addWidget(self._build_form_panel(), 1)
        shell.addWidget(card)

        self.employee_input.setFocus()
        self.setTabOrder(self.employee_input, self.password_input)
        self.setTabOrder(self.password_input, self.login_button)
        self._center_on_screen()

    def _center_on_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            return
        available = screen.availableGeometry()
        self.move(
            available.x() + (available.width() - self.width()) // 2,
            available.y() + (available.height() - self.height()) // 2,
        )

    def _build_brand_panel(self) -> QWidget:
        panel = _DragPanel()
        panel.setObjectName("brandPanel")
        panel.setFixedWidth(382)
        panel.setAttribute(Qt.WA_StyledBackground, True)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(40, 38, 36, 32)
        layout.setSpacing(0)

        logo = QLabel()
        logo.setPixmap(app_icon().pixmap(58, 58))
        logo.setFixedSize(58, 58)
        layout.addWidget(logo)
        layout.addSpacing(26)

        title = QLabel("生产计划\n协同管理平台")
        title.setObjectName("brandTitle")
        title.setWordWrap(True)
        layout.addWidget(title)
        layout.addSpacing(12)

        subtitle = QLabel("5 天新品计划 · 项目分布清单 · 工艺协同")
        subtitle.setObjectName("brandSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)
        layout.addSpacing(34)

        features = (
            "按装配 / 焊接分线查看五日生产计划",
        )
        for feature in features:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(10)
            dot = QFrame()
            dot.setFixedSize(5, 5)
            dot.setStyleSheet("background:#8C8C8C;border-radius:2px;")
            row.addWidget(dot, 0, Qt.AlignTop)
            label = QLabel(feature)
            label.setObjectName("brandFeature")
            label.setWordWrap(True)
            row.addWidget(label, 1)
            layout.addLayout(row)
            layout.addSpacing(15)

        layout.addStretch(1)
        badge = QLabel("301 计划 · 项目清单联动")
        badge.setObjectName("brandBadge")
        badge.setMaximumWidth(190)
        layout.addWidget(badge, 0, Qt.AlignLeft)
        layout.addSpacing(16)

        footer = QLabel(f"V{APP_VERSION} · 内部生产系统")
        footer.setObjectName("brandFooter")
        layout.addWidget(footer)
        return panel

    def _build_form_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(52, 12, 24, 30)
        layout.setSpacing(0)

        controls_row = QHBoxLayout()
        controls_row.setContentsMargins(0, 0, 0, 0)
        controls_row.addStretch(1)
        self.controls = WindowControls(theme="light", maximizable=False)
        controls_row.addWidget(self.controls)
        layout.addLayout(controls_row)
        layout.addSpacing(8)

        title = QLabel("欢迎回来")
        title.setObjectName("loginTitle")
        layout.addWidget(title)
        layout.addSpacing(8)
        subtitle = QLabel("使用工号与密码登录，进入生产计划工作台")
        subtitle.setObjectName("loginSubtitle")
        layout.addWidget(subtitle)
        layout.addSpacing(26)

        employee_label = QLabel("用户工号")
        employee_label.setObjectName("fieldLabel")
        layout.addWidget(employee_label)
        layout.addSpacing(8)
        self.employee_input = QLineEdit()
        self.employee_input.setObjectName("loginInput")
        self.employee_input.setPlaceholderText("请输入您的工号")
        self.employee_input.setFixedHeight(46)
        self.employee_input.addAction(
            line_icon("user", "#9B9B9B", 36), QLineEdit.LeadingPosition
        )
        self.employee_input.returnPressed.connect(lambda: self.password_input.setFocus())
        self.employee_input.textChanged.connect(self._clear_error)
        layout.addWidget(self.employee_input)
        layout.addSpacing(16)

        password_label = QLabel("密码")
        password_label.setObjectName("fieldLabel")
        layout.addWidget(password_label)
        layout.addSpacing(8)
        self.password_input = QLineEdit()
        self.password_input.setObjectName("loginInput")
        self.password_input.setPlaceholderText("请输入登录密码")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setFixedHeight(46)
        self.password_input.addAction(
            line_icon("lock", "#9B9B9B", 36), QLineEdit.LeadingPosition
        )
        self.password_input.returnPressed.connect(self._handle_login)
        self.password_input.textChanged.connect(self._clear_error)
        layout.addWidget(self.password_input)
        layout.addSpacing(8)

        self.error_label = QLabel("")
        self.error_label.setObjectName("loginError")
        self.error_label.setFixedHeight(18)
        layout.addWidget(self.error_label)
        layout.addSpacing(8)

        self.login_button = QPushButton("登 录")
        self.login_button.setObjectName("loginButton")
        self.login_button.setCursor(Qt.PointingHandCursor)
        self.login_button.clicked.connect(self._handle_login)
        layout.addWidget(self.login_button)

        layout.addStretch(1)

        hint = QFrame()
        hint.setObjectName("loginHintBox")
        hint_layout = QVBoxLayout(hint)
        hint_layout.setContentsMargins(14, 10, 14, 10)
        hint_label = QLabel(
            "<span style='color:#808080'>初始账号</span>&nbsp;&nbsp;"
            "<b style='color:#333333'>1001 / 123456</b><br>"
            "<span style='color:#9B9B9B;font-size:11px'>"
            "登录后请在“系统设置”中修改初始密码</span>"
        )
        hint_label.setObjectName("loginHint")
        hint_label.setTextFormat(Qt.RichText)
        hint_layout.addWidget(hint_label)
        layout.addWidget(hint)
        return panel

    def _clear_error(self) -> None:
        if self.error_label.text():
            self.error_label.setText("")

    def _handle_login(self) -> None:
        employee_no = self.employee_input.text().strip()
        password = self.password_input.text()
        if not employee_no or not password:
            self._show_error("请填写用户工号与密码。")
            return

        user = self.auth.authenticate(employee_no, password)
        if user is None:
            warn("[APP 登录]", f"登录失败 工号={employee_no}")
            self._show_error("工号或密码不正确，请重新输入。")
            self.password_input.clear()
            self.password_input.setFocus()
            return

        info("[APP 登录]", f"校验通过 工号={employee_no}")
        self.login_button.setEnabled(False)
        self.login_success.emit(user)

    def _show_error(self, message: str) -> None:
        self.error_label.setText(message)
