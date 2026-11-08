"""系统设置页：账户安全、终端 IP、数据源与软件更新。"""

from __future__ import annotations

import ipaddress
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.config import (
    APP_VERSION,
    DB_PATH,
    DB_PATH_ENV,
    resolve_db_path,
    save_local_config,
)
from app.database import Database
from app.models import User
from app.services.auth_service import AuthService
from app.services.settings_service import SettingsService
from app.utils.debug_log import info, warn
from app.utils.icons import line_icon
from app.widgets.avatar import Avatar


class SettingsPage(QWidget):
    saved = pyqtSignal()
    back_requested = pyqtSignal()

    def __init__(
        self,
        db: Database,
        user: User,
        settings_service: SettingsService,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.settings_service = settings_service
        self.auth_service = AuthService(db)
        self._build_ui()
        self.reload_values()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(28, 22, 28, 22)
        content_layout.setSpacing(16)

        inner = QWidget()
        inner.setMinimumWidth(860)
        inner.setMaximumWidth(1040)
        inner.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        inner_layout = QVBoxLayout(inner)
        inner_layout.setContentsMargins(0, 0, 0, 0)
        inner_layout.setSpacing(16)

        inner_layout.addWidget(self._build_account_card())
        inner_layout.addWidget(self._build_terminal_card())
        inner_layout.addWidget(self._build_database_card())
        inner_layout.addWidget(self._build_data_source_card())
        inner_layout.addWidget(self._build_update_card())

        centered = QHBoxLayout()
        centered.setContentsMargins(0, 0, 0, 0)
        centered.addStretch(1)
        centered.addWidget(inner, 20)
        centered.addStretch(1)
        content_layout.addLayout(centered)
        content_layout.addStretch(1)
        scroll.setWidget(content)
        root.addWidget(scroll, 1)

        footer = QFrame()
        footer.setObjectName("settingsFooter")
        footer.setFixedHeight(64)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(28, 10, 28, 10)
        footer_layout.setSpacing(12)

        back_button = QPushButton("返回首页")
        back_button.setObjectName("ghostButton")
        back_button.setCursor(Qt.PointingHandCursor)
        back_button.setIcon(line_icon("arrow-left", "#5C5C5C", 32))
        back_button.setIconSize(QSize(16, 16))
        back_button.clicked.connect(self.back_requested.emit)
        footer_layout.addWidget(back_button)

        self.footer_status = QLabel("")
        self.footer_status.setObjectName("fieldHint")
        footer_layout.addWidget(self.footer_status, 0, Qt.AlignVCenter)
        footer_layout.addStretch(1)

        self.save_button = QPushButton("保存设置")
        self.save_button.setObjectName("primaryButton")
        self.save_button.setCursor(Qt.PointingHandCursor)
        self.save_button.clicked.connect(self._save_settings)
        footer_layout.addWidget(self.save_button)
        root.addWidget(footer)

    def _card(self, title: str, hint: str, icon_kind: str) -> tuple[QFrame, QVBoxLayout]:
        frame = QFrame()
        frame.setObjectName("card")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(22, 18, 22, 20)
        layout.setSpacing(0)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(12)

        badge = QFrame()
        badge.setFixedSize(38, 38)
        badge.setStyleSheet("background:#F0F0F0;border-radius:10px;")
        badge_layout = QVBoxLayout(badge)
        badge_layout.setContentsMargins(0, 0, 0, 0)
        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setPixmap(line_icon(icon_kind, "#1A1A1A", 40).pixmap(20, 20))
        icon_label.setStyleSheet("background:transparent;")
        badge_layout.addWidget(icon_label)
        header.addWidget(badge)

        texts = QVBoxLayout()
        texts.setContentsMargins(0, 0, 0, 0)
        texts.setSpacing(2)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        hint_label = QLabel(hint)
        hint_label.setObjectName("sectionHint")
        hint_label.setWordWrap(True)
        texts.addWidget(title_label)
        texts.addWidget(hint_label)
        header.addLayout(texts, 1)
        layout.addLayout(header)
        layout.addSpacing(16)

        divider = QFrame()
        divider.setObjectName("divider")
        divider.setFixedHeight(1)
        layout.addWidget(divider)
        layout.addSpacing(16)
        return frame, layout

    def _build_account_card(self) -> QFrame:
        frame, layout = self._card(
            "账户与安全",
            "修改当前登录账号的密码，密码修改后立即生效。",
            "user",
        )

        profile_row = QHBoxLayout()
        profile_row.setContentsMargins(0, 0, 0, 0)
        profile_row.setSpacing(12)
        profile_row.addWidget(Avatar(self.user.name, 48))

        profile_text = QVBoxLayout()
        profile_text.setContentsMargins(0, 0, 0, 0)
        profile_text.setSpacing(3)
        name = QLabel(self.user.name)
        name.setObjectName("accountName")
        meta = QLabel(f"{self.user.role} · 工号 {self.user.employee_no}")
        meta.setObjectName("accountMeta")
        profile_text.addWidget(name)
        profile_text.addWidget(meta)
        profile_row.addLayout(profile_text)
        profile_row.addStretch(1)

        status = QLabel("账号状态正常")
        status.setObjectName("pill")
        profile_row.addWidget(status, 0, Qt.AlignVCenter)
        layout.addLayout(profile_row)
        layout.addSpacing(18)

        password_grid = QGridLayout()
        password_grid.setContentsMargins(0, 0, 0, 0)
        password_grid.setHorizontalSpacing(14)
        password_grid.setVerticalSpacing(7)
        fields = (
            ("原密码", "old_password", "请输入当前密码"),
            ("新密码", "new_password", "至少 6 位"),
            ("确认新密码", "confirm_password", "再次输入新密码"),
        )
        for column, (label_text, attr, placeholder) in enumerate(fields):
            label = QLabel(label_text)
            label.setObjectName("fieldLabel")
            password_grid.addWidget(label, 0, column)
            field = QLineEdit()
            field.setEchoMode(QLineEdit.Password)
            field.setPlaceholderText(placeholder)
            field.setFixedHeight(38)
            setattr(self, attr, field)
            password_grid.addWidget(field, 1, column)
            password_grid.setColumnStretch(column, 1)
        layout.addLayout(password_grid)
        layout.addSpacing(10)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(12)
        self.password_message = QLabel("")
        self.password_message.setFixedHeight(18)
        action_row.addWidget(self.password_message)
        action_row.addStretch(1)
        change_button = QPushButton("更新密码")
        change_button.setObjectName("secondaryButton")
        change_button.setCursor(Qt.PointingHandCursor)
        change_button.clicked.connect(self._change_password)
        action_row.addWidget(change_button)
        layout.addLayout(action_row)
        return frame

    def _build_terminal_card(self) -> QFrame:
        frame, layout = self._card(
            "IP连接",
            "配置三个用户IP地址。刷新计划时会先查询这些IP是否已有计算好的数据，"
            "有则直接取用最新的，本机不再重复计算。",
            "gear",
        )
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(7)
        self.ip_combos: list[QComboBox] = []
        for index in range(3):
            label = QLabel(f"终端IP地址{index + 1}")
            label.setObjectName("fieldLabel")
            grid.addWidget(label, 0, index)
            combo = QComboBox()
            combo.setEditable(True)
            combo.setInsertPolicy(QComboBox.NoInsert)
            combo.setFixedHeight(38)
            combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            grid.addWidget(combo, 1, index)
            grid.setColumnStretch(index, 1)
            self.ip_combos.append(combo)
        layout.addLayout(grid)
        layout.addSpacing(10)
        hint = QLabel("三个地址互不校验，可留空；以上设置仅对当前登录账号生效。")
        hint.setObjectName("fieldHint")
        layout.addWidget(hint)
        return frame

    def _build_database_card(self) -> QFrame:
        frame, layout = self._card(
            "共享数据库",
            "多台主机需要共享数据时，把数据库文件放到同一共享目录；修改后需重启程序生效。",
            "gear",
        )

        current = "网络共享（DELETE 日志）" if self.db.is_network else "本机（WAL 日志）"
        mode_label = QLabel(f"当前已连接：{self.db.path} · {current}")
        mode_label.setObjectName("fieldHint")
        mode_label.setWordWrap(True)
        layout.addWidget(mode_label)
        if os.environ.get(DB_PATH_ENV, "").strip():
            env_label = QLabel(
                f"已通过环境变量 {DB_PATH_ENV} 指定共享数据库位置，本页修改不会生效。"
            )
            env_label.setObjectName("fieldHint")
            env_label.setWordWrap(True)
            env_label.setStyleSheet("color:#1A1A1A;font-weight:600;")
            layout.addWidget(env_label)
        layout.addSpacing(12)

        label = QLabel("数据库文件位置")
        label.setObjectName("fieldLabel")
        layout.addWidget(label)
        layout.addSpacing(7)
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        self.db_path_edit = QLineEdit()
        self.db_path_edit.setFixedHeight(38)
        self.db_path_edit.setPlaceholderText(
            "如 \\\\服务器\\共享目录\\app.db；留空则使用本机默认位置"
        )
        self.db_path_edit.setText(str(resolve_db_path()))
        row.addWidget(self.db_path_edit, 1)
        browse = QPushButton("选择文件")
        browse.setObjectName("secondaryButton")
        browse.setCursor(Qt.PointingHandCursor)
        browse.setIcon(line_icon("folder", "#333333", 30))
        browse.setIconSize(QSize(15, 15))
        browse.clicked.connect(self._browse_db_path)
        row.addWidget(browse)
        layout.addLayout(row)
        layout.addSpacing(6)
        hint = QLabel(
            "统一填写 共享数据库文件位置"
        )
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        layout.addSpacing(10)
        lead_hint = QLabel(
            f"Project Lead 使用本机库（不参与共享）：{self.db.lead_path}\n"
        )
        lead_hint.setObjectName("fieldHint")
        lead_hint.setWordWrap(True)
        layout.addWidget(lead_hint)
        return frame

    def _browse_db_path(self) -> None:
        current = Path(self.db_path_edit.text().strip() or str(DB_PATH))
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "选择数据库文件位置",
            str(current),
            "SQLite 数据库 (*.db);;所有文件 (*.*)",
        )
        if selected:
            self.db_path_edit.setText(selected)

    def _build_data_source_card(self) -> QFrame:
        frame, layout = self._card(
            "数据源配置",
            "选择生产计划所在目录与项目分布清单文件，保存后自动重新同步；"
            "目录与文件按当前账号独立保存，不影响其他用户。",
            "folder",
        )

        plan_label = QLabel("生产计划文件目录")
        plan_label.setObjectName("fieldLabel")
        layout.addWidget(plan_label)
        layout.addSpacing(7)
        plan_row = QHBoxLayout()
        plan_row.setContentsMargins(0, 0, 0, 0)
        plan_row.setSpacing(10)
        self.plan_dir_edit = QLineEdit()
        self.plan_dir_edit.setReadOnly(True)
        self.plan_dir_edit.setFixedHeight(38)
        self.plan_dir_edit.setPlaceholderText("请选择生产计划 Excel 所在的上层目录")
        plan_row.addWidget(self.plan_dir_edit, 1)
        plan_browse = QPushButton("选择文件夹")
        plan_browse.setObjectName("secondaryButton")
        plan_browse.setCursor(Qt.PointingHandCursor)
        plan_browse.setIcon(line_icon("folder", "#333333", 30))
        plan_browse.setIconSize(QSize(15, 15))
        plan_browse.clicked.connect(self._browse_plan_dir)
        plan_row.addWidget(plan_browse)
        layout.addLayout(plan_row)
        layout.addSpacing(6)
        plan_hint = QLabel(
            "系统自动在目录中查找“当前年 / 次月或当月 / 文件名含 301”的 Excel，"
            "并读取真实当前日期附近的 5 天数据。"
        )
        plan_hint.setObjectName("fieldHint")
        plan_hint.setWordWrap(True)
        layout.addWidget(plan_hint)
        layout.addSpacing(16)

        project_label = QLabel("项目分布清单文件")
        project_label.setObjectName("fieldLabel")
        layout.addWidget(project_label)
        layout.addSpacing(7)
        project_row = QHBoxLayout()
        project_row.setContentsMargins(0, 0, 0, 0)
        project_row.setSpacing(10)
        self.project_file_edit = QLineEdit()
        self.project_file_edit.setReadOnly(True)
        self.project_file_edit.setFixedHeight(38)
        self.project_file_edit.setPlaceholderText("请选择项目分布清单 Excel 文件")
        project_row.addWidget(self.project_file_edit, 1)
        project_browse = QPushButton("选择文件")
        project_browse.setObjectName("secondaryButton")
        project_browse.setCursor(Qt.PointingHandCursor)
        project_browse.setIcon(line_icon("file", "#333333", 30))
        project_browse.setIconSize(QSize(15, 15))
        project_browse.clicked.connect(self._browse_project_file)
        project_row.addWidget(project_browse)
        layout.addLayout(project_row)
        layout.addSpacing(6)
        project_hint = QLabel(
            " "
        )
        project_hint.setObjectName("fieldHint")
        project_hint.setWordWrap(True)
        layout.addWidget(project_hint)
        return frame

    def _build_update_card(self) -> QFrame:
        frame, layout = self._card(
            "版本与更新",
            "检查当前软件版本，获取最新功能与修复。",
            "refresh",
        )
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(14)

        version_box = QVBoxLayout()
        version_box.setContentsMargins(0, 0, 0, 0)
        version_box.setSpacing(3)
        version = QLabel(f"当前版本 V{APP_VERSION}")
        version.setObjectName("sectionTitle")
        self.sync_label = QLabel("项目清单尚未同步")
        self.sync_label.setObjectName("fieldHint")
        version_box.addWidget(version)
        version_box.addWidget(self.sync_label)
        row.addLayout(version_box)
        row.addStretch(1)

        update_button = QPushButton("软件更新")
        update_button.setObjectName("secondaryButton")
        update_button.setCursor(Qt.PointingHandCursor)
        update_button.setIcon(line_icon("refresh", "#333333", 30))
        update_button.setIconSize(QSize(15, 15))
        update_button.clicked.connect(self._check_update)
        row.addWidget(update_button, 0, Qt.AlignVCenter)
        layout.addLayout(row)
        return frame

    def reload_values(self) -> None:
        values = self.settings_service.get_all(self.user.employee_no)
        ips = [values.get("ip1", ""), values.get("ip2", ""), values.get("ip3", "")]
        options: list[str] = []
        for ip in ips + list(SettingsService.DEFAULT_IPS) + [
            "192.168.0.101",
            "10.0.0.101",
            "127.0.0.1",
        ]:
            if ip and ip not in options:
                options.append(ip)
        for combo, value in zip(self.ip_combos, ips):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(options)
            combo.setCurrentText(value)
            combo.blockSignals(False)

        self.plan_dir_edit.setText(values.get("production_plan_dir", ""))
        self.project_file_edit.setText(values.get("project_lead_path", ""))
        self.db_path_edit.setText(str(resolve_db_path()))

        self._clear_password_fields()
        self.password_message.setText("")
        self.footer_status.setText("")

        synced_at = values.get("project_lead_synced_at", "")
        if synced_at:
            try:
                moment = datetime.fromisoformat(synced_at)
                self.sync_label.setText(
                    f"项目清单最近同步：{moment.strftime('%Y-%m-%d %H:%M')}"
                )
            except ValueError:
                self.sync_label.setText("项目清单已同步")
        else:
            self.sync_label.setText("项目清单尚未同步")

    def _clear_password_fields(self) -> None:
        for field in (self.old_password, self.new_password, self.confirm_password):
            field.clear()

    def _change_password(self) -> None:
        ok, message = self.auth_service.change_password(
            self.user,
            self.old_password.text(),
            self.new_password.text(),
            self.confirm_password.text(),
        )
        if ok:
            self.password_message.setObjectName("inlineSuccess")
            self.password_message.setStyleSheet("color:#333333;")
            self._clear_password_fields()
        else:
            self.password_message.setObjectName("inlineError")
            self.password_message.setStyleSheet("color:#111111;font-weight:600;")
        self.password_message.setText(message)

    def _browse_plan_dir(self) -> None:
        start = self.plan_dir_edit.text().strip() or str(Path.home())
        directory = QFileDialog.getExistingDirectory(
            self, "选择生产计划文件目录", start
        )
        if directory:
            self.plan_dir_edit.setText(directory)
            self._set_footer_status("目录已选择，请点击“保存设置”生效。")

    def _browse_project_file(self) -> None:
        start = self.project_file_edit.text().strip()
        if not start:
            start = str(Path.home())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择项目分布清单",
            start,
            "Excel 文件 (*.xlsx *.xlsm)",
        )
        if file_path:
            self.project_file_edit.setText(file_path)
            self._set_footer_status("文件已选择，请点击“保存设置”生效。")

    def _save_settings(self) -> None:
        ips = [combo.currentText().strip() for combo in self.ip_combos]
        labels = ("1", "2", "3")
        for index, ip in enumerate(ips):
            if not ip:
                continue
            try:
                ipaddress.ip_address(ip)
            except ValueError:
                warn("[UI 设置保存]", f"终端IP地址{labels[index]}格式不正确：{ip!r}")
                self._set_footer_status(
                    f"终端IP地址{labels[index]}格式不正确。", error=True
                )
                self.ip_combos[index].setFocus()
                return

        plan_dir = self.plan_dir_edit.text().strip()
        project_path = self.project_file_edit.text().strip()
        if plan_dir and not Path(plan_dir).is_dir():
            warn("[UI 设置保存]", f"生产计划目录不存在：{plan_dir}")
            QMessageBox.warning(self, "路径无效", "生产计划目录不存在，请重新选择。")
            return
        if project_path and not Path(project_path).is_file():
            warn("[UI 设置保存]", f"项目分布清单文件不存在：{project_path}")
            QMessageBox.warning(self, "路径无效", "项目分布清单文件不存在，请重新选择。")
            return

        db_text = self.db_path_edit.text().strip()
        target_db = Path(db_text) if db_text else DB_PATH
        if not target_db.parent.exists():
            warn("[UI 设置保存]", f"数据库目录不存在：{target_db.parent}")
            QMessageBox.warning(
                self, "路径无效", "数据库所在目录不存在，请重新选择共享位置。"
            )
            return
        probe = target_db.parent / f".db_write_test_{os.getpid()}"
        try:
            probe.write_text("", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            warn("[UI 设置保存]", f"数据库目录不可写：{target_db.parent} {exc}")
            QMessageBox.warning(
                self,
                "目录不可写",
                f"无法写入数据库目录：\n{target_db.parent}\n\n请确认共享权限。",
            )
            return

        db_changed = os.path.normcase(os.path.abspath(target_db)) != os.path.normcase(
            os.path.abspath(self.db.path)
        )

        info(
            "[UI 设置保存]",
            f"IP={ips} 计划目录={plan_dir!r} 项目清单={project_path!r} "
            f"数据库={str(target_db)!r} 变更={db_changed}",
        )
        self.settings_service.set_many(
            {
                "ip1": ips[0],
                "ip2": ips[1],
                "ip3": ips[2],
                "production_plan_dir": plan_dir,
                "project_lead_path": project_path,
            },
            self.user.employee_no,
        )
        if db_changed:
            if not target_db.exists():
                reply = QMessageBox.question(
                    self,
                    "初始化共享数据库",
                    "新位置还没有数据库文件。\n"
                    "是否将当前本机数据库（账号、设置、项目清单、文件登记）复制过去？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes,
                )
                if reply == QMessageBox.Yes:
                    try:
                        dest = sqlite3.connect(str(target_db), timeout=15)
                        try:
                            self.db.conn.backup(dest)
                        finally:
                            dest.close()
                    except sqlite3.Error as exc:
                        warn("[UI 设置保存]", f"复制数据库失败：{exc}")
                        QMessageBox.warning(
                            self,
                            "复制失败",
                            f"无法把当前数据库复制到：\n{target_db}\n\n原因：{exc}",
                        )
                        return
            save_local_config({"db_path": "" if not db_text else str(target_db)})
            self._set_footer_status("设置已保存，数据库位置将在重启后生效。")
            QMessageBox.information(
                self,
                "需要重启",
                "数据库位置已保存。\n请关闭并重新启动程序，新的数据库才会生效。",
            )
            return
        self._set_footer_status("设置已保存。")
        self.saved.emit()

    def _set_footer_status(self, text: str, error: bool = False) -> None:
        self.footer_status.setText(text)
        color = "#111111" if error else "#5C5C5C"
        self.footer_status.setStyleSheet(f"color:{color};font-size:11px;")

    def _check_update(self) -> None:
        QMessageBox.information(
            self,
            "软件更新",
            f"当前版本：V{APP_VERSION}\n\n已是最新版本，无需更新。",
        )
