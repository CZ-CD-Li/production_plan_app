"""独立的用户管理小程序：查看已有账号、新增登录用户。

直接运行：
    python build_tools/user_admin.py [--db 路径\\app.db]

打包：
    python -m PyInstaller --noconfirm --clean --onefile --windowed
        --name 用户管理工具 --icon build_tools/app.ico --paths .
        build_tools/user_admin.py
"""

from __future__ import annotations

import json
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.theme import build_stylesheet
from app.utils.icons import app_icon
from app.utils.security import hash_password

APP_TITLE = "生产计划协同管理平台"
TOOL_TITLE = "用户管理工具"
FROZEN = bool(getattr(sys, "frozen", False))
BASE_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "user_admin_config.json"
DEFAULT_ROLES = ("生产计划员", "工艺工程师")
SMOKE_TEST = "--smoke-test" in sys.argv

TOOL_QSS = """
QLabel#columnHint {
    color: #8A8A8A;
    font-size: 11px;
    font-weight: 600;
    padding-left: 12px;
}
QPushButton#removeRowButton {
    background: transparent;
    border: none;
    border-radius: 6px;
    color: #9B9B9B;
    font-size: 16px;
    padding: 0;
}
QPushButton#removeRowButton:hover {
    background: #F0F0F0;
    color: #111111;
}
QCheckBox {
    color: #6B6B6B;
    font-size: 12px;
    spacing: 6px;
}
QCheckBox::indicator {
    width: 14px;
    height: 14px;
    border: 1px solid #C9C9C9;
    border-radius: 3px;
    background: #FFFFFF;
}
QCheckBox::indicator:checked {
    background: #1A1A1A;
    border-color: #1A1A1A;
}
"""


# ---------- 数据库定位 ----------


def _cli_db() -> Path | None:
    args = sys.argv[1:]
    for index, arg in enumerate(args):
        if arg == "--db" and index + 1 < len(args):
            return Path(args[index + 1]).expanduser()
        if arg.startswith("--db="):
            return Path(arg.split("=", 1)[1]).expanduser()
    return None


def _saved_db() -> Path | None:
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    value = str(data.get("db_path", "")).strip()
    return Path(value).expanduser() if value else None


def _shared_db_from_app_config(folder: Path) -> Path | None:
    """主程序 data/local_config.json 里配置的共享数据库路径。"""
    try:
        data = json.loads(
            (folder / "data" / "local_config.json").read_text(encoding="utf-8")
        )
    except (OSError, ValueError):
        return None
    value = str(data.get("db_path", "")).strip()
    return Path(value).expanduser() if value else None


def default_db_path() -> Path | None:
    """按优先级猜测主程序的账号数据库位置。"""
    candidates: list[Path] = []
    cli = _cli_db()
    if cli is not None:
        return cli
    saved = _saved_db()
    if saved is not None:
        candidates.append(saved)
    env = os.environ.get("PRODUCTION_PLAN_DB", "").strip()
    if env:
        candidates.append(Path(env).expanduser())
    shared = _shared_db_from_app_config(BASE_DIR)
    if shared is not None:
        candidates.append(shared)
    candidates.append(BASE_DIR / "data" / "app.db")
    candidates.append(BASE_DIR / APP_TITLE / "data" / "app.db")
    candidates.append(ROOT_DIR / "data" / "app.db")
    for path in candidates:
        if path.is_file():
            return path
    return candidates[0] if candidates else None


class UserStore:
    """账号库的只读写封装（仅操作 users 表）。"""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.conn = sqlite3.connect(str(path), timeout=15, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA busy_timeout=15000")

    def has_users_table(self) -> bool:
        row = self.conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='users'"
        ).fetchone()
        return row is not None

    def users(self) -> list[sqlite3.Row]:
        return list(
            self.conn.execute(
                """
                SELECT employee_no, name, role, created_at
                FROM users ORDER BY employee_no
                """
            ).fetchall()
        )

    def roles(self) -> list[str]:
        return [
            row[0]
            for row in self.conn.execute(
                "SELECT DISTINCT role FROM users WHERE role <> '' ORDER BY role"
            ).fetchall()
        ]

    def existing(self, employee_nos: list[str]) -> set[str]:
        if not employee_nos:
            return set()
        marks = ", ".join("?" for _ in employee_nos)
        rows = self.conn.execute(
            f"SELECT employee_no FROM users WHERE employee_no IN ({marks})",
            tuple(employee_nos),
        ).fetchall()
        return {str(row[0]) for row in rows}

    def add_many(self, records: list[tuple[str, str, str, str]]) -> None:
        """一次性写入多条账号；任一条失败则全部回滚。"""
        created_at = datetime.now().isoformat(timespec="seconds")
        rows = [
            (employee_no, hash_password(password), name, role, created_at)
            for employee_no, name, role, password in records
        ]
        with self.conn:
            self.conn.executemany(
                """
                INSERT INTO users (employee_no, password_hash, name, role, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                rows,
            )

    def close(self) -> None:
        self.conn.close()


class UserAdminWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.store: UserStore | None = None
        self.batch_rows: list[dict] = []
        self.setWindowTitle(f"{TOOL_TITLE} · {APP_TITLE}")
        self.setMinimumSize(1020, 600)
        self.resize(1140, 660)
        self._build_ui()
        initial = default_db_path()
        if initial is not None:
            self._connect(initial, silent=True)

    # ---------- 界面 ----------

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 20, 24, 24)
        outer.setSpacing(14)

        header = QHBoxLayout()
        titles = QVBoxLayout()
        titles.setSpacing(3)
        title = QLabel(TOOL_TITLE)
        title.setObjectName("pageTitle")
        subtitle = QLabel("为生产计划协同管理平台新增登录账号，保存后立即生效，无需重启主程序。")
        subtitle.setObjectName("pageSubtitle")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        header.addLayout(titles)
        header.addStretch(1)
        outer.addLayout(header)

        outer.addWidget(self._build_db_card())
        outer.addLayout(self._build_body(), 1)
        self._reset_batch_rows()

    def _build_db_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        label = QLabel("账号数据库")
        label.setObjectName("fieldLabel")
        self.path_input = QLineEdit()
        self.path_input.setReadOnly(True)
        self.path_input.setPlaceholderText("未找到 app.db，请点击右侧按钮选择")
        self.choose_button = QPushButton("选择数据库…")
        self.choose_button.setObjectName("secondaryButton")
        self.choose_button.clicked.connect(self._choose_db)

        layout.addWidget(label)
        layout.addWidget(self.path_input, 1)
        layout.addWidget(self.choose_button)
        return card

    def _build_body(self) -> QHBoxLayout:
        body = QHBoxLayout()
        body.setSpacing(16)
        body.addWidget(self._build_table_card(), 1)
        body.addWidget(self._build_form_card(), 0)
        return body

    def _build_table_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(0)

        title = QLabel("已有账号")
        title.setObjectName("sectionTitle")
        self.count_pill = QLabel("共 0 个账号")
        self.count_pill.setObjectName("pillMuted")
        title_row = QHBoxLayout()
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(self.count_pill)
        hint = QLabel("数据来自主程序当前使用的账号库，工号即登录账号。")
        hint.setObjectName("sectionHint")
        layout.addLayout(title_row)
        layout.addSpacing(4)
        layout.addWidget(hint)
        layout.addSpacing(14)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(("工号", "姓名", "角色", "创建时间"))
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(42)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setFocusPolicy(Qt.NoFocus)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 100)
        self.table.setColumnWidth(1, 110)
        self.table.setColumnWidth(3, 170)
        layout.addWidget(self.table, 1)
        return card

    def _build_form_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        card.setFixedWidth(480)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(0)

        title = QLabel("新增用户")
        title.setObjectName("sectionTitle")
        hint = QLabel("可一次填写多个账号，全部通过校验后才会写入。")
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        layout.addWidget(title)
        layout.addSpacing(4)
        layout.addWidget(hint)
        layout.addSpacing(14)

        layout.addLayout(self._build_batch_header())
        layout.addSpacing(6)

        self.rows_container = QWidget()
        self.rows_layout = QVBoxLayout(self.rows_container)
        self.rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rows_layout.setSpacing(6)
        self.rows_layout.addStretch(1)
        self.rows_scroll = QScrollArea()
        self.rows_scroll.setWidgetResizable(True)
        self.rows_scroll.setFrameShape(QFrame.NoFrame)
        self.rows_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.rows_scroll.setWidget(self.rows_container)
        layout.addWidget(self.rows_scroll)
        layout.addStretch(1)

        self.status = QLabel("")
        self.status.setObjectName("fieldHint")
        self.status.setWordWrap(True)
        self.status.setMinimumHeight(30)
        layout.addWidget(self.status)
        layout.addSpacing(12)

        divider = QFrame()
        divider.setObjectName("divider")
        divider.setFixedHeight(1)
        layout.addWidget(divider)
        layout.addSpacing(12)

        tools = QHBoxLayout()
        tools.setSpacing(10)
        add_row_button = QPushButton("添加一行")
        add_row_button.setObjectName("secondaryButton")
        add_row_button.clicked.connect(lambda: self._add_batch_row())
        self.show_password = QCheckBox("显示密码")
        self.show_password.toggled.connect(self._toggle_password)
        tools.addWidget(add_row_button)
        tools.addStretch(1)
        tools.addWidget(self.show_password)
        layout.addLayout(tools)
        layout.addSpacing(12)

        self.add_button = QPushButton("全部新增")
        self.add_button.setObjectName("primaryButton")
        self.add_button.setMinimumHeight(42)
        self.add_button.setEnabled(False)
        self.add_button.clicked.connect(self._handle_batch_add)
        layout.addWidget(self.add_button)
        return card

    def _build_batch_header(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(6)
        for text, width in (("工号", 88), ("姓名", 88)):
            label = QLabel(text)
            label.setObjectName("columnHint")
            label.setFixedWidth(width)
            row.addWidget(label)
        role = QLabel("角色")
        role.setObjectName("columnHint")
        row.addWidget(role, 1)
        password = QLabel("密码")
        password.setObjectName("columnHint")
        password.setFixedWidth(100)
        row.addWidget(password)
        row.addSpacing(32)
        return row

    def _reset_batch_rows(self) -> None:
        for entry in self.batch_rows:
            self.rows_layout.removeWidget(entry["widget"])
            entry["widget"].deleteLater()
        self.batch_rows.clear()
        for _ in range(3):
            self._add_batch_row()
        self._update_rows_height()

    def _update_rows_height(self) -> None:
        shown = min(max(1, len(self.batch_rows)), 6)
        row_height = 36
        if self.batch_rows:
            row_height = self.batch_rows[0]["widget"].sizeHint().height()
        self.rows_scroll.setFixedHeight(shown * row_height + (shown - 1) * 6 + 4)

    def _role_options(self) -> list[str]:
        options = list(DEFAULT_ROLES)
        if self.store is not None:
            try:
                for role in self.store.roles():
                    if role and role not in options:
                        options.append(role)
            except sqlite3.Error:
                pass
        return options

    def _add_batch_row(self) -> None:
        first = not self.batch_rows
        employee = QLineEdit()
        employee.setPlaceholderText("1003" if first else "")
        employee.setMaxLength(32)
        employee.setFixedWidth(88)
        name = QLineEdit()
        name.setPlaceholderText("王强" if first else "")
        name.setMaxLength(32)
        name.setFixedWidth(88)
        role = QComboBox()
        role.setEditable(True)
        role.addItems(self._role_options())
        role.setCurrentText(DEFAULT_ROLES[0])
        password = QLineEdit()
        password.setEchoMode(
            QLineEdit.Normal if self.show_password.isChecked() else QLineEdit.Password
        )
        password.setPlaceholderText("密码" if first else "")
        password.setFixedWidth(100)
        remove = QPushButton("×")
        remove.setObjectName("removeRowButton")
        remove.setFixedSize(26, 26)
        remove.setToolTip("删除该行")
        remove.setCursor(Qt.PointingHandCursor)

        container = QWidget()
        row_layout = QHBoxLayout(container)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(6)
        row_layout.addWidget(employee)
        row_layout.addWidget(name)
        row_layout.addWidget(role, 1)
        row_layout.addWidget(password)
        row_layout.addWidget(remove)

        remove.clicked.connect(
            lambda _=False, widget=container: self._remove_batch_row(widget)
        )
        self.rows_layout.insertWidget(self.rows_layout.count() - 1, container)
        self.batch_rows.append(
            {
                "widget": container,
                "employee": employee,
                "name": name,
                "role": role,
                "password": password,
            }
        )
        self._update_rows_height()

    def _remove_batch_row(self, widget: QWidget) -> None:
        for entry in list(self.batch_rows):
            if entry["widget"] is widget:
                self.batch_rows.remove(entry)
                self.rows_layout.removeWidget(widget)
                widget.deleteLater()
                break
        if not self.batch_rows:
            self._add_batch_row()
        self._update_rows_height()

    def _toggle_password(self, checked: bool) -> None:
        mode = QLineEdit.Normal if checked else QLineEdit.Password
        for entry in self.batch_rows:
            entry["password"].setEchoMode(mode)

    # ---------- 交互 ----------

    def _set_status(self, text: str, ok: bool) -> None:
        self.status.setText(text)
        self.status.setObjectName("inlineSuccess" if ok else "inlineError")
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)

    def _choose_db(self) -> None:
        start = self.path_input.text().strip() or str(BASE_DIR)
        chosen, _ = QFileDialog.getOpenFileName(
            self, "选择主程序的账号数据库", start, "SQLite 数据库 (app.db *.db);;所有文件 (*)"
        )
        if chosen:
            self._connect(Path(chosen))

    def _connect(self, path: Path, silent: bool = False) -> bool:
        if not path.is_file():
            if not silent:
                QMessageBox.warning(self, "数据库不存在", f"找不到数据库文件：\n{path}")
            self._set_status(f"未找到数据库：{path}", False)
            self.add_button.setEnabled(False)
            return False
        try:
            store = UserStore(path)
        except sqlite3.Error as exc:
            if not silent:
                QMessageBox.warning(self, "无法打开数据库", f"{path}\n\n{exc}")
            self._set_status(f"无法打开数据库：{exc}", False)
            return False
        if not store.has_users_table():
            store.close()
            if not silent:
                QMessageBox.warning(
                    self,
                    "数据库不符",
                    "所选文件不是本系统的账号数据库（缺少 users 表）。\n\n"
                    "请选择主程序 data 目录下的 app.db。",
                )
            self._set_status("所选文件不是本系统的账号数据库。", False)
            self.add_button.setEnabled(False)
            return False

        if self.store is not None:
            self.store.close()
        self.store = store
        self.path_input.setText(str(path))
        self.path_input.setCursorPosition(0)
        self.add_button.setEnabled(True)
        self._load_table()
        self._load_roles()
        if not SMOKE_TEST:
            self._save_config(path)
        return True

    def _save_config(self, path: Path) -> None:
        try:
            CONFIG_FILE.write_text(
                json.dumps({"db_path": str(path)}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass

    def _load_table(self) -> None:
        if self.store is None:
            self.table.setRowCount(0)
            self.count_pill.setText("共 0 个账号")
            return
        try:
            rows = self.store.users()
        except sqlite3.Error as exc:
            self._set_status(f"读取账号列表失败：{exc}", False)
            return
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            created = str(row["created_at"] or "").replace("T", " ")[:16]
            values = (row["employee_no"], row["name"], row["role"], created)
            for col, text in enumerate(values):
                self.table.setItem(row_index, col, QTableWidgetItem(str(text)))
        self.count_pill.setText(f"共 {len(rows)} 个账号")

    def _load_roles(self) -> None:
        options = self._role_options()
        for entry in self.batch_rows:
            combo = entry["role"]
            current = combo.currentText().strip() or DEFAULT_ROLES[0]
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(options)
            combo.setCurrentText(current)
            combo.blockSignals(False)

    def _collect_batch_rows(self) -> list[dict]:
        records: list[dict] = []
        for index, entry in enumerate(self.batch_rows):
            employee_no = entry["employee"].text().strip()
            person = entry["name"].text().strip()
            secret = entry["password"].text()
            if not (employee_no or person or secret):
                continue
            records.append(
                {
                    "row": index + 1,
                    "employee_no": employee_no,
                    "name": person,
                    "role": entry["role"].currentText().strip() or DEFAULT_ROLES[0],
                    "password": secret,
                }
            )
        return records

    def _handle_batch_add(self) -> None:
        if self.store is None:
            self._set_status("请先选择账号数据库。", False)
            return
        records = self._collect_batch_rows()
        if not records:
            self._set_status("请至少填写一个用户的工号、姓名和密码。", False)
            return

        errors: list[str] = []
        seen: set[str] = set()
        for record in records:
            prefix = f"第 {record['row']} 行"
            employee_no = record["employee_no"]
            if not employee_no:
                errors.append(f"{prefix}：请填写工号")
            elif employee_no in seen:
                errors.append(f"{prefix}：工号 {employee_no} 在本次填写中重复")
            else:
                seen.add(employee_no)
            if not record["name"]:
                errors.append(f"{prefix}：请填写姓名")
            if len(record["password"]) < 6:
                errors.append(f"{prefix}：密码长度不能少于 6 位")

        try:
            existing = self.store.existing(
                [r["employee_no"] for r in records if r["employee_no"]]
            )
        except sqlite3.Error as exc:
            self._set_status(f"读取账号库失败：{exc}", False)
            return
        for record in records:
            if record["employee_no"] in existing:
                errors.append(f"第 {record['row']} 行：工号 {record['employee_no']} 已存在")

        if errors:
            summary = "；".join(errors[:2])
            if len(errors) > 2:
                summary += f" 等 {len(errors)} 处问题"
            self._set_status(f"未保存：{summary}。", False)
            QMessageBox.warning(self, "请修正后重试", "\n".join(errors))
            return

        try:
            self.store.add_many(
                [
                    (r["employee_no"], r["name"], r["role"], r["password"])
                    for r in records
                ]
            )
        except sqlite3.IntegrityError:
            self._set_status("工号已被其他用户占用，本次未写入任何账号，请重试。", False)
            return
        except sqlite3.Error as exc:
            self._set_status(f"新增失败：{exc}", False)
            return

        self._load_table()
        self._load_roles()
        self._reset_batch_rows()
        if self.batch_rows:
            self.batch_rows[0]["employee"].setFocus()
        summary = "、".join(f"{r['employee_no']} {r['name']}" for r in records[:5])
        if len(records) > 5:
            summary += f" 等 {len(records)} 人"
        self._set_status(f"已新增 {len(records)} 个账号：{summary}。", True)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.store is not None:
            self.store.close()
            self.store = None
        super().closeEvent(event)


def main() -> int:
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName(TOOL_TITLE)
    app.setStyle("Fusion")
    app.setStyleSheet(build_stylesheet() + TOOL_QSS)
    app.setWindowIcon(app_icon())

    try:
        window = UserAdminWindow()
    except Exception as exc:  # noqa: BLE001
        QMessageBox.critical(None, "启动失败", f"{type(exc).__name__}: {exc}")
        return 1
    window.show()
    if SMOKE_TEST:
        QTimer.singleShot(1500, app.quit)
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
