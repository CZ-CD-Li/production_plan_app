"""应用入口。"""

from __future__ import annotations

import sys

# 表格识图的 ONNX 运行时需在创建 QApplication 之前加载，
# 否则在部分 Windows 环境下会因 DLL 加载顺序冲突导致进程崩溃。
try:
    import onnxruntime  # noqa: F401
except Exception:  # noqa: BLE001
    pass

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QMessageBox

from app.config import (
    APP_NAME,
    APP_VERSION,
    ORG_NAME,
    ensure_runtime_dirs,
    resolve_db_path,
    resolve_lead_db_path,
)
from app.database import Database
from app.login.login_window import LoginWindow
from app.main.main_window import MainWindow
from app.theme import build_stylesheet
from app.updater import check_and_prompt
from app.utils import debug_log
from app.utils.debug_log import error, info
from app.utils.icons import app_icon


def main() -> int:
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    ensure_runtime_dirs()
    debug_log.start_session(f"{APP_NAME} v{APP_VERSION} 启动")
    db_path = resolve_db_path()
    info("[APP 启动]", f"共享数据库={db_path}")

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName(ORG_NAME)
    app.setStyle("Fusion")
    app.setStyleSheet(build_stylesheet())
    app.setWindowIcon(app_icon())

    if "--smoke-test" not in sys.argv and not check_and_prompt():
        info("[APP 启动]", "存在新版本但未升级，程序退出")
        return 1

    try:
        db = Database()
    except Exception as exc:  # noqa: BLE001
        error("[APP 启动]", f"数据库打开失败：{type(exc).__name__}: {exc}")
        QMessageBox.critical(
            None,
            "数据库连接失败",
            f"共享数据库：\n{db_path}\n"
            f"本机 Project Lead 库：\n{resolve_lead_db_path()}\n\n"
            f"原因：{exc}\n\n"
            "请确认共享目录可访问、当前账号有读写权限，且本机磁盘未被占用。",
        )
        return 1
    info("[APP 启动]", f"本机 Project Lead 库={db.lead_path}")
    windows: dict[str, object] = {}

    login = LoginWindow(db)
    windows["login"] = login

    def handle_login(user) -> None:
        info("[APP 登录]", f"登录成功 工号={user.employee_no} 姓名={user.name}")
        try:
            main_window = MainWindow(db, user)
        except Exception as exc:  # noqa: BLE001
            error("[APP 登录]", f"主窗口初始化失败：{type(exc).__name__}: {exc}")
            QMessageBox.critical(login, "启动失败", f"系统初始化失败：{exc}")
            login.login_button.setEnabled(True)
            return
        windows["main"] = main_window
        main_window.show()
        login.hide()

    login.login_success.connect(handle_login)
    login.show()

    if "--smoke-test" in sys.argv:
        # 打包验证用：自动登录并打开主界面，随后退出，不进入人工操作
        def _auto_login() -> None:
            login.employee_input.setText("1001")
            login.password_input.setText("123456")
            login._handle_login()

        QTimer.singleShot(200, _auto_login)
        QTimer.singleShot(2200, app.quit)

    try:
        return app.exec_()
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main())
