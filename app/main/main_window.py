"""主窗口：侧边栏 + 顶栏 + 页面栈。"""

from __future__ import annotations

from datetime import date, datetime, time, timedelta

from PyQt5.QtCore import QEvent, Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QMainWindow,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from app.capp.capp_view import CappQualityPlanView
from app.capp.control_plan_view import ControlPlanView
from app.config import APP_NAME, APP_VERSION
from app.database import Database
from app.main.sidebar import Sidebar
from app.main.topbar import TopBar
from app.models import User
from app.personal.process_document_view import PersonalProcessDocView
from app.personal.product_list_view import PersonalProductListView
from app.personal.tooling_view import PersonalToolingView
from app.personal.tracking_records_view import PersonalTrackingView
from app.services.data_manager import DataManager
from app.settings.settings_page import SettingsPage
from app.spc.spc_view import SpcGenerateView
from app.table_ocr.ocr_view import TableRecognitionView
from app.utils.debug_log import error, info, warn
from app.utils.icons import app_icon
from app.views.customization_view import CustomizationView
from app.views.plan_view import PlanView
from app.views.process_query_view import ProcessQueryView
from app.views.product_owner_view import ProductOwnerView
from app.views.tooling_query_view import ToolingQueryView
from app.views.tracking_query_view import TrackingQueryView
from app.work_instruction.instruction_view import WorkInstructionView
from app.widgets.frameless import FramelessMixin
from app.widgets.title_bar import WindowTitleBar

PAGE_META = {
    "plan": (
        "5天新品计划",
        "按装配 / 焊接工序查看前天至后天的生产计划",
    ),
    "process": (
        "工艺查询",
        "按产品代号、生产线、编制检索并下载工艺文件",
    ),
    "tooling_query": (
        "工装查询",
        "按产品代号、生产线、编制检索并下载工装文件",
    ),
    "owner_query": (
        "产品责任人查询",
        "按产品代号、主机、生产线查询装配与焊接负责人",
    ),
    "tracking_query": (
        "跟线记录查询",
        "正在开发中",
    ),
    "personal_products": (
        "个人-产品列表",
        "查看当前账号负责的产品，并进入工装或工艺文件",
    ),
    "personal_tooling": (
        "个人-工装管理",
        "上传、查看、下载与删除个人工装文件",
    ),
    "personal_process": (
        "个人-工艺管理",
        "上传、查看、下载与删除个人工艺文档",
    ),
    "personal_tracking": (
        "个人-跟线记录",
        "上传、查看、下载与删除个人跟线记录",
    ),
    "custom": (
        "个性化功能定制",
        "按岗位定制工作台入口与提醒",
    ),
    "capp_generate": (
        "转CAPP质量策划",
        "导入控制计划，映射字段并生成 CAPP 质量策划表",
    ),
    "control_plan": (
        "作业指导书转控制计划",
        "解析作业指导书各工序，按主机厂模板生成控制计划",
    ),
    "spc_generate": (
        "SPC生成",
        "导入 SPC 计划，确认数据范围后生成符合判异准则的 SPC 数据",
    ),
    "table_ocr": (
        "表格识图",
        "导入 PDF 图纸，识别表格中的序号、代号、版本、名称、数量列",
    ),
    "work_instruction": (
        "后减作业指导书生成",
        "导入多份图纸，按生产线生成产品代号-作业指导书（初版）",
    ),
    "settings": (
        "系统设置",
        "账户安全、终端连接与计划数据源",
    ),
}


class MainWindow(FramelessMixin, QMainWindow):
    def __init__(self, db: Database, user: User) -> None:
        super().__init__()
        self.enable_frameless(resizable=True)
        self.db = db
        self.user = user
        self.data_manager = DataManager(db, user)
        self._syncing = False

        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon())
        self.setMinimumSize(1180, 720)

        root = QWidget()
        root.setObjectName("windowFrame")
        root.setAttribute(Qt.WA_StyledBackground, True)
        outer_layout = QVBoxLayout(root)
        outer_layout.setContentsMargins(1, 1, 1, 1)
        outer_layout.setSpacing(0)

        self.title_bar = WindowTitleBar(
            APP_NAME, f"内部生产系统 · V{APP_VERSION}"
        )
        outer_layout.addWidget(self.title_bar)

        body = QWidget()
        body.setObjectName("appBody")
        body.setAttribute(Qt.WA_StyledBackground, True)
        root_layout = QHBoxLayout(body)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.sidebar = Sidebar(user)
        root_layout.addWidget(self.sidebar)

        content = QWidget()
        content.setObjectName("contentArea")
        content.setAttribute(Qt.WA_StyledBackground, True)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.topbar = TopBar()
        content_layout.addWidget(self.topbar)

        self.stack = QStackedWidget()
        self.pages: dict[str, QWidget] = {}

        self.plan_view = PlanView(self.data_manager)
        self.process_view = ProcessQueryView(db, user)
        self.tooling_query_view = ToolingQueryView(db, user)
        self.owner_query_view = ProductOwnerView(db, user)
        self.tracking_query_view = TrackingQueryView()
        self.personal_products_view = PersonalProductListView(db, user)
        self.personal_tooling_view = PersonalToolingView(db, user)
        self.personal_process_view = PersonalProcessDocView(db, user)
        self.personal_tracking_view = PersonalTrackingView(db, user)
        self.custom_view = CustomizationView()
        self.capp_view = CappQualityPlanView(db, user)
        self.control_plan_view = ControlPlanView(db, user)
        self.spc_view = SpcGenerateView(db, user)
        self.table_ocr_view = TableRecognitionView(db, user)
        self.work_instruction_view = WorkInstructionView(db, user)
        self.settings_view = SettingsPage(
            db, user, self.data_manager.settings_service
        )
        for key, page in (
            ("plan", self.plan_view),
            ("process", self.process_view),
            ("tooling_query", self.tooling_query_view),
            ("owner_query", self.owner_query_view),
            ("tracking_query", self.tracking_query_view),
            ("personal_products", self.personal_products_view),
            ("personal_tooling", self.personal_tooling_view),
            ("personal_process", self.personal_process_view),
            ("personal_tracking", self.personal_tracking_view),
            ("custom", self.custom_view),
            ("capp_generate", self.capp_view),
            ("control_plan", self.control_plan_view),
            ("spc_generate", self.spc_view),
            ("table_ocr", self.table_ocr_view),
            ("work_instruction", self.work_instruction_view),
            ("settings", self.settings_view),
        ):
            self.pages[key] = page
            self.stack.addWidget(page)
        content_layout.addWidget(self.stack, 1)
        root_layout.addWidget(content, 1)
        outer_layout.addWidget(body, 1)
        self.setCentralWidget(root)

        self.personal_products_view.tooling_requested.connect(
            lambda code, line: self.open_personal_page(
                "personal_tooling", code, line
            )
        )
        self.personal_products_view.process_requested.connect(
            lambda code, line: self.open_personal_page(
                "personal_process", code, line
            )
        )
        self.sidebar.nav_selected.connect(self.show_page)
        self.sidebar.settings_requested.connect(lambda: self.show_page("settings"))
        self.settings_view.back_requested.connect(lambda: self.show_page("plan"))
        self.settings_view.saved.connect(self.refresh_data)
        self.topbar.refresh_requested.connect(self.refresh_data)

        self.show_page("plan")
        self._fit_to_screen()
        QTimer.singleShot(120, self.refresh_data)
        self._schedule_next_daily_refresh()

    def _fit_to_screen(self) -> None:
        screen = QApplication.primaryScreen()
        available = screen.availableGeometry() if screen is not None else None
        width, height = 1400, 860
        if available is not None:
            width = max(1180, min(width, available.width() - 60))
            height = max(720, min(height, available.height() - 60))
        self.resize(width, height)
        if available is not None:
            self.move(
                available.x() + (available.width() - width) // 2,
                available.y() + (available.height() - height) // 2,
            )

    def changeEvent(self, event) -> None:  # noqa: N802
        if event.type() == QEvent.WindowStateChange and hasattr(self, "title_bar"):
            self.title_bar.sync_window_state(self.isMaximized())
        super().changeEvent(event)

    def show_page(self, key: str) -> None:
        page = self.pages.get(key)
        if page is None:
            return
        if key == "settings":
            self.settings_view.reload_values()
            self.sidebar.set_active(None)
            self.topbar.set_refresh_visible(False)
        else:
            self.sidebar.set_active(key)
            self.topbar.set_refresh_visible(key == "plan")
        self.stack.setCurrentWidget(page)
        title, subtitle = PAGE_META[key]
        self.topbar.set_page(title, subtitle)
        if hasattr(page, "on_activated"):
            page.on_activated()

    def open_personal_page(self, key: str, product_code: str, line: str) -> None:
        self.show_page(key)
        page = self.pages.get(key)
        if page is not None and hasattr(page, "select_product"):
            page.select_product(product_code, line)

    def refresh_data(self) -> None:
        if self._syncing:
            warn("[UI 数据同步]", "上一次同步尚未结束，本次请求被忽略")
            return
        self._syncing = True
        self.topbar.set_syncing()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        started = datetime.now()
        info("[UI 数据同步]", f"开始 today={date.today().isoformat()}")
        try:
            QApplication.processEvents()
            outcome = self.data_manager.reload()
            self.plan_view.refresh_all(outcome)
            self.personal_products_view.refresh()
            self.topbar.set_outcome(outcome)
            self.topbar.set_today(self.data_manager.today)
            elapsed = (datetime.now() - started).total_seconds()
            info(
                "[UI 数据同步]",
                f"结束 耗时={elapsed:.2f}s 计划行数={outcome.plan_count} "
                f"项目数={outcome.project_count} 匹配项目={outcome.matched_project_count} "
                f"errors={outcome.errors} notices={outcome.notices}",
            )
        except Exception as exc:  # noqa: BLE001
            error("[UI 数据同步]", f"同步异常：{type(exc).__name__}: {exc}")
            raise
        finally:
            QApplication.restoreOverrideCursor()
            self._syncing = False

    def _schedule_next_daily_refresh(self) -> None:
        now = datetime.now()
        next_day = datetime.combine(now.date() + timedelta(days=1), time.min)
        delay_ms = int((next_day - now).total_seconds() * 1000) + 5000
        QTimer.singleShot(delay_ms, self._on_daily_refresh)

    def _on_daily_refresh(self) -> None:
        self.refresh_data()
        self._schedule_next_daily_refresh()
