"""5 天生产计划展示页（仅负责展示与按日导出，数据由 DataManager 提供）。"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font as ExcelFont, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.config import MODES, TABLE_HEADERS
from app.models import LoadOutcome, PlanRow
from app.services.data_manager import DataManager
from app.utils.debug_log import detail, info
from app.widgets.date_selector import DateSelector
from app.widgets.segmented import ModeSelector


def write_plan_workbook(
    rows: list[PlanRow], mode: str, day: date, output_path: str | Path
) -> None:
    """把指定日期、指定工序的计划表格写成 Excel。"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = f"{mode}{day.strftime('%m%d')}"

    last_column = get_column_letter(len(TABLE_HEADERS))
    worksheet.merge_cells(f"A1:{last_column}1")
    title_cell = worksheet["A1"]
    title_cell.value = f"{day.strftime('%Y-%m-%d')} {mode}生产计划"
    title_cell.font = ExcelFont(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 30

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = ExcelFont(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D9E1F2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for column_index, header in enumerate(TABLE_HEADERS, start=1):
        cell = worksheet.cell(2, column_index, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    worksheet.row_dimensions[2].height = 26

    for row_index, row in enumerate(rows, start=3):
        values = (
            row.production_line,
            row.host,
            row.product_code,
            row.quantity,
            row.owner or "—",
            row.remark or "—",
        )
        for column_index, value in enumerate(values, start=1):
            cell = worksheet.cell(row_index, column_index, value)
            cell.border = border
            cell.alignment = (
                Alignment(horizontal="left", vertical="center")
                if column_index == len(TABLE_HEADERS)
                else Alignment(horizontal="center", vertical="center")
            )

    widths = (14, 18, 18, 12, 12, 30)
    for column_index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    workbook.save(str(output_path))


class PlanView(QWidget):
    def __init__(self, data_manager: DataManager, parent=None) -> None:
        super().__init__(parent)
        self.data_manager = data_manager
        self.current_mode = MODES[0]
        self.current_date: date = date.today()

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(14)

        self.banner = QFrame()
        self.banner.setObjectName("noticeBanner")
        self.banner.setVisible(False)
        banner_layout = QHBoxLayout(self.banner)
        banner_layout.setContentsMargins(14, 10, 14, 10)
        banner_layout.setSpacing(10)
        self.banner_dot = QFrame()
        self.banner_dot.setFixedSize(7, 7)
        self.banner_dot.setStyleSheet("background:#8C8C8C;border-radius:3px;")
        banner_layout.addWidget(self.banner_dot, 0, Qt.AlignTop)
        self.banner_label = QLabel("")
        self.banner_label.setObjectName("noticeText")
        self.banner_label.setWordWrap(True)
        banner_layout.addWidget(self.banner_label, 1)
        root.addWidget(self.banner)

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(22, 20, 22, 18)
        card_layout.setSpacing(14)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(14)
        self.mode_selector = ModeSelector(MODES)
        self.mode_selector.mode_changed.connect(self._on_mode_changed)
        header_row.addWidget(self.mode_selector)
        header_row.addStretch(1)

        self.meta_label = QLabel("等待数据同步")
        self.meta_label.setObjectName("tableMeta")
        header_row.addWidget(self.meta_label, 0, Qt.AlignVCenter)

        self.count_label = QLabel("共 0 条")
        self.count_label.setObjectName("countPill")
        header_row.addWidget(self.count_label, 0, Qt.AlignVCenter)

        self.export_button = QPushButton("导出")
        self.export_button.setObjectName("secondaryButton")
        self.export_button.setCursor(Qt.PointingHandCursor)
        self.export_button.setFixedWidth(96)
        self.export_button.clicked.connect(self._export_current_day)
        header_row.addWidget(self.export_button, 0, Qt.AlignVCenter)
        card_layout.addLayout(header_row)

        self.table = QTableWidget(0, len(TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(list(TABLE_HEADERS))
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(44)
        self.table.horizontalHeader().setFixedHeight(44)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setMinimumHeight(340)

        header = self.table.horizontalHeader()
        widths = (128, 150, 160, 110, 110)
        for index, width in enumerate(widths):
            self.table.setColumnWidth(index, width)
            header.setSectionResizeMode(index, QHeaderView.Fixed)
            item = self.table.horizontalHeaderItem(index)
            if item is not None:
                item.setTextAlignment(Qt.AlignCenter)
        header.setSectionResizeMode(len(TABLE_HEADERS) - 1, QHeaderView.Stretch)
        remark_header = self.table.horizontalHeaderItem(len(TABLE_HEADERS) - 1)
        if remark_header is not None:
            remark_header.setTextAlignment(Qt.AlignCenter)
        card_layout.addWidget(self.table, 1)

        self.empty_label = QLabel("所选工序与日期暂无生产计划")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.empty_label.setStyleSheet(
            "background:transparent;color:#9A9A9A;font-size:13px;"
        )
        self.empty_label.setParent(self.table.viewport())
        self.empty_label.hide()

        footer_row = QHBoxLayout()
        footer_row.setContentsMargins(0, 2, 0, 0)
        footer_row.setSpacing(14)
        date_label = QLabel("计划日期")
        date_label.setObjectName("fieldLabel")
        footer_row.addWidget(date_label, 0, Qt.AlignVCenter)

        self.date_selector = DateSelector(self.data_manager.dates())
        self.date_selector.date_changed.connect(self._on_date_changed)
        self.current_date = self.date_selector.selected_date() or date.today()
        footer_row.addWidget(self.date_selector, 1)
        card_layout.addLayout(footer_row)

        root.addWidget(card, 1)
        self.refresh_table()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self._position_empty_label()

    def _position_empty_label(self) -> None:
        if hasattr(self, "empty_label"):
            self.empty_label.setGeometry(self.table.viewport().rect())

    def _on_mode_changed(self, mode: str) -> None:
        self.current_mode = mode
        info("[UI 工序切换]", f"mode={mode}")
        self.refresh_table()

    def _on_date_changed(self, day: date) -> None:
        self.current_date = day
        info("[UI 日期切换]", f"date={day.isoformat()}")
        self.refresh_table()

    def apply_outcome(self, outcome: LoadOutcome) -> None:
        messages = outcome.errors or outcome.notices
        if not messages:
            self.banner.setVisible(False)
            self.banner_label.clear()
            return

        is_error = bool(outcome.errors)
        self.banner.setObjectName("noticeBannerError" if is_error else "noticeBanner")
        self.banner_dot.setStyleSheet(
            "background:#111111;border-radius:3px;"
            if is_error
            else "background:#8C8C8C;border-radius:3px;"
        )
        self.banner_label.setProperty("level", "error" if is_error else "info")
        prefix = "数据同步异常：" if is_error else "提示："
        self.banner_label.setText(prefix + "\n".join(messages))
        self.banner.style().unpolish(self.banner)
        self.banner.style().polish(self.banner)
        self.banner_label.style().unpolish(self.banner_label)
        self.banner_label.style().polish(self.banner_label)
        self.banner.setVisible(True)

    def _export_current_day(self) -> None:
        rows = self.data_manager.rows_for(self.current_mode, self.current_date)
        if not rows:
            QMessageBox.warning(self, "导出失败", "当前日期暂无计划数据。")
            return

        default_name = (
            f"新品计划-{self.current_mode}-{self.current_date.strftime('%Y%m%d')}.xlsx"
        )
        target, _ = QFileDialog.getSaveFileName(
            self, "导出当日计划", default_name, "Excel 文件 (*.xlsx)"
        )
        if not target:
            return
        if not target.lower().endswith(".xlsx"):
            target += ".xlsx"

        try:
            write_plan_workbook(rows, self.current_mode, self.current_date, target)
        except OSError as exc:
            QMessageBox.warning(self, "导出失败", f"无法写入文件：{exc}")
            return

        info(
            "[UI 计划导出]",
            f"mode={self.current_mode} date={self.current_date.isoformat()} "
            f"rows={len(rows)} path={target}",
        )
        QMessageBox.information(
            self, "导出成功", f"已导出 {len(rows)} 条计划数据：\n{target}"
        )

    def refresh_table(self) -> None:
        rows = self.data_manager.rows_for(self.current_mode, self.current_date)
        self.export_button.setEnabled(bool(rows))
        self.export_button.setToolTip(
            f"导出 {self.current_date.isoformat()} 的{self.current_mode}计划"
            if rows
            else "当前日期暂无数据可导出"
        )
        detail(
            "[UI 表格刷新]",
            f"mode={self.current_mode} date={self.current_date.isoformat()} rows={len(rows)}",
        )
        if not rows:
            available = {
                day.isoformat(): len(items)
                for day, items in self.data_manager.rows_by_mode_date.get(
                    self.current_mode, {}
                ).items()
            }
            info("[UI 表格刷新]", f"当前组合无数据，{self.current_mode} 各日期行数={available}")
        self.table.setRowCount(0)
        self.table.setRowCount(len(rows))

        for row_index, row in enumerate(rows):
            values = (
                row.production_line,
                row.host,
                row.product_code,
                row.quantity,
                row.owner or "—",
                row.remark or "—",
            )
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column_index == 5:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignCenter)
                item.setToolTip(value)
                if column_index == 2:
                    font = QFont("Microsoft YaHei UI")
                    font.setPixelSize(13)
                    font.setWeight(QFont.DemiBold)
                    item.setFont(font)
                    item.setForeground(Qt.black)
                self.table.setItem(row_index, column_index, item)

        self.count_label.setText(f"共 {len(rows)} 条")
        if self.data_manager.plan_result.file_path:
            plan_name = self.data_manager.plan_result.file_path.replace("\\", "/").split("/")[-1]
            sheet = self.data_manager.plan_result.sheet_name or "—"
            self.meta_label.setText(f"{plan_name} · 工作表 {sheet}")
            self.meta_label.setToolTip(
                f"{self.data_manager.plan_result.file_path}\n工作表：{sheet}"
            )
        else:
            self.meta_label.setText("尚未加载生产计划数据")
            self.meta_label.setToolTip("")

        self.empty_label.setVisible(not rows)
        self._position_empty_label()
        self.table.setVisible(True)

    def refresh_all(self, outcome: LoadOutcome) -> None:
        dates = self.data_manager.dates()
        info(
            "[UI 全量刷新]",
            f"dates={[d.isoformat() for d in dates]} "
            f"plan_count={outcome.plan_count} errors={len(outcome.errors)} "
            f"notices={len(outcome.notices)}",
        )
        if dates != self.date_selector.dates():
            self.date_selector.set_dates(dates)
            self.current_date = self.date_selector.selected_date() or date.today()
        self.apply_outcome(outcome)
        self.refresh_table()
