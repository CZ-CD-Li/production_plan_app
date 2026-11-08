"""个人文件列表：生产线 → 产品代号 → 文件，支持下载与删除。"""

from __future__ import annotations

import shutil
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.models import User
from app.services.personal_file_service import (
    CATEGORY_CONFIG,
    PersonalFileRecord,
    PersonalFileService,
)


class PersonalFileListPanel(QFrame):
    def __init__(self, db, user: User, category: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("card")
        self.user = user
        self.category = category
        self.config = CATEGORY_CONFIG[category]
        self.service = PersonalFileService(db)
        self._records: list[PersonalFileRecord] = []
        self._loading = False

        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(12)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel(self.config["list_title"])
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.count_label = QLabel("共 0 个")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        root.addLayout(header)

        panes = QHBoxLayout()
        panes.setContentsMargins(0, 0, 0, 0)
        panes.setSpacing(14)

        line_pane = self._pane("生产线")
        self.line_list = QListWidget()
        self.line_list.setFixedWidth(150)
        self.line_list.currentItemChanged.connect(self._on_line_changed)
        line_pane.addWidget(self.line_list, 1)
        panes.addLayout(line_pane, 0)

        product_pane = self._pane("产品代号")
        self.product_list = QListWidget()
        self.product_list.setFixedWidth(170)
        self.product_list.currentItemChanged.connect(self._on_product_changed)
        product_pane.addWidget(self.product_list, 1)
        panes.addLayout(product_pane, 0)

        file_pane = self._pane("文件列表")
        self.file_table = QTableWidget(0, 5)
        self.file_table.setHorizontalHeaderLabels(
            ["名称", "工序", "文件", "上传时间", "操作"]
        )
        self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.file_table.setAlternatingRowColors(True)
        self.file_table.setShowGrid(False)
        self.file_table.setWordWrap(False)
        self.file_table.setFocusPolicy(Qt.NoFocus)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.verticalHeader().setDefaultSectionSize(40)
        self.file_table.horizontalHeader().setFixedHeight(38)
        self.file_table.horizontalHeader().setHighlightSections(False)
        header_view = self.file_table.horizontalHeader()
        self.file_table.setColumnWidth(0, 140)
        self.file_table.setColumnWidth(1, 130)
        header_view.setSectionResizeMode(0, QHeaderView.Fixed)
        header_view.setSectionResizeMode(1, QHeaderView.Fixed)
        header_view.setSectionResizeMode(2, QHeaderView.Stretch)
        self.file_table.setColumnWidth(3, 150)
        self.file_table.setColumnWidth(4, 122)
        file_pane.addWidget(self.file_table, 1)
        panes.addLayout(file_pane, 1)

        root.addLayout(panes, 1)
        self.refresh()

    @staticmethod
    def _pane(title: str) -> QVBoxLayout:
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        label = QLabel(title)
        label.setObjectName("fieldLabel")
        layout.addWidget(label)
        return layout

    def on_activated(self) -> None:
        self.refresh()

    def select_context(self, product_code: str, line: str = "") -> None:
        """从产品列表跳转时，同步选中对应生产线与产品。"""
        self.refresh()
        lines = [self.line_list.item(i).text() for i in range(self.line_list.count())]
        target_line = line if line in lines else self._find_line_of(product_code, lines)
        if not target_line:
            return
        self._loading = True
        self._select_text(self.line_list, target_line)
        self._loading = False
        self._load_products(target_line, product_code)

    def _find_line_of(self, product_code: str, lines: list[str]) -> str:
        for line in lines:
            products = self.service.list_products(
                self.category, self.user.employee_no, line
            )
            if product_code in products:
                return line
        return ""

    def refresh(self) -> None:
        selected_line = self._current_text(self.line_list)
        selected_product = self._current_text(self.product_list)
        lines = self.service.list_lines(self.category, self.user.employee_no)

        self._loading = True
        self.line_list.clear()
        for line in lines:
            self.line_list.addItem(QListWidgetItem(line))
        self._loading = False

        target_line = selected_line if selected_line in lines else (lines[0] if lines else "")
        if target_line:
            self._loading = True
            self._select_text(self.line_list, target_line)
            self._loading = False
            self._load_products(target_line, selected_product)
        else:
            self._loading = True
            self.product_list.clear()
            self._loading = False
            self._records = []
            self._render_files()

    def _load_products(self, line: str, preferred: str = "") -> None:
        products = self.service.list_products(self.category, self.user.employee_no, line)
        self._loading = True
        self.product_list.clear()
        for code in products:
            self.product_list.addItem(QListWidgetItem(code))
        self._loading = False

        target = preferred if preferred in products else (products[0] if products else "")
        if target:
            self._loading = True
            self._select_text(self.product_list, target)
            self._loading = False
            self._load_files(line, target)
        else:
            self._records = []
            self._render_files()

    def _load_files(self, line: str, product_code: str) -> None:
        self._records = self.service.list_files(
            self.category, self.user.employee_no, line, product_code
        )
        self._render_files()

    def _render_files(self) -> None:
        self.file_table.setRowCount(0)
        self.file_table.setRowCount(len(self._records))
        for row_index, record in enumerate(self._records):
            values = (
                record.name,
                record.process,
                record.file_name,
                record.uploaded_at.replace("T", " ")[:16],
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if column == 2
                    else Qt.AlignCenter
                )
                item.setToolTip(record.file_path if column == 2 else value)
                self.file_table.setItem(row_index, column, item)
            self.file_table.setCellWidget(
                row_index, 4, self._build_actions(record)
            )
        self.count_label.setText(f"共 {len(self._records)} 个")

    def _build_actions(self, record: PersonalFileRecord) -> QWidget:
        holder = QWidget()
        holder.setStyleSheet("background:transparent;")
        layout = QHBoxLayout(holder)
        layout.setContentsMargins(4, 0, 4, 0)
        layout.setSpacing(2)

        download = QPushButton("下载")
        download.setObjectName("linkButton")
        download.setCursor(Qt.PointingHandCursor)
        download.clicked.connect(lambda: self._download(record))
        layout.addWidget(download)

        delete = QPushButton("删除")
        delete.setObjectName("dangerLinkButton")
        delete.setCursor(Qt.PointingHandCursor)
        delete.clicked.connect(lambda: self._delete(record))
        layout.addWidget(delete)
        return holder

    def _download(self, record: PersonalFileRecord) -> None:
        source = Path(record.file_path)
        if not source.is_file():
            QMessageBox.warning(
                self,
                "下载失败",
                f"文件不存在或已无法访问：\n{record.file_path}",
            )
            return
        target, _ = QFileDialog.getSaveFileName(
            self, "保存文件", record.file_name, "所有文件 (*.*)"
        )
        if not target:
            return
        try:
            shutil.copy2(source, target)
        except OSError as exc:
            QMessageBox.warning(self, "下载失败", f"无法保存文件：{exc}")
            return
        QMessageBox.information(self, "下载成功", f"文件已保存到：\n{target}")

    def _delete(self, record: PersonalFileRecord) -> None:
        answer = QMessageBox.question(
            self,
            "删除确认",
            f"确定删除“{record.name}”吗？\n生产网上的文件也会一并删除。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        self.service.delete(self.category, record.id, remove_file=True)
        self.refresh()

    def _on_line_changed(self, current, _previous) -> None:
        if self._loading or current is None:
            return
        self._load_products(current.text())

    def _on_product_changed(self, current, _previous) -> None:
        if self._loading or current is None:
            return
        line = self._current_text(self.line_list)
        self._load_files(line, current.text())

    @staticmethod
    def _current_text(widget: QListWidget) -> str:
        item = widget.currentItem()
        return item.text() if item is not None else ""

    @staticmethod
    def _select_text(widget: QListWidget, text: str) -> None:
        for index in range(widget.count()):
            item = widget.item(index)
            if item.text() == text:
                widget.setCurrentItem(item)
                return
