"""产品责任人查询：按产品代号 / 主机 / 生产线任选条件查询装配、焊接负责人。"""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.services.line_access import (
    ALL_LINES,
    allowed_lines,
    line_codes_in_text,
)
from app.services.product_owner_service import ProductOwnerService
from app.widgets.line_combo import LineCombo

COLUMNS = ("产品代号", "主机", "生产线", "装配负责人", "焊接负责人")


class ProductOwnerView(QWidget):
    def __init__(self, db, user, parent=None) -> None:
        super().__init__(parent)
        self.user = user
        self.allowed_lines = allowed_lines(user.role)
        self.restricted = set(self.allowed_lines) != set(ALL_LINES)
        self.service = ProductOwnerService(db)
        self.records = []

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(14)
        root.addWidget(self._build_filter_card())
        root.addWidget(self._build_table_card(), 1)
        self.load_records()

    def _build_filter_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("产品责任人查询")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("产品代号、主机、生产线可任选一项或多项")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        form = QGridLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(9)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)
        form.setColumnStretch(5, 1)

        form.addWidget(self._label("产品代号"), 0, 0)
        self.product_edit = QLineEdit()
        self.product_edit.setPlaceholderText("输入产品代号，支持模糊查询")
        self.product_edit.returnPressed.connect(self.load_records)
        form.addWidget(self.product_edit, 0, 1)

        form.addWidget(self._label("主机"), 0, 2)
        self.host_edit = QLineEdit()
        self.host_edit.setPlaceholderText("输入主机，支持模糊查询")
        self.host_edit.returnPressed.connect(self.load_records)
        form.addWidget(self.host_edit, 0, 3)

        form.addWidget(self._label("生产线"), 0, 4)
        self.line_combo = LineCombo()
        self.line_combo.set_allowed(self.allowed_lines)
        form.addWidget(self.line_combo, 0, 5)
        layout.addLayout(form)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 2, 0, 0)
        footer.setSpacing(10)
        self.filter_status = QLabel("")
        self.filter_status.setObjectName("fieldHint")
        footer.addWidget(self.filter_status, 1)
        reset = QPushButton("重置")
        reset.setObjectName("secondaryButton")
        reset.setCursor(Qt.PointingHandCursor)
        reset.clicked.connect(self._reset_filters)
        footer.addWidget(reset)
        search = QPushButton("搜索")
        search.setObjectName("primaryButton")
        search.setCursor(Qt.PointingHandCursor)
        search.clicked.connect(self.load_records)
        footer.addWidget(search)
        layout.addLayout(footer)
        return card

    def _build_table_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("责任人与产品")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.count_label = QLabel("共 0 条")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        layout.addLayout(header)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.horizontalHeader().setFixedHeight(38)
        self.table.horizontalHeader().setHighlightSections(False)
        for index in range(len(COLUMNS)):
            self.table.horizontalHeader().setSectionResizeMode(
                index, QHeaderView.Stretch
            )
        layout.addWidget(self.table, 1)
        return card

    @staticmethod
    def _label(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def load_records(self) -> None:
        line_text = self.line_combo.currentText().strip()
        line = ""
        if line_text:
            line = self.line_combo.current_line() or ""
            if not line:
                QMessageBox.warning(
                    self,
                    "产线代号无效",
                    "请输入当前账号可用的产线代号，并从联想列表中选择。",
                )
                return
        self.records = self.service.query(
            self.product_edit.text().strip(),
            self.host_edit.text().strip(),
            line,
        )
        if self.restricted:
            allowed = set(self.allowed_lines)
            self.records = [
                record
                for record in self.records
                if line_codes_in_text(record.production_line) & allowed
            ]
        self._render()

    def _render(self) -> None:
        self.table.setRowCount(len(self.records))
        for row, record in enumerate(self.records):
            values = (
                record.product_code,
                record.host,
                record.production_line,
                record.assembly or "未指定",
                record.welding or "未指定",
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if column in (1, 2)
                    else Qt.AlignCenter
                )
                item.setToolTip(value)
                self.table.setItem(row, column, item)
        self.count_label.setText(f"共 {len(self.records)} 条")
        self.filter_status.setText(
            f"产品代号：{self.product_edit.text().strip() or '全部'}　"
            f"主机：{self.host_edit.text().strip() or '全部'}　"
            f"生产线：{self.line_combo.current_line() or '全部'}"
        )

    def _reset_filters(self) -> None:
        self.product_edit.clear()
        self.host_edit.clear()
        self.line_combo.set_line("")
        self.load_records()

    def on_activated(self) -> None:
        self.load_records()
