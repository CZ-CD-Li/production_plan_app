"""个人-产品列表：展示当前用户负责的产品并提供工装/工艺入口。"""

from __future__ import annotations

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.models import User
from app.services.personal_product_service import PersonalProduct, PersonalProductService


class PersonalProductListView(QWidget):
    tooling_requested = pyqtSignal(str, str)
    process_requested = pyqtSignal(str, str)

    HEADERS = ("产品代号", "主机", "生产线", "操作")

    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(parent)
        self.user = user
        self.service = PersonalProductService(db)
        self._products: list[PersonalProduct] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(14)

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(22, 20, 22, 18)
        card_layout.setSpacing(12)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(10)
        title = QLabel("我负责的产品")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.count_label = QLabel("共 0 个")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        card_layout.addLayout(header)

        hint = QLabel("工艺负责人：装配或焊接字段包含当前用户姓名的产品")
        hint.setObjectName("sectionHint")
        card_layout.addWidget(hint)

        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(list(self.HEADERS))
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
        header_view = self.table.horizontalHeader()
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 180)
        self.table.setColumnWidth(2, 160)
        header_view.setSectionResizeMode(0, QHeaderView.Fixed)
        header_view.setSectionResizeMode(1, QHeaderView.Fixed)
        header_view.setSectionResizeMode(2, QHeaderView.Fixed)
        header_view.setSectionResizeMode(3, QHeaderView.Stretch)
        card_layout.addWidget(self.table, 1)

        self.empty_label = QLabel("暂无负责的产品，请先在项目分布清单中配置负责人")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.empty_label.setStyleSheet(
            "background:transparent;color:#9AA7B8;font-size:13px;"
        )
        self.empty_label.setParent(self.table.viewport())
        self.empty_label.hide()

        root.addWidget(card, 1)
        self.refresh()

    def resizeEvent(self, event) -> None:  # noqa: N802
        super().resizeEvent(event)
        self.empty_label.setGeometry(self.table.viewport().rect())

    def on_activated(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        self._products = self.service.products_for(self.user.name)
        self.table.setRowCount(0)
        self.table.setRowCount(len(self._products))

        for row_index, product in enumerate(self._products):
            for column, value in enumerate(
                (product.product_code, product.host, product.production_line)
            ):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                item.setToolTip(value)
                self.table.setItem(row_index, column, item)
            self.table.setCellWidget(
                row_index, 3, self._build_actions(product)
            )

        self.count_label.setText(f"共 {len(self._products)} 个")
        self.empty_label.setVisible(not self._products)
        self.empty_label.setGeometry(self.table.viewport().rect())

    def _build_actions(self, product: PersonalProduct) -> QWidget:
        holder = QWidget()
        holder.setStyleSheet("background:transparent;")
        layout = QHBoxLayout(holder)
        layout.setContentsMargins(6, 0, 6, 0)
        layout.setSpacing(6)

        tooling = QPushButton("工装")
        tooling.setObjectName("linkButton")
        tooling.setCursor(Qt.PointingHandCursor)
        tooling.clicked.connect(
            lambda: self.tooling_requested.emit(
                product.product_code, product.production_line
            )
        )
        layout.addWidget(tooling)

        process = QPushButton("工艺")
        process.setObjectName("linkButton")
        process.setCursor(Qt.PointingHandCursor)
        process.clicked.connect(
            lambda: self.process_requested.emit(
                product.product_code, product.production_line
            )
        )
        layout.addWidget(process)
        layout.addStretch(1)
        return holder
