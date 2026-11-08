"""工艺 / 工装文件公共查询页：按条件检索，或按生产线浏览，支持下载。"""

from __future__ import annotations

import shutil
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.services.line_access import ALL_LINES, allowed_lines, normalise_line
from app.services.personal_file_service import (
    CATEGORY_CONFIG,
    PersonalFileRecord,
    PersonalFileService,
)
from app.widgets.line_combo import LineCombo


class LibraryQueryView(QWidget):
    """所有用户都可查询、下载的公共文件库页面。"""

    def __init__(self, db, user, category: str, parent=None) -> None:
        super().__init__(parent)
        self.category = category
        self.user = user
        self.allowed_lines = allowed_lines(user.role)
        self.restricted = set(self.allowed_lines) != set(ALL_LINES)
        self.config = CATEGORY_CONFIG[category]
        self.service = PersonalFileService(db)
        self.records: list[PersonalFileRecord] = []
        self._loading = False
        self._flat = False

        self._columns: list[tuple[str, str, int]] = [
            ("产品代号", "product_code", 120),
            ("生产线", "production_line", 100),
        ]
        if category == "process":
            self._columns.append(("工序", "process", 120))
        self._columns.extend(
            (
                ("名称", "name", 180),
                ("文件", "file_name", 0),
                ("编制", "employee_name", 90),
                ("上传时间", "uploaded_at", 140),
            )
        )
        self._action_column = len(self._columns)

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 22, 28, 22)
        root.setSpacing(14)
        root.addWidget(self._build_filter_card())
        root.addWidget(self._build_list_card(), 1)

        self._reload_choices()
        self._show_browse()

    @property
    def label(self) -> str:
        return "工艺文件" if self.category == "process" else "工装文件"

    # ---------- 筛选区 ----------
    def _build_filter_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel(f"{self.label}查询")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("所有用户上传的文件均可查询与下载")
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
        self.product_edit.returnPressed.connect(self._handle_search)
        form.addWidget(self.product_edit, 0, 1)

        form.addWidget(self._label("生产线"), 0, 2)
        self.line_combo = LineCombo()
        self.line_combo.set_allowed(self.allowed_lines)
        form.addWidget(self.line_combo, 0, 3)

        form.addWidget(self._label("编制"), 0, 4)
        self.author_combo = QComboBox()
        self.author_combo.setEditable(True)
        self.author_combo.setInsertPolicy(QComboBox.NoInsert)
        self.author_combo.lineEdit().setPlaceholderText("选择或输入上传人")
        form.addWidget(self.author_combo, 0, 5)
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
        search.clicked.connect(self._handle_search)
        footer.addWidget(search)
        layout.addLayout(footer)
        return card

    @staticmethod
    def _label(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    # ---------- 列表区 ----------
    def _build_list_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        self.list_title = QLabel("按生产线浏览")
        self.list_title.setObjectName("sectionTitle")
        header.addWidget(self.list_title)
        header.addStretch(1)
        self.count_label = QLabel("共 0 个")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        layout.addLayout(header)

        self.browse_hint = QLabel(
            "未搜索时按生产线浏览，点击左侧生产线查看该产线全部文件。"
        )
        self.browse_hint.setObjectName("sectionHint")
        self.browse_hint.setWordWrap(True)
        layout.addWidget(self.browse_hint)

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(14)

        self.line_box = QWidget()
        line_layout = QVBoxLayout(self.line_box)
        line_layout.setContentsMargins(0, 0, 0, 0)
        line_layout.setSpacing(6)
        line_label = QLabel("生产线")
        line_label.setObjectName("fieldLabel")
        line_layout.addWidget(line_label)
        self.line_list = QListWidget()
        self.line_list.setFixedWidth(170)
        self.line_list.currentItemChanged.connect(self._on_line_changed)
        line_layout.addWidget(self.line_list, 1)
        body.addWidget(self.line_box, 0)

        self.file_table = QTableWidget(0, self._action_column + 1)
        headers = [item[0] for item in self._columns] + ["操作"]
        self.file_table.setHorizontalHeaderLabels(headers)
        self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.file_table.setAlternatingRowColors(True)
        self.file_table.setShowGrid(False)
        self.file_table.setWordWrap(False)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.verticalHeader().setDefaultSectionSize(40)
        self.file_table.horizontalHeader().setFixedHeight(38)
        self.file_table.horizontalHeader().setHighlightSections(False)
        self.file_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.Fixed
        )
        for index, (_title, _attr, width) in enumerate(self._columns):
            if not width:
                self.file_table.horizontalHeader().setSectionResizeMode(
                    index, QHeaderView.Stretch
                )
            else:
                self.file_table.setColumnWidth(index, width)
        self.file_table.setColumnWidth(self._action_column, 90)
        body.addWidget(self.file_table, 1)
        layout.addLayout(body, 1)
        return card

    # ---------- 数据加载 ----------
    def _reload_choices(self) -> None:
        current_line = self.line_combo.currentText().strip()
        current_author = self.author_combo.currentText().strip()
        self.line_combo.set_allowed(self.allowed_lines, current_line)
        self.author_combo.blockSignals(True)
        self.author_combo.clear()
        self.author_combo.addItems(self.service.list_uploaders(self.category))
        self.author_combo.setEditText(current_author)
        self.author_combo.blockSignals(False)

    def _show_browse(self) -> None:
        self._flat = False
        self.line_box.show()
        self.list_title.setText("按生产线浏览")
        self.browse_hint.setText(
            "未搜索时按生产线浏览，点击左侧生产线查看该产线全部文件。"
        )
        self.filter_status.setText("")
        selected = self._current_text(self.line_list)
        lines = list(self.allowed_lines)
        self._loading = True
        self.line_list.clear()
        for line in lines:
            self.line_list.addItem(QListWidgetItem(line))
        self._loading = False
        if not lines:
            self._render_records([])
            return
        available = {
            code
            for code in (
                normalise_line(record.production_line) for record in self.service.query_files(self.category)
            )
            if code in lines
        }
        if selected in available:
            target = selected
        else:
            target = next((line for line in lines if line in available), lines[0])
        self._loading = True
        self._select_text(self.line_list, target)
        self._loading = False
        self._load_line(target)

    def _load_line(self, line: str) -> None:
        records = self.service.query_files(self.category, line=line)
        self._render_records(records)

    def _handle_search(self) -> None:
        product_code = self.product_edit.text().strip()
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
        author = self.author_combo.currentText().strip()
        records = self.service.query_files(
            self.category, product_code=product_code, line=line, author=author
        )
        if not line and self.restricted:
            allowed = set(self.allowed_lines)
            records = [
                record
                for record in records
                if normalise_line(record.production_line) in allowed
            ]
        self._flat = True
        self.line_box.hide()
        self.list_title.setText("查询结果")
        self.browse_hint.setText(
            f"产品代号：{product_code or '全部'}　生产线：{line or '全部'}　"
            f"编制：{author or '全部'}"
        )
        self.filter_status.setText(f"共找到 {len(records)} 个{self.label}。")
        self._render_records(records)

    def _reset_filters(self) -> None:
        self.product_edit.clear()
        self.line_combo.set_line("")
        self.author_combo.setEditText("")
        self._reload_choices()
        self._show_browse()

    def _render_records(self, records: list[PersonalFileRecord]) -> None:
        self.records = records
        self.file_table.setRowCount(len(records))
        for row, record in enumerate(records):
            for column, (_title, attr, _width) in enumerate(self._columns):
                value = str(getattr(record, attr, "") or "")
                if attr == "uploaded_at":
                    value = value.replace("T", " ")[:16]
                item = QTableWidgetItem(value)
                item.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if attr in ("name", "file_name")
                    else Qt.AlignCenter
                )
                item.setToolTip(
                    record.file_path if attr == "file_name" else value
                )
                self.file_table.setItem(row, column, item)
            self.file_table.setCellWidget(
                row, self._action_column, self._build_action(record)
            )
        self.count_label.setText(f"共 {len(records)} 个")

    def _build_action(self, record: PersonalFileRecord) -> QWidget:
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

    def _on_line_changed(self, current, _previous) -> None:
        if self._loading or current is None or self._flat:
            return
        self._load_line(current.text())

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

    def on_activated(self) -> None:
        self._reload_choices()
        if self._flat:
            self._handle_search()
        else:
            self._show_browse()
