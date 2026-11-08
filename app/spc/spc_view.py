"""个性化功能定制 · SPC 生成页面。"""

from __future__ import annotations

import shutil
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.models import User
from app.personal.drop_area import FileDropArea
from app.spc.generator import (
    CHART_XBAR,
    CHART_XMR,
    PLAN_MAPPING_FIELDS,
    TEMPLATE_CHANGAN,
    TEMPLATE_LABELS,
    TEMPLATE_WULING,
    WULING_SAMPLE_COUNT,
    SpcGenerateError,
    SpcGenerationResult,
    generate_spc,
)
from app.spc.spec_parser import (
    SpcItem,
    SpcPlanError,
    SpcPlanImport,
    import_spc_plan,
)

COLUMNS = (
    "序号",
    "特性描述",
    "下限",
    "目标值",
    "上限",
    "单位",
    "量具名称",
    "量具精度",
    "图表类型",
    "状态",
)
EDITABLE_COLUMNS = (2, 3, 4, 5, 6, 7)
CHART_COLUMN = 8
STATUS_COLUMN = 9
INVALID_COLOR = QColor("#B42318")
VALID_COLOR = QColor("#067647")


def _text(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.6f}".rstrip("0").rstrip(".")


def _number(text: str) -> float | None:
    cleaned = text.strip().replace("，", "").replace(",", "")
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


class SpcGenerateView(QWidget):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.imported: SpcPlanImport | None = None
        self.file_path = ""
        self.items: list[SpcItem] = []
        self.confirmed = False
        self.result: SpcGenerationResult | None = None
        self.output_path = ""
        self._chart_backup: dict[int, str] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(14)

        layout.addWidget(self._build_upload_card())
        self.column_card = self._build_column_card()
        layout.addWidget(self.column_card)
        self.column_card.hide()
        self.items_card = self._build_items_card()
        layout.addWidget(self.items_card)
        self.items_card.hide()
        self.generate_card = self._build_generate_card()
        layout.addWidget(self.generate_card)
        self.generate_card.hide()
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card)
        self.result_card.hide()
        layout.addStretch(1)

        scroll.setWidget(content)
        root.addWidget(scroll)

    # ---------- 上传 ----------
    def _build_upload_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("导入 SPC 计划")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("支持 .xlsx / .xlsm，拖入或点击选择文件")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        self.drop_area = FileDropArea()
        self.drop_area.title_label.setText("拖入 SPC 计划 Excel，或点击右侧按钮选择")
        self.drop_area.file_label.setText(
            "自动识别序号列与标题栏，选择公差栏后解析数据范围、单位与量具精度"
        )
        self.drop_area.file_selected.connect(self._on_file_selected)
        layout.addWidget(self.drop_area)

        self.import_label = QLabel("")
        self.import_label.setObjectName("fieldHint")
        self.import_label.setWordWrap(True)
        layout.addWidget(self.import_label)
        return card

    def _on_file_selected(self, path: str) -> None:
        if not path:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            imported = import_spc_plan(path)
        except SpcPlanError as exc:
            QMessageBox.warning(self, "导入失败", str(exc))
            self.import_label.setText(str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.file_path = path
        self.imported = imported
        self.items = []
        self.confirmed = False
        self.result = None
        self.output_path = ""
        self.result_card.hide()
        self.items_card.hide()
        self.generate_card.hide()

        self.column_combo.clear()
        for column in imported.columns:
            self.column_combo.addItem(column.label, column.index)
        self._rebuild_plan_mapping()
        self.import_label.setText(
            f"已识别工作表“{imported.sheet_name}”，标题栏共 {len(imported.columns)} 列。"
            "请选择公差对应栏后点击“确认公差栏”。"
        )
        self.column_card.show()

    # ---------- 公差栏选择 ----------
    def _build_column_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("选择公差对应栏")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        layout.addLayout(header)

        hint = QLabel(
            "已定位内容为 1 且下一行为 2 的序号列，其上方为标题栏。"
            "请选择包含公差（数据范围）与单位的列，确认后再解析各项数据。"
        )
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        label = QLabel("公差栏")
        label.setObjectName("fieldLabel")
        row.addWidget(label)
        self.column_combo = QComboBox()
        self.column_combo.setMinimumWidth(320)
        row.addWidget(self.column_combo)
        row.addStretch(1)
        self.column_confirm_button = QPushButton("确认公差栏")
        self.column_confirm_button.setObjectName("primaryButton")
        self.column_confirm_button.setCursor(Qt.PointingHandCursor)
        self.column_confirm_button.clicked.connect(self._handle_column_confirm)
        row.addWidget(self.column_confirm_button)
        layout.addLayout(row)
        return card

    def _handle_column_confirm(self) -> None:
        if not self.file_path or self.imported is None:
            QMessageBox.warning(self, "无法确认", "请先导入 SPC 计划文件。")
            return
        column = self.column_combo.currentData()
        if column is None:
            QMessageBox.warning(self, "无法确认", "标题栏中没有可选的公差栏。")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            imported = import_spc_plan(self.file_path, tolerance_col=int(column))
        except SpcPlanError as exc:
            QMessageBox.warning(self, "解析失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.imported = imported
        self.items = imported.items
        self.confirmed = False
        self.result = None
        self.output_path = ""
        self.result_card.hide()
        self.generate_card.hide()
        self.confirm_status.setText("")
        self.confirm_status.setStyleSheet("")
        self.items_count.setText("")

        ready = sum(1 for item in self.items if item.ready)
        self.import_label.setText(
            f"已识别工作表“{imported.sheet_name}”，公差栏："
            f"{self.column_combo.currentText()}，共 {len(self.items)} 项，"
            f"自动识别完整 {ready} 项，需人工确认 {len(self.items) - ready} 项。"
            "请核对下表数据范围、单位与量具精度后点击“确认数据范围”。"
        )
        self._fill_table()
        self.items_card.show()

        pending = [item for item in self.items if not item.ready]
        if pending:
            lines = [
                f"序号 {item.index}：{item.note or item.problem}" for item in pending
            ]
            if len(lines) > 12:
                lines = lines[:12] + [f"…… 其余 {len(lines) - 12} 项同样需要确认"]
            QMessageBox.warning(
                self,
                "存在未识别项",
                "以下序号行的数据范围或单位未识别，请在表格中人工填写：\n\n"
                + "\n".join(lines),
            )

    # ---------- 数据确认表 ----------
    def _build_items_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("识别结果确认")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.items_count = QLabel("")
        self.items_count.setObjectName("countPill")
        header.addWidget(self.items_count)
        layout.addLayout(header)

        hint = QLabel(
            "下限 / 目标值 / 上限 / 单位 / 量具名称 / 量具精度可直接修改；"
            "单边公差（≥ / ≤）默认目标值 ±5、公差 ±5。"
            "图表类型可逐行选择，未给出中心值的数据范围（≥ / ≤ / 上下限）"
            "默认单值极差图（X-MR）。"
        )
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.items_table = QTableWidget(0, len(COLUMNS))
        self.items_table.setHorizontalHeaderLabels(list(COLUMNS))
        self.items_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.SelectedClicked
            | QAbstractItemView.EditKeyPressed
        )
        self.items_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.items_table.setAlternatingRowColors(True)
        self.items_table.setShowGrid(False)
        self.items_table.verticalHeader().setVisible(False)
        self.items_table.verticalHeader().setDefaultSectionSize(34)
        self.items_table.horizontalHeader().setFixedHeight(38)
        self.items_table.horizontalHeader().setHighlightSections(False)
        self.items_table.setColumnWidth(0, 60)
        self.items_table.setColumnWidth(1, 340)
        for index in (2, 3, 4, 6, 7):
            self.items_table.setColumnWidth(index, 96)
        self.items_table.setColumnWidth(5, 72)
        self.items_table.setColumnWidth(CHART_COLUMN, 150)
        self.items_table.horizontalHeader().setSectionResizeMode(
            STATUS_COLUMN, QHeaderView.Stretch
        )
        self.items_table.setMinimumHeight(360)
        layout.addWidget(self.items_table)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 4, 0, 0)
        footer.setSpacing(12)
        self.confirm_status = QLabel("")
        self.confirm_status.setObjectName("fieldHint")
        footer.addWidget(self.confirm_status, 1)
        self.confirm_button = QPushButton("确认数据范围")
        self.confirm_button.setObjectName("primaryButton")
        self.confirm_button.setCursor(Qt.PointingHandCursor)
        self.confirm_button.clicked.connect(self._handle_confirm)
        footer.addWidget(self.confirm_button, 0, Qt.AlignRight)
        layout.addLayout(footer)
        return card

    def _fill_table(self) -> None:
        self._chart_backup = {}
        self.items_table.setRowCount(len(self.items))
        for row, item in enumerate(self.items):
            values = (
                item.index,
                item.description,
                _text(item.lower),
                _text(item.target),
                _text(item.upper),
                item.unit,
                item.gauge_name,
                _text(item.precision),
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                if column in (1, 6):
                    cell.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                else:
                    cell.setTextAlignment(Qt.AlignCenter)
                if column not in EDITABLE_COLUMNS:
                    cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                if column == 1:
                    cell.setToolTip(item.description)
                self.items_table.setItem(row, column, cell)
            placeholder = QTableWidgetItem("")
            self.items_table.setItem(row, CHART_COLUMN, placeholder)
            combo = QComboBox()
            combo.addItem("均值极差图（Xbar-R）", CHART_XBAR)
            combo.addItem("单值极差图（X-MR）", CHART_XMR)
            combo.setCurrentIndex(
                1 if item.spec_kind in ("lower", "upper", "range") else 0
            )
            combo.setToolTip("没有给出中心值的数据范围建议使用单值极差图")
            self.items_table.setCellWidget(row, CHART_COLUMN, combo)
            self._update_status(row)
        self._apply_template_mode()
        self.items_count.setText(f"共 {len(self.items)} 项")

    def _template_key(self) -> str:
        data = self.template_combo.currentData()
        return data or TEMPLATE_CHANGAN

    def _apply_template_mode(self) -> None:
        """五菱模板固定 5 件/组 × 25 组，只允许均值极差图。"""
        wuling = self._template_key() == TEMPLATE_WULING
        for row in range(self.items_table.rowCount()):
            widget = self.items_table.cellWidget(row, CHART_COLUMN)
            if not isinstance(widget, QComboBox):
                continue
            if wuling:
                if row not in self._chart_backup:
                    self._chart_backup[row] = widget.currentData() or CHART_XBAR
                    widget.setCurrentIndex(0)
                widget.setEnabled(False)
            else:
                widget.setEnabled(True)
                saved = self._chart_backup.pop(row, None)
                if saved:
                    index = widget.findData(saved)
                    if index >= 0:
                        widget.setCurrentIndex(index)
        if hasattr(self, "plan_map_box"):
            self.plan_map_box.setVisible(wuling)
        if wuling:
            self.count_spin.setValue(WULING_SAMPLE_COUNT)
            self.count_spin.setEnabled(False)
            self.count_label.setText("数据数量（五菱固定 125）")
            self.count_note.setText(
                "五菱模板固定 25 组 × 5 件，共 125 个数据；"
                "每组 5 个值分别填入模板的 10 + 10 + 5 组数据区。"
            )
            self.generate_status.setText(
                "五菱模板：生成时自动校验数据点不超界、无连续 7 点升降、"
                "CPK 1.67~1.92、PPK 1.65~1.92，不满足会重新生成。"
            )
        else:
            self.count_spin.setEnabled(True)
            self.count_label.setText("数据数量（25 的倍数）")
            self.count_note.setText(
                "均值极差图：每组 2~5 个样本（模板 C9:AA13 共 5 行 25 列），"
                "默认 50 个 = 25 组 × 2 个；单值极差图固定 25 个单值。"
            )
            self.generate_status.setText(
                "生成时自动校验：数据点不超界、无连续 7 点升降、"
                "CPK 1.67~1.92、PPK 1.65~1.92，不满足会重新生成；"
                "单值极差图另校验 X 图 X̄±2.66MR̄ 与 MR 图 3.267MR̄ 控制限。"
            )

    def _update_status(self, row: int) -> None:
        item = self._row_item(row)
        status = self.items_table.item(row, STATUS_COLUMN)
        if status is None:
            status = QTableWidgetItem("")
            status.setTextAlignment(Qt.AlignCenter)
            status.setFlags(status.flags() & ~Qt.ItemIsEditable)
            self.items_table.setItem(row, STATUS_COLUMN, status)
        if item.ready:
            status.setText("已识别")
            status.setForeground(VALID_COLOR)
        else:
            status.setText("需确认")
            status.setForeground(INVALID_COLOR)

    def _row_item(self, row: int) -> SpcItem:
        source = self.items[row]
        widget = self.items_table.cellWidget(row, CHART_COLUMN)
        chart_type = CHART_XBAR
        if isinstance(widget, QComboBox):
            chart_type = widget.currentData() or CHART_XBAR
        return SpcItem(
            row=source.row,
            index=source.index,
            description=source.description,
            gauge_name=self._cell_text(row, 6),
            gauge_spec=source.gauge_spec,
            precision_text=self._cell_text(row, 7),
            precision=_number(self._cell_text(row, 7)) or 0.0,
            lower=_number(self._cell_text(row, 2)),
            target=_number(self._cell_text(row, 3)),
            upper=_number(self._cell_text(row, 4)),
            unit=self._cell_text(row, 5),
            note=source.note,
            spec_kind=source.spec_kind,
            chart_type=chart_type,
        )

    def _cell_text(self, row: int, column: int) -> str:
        cell = self.items_table.item(row, column)
        return cell.text().strip() if cell is not None else ""

    def _handle_confirm(self) -> None:
        if not self.items:
            QMessageBox.warning(self, "无法确认", "请先导入 SPC 计划文件。")
            return
        collected: list[SpcItem] = []
        problems: list[str] = []
        for row in range(len(self.items)):
            item = self._row_item(row)
            collected.append(item)
            self._update_status(row)
            if item.ready:
                continue
            problems.append(f"第 {row + 1} 行（序号 {item.index}）：{item.problem}")
            for column in range(len(COLUMNS)):
                cell = self.items_table.item(row, column)
                if cell is not None:
                    cell.setBackground(QColor("#FEE4E2"))
        if problems:
            header = "\n".join(problems[:10])
            if len(problems) > 10:
                header += f"\n…… 其余 {len(problems) - 10} 项"
            QMessageBox.warning(
                self,
                "数据未确认",
                "以下行的数据范围或单位不完整，请修改后再确认：\n\n" + header,
            )
            return
        for row in range(len(self.items)):
            for column in range(len(COLUMNS)):
                cell = self.items_table.item(row, column)
                if cell is not None:
                    cell.setBackground(QColor(0, 0, 0, 0))
        self.items = collected
        self.confirmed = True
        self.confirm_status.setText(
            f"已确认 {len(collected)} 项数据，可设置数据数量并生成。"
        )
        self.confirm_status.setStyleSheet(f"color:{VALID_COLOR.name()};")
        self.generate_card.show()

    # ---------- 生成 ----------
    def _build_generate_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(12)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("生成 SPC 数据")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        layout.addLayout(header)

        template_row = QHBoxLayout()
        template_row.setContentsMargins(0, 0, 0, 0)
        template_row.setSpacing(10)
        template_label = QLabel("SPC 模板")
        template_label.setObjectName("fieldLabel")
        template_row.addWidget(template_label)
        self.template_combo = QComboBox()
        self.template_combo.addItem(
            f"{TEMPLATE_LABELS[TEMPLATE_CHANGAN]}（现有模板）", TEMPLATE_CHANGAN
        )
        self.template_combo.addItem(
            f"{TEMPLATE_LABELS[TEMPLATE_WULING]}（5 件/组 × 25 组）",
            TEMPLATE_WULING,
        )
        self.template_combo.setFixedWidth(240)
        self.template_combo.currentIndexChanged.connect(
            lambda _index: self._apply_template_mode()
        )
        template_row.addWidget(self.template_combo)
        self.template_hint = QLabel(
            "长安：支持均值极差图 / 单值极差图；五菱：固定 125 个数据、全部均值极差图。"
        )
        self.template_hint.setObjectName("sectionHint")
        self.template_hint.setWordWrap(True)
        template_row.addWidget(self.template_hint, 1)
        layout.addLayout(template_row)

        self.plan_map_box = self._build_plan_mapping_box()
        layout.addWidget(self.plan_map_box)
        self.plan_map_box.hide()

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        self.count_label = QLabel("数据数量（25 的倍数）")
        self.count_label.setObjectName("fieldLabel")
        row.addWidget(self.count_label)
        self.count_spin = QSpinBox()
        self.count_spin.setRange(25, 2000)
        self.count_spin.setSingleStep(25)
        self.count_spin.setValue(50)
        self.count_spin.setSuffix(" 个")
        self.count_spin.setFixedWidth(140)
        row.addWidget(self.count_spin)
        self.count_note = QLabel(
            "均值极差图：每组 2~5 个样本（模板 C9:AA13 共 5 行 25 列），"
            "默认 50 个 = 25 组 × 2 个；单值极差图固定 25 个单值。"
        )
        self.count_note.setObjectName("sectionHint")
        self.count_note.setWordWrap(True)
        row.addWidget(self.count_note, 1)
        layout.addLayout(row)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(12)
        self.generate_status = QLabel(
            "生成时自动校验：数据点不超界、无连续 7 点升降、"
            "CPK 1.67~1.92、PPK 1.65~1.92，不满足会重新生成；"
            "单值极差图另校验 X 图 X̄±2.66MR̄ 与 MR 图 3.267MR̄ 控制限。"
        )
        self.generate_status.setObjectName("fieldHint")
        self.generate_status.setWordWrap(True)
        footer.addWidget(self.generate_status, 1)
        self.generate_button = QPushButton("生成 SPC 文件")
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self._handle_generate)
        footer.addWidget(self.generate_button, 0, Qt.AlignRight)
        layout.addLayout(footer)
        return card

    def _build_plan_mapping_box(self) -> QWidget:
        """五菱模板下，把确认表字段映射到导入计划表的标题栏列。"""
        box = QWidget()
        box.setStyleSheet("background:transparent;")
        layout = QVBoxLayout(box)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(8)

        hint = QLabel(
            "量具名称等信息回填：为确认表字段选择导入计划表中的标题栏列"
            "（序号 1 上一行），生成时写入“SPC计划”页对应列。"
        )
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(8)
        self.plan_map_combos: dict[str, QComboBox] = {}
        for index, field in enumerate(PLAN_MAPPING_FIELDS):
            row = index // 2
            pair = index % 2
            label = QLabel(field)
            label.setObjectName("fieldLabel")
            combo = QComboBox()
            combo.setMinimumWidth(240)
            combo.addItem("不回填", None)
            self.plan_map_combos[field] = combo
            grid.addWidget(label, row, pair * 3)
            grid.addWidget(combo, row, pair * 3 + 1)
        layout.addLayout(grid)
        return box

    _PLAN_MATCH_TOKENS: dict[str, tuple[str, ...]] = {
        "特性描述": ("特性描述", "特性", "测量项目", "项目", "description"),
        "下限": ("下限", "lsl", "下公差"),
        "目标值": ("目标值", "目标", "中心值", "中心", "target"),
        "上限": ("上限", "usl", "上公差"),
        "单位": ("单位", "unit"),
        "量具名称": ("量具名称", "检具", "名称", "gauge"),
        "量具精度": ("量具精度", "精度", "分辨力", "precision"),
    }

    def _rebuild_plan_mapping(self) -> None:
        columns = self.imported.columns if self.imported is not None else []
        for field, combo in self.plan_map_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("不回填", None)
            for column in columns:
                combo.addItem(column.label, column.index)
            guess = None
            tokens = self._PLAN_MATCH_TOKENS.get(field, (field,))
            for token in tokens:
                for column in columns:
                    if token.lower() in column.title.lower():
                        guess = column.index
                        break
                if guess is not None:
                    break
            index = combo.findData(guess) if guess is not None else -1
            combo.setCurrentIndex(index if index >= 0 else 0)
            combo.blockSignals(False)

    def _plan_column_map(self) -> dict[str, int]:
        return {
            field: combo.currentData()
            for field, combo in self.plan_map_combos.items()
            if combo.currentData() is not None
        }

    def _handle_generate(self) -> None:
        if not self.confirmed or not self.items:
            QMessageBox.warning(self, "无法生成", "请先确认数据范围。")
            return
        items = [self._row_item(row) for row in range(len(self.items))]
        template_key = self._template_key()
        if template_key == TEMPLATE_WULING:
            sample_count = WULING_SAMPLE_COUNT
        else:
            sample_count = self.count_spin.value()
            has_xbar = any(item.chart_type != CHART_XMR for item in items)
            if has_xbar and sample_count % 25 != 0:
                QMessageBox.warning(self, "无法生成", "数据数量必须是 25 的倍数。")
                return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.generate_button.setEnabled(False)
        try:
            result = generate_spc(
                self.imported,
                items,
                sample_count,
                self.user,
                template_key=template_key,
                plan_column_map=(
                    self._plan_column_map()
                    if template_key == TEMPLATE_WULING
                    else None
                ),
            )
        except SpcGenerateError as exc:
            QMessageBox.warning(self, "生成失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
            self.generate_button.setEnabled(True)

        self.items = items
        self.result = result
        self.output_path = result.output_path
        self.result_card.show()
        self._fill_result(result)
        xbar_count = sum(
            1 for item in result.results if item.chart_type != CHART_XMR
        )
        xmr_count = len(result.results) - xbar_count
        details = []
        if xbar_count:
            details.append(
                f"{xbar_count} 项均值极差图（每项 {result.sample_count} 个数据，"
                f"{result.per_group} 个/组）"
            )
        if xmr_count:
            details.append(f"{xmr_count} 项单值极差图（每项 25 个单值）")
        QMessageBox.information(
            self,
            "已完成",
            f"已按{result.template_label}模板生成 {len(result.results)} 项 SPC 数据："
            + "、".join(details)
            + "。\n请点击“下载 SPC 文件”保存后，用 Excel 打开确认图表与 CPK 是否正确。",
        )

    # ---------- 结果 ----------
    def _build_result_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("生成结果")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.download_button = QPushButton("下载 SPC 文件")
        self.download_button.setObjectName("primaryButton")
        self.download_button.setCursor(Qt.PointingHandCursor)
        self.download_button.clicked.connect(self._download)
        header.addWidget(self.download_button)
        layout.addLayout(header)

        self.result_label = QLabel("")
        self.result_label.setObjectName("fieldHint")
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

        self.result_table = QTableWidget(0, 6)
        self.result_table.setHorizontalHeaderLabels(
            ["序号", "图表类型", "工作表", "CPK", "PPK", "生成次数"]
        )
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setShowGrid(False)
        self.result_table.verticalHeader().setVisible(False)
        self.result_table.verticalHeader().setDefaultSectionSize(34)
        self.result_table.horizontalHeader().setFixedHeight(38)
        self.result_table.horizontalHeader().setHighlightSections(False)
        self.result_table.setColumnWidth(0, 60)
        self.result_table.setColumnWidth(1, 110)
        self.result_table.setColumnWidth(3, 90)
        self.result_table.setColumnWidth(4, 90)
        self.result_table.setColumnWidth(5, 90)
        self.result_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch
        )
        self.result_table.setFixedHeight(300)
        layout.addWidget(self.result_table)
        return card

    def _fill_result(self, result: SpcGenerationResult) -> None:
        xbar_count = sum(
            1 for item in result.results if item.chart_type != CHART_XMR
        )
        xmr_count = len(result.results) - xbar_count
        self.result_label.setText(
            f"模板：{result.template_label}　文件：{Path(result.output_path).name}"
            f"　共 {len(result.results)} 项"
            f"（均值极差图 {xbar_count} 项、单值极差图 {xmr_count} 项）。"
        )
        self.result_table.setRowCount(len(result.results))
        for row, item in enumerate(result.results):
            values = (
                item.index,
                item.chart_label,
                item.sheet_name,
                f"{item.cpk:.2f}",
                f"{item.ppk:.2f}",
                str(item.attempts),
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if column == 2
                    else Qt.AlignCenter
                )
                cell.setToolTip(f"{item.description}（{item.cpk_text}）")
                self.result_table.setItem(row, column, cell)

    def _download(self) -> None:
        if not self.output_path or not Path(self.output_path).is_file():
            QMessageBox.warning(self, "下载失败", "生成文件不存在，请重新生成。")
            return
        source = Path(self.output_path)
        target, _ = QFileDialog.getSaveFileName(
            self, "下载 SPC 文件", source.name, "Excel 文件 (*.xlsx)"
        )
        if not target:
            return
        try:
            shutil.copy2(source, target)
        except OSError as exc:
            QMessageBox.warning(self, "下载失败", f"无法保存文件：{exc}")
            return
        QMessageBox.information(self, "下载成功", f"文件已保存到：\n{target}")

    def on_activated(self) -> None:
        return
