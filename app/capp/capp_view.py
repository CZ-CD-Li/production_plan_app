"""个性化功能定制 · 转 CAPP 质量策划页面。"""

from __future__ import annotations

import shutil
from pathlib import Path

from PyQt5.QtCore import Qt
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
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.capp.generator import (
    CappGenerateError,
    GeneratedResult,
    generate,
    list_history,
)
from app.capp.importer import (
    ControlPlanError,
    ControlPlanImport,
    auto_mapping,
    import_control_plan,
)
from app.capp.rules import TARGET_FIELDS
from app.models import User
from app.personal.drop_area import FileDropArea

PREVIEW_LIMIT = 50


class CappQualityPlanView(QWidget):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.imported: ControlPlanImport | None = None
        self.mapping_combos: dict[str, QComboBox] = {}
        self.result: GeneratedResult | None = None
        self.output_path = ""

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
        self.mapping_card = self._build_mapping_card()
        layout.addWidget(self.mapping_card)
        self.mapping_card.hide()
        self.preview_card = self._build_preview_card()
        layout.addWidget(self.preview_card)
        self.preview_card.hide()
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card)
        self.result_card.hide()
        layout.addWidget(self._build_history_card(), 1)
        layout.addStretch(1)

        scroll.setWidget(content)
        root.addWidget(scroll)
        self.refresh_history()

    # ---------- 上传 ----------
    def _build_upload_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("导入控制计划")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("支持 .xlsx / .xlsm，拖入或点击选择文件")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        self.drop_area = FileDropArea()
        self.drop_area.title_label.setText("拖入控制计划 Excel，或点击右侧按钮选择")
        self.drop_area.file_label.setText("自动识别“零件/过程编号”及多级表头")
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
            imported = import_control_plan(path)
        except ControlPlanError as exc:
            QMessageBox.warning(self, "导入失败", str(exc))
            self.import_label.setText(str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.imported = imported
        self.result = None
        self.output_path = ""
        self.result_card.hide()
        header_rows = f"{imported.header_start}-{imported.header_end} 行"
        self.import_label.setText(
            f"已识别工作表“{imported.sheet_name}”，表头 {header_rows}，"
            f"源列 {len(imported.columns)} 个，数据 {len(imported.rows)} 行。"
            + ("　" + "；".join(imported.warnings) if imported.warnings else "")
        )
        self._fill_mapping(auto_mapping(imported.columns))
        self._fill_preview(imported)
        self.mapping_card.show()
        self.preview_card.show()

    # ---------- 字段映射 ----------
    def _build_mapping_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("字段映射")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("已自动匹配，请逐项确认；匹配不上的可手动选择")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(8)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        for position, target in enumerate(TARGET_FIELDS):
            block = position // 7
            row = position % 7
            holder = QWidget()
            holder_layout = QHBoxLayout(holder)
            holder_layout.setContentsMargins(0, 0, 0, 0)
            holder_layout.setSpacing(8)

            label = QLabel(target)
            label.setObjectName("fieldLabel")
            label.setFixedWidth(196)
            label.setWordWrap(True)
            holder_layout.addWidget(label, 0)

            combo = QComboBox()
            combo.setMinimumWidth(200)
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLength)
            holder_layout.addWidget(combo, 1)
            self.mapping_combos[target] = combo

            grid.addWidget(holder, row, block * 2)
        layout.addLayout(grid)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 4, 0, 0)
        footer.setSpacing(12)
        self.mapping_status = QLabel("")
        self.mapping_status.setObjectName("fieldHint")
        footer.addWidget(self.mapping_status, 1)

        self.generate_button = QPushButton("确认映射并生成质量策划")
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self._handle_generate)
        footer.addWidget(self.generate_button, 0, Qt.AlignRight)
        layout.addLayout(footer)
        return card

    def _fill_mapping(self, mapping: dict[str, int | None]) -> None:
        columns = self.imported.columns if self.imported is not None else []
        for target, combo in self.mapping_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("（未映射）", None)
            for column in columns:
                combo.addItem(column.display, column.index)
            index = combo.findData(mapping.get(target))
            combo.setCurrentIndex(index if index >= 0 else 0)
            combo.blockSignals(False)
        unmapped = sum(1 for value in mapping.values() if value is None)
        self.mapping_status.setText(
            f"自动匹配 {len(mapping) - unmapped} / {len(mapping)}，"
            f"未匹配 {unmapped} 个"
        )

    # ---------- 预览 ----------
    def _build_preview_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("控制计划预览")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.preview_label = QLabel("")
        self.preview_label.setObjectName("countPill")
        header.addWidget(self.preview_label)
        layout.addLayout(header)

        self.preview_table = QTableWidget(0, 0)
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setShowGrid(False)
        self.preview_table.setFocusPolicy(Qt.NoFocus)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.verticalHeader().setDefaultSectionSize(34)
        self.preview_table.horizontalHeader().setFixedHeight(38)
        self.preview_table.horizontalHeader().setHighlightSections(False)
        self.preview_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        self.preview_table.setFixedHeight(230)
        layout.addWidget(self.preview_table)
        return card

    def _fill_preview(self, imported: ControlPlanImport) -> None:
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(imported.columns))
        self.preview_table.setHorizontalHeaderLabels(
            [column.display for column in imported.columns]
        )
        rows = imported.rows[:PREVIEW_LIMIT]
        self.preview_table.setRowCount(len(rows))
        for row_index, values in enumerate(rows):
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                item.setToolTip(value)
                self.preview_table.setItem(row_index, col_index, item)
        for col_index in range(len(imported.columns)):
            self.preview_table.setColumnWidth(col_index, 150)
        self.preview_label.setText(
            f"显示前 {min(len(rows), len(imported.rows))} / {len(imported.rows)} 行"
        )

    # ---------- 生成结果 ----------
    def _build_result_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(22, 14, 22, 14)
        layout.setSpacing(14)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label, 1)

        self.download_button = QPushButton("下载质量策划文件")
        self.download_button.setObjectName("primaryButton")
        self.download_button.setCursor(Qt.PointingHandCursor)
        self.download_button.clicked.connect(self._download)
        layout.addWidget(self.download_button, 0, Qt.AlignVCenter)
        return card

    # ---------- 历史 ----------
    def _build_history_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("已完成的质量策划")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.history_count = QLabel("共 0 个")
        self.history_count.setObjectName("countPill")
        header.addWidget(self.history_count)
        layout.addLayout(header)

        self.history_table = QTableWidget(0, 3)
        self.history_table.setHorizontalHeaderLabels(["产品代号", "控制计划文件", "生成时间"])
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.setShowGrid(False)
        self.history_table.setFocusPolicy(Qt.NoFocus)
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.verticalHeader().setDefaultSectionSize(38)
        self.history_table.horizontalHeader().setFixedHeight(38)
        self.history_table.horizontalHeader().setHighlightSections(False)
        self.history_table.setColumnWidth(0, 150)
        self.history_table.setColumnWidth(1, 380)
        self.history_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch
        )
        self.history_table.setFixedHeight(200)
        layout.addWidget(self.history_table)
        return card

    def refresh_history(self) -> None:
        records = list_history(self.db, self.user)
        self.history_table.setRowCount(0)
        self.history_table.setRowCount(len(records))
        for row_index, record in enumerate(records):
            values = (
                record["product_code"],
                Path(record["source_file"]).name,
                record["created_at"].replace("T", " "),
            )
            for col_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if col_index == 1
                    else Qt.AlignCenter
                )
                item.setToolTip(value)
                self.history_table.setItem(row_index, col_index, item)
        self.history_count.setText(f"共 {len(records)} 个")

    def on_activated(self) -> None:
        self.refresh_history()

    # ---------- 生成与下载 ----------
    def _handle_generate(self) -> None:
        if self.imported is None:
            QMessageBox.warning(self, "无法生成", "请先导入控制计划 Excel。")
            return
        mapping = {
            target: combo.currentData()
            for target, combo in self.mapping_combos.items()
        }
        if mapping.get("过程编号（工序号）") is None:
            QMessageBox.warning(
                self, "无法生成", "请先为“过程编号（工序号）”选择对应的源列。"
            )
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            result = generate(self.imported, mapping, self.user, self.db)
        except CappGenerateError as exc:
            QMessageBox.warning(self, "生成失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.result = result
        self.output_path = result.output_path
        self.result_card.show()
        self.result_label.setText(
            f"已完成：产品代号 {result.product_code}，"
            f"源数据 {result.source_count} 行 → 质量策划 {result.row_count} 行。"
            f"文件：{Path(result.output_path).name}"
        )
        self.refresh_history()
        details = "\n".join(result.warnings) if result.warnings else ""
        QMessageBox.information(
            self,
            "已完成",
            "CAPP 质量策划表已建立完成。\n"
            "请点击“下载质量策划文件”保存后，打开文件确认内容是否正确。"
            + (f"\n\n{details}" if details else ""),
        )

    def _download(self) -> None:
        if not self.output_path or not Path(self.output_path).is_file():
            QMessageBox.warning(self, "下载失败", "生成文件不存在，请重新生成。")
            return
        source = Path(self.output_path)
        target, _ = QFileDialog.getSaveFileName(
            self, "下载质量策划文件", source.name, "Excel 文件 (*.xlsx)"
        )
        if not target:
            return
        try:
            shutil.copy2(source, target)
        except OSError as exc:
            QMessageBox.warning(self, "下载失败", f"无法保存文件：{exc}")
            return
        QMessageBox.information(self, "下载成功", f"文件已保存到：\n{target}")
