"""个性化功能定制 · 作业指导书转控制计划页面。"""

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

from app.capp.control_plan_writer import (
    ControlPlanGenerateError,
    generate_control_plan,
    list_templates,
    preview_text,
)
from app.capp.work_instruction import (
    PLAN_HEADERS,
    WorkInstructionError,
    WorkInstructionParse,
    ensure_grease_process,
    parse_work_instruction,
)
from app.config import DATA_DIR
from app.personal.drop_area import FileDropArea

OUTPUT_DIR = DATA_DIR / "control_plan_output"
PREVIEW_LIMIT = 60


class ControlPlanView(QWidget):
    def __init__(self, db, user, parent=None) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.parsed: WorkInstructionParse | None = None
        self.result_path = ""
        self.template_error = ""
        try:
            self.layouts = list_templates()
        except (ControlPlanGenerateError, OSError) as exc:
            self.layouts = []
            self.template_error = str(exc)

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
        self.template_card = self._build_template_card()
        layout.addWidget(self.template_card)
        self.template_card.hide()
        self.preview_card = self._build_preview_card()
        layout.addWidget(self.preview_card)
        self.preview_card.hide()
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card)
        self.result_card.hide()
        layout.addWidget(self._build_note_card())
        layout.addStretch(1)

        scroll.setWidget(content)
        root.addWidget(scroll)

        self._fill_templates()

    # ---------- 上传 ----------
    def _build_upload_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("导入作业指导书")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("支持 .xlsx / .xlsm，逐个工序工作表解析")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        self.drop_area = FileDropArea()
        self.drop_area.title_label.setText("拖入作业指导书 Excel，或点击右侧按钮选择")
        self.drop_area.file_label.setText("识别“制造条件确认”“品质特性确认”后生成控制计划")
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
            parsed = parse_work_instruction(path)
        except WorkInstructionError as exc:
            QMessageBox.warning(self, "导入失败", str(exc))
            self.import_label.setText(str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()

        self.parsed = parsed
        self.result_path = ""
        self.result_card.hide()

        grease_added = False
        if self.user.role.strip() == "装配":
            try:
                grease_added = ensure_grease_process(parsed)
            except WorkInstructionError as exc:
                parsed.warnings.append(f"未能自动补充涂润滑脂工序：{exc}")

        product = " ".join(
            item for item in (parsed.product_name, parsed.product_code) if item
        )
        text = (
            f"{product}　" if product else ""
        ) + (
            f"已解析工序 {len(parsed.blocks)} 个，控制计划数据 {len(parsed.rows)} 行；"
            f"跳过非工序工作表 {len(parsed.skipped_sheets)} 个。"
        )
        if grease_added:
            text += "　已自动补充“涂润滑脂”工序（序号 1）。"
        if parsed.warnings:
            text += "　" + "；".join(parsed.warnings[:5])
            if len(parsed.warnings) > 5:
                text += f"；其余 {len(parsed.warnings) - 5} 条略"
        self.import_label.setText(text)

        self._fill_templates()
        self._refresh_preview()
        self.template_card.show()
        self.preview_card.show()

    # ---------- 模板 ----------
    def _build_template_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("主机厂控制计划模板")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("模板来自《控制计划格式样板-3》")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)
        label = QLabel("选择主机厂")
        label.setObjectName("fieldLabel")
        label.setFixedWidth(96)
        row.addWidget(label)
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(260)
        self.template_combo.currentIndexChanged.connect(self._refresh_preview)
        row.addWidget(self.template_combo, 0)
        self.template_hint = QLabel("")
        self.template_hint.setObjectName("fieldHint")
        self.template_hint.setWordWrap(True)
        row.addWidget(self.template_hint, 1)
        layout.addLayout(row)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 4, 0, 0)
        footer.setSpacing(12)
        self.generate_status = QLabel(
            "生成时按模板保留表头、logo、复选框与打印设置，数据后保留 3 个空行。"
        )
        self.generate_status.setObjectName("fieldHint")
        self.generate_status.setWordWrap(True)
        footer.addWidget(self.generate_status, 1)
        self.generate_button = QPushButton("生成控制计划")
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self._handle_generate)
        footer.addWidget(self.generate_button, 0, Qt.AlignRight)
        layout.addLayout(footer)
        return card

    def _fill_templates(self) -> None:
        supported = [layout for layout in self.layouts if layout.supported]
        unsupported = [layout for layout in self.layouts if not layout.supported]
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        for layout in supported:
            self.template_combo.addItem(layout.display_name, layout.sheet_name)
        self.template_combo.blockSignals(False)
        if self.template_combo.count():
            self.template_combo.setCurrentIndex(0)
        if self.template_error:
            self.template_hint.setText(f"模板不可用：{self.template_error}")
        elif unsupported:
            self.template_hint.setText(
                f"可用 {len(supported)} 个主机厂模板；"
                f"{'、'.join(layout.display_name for layout in unsupported)} 格式不支持"
            )
        else:
            self.template_hint.setText(f"可用 {len(supported)} 个主机厂模板")

    def _current_layout(self):
        sheet_name = self.template_combo.currentData()
        if not sheet_name:
            return None
        for layout in self.layouts:
            if layout.sheet_name == sheet_name:
                return layout
        return None

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

        self.preview_table = QTableWidget(0, len(PLAN_HEADERS))
        self.preview_table.setHorizontalHeaderLabels(list(PLAN_HEADERS))
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setShowGrid(False)
        self.preview_table.setFocusPolicy(Qt.NoFocus)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.verticalHeader().setDefaultSectionSize(32)
        self.preview_table.horizontalHeader().setFixedHeight(38)
        self.preview_table.horizontalHeader().setHighlightSections(False)
        self.preview_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Interactive
        )
        widths = (74, 130, 150, 52, 110, 110, 90, 200, 120, 60, 70, 150, 150, 80, 70, 120, 110, 110)
        for index, width in enumerate(widths):
            self.preview_table.setColumnWidth(index, width)
        self.preview_table.setFixedHeight(280)
        layout.addWidget(self.preview_table)
        return card

    def _refresh_preview(self) -> None:
        if self.parsed is None:
            return
        layout = self._current_layout()
        symbols = layout.symbols if layout is not None else {}
        rows = self.parsed.rows[:PREVIEW_LIMIT]
        self.preview_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row.values):
                item = QTableWidgetItem(preview_text(value, symbols))
                item.setTextAlignment(Qt.AlignCenter)
                item.setToolTip(preview_text(value, symbols))
                self.preview_table.setItem(row_index, col_index, item)
        self.preview_label.setText(
            f"显示前 {len(rows)} / {len(self.parsed.rows)} 行"
        )
        if layout is not None:
            missing = f"；不写入列：{'、'.join(layout.missing_names)}" if layout.missing_names else ""
            self.generate_status.setText(
                f"{layout.display_name}：表头第 {layout.marker_row} 行为数据首行，"
                f"模板列 {len(layout.columns)} 个{missing}。"
                "生成时保留模板表头、logo、复选框与打印设置，数据后保留 3 个空行。"
            )

    # ---------- 结果 ----------
    def _build_result_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(22, 14, 22, 14)
        layout.setSpacing(14)

        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label, 1)

        self.download_button = QPushButton("下载控制计划文件")
        self.download_button.setObjectName("primaryButton")
        self.download_button.setCursor(Qt.PointingHandCursor)
        self.download_button.clicked.connect(self._download)
        layout.addWidget(self.download_button, 0, Qt.AlignVCenter)
        return card

    def _build_note_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 16, 22, 16)
        layout.setSpacing(6)
        title = QLabel("生成说明")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
        note = QLabel(
            "1. 每个含“品质特性确认”的工作表视为一道工序，按制造条件确认、品质特性确认两段提取数据；\n"
            "2. 过程特性含 Z / G 时按模板第 28 行 A~D 的特殊特性符号加“/”前缀；\n"
            "3. 维护保养、防错管理、审核行按 VBA 规则纵向/横向合并，行高按字数与列宽自动计算；\n"
            "4. 生成后请用 Excel 打开确认内容，确认无误再下发。"
        )
        note.setObjectName("fieldHint")
        note.setWordWrap(True)
        layout.addWidget(note)
        return card

    # ---------- 生成与下载 ----------
    def _handle_generate(self) -> None:
        if self.parsed is None:
            QMessageBox.warning(self, "无法生成", "请先导入作业指导书。")
            return
        layout = self._current_layout()
        if layout is None:
            QMessageBox.warning(self, "无法生成", "请选择主机厂控制计划模板。")
            return
        output = OUTPUT_DIR / f"{Path(self.parsed.file_path).stem}-{layout.display_name}.xlsx"
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.generate_button.setEnabled(False)
        try:
            generate_control_plan(self.parsed, layout.sheet_name, output)
        except ControlPlanGenerateError as exc:
            QMessageBox.warning(self, "生成失败", str(exc))
            return
        except OSError as exc:
            QMessageBox.warning(self, "生成失败", f"无法写入文件：{exc}")
            return
        finally:
            QApplication.restoreOverrideCursor()
            self.generate_button.setEnabled(True)

        self.result_path = str(output)
        self.result_card.show()
        self.result_label.setText(
            f"已生成：{layout.display_name}　{len(self.parsed.rows)} 行数据　"
            f"文件：{output.name}"
        )
        QMessageBox.information(
            self,
            "已完成",
            "控制计划已生成。\n"
            "请点击“下载控制计划文件”保存后，用 Excel 打开确认内容。",
        )

    def _download(self) -> None:
        if not self.result_path or not Path(self.result_path).is_file():
            QMessageBox.warning(self, "下载失败", "生成文件不存在，请重新生成。")
            return
        source = Path(self.result_path)
        target, _ = QFileDialog.getSaveFileName(
            self, "下载控制计划文件", source.name, "Excel 文件 (*.xlsx)"
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
