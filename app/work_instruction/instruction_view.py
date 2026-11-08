"""个性化功能定制 · 后减作业指导书生成页面。"""

from __future__ import annotations

import os
from pathlib import Path

from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from app.models import User
from app.table_ocr.recognizer import warm_up
from app.work_instruction.generator import (
    DEFAULT_LINE,
    LINE_OPTIONS,
    OUTPUT_DIR,
    InstructionError,
    InstructionResult,
    generate,
)


class _WarmupWorker(QThread):
    """页面打开时预加载 OCR 引擎，首次识别不必等待模型加载。"""

    def run(self) -> None:  # noqa: D102
        try:
            warm_up()
        except Exception:  # noqa: BLE001
            pass


class _GenerateWorker(QThread):
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, paths: list[str], line: str, warmup: QThread | None = None, parent=None) -> None:
        super().__init__(parent)
        self._paths = paths
        self._line = line
        self._warmup = warmup

    def run(self) -> None:  # noqa: D102
        try:
            if self._warmup is not None and self._warmup.isRunning():
                self._warmup.wait()
            result = generate(self._paths, self._line)
        except InstructionError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(f"生成失败：{type(exc).__name__}: {exc}")
        else:
            self.succeeded.emit(result)


class _DropArea(QFrame):
    """支持多选与多文件拖入的投放区域。"""

    files_added = pyqtSignal(list)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("dropArea")
        self.setAcceptDrops(True)
        self.setMinimumHeight(88)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(14)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(3)
        title = QLabel("拖入 PDF 图纸，可一次拖入多个，或点击右侧按钮选择")
        title.setStyleSheet(
            "background:transparent;color:#333333;font-size:13px;font-weight:600;"
        )
        hint = QLabel("总成图纸决定产品代号，其余图纸用于校正工序中的零件明细")
        hint.setObjectName("fieldHint")
        hint.setStyleSheet("background:transparent;")
        text_layout.addWidget(title)
        text_layout.addWidget(hint)
        layout.addLayout(text_layout, 1)

        self.choose_button = QPushButton("选择图纸")
        self.choose_button.setObjectName("secondaryButton")
        self.choose_button.setCursor(Qt.PointingHandCursor)
        self.choose_button.clicked.connect(self._choose_files)
        layout.addWidget(self.choose_button, 0, Qt.AlignVCenter)

    def _choose_files(self) -> None:
        from PyQt5.QtWidgets import QFileDialog

        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择图纸", "", "PDF 图纸 (*.pdf)"
        )
        if paths:
            self.files_added.emit(paths)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_dragging(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        self._set_dragging(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event) -> None:  # noqa: N802
        self._set_dragging(False)
        paths = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf")
        ]
        if paths:
            self.files_added.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _set_dragging(self, dragging: bool) -> None:
        self.setProperty("dragging", "true" if dragging else "false")
        self.style().unpolish(self)
        self.style().polish(self)


class WorkInstructionView(QWidget):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.file_paths: list[str] = []
        self.result: InstructionResult | None = None
        self._worker: _GenerateWorker | None = None
        self._warmup: _WarmupWorker | None = None
        self._busy = False

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
        layout.addWidget(self._build_setting_card())
        self.result_card = self._build_result_card()
        layout.addWidget(self.result_card)
        self.result_card.hide()
        layout.addStretch(1)

        scroll.setWidget(content)
        root.addWidget(scroll)

    # ---------- 导入 ----------
    def _build_upload_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("导入 PDF 图纸")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.count_label = QLabel("未导入")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        layout.addLayout(header)

        drop = _DropArea()
        drop.files_added.connect(self._on_files_added)
        layout.addWidget(drop)

        self.file_list = QListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.setAlternatingRowColors(True)
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.setMinimumHeight(132)
        layout.addWidget(self.file_list)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(10)
        self.remove_button = QPushButton("移除选中")
        self.remove_button.setObjectName("secondaryButton")
        self.remove_button.setCursor(Qt.PointingHandCursor)
        self.remove_button.clicked.connect(self._remove_selected)
        footer.addWidget(self.remove_button)
        clear_button = QPushButton("清空")
        clear_button.setObjectName("secondaryButton")
        clear_button.setCursor(Qt.PointingHandCursor)
        clear_button.clicked.connect(self._clear_files)
        footer.addWidget(clear_button)
        footer.addStretch(1)
        layout.addLayout(footer)
        return card

    def _on_files_added(self, paths: list[str]) -> None:
        added = 0
        for path in paths:
            if Path(path).suffix.lower() != ".pdf":
                continue
            if path in self.file_paths:
                continue
            self.file_paths.append(path)
            self.file_list.addItem(Path(path).name)
            self.file_list.item(self.file_list.count() - 1).setToolTip(path)
            added += 1
        if added:
            self.count_label.setText(f"共 {len(self.file_paths)} 份图纸")

    def _remove_selected(self) -> None:
        for index in sorted(
            (self.file_list.row(item) for item in self.file_list.selectedItems()),
            reverse=True,
        ):
            self.file_list.takeItem(index)
            del self.file_paths[index]
        self.count_label.setText(
            f"共 {len(self.file_paths)} 份图纸" if self.file_paths else "未导入"
        )

    def _clear_files(self) -> None:
        self.file_paths.clear()
        self.file_list.clear()
        self.count_label.setText("未导入")

    # ---------- 设置 ----------
    def _build_setting_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        title = QLabel("生成设置")
        title.setObjectName("sectionTitle")
        layout.addWidget(title)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)
        label = QLabel("生产线")
        label.setObjectName("fieldLabel")
        row.addWidget(label)
        self.line_combo = QComboBox()
        self.line_combo.setObjectName("lineCombo")
        self.line_combo.addItems(list(LINE_OPTIONS))
        self.line_combo.setCurrentText(DEFAULT_LINE)
        self.line_combo.setMinimumWidth(120)
        row.addWidget(self.line_combo)
        hint = QLabel("目前仅内置 C9 标准作业指导书模板，其他生产线选择后需补充模板文件")
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        row.addWidget(hint, 1)
        self.generate_button = QPushButton("生成作业指导书")
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self._on_generate)
        row.addWidget(self.generate_button)
        layout.addLayout(row)

        self.status_label = QLabel(
            "生成的文件保存在数据目录下的“作业指导书”文件夹，文件名格式为 产品代号-作业指导书（初版）.xlsx。"
        )
        self.status_label.setObjectName("fieldHint")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        return card

    def showEvent(self, event) -> None:  # noqa: N802
        super().showEvent(event)
        if self._warmup is None:
            self._warmup = _WarmupWorker(self)
            self._warmup.start()

    def _on_generate(self) -> None:
        if self._busy:
            QMessageBox.information(self, "正在生成", "上一次生成还未完成，请稍候。")
            return
        if not self.file_paths:
            QMessageBox.warning(self, "缺少图纸", "请先导入至少一份 PDF 图纸。")
            return

        self.result_card.hide()
        self.status_label.setText(
            f"正在识别 {len(self.file_paths)} 份图纸并生成作业指导书，图纸较多时需要几分钟。"
        )
        self._worker = _GenerateWorker(
            list(self.file_paths), self.line_combo.currentText(), self._warmup, self
        )
        self._worker.succeeded.connect(self._on_succeeded)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._on_finished)
        self._busy = True
        self.generate_button.setEnabled(False)
        self._worker.start()

    def _on_finished(self) -> None:
        self._busy = False
        self.generate_button.setEnabled(True)

    def _on_failed(self, message: str) -> None:
        self.status_label.setText(message)
        QMessageBox.warning(self, "生成失败", message)

    def _on_succeeded(self, result: InstructionResult) -> None:
        self.result = result
        self.status_label.setText(
            f"已生成 {result.output.name}（{result.line}），共 {len(result.drawings)} 份图纸。"
        )
        self._fill_result()
        self.result_card.show()

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
        self.open_file_button = QPushButton("打开文件")
        self.open_file_button.setObjectName("primaryButton")
        self.open_file_button.setCursor(Qt.PointingHandCursor)
        self.open_file_button.clicked.connect(self._open_file)
        header.addWidget(self.open_file_button)
        open_dir_button = QPushButton("打开所在文件夹")
        open_dir_button.setObjectName("secondaryButton")
        open_dir_button.setCursor(Qt.PointingHandCursor)
        open_dir_button.clicked.connect(self._open_dir)
        header.addWidget(open_dir_button)
        layout.addLayout(header)

        self.path_label = QLabel("")
        self.path_label.setObjectName("fieldHint")
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

        self.summary_list = QListWidget()
        self.summary_list.setObjectName("summaryList")
        self.summary_list.setAlternatingRowColors(True)
        self.summary_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.summary_list.setMinimumHeight(190)
        layout.addWidget(self.summary_list)

        self.warning_label = QLabel("")
        self.warning_label.setObjectName("fieldHint")
        self.warning_label.setWordWrap(True)
        self.warning_label.setStyleSheet("color:#B54708;")
        self.warning_label.hide()
        layout.addWidget(self.warning_label)
        return card

    def _fill_result(self) -> None:
        if self.result is None:
            return
        self.path_label.setText(str(self.result.output))
        self.summary_list.clear()
        self.summary_list.addItem(f"产品代号：{self.result.product_code}")
        for drawing in self.result.drawings:
            state = "正常" if drawing.rows else "未识别明细表"
            self.summary_list.addItem(
                f"图纸 {Path(drawing.path).name} · {drawing.code} · "
                f"{len(drawing.rows)} 行 · {state}"
            )
        for change in self.result.changes:
            self.summary_list.addItem(f"变更 {change}")
        if self.result.warnings:
            self.warning_label.setText("提示：" + "；".join(self.result.warnings))
            self.warning_label.show()
        else:
            self.warning_label.hide()

    def _open_file(self) -> None:
        if self.result is None:
            return
        try:
            os.startfile(str(self.result.output))  # noqa: S606
        except OSError as exc:
            QMessageBox.warning(self, "无法打开", str(exc))

    def _open_dir(self) -> None:
        target = self.result.output.parent if self.result else OUTPUT_DIR
        try:
            os.startfile(str(target))  # noqa: S606
        except OSError as exc:
            QMessageBox.warning(self, "无法打开", str(exc))

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait()
        if self._warmup is not None and self._warmup.isRunning():
            self._warmup.wait()
        super().closeEvent(event)
