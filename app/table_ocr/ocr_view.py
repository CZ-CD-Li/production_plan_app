"""个性化功能定制 · 表格识图页面。"""

from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
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

from app.models import User
from app.personal.drop_area import FileDropArea
from app.table_ocr.recognizer import (
    TableOcrError,
    TableRecognition,
    recognize_pdf,
    warm_up,
)

COLUMNS = ("序号", "代号", "版本", "名称", "数量")


class _WarmupWorker(QThread):
    """页面打开时预加载 OCR 引擎，首次识别不必再等待模型加载。"""

    def run(self) -> None:  # noqa: D102
        try:
            warm_up()
        except Exception:  # noqa: BLE001
            pass


class _RecognizeWorker(QThread):
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, path: str, warmup: QThread | None = None, parent=None) -> None:
        super().__init__(parent)
        self._path = path
        self._warmup = warmup

    def run(self) -> None:  # noqa: D102
        try:
            if self._warmup is not None and self._warmup.isRunning():
                self._warmup.wait()
            result = recognize_pdf(self._path)
        except TableOcrError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(f"识别失败：{type(exc).__name__}: {exc}")
        else:
            self.succeeded.emit(result)


class TableRecognitionView(QWidget):
    def __init__(self, db, user: User, parent=None) -> None:
        super().__init__(parent)
        self.db = db
        self.user = user
        self.file_path = ""
        self.recognition: TableRecognition | None = None
        self._worker: _RecognizeWorker | None = None
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
        title = QLabel("导入 PDF 图纸")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        hint = QLabel("支持 .pdf：图纸、扫描件、文字型文件均可识别")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        layout.addLayout(header)

        note = QLabel("识别结果需与图纸对照后下发。")
        note.setObjectName("fieldHint")
        note.setStyleSheet("color:#B54708;font-weight:600;")
        layout.addWidget(note)

        self.drop_area = FileDropArea()
        self.drop_area.title_label.setText("拖入 PDF 图纸，或点击右侧按钮选择")
        self.drop_area.file_label.setText(
            "自动识别图纸表格中的序号、代号、版本、名称、数量列，并按序号升序排列"
        )
        self.drop_area.file_selected.connect(self._on_file_selected)
        layout.addWidget(self.drop_area)

        self.import_label = QLabel(
            "图纸类文件先用线条定位表格，再只识别表格区域，通常几秒即可完成。"
        )
        self.import_label.setObjectName("fieldHint")
        self.import_label.setWordWrap(True)
        layout.addWidget(self.import_label)
        return card

    def showEvent(self, event) -> None:  # noqa: N802
        super().showEvent(event)
        if self._warmup is None:
            self._warmup = _WarmupWorker(self)
            self._warmup.start()

    def _on_file_selected(self, path: str) -> None:
        if not path:
            return
        if Path(path).suffix.lower() != ".pdf":
            QMessageBox.warning(self, "无法导入", "仅支持 .pdf 文件。")
            return
        if self._busy:
            QMessageBox.information(self, "正在识别", "上一个文件仍在识别，请稍候。")
            return

        self.file_path = path
        self.recognition = None
        self.result_card.hide()
        self.import_label.setText(
            f"正在识别 {Path(path).name}，图纸越大耗时越长，请稍候。"
        )
        self._start_worker(path)

    def _start_worker(self, path: str) -> None:
        self._worker = _RecognizeWorker(path, self._warmup, self)
        self._worker.succeeded.connect(self._on_recognized)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._on_worker_finished)
        self._busy = True
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self._worker.start()

    def _on_worker_finished(self) -> None:
        if self._busy:
            QApplication.restoreOverrideCursor()
            self._busy = False

    def _on_failed(self, message: str) -> None:
        self._on_worker_finished()
        self.import_label.setText(message)
        QMessageBox.warning(self, "识别失败", message)

    def _on_recognized(self, result: TableRecognition) -> None:
        self._on_worker_finished()
        self.recognition = result
        self._fill_table()
        self.result_card.show()

        pages = []
        if result.text_pages:
            pages.append(f"文字层解析 {result.text_pages} 页")
        if result.ocr_pages:
            pages.append(f"OCR 识别 {result.ocr_pages} 页")
        method = "，".join(pages) if pages else f"共 {result.page_count} 页"
        self.import_label.setText(
            f"已识别 {Path(self.file_path).name}：{method}，"
            f"共 {len(result.rows)} 行，已按序号升序排列。"
        )
        self.warning_label.setText(
            "提示：" + "；".join(result.warnings) if result.warnings else ""
        )
        self.warning_label.setVisible(bool(result.warnings))

    # ---------- 结果 ----------
    def _build_result_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        title = QLabel("识别结果")
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        self.count_label = QLabel("")
        self.count_label.setObjectName("countPill")
        header.addWidget(self.count_label)
        self.copy_button = QPushButton("复制表格内容")
        self.copy_button.setObjectName("primaryButton")
        self.copy_button.setCursor(Qt.PointingHandCursor)
        self.copy_button.clicked.connect(self._copy_table)
        header.addWidget(self.copy_button)
        layout.addLayout(header)

        hint = QLabel(
            "复制内容为制表符分隔，粘贴到 Excel 会自动分列；"
            "也可拖动选中单元格后用 Ctrl+C 复制局部内容。"
        )
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(list(COLUMNS))
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setFixedHeight(38)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.setColumnWidth(0, 64)
        self.table.setColumnWidth(1, 180)
        self.table.setColumnWidth(2, 72)
        self.table.setColumnWidth(4, 72)
        self.table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.Stretch
        )
        self.table.setMinimumHeight(380)
        layout.addWidget(self.table)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 2, 0, 0)
        footer.setSpacing(12)
        self.copy_status = QLabel("")
        self.copy_status.setObjectName("fieldHint")
        footer.addWidget(self.copy_status, 1)
        layout.addLayout(footer)

        self.warning_label = QLabel("")
        self.warning_label.setObjectName("fieldHint")
        self.warning_label.setWordWrap(True)
        self.warning_label.setStyleSheet("color:#B54708;")
        self.warning_label.hide()
        layout.addWidget(self.warning_label)
        return card

    def _fill_table(self) -> None:
        rows = self.recognition.rows if self.recognition is not None else []
        self.table.setRowCount(len(rows))
        for row, item in enumerate(rows):
            values = (
                item.index,
                item.code,
                item.version,
                item.name,
                item.quantity,
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(value)
                cell.setTextAlignment(
                    Qt.AlignLeft | Qt.AlignVCenter
                    if column in (1, 3)
                    else Qt.AlignCenter
                )
                cell.setToolTip(value)
                self.table.setItem(row, column, cell)
        self.count_label.setText(f"共 {len(rows)} 行")
        self.copy_status.setText("")

    def _copy_table(self) -> None:
        if self.recognition is None or not self.recognition.rows:
            QMessageBox.warning(self, "无法复制", "请先导入并识别 PDF 表格。")
            return
        lines = ["\t".join(COLUMNS)]
        for item in self.recognition.rows:
            lines.append(
                "\t".join(
                    (
                        item.index,
                        item.code,
                        item.version,
                        item.name,
                        item.quantity,
                    )
                )
            )
        QApplication.clipboard().setText("\r\n".join(lines))
        self.copy_status.setText(
            f"已复制 {len(self.recognition.rows)} 行（含表头），"
            "可直接粘贴到 Excel。"
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait()
        if self._warmup is not None and self._warmup.isRunning():
            self._warmup.wait()
        super().closeEvent(event)
