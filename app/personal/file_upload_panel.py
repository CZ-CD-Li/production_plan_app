"""个人文件上传表单：拖拽文件、工序联想、UDC 地址校验。"""

from __future__ import annotations

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QCompleter,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from app.models import User
from app.personal.drop_area import FileDropArea
from app.personal.process_names import PROCESS_NAMES
from app.services.line_access import allowed_lines, line_in_text
from app.services.personal_file_service import (
    CATEGORY_CONFIG,
    PersonalFileError,
    PersonalFileService,
    UdcPathError,
)
from app.services.personal_product_service import PersonalProduct, PersonalProductService
from app.widgets.line_combo import LineCombo


class FileUploadPanel(QFrame):
    uploaded = pyqtSignal()

    def __init__(self, db, user: User, category: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("card")
        self.user = user
        self.category = category
        self.config = CATEGORY_CONFIG[category]
        self.file_service = PersonalFileService(db)
        self.product_service = PersonalProductService(db)
        self._products: list[PersonalProduct] = []
        self.require_process = category != "process"
        self._allowed_lines = allowed_lines(user.role)

        root = QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(12)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(10)
        title = QLabel(self.config["upload_title"])
        title.setObjectName("sectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        file_pattern = (
            "产品代号-产线-工序-名称"
            if self.require_process
            else "产品代号-产线-名称"
        )
        hint = QLabel(f"文件名将自动改为：{file_pattern}")
        hint.setObjectName("sectionHint")
        header.addWidget(hint)
        root.addLayout(header)

        self.drop_area = FileDropArea()
        root.addWidget(self.drop_area)

        form = QGridLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(9)
        form.setColumnStretch(1, 1)
        form.setColumnStretch(3, 1)

        form.addWidget(self._label("产品代号"), 0, 0)
        self.product_combo = QComboBox()
        self.product_combo.setEditable(True)
        self.product_combo.setInsertPolicy(QComboBox.NoInsert)
        self.product_combo.currentIndexChanged.connect(self._on_product_changed)
        form.addWidget(self.product_combo, 0, 1)

        form.addWidget(self._label("生产线"), 0, 2)
        self.line_combo = LineCombo()
        self.line_combo.set_allowed(self._allowed_lines)
        form.addWidget(self.line_combo, 0, 3)

        self.process_combo: QComboBox | None = None
        if self.require_process:
            form.addWidget(self._label("工序"), 1, 0)
            self.process_combo = QComboBox()
            self.process_combo.setEditable(True)
            self.process_combo.setInsertPolicy(QComboBox.NoInsert)
            self.process_combo.addItems(PROCESS_NAMES)
            self.process_combo.setCurrentIndex(-1)
            self.process_combo.setEditText("")
            completer = self.process_combo.completer()
            if completer is not None:
                completer.setCompletionMode(QCompleter.PopupCompletion)
                completer.setFilterMode(Qt.MatchContains)
                completer.setCaseSensitivity(Qt.CaseInsensitive)
            form.addWidget(self.process_combo, 1, 1)

        for combo, placeholder in (
            (self.product_combo, "选择或输入产品代号"),
            (self.line_combo, "选择或输入生产线"),
            (self.process_combo, "输入工序，自动联想"),
        ):
            if combo is None:
                continue
            edit = combo.lineEdit()
            if edit is not None:
                edit.setPlaceholderText(placeholder)

        if self.require_process:
            form.addWidget(self._label(self.config["name_label"]), 1, 2)
            self.name_edit = QLineEdit()
            self.name_edit.setPlaceholderText(f"请输入{self.config['label']}名称")
            form.addWidget(self.name_edit, 1, 3)
        else:
            form.addWidget(self._label(self.config["name_label"]), 1, 0)
            self.name_edit = QLineEdit()
            self.name_edit.setPlaceholderText(f"请输入{self.config['label']}名称")
            form.addWidget(self.name_edit, 1, 1, 1, 3)

        form.addWidget(self._label("保存地址"), 2, 0)
        self.save_path_edit = QLineEdit()
        self.save_path_edit.setPlaceholderText(
            "选择文件保存目录，测试阶段支持本地路径"
        )
        form.addWidget(self.save_path_edit, 2, 1, 1, 2)

        browse = QPushButton("选择目录")
        browse.setObjectName("secondaryButton")
        browse.setFixedWidth(104)
        browse.setCursor(Qt.PointingHandCursor)
        browse.clicked.connect(self._browse_save_dir)
        form.addWidget(browse, 2, 3, Qt.AlignRight)

        root.addLayout(form)

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(12)
        self.status_label = QLabel("")
        self.status_label.setObjectName("fieldHint")
        footer.addWidget(self.status_label, 1)

        self.upload_button = QPushButton("上传")
        self.upload_button.setObjectName("primaryButton")
        self.upload_button.setCursor(Qt.PointingHandCursor)
        self.upload_button.clicked.connect(self._handle_upload)
        footer.addWidget(self.upload_button, 0, Qt.AlignRight)
        root.addLayout(footer)

        self._reload_products()

    @staticmethod
    def _label(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _reload_products(self) -> None:
        self._products = self.product_service.products_for(self.user.name)
        current_code = self.product_combo.currentText().strip()
        current_line = self.line_combo.currentText().strip()

        self.product_combo.blockSignals(True)
        self.product_combo.clear()
        for product in self._products:
            self.product_combo.addItem(product.product_code)
        self.product_combo.setEditText(current_code)
        self.product_combo.blockSignals(False)

        self.line_combo.set_allowed(self._allowed_lines, current_line)

    def on_activated(self) -> None:
        self._reload_products()

    def select_product(self, product_code: str, line: str = "") -> None:
        self.product_combo.setEditText(product_code)
        if line and self.line_combo.set_line(line):
            return
        self._apply_line_for_product(product_code)

    def _on_product_changed(self) -> None:
        self._apply_line_for_product(self.product_combo.currentText().strip())

    def _apply_line_for_product(self, product_code: str) -> None:
        self.line_combo.set_line("")
        for product in self._products:
            if product.product_code != product_code:
                continue
            code = line_in_text(product.production_line, self._allowed_lines)
            if code:
                self.line_combo.set_line(code)
                return

    def _browse_save_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择保存地址")
        if directory:
            self.save_path_edit.setText(directory)

    def _handle_upload(self) -> None:
        file_path = self.drop_area.file_path()
        if not file_path:
            self._warn("请选择或拖入要上传的文件。")
            return
        line = self.line_combo.current_line()
        if line is None:
            self._warn(
                "请输入有效的产线代号（可在下拉列表中联想选择，"
                f"当前账号可用：{self._allowed_lines[0]} ~ "
                f"{self._allowed_lines[-1]}）。"
            )
            return
        try:
            record = self.file_service.save_upload(
                category=self.category,
                user=self.user,
                source_file=file_path,
                save_dir=self.save_path_edit.text().strip(),
                product_code=self.product_combo.currentText().strip(),
                production_line=line,
                process=(
                    self.process_combo.currentText().strip()
                    if self.process_combo is not None
                    else ""
                ),
                name=self.name_edit.text().strip(),
            )
        except UdcPathError as exc:
            self._warn(str(exc))
            return
        except PersonalFileError as exc:
            self._warn(str(exc))
            return

        self.status_label.setText(f"已保存为 {record.file_name}")
        QMessageBox.information(
            self,
            "上传成功",
            f"{self.config['label']}上传成功。\n文件已保存为：{record.file_name}",
        )
        self.drop_area.set_file("")
        self.name_edit.clear()
        if self.process_combo is not None:
            self.process_combo.setEditText("")
        self.uploaded.emit()

    def _warn(self, message: str) -> None:
        self.status_label.setText(message)
        QMessageBox.warning(self, "上传失败", message)
