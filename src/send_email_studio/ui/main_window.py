from pathlib import Path
from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtCore import Qt
from PySide6.QtGui import QCursor, QPainter
from PySide6.QtWidgets import (QAbstractButton, QComboBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QMainWindow, QPlainTextEdit, QProgressBar, QPushButton,
                               QScrollArea, QSpinBox, QVBoxLayout, QWidget)
import qtawesome as qta

from ..application.send_service import send_from_excel
from ..domain.models import MailSettings
from ..infrastructure.excel_reader import read_excel_rows


class _Worker(QObject):
    finished = Signal(int)
    failed = Signal(str)
    progress = Signal(int, int, str)

    def __init__(self, settings: MailSettings, excel: str):
        super().__init__()
        self.settings, self.excel = settings, excel

    @Slot()
    def run(self):
        try:
            # 发送器会通过 progress 回调同时传递进度和成功信息；界面只接收这一条通道，
            # 避免 log 与 progress 各追加一次而造成重复显示。
            count = send_from_excel(self.settings, self.excel, None, self.progress.emit)
            self.finished.emit(count)
        except Exception as exc:
            self.failed.emit(str(exc))


class _NoWheelSpinBox(QSpinBox):
    """端口输入框不响应滚轮，避免页面滚动时误改端口。"""

    def wheelEvent(self, event):
        event.ignore()


class _ArrowComboBox(QComboBox):
    """参考 cloud-report 的下拉框：滚轮不改值，箭头随弹出状态切换。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._popup_open = False

    def wheelEvent(self, event):
        event.ignore()

    def showPopup(self):
        self._popup_open = True
        self.update()
        super().showPopup()

    def hidePopup(self):
        super().hidePopup()
        self._popup_open = False
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        icon_name = "fa5s.chevron-up" if self._popup_open else "fa5s.chevron-down"
        color = "#1A2A43" if self.isEnabled() else "#9AA4B3"
        painter = QPainter(self)
        qta.icon(icon_name, color=color).paint(painter, self.width() - 24, (self.height() - 12) // 2, 12, 12)
        painter.end()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量发送邮件工具")
        self.setMinimumSize(780, 620)
        self.resize(1180, 820)
        self.thread: QThread | None = None
        self.worker: _Worker | None = None
        self._build_ui()

    def _build_ui(self):
        form = QVBoxLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(14)
        self.smtp_host = QLineEdit("smtp.exmail.qq.com")
        self.smtp_port = _NoWheelSpinBox(); self.smtp_port.setRange(1, 65535); self.smtp_port.setValue(587)
        self.smtp_user = QLineEdit(); self.smtp_password = QLineEdit(); self.smtp_password.setEchoMode(QLineEdit.Password)
        self.password_toggle_action = self.smtp_password.addAction(
            qta.icon("fa5s.eye", color="#52627A"), QLineEdit.ActionPosition.TrailingPosition
        )
        self.password_toggle_action.setToolTip("显示密码")
        self.password_toggle_action.triggered.connect(self._toggle_password_visibility)
        self.subject = QLineEdit(); self.subject.setPlaceholderText("若 Excel 存在“主题”列，将按行覆盖此处主题")
        smtp_row = _field_row((_field("SMTP 服务器", self.smtp_host), _field("SMTP 端口", self.smtp_port)))
        account_row = _field_row((_field("邮箱账号", self.smtp_user), _field("邮箱密码", self.smtp_password)))
        form.addWidget(smtp_row); form.addWidget(account_row); form.addWidget(_field("邮件主题", self.subject))

        self.content = QPlainTextEdit("<p>你好，这是使用 Python 自动发送的 HTML 格式邮件。</p>")
        self.content.setPlaceholderText("填写 HTML 正文；如选择下方模板，则以模板为准。")
        self.template = QLineEdit(); self.template.setPlaceholderText("可选：HTML、TXT 或 DOCX 模板；变量格式为 {列名}")
        template_btn = QPushButton("浏览"); template_btn.setIcon(qta.icon("fa5s.file-alt", color="#FFFFFF"))
        template_btn.clicked.connect(self._choose_template)
        template_row = QHBoxLayout(); template_row.addWidget(self.template); template_row.addWidget(template_btn)
        template_widget = QWidget(); template_widget.setLayout(template_row)
        self.excel = QLineEdit(); self.excel.setPlaceholderText("选择包含收件/抄送列的 Excel 文件")
        excel_btn = QPushButton("浏览"); excel_btn.setIcon(qta.icon("fa5s.file-excel", color="#FFFFFF")); excel_btn.clicked.connect(self._choose_excel)
        excel_row = QHBoxLayout(); excel_row.addWidget(self.excel); excel_row.addWidget(excel_btn)
        excel_widget = QWidget(); excel_widget.setLayout(excel_row)
        self.mode = _ArrowComboBox(); self.mode.addItem("每个收件人单独发送", "individual"); self.mode.addItem("一封邮件发送给全部收件人", "grouped")
        form.addWidget(_field("邮件正文 HTML", self.content))
        template_hint = QLabel("正文与模板二选一：选择模板后会忽略上方正文，并可使用 {列名} 替换 Excel 当前行数据。")
        template_hint.setObjectName("ruleDetails"); template_hint.setWordWrap(True)
        form.addWidget(template_hint); form.addWidget(_field("Word/HTML 模板（可选）", template_widget))
        form.addWidget(_field("Excel 文件", excel_widget))
        hint = QLabel("收件列名包含“收件”；抄送列名包含“抄送”。单元格邮箱支持逗号、分号和顿号分隔。模板占位符格式：{列名}。")
        hint.setWordWrap(True)
        self.send_button = QPushButton("发送邮件"); self.send_button.setIcon(qta.icon("fa5s.paper-plane", color="#FFFFFF")); self.send_button.clicked.connect(self._start)
        self.progress = QProgressBar(); self.progress.setRange(0, 100)
        self.status = QLabel("就绪"); self.status.setObjectName("status"); self.status.setProperty("state", "idle"); self.status.setWordWrap(True)
        self.log = QPlainTextEdit(); self.log.setReadOnly(True)
        header = QFrame(); header.setObjectName("header")
        header_layout = QVBoxLayout(header); header_layout.setContentsMargins(0, 0, 0, 16); header_layout.setSpacing(2)
        title = QLabel("SEND EMAIL STUDIO"); title.setObjectName("title")
        subtitle = QLabel("OPERATIONS MAILER"); subtitle.setObjectName("subtitle")
        header_layout.addWidget(title); header_layout.addWidget(subtitle)
        form_frame = QFrame(); form_frame.setObjectName("section")
        form_layout = QVBoxLayout(form_frame); form_layout.setContentsMargins(28, 20, 28, 22); form_layout.setSpacing(16)
        section_heading = QHBoxLayout(); number = QLabel("01"); number.setObjectName("sectionNumber"); section_title = QLabel("邮件配置"); section_title.setObjectName("sectionTitle")
        section_heading.addWidget(number); section_heading.addWidget(section_title); section_heading.addStretch(1)
        config_hint = QLabel("提示：SMTP 587 端口使用 STARTTLS；Excel 列名包含“收件”或“抄送”即可识别，邮箱支持逗号、分号和顿号分隔。")
        config_hint.setObjectName("ruleDetails"); config_hint.setWordWrap(True)
        form_layout.addLayout(section_heading); form_layout.addLayout(form); form_layout.addWidget(config_hint)
        rules_frame = QFrame(); rules_frame.setObjectName("section")
        rules_layout = QVBoxLayout(rules_frame); rules_layout.setContentsMargins(28, 20, 28, 22); rules_layout.setSpacing(14)
        rules_heading = QHBoxLayout(); number_rules = QLabel("02"); number_rules.setObjectName("sectionNumber"); title_rules = QLabel("发送规则"); title_rules.setObjectName("sectionTitle")
        rules_heading.addWidget(number_rules); rules_heading.addWidget(title_rules); rules_heading.addStretch(1)
        rules_layout.addLayout(rules_heading)
        mode_label = QLabel("收件人发送方式"); mode_label.setObjectName("fieldLabel")
        rules_layout.addWidget(mode_label); rules_layout.addWidget(self.mode); rules_layout.addWidget(hint)
        delivery_frame = QFrame(); delivery_frame.setObjectName("section")
        delivery_layout = QVBoxLayout(delivery_frame); delivery_layout.setContentsMargins(28, 20, 28, 22); delivery_layout.setSpacing(14)
        delivery_heading = QHBoxLayout(); number2 = QLabel("03"); number2.setObjectName("sectionNumber"); title2 = QLabel("执行与结果"); title2.setObjectName("sectionTitle")
        delivery_heading.addWidget(number2); delivery_heading.addWidget(title2); delivery_heading.addStretch(1)
        delivery_layout.addLayout(delivery_heading); delivery_layout.addWidget(self.send_button); delivery_layout.addWidget(self.progress); delivery_layout.addWidget(self.status); delivery_layout.addWidget(self.log)
        content = QWidget(); content.setObjectName("content")
        layout = QVBoxLayout(content); layout.setContentsMargins(56, 32, 56, 44); layout.setSpacing(20)
        layout.addWidget(header); layout.addWidget(form_frame); layout.addWidget(rules_frame); layout.addWidget(delivery_frame); layout.addStretch(1)
        scroll = QScrollArea(); scroll.setObjectName("mainScrollArea"); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame); scroll.setWidget(content)
        self.setCentralWidget(scroll)
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow, QScrollArea#mainScrollArea { background: #F7F3E7; color: #182843; }
            QWidget#content { background: #F7F3E7; }
            QFrame#header { border-bottom: 2px solid #1A2A43; }
            QFrame#section { background: #FFFDFC; border: 2px solid #1A2A43; border-radius: 0; }
            QLabel { color: #182843; font-family: "Microsoft YaHei UI"; }
            QLabel#title { font-family: "Bahnschrift SemiBold"; font-size: 28px; font-weight: 800; color: #172742; }
            QLabel#subtitle { color: #00A88E; font-family: "Consolas"; font-size: 11px; font-weight: 700; letter-spacing: 1px; }
            QLabel#sectionNumber { color: #6464E9; font-family: "Consolas"; font-size: 12px; font-weight: 700; }
            QLabel#sectionTitle { font-size: 18px; font-weight: 800; margin-left: 8px; }
            QLabel#fieldLabel { color: #66738B; font-size: 12px; }
            QLabel#ruleDetails { color: #66738B; font-size: 12px; }
            QLabel#status[state="error"] { color: #B42318; font-weight: 700; }
            QLabel#status[state="success"] { color: #007D68; font-weight: 700; }
            QLabel#status[state="warning"] { color: #A65B00; font-weight: 700; }
            QLineEdit, QComboBox, QSpinBox { min-height: 38px; border: 2px solid #1A2A43; border-radius: 0; padding: 1px 10px; background: #FFFFFF; color: #1B2940; selection-background-color: #6464E9; }
            QPlainTextEdit { border: 2px solid #1A2A43; border-radius: 0; padding: 8px 10px; background: #FFFFFF; color: #1B2940; font-family: "Consolas"; font-size: 12px; selection-background-color: #6464E9; }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QPlainTextEdit:focus { border-color: #5D5CE2; background: #FFFEFC; }
            QComboBox::drop-down { border: 0; width: 28px; }
            QComboBox QAbstractItemView { background: #FFFFFF; color: #1B2940; border: 2px solid #1A2A43; selection-background-color: #E7E6FF; }
            QSpinBox::up-button, QSpinBox::down-button { width: 0; height: 0; border: 0; }
            QPushButton { min-height: 38px; border: 2px solid #1A2A43; border-radius: 0; padding: 0 18px; background: #5D5CE2; color: #FFFFFF; font-family: "Microsoft YaHei UI"; font-weight: 800; }
            QPushButton:hover:enabled { background: #4847C7; }
            QPushButton:pressed:enabled { background: #3A399E; }
            QPushButton:disabled { background: #ECEBE5; border-color: #A8B0BC; color: #9AA4B3; }
            QProgressBar { min-height: 20px; border: 2px solid #1A2A43; border-radius: 0; text-align: center; background: #E7EDF4; color: #182843; font-weight: 800; }
            QProgressBar::chunk { background: #00A88E; }
            QScrollBar:vertical { width: 10px; background: #F7F3E7; margin: 0; }
            QScrollBar::handle:vertical { min-height: 36px; background: #9AA4B3; border: 2px solid #F7F3E7; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)
        for button in self.findChildren(QAbstractButton):
            button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def _choose_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xlsm *.xls)")
        if path: self.excel.setText(path)

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择模板文件", "", "模板文件 (*.html *.htm *.txt *.docx)")
        if path: self.template.setText(path)

    def _toggle_password_visibility(self):
        showing = self.smtp_password.echoMode() == QLineEdit.EchoMode.Password
        self.smtp_password.setEchoMode(QLineEdit.EchoMode.Normal if showing else QLineEdit.EchoMode.Password)
        self.password_toggle_action.setIcon(
            qta.icon("fa5s.eye-slash" if showing else "fa5s.eye", color="#52627A")
        )
        self.password_toggle_action.setToolTip("隐藏密码" if showing else "显示密码")

    def _start(self):
        if not self.smtp_user.text().strip() or not self.smtp_password.text():
            self._set_status("邮箱账号和密码不能为空。", "error"); return
        subject_text = self.subject.text().strip()
        if not self.excel.text().strip():
            self._set_status("Excel 文件不能为空。", "error"); return
        excel_path = Path(self.excel.text().strip())
        if not excel_path.is_file(): self._set_status("Excel 文件不存在，请重新选择。", "error"); return
        if not subject_text:
            try:
                columns, _ = read_excel_rows(excel_path)
            except Exception as exc:
                self._set_status(f"Excel 文件读取失败：{exc}", "error"); return
            if not any("主题" in column for column in columns):
                self._set_status("邮件主题和 Excel 文件不能为空。", "error"); return
        template_path = Path(self.template.text().strip()) if self.template.text().strip() else None
        if template_path and not template_path.is_file(): self._set_status("模板文件不存在，请重新选择。", "error"); return
        settings = MailSettings(self.smtp_host.text().strip(), self.smtp_port.value(), self.smtp_user.text().strip(), self.smtp_password.text(), subject_text, self.content.toPlainText(), template_path, self.mode.currentData())
        self.log.clear(); self.progress.setValue(0); self.send_button.setEnabled(False)
        self.thread = QThread(self); self.worker = _Worker(settings, str(excel_path)); self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run); self.worker.progress.connect(self._on_progress); self.worker.finished.connect(self._done); self.worker.failed.connect(self._failed)
        self.worker.finished.connect(self.thread.quit); self.worker.failed.connect(self.thread.quit); self.thread.finished.connect(self._thread_finished); self.thread.start()

    @Slot(int, int, str)
    def _on_progress(self, current: int, total: int, message: str):
        if message: self.log.appendPlainText(message)
        if total: self.progress.setValue(int(current * 100 / total)); self.status.setText(f"已完成 {current}/{total}")

    def _done(self, count: int): self.progress.setValue(100); self._set_status(f"发送完成，共 {count} 封。", "success")
    def _failed(self, message: str): self._set_status(f"发送失败：{message}", "error"); self.log.appendPlainText(f"错误：{message}")
    def _thread_finished(self): self.send_button.setEnabled(True); self.thread = None; self.worker = None

    def _set_status(self, message: str, state: str = "idle"):
        self.status.setText(message)
        self.status.setProperty("state", state)
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)
        self.status.update()


def _field(label_text: str, control: QWidget) -> QWidget:
    container = QWidget()
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    label = QLabel(label_text)
    label.setObjectName("fieldLabel")
    layout.addWidget(label)
    layout.addWidget(control)
    return container


def _field_row(fields: tuple[QWidget, QWidget]) -> QWidget:
    row = QWidget()
    layout = QHBoxLayout(row)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(18)
    for field in fields:
        layout.addWidget(field, 1)
    return row
