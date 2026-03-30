import csv
import importlib.util
import mimetypes
import os
import ssl
import time
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from typing import Iterable, List, Optional

import smtplib

def _configure_qt_env() -> None:
    """
    On some macOS setups Qt can't load the 'cocoa' platform plugin unless
    plugin/library paths are present *before* importing PySide6.
    """
    try:
        spec = importlib.util.find_spec("PySide6")
        if spec is None or not spec.submodule_search_locations:
            return
        pyside_dir = str(list(spec.submodule_search_locations)[0])
        qt_dir = os.path.join(pyside_dir, "Qt")
        plugins_dir = os.path.join(qt_dir, "plugins")
        platforms_dir = os.path.join(plugins_dir, "platforms")
        qt_lib_dir = os.path.join(qt_dir, "lib")

        if os.path.isdir(plugins_dir):
            os.environ.setdefault("QT_PLUGIN_PATH", plugins_dir)
        if os.path.isdir(platforms_dir):
            os.environ.setdefault("QT_QPA_PLATFORM_PLUGIN_PATH", platforms_dir)

        # Help the dynamic loader find Qt frameworks shipped with PySide6.
        for env_key in ("DYLD_LIBRARY_PATH", "DYLD_FRAMEWORK_PATH"):
            if os.path.isdir(qt_lib_dir):
                cur = os.environ.get(env_key, "")
                if qt_lib_dir not in cur.split(":"):
                    os.environ[env_key] = qt_lib_dir + (":" + cur if cur else "")
    except Exception:
        return


_configure_qt_env()


from PySide6.QtCore import QThread, Signal, Qt  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

try:
    import certifi  # type: ignore
except Exception:  # pragma: no cover
    certifi = None


PROVIDERS = {
    "Gmail": {"host": "smtp.gmail.com", "port": 587, "starttls": True},
    "Mail.ru": {"host": "smtp.mail.ru", "port": 587, "starttls": True},
    "Custom": {"host": "", "port": 587, "starttls": True},
}


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def read_emails_from_csv(path: str) -> List[str]:
    emails: List[str] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel

        reader = csv.reader(f, dialect)
        rows = list(reader)

    if not rows:
        return []

    # Header support: email
    first = [c.strip().lower() for c in rows[0]]
    start_idx = 0
    email_col = 0
    if "email" in first:
        email_col = first.index("email")
        start_idx = 1

    for r in rows[start_idx:]:
        if not r:
            continue
        if email_col >= len(r):
            continue
        e = r[email_col].strip()
        if not e:
            continue
        emails.append(e)

    # de-dup, keep order
    seen = set()
    out: List[str] = []
    for e in emails:
        key = e.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(e)
    return out


def build_message(
    from_email: str,
    to_email: str,
    subject: str,
    body_text: str,
    is_html: bool,
    attachments: Iterable[str],
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject
    if is_html:
        plain_fallback = "Это HTML-письмо. Если вы видите этот текст, включите отображение HTML."
        msg.set_content(plain_fallback)
        msg.add_alternative(body_text or "", subtype="html")
    else:
        msg.set_content(body_text or "")

    for p in attachments:
        if not p:
            continue
        filename = os.path.basename(p)
        ctype, encoding = mimetypes.guess_type(p)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(p, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)

    return msg


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    starttls: bool
    verify_tls: bool
    username: str
    password: str
    from_email: str


class SenderWorker(QThread):
    progress = Signal(int, int)  # sent, total
    log = Signal(str, str, str)  # ts, email, message
    counters = Signal(int, int, int)  # sent, failed, remaining
    finished_ok = Signal()
    finished_err = Signal(str)

    def __init__(
        self,
        smtp_cfg: SmtpConfig,
        emails: List[str],
        subject: str,
        body_text: str,
        is_html: bool,
        attachments: List[str],
        emails_per_minute: int,
    ) -> None:
        super().__init__()
        self._smtp_cfg = smtp_cfg
        self._emails = emails
        self._subject = subject
        self._body_text = body_text
        self._is_html = is_html
        self._attachments = attachments
        self._epm = max(1, int(emails_per_minute))
        self._stop = False

    def request_stop(self) -> None:
        self._stop = True

    def _connect(self) -> smtplib.SMTP:
        smtp = smtplib.SMTP(self._smtp_cfg.host, self._smtp_cfg.port, timeout=30)
        smtp.ehlo()
        if self._smtp_cfg.starttls:
            if self._smtp_cfg.verify_tls:
                if certifi is not None:
                    ctx = ssl.create_default_context(cafile=certifi.where())
                else:
                    ctx = ssl.create_default_context()
            else:
                ctx = ssl._create_unverified_context()
            smtp.starttls(context=ctx)
            smtp.ehlo()
        smtp.login(self._smtp_cfg.username, self._smtp_cfg.password)
        return smtp

    def run(self) -> None:
        total = len(self._emails)
        sent = 0
        failed = 0
        delay_s = 60.0 / float(self._epm)

        smtp: Optional[smtplib.SMTP] = None
        try:
            smtp = self._connect()
        except Exception as e:
            self.finished_err.emit(f"SMTP login/connect failed: {e}")
            return

        try:
            for idx, to_email in enumerate(self._emails, start=1):
                if self._stop:
                    self.log.emit(now_ts(), "-", "Stopped by user")
                    break

                msg = build_message(
                    from_email=self._smtp_cfg.from_email,
                    to_email=to_email,
                    subject=self._subject,
                    body_text=self._body_text,
                    is_html=self._is_html,
                    attachments=self._attachments,
                )

                try:
                    smtp.send_message(msg)
                    sent += 1
                    self.log.emit(now_ts(), to_email, "SENT")
                except Exception as e:
                    failed += 1
                    self.log.emit(now_ts(), to_email, f"FAILED: {e}")
                    # Reconnect once on failure (common when server drops connection)
                    try:
                        smtp.quit()
                    except Exception:
                        pass
                    try:
                        smtp = self._connect()
                    except Exception:
                        # keep going; next send will try reconnect again
                        pass

                remaining = max(0, total - idx)
                self.progress.emit(sent + failed, total)
                self.counters.emit(sent, failed, remaining)

                if idx < total and not self._stop:
                    time.sleep(delay_s)

            self.finished_ok.emit()
        except Exception as e:
            self.finished_err.emit(str(e))
        finally:
            try:
                if smtp is not None:
                    smtp.quit()
            except Exception:
                pass


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Mail Notifier (macOS)")
        self.resize(980, 720)

        self.worker: Optional[SenderWorker] = None
        self.emails: List[str] = []
        self.csv_path: Optional[str] = None

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        layout.addWidget(self._build_smtp_box())
        layout.addWidget(self._build_campaign_box())
        layout.addWidget(self._build_controls_box())
        layout.addWidget(self._build_stats_box())
        layout.addWidget(self._build_log_box())

        self._apply_provider_defaults()
        self._set_running(False)

    def _build_smtp_box(self) -> QGroupBox:
        box = QGroupBox("SMTP настройки")
        form = QFormLayout(box)

        self.provider = QComboBox()
        self.provider.addItems(list(PROVIDERS.keys()))
        self.provider.currentTextChanged.connect(self._apply_provider_defaults)

        self.smtp_host = QLineEdit()
        self.smtp_port = QSpinBox()
        self.smtp_port.setRange(1, 65535)
        self.smtp_starttls = QCheckBox("STARTTLS")
        self.smtp_starttls.setChecked(True)
        self.smtp_verify_tls = QCheckBox("Проверять TLS сертификат (рекомендуется)")
        self.smtp_verify_tls.setChecked(True)

        self.username = QLineEdit()
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.from_email = QLineEdit()

        form.addRow("Провайдер:", self.provider)
        form.addRow("SMTP host:", self.smtp_host)
        form.addRow("SMTP port:", self.smtp_port)
        form.addRow("", self.smtp_starttls)
        form.addRow("", self.smtp_verify_tls)
        form.addRow("Логин (email):", self.username)
        form.addRow("Пароль / app-password:", self.password)
        form.addRow("From (если отличается):", self.from_email)

        return box

    def _build_campaign_box(self) -> QGroupBox:
        box = QGroupBox("Письмо и база")
        grid = QGridLayout(box)

        self.csv_label = QLabel("CSV не выбран")
        self.btn_pick_csv = QPushButton("Выбрать CSV…")
        self.btn_pick_csv.clicked.connect(self.pick_csv)

        self.subject = QLineEdit()
        self.body = QPlainTextEdit()
        self.body.setPlaceholderText("Текст письма (обычный текст) или HTML (если включишь режим HTML).")

        self.body_is_html = QCheckBox("HTML режим")
        self.body_is_html.setChecked(False)

        self.attachments = QListWidget()
        self.btn_add_attach = QPushButton("Добавить вложение…")
        self.btn_remove_attach = QPushButton("Убрать выбранное")
        self.btn_add_attach.clicked.connect(self.add_attachment)
        self.btn_remove_attach.clicked.connect(self.remove_attachment)

        self.rate = QSpinBox()
        self.rate.setRange(1, 600)
        self.rate.setValue(30)
        self.rate.setSuffix(" писем/мин")

        grid.addWidget(QLabel("База:"), 0, 0)
        grid.addWidget(self.csv_label, 0, 1)
        grid.addWidget(self.btn_pick_csv, 0, 2)

        grid.addWidget(QLabel("Тема:"), 1, 0)
        grid.addWidget(self.subject, 1, 1, 1, 2)

        grid.addWidget(QLabel("Текст:"), 2, 0, Qt.AlignmentFlag.AlignTop)
        grid.addWidget(self.body, 2, 1, 1, 2)
        grid.addWidget(self.body_is_html, 3, 1)

        grid.addWidget(QLabel("Вложения:"), 4, 0, Qt.AlignmentFlag.AlignTop)
        grid.addWidget(self.attachments, 4, 1, 2, 1)
        btn_col = QVBoxLayout()
        btn_col.addWidget(self.btn_add_attach)
        btn_col.addWidget(self.btn_remove_attach)
        btn_col.addStretch(1)
        grid.addLayout(btn_col, 4, 2, 1, 1)

        grid.addWidget(QLabel("Скорость:"), 6, 0)
        grid.addWidget(self.rate, 6, 1)

        return box

    def _build_controls_box(self) -> QGroupBox:
        box = QGroupBox("Управление")
        row = QHBoxLayout(box)
        self.btn_start = QPushButton("Старт")
        self.btn_stop = QPushButton("Стоп")
        self.btn_start.clicked.connect(self.start_sending)
        self.btn_stop.clicked.connect(self.stop_sending)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)

        row.addWidget(self.btn_start)
        row.addWidget(self.btn_stop)
        row.addWidget(self.progress, 1)
        return box

    def _build_stats_box(self) -> QGroupBox:
        box = QGroupBox("Статистика")
        row = QHBoxLayout(box)
        self.lbl_total = QLabel("Всего: 0")
        self.lbl_sent = QLabel("Отправлено: 0")
        self.lbl_failed = QLabel("Ошибки: 0")
        self.lbl_remaining = QLabel("Осталось: 0")
        row.addWidget(self.lbl_total)
        row.addWidget(self.lbl_sent)
        row.addWidget(self.lbl_failed)
        row.addWidget(self.lbl_remaining)
        row.addStretch(1)
        return box

    def _build_log_box(self) -> QGroupBox:
        box = QGroupBox("Лог отправки")
        v = QVBoxLayout(box)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Time", "Email", "Status"])
        self.table.horizontalHeader().setStretchLastSection(True)
        v.addWidget(self.table)
        return box

    def _apply_provider_defaults(self) -> None:
        p = self.provider.currentText()
        cfg = PROVIDERS.get(p, PROVIDERS["Custom"])
        self.smtp_host.setText(cfg["host"])
        self.smtp_port.setValue(cfg["port"])
        self.smtp_starttls.setChecked(bool(cfg["starttls"]))

    def _set_running(self, running: bool) -> None:
        self.btn_start.setEnabled(not running)
        self.btn_stop.setEnabled(running)
        self.btn_pick_csv.setEnabled(not running)
        self.btn_add_attach.setEnabled(not running)
        self.btn_remove_attach.setEnabled(not running)

        self.provider.setEnabled(not running)
        self.smtp_host.setEnabled(not running)
        self.smtp_port.setEnabled(not running)
        self.smtp_starttls.setEnabled(not running)
        self.smtp_verify_tls.setEnabled(not running)
        self.username.setEnabled(not running)
        self.password.setEnabled(not running)
        self.from_email.setEnabled(not running)
        self.subject.setEnabled(not running)
        self.body.setEnabled(not running)
        self.body_is_html.setEnabled(not running)
        self.rate.setEnabled(not running)

    def _append_log(self, ts: str, email: str, status: str) -> None:
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(ts))
        self.table.setItem(r, 1, QTableWidgetItem(email))
        self.table.setItem(r, 2, QTableWidgetItem(status))
        self.table.scrollToBottom()

    def pick_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать CSV", "", "CSV (*.csv);;All files (*)")
        if not path:
            return
        try:
            emails = read_emails_from_csv(path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка CSV", str(e))
            return

        self.csv_path = path
        self.emails = emails
        self.csv_label.setText(f"{os.path.basename(path)} (email: {len(emails)})")
        self.lbl_total.setText(f"Всего: {len(emails)}")
        self.lbl_sent.setText("Отправлено: 0")
        self.lbl_failed.setText("Ошибки: 0")
        self.lbl_remaining.setText(f"Осталось: {len(emails)}")
        self.progress.setValue(0)

    def add_attachment(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Добавить вложения", "", "All files (*)")
        for p in paths:
            if not p:
                continue
            item = QListWidgetItem(p)
            self.attachments.addItem(item)

    def remove_attachment(self) -> None:
        for item in self.attachments.selectedItems():
            self.attachments.takeItem(self.attachments.row(item))

    def _get_attachments(self) -> List[str]:
        out: List[str] = []
        for i in range(self.attachments.count()):
            out.append(self.attachments.item(i).text())
        return out

    def _validate(self) -> Optional[str]:
        if not self.csv_path or not self.emails:
            return "Выбери CSV с email адресами."
        host = self.smtp_host.text().strip()
        if not host:
            return "Укажи SMTP host."
        if not self.username.text().strip():
            return "Укажи логин (email)."
        if not self.password.text():
            return "Укажи пароль / app-password."
        if not self.subject.text().strip():
            return "Укажи тему письма."
        return None

    def start_sending(self) -> None:
        err = self._validate()
        if err:
            QMessageBox.warning(self, "Проверка", err)
            return

        from_email = self.from_email.text().strip() or self.username.text().strip()
        smtp_cfg = SmtpConfig(
            host=self.smtp_host.text().strip(),
            port=int(self.smtp_port.value()),
            starttls=bool(self.smtp_starttls.isChecked()),
            verify_tls=bool(self.smtp_verify_tls.isChecked()),
            username=self.username.text().strip(),
            password=self.password.text(),
            from_email=from_email,
        )

        self.table.setRowCount(0)
        self.progress.setValue(0)
        self._set_running(True)

        self.worker = SenderWorker(
            smtp_cfg=smtp_cfg,
            emails=self.emails,
            subject=self.subject.text().strip(),
            body_text=self.body.toPlainText(),
            is_html=bool(self.body_is_html.isChecked()),
            attachments=self._get_attachments(),
            emails_per_minute=int(self.rate.value()),
        )
        self.worker.log.connect(self._append_log)
        self.worker.progress.connect(self._on_progress)
        self.worker.counters.connect(self._on_counters)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_err.connect(self._on_finished_err)
        self.worker.start()

    def stop_sending(self) -> None:
        if self.worker is not None:
            self.worker.request_stop()

    def _on_progress(self, done: int, total: int) -> None:
        if total <= 0:
            self.progress.setValue(0)
            return
        self.progress.setValue(int((done / total) * 100))

    def _on_counters(self, sent: int, failed: int, remaining: int) -> None:
        self.lbl_sent.setText(f"Отправлено: {sent}")
        self.lbl_failed.setText(f"Ошибки: {failed}")
        self.lbl_remaining.setText(f"Осталось: {remaining}")

    def _cleanup_worker(self) -> None:
        if self.worker is None:
            return
        self.worker.wait(1000)
        self.worker = None
        self._set_running(False)

    def _on_finished_ok(self) -> None:
        self._append_log(now_ts(), "-", "DONE")
        self._cleanup_worker()

    def _on_finished_err(self, msg: str) -> None:
        self._append_log(now_ts(), "-", f"ERROR: {msg}")
        QMessageBox.critical(self, "Ошибка", msg)
        self._cleanup_worker()


def main() -> None:
    app = QApplication([])
    w = MainWindow()
    w.show()
    app.exec()


if __name__ == "__main__":
    main()

