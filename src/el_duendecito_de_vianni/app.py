from __future__ import annotations

import logging
import os
import sys
import time
from datetime import date
from pathlib import Path

from PySide6.QtCore import QDate, QLockFile, QPoint, QStandardPaths, QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QAction, QBrush, QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSystemTrayIcon,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .config import AppConfig, ensure_directories, get_app_root, load_config, save_config
from .credentials import has_mercury_password, load_mercury_password, save_mercury_password
from .logging_utils import configure_logging
from .mercury import MercuryAutomationError, run_mercury_export
from .office import open_folder, print_file, set_start_with_windows
from .processor import DocumentProcessor, RunReport
from .spreadsheet import has_employee_rows


APP_STYLESHEET = """
QMainWindow, QDialog {
    background: #f7f1e5;
    color: #283324;
    font-family: "Segoe UI";
    font-size: 10pt;
}
QDialog QLabel, QDialog QCheckBox {
    color: #243325;
    font-weight: 500;
}
QDialog#ConfigDialog {
    background: #fff8e8;
}
QWidget#ConfigPathRow {
    background: transparent;
}
QLabel#TitleLabel {
    color: #245236;
    font-size: 26px;
    font-weight: 800;
}
QLabel#SubtitleLabel {
    color: #6a4f24;
    font-size: 12px;
}
QLabel#StatusCard, QLabel#PathCard {
    background: #fffaf0;
    border: 1px solid #d9c99f;
    border-radius: 8px;
    color: #243325;
    padding: 12px;
}
QWidget#ActionPanel {
    background: #fffaf0;
    border: 1px solid #d9c99f;
    border-radius: 8px;
    padding: 12px;
}
QLabel#ActionTitle {
    color: #245236;
    font-size: 17px;
    font-weight: 800;
}
QLabel#ActionHint {
    color: #6a4f24;
}
QLabel#ProgressLabel {
    color: #245236;
    font-weight: 650;
}
QProgressBar {
    background: #fffdf7;
    border: 1px solid #cdbb8a;
    border-radius: 6px;
    color: #243325;
    min-height: 18px;
    text-align: center;
}
QProgressBar::chunk {
    background: #5b8c48;
    border-radius: 5px;
}
QTextEdit {
    background: #fffdf7;
    border: 1px solid #cdbb8a;
    border-radius: 8px;
    color: #243325;
    padding: 8px;
    selection-background-color: #5b8c48;
    selection-color: #ffffff;
}
QPushButton {
    background: #315f3b;
    border: 1px solid #214429;
    border-radius: 6px;
    color: white;
    font-weight: 600;
    padding: 8px 12px;
}
QPushButton:hover {
    background: #3d7449;
}
QPushButton:pressed {
    background: #24482d;
}
QLineEdit, QSpinBox, QComboBox, QDateEdit {
    background: #fffdf7;
    border: 1px solid #cdbb8a;
    border-radius: 6px;
    color: #1f2b20;
    padding: 6px;
}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QDateEdit:focus {
    border: 1px solid #4f7c43;
    background: #ffffff;
}
QComboBox QAbstractItemView {
    background: #fffdf7;
    color: #1f2b20;
    selection-background-color: #dfeecf;
    selection-color: #1f2b20;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #8e7a4d;
    border-radius: 3px;
    background: #fffdf7;
}
QCheckBox::indicator:checked {
    background: #315f3b;
    border: 1px solid #214429;
}
QPushButton#PathButton {
    background: #d6a540;
    border: 1px solid #9b7426;
    color: #243325;
    min-width: 34px;
    padding: 6px 8px;
}
QPushButton#PathButton:hover {
    background: #e0b956;
}
QMenu {
    background: #fffaf0;
    border: 1px solid #cdbb8a;
}
QMenu::item {
    padding: 7px 24px;
}
QMenu::item:selected {
    background: #dfeecf;
}
"""


class MercuryWorker(QObject):
    finished = Signal(object)
    failed = Signal(object)

    def __init__(self, config: AppConfig, password: str, report_date: date):
        super().__init__()
        self.config = config
        self.password = password
        self.report_date = report_date

    @Slot()
    def run(self) -> None:
        try:
            result = run_mercury_export(self.config, self.password, report_date=self.report_date)
        except Exception as exc:
            self.failed.emit(exc)
        else:
            self.finished.emit(result)


def make_icon() -> QIcon:
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#f0d7a3"))
    painter.drawEllipse(18, 23, 28, 30)

    painter.setBrush(QColor("#d8a15f"))
    left_ear = QPainterPath()
    left_ear.moveTo(18, 35)
    left_ear.lineTo(5, 28)
    left_ear.lineTo(17, 43)
    left_ear.closeSubpath()
    painter.drawPath(left_ear)
    right_ear = QPainterPath()
    right_ear.moveTo(46, 35)
    right_ear.lineTo(59, 28)
    right_ear.lineTo(47, 43)
    right_ear.closeSubpath()
    painter.drawPath(right_ear)

    hat = QPainterPath()
    hat.moveTo(13, 28)
    hat.cubicTo(17, 12, 35, 5, 50, 13)
    hat.cubicTo(42, 15, 41, 22, 48, 27)
    hat.cubicTo(37, 23, 24, 23, 13, 28)
    painter.setBrush(QColor("#2f6f3e"))
    painter.drawPath(hat)
    painter.setBrush(QColor("#d6a540"))
    painter.drawEllipse(45, 10, 8, 8)

    painter.setPen(QPen(QColor("#5b4028"), 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawPoint(QPoint(27, 37))
    painter.drawPoint(QPoint(38, 37))
    painter.setPen(QPen(QColor("#7b4e2f"), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawArc(27, 38, 11, 8, 200 * 16, 140 * 16)

    painter.setPen(QPen(QColor("#1f4a2c"), 2))
    painter.setBrush(Qt.BrushStyle.NoBrush)
    painter.drawEllipse(18, 23, 28, 30)
    painter.end()
    return QIcon(pixmap)


def make_dancing_elf_frame(frame: int) -> QPixmap:
    pixmap = QPixmap(42, 42)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    step = -2 if frame % 2 == 0 else 2
    arm_lift = -5 if frame % 2 == 0 else 3
    foot_lift = 4 if frame % 2 == 0 else -1

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#d8a15f"))
    left_ear = QPainterPath()
    left_ear.moveTo(13, 17)
    left_ear.lineTo(4, 13)
    left_ear.lineTo(12, 23)
    left_ear.closeSubpath()
    painter.drawPath(left_ear)
    right_ear = QPainterPath()
    right_ear.moveTo(29, 17)
    right_ear.lineTo(38, 13)
    right_ear.lineTo(30, 23)
    right_ear.closeSubpath()
    painter.drawPath(right_ear)

    painter.setBrush(QColor("#f0d7a3"))
    painter.drawEllipse(12 + step, 11, 18, 20)

    hat = QPainterPath()
    hat.moveTo(10 + step, 15)
    hat.cubicTo(13 + step, 4, 26 + step, 2, 33 + step, 8)
    hat.cubicTo(27 + step, 9, 27 + step, 14, 33 + step, 17)
    hat.cubicTo(25 + step, 14, 17 + step, 13, 10 + step, 15)
    painter.setBrush(QColor("#2f6f3e"))
    painter.drawPath(hat)
    painter.setBrush(QColor("#d6a540"))
    painter.drawEllipse(31 + step, 5, 5, 5)

    painter.setPen(QPen(QColor("#5b4028"), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawPoint(QPoint(18 + step, 20))
    painter.drawPoint(QPoint(24 + step, 20))
    painter.setPen(QPen(QColor("#7b4e2f"), 1, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawArc(18 + step, 21, 7, 5, 200 * 16, 140 * 16)

    painter.setPen(QPen(QColor("#315f3b"), 4, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawLine(21 + step, 30, 21 - step, 37)
    painter.setPen(QPen(QColor("#d6a540"), 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawLine(14 + step, 29, 8, 25 + arm_lift)
    painter.drawLine(28 + step, 29, 35, 25 - arm_lift)
    painter.setPen(QPen(QColor("#24482d"), 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    painter.drawLine(18 + step, 36, 12, 39 + foot_lift)
    painter.drawLine(24 + step, 36, 31, 39 - foot_lift)
    painter.end()
    return pixmap


class ConfigDialog(QDialog):
    def __init__(self, config: AppConfig, parent=None):
        super().__init__(parent)
        self.setObjectName("ConfigDialog")
        self.setWindowTitle("Configuracion")
        self.setWindowIcon(make_icon())
        self.setStyleSheet(APP_STYLESHEET)
        self.config = config
        layout = QFormLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setHorizontalSpacing(14)
        layout.setVerticalSpacing(10)
        self.downloads = self._path_row(config.downloads_folder)
        self.templates = self._path_row(config.template_folder)
        self.output = self._path_row(config.output_folder)
        self.work_schedule_lookup = self._file_row(config.work_schedule_lookup)
        self.mercury_url = QLineEdit(config.mercury_url)
        self.mercury_username = QLineEdit(config.mercury_username)
        self.mercury_company = QLineEdit(config.mercury_company)
        self.mercury_companies = QLineEdit(config.mercury_companies)
        self.mercury_report_name = QLineEdit(config.mercury_report_name)
        self.mercury_password = QLineEdit()
        self.mercury_password.setEchoMode(QLineEdit.EchoMode.Password)
        if has_mercury_password():
            self.mercury_password.setPlaceholderText("Contrasena guardada; escriba una nueva para cambiarla")
        else:
            self.mercury_password.setPlaceholderText("Contrasena de Mercury")
        self.mercury_headless = QCheckBox("Ejecutar Mercury invisible")
        self.mercury_headless.setChecked(config.mercury_headless)
        self.interval = QSpinBox()
        self.interval.setRange(1, 1440)
        self.interval.setValue(config.scan_interval_minutes)
        self.ask_delete = QCheckBox("Preguntar antes de borrar el archivo descargado")
        self.ask_delete.setChecked(config.ask_before_delete_original)
        self.start_minimized = QCheckBox("Iniciar minimizado en la bandeja del sistema")
        self.start_minimized.setChecked(config.start_minimized_to_tray)
        self.start_windows = QCheckBox("Iniciar automaticamente con Windows")
        self.start_windows.setChecked(config.start_with_windows)
        self.printer = QComboBox()
        self.printer.setEditable(True)
        self.printer.addItem(config.selected_printer)
        layout.addRow("Carpeta de Descargas", self.downloads)
        layout.addRow("Carpeta de plantillas", self.templates)
        layout.addRow("Carpeta de salida", self.output)
        layout.addRow("Tabla de horarios", self.work_schedule_lookup)
        layout.addRow("Mercury URL", self.mercury_url)
        layout.addRow("Usuario Mercury", self.mercury_username)
        layout.addRow("Companias Mercury", self.mercury_companies)
        layout.addRow("Reporte Mercury", self.mercury_report_name)
        layout.addRow("Contrasena Mercury", self.mercury_password)
        layout.addRow(self.mercury_headless)
        layout.addRow("Intervalo (minutos)", self.interval)
        layout.addRow("Impresora", self.printer)
        layout.addRow(self.ask_delete)
        layout.addRow(self.start_minimized)
        layout.addRow(self.start_windows)
        buttons = QHBoxLayout()
        save = QPushButton("Guardar")
        cancel = QPushButton("Cancelar")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addWidget(save)
        buttons.addWidget(cancel)
        layout.addRow(buttons)

    def _path_row(self, value: str) -> QWidget:
        widget = QWidget()
        widget.setObjectName("ConfigPathRow")
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        edit = QLineEdit(value)
        button = QPushButton("...")
        button.setObjectName("PathButton")
        button.clicked.connect(lambda: self._choose_folder(edit))
        layout.addWidget(edit)
        layout.addWidget(button)
        widget.edit = edit  # type: ignore[attr-defined]
        return widget

    def _choose_folder(self, edit: QLineEdit) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta", edit.text())
        if folder:
            edit.setText(folder)

    def _file_row(self, value: str) -> QWidget:
        widget = QWidget()
        widget.setObjectName("ConfigPathRow")
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        edit = QLineEdit(value)
        button = QPushButton("...")
        button.setObjectName("PathButton")
        button.clicked.connect(lambda: self._choose_file(edit))
        layout.addWidget(edit)
        layout.addWidget(button)
        widget.edit = edit  # type: ignore[attr-defined]
        return widget

    def _choose_file(self, edit: QLineEdit) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar tabla de horarios",
            edit.text(),
            "Excel (*.xlsx *.xlsm *.xltx *.xltm)",
        )
        if file_path:
            edit.setText(file_path)

    def updated_config(self) -> AppConfig:
        return AppConfig(
            downloads_folder=self.downloads.edit.text(),  # type: ignore[attr-defined]
            template_folder=self.templates.edit.text(),  # type: ignore[attr-defined]
            output_folder=self.output.edit.text(),  # type: ignore[attr-defined]
            imported_folder=self.config.imported_folder,
            logs_folder=self.config.logs_folder,
            work_schedule_lookup=self.work_schedule_lookup.edit.text(),  # type: ignore[attr-defined]
            mercury_url=self.mercury_url.text(),
            mercury_username=self.mercury_username.text(),
            mercury_company=self.mercury_companies.text().split(";")[0].strip() or self.mercury_company.text(),
            mercury_companies=self.mercury_companies.text(),
            mercury_report_name=self.mercury_report_name.text(),
            mercury_headless=self.mercury_headless.isChecked(),
            scan_interval_minutes=self.interval.value(),
            ask_before_delete_original=self.ask_delete.isChecked(),
            selected_printer=self.printer.currentText(),
            start_minimized_to_tray=self.start_minimized.isChecked(),
            start_with_windows=self.start_windows.isChecked(),
            monitoring_enabled=self.config.monitoring_enabled,
        )


class MainWindow(QMainWindow):
    def __init__(self, config: AppConfig, log_path: Path):
        super().__init__()
        self.setStyleSheet(APP_STYLESHEET)
        self.config = config
        self.log_path = log_path
        self.processor = DocumentProcessor(config)
        self.last_scan = "Nunca"
        self.last_spreadsheet = ""
        self.last_employees = 0
        self.dance_frame = 0
        self.monitoring = config.monitoring_enabled
        self.exiting = False
        self.mercury_thread: QThread | None = None
        self.mercury_worker: MercuryWorker | None = None
        self.mercury_busy = False
        self.busy_status_timer = QTimer(self)
        self.busy_status_timer.timeout.connect(self._advance_busy_status)
        self.busy_started_at = 0.0
        self.busy_status_index = 0
        self.busy_status_phases = (
            "Conectando con Mercury",
            "Seleccionando la compania",
            "Abriendo Recursos Humanos",
            "Preparando el reporte",
            "Aplicando el filtro de fecha",
            "Esperando la respuesta de Mercury",
            "Descargando el reporte",
        )
        self.setWindowTitle("El duendecito de Vianni")
        self.setWindowIcon(make_icon())
        self._build_ui()
        self._build_tray()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.scan_now)
        self.dance_timer = QTimer(self)
        self.dance_timer.timeout.connect(self._advance_dancing_elf)
        self._sync_timer()
        self.refresh_status()
        QTimer.singleShot(1200, self.show_startup_greeting)

    def _build_ui(self) -> None:
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        title = QLabel("El duendecito de Vianni")
        title.setObjectName("TitleLabel")
        subtitle = QLabel("Un ayudante discreto para preparar documentos de entradas y salidas")
        subtitle.setObjectName("SubtitleLabel")
        layout.addWidget(title)
        layout.addWidget(subtitle)
        actions = QHBoxLayout()
        actions.setSpacing(12)
        actions.addWidget(self._workflow_panel("Entradas", "Procesar nuevas entradas de un solo dia.", self.run_entradas))
        actions.addWidget(self._workflow_panel("Salidas", "Preparar salidas de un solo dia.", self.run_salidas))
        layout.addLayout(actions)
        progress_row = QHBoxLayout()
        progress_row.setSpacing(8)
        self.busy_elf = QLabel()
        self.busy_elf.setFixedSize(42, 42)
        self.busy_elf.setPixmap(make_dancing_elf_frame(0))
        self.busy_elf.setVisible(False)
        self.progress_label = QLabel("Listo para trabajar.")
        self.progress_label.setObjectName("ProgressLabel")
        self.progress_label.setWordWrap(True)
        progress_row.addWidget(self.busy_elf)
        progress_row.addWidget(self.progress_label, 1)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addLayout(progress_row)
        layout.addWidget(self.progress_bar)

        self.advanced_toggle = QPushButton("Mostrar opciones avanzadas")
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.clicked.connect(self._toggle_advanced_options)
        layout.addWidget(self.advanced_toggle)

        self.advanced_widget = QWidget()
        advanced_layout = QVBoxLayout(self.advanced_widget)
        advanced_layout.setContentsMargins(0, 0, 0, 0)
        advanced_layout.setSpacing(12)
        self.status = QLabel()
        self.status.setObjectName("StatusCard")
        self.paths = QLabel()
        self.paths.setObjectName("PathCard")
        self.paths.setWordWrap(True)
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        advanced_layout.addWidget(self.status)
        advanced_layout.addWidget(self.paths)
        advanced_layout.addWidget(self.log_area)
        buttons = QHBoxLayout()
        for text, callback in (
            ("Procesar archivo", self.scan_now),
            ("Plantillas", lambda: open_folder(self.config.template_folder)),
            ("Salida", lambda: open_folder(self.config.output_folder)),
            ("Configuracion", self.open_config),
            ("Registro", self.open_log),
            ("Iniciar", self.start_monitoring),
            ("Detener", self.stop_monitoring),
        ):
            button = QPushButton(text)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        advanced_layout.addLayout(buttons)
        self.advanced_widget.setVisible(False)
        layout.addWidget(self.advanced_widget)
        self.setCentralWidget(central)
        self.resize(920, 560)

    def _workflow_panel(self, title: str, hint: str, callback) -> QWidget:
        panel = QWidget()
        panel.setObjectName("ActionPanel")
        layout = QVBoxLayout(panel)
        layout.setSpacing(8)
        title_label = QLabel(title)
        title_label.setObjectName("ActionTitle")
        hint_label = QLabel(hint)
        hint_label.setObjectName("ActionHint")
        hint_label.setWordWrap(True)
        date_edit = QDateEdit(QDate.currentDate())
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("dd/MM/yyyy")
        button = QPushButton(title)
        button.clicked.connect(callback)
        layout.addWidget(title_label)
        layout.addWidget(hint_label)
        layout.addWidget(date_edit)
        layout.addWidget(button)
        if title == "Entradas":
            self.entries_date = date_edit
        else:
            self.departures_date = date_edit
        return panel

    def _toggle_advanced_options(self) -> None:
        visible = self.advanced_toggle.isChecked()
        self.advanced_widget.setVisible(visible)
        self.advanced_toggle.setText("Ocultar opciones avanzadas" if visible else "Mostrar opciones avanzadas")

    def _build_tray(self) -> None:
        self.tray = QSystemTrayIcon(make_icon(), self)
        self.tray.setToolTip("El duendecito de Vianni - vigilando nuevas entradas")
        menu = QMenu()
        actions = [
            ("Abrir El duendecito de Vianni", self.show_window),
            ("Entradas", self.run_entradas),
            ("Procesar archivo de Descargas", self.scan_now),
            ("Iniciar monitoreo", self.start_monitoring),
            ("Detener monitoreo", self.stop_monitoring),
            ("Abrir carpeta de salida", lambda: open_folder(self.config.output_folder)),
            ("Configuracion", self.open_config),
            ("Ver registro", self.open_log),
            ("Salir", self.quit_app),
        ]
        for text, callback in actions:
            action = QAction(text, self)
            action.triggered.connect(callback)
            menu.addAction(action)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self._tray_activated)
        self.tray.show()

    def _tray_activated(self, reason) -> None:
        if reason in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick):
            self.show_window()

    def show_startup_greeting(self) -> None:
        if not self.tray.isVisible():
            return
        if self.monitoring:
            message = "Estoy despierto y vigilando. Siempre chequeando si hay nuevos empleados. Todo esta funcionando bien."
        else:
            message = "Estoy listo en la bandeja. Puede iniciar el monitoreo cuando desee."
        self.tray.showMessage("El duendecito de Vianni", message, make_icon(), 6000)

    def _sync_timer(self) -> None:
        if self.monitoring:
            self.timer.start(self.config.scan_interval_minutes * 60 * 1000)
        else:
            self.timer.stop()

    def refresh_status(self) -> None:
        active = "activo" if self.monitoring else "inactivo"
        charm = "Listo para ayudar" if self.monitoring else "Descansando en la bandeja"
        self.status.setText(
            f"Estado del duendecito: {charm}\nMonitoreo: {active}\nUltimo escaneo: {self.last_scan}\n"
            f"Ultimo archivo procesado: {self.last_spreadsheet or 'Ninguno'}\n"
            f"Empleados procesados en la ultima corrida: {self.last_employees}"
        )
        self.paths.setText(
            f"Descargas: {self.config.downloads_folder}\n"
            f"Plantillas: {self.config.template_folder}\n"
            f"Salida: {self.config.output_folder}"
        )

    def append_log(self, text: str) -> None:
        self.log_area.append(text)
        logging.info(text)

    def set_progress(self, value: int, message: str) -> None:
        self.busy_status_timer.stop()
        self.stop_busy_elf()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(max(0, min(100, value)))
        self.progress_label.setText(message)
        QApplication.processEvents()

    def start_busy_progress(self, message: str) -> None:
        self.start_busy_elf()
        self.progress_bar.setRange(0, 0)
        self.busy_started_at = time.monotonic()
        self.busy_status_index = 0
        self.progress_label.setText(message)
        self.busy_status_timer.start(1200)
        QApplication.processEvents()

    def _advance_busy_status(self) -> None:
        if not self.busy_status_timer.isActive():
            return
        self.busy_status_index = (self.busy_status_index + 1) % len(self.busy_status_phases)
        elapsed = int(time.monotonic() - self.busy_started_at)
        minutes, seconds = divmod(elapsed, 60)
        phase = self.busy_status_phases[self.busy_status_index]
        self.progress_label.setText(f"{phase}...  {minutes:02d}:{seconds:02d}")

    def start_busy_elf(self) -> None:
        self.busy_elf.setVisible(True)
        if not self.dance_timer.isActive():
            self.dance_timer.start(260)

    def stop_busy_elf(self) -> None:
        self.dance_timer.stop()
        self.busy_elf.setVisible(False)

    def _advance_dancing_elf(self) -> None:
        self.dance_frame = (self.dance_frame + 1) % 2
        self.busy_elf.setPixmap(make_dancing_elf_frame(self.dance_frame))

    def scan_now(self) -> None:
        from datetime import datetime

        self.last_scan = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.append_log("El duendecito esta revisando Descargas...")
        self.set_progress(10, "Revisando archivo en Descargas...")
        try:
            report = self.processor.process_next_export(delete_original=not self.config.ask_before_delete_original)
        except Exception as exc:
            logging.exception("Error al procesar")
            self.set_progress(0, "No se pudo procesar el archivo.")
            QMessageBox.critical(self, "El duendecito necesita ayuda", f"No se pudo procesar el archivo.\n\n{exc}")
            self.refresh_status()
            return
        self.set_progress(100, "Documentos generados.")
        self.handle_report(report)

    def run_entradas(self) -> None:
        self.run_mercury(self._selected_entries_date())

    def run_salidas(self) -> None:
        selected = self._selected_departures_date().strftime("%d/%m/%Y")
        QMessageBox.information(
            self,
            "Salidas",
            f"La funcion de salidas para {selected} sera el proximo flujo que vamos a construir.",
        )

    def _selected_entries_date(self) -> date:
        return self.entries_date.date().toPython()

    def _selected_departures_date(self) -> date:
        return self.departures_date.date().toPython()

    def handle_report(self, report: RunReport) -> None:
        self.last_spreadsheet = report.imported_spreadsheet or report.source_spreadsheet
        self.last_employees = report.employees_processed
        self.append_log(report.message)
        if report.already_processed:
            QMessageBox.information(self, "Archivo ya revisado", report.message)
        elif report.employees_processed:
            self.tray.showMessage("El duendecito de Vianni", "Listo: los documentos fueron generados correctamente.")
            response = QMessageBox.question(
                self,
                "Trabajo terminado",
                f"{report.message}\n\nDesea imprimir los documentos ahora?",
            )
            if response == QMessageBox.StandardButton.Yes:
                self.print_report(report)
            else:
                open_response = QMessageBox.question(self, "Carpeta de salida", "Desea abrir la carpeta de salida?")
                if open_response == QMessageBox.StandardButton.Yes:
                    open_folder(report.output_folder)
            if self.config.ask_before_delete_original and report.source_spreadsheet:
                delete_response = QMessageBox.question(
                    self,
                    "Limpiar Descargas",
                    "Desea borrar el archivo original de la carpeta Descargas?",
                )
                if delete_response == QMessageBox.StandardButton.Yes:
                    Path(report.source_spreadsheet).unlink(missing_ok=True)
        self.refresh_status()

    def print_report(self, report: RunReport) -> None:
        for document in report.generated_documents:
            print_file(document, self.config.selected_printer)
        self.append_log("El duendecito envio la impresion a Windows.")

    def open_config(self) -> None:
        dialog = ConfigDialog(self.config, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_mercury_password = dialog.mercury_password.text()
            self.config = dialog.updated_config()
            save_config(self.config, get_app_root())
            if new_mercury_password:
                try:
                    save_mercury_password(new_mercury_password)
                except Exception as exc:
                    QMessageBox.warning(
                        self,
                        "Contrasena Mercury",
                        f"No se pudo guardar la contrasena de Mercury.\n\n{exc}",
                    )
                    logging.exception("No se pudo guardar la contrasena de Mercury")
                else:
                    self.append_log("Contrasena de Mercury guardada en Windows.")
            set_start_with_windows(self.config.start_with_windows)
            ensure_directories(self.config)
            self.processor = DocumentProcessor(self.config)
            self.monitoring = self.config.monitoring_enabled
            self._sync_timer()
            self.refresh_status()

    def open_log(self) -> None:
        if os.name == "nt":
            os.startfile(self.log_path)  # type: ignore[attr-defined]

    def run_mercury(self, report_date: date | None = None) -> None:
        report_date = report_date or date.today()
        if self.mercury_busy:
            self.append_log("El duendecito ya esta trabajando con Mercury.")
            return
        self.append_log(f"El duendecito esta buscando entradas de {report_date:%d/%m/%Y} en Mercury...")
        self.start_busy_progress("Entrando a Mercury y preparando la descarga...")
        self.mercury_busy = True
        self.mercury_report_date = report_date
        self.mercury_thread = QThread(self)
        self.mercury_worker = MercuryWorker(self.config, load_mercury_password(), report_date)
        self.mercury_worker.moveToThread(self.mercury_thread)
        self.mercury_thread.started.connect(self.mercury_worker.run)
        self.mercury_worker.finished.connect(self._handle_mercury_result)
        self.mercury_worker.failed.connect(self._handle_mercury_error)
        self.mercury_worker.finished.connect(self.mercury_thread.quit)
        self.mercury_worker.failed.connect(self.mercury_thread.quit)
        self.mercury_worker.finished.connect(self.mercury_worker.deleteLater)
        self.mercury_worker.failed.connect(self.mercury_worker.deleteLater)
        self.mercury_thread.finished.connect(self.mercury_thread.deleteLater)
        self.mercury_thread.finished.connect(self._clear_mercury_worker)
        self.mercury_thread.start()

    def _clear_mercury_worker(self) -> None:
        self.mercury_worker = None
        self.mercury_thread = None

    def _handle_mercury_error(self, exc: Exception) -> None:
        self.mercury_busy = False
        if isinstance(exc, MercuryAutomationError):
            self.set_progress(0, "Mercury necesita configuracion.")
            QMessageBox.warning(self, "Mercury necesita configuracion", str(exc))
            self.append_log(str(exc))
        else:
            logging.exception("Error al usar Mercury")
            self.set_progress(0, "No se pudo completar el trabajo en Mercury.")
            QMessageBox.critical(self, "Mercury necesita ayuda", f"No se pudo completar el trabajo en Mercury.\n\n{exc}")
        self.refresh_status()

    def _handle_mercury_result(self, result) -> None:
        report_date = self.mercury_report_date
        try:
            self.set_progress(45, "Reporte descargado. Revisando datos...")
            self.append_log(result.message)
            processed_reports: list[RunReport] = []
            empty_files: list[str] = []
            total_files = max(1, len(result.downloaded_files))
            for index, downloaded_file in enumerate(result.downloaded_files, start=1):
                progress = 45 + int((index - 1) / total_files * 45)
                self.set_progress(progress, f"Procesando documentos {index} de {total_files}...")
                if has_employee_rows(downloaded_file):
                    processed_reports.append(
                        self.processor.process_export_file(
                            downloaded_file,
                            force=True,
                            delete_original=not self.config.ask_before_delete_original,
                            run_date=report_date,
                        )
                    )
                else:
                    empty_files.append(Path(downloaded_file).name)
                self.set_progress(45 + int(index / total_files * 45), f"Documentos procesados {index} de {total_files}.")
        except Exception as exc:
            self.mercury_busy = False
            logging.exception("Error al usar Mercury")
            self.set_progress(0, "No se pudo completar el trabajo en Mercury.")
            QMessageBox.critical(self, "Mercury necesita ayuda", f"No se pudo completar el trabajo en Mercury.\n\n{exc}")
            return
        self.mercury_busy = False
        for company in result.companies_without_download:
            self.append_log(f"Mercury no genero archivo para {company}.")
        for filename in empty_files:
            self.append_log(f"{filename} no tiene nuevas entradas.")
        if not processed_reports:
            self.set_progress(100, "No hay nuevas entradas para procesar.")
            QMessageBox.information(self, "Mercury", "No hay nuevas entradas para procesar.")
            self.refresh_status()
            return
        self.set_progress(100, "Documentos generados.")
        for report in processed_reports:
            self.handle_report(report)
        self.set_progress(100, "Trabajo terminado.")

    def start_monitoring(self) -> None:
        self.monitoring = True
        self.config.monitoring_enabled = True
        self._sync_timer()
        self.tray.showMessage("El duendecito de Vianni", "El duendecito esta atento en Descargas.")
        self.refresh_status()

    def stop_monitoring(self) -> None:
        self.monitoring = False
        self.config.monitoring_enabled = False
        self._sync_timer()
        self.refresh_status()

    def show_window(self) -> None:
        self.show()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event) -> None:
        if self.exiting:
            event.accept()
        else:
            event.ignore()
            self.hide()
            self.tray.showMessage("El duendecito de Vianni", "Sigo trabajando desde la bandeja del sistema.")

    def quit_app(self) -> None:
        response = QMessageBox.question(self, "Salir", "Desea dejar descansar al duendecito y cerrar la aplicacion?")
        if response == QMessageBox.StandardButton.Yes:
            self.exiting = True
            QApplication.quit()


def main() -> int:
    app_root = get_app_root()
    config = load_config(app_root)
    ensure_directories(config)
    log_path = configure_logging(config.logs_folder)
    app = QApplication(sys.argv)
    app.setApplicationName("ElDuendecitoDeVianni")
    app.setOrganizationName("Vianni")
    app.setWindowIcon(make_icon())
    app.setQuitOnLastWindowClosed(False)

    lock_folder = Path(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation))
    lock_folder.mkdir(parents=True, exist_ok=True)
    lock_file = QLockFile(str(lock_folder / "ElDuendecitoDeVianni.lock"))
    lock_file.setStaleLockTime(30_000)
    if not lock_file.tryLock(100):
        QMessageBox.information(
            None,
            "El duendecito ya esta despierto",
            "Ya hay un duendecito de Vianni trabajando en este equipo.",
        )
        return 0
    app.instance_lock = lock_file  # type: ignore[attr-defined]

    window = MainWindow(config, log_path)
    if config.start_minimized_to_tray:
        window.hide()
    else:
        window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
