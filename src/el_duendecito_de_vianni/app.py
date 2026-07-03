from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

from PySide6.QtCore import QPoint, QTimer, Qt
from PySide6.QtGui import QAction, QBrush, QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSystemTrayIcon,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .config import AppConfig, ensure_directories, get_app_root, load_config, save_config
from .logging_utils import configure_logging
from .office import open_folder, print_file, set_start_with_windows
from .processor import DocumentProcessor, RunReport


APP_STYLESHEET = """
QMainWindow, QDialog {
    background: #f7f1e5;
    color: #283324;
    font-family: "Segoe UI";
    font-size: 10pt;
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
    padding: 12px;
}
QTextEdit {
    background: #fffdf7;
    border: 1px solid #cdbb8a;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #5b8c48;
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
QLineEdit, QSpinBox, QComboBox {
    background: #fffdf7;
    border: 1px solid #cdbb8a;
    border-radius: 6px;
    padding: 6px;
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


class ConfigDialog(QDialog):
    def __init__(self, config: AppConfig, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuracion")
        self.setWindowIcon(make_icon())
        self.setStyleSheet(APP_STYLESHEET)
        self.config = config
        layout = QFormLayout(self)
        self.downloads = self._path_row(config.downloads_folder)
        self.templates = self._path_row(config.template_folder)
        self.output = self._path_row(config.output_folder)
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
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        edit = QLineEdit(value)
        button = QPushButton("...")
        button.clicked.connect(lambda: self._choose_folder(edit))
        layout.addWidget(edit)
        layout.addWidget(button)
        widget.edit = edit  # type: ignore[attr-defined]
        return widget

    def _choose_folder(self, edit: QLineEdit) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta", edit.text())
        if folder:
            edit.setText(folder)

    def updated_config(self) -> AppConfig:
        return AppConfig(
            downloads_folder=self.downloads.edit.text(),  # type: ignore[attr-defined]
            template_folder=self.templates.edit.text(),  # type: ignore[attr-defined]
            output_folder=self.output.edit.text(),  # type: ignore[attr-defined]
            imported_folder=self.config.imported_folder,
            logs_folder=self.config.logs_folder,
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
        self.monitoring = config.monitoring_enabled
        self.exiting = False
        self.setWindowTitle("El duendecito de Vianni")
        self.setWindowIcon(make_icon())
        self._build_ui()
        self._build_tray()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.scan_now)
        self._sync_timer()
        self.refresh_status()
        QTimer.singleShot(1200, self.show_startup_greeting)

    def _build_ui(self) -> None:
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        title = QLabel("El duendecito de Vianni")
        title.setObjectName("TitleLabel")
        subtitle = QLabel("Un ayudante discreto para preparar documentos de nuevas entradas")
        subtitle.setObjectName("SubtitleLabel")
        layout.addWidget(title)
        layout.addWidget(subtitle)
        self.status = QLabel()
        self.status.setObjectName("StatusCard")
        self.paths = QLabel()
        self.paths.setObjectName("PathCard")
        self.paths.setWordWrap(True)
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.status)
        layout.addWidget(self.paths)
        layout.addWidget(self.log_area)
        buttons = QHBoxLayout()
        for text, callback in (
            ("Procesar ahora", self.scan_now),
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
        layout.addLayout(buttons)
        self.setCentralWidget(central)
        self.resize(920, 560)

    def _build_tray(self) -> None:
        self.tray = QSystemTrayIcon(make_icon(), self)
        self.tray.setToolTip("El duendecito de Vianni - vigilando nuevas entradas")
        menu = QMenu()
        actions = [
            ("Abrir El duendecito de Vianni", self.show_window),
            ("Procesar ahora", self.scan_now),
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
            message = "El duendecito esta despierto y vigilando Descargas. Todo esta funcionando bien."
        else:
            message = "El duendecito esta listo en la bandeja. Puede iniciar el monitoreo cuando desee."
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

    def scan_now(self) -> None:
        from datetime import datetime

        self.last_scan = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.append_log("El duendecito esta revisando Descargas...")
        try:
            report = self.processor.process_next_export(delete_original=not self.config.ask_before_delete_original)
        except Exception as exc:
            logging.exception("Error al procesar")
            QMessageBox.critical(self, "El duendecito necesita ayuda", f"No se pudo procesar el archivo.\n\n{exc}")
            self.refresh_status()
            return
        self.handle_report(report)

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
            self.config = dialog.updated_config()
            save_config(self.config, get_app_root())
            set_start_with_windows(self.config.start_with_windows)
            ensure_directories(self.config)
            self.processor = DocumentProcessor(self.config)
            self.monitoring = self.config.monitoring_enabled
            self._sync_timer()
            self.refresh_status()

    def open_log(self) -> None:
        if os.name == "nt":
            os.startfile(self.log_path)  # type: ignore[attr-defined]

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
    app.setWindowIcon(make_icon())
    app.setQuitOnLastWindowClosed(False)
    window = MainWindow(config, log_path)
    if config.start_minimized_to_tray:
        window.hide()
    else:
        window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
