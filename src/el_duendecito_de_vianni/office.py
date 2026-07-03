from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path


def office_available() -> bool:
    if os.name != "nt":
        return False
    try:
        import win32com.client  # type: ignore

        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except Exception:
        return False


def open_folder(path: str | Path) -> None:
    folder = Path(path)
    folder.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(folder)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(folder)])


def print_file(path: str | Path, printer: str = "") -> bool:
    if os.name != "nt":
        logging.warning("La impresion automatica solo esta disponible en Windows.")
        return False
    try:
        if printer:
            logging.info("Imprimiendo %s en %s", path, printer)
        os.startfile(Path(path), "print")  # type: ignore[attr-defined]
        return True
    except Exception as exc:
        logging.exception("No se pudo imprimir %s: %s", path, exc)
        return False


def set_start_with_windows(enabled: bool) -> None:
    if os.name != "nt":
        return
    import winreg

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "ElDuendecitoDeVianni"
    command = f'"{sys.executable}"'
    if not getattr(sys, "frozen", False):
        command = f'"{sys.executable}" -m el_duendecito_de_vianni.app'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, command)
        else:
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
