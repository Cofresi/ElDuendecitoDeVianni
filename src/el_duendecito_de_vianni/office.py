from __future__ import annotations

import logging
import os
import subprocess
import sys
from collections.abc import Iterable
from pathlib import Path
from zipfile import BadZipFile, ZipFile


WORD_FIELD_MARKERS = (b"<w:instrText", b"<w:fldSimple")


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


def docx_has_word_fields(path: str | Path) -> bool:
    document_path = Path(path)
    try:
        with ZipFile(document_path) as archive:
            for name in archive.namelist():
                if not name.startswith("word/") or not name.endswith(".xml"):
                    continue
                content = archive.read(name)
                if any(marker in content for marker in WORD_FIELD_MARKERS):
                    return True
    except (OSError, BadZipFile) as exc:
        raise RuntimeError(f"No se pudo inspeccionar el documento de Word: {document_path}") from exc
    return False


def update_word_fields(paths: Iterable[str | Path]) -> None:
    documents = [
        Path(path).resolve()
        for path in paths
        if Path(path).suffix.casefold() == ".docx" and docx_has_word_fields(path)
    ]
    if not documents:
        return

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Microsoft Word y pywin32 son necesarios para actualizar las fechas dinamicas.") from exc

    word = None
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for document_path in documents:
            document = None
            try:
                document = word.Documents.Open(
                    str(document_path),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                )
                document.Fields.Update()
                for story_range in document.StoryRanges:
                    current_range = story_range
                    while current_range is not None:
                        current_range.Fields.Update()
                        current_range = current_range.NextStoryRange
                document.Save()
            except Exception as exc:
                raise RuntimeError(
                    f"No se pudieron actualizar las fechas dinamicas en {document_path.name}."
                ) from exc
            finally:
                if document is not None:
                    document.Close(SaveChanges=False)
        logging.info("Campos dinamicos de Word actualizados: %s documentos.", len(documents))
    finally:
        try:
            if word is not None:
                word.Quit()
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()


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
