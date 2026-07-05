from __future__ import annotations

import os


MERCURY_PASSWORD_TARGET = "ElDuendecitoDeVianni/MercuryPassword"


def save_mercury_password(password: str) -> None:
    if not password:
        return
    _require_windows()
    import win32cred  # type: ignore

    credential = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": MERCURY_PASSWORD_TARGET,
        "UserName": "Mercury",
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(credential, 0)
    saved = load_mercury_password()
    if saved != password:
        raise RuntimeError("Windows no confirmo el guardado de la contrasena de Mercury.")


def load_mercury_password() -> str:
    _require_windows()
    import win32cred  # type: ignore

    try:
        credential = win32cred.CredRead(MERCURY_PASSWORD_TARGET, win32cred.CRED_TYPE_GENERIC)
    except Exception:
        return ""
    blob = credential.get("CredentialBlob", b"")
    if isinstance(blob, bytes):
        return blob.decode("utf-16-le", errors="ignore").rstrip("\x00")
    return str(blob)


def has_mercury_password() -> bool:
    return bool(load_mercury_password())


def delete_mercury_password() -> None:
    _require_windows()
    import win32cred  # type: ignore

    try:
        win32cred.CredDelete(MERCURY_PASSWORD_TARGET, win32cred.CRED_TYPE_GENERIC)
    except Exception:
        pass


def _require_windows() -> None:
    if os.name != "nt":
        raise RuntimeError("El guardado seguro de contrasenas solo esta disponible en Windows.")
