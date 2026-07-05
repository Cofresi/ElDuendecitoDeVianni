from __future__ import annotations

import os
from pathlib import Path


MERCURY_PASSWORD_TARGET = "ElDuendecitoDeVianni/MercuryPassword"
MERCURY_PASSWORD_FILE = "mercury_password.dpapi"


def save_mercury_password(password: str) -> None:
    if not password:
        return
    _require_windows()
    import win32crypt  # type: ignore

    protected = win32crypt.CryptProtectData(
        password.encode("utf-8"),
        "ElDuendecitoDeVianni Mercury",
        None,
        None,
        None,
        0,
    )
    path = _password_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(protected)
    saved = load_mercury_password()
    if saved != password:
        raise RuntimeError("Windows no confirmo el guardado de la contrasena de Mercury.")


def load_mercury_password() -> str:
    _require_windows()
    import win32crypt  # type: ignore

    path = _password_path()
    if path.exists():
        try:
            _, data = win32crypt.CryptUnprotectData(path.read_bytes(), None, None, None, 0)
            return data.decode("utf-8")
        except Exception:
            return ""

    # Backward compatibility for early builds that used Windows Credential Manager.
    try:
        import win32cred  # type: ignore

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
    _password_path().unlink(missing_ok=True)

    try:
        import win32cred  # type: ignore

        win32cred.CredDelete(MERCURY_PASSWORD_TARGET, win32cred.CRED_TYPE_GENERIC)
    except Exception:
        pass


def _require_windows() -> None:
    if os.name != "nt":
        raise RuntimeError("El guardado seguro de contrasenas solo esta disponible en Windows.")


def _password_path() -> Path:
    base = os.getenv("EL_DUENDECITO_CREDENTIAL_DIR")
    if base:
        return Path(base) / MERCURY_PASSWORD_FILE
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "Vianni" / "ElDuendecitoDeVianni" / MERCURY_PASSWORD_FILE
    return Path.home() / "AppData" / "Local" / "Vianni" / "ElDuendecitoDeVianni" / MERCURY_PASSWORD_FILE
