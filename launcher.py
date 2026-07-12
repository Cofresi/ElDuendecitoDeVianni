from pathlib import Path
import sys


# PyInstaller supplies the package path when bundled; add it for source runs.
if getattr(sys, "frozen", False) is False:
    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))


from el_duendecito_de_vianni.app import main


if __name__ == "__main__":
    raise SystemExit(main())
