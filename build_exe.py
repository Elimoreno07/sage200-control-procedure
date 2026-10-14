"""Genera el ejecutable único de Windows mediante PyInstaller."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path


NOMBRE_APLICACION = "Sage200ControlProcedure"


def main() -> int:
    try:
        import PyInstaller.__main__
    except ImportError:
        print(
            "ERROR: PyInstaller no está instalado. Ejecuta primero "
            "'python -m pip install -r requirements.txt'.",
            file=sys.stderr,
        )
        return 1

    raiz = Path(__file__).resolve().parent
    entrada = raiz / "app.py"
    if not entrada.exists():
        print(f"ERROR: No se encuentra {entrada}", file=sys.stderr)
        return 1

    carpeta_temporal = Path(tempfile.gettempdir()) / f"{NOMBRE_APLICACION}-build"

    PyInstaller.__main__.run(
        [
            str(entrada),
            "--name",
            NOMBRE_APLICACION,
            "--onefile",
            "--windowed",
            "--clean",
            "--noconfirm",
            "--distpath",
            str(raiz / "dist"),
            "--workpath",
            str(carpeta_temporal),
            "--specpath",
            str(carpeta_temporal),
            "--collect-all",
            "openpyxl",
        ]
    )

    ejecutable = raiz / "dist" / f"{NOMBRE_APLICACION}.exe"
    if not ejecutable.exists():
        print("ERROR: PyInstaller terminó sin crear el ejecutable.", file=sys.stderr)
        return 1

    print("Ejecutable generado correctamente:")
    print(ejecutable)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
