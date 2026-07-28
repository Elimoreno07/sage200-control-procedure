"""Crea el Excel Control Procedure de una nueva versión de Sage 200."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook


EXTENSIONES_EXCEL = {".xlsx", ".xlsm"}
CARPETA_CONTROL_PROCEDURE = Path(
    r"C:\Users\elisabeth.moreno\OneDrive - Sage Software, Inc\CONTROL PROCEDURE"
)


class ErrorControlProcedure(Exception):
    """Error esperado durante la creación del Control Procedure."""


def pedir_dato(mensaje: str) -> str:
    """Solicita un valor obligatorio por consola."""
    valor = input(mensaje).strip()
    if not valor:
        raise ErrorControlProcedure("El valor indicado no puede estar vacío.")
    return valor


def buscar_fichero_origen(carpeta: Path, version_anterior: str) -> Path:
    """Localiza un único Excel cuyo nombre contenga la versión anterior."""
    candidatos = sorted(
        (
            fichero
            for fichero in carpeta.iterdir()
            if fichero.is_file()
            and fichero.suffix.lower() in EXTENSIONES_EXCEL
            and not fichero.name.startswith("~$")
            and version_anterior in fichero.name
        ),
        key=lambda fichero: fichero.name.lower(),
    )

    if not candidatos:
        raise ErrorControlProcedure(
            f'No se ha encontrado ningún Excel que contenga la versión '
            f'"{version_anterior}" en:\n{carpeta}'
        )

    if len(candidatos) > 1:
        listado = "\n".join(f"  - {fichero.name}" for fichero in candidatos)
        raise ErrorControlProcedure(
            "Se han encontrado varios Excel posibles. Deja un único fichero "
            f"que coincida o concreta la versión:\n{listado}"
        )

    return candidatos[0]


def obtener_ultima_fila_con_datos(hoja) -> int:
    """Devuelve la última fila con algún valor, ignorando celdas solo formateadas."""
    for numero_fila in range(hoja.max_row, 0, -1):
        if any(celda.value is not None for celda in hoja[numero_fila]):
            return numero_fila
    return 1


def limpiar_columna_f(fichero: Path) -> None:
    """Limpia los valores de F2:F<última fila> en la hoja activa."""
    conservar_macros = fichero.suffix.lower() == ".xlsm"

    try:
        libro = load_workbook(fichero, keep_vba=conservar_macros)
    except PermissionError as exc:
        raise ErrorControlProcedure(
            f'No se puede abrir "{fichero.name}". Comprueba que el Excel esté cerrado.'
        ) from exc
    except OSError as exc:
        raise ErrorControlProcedure(
            f'No se ha podido abrir el Excel "{fichero.name}": {exc}'
        ) from exc

    try:
        hoja = libro.active
        ultima_fila = obtener_ultima_fila_con_datos(hoja)

        for numero_fila in range(2, ultima_fila + 1):
            hoja.cell(row=numero_fila, column=6).value = None

        try:
            libro.save(fichero)
        except PermissionError as exc:
            raise ErrorControlProcedure(
                f'No se puede guardar "{fichero.name}". Cierra el fichero en Excel '
                "y vuelve a ejecutar el script."
            ) from exc
        except OSError as exc:
            raise ErrorControlProcedure(
                f'No se ha podido guardar el Excel "{fichero.name}": {exc}'
            ) from exc
    finally:
        libro.close()


def crear_control_procedure(
    version_anterior: str, nueva_version: str, carpeta: Path
) -> Path:
    """Copia el Control Procedure anterior, limpia F y devuelve el destino."""
    if version_anterior == nueva_version:
        raise ErrorControlProcedure(
            "La versión anterior y la nueva versión deben ser diferentes."
        )

    if not carpeta.exists():
        raise ErrorControlProcedure(f"La ruta indicada no existe:\n{carpeta}")
    if not carpeta.is_dir():
        raise ErrorControlProcedure(f"La ruta indicada no es una carpeta:\n{carpeta}")

    origen = buscar_fichero_origen(carpeta, version_anterior)
    print(f"Fichero origen encontrado: {origen}")

    nombre_destino = origen.name.replace(version_anterior, nueva_version)
    destino = origen.with_name(nombre_destino)

    if destino.exists():
        raise ErrorControlProcedure(
            f"El fichero destino ya existe y no se sobrescribirá:\n{destino}"
        )

    try:
        shutil.copy2(origen, destino)
    except PermissionError as exc:
        raise ErrorControlProcedure(
            "No se ha podido crear la copia. Comprueba que los Excel estén cerrados "
            "y que tengas permiso de escritura en la carpeta."
        ) from exc
    except OSError as exc:
        raise ErrorControlProcedure(f"No se ha podido copiar el fichero: {exc}") from exc

    print(f"Fichero nuevo creado: {destino.name}")

    try:
        limpiar_columna_f(destino)
    except ErrorControlProcedure:
        # La copia se conserva para no borrar automáticamente un fichero del usuario.
        raise

    print("Columna F limpiada correctamente.")
    return destino


def main() -> int:
    """Punto de entrada interactivo del script."""
    print("Creación de Control Procedure para Sage 200")
    print("-" * 44)

    try:
        version_anterior = pedir_dato("Versión anterior (ej. 2026.85.000): ")
        nueva_version = pedir_dato("Nueva versión (ej. 2026.86.000): ")
        carpeta = CARPETA_CONTROL_PROCEDURE
        print(f"Carpeta de trabajo: {carpeta}")

        destino = crear_control_procedure(
            version_anterior, nueva_version, carpeta
        )
    except (ErrorControlProcedure, EOFError, KeyboardInterrupt) as exc:
        if isinstance(exc, KeyboardInterrupt):
            print("\nOperación cancelada por el usuario.", file=sys.stderr)
        elif isinstance(exc, EOFError):
            print("\nERROR: No se han podido leer los datos de entrada.", file=sys.stderr)
        else:
            print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # Red de seguridad para errores inesperados de openpyxl.
        print(f"\nERROR inesperado: {exc}", file=sys.stderr)
        return 1

    print(f"Ruta final del fichero generado: {destino.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
