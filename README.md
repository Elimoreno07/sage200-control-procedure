# Generador de Control Procedure para Sage 200

Script para crear el fichero Excel **Control Procedure** de una nueva versión a
partir del fichero de la versión anterior.

## Qué hace

1. Solicita la versión anterior, la nueva versión y la carpeta local de los Excel.
2. Busca un `.xlsx` o `.xlsm` que contenga la versión anterior en el nombre.
3. Crea una copia cuyo nombre contiene la nueva versión.
4. Limpia el contenido de la columna F, desde la fila 2 hasta la última fila con
   datos, en la hoja activa.
5. Mantiene la cabecera, los formatos y el resto del libro.

El script no sobrescribe un fichero de destino existente.

## Requisitos

- Windows
- Python 3.11 o posterior

## Instalación

Abre PowerShell dentro de esta carpeta y ejecuta:

```powershell
python -m pip install -r requirements.txt
```

## Uso

```powershell
python crear_control_procedure.py
```

Ejemplo de datos solicitados:

```text
Versión anterior (ej. 2026.85.000): 2026.85.000
Nueva versión (ej. 2026.86.000): 2026.86.000
Ruta local de la carpeta de los Excel: C:\...\CONTROL-PROCEDURE\2026
```

Si la ruta contiene espacios se puede pegar con o sin comillas.

## Recomendación

Cierra en Excel el fichero de destino antes de ejecutar el script. Si Windows
impide abrirlo o guardarlo, el programa mostrará un mensaje indicando que debe
cerrarse.
