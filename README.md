# Sage 200 Control Procedure

Aplicación de escritorio para crear una nueva versión del Excel **Control
Procedure** de Sage 200 sin usar Python, Git ni Visual Studio Code.

## Funcionalidad

- Interfaz gráfica para Windows creada con Tkinter.
- Selección de la carpeta de trabajo mediante el Explorador.
- Detección automática de la última versión numérica disponible.
- Propuesta de la siguiente revisión.
- Creación del destino sin sobrescribir ficheros existentes.
- Sustitución de la versión origen por la versión destino en todas las hojas.
- Limpieza de la columna F desde la fila 2 en la hoja activa.
- Conservación de estilos, formatos, fórmulas y estructura del libro.
- Ventana de progreso y resumen final.
- Log independiente por ejecución.
- Publicación automática en la biblioteca corporativa de SharePoint mediante el
  acceso sincronizado `TechDoc - 2026`.
- Compatibilidad con `.xlsx` y `.xlsm`.

## Uso de la aplicación

Ejecuta:

```text
Sage200ControlProcedure.exe
```

Después:

1. Selecciona la carpeta si la propuesta no es correcta.
2. Indica la versión origen y la versión destino.
3. Pulsa **Generar nueva versión**.

El botón **Duplicar última versión** detecta el Excel más reciente, completa la
versión origen y propone la revisión siguiente. Solo es necesario revisar o
indicar la versión destino y pulsar **Generar nueva versión**.

Consulta [MANUAL_USO.md](MANUAL_USO.md) para las instrucciones completas.

## Carpeta predeterminada

La aplicación propone este acceso de OneDrive:

```text
C:\Users\<usuario>\OneDrive - Sage Software, Inc\TechDoc - 2026
```

Este acceso debe estar sincronizado con `CONTROL-PROCEDURE/2026` en SharePoint.
La ruta se construye con el usuario de Windows actual y se puede cambiar desde la
interfaz.

Aunque se elija otra carpeta de trabajo, la aplicación copia también el resultado
al acceso corporativo `TechDoc - 2026`. El resumen final confirma la publicación
y ofrece un botón **Abrir SharePoint**.

## Desarrollo

Requisitos:

- Windows
- Python 3.11 o posterior

Instalación:

```powershell
py -m pip install -r requirements.txt
```

Ejecutar la interfaz:

```powershell
py app.py
```

Ejecutar la alternativa por consola:

```powershell
py crear_control_procedure.py
```

Ejecutar las pruebas:

```powershell
py -m unittest discover -s tests -v
```

## Generar el `.exe`

Haz doble clic en `build.bat` o ejecuta:

```powershell
build.bat
```

El resultado será un único archivo:

```text
dist\Sage200ControlProcedure.exe
```

`build_exe.py` contiene la configuración reproducible de PyInstaller. El
ejecutable se genera en modo ventana, sin consola, e incluye Python, Tkinter y
openpyxl.

## Estructura

```text
app.py                         Interfaz gráfica
crear_control_procedure.py     Lógica de negocio y entrada por consola
build.bat                      Compilación sencilla para Windows
build_exe.py                   Configuración de PyInstaller
requirements.txt               Dependencias
MANUAL_USO.md                  Manual para QA y mantenimiento
tests/                         Pruebas automáticas
```

## Logs

Cada ejecución crea un log en:

```text
%LOCALAPPDATA%\Sage200ControlProcedure\logs
```
