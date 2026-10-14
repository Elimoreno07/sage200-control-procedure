# Manual de uso

## Para usuarios de QA

No es necesario instalar Python, Git ni Visual Studio Code.

1. Copia `Sage200ControlProcedure.exe` en una ubicación accesible.
2. Abre el ejecutable con doble clic.
3. Comprueba la carpeta de trabajo. Por defecto se propone el acceso sincronizado
   `TechDoc - 2026` de OneDrive.
4. Indica la **versión origen** y la **versión destino**.
5. Pulsa **Generar nueva versión**.
6. Espera a que termine la ventana de progreso.
7. Revisa el resumen y usa **Abrir carpeta** para ver el fichero generado.
8. Comprueba **Publicado en SharePoint: Sí**. El botón **Abrir SharePoint** abre
   la biblioteca corporativa en el navegador.

### Duplicar la última versión

1. Pulsa **Duplicar última versión**.
2. La aplicación detectará el Control Procedure con la versión numérica más alta.
3. La versión origen se completará automáticamente.
4. Se propondrá la siguiente revisión como destino. Por ejemplo,
   `2026.87.000` propondrá `2026.87.001`.
5. Cambia la versión destino si es necesario y pulsa **Generar nueva versión**.

### Carpeta de trabajo

Pulsa **Seleccionar...** para elegir otra carpeta. La aplicación recuerda la última
carpeta seleccionada.

Para subir automáticamente a SharePoint debe seleccionarse el acceso directo que
OneDrive sincroniza con la biblioteca corporativa. En el equipo habitual es:

```text
C:\Users\<usuario>\OneDrive - Sage Software, Inc\TechDoc - 2026
```

### Qué realiza la aplicación

- Localiza el Excel de la versión origen.
- Crea el nuevo fichero sin sobrescribir destinos existentes.
- Reemplaza la versión anterior por la nueva en todas las hojas.
- Limpia la columna F desde la fila 2 en la hoja activa.
- Conserva estilos, formatos, fórmulas y la estructura del libro.
- Genera un fichero de log por cada ejecución.
- Publica el resultado en `CONTROL-PROCEDURE/2026` de SharePoint mediante el
  acceso sincronizado `TechDoc - 2026`.

Los logs se guardan en:

```text
%LOCALAPPDATA%\Sage200ControlProcedure\logs
```

## Errores habituales

- **No se encuentra la versión origen:** comprueba el nombre del fichero y la
  carpeta seleccionada.
- **El destino ya existe:** el programa nunca lo sobrescribe. Elimina el destino
  solo si estás segura de que no se necesita o indica otra versión.
- **No se puede abrir o guardar:** cierra el fichero en Excel y vuelve a intentarlo.
- **Varios ficheros con la misma versión:** deja un único Control Procedure para
  esa versión.
- **No aparece en SharePoint:** comprueba que OneDrive esté iniciado y que la
  carpeta `TechDoc - 2026` exista dentro de `OneDrive - Sage Software, Inc`.

## Para generar el ejecutable

Esta sección es solo para mantenimiento técnico.

1. Instala Python 3.11 o posterior.
2. Ejecuta `build.bat`.
3. El ejecutable se generará en:

```text
dist\Sage200ControlProcedure.exe
```

PyInstaller empaqueta Python, Tkinter, openpyxl y el código del proyecto en un
único archivo. Los usuarios de QA solo necesitan el `.exe` resultante.
