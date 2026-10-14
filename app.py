"""Interfaz gráfica de Control Procedure para usuarios de QA."""

from __future__ import annotations

import json
import os
import queue
import subprocess
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from crear_control_procedure import (
    CARPETA_CONTROL_PROCEDURE,
    ErrorControlProcedure,
    ResultadoGeneracion,
    URL_SHAREPOINT,
    detectar_ultima_version,
    generar_control_procedure,
    proponer_siguiente_version,
)


TITULO = "Sage 200 · Control Procedure"
COLOR_PRINCIPAL = "#007F64"
COLOR_PRINCIPAL_ACTIVO = "#006A54"
COLOR_FONDO = "#F4F6F8"
COLOR_TEXTO = "#1F2933"


def _carpeta_configuracion() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    carpeta = base / "Sage200ControlProcedure"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


class AplicacionControlProcedure(tk.Tk):
    """Ventana principal de la aplicación."""

    def __init__(self) -> None:
        super().__init__()
        self.title(TITULO)
        self.geometry("760x500")
        self.minsize(680, 460)
        self.configure(bg=COLOR_FONDO)

        self.cola: queue.Queue[tuple[str, object]] = queue.Queue()
        self.progreso: tk.Toplevel | None = None
        self.barra_progreso: ttk.Progressbar | None = None
        self.texto_progreso = tk.StringVar(value="Preparando...")
        self.config_path = _carpeta_configuracion() / "config.json"

        self.carpeta_var = tk.StringVar(value=str(self._cargar_carpeta()))
        self.origen_var = tk.StringVar()
        self.destino_var = tk.StringVar()
        self.detectado_var = tk.StringVar(value="Última versión: pendiente de detectar")
        self.estado_var = tk.StringVar(value="Selecciona una carpeta o usa la propuesta automática.")

        self._configurar_estilos()
        self._crear_interfaz()
        self.after(150, lambda: self._detectar_ultima(silencioso=True))

    def _configurar_estilos(self) -> None:
        estilo = ttk.Style(self)
        estilo.theme_use("vista" if "vista" in estilo.theme_names() else "clam")
        estilo.configure("App.TFrame", background=COLOR_FONDO)
        estilo.configure(
            "Titulo.TLabel",
            background=COLOR_FONDO,
            foreground=COLOR_TEXTO,
            font=("Segoe UI", 20, "bold"),
        )
        estilo.configure(
            "Subtitulo.TLabel",
            background=COLOR_FONDO,
            foreground="#52606D",
            font=("Segoe UI", 10),
        )
        estilo.configure(
            "Campo.TLabel",
            background=COLOR_FONDO,
            foreground=COLOR_TEXTO,
            font=("Segoe UI", 10, "bold"),
        )
        estilo.configure(
            "Detectado.TLabel",
            background="#E8F5F1",
            foreground="#095C49",
            padding=10,
            font=("Segoe UI", 10, "bold"),
        )
        estilo.configure("TEntry", padding=7, font=("Segoe UI", 10))
        estilo.configure("TButton", padding=(12, 8), font=("Segoe UI", 10))
        estilo.configure(
            "Principal.TButton",
            padding=(18, 11),
            font=("Segoe UI", 11, "bold"),
            foreground="white",
            background=COLOR_PRINCIPAL,
        )
        estilo.map(
            "Principal.TButton",
            background=[("active", COLOR_PRINCIPAL_ACTIVO), ("disabled", "#9FB8B1")],
        )

    def _crear_interfaz(self) -> None:
        contenedor = ttk.Frame(self, style="App.TFrame", padding=28)
        contenedor.pack(fill="both", expand=True)
        contenedor.columnconfigure(1, weight=1)

        ttk.Label(
            contenedor, text="Generador de Control Procedure", style="Titulo.TLabel"
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            contenedor,
            text="Crea una nueva versión de Sage 200 conservando la estructura del Excel.",
            style="Subtitulo.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 24))

        ttk.Label(contenedor, text="Carpeta de trabajo", style="Campo.TLabel").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 6)
        )
        ttk.Entry(contenedor, textvariable=self.carpeta_var).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(contenedor, text="Seleccionar...", command=self._seleccionar_carpeta).grid(
            row=3, column=2, sticky="ew"
        )

        ttk.Label(contenedor, textvariable=self.detectado_var, style="Detectado.TLabel").grid(
            row=4, column=0, columnspan=3, sticky="ew", pady=(18, 20)
        )

        ttk.Label(contenedor, text="Versión origen", style="Campo.TLabel").grid(
            row=5, column=0, sticky="w", padx=(0, 12)
        )
        ttk.Label(contenedor, text="Versión destino", style="Campo.TLabel").grid(
            row=5, column=1, sticky="w", padx=(0, 12)
        )
        ttk.Entry(contenedor, textvariable=self.origen_var).grid(
            row=6, column=0, sticky="ew", padx=(0, 12), pady=(6, 0)
        )
        self.entrada_destino = ttk.Entry(contenedor, textvariable=self.destino_var)
        self.entrada_destino.grid(
            row=6, column=1, sticky="ew", padx=(0, 12), pady=(6, 0)
        )
        ttk.Button(
            contenedor,
            text="Duplicar última versión",
            command=lambda: self._detectar_ultima(silencioso=False),
        ).grid(row=6, column=2, sticky="ew", pady=(6, 0))

        separador = ttk.Separator(contenedor)
        separador.grid(row=7, column=0, columnspan=3, sticky="ew", pady=24)

        self.boton_generar = ttk.Button(
            contenedor,
            text="Generar nueva versión",
            style="Principal.TButton",
            command=self._iniciar_generacion,
        )
        self.boton_generar.grid(row=8, column=0, columnspan=3, sticky="ew")

        ttk.Label(
            contenedor,
            textvariable=self.estado_var,
            style="Subtitulo.TLabel",
            wraplength=690,
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(18, 0))

    def _cargar_carpeta(self) -> Path:
        try:
            datos = json.loads(self.config_path.read_text(encoding="utf-8"))
            carpeta = Path(datos.get("carpeta", ""))
            if carpeta.is_dir():
                return carpeta
        except (OSError, ValueError, TypeError):
            pass
        return CARPETA_CONTROL_PROCEDURE

    def _guardar_carpeta(self, carpeta: Path) -> None:
        try:
            self.config_path.write_text(
                json.dumps({"carpeta": str(carpeta)}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass

    def _seleccionar_carpeta(self) -> None:
        carpeta = filedialog.askdirectory(
            title="Selecciona la carpeta de Control Procedure",
            initialdir=self.carpeta_var.get() or str(Path.home()),
            mustexist=True,
        )
        if not carpeta:
            return
        ruta = Path(carpeta)
        self.carpeta_var.set(str(ruta))
        self._guardar_carpeta(ruta)
        self._detectar_ultima(silencioso=False)

    def _detectar_ultima(self, silencioso: bool) -> None:
        try:
            carpeta = Path(self.carpeta_var.get().strip())
            version, fichero = detectar_ultima_version(carpeta)
            propuesta = proponer_siguiente_version(version)
        except ErrorControlProcedure as exc:
            self.detectado_var.set("Última versión: no disponible")
            self.estado_var.set(str(exc).replace("\n", " "))
            if not silencioso:
                messagebox.showerror("No se pudo detectar la última versión", str(exc))
            return

        self.origen_var.set(version)
        self.destino_var.set(propuesta)
        self.detectado_var.set(f"Detectado: {fichero.name}")
        self.estado_var.set(
            f"Origen detectado automáticamente. Propuesta de destino: {propuesta}"
        )
        self.entrada_destino.focus_set()
        self.entrada_destino.selection_range(0, tk.END)

    def _iniciar_generacion(self) -> None:
        carpeta = Path(self.carpeta_var.get().strip())
        origen = self.origen_var.get().strip()
        destino = self.destino_var.get().strip()
        if not origen or not destino:
            messagebox.showwarning(
                "Faltan datos", "Indica la versión origen y la versión destino."
            )
            return

        self._guardar_carpeta(carpeta)
        self._mostrar_progreso()
        self.boton_generar.configure(state="disabled")

        hilo = threading.Thread(
            target=self._generar_en_segundo_plano,
            args=(origen, destino, carpeta),
            daemon=True,
        )
        hilo.start()
        self.after(100, self._procesar_cola)

    def _generar_en_segundo_plano(
        self, origen: str, destino: str, carpeta: Path
    ) -> None:
        def progreso(mensaje: str, porcentaje: int) -> None:
            self.cola.put(("progreso", (mensaje, porcentaje)))

        try:
            resultado = generar_control_procedure(
                origen, destino, carpeta, progress_callback=progreso
            )
            self.cola.put(("resultado", resultado))
        except ErrorControlProcedure as exc:
            self.cola.put(("error", str(exc)))
        except Exception as exc:
            self.cola.put(("error", f"Error inesperado: {exc}"))

    def _procesar_cola(self) -> None:
        terminado = False
        try:
            while True:
                tipo, contenido = self.cola.get_nowait()
                if tipo == "progreso":
                    mensaje, porcentaje = contenido  # type: ignore[misc]
                    self.texto_progreso.set(str(mensaje))
                    if self.barra_progreso:
                        self.barra_progreso["value"] = int(porcentaje)
                elif tipo == "resultado":
                    terminado = True
                    self._cerrar_progreso()
                    self._mostrar_resumen(contenido)  # type: ignore[arg-type]
                elif tipo == "error":
                    terminado = True
                    self._cerrar_progreso()
                    self.estado_var.set("No se ha generado el fichero.")
                    messagebox.showerror("No se pudo generar", str(contenido))
        except queue.Empty:
            pass

        if not terminado and self.progreso:
            self.after(100, self._procesar_cola)

    def _mostrar_progreso(self) -> None:
        ventana = tk.Toplevel(self)
        ventana.title("Generando Control Procedure")
        ventana.geometry("440x160")
        ventana.resizable(False, False)
        ventana.transient(self)
        ventana.grab_set()
        ventana.protocol("WM_DELETE_WINDOW", lambda: None)

        marco = ttk.Frame(ventana, padding=24)
        marco.pack(fill="both", expand=True)
        ttk.Label(marco, text="Procesando el Excel...", style="Campo.TLabel").pack(
            anchor="w"
        )
        ttk.Label(marco, textvariable=self.texto_progreso).pack(anchor="w", pady=(10, 12))
        self.barra_progreso = ttk.Progressbar(
            marco, mode="determinate", maximum=100, value=0
        )
        self.barra_progreso.pack(fill="x")
        self.progreso = ventana

    def _cerrar_progreso(self) -> None:
        self.boton_generar.configure(state="normal")
        if self.progreso:
            self.progreso.grab_release()
            self.progreso.destroy()
            self.progreso = None
        self.barra_progreso = None

    def _mostrar_resumen(self, resultado: ResultadoGeneracion) -> None:
        self.estado_var.set(f"Generado correctamente: {resultado.destino.name}")
        resumen = tk.Toplevel(self)
        resumen.title("Control Procedure generado")
        resumen.geometry("700x410")
        resumen.minsize(620, 380)
        resumen.transient(self)
        resumen.grab_set()

        marco = ttk.Frame(resumen, padding=24)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(1, weight=1)
        ttk.Label(marco, text="Proceso completado", style="Titulo.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 18)
        )

        datos = [
            ("Archivo origen", str(resultado.origen)),
            ("Archivo destino", str(resultado.destino)),
            ("Nº de hojas modificadas", str(resultado.hojas_modificadas)),
            ("Nº de reemplazos realizados", str(resultado.reemplazos_realizados)),
            ("Celdas limpiadas en columna F", str(resultado.celdas_columna_f_limpiadas)),
            (
                "Publicado en SharePoint",
                "Sí" if resultado.publicado_en_sharepoint else "No",
            ),
            (
                "Destino SharePoint",
                str(resultado.destino_sharepoint or "No publicado"),
            ),
            ("Errores encontrados", str(len(resultado.errores))),
            ("Fichero de log", str(resultado.log)),
        ]
        for fila, (etiqueta, valor) in enumerate(datos, start=1):
            ttk.Label(marco, text=etiqueta, style="Campo.TLabel").grid(
                row=fila, column=0, sticky="nw", padx=(0, 14), pady=5
            )
            ttk.Label(marco, text=valor, wraplength=450).grid(
                row=fila, column=1, sticky="nw", pady=5
            )

        botones = ttk.Frame(marco)
        botones.grid(row=len(datos) + 1, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        ttk.Button(
            botones,
            text="Abrir carpeta",
            command=lambda: self._abrir_ruta(resultado.destino.parent),
        ).pack(side="left")
        ttk.Button(
            botones,
            text="Abrir log",
            command=lambda: self._abrir_ruta(resultado.log),
        ).pack(side="left", padx=8)
        if resultado.publicado_en_sharepoint:
            ttk.Button(
                botones,
                text="Abrir SharePoint",
                command=lambda: webbrowser.open(URL_SHAREPOINT),
            ).pack(side="left")
        ttk.Button(botones, text="Cerrar", command=resumen.destroy).pack(side="right")

    @staticmethod
    def _abrir_ruta(ruta: Path) -> None:
        try:
            os.startfile(ruta)  # type: ignore[attr-defined]
        except (OSError, AttributeError):
            subprocess.Popen(["explorer", str(ruta)])


def main() -> None:
    aplicacion = AplicacionControlProcedure()
    aplicacion.mainloop()


if __name__ == "__main__":
    main()
