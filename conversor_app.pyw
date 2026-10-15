"""
Ventana del Conversor de Asistencia (para Recursos Humanos).

Usa la misma lógica de convertir_reporte.py: separa "Fecha/Hora" en
"Fecha" y "Hora" y guarda <nombre>_convertido.xlsx junto al original,
sin modificar el original.
"""

import ctypes
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP = "Conversor de Asistencia"
AZUL = "#305496"
VERDE = "#2E7D32"
NARANJA = "#B26A00"
ROJO = "#C62828"

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from convertir_reporte import ErrorReporte, procesar
except Exception as e:
    raiz = tk.Tk()
    raiz.withdraw()
    messagebox.showerror(
        APP,
        f"No se pudo iniciar el conversor:\n\n{e}\n\n"
        "Verifique que Python tenga instalado openpyxl (pip install openpyxl).",
    )
    sys.exit(1)


class App(tk.Tk):
    def __init__(self, archivos):
        super().__init__()
        self.title(APP)
        icono = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "icono.ico"
        if icono.exists():
            self.iconbitmap(default=str(icono))

        self.escala = max(1.0, float(self.tk.call("tk", "scaling")) * 72 / 96)
        ancho = min(self.px(1180), int(self.winfo_screenwidth() * 0.92))
        alto = min(self.px(680), int(self.winfo_screenheight() * 0.85))
        self.geometry(f"{ancho}x{alto}")
        self.minsize(self.px(820), self.px(520))

        self.resultados = {}
        self.mensajes_error = {}
        self.filas_por_origen = {}
        self.carpeta_inicial = Path.home() / "Downloads"
        self.abrir_al_terminar = tk.BooleanVar(value=False)

        self._estilos()
        self._construir()
        self.bind("<Control-o>", lambda _e: self.seleccionar())
        if archivos:
            self.after(200, lambda: self.convertir(archivos))

    def px(self, valor):
        return int(valor * self.escala)

    def _estilos(self):
        st = ttk.Style(self)
        st.configure(".", font=("Segoe UI", 10))
        st.configure("Grande.TButton", font=("Segoe UI", 12, "bold"), padding=(18, 10))
        st.configure("TButton", padding=(10, 6))
        st.configure("Treeview", rowheight=self.px(24))
        st.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        st.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))

    def _construir(self):
        cab = tk.Frame(self, bg=AZUL, padx=20, pady=14)
        cab.pack(fill="x")
        tk.Label(cab, text=APP, bg=AZUL, fg="white",
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(cab, bg=AZUL, fg="#DCE6F5", font=("Segoe UI", 10),
                 text="Separa la columna Fecha/Hora (o Tiempo) del reporte del reloj en Fecha y Hora. "
                      "El archivo original no se modifica.").pack(anchor="w")

        self.estado = tk.StringVar(value="Seleccione el reporte exportado del software (Ctrl+O).")
        ttk.Label(self, textvariable=self.estado, padding=(20, 6),
                  background="#EEF1F6").pack(fill="x", side="bottom")

        cuerpo = ttk.Frame(self, padding=(20, 14, 20, 6))
        cuerpo.pack(fill="both", expand=True)

        acciones = ttk.Frame(cuerpo)
        acciones.pack(fill="x", pady=(0, 12))
        ttk.Button(acciones, text="Seleccionar reporte(s)…", style="Grande.TButton",
                   command=self.seleccionar).pack(side="left")
        ttk.Checkbutton(acciones, text="Abrir en Excel al terminar",
                        variable=self.abrir_al_terminar).pack(side="left", padx=16)
        self.btn_carpeta = ttk.Button(acciones, text="Abrir carpeta",
                                      command=self.abrir_carpeta, state="disabled")
        self.btn_carpeta.pack(side="right")
        self.btn_excel = ttk.Button(acciones, text="Abrir en Excel",
                                    command=self.abrir_excel, state="disabled")
        self.btn_excel.pack(side="right", padx=(0, 8))

        # Lista de reportes convertidos en esta sesión
        marco = ttk.LabelFrame(cuerpo, text="Reportes procesados", padding=8)
        marco.pack(fill="x")
        cols = ("archivo", "periodo", "registros", "separados", "revisar", "estado")
        self.tabla = ttk.Treeview(marco, columns=cols, show="headings", height=4,
                                  selectmode="browse")
        for col, titulo, ancho, anchor in (
            ("archivo", "Archivo", 180, "w"),
            ("periodo", "Periodo", 250, "w"),
            ("registros", "Registros", 85, "center"),
            ("separados", "Separados", 90, "center"),
            ("revisar", "A revisar", 85, "center"),
            ("estado", "Estado", 330, "w"),
        ):
            self.tabla.heading(col, text=titulo)
            self.tabla.column(col, width=self.px(ancho), minwidth=self.px(60),
                              anchor=anchor, stretch=(col == "estado"))
        self.tabla.tag_configure("ok", foreground=VERDE)
        self.tabla.tag_configure("aviso", foreground=NARANJA)
        self.tabla.tag_configure("error", foreground=ROJO)
        self.tabla.pack(fill="x")
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self.mostrar_seleccion())
        self.tabla.bind("<Double-1>", lambda _e: self.abrir_excel())

        # Vista previa del reporte seleccionado
        self.marco_prev = ttk.LabelFrame(cuerpo, text="Vista previa", padding=8)
        self.marco_prev.pack(fill="both", expand=True, pady=(12, 0))
        cols = ("id", "nombre", "fecha", "hora", "original", "tipo", "terminal")
        self.prev = ttk.Treeview(self.marco_prev, columns=cols, show="headings")
        for col, titulo, ancho, anchor in (
            ("id", "ID de Usuario", 100, "center"),
            ("nombre", "Nombre Completo", 240, "w"),
            ("fecha", "Fecha", 110, "center"),
            ("hora", "Hora", 80, "center"),
            ("original", "Fecha/Hora original", 160, "center"),
            ("tipo", "Tipo de Registro", 130, "w"),
            ("terminal", "Terminal", 120, "w"),
        ):
            self.prev.heading(col, text=titulo)
            self.prev.column(col, width=self.px(ancho), anchor=anchor)
        self.prev.tag_configure("aviso", background="#FFF2CC")
        barra = ttk.Scrollbar(self.marco_prev, orient="vertical", command=self.prev.yview)
        self.prev.configure(yscrollcommand=barra.set)
        self.prev.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

    def seleccionar(self):
        archivos = filedialog.askopenfilenames(
            parent=self,
            title="Seleccione el/los reporte(s) de eventos",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(self.carpeta_inicial),
        )
        if archivos:
            self.convertir(archivos)

    def convertir(self, archivos):
        self.config(cursor="watch")
        self.estado.set("Convirtiendo…")
        self.update_idletasks()
        ultimo, errores, avisos = None, [], 0
        for a in archivos:
            ruta = Path(a).resolve()
            self.carpeta_inicial = ruta.parent
            try:
                r = procesar(ruta)
            except ErrorReporte as e:
                r, msg = None, str(e)
            except Exception as e:
                r, msg = None, f"Error inesperado: {e}"

            # Si el mismo archivo se convierte otra vez, reemplaza su fila
            anterior = self.filas_por_origen.pop(ruta, None)
            if anterior:
                self.tabla.delete(anterior)
                self.resultados.pop(anterior, None)
                self.mensajes_error.pop(anterior, None)

            if r is None:
                errores.append(f"{ruta.name}: {msg}")
                iid = self.tabla.insert("", 0, tags=("error",),
                                        values=(ruta.name, "", "", "", "", msg))
                self.mensajes_error[iid] = f"{ruta.name}: {msg}"
            else:
                n_rev = len(r["revisar"])
                avisos += bool(n_rev)
                estado = (f"Revisar {n_rev} fila(s) — ver hoja 'Revisar'" if n_rev
                          else f"Listo → {r['salida'].name}")
                iid = self.tabla.insert(
                    "", 0, tags=("aviso" if n_rev else "ok",),
                    values=(ruta.name, r["rango"], len(r["registros"]),
                            r["convertidos"], n_rev, estado))
                self.resultados[iid] = r
            self.filas_por_origen[ruta] = iid
            ultimo = iid

        self.config(cursor="")
        if ultimo:
            self.tabla.selection_set(ultimo)
            self.tabla.focus(ultimo)
        ok = len(archivos) - len(errores)
        self.estado.set(f"{ok} de {len(archivos)} reporte(s) convertidos"
                        + (f", {avisos} con filas a revisar." if avisos else "."))
        if errores:
            messagebox.showwarning(APP, "No se pudieron convertir:\n\n" + "\n\n".join(errores),
                                   parent=self)
            return
        if self.abrir_al_terminar.get() and ok == 1:
            self.abrir_excel()

    def _actual(self):
        sel = self.tabla.selection()
        return self.resultados.get(sel[0]) if sel else None

    def mostrar_seleccion(self):
        self.prev.delete(*self.prev.get_children())
        r = self._actual()
        estado = "normal" if r else "disabled"
        self.btn_excel.config(state=estado)
        self.btn_carpeta.config(state=estado)
        if not r:
            self.marco_prev.config(text="Vista previa")
            sel = self.tabla.selection()
            if sel and sel[0] in self.mensajes_error:
                self.estado.set(self.mensajes_error[sel[0]])
            return
        self.marco_prev.config(
            text=f"Vista previa — {r['salida'].name} ({len(r['registros'])} registros)")
        for reg in r["registros"]:
            partes = reg["partes"]
            if partes:
                fecha = partes[0].strftime("%d/%m/%Y")
                hora = partes[1].strftime("%H:%M:%S" if partes[2] else "%H:%M")
            else:
                fecha = hora = "—"
            self.prev.insert("", "end", tags=() if partes else ("aviso",), values=(
                reg["ID de Usuario"] or "", reg["Nombre Completo"] or "", fecha, hora,
                reg["Fecha/Hora"] or "", reg["Tipo de Registro"] or "", reg["Terminal"] or ""))

    def abrir_excel(self):
        r = self._actual()
        if r and r["salida"].exists():
            os.startfile(r["salida"])

    def abrir_carpeta(self):
        r = self._actual()
        if r:
            subprocess.Popen(f'explorer /select,"{r["salida"]}"')


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App(sys.argv[1:]).mainloop()


if __name__ == "__main__":
    main()
