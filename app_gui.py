import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import traceback
import generador  # Importa tu módulo de procesamiento

# -------------------------------
# Utilidad para EXE (carpeta real)
# -------------------------------
def resource_path(relative_path):
    """ Obtiene el path real dentro de un .exe """
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# -------------------------------
# Clase principal GUI
# -------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Generador Automático de Protocolos v1.0")
        self.geometry("700x520")
        self.resizable(False, False)

        # Variables
        self.plantilla_path = tk.StringVar()
        self.excels_paths = []

        # Layout
        self.build_gui()

    # -------------------------------
    # Interfaz gráfica
    # -------------------------------
    def build_gui(self):
        # Marco superior con título
        frame_title = tk.Frame(self, pady=10)
        frame_title.pack(fill="x")
        ttk.Label(frame_title,
                  text="Generador Automático de Protocolos v1.0",
                  font=("Segoe UI", 16, "bold")).pack()

        # Marco de selección de archivos
        frame_files = tk.Frame(self, pady=10)
        frame_files.pack(fill="x")

        # Selector plantilla
        ttk.Label(frame_files, text="Plantilla (.docx):",
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.plantilla_path, width=60).grid(row=1, column=0, padx=5)
        ttk.Button(frame_files, text="Seleccionar...",
                   command=self.seleccionar_plantilla).grid(row=1, column=1, padx=5)

        # Selector Excels
        ttk.Label(frame_files, text="\nArchivos Excel (.xlsx):",
                  font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w")
        self.excel_listbox = tk.Listbox(frame_files, width=60, height=5)
        self.excel_listbox.grid(row=3, column=0, padx=5)
        ttk.Button(frame_files, text="Agregar Excel...",
                   command=self.agregar_excels).grid(row=3, column=1, padx=5)

        # Botón generar
        frame_button = tk.Frame(self, pady=15)
        frame_button.pack(fill="x")
        ttk.Button(frame_button,
                   text="GENERAR DOCUMENTOS",
                   command=self.thread_generar,
                   width=40).pack()

        # Área de logs
        ttk.Label(self, text="Registro / Log:", font=("Segoe UI", 10, "bold")).pack()
        frame_log = tk.Frame(self)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)

        self.text_log = tk.Text(frame_log, height=12, wrap="word")
        self.text_log.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame_log, command=self.text_log.yview)
        scrollbar.pack(side="right", fill="y")

        self.text_log.config(yscrollcommand=scrollbar.set)

    # -------------------------------
    # Métodos de interacción
    # -------------------------------
    def seleccionar_plantilla(self):
        path = filedialog.askopenfilename(
            title="Seleccionar plantilla DOCX",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if path:
            self.plantilla_path.set(path)
            self.log(f"Plantilla seleccionada:\n{path}\n")

    def agregar_excels(self):
        paths = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if paths:
            for p in paths:
                self.excels_paths.append(p)
                self.excel_listbox.insert(tk.END, p)
            self.log(f"{len(paths)} archivo(s) Excel añadido(s).\n")

    # -------------------------------
    # Ejecución en hilos
    # -------------------------------
    def thread_generar(self):
        t = threading.Thread(target=self.generar_documentos)
        t.daemon = True
        t.start()

    # -------------------------------
    # Generación de documentos
    # -------------------------------
    def generar_documentos(self):
        try:
            plantilla = self.plantilla_path.get().strip()
            excels = self.excels_paths

            if not plantilla:
                messagebox.showerror("Error", "Debe seleccionar una plantilla Word.")
                return

            if not excels:
                messagebox.showerror("Error", "Debe agregar al menos un archivo Excel.")
                return

            self.log("Iniciando generación...\n")

            output_folder = "output"
            os.makedirs(output_folder, exist_ok=True)

            # Llamada al generador real
            resultados = generador.procesar_archivos(plantilla, excels, output_folder, self.log)

            self.log("\n" + "="*50)
            self.log("\nProceso terminado.\n")
            self.log(f"Documentos generados en:\n{os.path.abspath(output_folder)}\n")
            self.log("="*50 + "\n")

            messagebox.showinfo("Finalizado", "Documentos generados correctamente.")

        except Exception as e:
            self.log(f"\nERROR FATAL:\n{str(e)}\n")
            self.log(traceback.format_exc())
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    # -------------------------------
    # Log
    # -------------------------------
    def log(self, text):
        self.text_log.insert(tk.END, text)
        self.text_log.see(tk.END)


# -------------------------------
# Inicio del programa
# -------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
