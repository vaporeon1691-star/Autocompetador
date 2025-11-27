import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pathlib

from generador import procesar_archivos   # ⬅ IMPORTA TU MOTOR REAL DE PROCESO


# ============================================================
#                 INTERFAZ GRAFICA (TKINTER)
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Generador Automático de Protocolos v1.0")
        root.geometry("900x600")

        self.docx_path = tk.StringVar()
        self.xlsx_files = []

        # === Crear carpeta de salida estándar ===
        documentos = pathlib.Path.home() / "Documents" / "Protocolos_Generados"
        os.makedirs(documentos, exist_ok=True)
        self.output_dir = str(documentos)

        # ============================================================
        #                            UI
        # ============================================================

        tk.Label(root, text="Generador Automático de Protocolos",
                 font=("Arial", 18, "bold")).pack(pady=10)

        frame = tk.Frame(root)
        frame.pack()

        # -----------------------------
        #   SELECCIONAR PLANTILLA DOCX
        # -----------------------------
        tk.Label(frame, text="Plantilla (.docx):",
                 font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w")

        tk.Entry(frame, textvariable=self.docx_path, width=70)\
            .grid(row=0, column=1, padx=5)

        tk.Button(frame, text="Seleccionar...", command=self.seleccionar_docx)\
            .grid(row=0, column=2, padx=5)

        # -----------------------------
        #   LISTA DE EXCELS
        # -----------------------------
        tk.Label(frame, text="Archivos Excel (.xlsx):",
                 font=("Arial", 12, "bold")).grid(row=1, column=0, sticky="nw")

        self.listbox = tk.Listbox(frame, width=70, height=5)
        self.listbox.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(frame, text="Agregar Excel...", command=self.agregar_excel)\
            .grid(row=1, column=2, padx=5)

        # -----------------------------
        #   BOTÓN GENERAR
        # -----------------------------
        tk.Button(root, text="GENERAR DOCUMENTOS",
                  font=("Arial", 14, "bold"),
                  command=self.generar_documentos)\
            .pack(pady=10)

        # -----------------------------
        #   LOG
        # -----------------------------
        tk.Label(root, text="Registro / Log:",
                 font=("Arial", 12, "bold")).pack()

        self.log_area = ScrolledText(root, width=110, height=15)
        self.log_area.pack(pady=5)

        self.log(f"Carpeta de salida: {self.output_dir}")

    # ============================================================
    #                         FUNCIONES GUI
    # ============================================================

    def log(self, msg):
        """Escribir en el cuadro de log."""
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)

    def seleccionar_docx(self):
        """Diálogo para seleccionar plantilla."""
        path = filedialog.askopenfilename(
            filetypes=[("Documento Word", "*.docx")]
        )
        if path:
            self.docx_path.set(path)

    def agregar_excel(self):
        """Diálogo para agregar Excel y listarlos."""
        path = filedialog.askopenfilename(
            filetypes=[("Archivo Excel", "*.xlsx")]
        )
        if path:
            self.xlsx_files.append(path)
            self.listbox.insert(tk.END, path)

    # ============================================================
    #                 PROCESO PRINCIPAL DESDE LA GUI
    # ============================================================

    def generar_documentos(self):

        if not self.docx_path.get():
            messagebox.showerror("Error", "Selecciona una plantilla .docx")
            return

        if not self.xlsx_files:
            messagebox.showerror("Error", "Agrega al menos un archivo Excel")
            return

        self.log("\n------------------------------------------")
        self.log("Iniciando generación...")
        self.log("------------------------------------------\n")

        try:
            resultados = procesar_archivos(
                self.docx_path.get(),
                self.xlsx_files,
                self.output_dir,
                self.log
            )

            total = len(self.xlsx_files)
            ok = len(resultados)

            messagebox.showinfo(
                "Proceso finalizado",
                f"Generados correctamente: {ok}/{total}\n"
                f"Carpeta: {self.output_dir}"
            )

        except Exception as e:
            self.log(f"ERROR FATAL: {e}")
            messagebox.showerror("Error grave", str(e))


# ============================================================
#                          INICIO
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
