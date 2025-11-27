import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pathlib

# IMPORTAMOS TU MOTOR REAL
from generador import procesar_archivo_excel


# ====================================
#     INTERFAZ GRÁFICA (GUI)
# ====================================

class App:
    def __init__(self, root):
        self.root = root
        root.title("Generador Automático de Protocolos v1.0")
        root.geometry("850x600")

        self.docx_path = tk.StringVar()
        self.xlsx_files = []

        # Crear carpeta segura en Documentos
        documents = pathlib.Path.home() / "Documents" / "Protocolos_Generados"
        os.makedirs(documents, exist_ok=True)
        self.output_dir = str(documents)

        # -------------------------
        #         UI
        # -------------------------
        tk.Label(root, text="Generador Automático de Protocolos v1.0",
                 font=("Arial", 18, "bold")).pack(pady=10)

        frame = tk.Frame(root)
        frame.pack(pady=5)

        # PLANTILLA DOCX
        tk.Label(frame, text="Plantilla (.docx):", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.docx_path, width=70).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Seleccionar...", command=self.seleccionar_docx).grid(row=0, column=2, padx=5)

        # EXCELS
        tk.Label(frame, text="Archivos Excel (.xlsx):", font=("Arial", 12, "bold")).grid(row=1, column=0, sticky="nw")
        self.listbox = tk.Listbox(frame, width=70, height=5)
        self.listbox.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(frame, text="Agregar Excel...", command=self.agregar_excel).grid(row=1, column=2, padx=5)

        # BOTÓN GENERAR
        tk.Button(root, text="GENERAR DOCUMENTOS", font=("Arial", 14, "bold"),
                  command=self.generar_documentos).pack(pady=10)

        # LOG
        tk.Label(root, text="Registro / Log:", font=("Arial", 12, "bold")).pack()
        self.log_area = ScrolledText(root, width=100, height=15)
        self.log_area.pack(pady=5)

        self.log(f"Carpeta de salida: {self.output_dir}")

    # --------------------------------
    #           LOG
    # --------------------------------
    def log(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)

    # --------------------------------
    #       Seleccionar archivos
    # --------------------------------
    def seleccionar_docx(self):
        path = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if path:
            self.docx_path.set(path)

    def agregar_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.xlsx_files.append(path)
            self.listbox.insert(tk.END, path)

    # --------------------------------
    #       GENERAR DOCUMENTOS
    # --------------------------------
    def generar_documentos(self):
        if not self.docx_path.get():
            messagebox.showerror("Error", "Selecciona una plantilla .docx")
            return

        if not self.xlsx_files:
            messagebox.showerror("Error", "Agrega al menos un archivo Excel")
            return

        self.log("Iniciando generación...\n")

        template = self.docx_path.get()
        ok = 0

        for x in self.xlsx_files:
            self.log(f"Procesando: {x}")

            éxito = procesar_archivo_excel(
                template_path=template,
                excel_path=x,
                output_folder=self.output_dir,
                log_func=self.log
            )

            if éxito:
                ok += 1

        messagebox.showinfo(
            "Proceso finalizado",
            f"Generados correctamente: {ok}/{len(self.xlsx_files)}\n"
            f"Carpeta: {self.output_dir}"
        )


# MAIN
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
