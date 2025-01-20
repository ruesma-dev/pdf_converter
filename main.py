# main.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from image_to_pdf import images_to_pdf
from pdf_to_word import pdf_to_word
from pdf_to_jpg import pdf_to_jpg
from extract_pages import extract_pages_from_pdf

class AppUnificada(tk.Tk):
    """
    Ventana principal que unifica las funcionalidades:
    1) Convertir imágenes a PDF
    2) Convertir PDF a Word
    3) Convertir PDF a JPG
    4) Extraer páginas de PDF
    """

    def __init__(self):
        super().__init__()
        self.title("App Unificada PDF / Imágenes")
        self.geometry("550x450")
        self.resizable(False, False)

        # Variables de control
        self.option_var = tk.StringVar(value="img_to_pdf")  # Valor por defecto
        self.file_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.ocr_var = tk.BooleanVar(value=False)
        self.dpi_var = tk.StringVar(value="300")  # Solo para PDF->JPG
        self.start_page_var = tk.StringVar()
        self.end_page_var = tk.StringVar()

        # Crear interfaz
        self.create_widgets()

    def create_widgets(self):
        """Crea los elementos de la interfaz gráfica."""
        # Frame para opciones (radio buttons)
        options_frame = tk.LabelFrame(self, text="Selecciona la funcionalidad")
        options_frame.pack(fill="x", padx=10, pady=10)

        tk.Radiobutton(
            options_frame, text="Imágenes a PDF",
            variable=self.option_var, value="img_to_pdf",
            command=self.update_ui
        ).pack(anchor="w", padx=5, pady=2)

        tk.Radiobutton(
            options_frame, text="PDF a Word",
            variable=self.option_var, value="pdf_to_word",
            command=self.update_ui
        ).pack(anchor="w", padx=5, pady=2)

        tk.Radiobutton(
            options_frame, text="PDF a JPG",
            variable=self.option_var, value="pdf_to_jpg",
            command=self.update_ui
        ).pack(anchor="w", padx=5, pady=2)

        tk.Radiobutton(
            options_frame, text="Extraer páginas de PDF",
            variable=self.option_var, value="extract_pages",
            command=self.update_ui
        ).pack(anchor="w", padx=5, pady=2)

        # Frame para seleccionar archivo(s) de entrada
        input_frame = tk.LabelFrame(self, text="Entrada")
        input_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(input_frame, text="Seleccionar archivo o carpeta").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(input_frame, textvariable=self.file_path_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="Examinar", command=self.browse_input).grid(row=0, column=2, padx=5, pady=5)

        # Frame para opciones específicas
        self.options_specific_frame = tk.LabelFrame(self, text="Opciones adicionales")
        self.options_specific_frame.pack(fill="x", padx=10, pady=5)

        # # 1) OCR (para PDF a Word)
        # self.ocr_check = tk.Checkbutton(
        #     self.options_specific_frame, text="Usar OCR (para PDF->Word)",
        #     variable=self.ocr_var
        # )
        # self.ocr_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # 2) DPI (para PDF->JPG)
        self.dpi_label = tk.Label(self.options_specific_frame, text="DPI (para PDF->JPG):")
        self.dpi_entry = tk.Entry(self.options_specific_frame, textvariable=self.dpi_var, width=6)

        # 3) Rango de páginas (para Extraer páginas)
        self.start_page_label = tk.Label(self.options_specific_frame, text="Página inicial:")
        self.start_page_entry = tk.Entry(self.options_specific_frame, textvariable=self.start_page_var, width=6)

        self.end_page_label = tk.Label(self.options_specific_frame, text="Página final:")
        self.end_page_entry = tk.Entry(self.options_specific_frame, textvariable=self.end_page_var, width=6)

        # Frame para la salida (PDF, carpeta, etc.)
        output_frame = tk.LabelFrame(self, text="Salida")
        output_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(output_frame, text="Archivo/Carpeta destino").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(output_frame, textvariable=self.output_path_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(output_frame, text="Guardar en...", command=self.browse_output).grid(row=0, column=2, padx=5, pady=5)

        # Botón ejecutar
        tk.Button(self, text="Ejecutar", command=self.execute_action, bg="green", fg="white", font=("Arial", 12)).pack(pady=20)

        self.update_ui()

    def browse_input(self):
        """Selecciona archivo(s) o carpeta según la opción activa."""
        option = self.option_var.get()

        if option == "img_to_pdf":
            # Primero intentamos seleccionar una carpeta; si no, seleccionamos una imagen
            folder = filedialog.askdirectory(title="Selecciona una carpeta de imágenes")
            if folder:
                self.file_path_var.set(folder)
            else:
                # Seleccionar archivo de imagen
                filetypes = [("Imágenes", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")]
                file = filedialog.askopenfilename(title="Selecciona una imagen", filetypes=filetypes)
                if file:
                    self.file_path_var.set(file)
        else:
            # En las otras opciones, seleccionamos un PDF
            filetypes = [("Archivos PDF", "*.pdf")]
            file = filedialog.askopenfilename(title="Selecciona un archivo PDF", filetypes=filetypes)
            if file:
                self.file_path_var.set(file)

    def browse_output(self):
        """Selecciona la ruta de salida, que puede ser archivo o carpeta."""
        option = self.option_var.get()

        if option in ["img_to_pdf", "extract_pages"]:
            # Guardar como PDF
            file = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Archivos PDF", "*.pdf")],
                title="Guardar PDF"
            )
            if file:
                self.output_path_var.set(file)
        elif option in ["pdf_to_word", "pdf_to_jpg"]:
            # Seleccionar carpeta de salida
            folder = filedialog.askdirectory(title="Selecciona carpeta de salida")
            if folder:
                self.output_path_var.set(folder)

    def update_ui(self):
        """Muestra u oculta las opciones relevantes según la funcionalidad."""
        option = self.option_var.get()

        # Ocultar todos los widgets específicos
        # self.ocr_check.grid_remove()
        self.dpi_label.grid_remove()
        self.dpi_entry.grid_remove()
        self.start_page_label.grid_remove()
        self.start_page_entry.grid_remove()
        self.end_page_label.grid_remove()
        self.end_page_entry.grid_remove()

        if option == "pdf_to_word":
        #     self.ocr_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            pass
        elif option == "pdf_to_jpg":
            self.dpi_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
            self.dpi_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        elif option == "extract_pages":
            self.start_page_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
            self.start_page_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
            self.end_page_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
            self.end_page_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    def execute_action(self):
        """Ejecuta la funcionalidad seleccionada."""
        option = self.option_var.get()
        input_path = self.file_path_var.get().strip()
        output_path = self.output_path_var.get().strip()

        if not input_path:
            messagebox.showwarning("Aviso", "Por favor, selecciona la ruta de entrada.")
            return

        if option in ["img_to_pdf", "extract_pages"]:
            if not output_path:
                messagebox.showwarning("Aviso", "Por favor, selecciona la ruta de salida.")
                return
        elif option in ["pdf_to_word", "pdf_to_jpg"]:
            if not output_path or not os.path.isdir(output_path):
                messagebox.showwarning("Aviso", "Por favor, selecciona una carpeta de salida válida.")
                return

        try:
            if option == "img_to_pdf":
                images_to_pdf(input_path, output_path)

            elif option == "pdf_to_word":
                use_ocr = self.ocr_var.get()
                pdf_to_word(input_path, output_path, use_ocr)

            elif option == "pdf_to_jpg":
                try:
                    dpi_value = int(self.dpi_var.get())
                    if dpi_value <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showwarning("Aviso", "Por favor, introduce un DPI válido (número entero positivo).")
                    return
                pdf_to_jpg(input_path, output_path, dpi=dpi_value)

            elif option == "extract_pages":
                # Validar las entradas de páginas
                start_page_str = self.start_page_var.get().strip()
                end_page_str = self.end_page_var.get().strip()

                if not (start_page_str.isdigit() and end_page_str.isdigit()):
                    messagebox.showwarning(
                        "Aviso",
                        "Por favor, introduce números válidos en 'Página inicial' y 'Página final'."
                    )
                    return

                start_page = int(start_page_str)
                end_page = int(end_page_str)

                extract_pages_from_pdf(input_path, output_path, start_page, end_page)

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")

if __name__ == "__main__":
    app = AppUnificada()
    app.mainloop()
