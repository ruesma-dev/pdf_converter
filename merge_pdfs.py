# merge_pdf,py

import fitz  # PyMuPDF
from tkinter import Tk, filedialog, Toplevel, Button, Listbox, EXTENDED
import os

def merge_pdfs(pdf_list, output_pdf):
    """
    Une varios archivos PDF en un solo archivo.

    Args:
        pdf_list (list): Lista de rutas de archivos PDF a combinar.
        output_pdf (str): Ruta del archivo PDF resultante.

    Returns:
        None
    """
    output_document = fitz.open()

    for pdf in pdf_list:
        with fitz.open(pdf) as pdf_document:
            output_document.insert_pdf(pdf_document)

    output_document.save(output_pdf)
    output_document.close()
    print(f"PDF combinado guardado en: {output_pdf}")


def reorder_pdfs(selected_pdfs):
    """
    Abre una ventana para reorganizar los PDFs seleccionados.

    Args:
        selected_pdfs (list): Lista de rutas de archivos PDF seleccionados.

    Returns:
        None
    """
    def move_up():
        selected_indices = pdf_listbox.curselection()
        for idx in selected_indices:
            if idx > 0:
                pdf_list[idx - 1], pdf_list[idx] = pdf_list[idx], pdf_list[idx - 1]
                update_listbox()

    def move_down():
        selected_indices = pdf_listbox.curselection()
        for idx in reversed(selected_indices):
            if idx < len(pdf_list) - 1:
                pdf_list[idx], pdf_list[idx + 1] = pdf_list[idx + 1], pdf_list[idx]
                update_listbox()

    def update_listbox():
        pdf_listbox.delete(0, 'end')
        for pdf in pdf_list:
            pdf_listbox.insert('end', os.path.basename(pdf))

    def execute_merge():
        if pdf_list:
            output_pdf = os.path.join(os.path.dirname(pdf_list[0]), "merged_output.pdf")
            merge_pdfs(pdf_list, output_pdf)
            reorder_window.destroy()

    # Crear la ventana para reorganizar
    reorder_window = Toplevel()
    reorder_window.title("Reorganizar PDFs")
    reorder_window.geometry("400x400")

    pdf_listbox = Listbox(reorder_window, selectmode=EXTENDED)
    pdf_listbox.pack(fill="both", expand=True, padx=10, pady=10)

    pdf_list = list(selected_pdfs)  # Convertir la tupla en una lista
    update_listbox()

    Button(reorder_window, text="Subir", command=move_up).pack(pady=5)
    Button(reorder_window, text="Bajar", command=move_down).pack(pady=5)
    Button(reorder_window, text="Ejecutar", command=execute_merge).pack(pady=20)


def select_pdfs():
    """Permite seleccionar múltiples archivos PDF y abrir la ventana de reorganización."""
    pdfs = filedialog.askopenfilenames(title="Seleccionar PDFs", filetypes=[("Archivos PDF", "*.pdf")])
    if pdfs:
        reorder_pdfs(pdfs)


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Ocultar ventana principal
    select_pdfs()
