# extract_pages.py

import os
import fitz  # PyMuPDF
from tkinter import messagebox

def extract_pages_from_pdf(input_pdf: str, output_pdf: str, start_page: int, end_page: int) -> None:
    """
    Extrae páginas específicas de un archivo PDF y las guarda en un nuevo PDF.
    """
    try:
        pdf_document = fitz.open(input_pdf)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el PDF de entrada.\n{e}")
        return

    if (
        start_page < 1 or
        end_page > len(pdf_document) or
        start_page > end_page
    ):
        messagebox.showwarning("Aviso", "Rango de páginas inválido.")
        pdf_document.close()
        return

    output_document = fitz.open()

    try:
        for page_num in range(start_page - 1, end_page):
            output_document.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron extraer las páginas.\n{e}")
        output_document.close()
        pdf_document.close()
        return

    try:
        output_document.save(output_pdf)
        messagebox.showinfo(
            "Proceso completado",
            f"Páginas {start_page}-{end_page} extraídas y guardadas en:\n{output_pdf}"
        )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el PDF de salida.\n{e}")
    finally:
        output_document.close()
        pdf_document.close()
