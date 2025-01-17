# extraer_paginas.py

import fitz  # PyMuPDF
from tkinter import Tk, filedialog, simpledialog


def extract_pages_from_pdf(input_pdf, output_pdf, start_page, end_page):
    """
    Extrae páginas específicas de un archivo PDF y las guarda en un nuevo PDF.

    Args:
        input_pdf (str): Ruta del archivo PDF de entrada.
        output_pdf (str): Ruta para guardar el archivo PDF resultante.
        start_page (int): Página inicial (base 1).
        end_page (int): Página final (base 1).

    Returns:
        None
    """
    # Abrir el PDF original
    pdf_document = fitz.open(input_pdf)

    # Verificar que los índices estén dentro del rango
    if start_page < 1 or end_page > len(pdf_document) or start_page > end_page:
        print("Rango de páginas inválido.")
        return

    # Crear un nuevo PDF
    output_document = fitz.open()

    # Extraer las páginas especificadas
    for page_num in range(start_page - 1, end_page):
        output_document.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

    # Guardar el nuevo PDF
    output_document.save(output_pdf)
    output_document.close()
    pdf_document.close()

    print(f"Páginas {start_page}-{end_page} extraídas y guardadas en: {output_pdf}")


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Ocultar ventana principal

    # Seleccionar archivo PDF de entrada
    input_pdf = filedialog.askopenfilename(title="Selecciona un archivo PDF", filetypes=[("Archivos PDF", "*.pdf")])
    if not input_pdf:
        print("No se seleccionó ningún archivo.")
        exit()

    # Seleccionar archivo PDF de salida
    output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")])
    if not output_pdf:
        print("No se seleccionó una ubicación de salida.")
        exit()

    # Pedir rango de páginas
    start_page = simpledialog.askinteger("Página Inicial", "Ingresa la página inicial (base 1):")
    end_page = simpledialog.askinteger("Página Final", "Ingresa la página final (base 1):")

    if start_page and end_page:
        extract_pages_from_pdf(input_pdf, output_pdf, start_page, end_page)
