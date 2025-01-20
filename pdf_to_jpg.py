# pdf_to_jpg.py

import os
from PIL import Image
import fitz  # PyMuPDF
from tkinter import messagebox

def pdf_to_jpg(pdf_path: str, output_folder: str, dpi: int = 300) -> None:
    """
    Convierte cada página de un PDF en una imagen JPG.
    """
    try:
        pdf_document = fitz.open(pdf_path)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el PDF.\n{e}")
        return

    total_pages = len(pdf_document)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for page_number in range(total_pages):
        try:
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Nombre de salida con número de página
            output_path = os.path.join(output_folder, f"{pdf_name}_page_{page_number + 1}.jpg")
            img.save(output_path, "JPEG", quality=95)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo convertir la página {page_number + 1}.\n{e}")
            continue

    pdf_document.close()
    messagebox.showinfo(
        "Proceso completado",
        f"Imágenes JPG generadas en la carpeta:\n{output_folder}"
    )
