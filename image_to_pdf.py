# image_to_pdf.py

import os
from PIL import Image
from tkinter import messagebox

def images_to_pdf(input_path: str, output_path: str) -> None:
    """
    Convierte una carpeta con imágenes (o una sola imagen)
    en un único archivo PDF.
    """
    valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
    images = []

    if os.path.isdir(input_path):
        # Obtener todas las imágenes de la carpeta
        for file in sorted(os.listdir(input_path)):
            if os.path.splitext(file)[1].lower() in valid_extensions:
                img_path = os.path.join(input_path, file)
                try:
                    img = Image.open(img_path).convert('RGB')
                    images.append(img)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo abrir la imagen {file}.\n{e}")
    else:
        # Procesar una sola imagen
        if os.path.splitext(input_path)[1].lower() in valid_extensions:
            try:
                img = Image.open(input_path).convert('RGB')
                images.append(img)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir la imagen.\n{e}")

    if not images:
        messagebox.showwarning("Aviso", "No se encontraron imágenes válidas para convertir.")
        return

    try:
        images[0].save(output_path, save_all=True, append_images=images[1:])
        messagebox.showinfo("Proceso completado", f"PDF guardado en:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el PDF.\n{e}")
