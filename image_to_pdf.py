from PIL import Image
import os
from tkinter import Tk, filedialog, Button, Label

def images_to_pdf(input_path, output_path):
    """
    Convierte una carpeta o una lista de imágenes en un único archivo PDF.

    Args:
        input_path (str): Ruta a la carpeta o imagen.
        output_path (str): Ruta para guardar el archivo PDF.

    Returns:
        None
    """
    images = []

    # Verificar si es una carpeta o un archivo individual
    if os.path.isdir(input_path):
        # Obtener todas las imágenes de la carpeta
        valid_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
        for file in sorted(os.listdir(input_path)):
            if os.path.splitext(file)[1].lower() in valid_extensions:
                img_path = os.path.join(input_path, file)
                img = Image.open(img_path).convert('RGB')
                images.append(img)
    else:
        # Procesar una sola imagen
        img = Image.open(input_path).convert('RGB')
        images.append(img)

    # Verificar que haya imágenes
    if not images:
        print("No se encontraron imágenes válidas para convertir.")
        return

    # Guardar el PDF
    images[0].save(output_path, save_all=True, append_images=images[1:])
    print(f"PDF guardado en: {output_path}")


def select_folder_or_file():
    """Selecciona una carpeta o un archivo individual para la conversión."""
    input_path = filedialog.askdirectory(title="Selecciona una carpeta con imágenes")
    if not input_path:
        input_path = filedialog.askopenfilename(title="Selecciona una imagen", filetypes=[("Imágenes", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
    return input_path


def select_output_file():
    """Selecciona la ubicación y el nombre del archivo PDF de salida."""
    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")])
    return output_path


def main_menu():
    """Crea un menú gráfico para seleccionar opciones."""
    root = Tk()
    root.title("Conversor de Imágenes a PDF")
    root.geometry("400x200")

    Label(root, text="Conversor de Imágenes a PDF", font=("Arial", 14)).pack(pady=10)

    def execute_conversion():
        input_path = select_folder_or_file()
        if input_path:
            output_path = select_output_file()
            if output_path:
                images_to_pdf(input_path, output_path)
                root.destroy()

    Button(root, text="Seleccionar imágenes y convertir a PDF", command=execute_conversion).pack(pady=20)
    Button(root, text="Salir", command=root.destroy).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main_menu()
