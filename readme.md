# Conversor de PDF a Word, Imágenes o PDF desde Imágenes con OCR

Este repositorio contiene un conjunto de scripts en Python para:
- Convertir archivos PDF en documentos Word o imágenes JPG.
- Crear archivos PDF a partir de imágenes o carpetas de imágenes.
- Incluir soporte para reconocimiento óptico de caracteres (OCR) mediante la biblioteca `pytesseract` en caso de que el PDF contenga texto incrustado en imágenes.

## Características
- **Conversión a Word:** Convierte el contenido del PDF en un documento Word, incluyendo texto e imágenes.
- **Conversión a JPG:** Guarda cada página del PDF como imágenes en formato JPG.
- **OCR Integrado:** Utiliza Tesseract-OCR para reconocer texto en imágenes incrustadas en el PDF.
- **Conversión de Imágenes a PDF:** Permite convertir una carpeta de imágenes o imágenes individuales en un archivo PDF.
- **Interfaz Gráfica de Usuario (GUI):** Selecciona archivos o carpetas y elige el formato de salida mediante una interfaz gráfica.
- **Gestión Automática de Carpetas:** Guarda los archivos de salida en subcarpetas con nombres descriptivos.

## Requisitos
- Python 3.x
- Bibliotecas necesarias:
  - `pymupdf` (PyMuPDF)
  - `pillow` (PIL)
  - `pytesseract`
  - `tkinter`

## Instalación
1. Clona el repositorio:
   ```bash
   git clone https://github.com/tuusuario/conversor-pdf.git
   cd conversor-pdf
   ```
2. Instala las dependencias:
   ```bash
   pip install pymupdf pillow pytesseract
   ```
3. Instala Tesseract-OCR:
   - **Windows:** Descarga e instala desde [Tesseract OCR](https://github.com/tesseract-ocr/tesseract).
   - **Linux:**
     ```bash
     sudo apt install tesseract-ocr
     ```
   - **MacOS:**
     ```bash
     brew install tesseract
     ```
4. Configura la ruta de Tesseract si no está en el PATH:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

## Uso
### 1. Conversión de PDF a Word o JPG
1. Ejecuta el script:
   ```bash
   python pdf_converter.py
   ```
2. Selecciona un archivo PDF mediante el cuadro de diálogo.
3. Elige el formato de salida (Word o JPG).
4. (Opcional) Marca el checkbox para activar OCR en caso de que el PDF contenga imágenes.
5. Haz clic en **Ejecutar**.
6. El archivo de salida se guardará en una subcarpeta dentro de la carpeta donde está ubicado el archivo PDF original.

### 2. Conversión de Imágenes a PDF
1. Ejecuta el script:
   ```bash
   python images_to_pdf.py
   ```
2. Selecciona una carpeta de imágenes o una imagen individual mediante el cuadro de diálogo.
3. Especifica el nombre y ubicación del archivo PDF de salida.
4. Haz clic en **Ejecutar**.
5. El archivo PDF se guardará en la ubicación seleccionada.

## Estructura del Proyecto
```
.
├── pdf_converter.py        # Conversor de PDF a Word o JPG con OCR
├── images_to_pdf.py        # Conversor de imágenes a PDF
├── README.md               # Documentación del proyecto
└── requirements.txt        # Dependencias
```

## Licencia
Este proyecto está licenciado bajo la licencia MIT. Consulte el archivo LICENSE para más detalles.

## Contribuciones
¡Las contribuciones son bienvenidas! Por favor, crea un fork del repositorio y envía un pull request con tus mejoras.

## Autor
- **Tu Nombre** - [GitHub](https://github.com/tuusuario)
