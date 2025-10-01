from PIL import ImageGrab
import pytesseract
from docx import Document
import os

# Configura Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Capturar imagen del portapapeles
imagen = ImageGrab.grabclipboard()
if imagen is None:
    print("No hay imagen en el portapapeles")
    exit()

# Extraer texto y limpiar caracteres problemáticos
texto = pytesseract.image_to_string(imagen)
texto = texto.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
texto = texto.strip()

print("Texto extraído:\n", texto)

# Guardar en Word en la carpeta de OneDrive
ruta_carpeta = r" / " #agregar ruta de carpeta a la cual ira el archivo 
if not os.path.exists(ruta_carpeta):
    os.makedirs(ruta_carpeta)  # Crear la carpeta si no existe

ruta_word = os.path.join(ruta_carpeta, "codigo_recortado.docx")

doc = Document()
doc.add_paragraph(texto)

try:
    doc.save(ruta_word)
    print(f"Archivo Word generado en: {ruta_word}")
except Exception as e:
    print(f"Error al guardar el archivo: {e}")
