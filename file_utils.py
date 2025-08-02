# file_utils.py
import os
from docx2pdf import convert
from PIL import Image

def secure_delete_file(file_path):
    """Sobrescribe un archivo con ceros antes de eliminarlo."""
    if os.path.exists(file_path):
        try:
            with open(file_path, "wb") as f:
                f.seek(0)
                f.write(b'\0' * os.path.getsize(file_path))
            os.remove(file_path)
        except Exception:
            # En caso de error, simplemente borra el archivo
            os.remove(file_path)

def image_to_pdf(image_path, output_pdf_path):
    """Convierte un archivo de imagen (JPG, PNG) a PDF."""
    image = Image.open(image_path)
    if image.mode != "RGB":
        image = image.convert("RGB")
    image.save(output_pdf_path, "PDF", resolution=100.0)

def word_to_pdf(word_path, output_pdf_path):
    """Convierte un archivo de Word (DOCX) a PDF."""
    try:
        convert(word_path, output_pdf_path)
        return True
    except Exception as e:
        print(f"Error al convertir DOCX a PDF: {e}")
        return False