# compressor.py
import subprocess
import os
import sys
import logging

def get_resource_path(relative_path: str) -> str:
    """Obtiene la ruta absoluta a un recurso en dev, onedir y onefile.

    - onefile: usa sys._MEIPASS
    - onedir: carpeta del ejecutable
    - dev: carpeta del proyecto (cwd)
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    elif getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def compress_pdf(input_pdf, output_pdf, quality):
    """
    Comprime un archivo PDF usando Ghostscript.
    """
    try:
        gs_path = get_resource_path(os.path.join("recursos", "gswin64c.exe"))

        if not os.path.exists(input_pdf):
            logging.error(f"El archivo de entrada no existe: {input_pdf}")
            return False, f"El archivo de entrada no existe: {input_pdf}"

        
        if not os.path.exists(gs_path):
            error_msg = f"No se encontró el ejecutable de Ghostscript en la ruta esperada: {gs_path}"
            logging.error(error_msg)
            return False, error_msg

        command = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_pdf}",
            f"{input_pdf}"
        ]

        
        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.CREATE_NO_WINDOW
        
        
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=creationflags)
        
        logging.info("PDF comprimido y guardado exitosamente.")
        return True, "PDF comprimido y guardado exitosamente."

    except FileNotFoundError:
        error_msg = "Error: El ejecutable de Ghostscript no se pudo encontrar."
        logging.error(error_msg)
        return False, error_msg
    except subprocess.CalledProcessError as e:
        error_msg = f"Error en el proceso de Ghostscript. Detalles: {e.stderr}"
        logging.error(error_msg)
        if os.path.exists(output_pdf):
            os.remove(output_pdf)
        return False, error_msg
    except Exception as e:
        error_msg = f"Ocurrió un error inesperado al comprimir el PDF: {e}"
        logging.error(error_msg)
        if os.path.exists(output_pdf):
            os.remove(output_pdf)
        return False, error_msg
