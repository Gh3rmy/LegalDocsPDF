import subprocess
import os

def compress_pdf(input_path, output_path, quality='ebook'):
    try:
        # Configuración para ejecutar sin ventana en Windows
        startupinfo = None
        if os.name == 'nt':  # Verifica si el sistema operativo es Windows
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        subprocess.run([
            "gswin64c",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            input_path
        ], check=True, startupinfo=startupinfo)
        
    except FileNotFoundError:
        print("⚠️ Ghostscript no está instalado o no se encuentra en el PATH.")
        raise