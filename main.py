# main.py
import sys
import os
import subprocess
import logging
import math
from PyQt5.QtWidgets import QApplication, QWidget, QListWidgetItem, QComboBox, QProgressBar, QTabWidget, QGridLayout, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox, QHBoxLayout
from PyQt5.QtGui import QPixmap, QImage, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QUrl, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtWidgets import QApplication, QWidget, QListWidgetItem, QListWidget, QComboBox, QProgressBar, QTabWidget, QGridLayout, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox, QHBoxLayout, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import QRect, QPoint, Qt
from pdf_utils import remove_selected_pages
from compressor import compress_pdf
from file_utils import secure_delete_file, image_to_pdf, word_to_pdf
import fitz 
from PIL import Image

logging.basicConfig(filename='app_activity.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

class Worker(QThread):
    finished = pyqtSignal(list, str)
    progress_update = pyqtSignal(int)
    error = pyqtSignal(str)
    
    def __init__(self, task, input_data=None):
        super().__init__()
        self.task = task
        self.input_data = input_data

    def run(self):
        try:
            if self.task == "render_pdf":
                pdf_path = self.input_data
                doc = fitz.open(pdf_path)
                images = []
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=50)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
                    self.progress_update.emit(int((i + 1) / len(doc) * 100))
                doc.close()
                self.finished.emit(images, pdf_path)
            
            elif self.task == "process_pdf":
                input_pdf, out_path, quality = self.input_data
                
                self.progress_update.emit(10)
                compress_pdf(input_pdf, out_path, quality)
                self.progress_update.emit(100)
                
                self.finished.emit([], out_path)

        except Exception as e:
            self.error.emit(str(e))

class WordToPDFTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.output_pdf = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.label = QLabel("Arrastra y suelta un archivo de Word aquí o usa el botón para convertirlo a PDF.")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 16px; color: gray;")
        layout.addWidget(self.label)
        
        self.open_button = QPushButton("Seleccionar Archivo de Word")
        self.open_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.open_button)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.info_label = QLabel("")
        layout.addWidget(self.info_label)

        self.setLayout(layout)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.isLocalFile() for url in event.mimeData().urls()):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.handle_file(url.toLocalFile())
            break

    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo de Word", "", "Archivos Word (*.docx)")
        if file_name:
            self.handle_file(file_name)

    def handle_file(self, file_path):
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Advertencia", "El archivo no existe.")
            return
        
        if not file_path.lower().endswith(".docx"):
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona un archivo de Word (.docx).")
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.info_label.setText("Convirtiendo a PDF...")
        
        self.output_pdf = file_path.replace(".docx", "_convertido.pdf")
        
        if word_to_pdf(file_path, self.output_pdf):
            self.progress_bar.setValue(100)
            self.info_label.setText(f"¡Conversión exitosa! Archivo guardado en:\n{self.output_pdf}")
            logging.info(f"Conversión de Word a PDF exitosa. Archivo: {file_path}")
        else:
            self.progress_bar.setValue(0)
            self.info_label.setText("Error al convertir. Asegúrate de tener Microsoft Word instalado.")
            logging.error(f"Error al convertir Word a PDF: {file_path}")
        
        self.progress_bar.setVisible(False)


class PageRemoverTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.input_pdf = None
        self.pages = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.open_button = QPushButton("Abrir PDF")
        self.open_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.open_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)
        self.list_widget.setViewMode(QListWidget.IconMode)
        self.list_widget.setIconSize(QSize(300, 400))
        self.list_widget.setGridSize(QSize(320, 420))
        layout.addWidget(self.list_widget)

        hlayout = QHBoxLayout()
        self.remove_button = QPushButton("Eliminar Páginas Seleccionadas y Guardar como...")
        self.remove_button.clicked.connect(self.remove_pages_and_save)
        self.remove_button.setEnabled(False)
        hlayout.addWidget(self.remove_button)

        layout.addLayout(hlayout)
        self.setLayout(layout)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.isLocalFile() for url in event.mimeData().urls()):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.handle_file(url.toLocalFile())
            break

    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf)")
        if file_name:
            self.handle_file(file_name)

    def handle_file(self, file_path):
        self.input_pdf = file_path
        self.list_widget.clear()
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.remove_button.setEnabled(False)
        
        self.worker = Worker("render_pdf", file_path)
        self.worker.progress_update.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self.on_pdf_rendered)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        logging.info(f"PDF abierto para eliminación de páginas: {file_path}")

    def on_pdf_rendered(self, images, pdf_path):
        self.pages = images
        for i, img in enumerate(self.pages):
            pixmap = QPixmap.fromImage(QImage(img.tobytes(), img.width, img.height, img.width*3, QImage.Format_RGB888))
            item = QListWidgetItem(f"Página {i+1}")
            item.setIcon(QIcon(pixmap.scaled(self.list_widget.iconSize(), Qt.KeepAspectRatio, Qt.SmoothTransformation)))
            self.list_widget.addItem(item)
        
        self.progress_bar.setVisible(False)
        self.remove_button.setEnabled(True)

    def remove_pages_and_save(self):
        selected = self.list_widget.selectedIndexes()
        to_remove = [i.row() for i in selected]
        
        if not to_remove:
            QMessageBox.warning(self, "Advertencia", "No has seleccionado ninguna página para eliminar.")
            return

        out_path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF sin páginas", "", "Archivos PDF (*.pdf)")
        if out_path:
            remove_selected_pages(self.input_pdf, out_path, to_remove)
            QMessageBox.information(self, "Éxito", "PDF guardado exitosamente sin las páginas seleccionadas.")
            logging.info(f"Páginas eliminadas de {self.input_pdf}. Nuevo archivo guardado en: {out_path}")
            
    def on_error(self, message):
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Error", f"Ocurrió un error: {message}")
        logging.error(f"Error en el eliminador de páginas: {message}")


class PDFCompressorTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.input_pdf = None
        self.original_size = 0
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.label = QLabel("Selecciona un archivo PDF para comprimirlo y guardarlo.")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 16px; color: gray;")
        layout.addWidget(self.label)
        
        self.open_button = QPushButton("Abrir PDF")
        self.open_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.open_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.size_layout = QHBoxLayout()
        self.original_size_label = QLabel("Tamaño Original:")
        self.compressed_size_label = QLabel("Tamaño Comprimido:")
        self.original_size_label.setStyleSheet("font-weight: bold;")
        self.compressed_size_label.setStyleSheet("font-weight: bold;")
        self.size_layout.addWidget(self.original_size_label)
        self.size_layout.addStretch()
        self.size_layout.addWidget(self.compressed_size_label)
        layout.addLayout(self.size_layout)

        hlayout = QHBoxLayout()
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["Compresion Extrema (menos calidad)", "Compresion Recomendada (Buena Calidad)", "Baja Compresion (Alta Calidad)"])
        hlayout.addWidget(QLabel("Tipo de Compresion:"))
        hlayout.addWidget(self.quality_combo)
        
        self.compress_button = QPushButton("Comprimir PDF y Guardar como...")
        self.compress_button.clicked.connect(self.process_pdf)
        self.compress_button.setEnabled(False)
        hlayout.addWidget(self.compress_button)

        layout.addLayout(hlayout)
        self.setLayout(layout)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.isLocalFile() for url in event.mimeData().urls()):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.handle_file(url.toLocalFile())
            break

    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf)")
        if file_name:
            self.handle_file(file_name)
    
    def handle_file(self, file_path):
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Advertencia", "El archivo no existe.")
            return
        
        if not file_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona un archivo PDF.")
            return

        self.input_pdf = file_path
        self.original_size = os.path.getsize(file_path)
        self.original_size_label.setText(f"Tamaño Original: {self.format_size(self.original_size)}")
        self.compressed_size_label.setText("Tamaño Comprimido:")
        self.compress_button.setEnabled(True)
        logging.info(f"PDF abierto para compresión: {file_path}")

    def process_pdf(self):
        if not self.input_pdf:
            QMessageBox.warning(self, "Advertencia", "Primero seleccioná un archivo PDF.")
            return

        try:
            subprocess.run(["gswin64c", "--version"], check=True, capture_output=True, text=True)
        except (FileNotFoundError, subprocess.CalledProcessError):
            QMessageBox.critical(self, "Error", 
                "Ghostscript no está instalado o no se encuentra en el PATH.\n"
                "Por favor, instalalo para usar la función de compresión.")
            return
        
        out_path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF comprimido", "", "Archivos PDF (*.pdf)")
        if out_path:
            quality_map = {
                0: "screen",
                1: "ebook",
                2: "printer"
            }
            quality = quality_map[self.quality_combo.currentIndex()]

            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.compress_button.setEnabled(False)
            
            logging.info(f"Compresión iniciada para: {self.input_pdf}. Tipo de Compresion: {quality}")

            self.worker = Worker("process_pdf", (self.input_pdf, out_path, quality))
            self.worker.progress_update.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.on_pdf_processed)
            self.worker.error.connect(self.on_error)
            self.worker.start()

    def on_pdf_processed(self, _, out_path):
        self.progress_bar.setVisible(False)
        self.compress_button.setEnabled(True)
        QMessageBox.information(self, "Éxito", "PDF comprimido y guardado exitosamente.")
        
        compressed_size = os.path.getsize(out_path)
        self.compressed_size_label.setText(f"Tamaño Comprimido: {self.format_size(compressed_size)}")
        logging.info(f"Compresión finalizada. Archivo guardado en: {out_path}")

    def on_error(self, message):
        self.progress_bar.setVisible(False)
        self.compress_button.setEnabled(True)
        QMessageBox.critical(self, "Error", f"Ocurrió un error: {message}")
        logging.error(f"Error en el compresor de PDF: {message}")

    def format_size(self, size_bytes):
        if size_bytes == 0:
            return "0 B"
        size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 2)
        return f"{s} {size_name[i]}"

class PDFToolApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Herramientas de Compresion Judiciales - By Germán Rojas")
        self.setGeometry(100, 100, 1200, 900)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        self.setWindowIcon(QIcon(resource_path('icon.ico')))
        
        self.word_to_pdf_tab = WordToPDFTab()
        self.page_remover_tab = PageRemoverTab()
        self.pdf_compressor_tab = PDFCompressorTab()

        self.tabs.addTab(self.word_to_pdf_tab, "Convertir Word a PDF")
        self.tabs.addTab(self.page_remover_tab, "Eliminar Páginas")
        self.tabs.addTab(self.pdf_compressor_tab, "Comprimir PDF")
        
        layout.addWidget(self.tabs)

        credits_label = QLabel("Creado por German Rojas")
        credits_label.setAlignment(Qt.AlignCenter)
        credits_label.setStyleSheet("font-size: 10px; color: gray;")
        layout.addWidget(credits_label)

        self.setLayout(layout)

def resource_path(relative_path):
    """Obtiene la ruta absoluta a un recurso, para que funcione tanto en desarrollo como en PyInstaller."""
    try:
        
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    
    app.setWindowIcon(QIcon(resource_path('icon.ico')))
    
    
    splash_pixmap = QPixmap(resource_path('logo.png'))
    splash_pixmap = splash_pixmap.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    splash = QSplashScreen(splash_pixmap)
    splash.setWindowFlags(Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
    
   
    splash.show()
    
    
    load_timer = QTimer()
    global step
    step = 0
    def update_splash_message():
        global step
        messages = ["Iniciando.", "Iniciando..", "Iniciando...", "Cargando.", "Cargando..", "Cargando..."]
        step = (step + 1) % len(messages)
        splash.showMessage(messages[step], Qt.AlignBottom | Qt.AlignCenter, Qt.white)

    load_timer.timeout.connect(update_splash_message)
    load_timer.start(500) 
    

    window = PDFToolApp()
    window.setWindowOpacity(0.0) 
    
    
    anim = QPropertyAnimation(splash, b"windowOpacity")
    anim.setDuration(1500)
    anim.setStartValue(1)
    anim.setEndValue(0)
    anim.setEasingCurve(QEasingCurve.InQuad)
    
    
    main_anim = QPropertyAnimation(window, b"windowOpacity")
    main_anim.setDuration(1000) 
    main_anim.setStartValue(0)
    main_anim.setEndValue(1)
    
    anim.finished.connect(splash.close)
    anim.finished.connect(main_anim.start)

    
    def finish_splash_and_show_main():
        load_timer.stop()
        window.show() 
        anim.start()
    
   
    QTimer.singleShot(3000, finish_splash_and_show_main)
    
    sys.exit(app.exec_())
