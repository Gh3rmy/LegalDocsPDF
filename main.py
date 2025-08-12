import sys
import os
import subprocess
import logging
if sys.platform == "win32" and hasattr(sys, 'frozen'):
    # Usamos la carpeta AppData/Local de manera estándar
    win32com_cache_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'win32com', 'gen_py')
    os.environ['PYWIN32_GENPY_DIR'] = win32com_cache_dir
    # También nos aseguramos de que el directorio exista
    if not os.path.exists(win32com_cache_dir):
        os.makedirs(win32com_cache_dir)



import math
import io
import fitz 
from PIL import Image
from docx2pdf import convert
from PyQt5.QtWidgets import (QApplication, QWidget, QListWidgetItem, QComboBox, 
                             QProgressBar, QTabWidget, QGridLayout, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QMessageBox, 
                             QHBoxLayout, QListView, QCheckBox, QSizePolicy, 
                             QAbstractItemView, QListWidget, QSplashScreen)
from PyQt5.QtGui import QPixmap, QImage, QIcon
from PyQt5.QtCore import (Qt, QThread, pyqtSignal, QSize, QUrl, QTimer, 
                          QPropertyAnimation, QEasingCurve, QRect, QPoint)
import win32com

# --- Añadir carpeta gen_py para win32com antes de importar módulos que la usen ---
genpy_path = os.path.join(os.path.dirname(__file__), 'recursos', 'gen_py')
if os.path.exists(genpy_path):
    sys.path.append(genpy_path)


if sys.platform == "win32" and hasattr(sys, 'frozen'):
    try:
        temp_dir = os.path.join(os.environ["TEMP"], "win32com_gen_py")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        win32com.__path__.insert(0, temp_dir)
    except Exception as e:
        logging.error(f"Error al configurar el directorio temporal de win32com: {e}")

import pythoncom
import win32com.client
import pywintypes
try:
    win32com.client.gencache.SetEnabled = 0
except Exception as e:
    logging.warning(f"No se pudo deshabilitar la caché de win32com: {e}")
# Importaciones de tus utilidades
from pdf_utils import remove_selected_pages
from compressor import compress_pdf
from file_utils import secure_delete_file, image_to_pdf, word_to_pdf


def configure_logging():
    if sys.platform == "win32":
        appdata_path = os.path.join(os.getenv('LOCALAPPDATA'), "LegalDocs")
    else:
        appdata_path = os.path.join(os.path.expanduser("~"), ".LegalDocs")

    if not os.path.exists(appdata_path):
        os.makedirs(appdata_path)

    log_file_path = os.path.join(appdata_path, 'app_activity.log')

    logging.basicConfig(
        filename=log_file_path,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

configure_logging()


if sys.stdout is None:
    try:
        sys.stdout = open(os.devnull, 'w')
    except Exception:
        sys.stdout = io.StringIO()
if sys.stderr is None:
    try:
        sys.stderr = open(os.devnull, 'w')
    except Exception:
        sys.stderr = io.StringIO()


def ensure_com_modules():
    import win32com.client
    try:
        win32com.client.Dispatch('Word.Application')
    except Exception as e:
        
        print(f"Warning: fallo al crear instancia COM: {e}")


class Worker(QThread):
    finished = pyqtSignal(list, str)
    finished_compression = pyqtSignal(bool, str) 
    progress_update = pyqtSignal(int)
    error = pyqtSignal(str)
    sizes_updated = pyqtSignal(float, float) 
    
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
                
                original_size_bytes = os.path.getsize(input_pdf)
                original_mb = original_size_bytes / (1024 * 1024)

                self.progress_update.emit(10)
                
                success, message = compress_pdf(input_pdf, out_path, quality)
                
                if success:
                    try:
                        compressed_size_bytes = os.path.getsize(out_path)
                        compressed_mb = compressed_size_bytes / (1024 * 1024)
                        self.sizes_updated.emit(original_mb, compressed_mb)
                        self.finished_compression.emit(True, "PDF comprimido y guardado exitosamente.")
                    except Exception as e:
                        error_msg = f"Compresión exitosa, pero no se pudo obtener el tamaño del archivo: {e}"
                        self.finished_compression.emit(False, error_msg)
                    finally:
                        self.progress_update.emit(100)
                else:
                    self.finished_compression.emit(False, message)

        except Exception as e:
            self.error.emit(str(e))


class WordToPDFWorker(QThread):
    progress_update = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, docx_path, output_pdf_path):
        super().__init__()
        self.docx_path = docx_path
        self.output_pdf_path = output_pdf_path

    def run(self):
        try:
            
            pythoncom.CoInitialize()

            self.progress_update.emit(10)

            
            src = os.path.normpath(os.path.abspath(self.docx_path))
            dst = os.path.normpath(os.path.abspath(self.output_pdf_path))
            os.makedirs(os.path.dirname(dst), exist_ok=True)

            
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx('Word.Application')
                word.Visible = False
                try:
                    word.DisplayAlerts = 0  
                except Exception:
                    pass

                doc = word.Documents.Open(src, ReadOnly=True)
                
                doc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17)
            finally:
                try:
                    if doc is not None:
                        doc.Close(False)
                except Exception:
                    pass
                try:
                    if word is not None:
                        word.Quit()
                except Exception:
                    pass

            self.progress_update.emit(100)
            self.finished.emit(True, self.output_pdf_path)

        except pywintypes.com_error as e:
            self.finished.emit(False, f"Error de conversión (COM). Verifica que Microsoft Word esté instalado y activado. Detalle: {e}")
        except Exception as e:
            self.finished.emit(False, f"Error de conversión. Verifica que Microsoft Word esté instalado y activado. Detalle: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


class WordToPDFTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.output_pdf = None
        self.init_ui()
        self.setAcceptDrops(True)

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.label = QLabel("Arrastra y suelta un archivo de Word aquí o usa el botón para convertirlo a PDF.")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
        font-size: 16px; 
        color: gray;
        min-height: 300px;
        """)
        layout.addWidget(self.label)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.info_label = QLabel("")
        layout.addWidget(self.info_label)

        layout.addStretch()

        self.open_button = QPushButton("Seleccionar Archivo de Word")
        self.open_button.clicked.connect(self.open_file_dialog)
        
        self.open_button.setFixedSize(250, 40)
        self.open_button.setToolTip("Abre el explorador de archivos para seleccionar un documento Word (.docx) y convertirlo a PDF.")
        self.open_button.setIcon(QIcon(resource_path('recursos/folder.png')))
        self.open_button.setStyleSheet("font-size: 14px;")

        layout.addWidget(self.open_button, alignment=Qt.AlignCenter)

        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.isLocalFile() and url.toLocalFile().lower().endswith(".docx") for url in event.mimeData().urls()):
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
        self.info_label.setAlignment(Qt.AlignCenter)
        
        output_pdf = os.path.abspath(file_path.replace(".docx", "_convertido.pdf"))
        
        self.worker = WordToPDFWorker(file_path, output_pdf)
        self.worker.progress_update.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.start()
        logging.info(f"Conversión de Word a PDF iniciada para: {file_path}")

    def on_conversion_finished(self, success, output_path):
        self.progress_bar.setVisible(False)
        if success:
            self.info_label.setText(f"¡Conversión exitosa! Archivo guardado en:\n{output_path}")
            logging.info(f"Conversión de Word a PDF exitosa. Archivo: {output_path}")
            QMessageBox.information(self, "Conversión exitosa", f"Archivo guardado en:\n{output_path}")
        else:
            self.info_label.setText(f"Error al convertir. El problema podría ser Pandoc o el archivo.")
            logging.error(f"Error al convertir Word a PDF: {output_path}")


class PageRemoverTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.input_pdf = None
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.list_widget = QListWidget()
        self.list_widget.setFlow(QListView.LeftToRight)
        self.list_widget.setViewMode(QListView.IconMode)
        self.list_widget.setIconSize(QSize(300, 400))
        self.list_widget.setGridSize(QSize(320, 420))
        self.list_widget.setResizeMode(QListView.Adjust)
        self.list_widget.setWrapping(True)
        self.list_widget.setSpacing(10)
        self.list_widget.setDragEnabled(False)
        self.list_widget.setAcceptDrops(False)
        self.list_widget.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        self.list_widget.setStyleSheet("""
            QListWidget::item:selected {
                background-color: #007bff;
                color: white;
            }
        """)
        
        self.list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_layout.addWidget(self.list_widget)

        hlayout = QHBoxLayout()
        hlayout.addStretch()

        self.open_button = QPushButton("Seleccionar PDF")
        self.open_button.setFixedSize(150, 40)
        self.open_button.clicked.connect(self.open_file_dialog)
        self.open_button.setIcon(QIcon(resource_path('recursos/pdflogo.png')))
        self.open_button.setToolTip("Abre el explorador de archivos para seleccionar un PDF.")
        self.open_button.setStyleSheet("""font-size: 14px;""")
        hlayout.addWidget(self.open_button)

        self.remove_button = QPushButton("Eliminar Páginas Seleccionadas y Guardar como...")
        self.remove_button.setFixedSize(350, 40)
        self.remove_button.clicked.connect(self.remove_pages_and_save)
        self.remove_button.setEnabled(False)
        self.remove_button.setIcon(QIcon(resource_path('recursos/trash.png')))
        self.remove_button.setToolTip("Elimina las páginas seleccionadas y guarda el PDF resultante.")
        self.remove_button.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                background-color: #e74c3c;
                color: white;
                border-radius: 5px;
            }
            QPushButton:pressed {
                background-color: #c0392b;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        hlayout.addWidget(self.remove_button)
        hlayout.addStretch()

        main_layout.addLayout(hlayout)
        
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addStretch()
        
        self.secure_delete_checkbox = QCheckBox("Eliminar archivo original de forma segura")
        self.secure_delete_checkbox.setToolTip("Sobrescribe el archivo original para que no pueda ser recuperado.")
        checkbox_layout.addWidget(self.secure_delete_checkbox)
        
        checkbox_layout.addStretch()
        main_layout.addLayout(checkbox_layout)
        
        self.setLayout(main_layout)

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
        self.secure_delete_checkbox.setChecked(False)  
        
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
            
            if self.secure_delete_checkbox.isChecked():
                try:
                    secure_delete_file(self.input_pdf)
                    QMessageBox.information(self, "Éxito", "PDF guardado exitosamente sin las páginas seleccionadas. El archivo original ha sido eliminado de forma segura.")
                    logging.info(f"Archivo original {self.input_pdf} eliminado de forma segura.")
                except Exception as e:
                    QMessageBox.warning(self, "Advertencia", f"El archivo original se guardó, pero no se pudo eliminar de forma segura: {e}")
                    logging.error(f"Fallo en el borrado seguro de {self.input_pdf}: {e}")
            else:
                QMessageBox.information(self, "Éxito", "PDF guardado exitosamente sin las páginas seleccionadas.")
            
            logging.info(f"Páginas eliminadas de {self.input_pdf}. Nuevo archivo guardado en: {out_path}")
            
    def on_error(self, message):
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Error", f"Ocurrió un error: {message}")
        logging.error(f"Error en el eliminador de páginas: {message}")

class PDFCompressorTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.input_pdf = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        self.label = QLabel("Selecciona un archivo PDF para comprimirlo y guardarlo.")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 16px; color: gray; min-height: 300px;")
        layout.addWidget(self.label)
        
        self.open_button = QPushButton("Abrir PDF")
        self.open_button.clicked.connect(self.open_file_dialog)
        self.open_button.setFixedSize(150, 40) 
        self.open_button.setToolTip("Abre el explorador de archivos para seleccionar un documento PDF a comprimir.")
        self.open_button.setIcon(QIcon(resource_path('recursos/pdflogo.png')))
        layout.addWidget(self.open_button, alignment=Qt.AlignCenter)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        combo_layout = QHBoxLayout()
        combo_layout.addStretch()
        
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["Compresion Extrema (menos calidad)", "Compresion Recomendada (Buena Calidad)", "Baja Compresion (Alta Calidad)"])
        self.quality_combo.setFixedSize(280, 30)
        self.quality_combo.setStyleSheet("""
            QComboBox {
                font-size: 12px;
                border: 1px solid gray;
                border-radius: 3px;
                padding: 1px 18px 1px 3px;
                background-color: white;
            }
        """)
        combo_layout.addWidget(self.quality_combo)
        combo_layout.addStretch()
        
        layout.addLayout(combo_layout)

        self.size_layout = QHBoxLayout()
        self.size_layout.addStretch()  

        self.original_size_label = QLabel("Tamaño Original: N/A")
        self.original_size_label.setStyleSheet("font-weight: bold; font-size: 12px;") 
        
        self.separator_label = QLabel(" - ") 
        self.separator_label.setStyleSheet("font-weight: bold; font-size: 12px; color: gray;")

        self.compressed_size_label = QLabel("Tamaño Comprimido: N/A")
        self.compressed_size_label.setStyleSheet("font-weight: bold; font-size: 12px;") 
        
        self.size_layout.addWidget(self.original_size_label)
        self.size_layout.addWidget(self.separator_label)
        self.size_layout.addWidget(self.compressed_size_label)
        self.size_layout.addStretch()  
        layout.addLayout(self.size_layout)

        layout.addStretch()
        
        self.status_label = QLabel("Estado: Esperando archivo...")
        self.status_label.setStyleSheet("font-style: italic; color: #3498db;")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.compress_button = QPushButton("Comprimir PDF y Guardar como...")
        self.compress_button.setFixedSize(300, 40)
        self.compress_button.clicked.connect(self.process_pdf) 
        
        self.compress_button.setEnabled(False)
        
        self.compress_button.setStyleSheet("""
            QPushButton {
                background-color: #bdc3c7;
                color: #808080;
                font-weight: bold;
                border-radius: 5px;
            }
        """)
        self.compress_button.setIcon(QIcon(resource_path('recursos/compress.png')))
        
        layout.addWidget(self.compress_button, alignment=Qt.AlignCenter)

        self.setLayout(layout)
        
    def open_file_dialog(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", "", "Archivos PDF (*.pdf)")
        if file_name:
            self.handle_file(file_name)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and all(url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf") for url in event.mimeData().urls()):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.handle_file(url.toLocalFile())
            break
            
    def handle_file(self, file_path):
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Advertencia", "El archivo no existe.")
            return
        
        if not file_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona un archivo PDF.")
            return

        self.input_pdf = file_path
        self.original_size = os.path.getsize(file_path)
        
        self.status_label.setText(f"Estado: Archivo cargado - {os.path.basename(file_path)}")
        self.status_label.setStyleSheet("font-style: italic; color: #2ecc71;")
        
        self.original_size_label.setText(f"Tamaño Original: {self.format_size(self.original_size)}")
        self.compressed_size_label.setText("Tamaño Comprimido: N/A")
        
        self.compress_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }
        """)
        self.compress_button.setEnabled(True)

    def process_pdf(self):
        if not self.input_pdf:
            QMessageBox.warning(self, "Advertencia", "Por favor, selecciona un archivo PDF primero.")
            return
        
        out_path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF Comprimido", "", "Archivos PDF (*.pdf)")
        
        if not out_path:
            QMessageBox.information(self, "Cancelado", "La operación de guardado ha sido cancelada.")
            return

        gs_path = None
        if hasattr(sys, '_MEIPASS'):
            gs_path = os.path.join(sys._MEIPASS, "recursos", "gswin64c.exe")
        else:
            gs_path = "gswin64c"
        
        try:
            
            creationflags = 0
            if sys.platform == "win32":
                creationflags = subprocess.CREATE_NO_WINDOW
            
            subprocess.run([gs_path, "--version"], check=True, capture_output=True, text=True, creationflags=creationflags)
        except (FileNotFoundError, subprocess.CalledProcessError):
            QMessageBox.critical(self, "Error", 
                "No se encontró Ghostscript. La función de compresión no está disponible.")
            return

        self.status_label.setText("Estado: Comprimiendo...")
        self.status_label.setStyleSheet("font-style: italic; color: #f39c12;")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.compress_button.setEnabled(False)
        self.compress_button.setStyleSheet("""
            QPushButton {
                background-color: #bdc3c7;
                color: #808080;
                font-weight: bold;
                border-radius: 5px;
            }
        """)
        
        quality_map = {
            0: "screen",
            1: "ebook",
            2: "printer"
        }
        quality = quality_map.get(self.quality_combo.currentIndex())
        
        self.worker = Worker("process_pdf", (self.input_pdf, out_path, quality))
        self.worker.progress_update.connect(self.progress_bar.setValue)
        self.worker.sizes_updated.connect(self.update_sizes)
        self.worker.finished_compression.connect(self.on_compression_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        
    def on_compression_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.compress_button.setEnabled(True)
        
        if success:
            self.compress_button.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    font-weight: bold;
                    border-radius: 5px;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
                QPushButton:pressed {
                    background-color: #3e8e41;
                }
            """)
            self.status_label.setText(f"Estado: {message}")
            self.status_label.setStyleSheet("font-style: italic; color: #2ecc71;")
            QMessageBox.information(self, "Éxito", message)
        else:
            self.compress_button.setStyleSheet("""
                QPushButton {
                    background-color: #bdc3c7;
                    color: #808080;
                    font-weight: bold;
                    border-radius: 5px;
                }
            """)
            self.status_label.setText(f"Estado: Error - {message}")
            self.status_label.setStyleSheet("font-style: italic; color: #e74c3c;")
            QMessageBox.critical(self, "Error", f"Ocurrió un error durante la compresión: {message}")
            self.reset_state()
            
    def update_sizes(self, original_mb, compressed_mb):
        self.original_size_label.setText(f"Tamaño Original: {original_mb:.2f} MB")
        self.compressed_size_label.setText(f"Tamaño Comprimido: {compressed_mb:.2f} MB")

    def on_error(self, message):
        self.progress_bar.setVisible(False)
        
        QMessageBox.critical(self, "Error", f"Ocurrió un error: {message}")
        
        self.reset_state()

    def reset_state(self):
        """Reinicia la interfaz a su estado inicial de 'esperando archivo'."""
        self.input_pdf = None
        self.status_label.setText("Estado: Esperando archivo...")
        self.status_label.setStyleSheet("font-style: italic; color: #3498db;")
        self.original_size_label.setText("Tamaño Original: N/A")
        self.compressed_size_label.setText("Tamaño Comprimido: N/A")
        self.compress_button.setEnabled(False)
        self.compress_button.setStyleSheet("""
            QPushButton {
                background-color: #bdc3c7;
                color: #808080;
                font-weight: bold;
                border-radius: 5px;
            }
        """)

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
        self.setWindowTitle("Herramientas de Compresion Judiciales")
        self.setGeometry(100, 100, 1200, 900)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        self.setWindowIcon(QIcon(resource_path('recursos/icon.ico')))
        
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
    
    app.setWindowIcon(QIcon(resource_path('recursos/icon.ico')))
    
    splash_pixmap = QPixmap(resource_path('recursos/logo.png'))
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
