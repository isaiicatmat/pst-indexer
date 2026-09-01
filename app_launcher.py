import sys
import os
import sqlite3
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QProgressBar, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QIcon
import subprocess

class IndexWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(int)

    def run(self):
        try:
            self.progress.emit("Iniciando indexación desde Outlook...\n")
            # Pasar argumento 2 para indexar todas las carpetas sin pedir entrada
            result = subprocess.run(['python', 'indexar_desde_outlook_v2.py', '2'],
                                  capture_output=True, text=True, timeout=600)

            if result.returncode != 0:
                self.progress.emit(f"Advertencia: {result.stderr}\n")

            # Contar correos indexados
            conn = sqlite3.connect('email_index.db')
            c = conn.cursor()
            c.execute('SELECT COUNT(*) FROM emails')
            count = c.fetchone()[0]
            conn.close()

            self.progress.emit(f"Indexacion completada: {count} correos\n")
            self.finished.emit(count)
        except Exception as e:
            self.progress.emit(f"Error: {str(e)}\n")
            self.finished.emit(0)

class Launcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.indexer_thread = None

    def init_ui(self):
        self.setWindowTitle("🔍 Buscador de Correos - Panel de Control")
        self.setGeometry(100, 100, 600, 500)
        self.setStyleSheet(self.get_stylesheet())

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)

        # Título
        title = QLabel("📧 Buscador de Correos Outlook")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Subtítulo
        subtitle = QLabel("Panel de Control")
        subtitle_font = QFont()
        subtitle_font.setPointSize(10)
        subtitle.setFont(subtitle_font)
        subtitle.setStyleSheet("color: #888;")
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        # Sección 1: Estado
        status_label = QLabel("📊 Estado de la Indexación")
        status_label.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(status_label)

        self.status_text = QLabel(self.get_status())
        self.status_text.setStyleSheet("background-color: #f5f5f5; padding: 10px; border-radius: 5px;")
        layout.addWidget(self.status_text)

        layout.addSpacing(10)

        # Sección 2: Acciones
        actions_label = QLabel("⚙️ Acciones")
        actions_label.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(actions_label)

        # Botones en grid
        button_layout = QVBoxLayout()

        btn_index = self.create_button("🔄 Indexar/Actualizar Correos", self.index_emails)
        button_layout.addWidget(btn_index)

        btn_search = self.create_button("🔍 Abrir Buscador", self.open_searcher)
        button_layout.addWidget(btn_search)

        btn_clean = self.create_button("🗑️ Limpiar Índice", self.clean_index)
        button_layout.addWidget(btn_clean)

        layout.addLayout(button_layout)

        layout.addSpacing(10)

        # Sección 3: Log
        log_label = QLabel("📝 Actividad")
        log_label.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(log_label)

        self.log_text = QLabel("")
        self.log_text.setStyleSheet("background-color: #1e1e1e; color: #00ff00; padding: 10px; border-radius: 5px; font-family: Courier;")
        self.log_text.setWordWrap(True)
        layout.addWidget(self.log_text)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(self.get_progress_stylesheet())
        layout.addWidget(self.progress_bar)

        layout.addStretch()

        self.update_status()

    def create_button(self, text, callback):
        btn = QPushButton(text)
        btn.setFont(QFont("Arial", 10))
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(45)
        btn.setStyleSheet(self.get_button_stylesheet())
        btn.clicked.connect(callback)
        return btn

    def get_status(self):
        try:
            conn = sqlite3.connect('email_index.db')
            c = conn.cursor()
            c.execute('SELECT COUNT(*) FROM emails')
            count = c.fetchone()[0]

            c.execute('SELECT COUNT(DISTINCT pst_file) FROM emails')
            files = c.fetchone()[0]

            conn.close()
            return f"✓ {count:,} correos indexados | {files} fuente(s)"
        except:
            return "⚠️ Base de datos no inicializada"

    def update_status(self):
        self.status_text.setText(self.get_status())

    def index_emails(self):
        self.log_text.setText("Iniciando indexación...\n")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(50)

        self.indexer_thread = IndexWorker()
        self.indexer_thread.progress.connect(self.update_log)
        self.indexer_thread.finished.connect(self.on_index_finished)
        self.indexer_thread.start()

    def update_log(self, message):
        self.log_text.setText(self.log_text.text() + message)

    def on_index_finished(self, count):
        self.progress_bar.setVisible(False)
        self.update_status()
        if count > 0:
            QMessageBox.information(self, "✓ Éxito", f"Se indexaron {count} correos correctamente")
        else:
            QMessageBox.warning(self, "⚠️ Advertencia", "No se indexaron correos")

    def open_searcher(self):
        try:
            subprocess.Popen(['python', 'email_searcher_modern.py'])
            self.log_text.setText("✓ Buscador abierto\n")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir el buscador: {str(e)}")

    def clean_index(self):
        reply = QMessageBox.question(self, "Confirmación",
                                     "¿Eliminar toda la indexación?\n(Los correos no se borrarán, solo el índice)",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                if os.path.exists('email_index.db'):
                    os.remove('email_index.db')
                self.log_text.setText("✓ Índice eliminado\n")
                self.update_status()
                QMessageBox.information(self, "✓ Éxito", "Índice eliminado. Ejecuta 'Indexar' para recrearlo")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo eliminar: {str(e)}")

    def get_stylesheet(self):
        return """
        QMainWindow {
            background-color: #ffffff;
        }
        QLabel {
            color: #333333;
        }
        """

    def get_button_stylesheet(self):
        return """
        QPushButton {
            background-color: #007AFF;
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: bold;
            padding: 10px;
        }
        QPushButton:hover {
            background-color: #0051D5;
        }
        QPushButton:pressed {
            background-color: #003DA1;
        }
        """

    def get_progress_stylesheet(self):
        return """
        QProgressBar {
            border: 2px solid #ddd;
            border-radius: 5px;
            text-align: center;
            height: 30px;
        }
        QProgressBar::chunk {
            background-color: #007AFF;
        }
        """

if __name__ == '__main__':
    app = QApplication(sys.argv)
    launcher = Launcher()
    launcher.show()
    sys.exit(app.exec_())
