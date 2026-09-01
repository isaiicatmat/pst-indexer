import sys
import sqlite3
import logging
from datetime import datetime
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QLabel,
                             QMessageBox, QDialog, QTextEdit, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class SearchWorker(QThread):
    results_ready = pyqtSignal(list)
    searching = pyqtSignal(bool)

    def __init__(self, sender, subject, content, date_from, date_to):
        super().__init__()
        self.sender = sender
        self.subject = subject
        self.content = content
        self.date_from = date_from
        self.date_to = date_to

    def run(self):
        self.searching.emit(True)
        try:
            conn = sqlite3.connect('email_index.db')
            c = conn.cursor()

            query = 'SELECT sender, subject, body, date, pst_file FROM emails WHERE 1=1'
            params = []

            if self.sender:
                query += ' AND sender LIKE ?'
                params.append(f'%{self.sender}%')
            if self.subject:
                query += ' AND subject LIKE ?'
                params.append(f'%{self.subject}%')
            if self.content:
                query += ' AND body LIKE ?'
                params.append(f'%{self.content}%')
            if self.date_from:
                query += ' AND date >= ?'
                params.append(self.date_from)
            if self.date_to:
                query += ' AND date <= ?'
                params.append(self.date_to)

            query += ' ORDER BY date DESC LIMIT 1000'
            c.execute(query, params)
            results = c.fetchall()
            conn.close()

            self.results_ready.emit(results)
        except Exception as e:
            logger.error(f"Error en búsqueda: {e}")
            self.results_ready.emit([])
        finally:
            self.searching.emit(False)


class EmailSearcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.search_worker = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("🔍 Buscador de Correos")
        self.setGeometry(100, 100, 1200, 700)
        self.setStyleSheet(self.get_stylesheet())

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Título
        title = QLabel("📧 Buscador de Correos")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        main_layout.addWidget(title)

        # Área de búsqueda
        search_layout = QVBoxLayout()

        # Fila 1: Remitente y Asunto
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Remitente:"))
        self.sender_input = QLineEdit()
        self.sender_input.setPlaceholderText("Busca por remitente...")
        self.sender_input.returnPressed.connect(self.search)
        row1.addWidget(self.sender_input)

        row1.addWidget(QLabel("Asunto:"))
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("Busca por asunto...")
        self.subject_input.returnPressed.connect(self.search)
        row1.addWidget(self.subject_input)

        search_layout.addLayout(row1)

        # Fila 2: Contenido y Fecha
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Contenido:"))
        self.content_input = QLineEdit()
        self.content_input.setPlaceholderText("Busca en el contenido...")
        self.content_input.returnPressed.connect(self.search)
        row2.addWidget(self.content_input)

        row2.addWidget(QLabel("Desde:"))
        self.date_from_input = QLineEdit()
        self.date_from_input.setPlaceholderText("YYYY-MM-DD")
        self.date_from_input.returnPressed.connect(self.search)
        row2.addWidget(self.date_from_input)

        row2.addWidget(QLabel("Hasta:"))
        self.date_to_input = QLineEdit()
        self.date_to_input.setPlaceholderText("YYYY-MM-DD")
        self.date_to_input.returnPressed.connect(self.search)
        row2.addWidget(self.date_to_input)

        search_layout.addLayout(row2)

        # Botones
        button_layout = QHBoxLayout()
        self.search_btn = QPushButton("🔍 Buscar (Enter)")
        self.search_btn.setFont(QFont("Arial", 10, QFont.Bold))
        self.search_btn.clicked.connect(self.search)
        self.search_btn.setMinimumHeight(40)
        button_layout.addWidget(self.search_btn)

        self.clear_btn = QPushButton("✕ Limpiar")
        self.clear_btn.clicked.connect(self.clear_search)
        self.clear_btn.setMinimumHeight(40)
        button_layout.addWidget(self.clear_btn)

        search_layout.addLayout(button_layout)

        main_layout.addLayout(search_layout)

        # Estado
        self.status_label = QLabel("Listo")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        main_layout.addWidget(self.status_label)

        # Tabla de resultados
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["Remitente", "Asunto", "Fecha", "Archivo"])
        self.results_table.horizontalHeader().setStretchLastSection(False)
        self.results_table.setColumnWidth(0, 250)
        self.results_table.setColumnWidth(1, 400)
        self.results_table.setColumnWidth(2, 150)
        self.results_table.setColumnWidth(3, 150)
        self.results_table.itemDoubleClicked.connect(self.show_email_details)
        main_layout.addWidget(self.results_table)

        self.load_recent_emails()

    def search(self):
        if not any([self.sender_input.text(), self.subject_input.text(),
                   self.content_input.text(), self.date_from_input.text(),
                   self.date_to_input.text()]):
            QMessageBox.warning(self, "Advertencia", "Ingresa al menos un criterio de búsqueda")
            return

        self.status_label.setText("🔍 Buscando...")
        self.search_btn.setEnabled(False)

        self.search_worker = SearchWorker(
            self.sender_input.text(),
            self.subject_input.text(),
            self.content_input.text(),
            self.date_from_input.text(),
            self.date_to_input.text()
        )
        self.search_worker.results_ready.connect(self.display_results)
        self.search_worker.searching.connect(self.on_search_state_changed)
        self.search_worker.start()

    def on_search_state_changed(self, searching):
        self.search_btn.setEnabled(not searching)

    def display_results(self, results):
        self.results_table.setRowCount(0)

        for row_idx, (sender, subject, body, date, file) in enumerate(results):
            self.results_table.insertRow(row_idx)

            sender_display = (sender[:60] + "...") if len(sender) > 60 else sender
            subject_display = (subject[:80] + "...") if len(subject) > 80 else subject

            self.results_table.setItem(row_idx, 0, QTableWidgetItem(sender_display))
            self.results_table.setItem(row_idx, 1, QTableWidgetItem(subject_display))
            self.results_table.setItem(row_idx, 2, QTableWidgetItem(date))
            self.results_table.setItem(row_idx, 3, QTableWidgetItem(Path(file).name))

        self.status_label.setText(f"✓ {len(results)} resultado(s) encontrado(s)")

    def show_email_details(self, item):
        row = item.row()
        sender = self.results_table.item(row, 0).text()
        subject = self.results_table.item(row, 1).text()
        date = self.results_table.item(row, 2).text()

        # Obtener body
        try:
            conn = sqlite3.connect('email_index.db')
            c = conn.cursor()
            c.execute('SELECT body FROM emails WHERE sender LIKE ? AND subject LIKE ? AND date LIKE ? LIMIT 1',
                     (f"%{sender}%", f"%{subject}%", f"%{date}%"))
            result = c.fetchone()
            conn.close()

            body = result[0] if result and result[0] else "[Sin contenido]"
        except:
            body = "[Error al cargar]"

        # Ventana de detalles
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Detalles - {subject}")
        dialog.setGeometry(200, 200, 800, 600)
        layout = QVBoxLayout()

        info = QLabel(f"De: {sender}\nFecha: {date}\nAsunto: {subject}")
        info.setFont(QFont("Arial", 9))
        layout.addWidget(info)

        body_text = QTextEdit()
        body_text.setText(body)
        body_text.setReadOnly(True)
        layout.addWidget(body_text)

        dialog.setLayout(layout)
        dialog.exec_()

    def load_recent_emails(self):
        try:
            conn = sqlite3.connect('email_index.db')
            c = conn.cursor()
            c.execute('SELECT sender, subject, date, pst_file FROM emails ORDER BY date DESC LIMIT 50')
            results = c.fetchall()
            conn.close()

            self.results_table.setRowCount(0)
            for row_idx, (sender, subject, date, file) in enumerate(results):
                self.results_table.insertRow(row_idx)
                sender_display = (sender[:60] + "...") if len(sender) > 60 else sender
                subject_display = (subject[:80] + "...") if len(subject) > 80 else subject
                self.results_table.setItem(row_idx, 0, QTableWidgetItem(sender_display))
                self.results_table.setItem(row_idx, 1, QTableWidgetItem(subject_display))
                self.results_table.setItem(row_idx, 2, QTableWidgetItem(date))
                self.results_table.setItem(row_idx, 3, QTableWidgetItem(Path(file).name))

            self.status_label.setText(f"✓ {len(results)} correos recientes")
        except:
            self.status_label.setText("⚠️ Base de datos no inicializada")

    def clear_search(self):
        self.sender_input.clear()
        self.subject_input.clear()
        self.content_input.clear()
        self.date_from_input.clear()
        self.date_to_input.clear()
        self.load_recent_emails()

    def get_stylesheet(self):
        return """
        QMainWindow {
            background-color: #f8f9fa;
        }
        QLineEdit {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 10pt;
        }
        QLineEdit:focus {
            border: 2px solid #007AFF;
        }
        QPushButton {
            background-color: #007AFF;
            color: white;
            border: none;
            border-radius: 5px;
            font-weight: bold;
            padding: 8px;
        }
        QPushButton:hover {
            background-color: #0051D5;
        }
        QPushButton:pressed {
            background-color: #003DA1;
        }
        QTableWidget {
            border: 1px solid #ddd;
            border-radius: 5px;
            gridline-color: #eee;
        }
        QHeaderView::section {
            background-color: #f0f0f0;
            padding: 5px;
            border: none;
        }
        """


if __name__ == '__main__':
    app = QApplication(sys.argv)
    searcher = EmailSearcher()
    searcher.show()
    sys.exit(app.exec_())
