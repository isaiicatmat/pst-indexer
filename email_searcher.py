import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
import sqlite3
import os
import json
import threading
from datetime import datetime
from pathlib import Path
import email
from email import policy
from email.parser import BytesParser
import logging

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_searcher.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

try:
    import extract_msg
except ImportError:
    extract_msg = None


class EmailIndexer:
    def __init__(self, db_path='email_index.db'):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Inicializar la base de datos SQLite"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        # Crear tabla si no existe
        c.execute('''
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sender TEXT,
                recipient TEXT,
                subject TEXT,
                body TEXT,
                date TEXT,
                pst_file TEXT,
                indexed_date TEXT,
                UNIQUE(sender, subject, date, pst_file)
            )
        ''')

        # Crear índices para búsquedas rápidas
        c.execute('CREATE INDEX IF NOT EXISTS idx_sender ON emails(sender)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_subject ON emails(subject)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_date ON emails(date)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_body ON emails(body)')

        conn.commit()
        conn.close()

    def index_pst_file(self, pst_path, progress_callback=None):
        """Indexar un archivo PST"""
        logger.info(f"Iniciando indexación de: {pst_path}")

        # Intentar con extract_msg
        if extract_msg:
            try:
                return self._index_with_extract_msg(pst_path, progress_callback)
            except Exception as e:
                logger.warning(f"Error con extract_msg: {e}")

        # Fallback a lectura manual
        try:
            return self._index_manual(pst_path, progress_callback)
        except Exception as e:
            logger.error(f"Error indexando {pst_path}: {e}")
            return 0

    def _index_with_extract_msg(self, pst_path, progress_callback):
        """Indexar archivos .msg dentro de PST"""
        try:
            import subprocess
            import tempfile

            # Intentar extraer PST a MSG usando herramienta externa
            temp_dir = tempfile.mkdtemp()

            # Nota: Esto requiere readpst de libpst instalado
            # readpst -e -o temp_dir archivo.pst
            try:
                subprocess.run(['readpst', '-e', '-o', temp_dir, pst_path],
                             capture_output=True, check=True)
            except (FileNotFoundError, subprocess.CalledProcessError) as e:
                logger.warning(f"readpst no disponible: {e}")
                return 0

            # Indexar archivos .msg extraídos
            msg_files = list(Path(temp_dir).rglob('*.msg'))
            total = len(msg_files)

            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()

            for idx, msg_file in enumerate(msg_files):
                try:
                    msg = extract_msg.Message(str(msg_file))

                    c.execute('''
                        INSERT OR IGNORE INTO emails
                        (sender, recipient, subject, body, date, pst_file, indexed_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        msg.sender or '',
                        msg.to or '',
                        msg.subject or '',
                        msg.body or '',
                        str(msg.date) if msg.date else '',
                        pst_path,
                        datetime.now().isoformat()
                    ))

                    if progress_callback and idx % 10 == 0:
                        progress_callback(idx, total)
                except Exception as e:
                    logger.error(f"Error procesando {msg_file}: {e}")
                    continue

            conn.commit()
            conn.close()
            logger.info(f"Indexación completada: {total} correos")
            return total

        except Exception as e:
            logger.error(f"Error con extract_msg: {e}")
            return 0

    def _index_manual(self, pst_path, progress_callback):
        """Intentar lectura manual de PST como contenedor OLE"""
        try:
            import olefile

            ole = olefile.OleFileIO(pst_path)

            if not ole.exists('Outlook Message Database'):
                logger.error("No es un archivo PST válido")
                return 0

            logger.info("Lectura manual de PST iniciada")
            return 0  # La lectura manual de PST es muy compleja

        except ImportError:
            logger.error("Se requiere instalar olefile o libpst-python")
            return 0
        except Exception as e:
            logger.error(f"Error en lectura manual: {e}")
            return 0

    def search(self, sender='', subject='', content='', date_from='', date_to=''):
        """Buscar en la base de datos"""
        # Validar que al menos un criterio está presente
        if not any([sender, subject, content, date_from, date_to]):
            logger.warning("Búsqueda sin criterios rechazada")
            return []

        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        query = 'SELECT sender, subject, body, date, pst_file FROM emails WHERE 1=1'
        params = []

        if sender:
            query += ' AND sender LIKE ?'
            params.append(f'%{sender}%')

        if subject:
            query += ' AND subject LIKE ?'
            params.append(f'%{subject}%')

        if content:
            query += ' AND body LIKE ?'
            params.append(f'%{content}%')

        if date_from:
            query += ' AND date >= ?'
            params.append(date_from)

        if date_to:
            query += ' AND date <= ?'
            params.append(date_to)

        query += ' ORDER BY date DESC LIMIT 1000'

        c.execute(query, params)
        results = c.fetchall()
        conn.close()

        return results

    def get_index_stats(self):
        """Obtener estadísticas del índice"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('SELECT COUNT(*) FROM emails')
        total = c.fetchone()[0]

        c.execute('SELECT COUNT(DISTINCT pst_file) FROM emails')
        files = c.fetchone()[0]

        c.execute('SELECT MIN(date), MAX(date) FROM emails')
        dates = c.fetchone()

        conn.close()

        return {
            'total_emails': total,
            'total_files': files,
            'date_range': dates
        }


class EmailSearcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Correos Outlook")
        self.root.geometry("1000x700")

        self.indexer = EmailIndexer()

        # Estilo
        self.root.configure(bg='#f0f0f0')
        self.setup_styles()

        # Main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # Header
        self.create_header(main_frame)

        # Search section
        self.create_search_section(main_frame)

        # Results section
        self.create_results_section(main_frame)

        # Status bar
        self.create_status_bar(main_frame)

        self.update_stats()

    def setup_styles(self):
        """Configurar estilos"""
        style = ttk.Style()
        style.theme_use('clam')

        # Colores personalizados
        style.configure('Title.TLabel', font=('Helvetica', 16, 'bold'), background='#f0f0f0')
        style.configure('Header.TFrame', background='#f0f0f0')
        style.configure('Header.TLabel', background='#f0f0f0')

    def create_header(self, parent):
        """Crear encabezado"""
        header_frame = ttk.Frame(parent, padding="10")
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        header_frame.columnconfigure(1, weight=1)

        title_label = ttk.Label(header_frame, text="🔍 Buscador de Correos",
                               font=('Helvetica', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)

        # Botones de acción
        button_frame = ttk.Frame(header_frame)
        button_frame.grid(row=0, column=1, sticky=tk.E)

        ttk.Button(button_frame, text="📁 Seleccionar carpeta PST",
                  command=self.select_pst_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🔄 Actualizar índice",
                  command=self.refresh_index).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="ℹ️ Estadísticas",
                  command=self.show_stats).pack(side=tk.LEFT, padx=5)

    def create_search_section(self, parent):
        """Crear sección de búsqueda"""
        search_frame = ttk.LabelFrame(parent, text="Filtros de búsqueda", padding="10")
        search_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        search_frame.columnconfigure(1, weight=1)
        search_frame.columnconfigure(3, weight=1)
        search_frame.columnconfigure(5, weight=1)

        # Remitente
        ttk.Label(search_frame, text="Remitente:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.sender_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.sender_var, width=25).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        # Asunto
        ttk.Label(search_frame, text="Asunto:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.subject_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.subject_var, width=25).grid(row=0, column=3, sticky=(tk.W, tk.E), padx=5)

        # Contenido
        ttk.Label(search_frame, text="Contenido:").grid(row=0, column=4, sticky=tk.W, padx=5)
        self.content_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.content_var, width=25).grid(row=0, column=5, sticky=(tk.W, tk.E), padx=5)

        # Fecha
        ttk.Label(search_frame, text="Desde:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.date_from_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.date_from_var, width=25).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)

        ttk.Label(search_frame, text="Hasta:").grid(row=1, column=2, sticky=tk.W, padx=5)
        self.date_to_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.date_to_var, width=25).grid(row=1, column=3, sticky=(tk.W, tk.E), padx=5)

        # Botones
        button_frame = ttk.Frame(search_frame)
        button_frame.grid(row=1, column=4, columnspan=2, sticky=tk.E, padx=5)

        ttk.Button(button_frame, text="🔍 Buscar", command=self.search).pack(side=tk.LEFT, padx=3)
        ttk.Button(button_frame, text="✕ Limpiar", command=self.clear_search).pack(side=tk.LEFT, padx=3)

    def create_results_section(self, parent):
        """Crear sección de resultados"""
        results_frame = ttk.LabelFrame(parent, text="Resultados", padding="10")
        results_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

        # Treeview con scrollbars
        tree_frame = ttk.Frame(results_frame)
        tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))

        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.tree = ttk.Treeview(tree_frame,
                                columns=('sender', 'subject', 'date', 'file'),
                                height=15,
                                yscrollcommand=vsb.set,
                                xscrollcommand=hsb.set)

        vsb.configure(command=self.tree.yview)
        hsb.configure(command=self.tree.xview)

        self.tree.column('#0', width=0, stretch=tk.NO)
        self.tree.column('sender', width=200, heading='Remitente')
        self.tree.column('subject', width=300, heading='Asunto')
        self.tree.column('date', width=150, heading='Fecha')
        self.tree.column('file', width=150, heading='Archivo')

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Bind para mostrar detalles
        self.tree.bind('<Double-1>', self.show_email_details)

    def create_status_bar(self, parent):
        """Crear barra de estado"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))

        self.status_var = tk.StringVar(value="Listo")
        ttk.Label(status_frame, textvariable=self.status_var, relief=tk.SUNKEN).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.stats_var = tk.StringVar(value="")
        ttk.Label(status_frame, textvariable=self.stats_var, relief=tk.SUNKEN).pack(side=tk.RIGHT, padx=5)

    def select_pst_folder(self):
        """Seleccionar carpeta con archivos PST"""
        folder = filedialog.askdirectory(title="Selecciona la carpeta con archivos PST")
        if folder:
            self.index_folder(folder)

    def index_folder(self, folder_path):
        """Indexar todos los PST en una carpeta"""
        pst_files = list(Path(folder_path).rglob('*.pst'))

        if not pst_files:
            messagebox.showwarning("Advertencia", "No se encontraron archivos .pst en la carpeta")
            return

        # Ejecutar en thread para no bloquear GUI
        thread = threading.Thread(target=self._index_thread, args=(pst_files,))
        thread.start()

    def _index_thread(self, pst_files):
        """Thread para indexación"""
        try:
            total_emails = 0
            for idx, pst_file in enumerate(pst_files):
                self.status_var.set(f"Indexando {idx+1}/{len(pst_files)}: {pst_file.name}")
                self.root.update()

                try:
                    count = self.indexer.index_pst_file(str(pst_file))
                    total_emails += count
                except Exception as e:
                    logger.error(f"Error indexando {pst_file}: {e}")
                    messagebox.showerror("Error", f"Error indexando {pst_file.name}: {str(e)}")

            self.status_var.set(f"✓ Indexación completada: {total_emails} correos")
            self.update_stats()
            messagebox.showinfo("Éxito", f"Se indexaron {total_emails} correos de {len(pst_files)} archivo(s)")

        except Exception as e:
            self.status_var.set("Error en indexación")
            messagebox.showerror("Error", f"Error durante indexación: {str(e)}")

    def search(self):
        """Ejecutar búsqueda"""
        sender = self.sender_var.get()
        subject = self.subject_var.get()
        content = self.content_var.get()
        date_from = self.date_from_var.get()
        date_to = self.date_to_var.get()

        if not any([sender, subject, content, date_from, date_to]):
            messagebox.showwarning("Advertencia", "Ingresa al menos un criterio de búsqueda")
            return

        # Ejecutar en thread
        thread = threading.Thread(target=self._search_thread,
                                 args=(sender, subject, content, date_from, date_to))
        thread.start()

    def _search_thread(self, sender, subject, content, date_from, date_to):
        """Thread para búsqueda"""
        self.status_var.set("Buscando...")
        self.root.update()

        try:
            results = self.indexer.search(sender, subject, content, date_from, date_to)

            # Limpiar tree
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Agregar resultados
            for idx, (sender_val, subject_val, body, date_val, file) in enumerate(results):
                # Truncar campos para visualización
                sender_display = (sender_val[:50] + '...') if len(sender_val) > 50 else sender_val
                subject_display = (subject_val[:50] + '...') if len(subject_val) > 50 else subject_val
                file_display = os.path.basename(file)

                self.tree.insert('', 'end', iid=idx,
                               values=(sender_display, subject_display, date_val, file_display),
                               tags=('result',))

            self.status_var.set(f"✓ {len(results)} resultado(s) encontrado(s)")

        except Exception as e:
            logger.error(f"Error en búsqueda: {e}")
            messagebox.showerror("Error", f"Error en búsqueda: {str(e)}")
            self.status_var.set("Error en búsqueda")

    def show_email_details(self, event):
        """Mostrar detalles del correo seleccionado"""
        selection = self.tree.selection()
        if not selection:
            return

        item_id = selection[0]
        values = self.tree.item(item_id)['values']

        if len(values) >= 3:
            sender, subject, date = values[0], values[1], values[2]

            # Crear ventana de detalles
            detail_window = tk.Toplevel(self.root)
            detail_window.title(f"Detalles - {subject}")
            detail_window.geometry("700x400")

            # Frame principal
            main_frame = ttk.Frame(detail_window, padding="10")
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Información del correo
            info_text = f"""
De: {sender}
Fecha: {date}
Asunto: {subject}
"""

            ttk.Label(main_frame, text=info_text, justify=tk.LEFT).pack(fill=tk.X, pady=10)

            ttk.Label(main_frame, text="Contenido:").pack(fill=tk.X)

            # Texto del cuerpo
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            text_widget = tk.Text(text_frame, yscrollcommand=scrollbar.set, height=15, width=80)
            text_widget.pack(fill=tk.BOTH, expand=True)
            scrollbar.config(command=text_widget.yview)

            # Obtener cuerpo del correo desde BD
            try:
                conn = sqlite3.connect(self.indexer.db_path)
                c = conn.cursor()
                c.execute('SELECT body FROM emails WHERE sender=? AND subject=? AND date=?',
                         (sender, subject, date))
                result = c.fetchone()
                conn.close()

                if result and result[0]:
                    text_widget.insert(1.0, result[0])
                else:
                    text_widget.insert(1.0, "[Sin contenido]")
            except Exception as e:
                text_widget.insert(1.0, f"[Error al cargar: {str(e)}]")

            text_widget.config(state=tk.DISABLED)

    def clear_search(self):
        """Limpiar búsqueda"""
        self.sender_var.set('')
        self.subject_var.set('')
        self.content_var.set('')
        self.date_from_var.set('')
        self.date_to_var.set('')

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.status_var.set("Listo")

    def refresh_index(self):
        """Actualizar índice"""
        self.select_pst_folder()

    def show_stats(self):
        """Mostrar estadísticas"""
        stats = self.indexer.get_index_stats()

        msg = f"""
Estadísticas del índice:

Total de correos: {stats['total_emails']:,}
Archivos PST indexados: {stats['total_files']}

Rango de fechas:
  Desde: {stats['date_range'][0] or 'N/A'}
  Hasta: {stats['date_range'][1] or 'N/A'}
"""

        messagebox.showinfo("Estadísticas", msg)

    def update_stats(self):
        """Actualizar estadísticas en barra de estado"""
        try:
            stats = self.indexer.get_index_stats()
            self.stats_var.set(f"📊 {stats['total_emails']:,} correos | {stats['total_files']} archivo(s)")
        except Exception as e:
            logger.error(f"Error actualizando estadísticas: {e}")


def main():
    root = tk.Tk()
    app = EmailSearcherGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
