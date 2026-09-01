import sqlite3
import sys
from datetime import datetime
import logging

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
    import win32com.client
except ImportError:
    logger.error("Se requiere pywin32. Instala con: pip install pywin32")
    sys.exit(1)


class OutlookIndexer:
    def __init__(self, db_path='email_index.db'):
        self.db_path = db_path
        self.init_database()
        self.outlook = None
        self.namespace = None

    def init_database(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

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

        c.execute('CREATE INDEX IF NOT EXISTS idx_sender ON emails(sender)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_subject ON emails(subject)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_date ON emails(date)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_body ON emails(body)')

        conn.commit()
        conn.close()

    def connect_outlook(self):
        try:
            self.outlook = win32com.client.GetObject(None, "Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            logger.info("✓ Conectado a Outlook")
            return True
        except Exception as e:
            logger.error(f"Error conectando a Outlook: {e}")
            return False

    def extract_body(self, item):
        """Extraer body con mejor manejo de formatos"""
        body = ""

        # Intentar Plain Text Body primero
        try:
            if hasattr(item, 'Body') and item.Body:
                body = str(item.Body).strip()
                if body:
                    return body
        except:
            pass

        # Intentar HTML Body
        try:
            if hasattr(item, 'HTMLBody') and item.HTMLBody:
                html = str(item.HTMLBody)
                # Limpiar tags HTML básicos
                import re
                text = re.sub('<[^<]+?>', '', html)
                body = text.strip()
                if body:
                    return body[:5000]  # Limitar a 5000 caracteres
        except:
            pass

        # Intentar PropAccessor para propiedades especiales
        try:
            pa = item.PropertyAccessor
            body = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1000001F")
            if body:
                return str(body).strip()
        except:
            pass

        return body

    def indexar_carpeta(self, folder_name="Bandeja de entrada"):
        if not self.connect_outlook():
            return 0

        try:
            inbox = self.namespace.GetDefaultFolder(6)

            if folder_name.lower() != "bandeja de entrada":
                try:
                    for folder in inbox.Parent.Folders:
                        if folder.Name.lower() == folder_name.lower():
                            inbox = folder
                            break
                except:
                    pass

            logger.info(f"Indexando: {inbox.Name}")
            logger.info(f"Total de correos: {inbox.Items.Count}")

            total = inbox.Items.Count
            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()

            for idx, item in enumerate(inbox.Items):
                try:
                    if idx % 50 == 0:
                        logger.info(f"Procesados: {idx}/{total}")

                    sender = str(item.SenderName) if hasattr(item, 'SenderName') else ""
                    recipient = str(item.To) if hasattr(item, 'To') else ""
                    subject = str(item.Subject) if hasattr(item, 'Subject') else ""
                    body = self.extract_body(item)
                    date = str(item.ReceivedTime) if hasattr(item, 'ReceivedTime') else ""

                    c.execute('''
                        INSERT OR IGNORE INTO emails
                        (sender, recipient, subject, body, date, pst_file, indexed_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (sender, recipient, subject, body, date,
                          f"Outlook:{inbox.Name}", datetime.now().isoformat()))

                except Exception as e:
                    logger.debug(f"Error procesando correo: {e}")
                    continue

            conn.commit()
            conn.close()

            logger.info(f"✓ Indexación completada: {total} correos")
            return total

        except Exception as e:
            logger.error(f"Error indexando carpeta: {e}")
            return 0

    def indexar_todas_carpetas(self):
        if not self.connect_outlook():
            return 0

        try:
            inbox = self.namespace.GetDefaultFolder(6)
            total_indexados = 0

            for folder in inbox.Parent.Folders:
                logger.info(f"\nIndexando carpeta: {folder.Name}")
                count = self.indexar_carpeta_obj(folder)
                total_indexados += count

            logger.info(f"\n✓ Total indexado: {total_indexados} correos")
            return total_indexados

        except Exception as e:
            logger.error(f"Error indexando carpetas: {e}")
            return 0

    def indexar_carpeta_obj(self, folder):
        try:
            total = folder.Items.Count
            logger.info(f"  Correos: {total}")

            conn = sqlite3.connect(self.db_path)
            c = conn.cursor()

            for idx, item in enumerate(folder.Items):
                try:
                    if idx % 50 == 0:
                        logger.info(f"    {idx}/{total}")

                    sender = str(item.SenderName) if hasattr(item, 'SenderName') else ""
                    recipient = str(item.To) if hasattr(item, 'To') else ""
                    subject = str(item.Subject) if hasattr(item, 'Subject') else ""
                    body = self.extract_body(item)
                    date = str(item.ReceivedTime) if hasattr(item, 'ReceivedTime') else ""

                    c.execute('''
                        INSERT OR IGNORE INTO emails
                        (sender, recipient, subject, body, date, pst_file, indexed_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (sender, recipient, subject, body, date,
                          f"Outlook:{folder.Name}", datetime.now().isoformat()))

                except Exception as e:
                    logger.debug(f"Error: {e}")
                    continue

            conn.commit()
            conn.close()
            return total

        except Exception as e:
            logger.error(f"Error indexando {folder.Name}: {e}")
            return 0


def main():
    print("\n" + "="*50)
    print("  Indexador de Outlook v2")
    print("="*50 + "\n")

    indexer = OutlookIndexer()

    print("Opciones:")
    print("1. Indexar solo Bandeja de entrada")
    print("2. Indexar TODAS las carpetas")
    print()

    opcion = input("Elige opción (1 o 2): ").strip()

    if opcion == "1":
        count = indexer.indexar_carpeta("Bandeja de entrada")
    elif opcion == "2":
        count = indexer.indexar_todas_carpetas()
    else:
        print("Opción inválida")
        return

    print(f"\n✓ Se indexaron {count} correos")
    print("Ahora puedes usar: python email_searcher_modern.py\n")

    input("Presiona Enter para salir...")


if __name__ == '__main__':
    main()
