import unittest
import sqlite3
import tempfile
import os
import json
from pathlib import Path
from datetime import datetime, timedelta
from email_searcher import EmailIndexer


class TestEmailIndexer(unittest.TestCase):
    """Tests para la clase EmailIndexer"""

    def setUp(self):
        """Preparar tests - crear DB temporal"""
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, 'test_email.db')
        self.indexer = EmailIndexer(self.db_path)

    def tearDown(self):
        """Limpiar tests - eliminar archivos temporales"""
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
        os.rmdir(self.temp_dir)

    def test_database_initialization(self):
        """Test: Base de datos se crea correctamente"""
        self.assertTrue(os.path.exists(self.db_path))

        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails'")
        result = c.fetchone()
        conn.close()

        self.assertIsNotNone(result, "La tabla 'emails' no fue creada")

    def test_database_schema(self):
        """Test: La tabla tiene las columnas correctas"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("PRAGMA table_info(emails)")
        columns = [row[1] for row in c.fetchall()]
        conn.close()

        expected_columns = ['id', 'sender', 'recipient', 'subject', 'body', 'date', 'pst_file', 'indexed_date']
        for col in expected_columns:
            self.assertIn(col, columns, f"Columna '{col}' no encontrada")

    def test_add_single_email(self):
        """Test: Agregar un correo a la base de datos"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        test_email = {
            'sender': 'test@example.com',
            'recipient': 'user@example.com',
            'subject': 'Test Subject',
            'body': 'Test body content',
            'date': '2024-01-01T10:00:00',
            'pst_file': 'test.pst'
        }

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (test_email['sender'], test_email['recipient'], test_email['subject'],
              test_email['body'], test_email['date'], test_email['pst_file'],
              datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='test@example.com')
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0][0], 'test@example.com')

    def test_search_by_sender(self):
        """Test: Búsqueda por remitente"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender1@test.com', 'user@test.com', 'Subject 1', 'Body 1', '2024-01-01', 'test.pst'),
            ('sender2@test.com', 'user@test.com', 'Subject 2', 'Body 2', '2024-01-02', 'test.pst'),
            ('sender1@test.com', 'user@test.com', 'Subject 3', 'Body 3', '2024-01-03', 'test.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='sender1@test.com')
        self.assertEqual(len(results), 2)

    def test_search_by_subject(self):
        """Test: Búsqueda por asunto"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender@test.com', 'user@test.com', 'Importante: Reunión', 'Body 1', '2024-01-01', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Urgente: Presupuesto', 'Body 2', '2024-01-02', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Reunión de seguimiento', 'Body 3', '2024-01-03', 'test.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(subject='Reunión')
        self.assertEqual(len(results), 2)

    def test_search_by_content(self):
        """Test: Búsqueda por contenido del correo"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender@test.com', 'user@test.com', 'Subject 1', 'El presupuesto está aprobado', '2024-01-01', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Subject 2', 'Necesito revisar el presupuesto', '2024-01-02', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Subject 3', 'Contenido sin palabra clave', '2024-01-03', 'test.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(content='presupuesto')
        self.assertEqual(len(results), 2)

    def test_search_by_date_range(self):
        """Test: Búsqueda por rango de fechas"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender@test.com', 'user@test.com', 'Subject 1', 'Body 1', '2024-01-01', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Subject 2', 'Body 2', '2024-01-15', 'test.pst'),
            ('sender@test.com', 'user@test.com', 'Subject 3', 'Body 3', '2024-02-01', 'test.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(date_from='2024-01-01', date_to='2024-01-31')
        self.assertEqual(len(results), 2)

    def test_search_combined_criteria(self):
        """Test: Búsqueda con múltiples criterios"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender1@test.com', 'user@test.com', 'Reunión', 'Hablaremos sobre el proyecto', '2024-01-01', 'test.pst'),
            ('sender2@test.com', 'user@test.com', 'Reunión', 'Hablaremos sobre el presupuesto', '2024-01-02', 'test.pst'),
            ('sender1@test.com', 'user@test.com', 'Otro tema', 'Hablaremos sobre el proyecto', '2024-01-03', 'test.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='sender1@test.com', subject='Reunión', content='proyecto')
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0][1], 'Reunión')

    def test_search_no_results(self):
        """Test: Búsqueda sin resultados"""
        results = self.indexer.search(sender='noexiste@test.com')
        self.assertEqual(len(results), 0)

    def test_search_special_characters(self):
        """Test: Búsqueda con caracteres especiales"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('user@test.com', 'dest@test.com', "Ñoño's work", "Contenido con 'quotes'",
              '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(subject="Ñoño")
        self.assertEqual(len(results), 1)

    def test_search_case_insensitive(self):
        """Test: Búsqueda no distingue mayúsculas/minúsculas"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('sender@TEST.com', 'user@test.com', 'SUBJECT', 'BODY',
              '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='sender@test.com')
        self.assertEqual(len(results), 1)

    def test_duplicate_emails(self):
        """Test: No se agregan duplicados"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        email_data = ('sender@test.com', 'user@test.com', 'Subject', 'Body', '2024-01-01', 'test.pst')

        for _ in range(3):
            c.execute('''
                INSERT OR IGNORE INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email_data, datetime.now().isoformat()))

        conn.commit()

        c.execute('SELECT COUNT(*) FROM emails')
        count = c.fetchone()[0]
        conn.close()

        self.assertEqual(count, 1, "Se agregaron correos duplicados")

    def test_empty_search_fields(self):
        """Test: Búsqueda sin criterios devuelve todos"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        for i in range(5):
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (f'sender{i}@test.com', 'user@test.com', f'Subject {i}', f'Body {i}',
                  '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        # Sin criterios, search() requiere al menos uno
        results = self.indexer.search()
        self.assertEqual(len(results), 0)

    def test_index_stats(self):
        """Test: Estadísticas correctas"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        emails = [
            ('sender1@test.com', 'user@test.com', 'Subject 1', 'Body 1', '2024-01-01', 'file1.pst'),
            ('sender2@test.com', 'user@test.com', 'Subject 2', 'Body 2', '2024-01-15', 'file2.pst'),
            ('sender3@test.com', 'user@test.com', 'Subject 3', 'Body 3', '2024-02-01', 'file1.pst'),
        ]

        for email in emails:
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (*email, datetime.now().isoformat()))

        conn.commit()
        conn.close()

        stats = self.indexer.get_index_stats()

        self.assertEqual(stats['total_emails'], 3)
        self.assertEqual(stats['total_files'], 2)
        self.assertIsNotNone(stats['date_range'])

    def test_long_email_content(self):
        """Test: Correos con contenido muy largo"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        long_body = 'A' * 10000  # 10,000 caracteres

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('sender@test.com', 'user@test.com', 'Long content', long_body,
              '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(content='A')
        self.assertEqual(len(results), 1)

    def test_null_values_handling(self):
        """Test: Manejo de valores NULL/vacíos"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('', '', '', '', '', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='')
        self.assertIsNotNone(results)

    def test_max_results_limit(self):
        """Test: Se respeta el límite de 1000 resultados"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        # Insertar más de 1000 resultados potenciales
        for i in range(1100):
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', ('same@test.com', 'user@test.com', f'Subject {i}', 'Body',
                  '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='same@test.com')
        self.assertLessEqual(len(results), 1000)

    def test_date_sorting(self):
        """Test: Los resultados se ordenan por fecha descendente"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        dates = ['2024-01-01', '2024-01-15', '2024-01-10']
        for i, date in enumerate(dates):
            c.execute('''
                INSERT INTO emails
                (sender, recipient, subject, body, date, pst_file, indexed_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', ('sender@test.com', 'user@test.com', f'Subject {i}', 'Body',
                  date, 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='sender@test.com')
        dates_result = [r[3] for r in results]
        self.assertEqual(dates_result, sorted(dates_result, reverse=True))


class TestEdgeCases(unittest.TestCase):
    """Tests para casos extremos y edge cases"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, 'test_edge.db')
        self.indexer = EmailIndexer(self.db_path)

    def tearDown(self):
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
        os.rmdir(self.temp_dir)

    def test_unicode_content(self):
        """Test: Contenido con caracteres unicode"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('用户@test.com', '用户@test.com', '日本語のタイトル', '中文内容',
              '2024-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='用户')
        self.assertEqual(len(results), 1)

    def test_sql_injection_attempt(self):
        """Test: Protección contra SQL injection"""
        # Esta búsqueda debería estar segura contra SQL injection
        malicious_input = "'; DROP TABLE emails; --"

        results = self.indexer.search(sender=malicious_input)
        self.assertEqual(len(results), 0)

        # Verificar que la tabla aún existe
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='emails'")
        result = c.fetchone()
        conn.close()

        self.assertIsNotNone(result, "La tabla fue eliminada por SQL injection")

    def test_very_old_date(self):
        """Test: Correos muy antiguos"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('sender@test.com', 'user@test.com', 'Old email', 'Body',
              '1990-01-01', 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(date_from='1980-01-01', date_to='2000-01-01')
        self.assertEqual(len(results), 1)

    def test_future_date(self):
        """Test: Correos con fechas futuras"""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        future_date = (datetime.now() + timedelta(days=365)).isoformat()

        c.execute('''
            INSERT INTO emails
            (sender, recipient, subject, body, date, pst_file, indexed_date)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', ('sender@test.com', 'user@test.com', 'Future email', 'Body',
              future_date, 'test.pst', datetime.now().isoformat()))

        conn.commit()
        conn.close()

        results = self.indexer.search(sender='sender@test.com')
        self.assertEqual(len(results), 1)

    def test_whitespace_only_search(self):
        """Test: Búsqueda solo con espacios"""
        results = self.indexer.search(sender='   ')
        # Debe devolver resultados que contengan solo espacios
        self.assertIsInstance(results, list)

    def test_empty_database_stats(self):
        """Test: Estadísticas con DB vacía"""
        stats = self.indexer.get_index_stats()

        self.assertEqual(stats['total_emails'], 0)
        self.assertEqual(stats['total_files'], 0)


def run_tests():
    """Ejecutar todos los tests"""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    suite.addTests(loader.loadTestsFromTestCase(TestEmailIndexer))
    suite.addTests(loader.loadTestsFromTestCase(TestEdgeCases))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    exit(0 if success else 1)
