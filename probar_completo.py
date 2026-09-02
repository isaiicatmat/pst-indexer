# -*- coding: utf-8 -*-
"""Prueba de extremo a extremo: abrir la app, indexar desde Outlook (simulado),
buscar y leer el contenido. Ejecuta:  python probar_completo.py"""
import os, sys, tempfile, unittest
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication

# Qt en modo sin ventana avisa de fuentes que no encuentra, cientos de veces.
# No afecta a la aplicacion real: solo ensucia el reporte de las pruebas.
def _silenciar_qt(tipo, contexto, mensaje):
    ruido = ("qfontdatabase", "cannot find font", "qt.qpa", "populating font")
    if any(r in str(mensaje).lower() for r in ruido):
        return
    sys.stderr.write(str(mensaje) + "\n")


from PyQt5.QtCore import qInstallMessageHandler
qInstallMessageHandler(_silenciar_qt)

from PyQt5.QtCore import Qt, QElapsedTimer

import indexador_outlook as ix
import buscador_correos as bc
from probar_outlook import outlook_de_prueba, instalar_falso, FakeItem

app = QApplication.instance() or QApplication(sys.argv)


def esperar(hilo, ms=15000):
    t = QElapsedTimer(); t.start()
    while hilo.isRunning() and t.elapsed() < ms:
        app.processEvents()
        hilo.wait(20)
    for _ in range(10):          # dejar llegar la senal de "terminado"
        app.processEvents()


class FlujoCompletoTests(unittest.TestCase):
    def setUp(self):
        instalar_falso(outlook_de_prueba())
        bc.outlook_disponible = lambda: (True, "")
        self.ruta = tempfile.mktemp(suffix=".db")
        self.v = bc.Ventana(self.ruta)
        self.v.show()               # para que isVisible() signifique algo
        app.processEvents()

    def tearDown(self):
        if self.v.hilo and self.v.hilo.isRunning():
            self.v.hilo.cancelar(); self.v.hilo.wait(3000)
        self.v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.ruta + s)
            except OSError: pass

    def test_01_flujo_del_usuario_de_principio_a_fin(self):
        # 1. Abre la app por primera vez: ve la bienvenida, sin barra de busqueda
        self.assertEqual(self.v.pilas.currentIndex(), 1)
        self.assertFalse(self.v.barra_busqueda.isVisible())

        # 2. Pulsa "Actualizar correos"
        self.v.actualizar_correos()
        self.assertIsNotNone(self.v.hilo)
        esperar(self.v.hilo)
        app.processEvents()

        # 3. Ya hay correos: aparece la lista y la barra de busqueda
        self.assertEqual(self.v.pilas.currentIndex(), 0)
        self.assertTrue(self.v.barra_busqueda.isVisible())
        self.assertEqual(self.v.base.total(), 7)
        self.assertEqual(self.v.lista.count(), 7)

        # 4. El dialogo de progreso se cerro solo
        self.assertIsNone(self.v.dialogo)
        self.assertTrue(self.v.btn_actualizar.isEnabled())

        # 5. Busca "factura" y ve el contenido sin hacer doble clic
        self.v.caja.setText("factura"); self.v.buscar(); app.processEvents()
        self.assertEqual(self.v.lista.count(), 2, "la factura y su respuesta")
        # se elige el correo concreto, no se asume el orden
        fila = next(i for i in range(self.v.lista.count())
                    if self.v.lista.item(i).data(Qt.UserRole)["remitente"] == "María Gómez")
        self.v.lista.setCurrentRow(fila); app.processEvents()
        self.assertIn("12,450", self.v.cuerpo.toPlainText())
        self.assertIn("Factura 4471", self.v.cabecera.text())

        # 6. Todos los resultados muestran contenido real
        for fila in range(self.v.lista.count()):
            self.v.lista.setCurrentRow(fila); app.processEvents()
            self.assertTrue(self.v.cuerpo.toPlainText().strip())
            self.assertNotIn("no tiene texto guardado", self.v.cuerpo.toPlainText())

        # 7. El desplegable de carpetas incluye las subcarpetas anidadas
        carpetas = [self.v.f_carpeta.itemData(i) for i in range(self.v.f_carpeta.count())]
        self.assertTrue(any("Cliente Norte" in (c or "") for c in carpetas), carpetas)

        # 8. La barra de estado reporta 100% con contenido
        self.assertIn("(100%)", self.v.estado.text())

    def test_02_actualizar_dos_veces_no_duplica(self):
        for _ in range(2):
            self.v.actualizar_correos(); esperar(self.v.hilo)
        self.assertEqual(self.v.base.total(), 7)

    def test_03_correo_nuevo_aparece_al_actualizar(self):
        self.v.actualizar_correos(); esperar(self.v.hilo)

        ns = outlook_de_prueba()
        ns.Folders[0].Folders[0].Items._i.append(
            FakeItem("E100", "Cotización urgente del cliente",
                     body="Necesitamos la cotización antes del viernes."))
        instalar_falso(ns)
        bc.outlook_disponible = lambda: (True, "")

        self.v.actualizar_correos(); esperar(self.v.hilo)

        self.assertEqual(self.v.base.total(), 8)
        self.v.caja.setText("cotización"); self.v.buscar(); app.processEvents()
        self.assertEqual(self.v.lista.count(), 1)
        self.assertIn("antes del viernes", self.v.cuerpo.toPlainText())

    def test_04_error_de_outlook_da_mensaje_claro_y_no_tumba_la_app(self):
        """Si Outlook falla, el usuario debe ver una explicacion, no un cierre brusco."""
        from PyQt5.QtWidgets import QMessageBox
        avisos = []
        original = QMessageBox.warning
        QMessageBox.warning = staticmethod(lambda p, t, m, *a, **k: avisos.append(m))
        try:
            import win32com.client as wc
            def revienta(_):
                raise Exception("El servidor de Outlook no esta disponible")
            wc.Dispatch = revienta

            self.v.actualizar_correos()
            esperar(self.v.hilo)

            self.assertEqual(len(avisos), 1, "debe avisar exactamente una vez")
            self.assertIn("Outlook", avisos[0])
            # la ventana sigue usable y el boton vuelve a habilitarse
            self.assertTrue(self.v.isEnabled())
            self.assertTrue(self.v.btn_actualizar.isEnabled())
            self.assertIsNone(self.v.dialogo)
            # y sigue mostrando la pantalla de bienvenida, no una lista rota
            self.assertEqual(self.v.pilas.currentIndex(), 1)
        finally:
            QMessageBox.warning = original

    def test_05_buscar_sin_acentos_encuentra_con_acentos(self):
        self.v.actualizar_correos(); esperar(self.v.hilo)
        self.v.caja.setText("promocion"); self.v.buscar(); app.processEvents()
        self.assertEqual(self.v.lista.count(), 1, "escribir sin acento debe funcionar")


class MigracionSinDuplicadosTests(unittest.TestCase):
    """La base anterior se usa como puente, pero no debe dejar duplicados."""

    def setUp(self):
        import sqlite3
        instalar_falso(outlook_de_prueba())
        bc.outlook_disponible = lambda: (True, "")
        self.dir = tempfile.mkdtemp()
        self.vieja = os.path.join(self.dir, "email_index.db")
        con = sqlite3.connect(self.vieja)
        con.execute("CREATE TABLE emails (id INTEGER PRIMARY KEY, sender TEXT, recipient TEXT,"
                    " subject TEXT, body TEXT, date TEXT, pst_file TEXT, indexed_date TEXT)")
        con.executemany("INSERT INTO emails (sender,recipient,subject,body,date,pst_file)"
                        " VALUES (?,?,?,?,?,?)",
                        [("María Gómez", "isai@x.com", "Factura 4471", "Cuerpo antiguo",
                          "2024-03-15 10:23:45", "Outlook:Bandeja de entrada")])
        con.commit(); con.close()
        # A proposito NO cambiamos de directorio: la app debe encontrar la base
        # anterior junto a correos.db, sin depender del directorio de trabajo.

    def test_puente_y_luego_limpieza(self):
        ruta = os.path.join(self.dir, "correos.db")
        v = bc.Ventana(ruta); v.show(); app.processEvents()

        # 1. Al abrir, ya se ven los correos de la version anterior
        self.assertEqual(v.base.total(), 1)
        self.assertEqual(v.pilas.currentIndex(), 0)

        # 2. Tras traerlos de Outlook, los heredados desaparecen: sin duplicados
        v.actualizar_correos(); esperar(v.hilo)
        self.assertEqual(v.base.total(), 7, "7 reales, 0 heredados")
        heredados = v.base.con.execute(
            "SELECT COUNT(*) FROM correos WHERE entry_id LIKE 'legacy:%'").fetchone()[0]
        self.assertEqual(heredados, 0)

        # 3. Y la factura aparece una sola vez, con el cuerpo bueno
        v.caja.setText("factura 4471"); v.buscar(); app.processEvents()
        # coinciden la factura y su respuesta, pero cada una UNA sola vez
        remitentes = [v.lista.item(i).data(Qt.UserRole)["remitente"]
                      for i in range(v.lista.count())]
        self.assertEqual(len(remitentes), len(set(remitentes)), f"hay duplicados: {remitentes}")
        self.assertEqual(remitentes.count("María Gómez"), 1)
        fila = remitentes.index("María Gómez")
        v.lista.setCurrentRow(fila); app.processEvents()
        # y muestra el cuerpo BUENO de Outlook, no el pobre de la version anterior
        self.assertIn("12,450", v.cuerpo.toPlainText())
        self.assertNotIn("Cuerpo antiguo", v.cuerpo.toPlainText())
        v.base.cerrar()


if __name__ == "__main__":
    unittest.main(verbosity=2)
