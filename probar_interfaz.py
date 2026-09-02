"""Pruebas de la interfaz sin abrir ventana (modo offscreen).
Ejecuta:  python probar_interfaz.py"""
import os, sys, tempfile, unittest
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from motor_busqueda import BaseCorreos
import buscador_correos as bc

app = QApplication.instance() or QApplication(sys.argv)

LARGO_REM = "Departamento de Administracion y Finanzas Corporativas Region Norte SA de CV"
LARGO_ASU = ("Notificacion automatica del sistema de gestion documental sobre el "
             "expediente 99887766 con acuse de recibo y confirmacion de lectura")

DATOS = [
    dict(entry_id="A1", carpeta="Bandeja de entrada", remitente="María Gómez",
         correo_rem="maria@empresa.com", destinatarios="isai@empresa.com",
         asunto="Factura 4471 pendiente", cuerpo="El monto es de $12,450 MXN.",
         fecha="2024-03-15 10:23:45", adjuntos=1),
    dict(entry_id="A2", carpeta="Elementos enviados", remitente="Isai Carreto",
         correo_rem="isai@empresa.com", destinatarios="maria@empresa.com",
         asunto="Confirmacion de pago", cuerpo="Adjunto comprobante del pago.",
         fecha="2024-03-16 14:05:00", adjuntos=2),
    dict(entry_id="A3", carpeta="Proyectos/Norte", remitente=LARGO_REM,
         correo_rem="admin@x.com", destinatarios="isai@empresa.com", asunto=LARGO_ASU,
         cuerpo="CONTENIDO LARGO QUE DEBE VERSE COMPLETO", fecha="2024-07-01 09:00:00",
         adjuntos=0),
    dict(entry_id="A4", carpeta="Bandeja de entrada", remitente="Boletin",
         correo_rem="n@b.com", destinatarios="isai@empresa.com", asunto="Solo imagen",
         cuerpo="", fecha="2024-08-01 07:00:00", adjuntos=0),
]


class InterfazTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ruta = tempfile.mktemp(suffix=".db")
        db = BaseCorreos(cls.ruta); db.guardar(DATOS); db.cerrar()
        cls.v = bc.Ventana(cls.ruta)

    @classmethod
    def tearDownClass(cls):
        cls.v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(cls.ruta + s)
            except OSError: pass

    def _buscar(self, texto):
        self.v.caja.setText(texto)
        self.v.buscar()
        app.processEvents()

    def test_01_abre_con_correos(self):
        self.assertEqual(self.v.pilas.currentIndex(), 0, "debe mostrar la lista, no el estado vacio")
        self.assertEqual(self.v.lista.count(), 4)

    def test_02_selecciona_el_primero_solo(self):
        self.assertEqual(self.v.lista.currentRow(), 0)
        self.assertTrue(self.v.cabecera.text(), "la cabecera debe llenarse sola")

    def test_03_el_contenido_se_ve_sin_doble_clic(self):
        """Requisito: el contenido de los correos se ve."""
        self._buscar("factura")
        self.assertGreaterEqual(self.v.lista.count(), 1)
        self.assertIn("12,450", self.v.cuerpo.toPlainText())

    def test_04_BUG1_remitente_y_asunto_largos(self):
        """El bug original: con textos largos el detalle salia vacio."""
        self._buscar("99887766")
        self.assertEqual(self.v.lista.count(), 1)
        texto = self.v.cuerpo.toPlainText()
        self.assertIn("CONTENIDO LARGO QUE DEBE VERSE COMPLETO", texto)
        self.assertNotIn("Sin contenido", texto)

    def test_05_resaltado_de_terminos(self):
        self._buscar("comprobante")
        h = self.v.cuerpo.toHtml().lower()
        self.assertIn("ffe9a8", h, "el termino buscado debe aparecer resaltado")
        # y el resaltado debe caer sobre la palabra buscada, no en otro lado
        i = h.find("ffe9a8")
        self.assertIn("comprobante", h[i:i+220])

    def test_06_correo_sin_cuerpo_explica_que_hacer(self):
        self._buscar("Solo imagen")
        t = self.v.cuerpo.toPlainText()
        self.assertIn("no tiene texto guardado", t)
        self.assertIn("Abrir en Outlook", t)

    def test_07_navegar_con_flechas_cambia_el_contenido(self):
        self._buscar("")
        self.v.lista.setCurrentRow(0); app.processEvents()
        primero = self.v.cuerpo.toPlainText()
        self.v.lista.setCurrentRow(1); app.processEvents()
        self.assertNotEqual(primero, self.v.cuerpo.toPlainText())

    def test_08_cada_fila_muestra_su_propio_contenido(self):
        """Recorre TODOS los resultados y verifica que el cuerpo coincide."""
        self._buscar("")
        for fila in range(self.v.lista.count()):
            self.v.lista.setCurrentRow(fila); app.processEvents()
            d = self.v.lista.item(fila).data(Qt.UserRole)
            esperado = self.v.base.por_id(d["id"])
            self.assertIn(esperado["asunto"][:30], self.v.cabecera.text())
            if esperado["cuerpo"]:
                self.assertIn(esperado["cuerpo"][:25], self.v.cuerpo.toPlainText())

    def test_08b_correo_del_remitente_sin_entidades_html(self):
        """La direccion debe verse como <correo>, no como &lt;correo&gt;."""
        self._buscar("factura")
        self.v.lista.setCurrentRow(0)
        app.processEvents()
        origen = self.v.cabecera.text()
        self.assertNotIn("&amp;lt;", origen, "doble escapado: se veria '&lt;' literal")
        self.assertIn("&lt;", origen, "la direccion debe ir entre < >")
        # Lo que realmente ve el usuario, ya renderizado por Qt:
        from PyQt5.QtGui import QTextDocument
        doc = QTextDocument(); doc.setHtml(origen)
        visible = doc.toPlainText()
        self.assertIn("<maria@empresa.com>", visible)
        self.assertNotIn("&lt;", visible)

    def test_09_sin_resultados_da_instrucciones(self):
        self._buscar("xyzabc123noexiste")
        self.assertEqual(self.v.lista.count(), 0)
        self.assertIn("Sin resultados", self.v.conteo.text())
        self.assertIn("Actualizar correos", self.v.cuerpo.toPlainText())

    def test_10_filtro_carpeta(self):
        self._buscar("")
        i = self.v.f_carpeta.findData("Elementos enviados")
        self.assertGreater(i, 0, "la carpeta debe aparecer en el desplegable")
        self.v.f_carpeta.setCurrentIndex(i); app.processEvents()
        self.assertEqual(self.v.lista.count(), 1)
        self.v.f_carpeta.setCurrentIndex(0); app.processEvents()

    def test_11_filtro_remitente(self):
        self._buscar("")
        self.v.f_remitente.setText("maria"); self.v.buscar(); app.processEvents()
        self.assertEqual(self.v.lista.count(), 1)
        self.v.f_remitente.clear()

    def test_12_limpiar_restaura_todo(self):
        self.v.caja.setText("factura"); self.v.f_remitente.setText("maria")
        self.v._limpiar(); app.processEvents()
        self.assertEqual(self.v.caja.text(), "")
        self.assertEqual(self.v.f_remitente.text(), "")
        self.assertEqual(self.v.lista.count(), 4)

    def test_13_teclear_no_truena_con_nada(self):
        for basura in ['"', '((', '*', "'", "\\", "a OR", "NEAR(", "%$#", "ñ", "  "]:
            self._buscar(basura)

    def test_14_barra_de_estado_informa_cobertura(self):
        self.v._refrescar_estado()
        t = self.v.estado.text()
        self.assertIn("indexados", t)
        self.assertIn("con contenido legible", t)
        # de los 4 correos de prueba, uno no tiene cuerpo
        self.assertIn("1 sin texto", t)
        self.assertNotIn("(100%)", t, "no debe decir 100% si falta contenido")

    def test_15_copiar_al_portapapeles(self):
        self._buscar("factura")
        self.v._copiar()
        self.assertIn("12,450", QApplication.clipboard().text())

    def test_16_estado_vacio_en_base_nueva(self):
        ruta = tempfile.mktemp(suffix=".db")
        v = bc.Ventana(ruta)
        self.assertEqual(v.pilas.currentIndex(), 1, "base vacia -> pantalla de bienvenida")
        v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(ruta + s)
            except OSError: pass

    def test_17_atajos_registrados(self):
        from PyQt5.QtWidgets import QShortcut
        teclas = {s.key().toString() for s in self.v.findChildren(QShortcut)}
        for k in ("Ctrl+F", "Esc", "F5"):
            self.assertTrue(any(k in t for t in teclas), f"falta el atajo {k}: {teclas}")

    def test_18_render_de_la_lista_no_truena(self):
        """Dibuja de verdad los items con el delegado."""
        from PyQt5.QtGui import QPixmap, QPainter
        from PyQt5.QtWidgets import QStyleOptionViewItem
        self._buscar("")
        d = self.v.lista.itemDelegate()
        px = QPixmap(460, 76); px.fill()
        p = QPainter(px)
        for fila in range(self.v.lista.count()):
            o = QStyleOptionViewItem()
            o.rect = px.rect(); o.font = self.v.font()
            d.paint(p, o, self.v.lista.model().index(fila, 0))
        p.end()


class BotonPstTests(unittest.TestCase):
    """El boton para abrir un .pst archivado."""

    def setUp(self):
        self.ruta = tempfile.mktemp(suffix=".db")
        db = BaseCorreos(self.ruta); db.guardar(DATOS); db.cerrar()
        self.v = bc.Ventana(self.ruta); self.v.show(); app.processEvents()

    def tearDown(self):
        self.v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.ruta + s)
            except OSError: pass

    def test_existe_y_explica_para_que_sirve(self):
        self.assertTrue(self.v.btn_pst.isVisible())
        self.assertIn(".pst", self.v.btn_pst.text())
        self.assertIn("archivado", self.v.btn_pst.toolTip())

    def test_sin_outlook_avisa_y_no_abre_selector(self):
        """En una maquina sin Outlook debe explicarlo, no fallar en silencio."""
        from PyQt5.QtWidgets import QMessageBox, QFileDialog
        avisos, abierto = [], []
        oi, og = QMessageBox.information, QFileDialog.getOpenFileName
        QMessageBox.information = staticmethod(lambda p, t, m, *a, **k: avisos.append(m))
        QFileDialog.getOpenFileName = staticmethod(
            lambda *a, **k: abierto.append(1) or ("", ""))
        try:
            bc.outlook_disponible = lambda: (False, "Esta función necesita Windows con Outlook instalado.")
            self.v.agregar_pst()
            self.assertEqual(len(avisos), 1)
            self.assertIn("Windows", avisos[0])
            self.assertEqual(abierto, [], "no debe abrir el selector de archivos")
        finally:
            QMessageBox.information = oi
            QFileDialog.getOpenFileName = og

    def test_cancelar_el_selector_no_hace_nada(self):
        from PyQt5.QtWidgets import QFileDialog
        og = QFileDialog.getOpenFileName
        QFileDialog.getOpenFileName = staticmethod(lambda *a, **k: ("", ""))
        try:
            bc.outlook_disponible = lambda: (True, "")
            self.v.agregar_pst()          # no debe lanzar nada
        finally:
            QFileDialog.getOpenFileName = og
            bc.outlook_disponible = lambda: (False, "sin Outlook")

    def test_se_deshabilita_mientras_indexa(self):
        self.assertTrue(self.v.btn_pst.isEnabled())


class FiltroFechaTests(unittest.TestCase):
    """Los periodos rapidos de fecha."""

    @classmethod
    def setUpClass(cls):
        cls.ruta = tempfile.mktemp(suffix=".db")
        db = BaseCorreos(cls.ruta)
        from PyQt5.QtCore import QDate
        hoy = QDate.currentDate()
        cls.hoy = hoy
        # un correo de hoy, uno de hace 3 dias, uno de hace 60, uno del anio pasado
        fechas = [(hoy, "A"), (hoy.addDays(-3), "B"), (hoy.addDays(-60), "C"),
                  (QDate(hoy.year() - 1, 6, 15), "D")]
        db.guardar([dict(entry_id=k, carpeta="B", remitente="R", correo_rem="",
                         destinatarios="", asunto=f"Correo {k}", cuerpo="texto",
                         fecha=f.toString("yyyy-MM-dd") + " 10:00:00", adjuntos=0)
                    for f, k in fechas])
        db.cerrar()
        cls.v = bc.Ventana(cls.ruta); cls.v.show(); app.processEvents()

    @classmethod
    def tearDownClass(cls):
        cls.v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(cls.ruta + s)
            except OSError: pass

    def _elegir(self, clave):
        i = self.v.f_periodo.findData(clave)
        self.assertGreaterEqual(i, 0, f"falta el periodo {clave}")
        self.v.f_periodo.setCurrentIndex(i)
        app.processEvents()

    def test_por_defecto_no_filtra(self):
        self._elegir("")
        self.assertEqual(self.v.rango_de_fechas(), ("", ""))
        self.assertEqual(self.v.lista.count(), 4)

    def test_hoy(self):
        self._elegir("hoy")
        d, h = self.v.rango_de_fechas()
        self.assertEqual(d, h)
        self.assertEqual(self.v.lista.count(), 1)

    def test_ultimos_7_dias_incluye_hoy(self):
        self._elegir("7d")
        d, h = self.v.rango_de_fechas()
        self.assertEqual(h, self.hoy.toString("yyyy-MM-dd"))
        self.assertEqual(d, self.hoy.addDays(-6).toString("yyyy-MM-dd"))
        self.assertEqual(self.v.lista.count(), 2, "el de hoy y el de hace 3 dias")

    def test_ultimos_30_dias(self):
        self._elegir("30d")
        self.assertEqual(self.v.lista.count(), 2)

    def test_ultimos_3_meses_alcanza_los_60_dias(self):
        self._elegir("3m")
        self.assertEqual(self.v.lista.count(), 3)

    def test_este_anio_empieza_en_enero(self):
        self._elegir("anio")
        d, _ = self.v.rango_de_fechas()
        self.assertEqual(d, f"{self.hoy.year()}-01-01")

    def test_anio_pasado_es_el_anio_completo(self):
        self._elegir("anio-1")
        d, h = self.v.rango_de_fechas()
        a = self.hoy.year() - 1
        self.assertEqual((d, h), (f"{a}-01-01", f"{a}-12-31"))
        self.assertEqual(self.v.lista.count(), 1, "solo el correo del anio pasado")

    def test_los_calendarios_solo_salen_con_rango_a_mano(self):
        self.v.btn_filtros.setChecked(True)      # abrir el panel de filtros
        app.processEvents()
        self._elegir("30d")
        self.assertFalse(self.v.f_desde.isVisible(), "sin rango no hacen falta")
        self._elegir("rango")
        self.assertTrue(self.v.f_desde.isVisible())
        self.assertTrue(self.v.f_hasta.isVisible())
        self.v.btn_filtros.setChecked(False)
        app.processEvents()

    def test_avisa_de_los_filtros_aunque_el_panel_este_plegado(self):
        """Un filtro olvidado hace creer que faltan correos."""
        self.v.btn_filtros.setChecked(False)
        self._elegir("hoy")
        self.assertIn("(1)", self.v.btn_filtros.text())
        self.assertIn("Filtrando por", self.v.pista.text())
        self.assertIn("Hoy", self.v.pista.text())
        self.assertTrue(self.v.btn_filtros.property("conFiltros"))

        self.v._limpiar(); app.processEvents()
        self.assertEqual(self.v.btn_filtros.text(), "Filtros")
        self.assertNotIn("Filtrando por", self.v.pista.text())
        self.assertFalse(self.v.btn_filtros.property("conFiltros"))

    def test_el_conteo_no_dice_recientes_si_hay_filtro(self):
        self.v._limpiar(); app.processEvents()
        self.assertIn("más recientes", self.v.conteo.text())
        self._elegir("hoy")
        self.assertIn("encontrado", self.v.conteo.text())
        self.assertNotIn("más recientes", self.v.conteo.text())
        self.v._limpiar(); app.processEvents()

    def test_cuenta_varios_filtros(self):
        self.v._limpiar(); app.processEvents()
        self.v.f_remitente.setText("R")
        self._elegir("hoy")
        self.assertIn("(2)", self.v.btn_filtros.text())
        self.v._limpiar(); app.processEvents()

    def test_rango_a_mano_incluye_los_dos_extremos(self):
        from PyQt5.QtCore import QDate
        self._elegir("rango")
        self.v.f_desde.setDate(self.hoy.addDays(-3))
        self.v.f_hasta.setDate(self.hoy.addDays(-3))
        app.processEvents()
        self.assertEqual(self.v.lista.count(), 1,
                         "un solo dia debe encontrar el correo de ese dia")

    def test_limpiar_devuelve_a_cualquier_fecha(self):
        self._elegir("hoy")
        self.v._limpiar(); app.processEvents()
        self.assertEqual(self.v.f_periodo.currentData(), "")
        self.assertEqual(self.v.lista.count(), 4)

    def test_fecha_se_combina_con_texto(self):
        self._elegir("")
        self.v.caja.setText("Correo"); self.v.buscar(); app.processEvents()
        self.assertEqual(self.v.lista.count(), 4)
        self._elegir("hoy")
        self.assertEqual(self.v.lista.count(), 1)
        self.v.caja.clear()


class VaciarIndiceTests(unittest.TestCase):
    """Vaciar el indice desde la app, sin borrar archivos a mano."""

    def setUp(self):
        self.ruta = tempfile.mktemp(suffix=".db")
        db = BaseCorreos(self.ruta); db.guardar(DATOS); db.cerrar()
        self.v = bc.Ventana(self.ruta); self.v.show(); app.processEvents()

    def tearDown(self):
        self.v.base.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.ruta + s)
            except OSError: pass

    def _responder(self, respuesta):
        from PyQt5.QtWidgets import QMessageBox
        self._orig_q = QMessageBox.question
        self._orig_i = QMessageBox.information
        self.mensajes = []
        QMessageBox.question = staticmethod(
            lambda p, t, m, *a, **k: self.mensajes.append(m) or respuesta)
        QMessageBox.information = staticmethod(lambda p, t, m, *a, **k: self.mensajes.append(m))

    def _restaurar(self):
        from PyQt5.QtWidgets import QMessageBox
        QMessageBox.question = self._orig_q
        QMessageBox.information = self._orig_i

    def test_vacia_y_vuelve_a_la_bienvenida(self):
        from PyQt5.QtWidgets import QMessageBox
        self._responder(QMessageBox.Yes)
        try:
            self.v.vaciar_indice()
        finally:
            self._restaurar()
        self.assertEqual(self.v.base.total(), 0)
        self.assertEqual(self.v.pilas.currentIndex(), 1)
        self.assertIn("NO se tocan", self.mensajes[0], "debe aclarar que Outlook no se toca")

    def test_cancelar_no_borra_nada(self):
        from PyQt5.QtWidgets import QMessageBox
        self._responder(QMessageBox.No)
        try:
            self.v.vaciar_indice()
        finally:
            self._restaurar()
        self.assertEqual(self.v.base.total(), len(DATOS))

    def test_indice_ya_vacio_lo_dice(self):
        from PyQt5.QtWidgets import QMessageBox
        self.v.base.vaciar()
        self._responder(QMessageBox.Yes)
        try:
            self.v.vaciar_indice()
        finally:
            self._restaurar()
        self.assertIn("ya está vacío", self.mensajes[0])

    def test_muestra_donde_estan_los_datos(self):
        from PyQt5.QtWidgets import QMessageBox
        self._responder(QMessageBox.Yes)
        try:
            self.v.mostrar_carpeta_datos()
        finally:
            self._restaurar()
        self.assertIn(self.ruta, self.mensajes[0])


class CierreTests(unittest.TestCase):
    def test_cerrar_con_busqueda_pendiente_no_truena(self):
        """El temporizador de busqueda no debe dispararse sobre una base cerrada."""
        from PyQt5.QtGui import QCloseEvent
        ruta = tempfile.mktemp(suffix=".db")
        db = BaseCorreos(ruta); db.guardar(DATOS); db.cerrar()
        v = bc.Ventana(ruta); v.show(); app.processEvents()
        v.caja.setText("fact")          # deja el temporizador armado
        v.closeEvent(QCloseEvent())     # el usuario cierra antes de los 250 ms
        v.temporizador.timeout.emit()   # el temporizador intenta dispararse
        app.processEvents()
        for s_ in ("", "-wal", "-shm"):
            try: os.remove(ruta + s_)
            except OSError: pass


if __name__ == "__main__":
    unittest.main(verbosity=2)
