"""Pruebas automaticas del sistema. Ejecuta:  python probar_sistema.py"""
import os, sys, tempfile, unittest
from motor_busqueda import (BaseCorreos, normalizar_fecha, importar_base_antigua,
                            carpeta_datos, porcentaje)
from indexador_outlook import html_a_texto, outlook_disponible

CORREOS = [
    dict(entry_id="A1", carpeta="Bandeja de entrada", remitente="María Gómez",
         correo_rem="maria.gomez@empresa.com", destinatarios="isai@empresa.com",
         asunto="Factura 4471 pendiente de pago",
         cuerpo="Estimado Isai, le recuerdo que la factura 4471 sigue pendiente. "
                "El monto es de $12,450 MXN con vencimiento el 30 de marzo.",
         fecha="2024-03-15 10:23:45", adjuntos=1),
    dict(entry_id="A2", carpeta="Bandeja de entrada", remitente="Soporte Técnico",
         correo_rem="soporte@proveedor.com", destinatarios="isai@empresa.com",
         asunto="RE: Problema con el servidor de correo",
         cuerpo="Ya reiniciamos el servidor. El incidente quedó cerrado.",
         fecha="2024-05-02 08:00:00", adjuntos=0),
    dict(entry_id="A3", carpeta="Elementos enviados", remitente="Isai Carreto",
         correo_rem="isai@empresa.com", destinatarios="maria.gomez@empresa.com",
         asunto="Confirmación de pago factura 4471",
         cuerpo="Hola María, adjunto el comprobante del pago de la factura 4471.",
         fecha="2024-03-16 14:05:00", adjuntos=2),
    dict(entry_id="A4", carpeta="Proyectos/Cliente Norte", remitente="Ana Ruiz",
         correo_rem="ana@clientenorte.com", destinatarios="isai@empresa.com",
         asunto="Propuesta comercial 2025",
         cuerpo="Buen día, enviamos la propuesta comercial actualizada para el 2025.",
         fecha="2025-01-10 09:30:00", adjuntos=1),
    dict(entry_id="A5", carpeta="Bandeja de entrada", remitente="Boletín Noticias",
         correo_rem="news@boletin.com", destinatarios="isai@empresa.com",
         asunto="Resumen semanal", cuerpo="", fecha="2024-06-01 07:00:00", adjuntos=0),
]


class BaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mktemp(suffix=".db")
        self.db = BaseCorreos(self.tmp)
        self.db.guardar(CORREOS)

    def tearDown(self):
        self.db.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.tmp + s)
            except OSError: pass

    def test_fts5_activo(self):
        self.assertTrue(self.db.fts, "FTS5 deberia estar disponible")

    def test_total(self):
        self.assertEqual(self.db.total(), 5)

    def test_upsert_actualiza_cuerpo(self):
        """BUG 3 corregido: reindexar SI actualiza el contenido."""
        nuevo = dict(CORREOS[4]); nuevo["cuerpo"] = "Contenido recuperado"
        self.db.guardar([nuevo])
        self.assertEqual(self.db.total(), 5, "no debe duplicar")
        fila = [c for c in self.db.recientes(50) if c["id"]][0]
        recuperado = [c for c in self.db.buscar(texto="recuperado")]
        self.assertEqual(len(recuperado), 1)

    def test_buscar_texto_libre(self):
        r = self.db.buscar(texto="factura")
        self.assertEqual(len(r), 2)

    def test_buscar_por_numero(self):
        self.assertEqual(len(self.db.buscar(texto="4471")), 2)

    def test_buscar_sin_acentos(self):
        """Escribir sin acentos debe encontrar palabras con acento."""
        self.assertGreaterEqual(len(self.db.buscar(texto="Maria")), 1)
        self.assertGreaterEqual(len(self.db.buscar(texto="tecnico")), 1)

    def test_buscar_prefijo(self):
        self.assertGreaterEqual(len(self.db.buscar(texto="confirm")), 1)

    def test_varias_palabras_es_Y(self):
        self.assertEqual(len(self.db.buscar(texto="factura pendiente")), 1)

    def test_frase_exacta(self):
        self.assertEqual(len(self.db.buscar(texto='"propuesta comercial"')), 1)

    def test_caracteres_raros_no_rompen(self):
        """El usuario puede escribir cualquier cosa sin que truene."""
        for basura in ['"', '((', 'a OR', '*', 'NEAR(', "'", '\\', 'a AND AND b', '^%$#@!']:
            self.db.buscar(texto=basura)   # no debe lanzar excepcion

    def test_filtro_remitente(self):
        self.assertEqual(len(self.db.buscar(remitente="maria")), 1)
        self.assertEqual(len(self.db.buscar(remitente="clientenorte.com")), 1)

    def test_filtro_fechas(self):
        self.assertEqual(len(self.db.buscar(desde="2024-03-01", hasta="2024-03-31")), 2)
        self.assertEqual(len(self.db.buscar(desde="2025-01-01")), 1)

    def test_filtro_carpeta(self):
        self.assertEqual(len(self.db.buscar(carpeta="Elementos enviados")), 1)

    def test_combinado(self):
        r = self.db.buscar(texto="factura", carpeta="Bandeja de entrada")
        self.assertEqual(len(r), 1)

    def test_orden_por_fecha_desc(self):
        f = [c["fecha"] for c in self.db.recientes(10)]
        self.assertEqual(f, sorted(f, reverse=True))

    def test_por_id_devuelve_cuerpo_completo(self):
        """BUG 1 corregido: el detalle se obtiene por id, no por texto recortado."""
        r = self.db.buscar(texto="factura")
        self.assertEqual([c["fecha"] for c in r], sorted([c["fecha"] for c in r], reverse=True))
        # tomamos el correo concreto de Maria, no el primero de la lista
        objetivo = [c for c in r if c["remitente"] == "María Gómez"][0]
        completo = self.db.por_id(objetivo["id"])
        self.assertIn("12,450", completo["cuerpo"])
        self.assertEqual(completo["asunto"], "Factura 4471 pendiente de pago")

    def test_remitente_largo_conserva_cuerpo(self):
        """Reproduce el BUG 1 original con nombres largos."""
        largo = "Departamento de Administracion y Finanzas Corporativas Region Norte SA de CV"
        asunto = ("Notificacion automatica del sistema de gestion documental sobre "
                  "expediente numero 99887766 con acuse de recibo")
        self.db.guardar([dict(entry_id="B1", carpeta="Bandeja de entrada", remitente=largo,
                              correo_rem="admin@x.com", destinatarios="isai@empresa.com",
                              asunto=asunto, cuerpo="CONTENIDO IMPORTANTE QUE DEBE VERSE",
                              fecha="2024-07-01 10:00:00", adjuntos=0)])
        r = self.db.buscar(texto="99887766")
        self.assertEqual(len(r), 1)
        self.assertIn("CONTENIDO IMPORTANTE", self.db.por_id(r[0]["id"])["cuerpo"])

    def test_cobertura_de_cuerpo(self):
        self.assertEqual(self.db.total_con_cuerpo(), 4)

    def test_carpetas(self):
        self.assertEqual(len(self.db.carpetas()), 3)

    def test_limite(self):
        self.assertLessEqual(len(self.db.buscar(limite=2)), 2)


class FechaTests(unittest.TestCase):
    def test_formatos(self):
        casos = [
            ("2024-03-15 10:23:45", "2024-03-15 10:23:45"),
            ("2024-03-15 10:23:45+00:00", "2024-03-15 10:23:45"),
            ("2024-03-15T10:23:45", "2024-03-15 10:23:45"),
            ("2024-03-15", "2024-03-15 00:00:00"),
            ("03/15/2024 10:23:45 AM", "2024-03-15 10:23:45"),
            ("15/03/2024 14:05", "2024-03-15 14:05:00"),
            ("03/15/24 1:05 PM", "2024-03-15 13:05:00"),
            ("12/15/2024 12:30 AM", "2024-12-15 00:30:00"),
            ("", ""),
            ("basura", ""),
        ]
        for entrada, esperado in casos:
            self.assertEqual(normalizar_fecha(entrada), esperado, f"fallo con {entrada!r}")


class HtmlTests(unittest.TestCase):
    def test_limpia_html_real(self):
        html = """<html><head><style>.x{color:red}</style></head><body>
        <!-- comentario --><div>Hola&nbsp;Isai,</div><p>El pago de <b>$1,200</b> ya se aplic&oacute;.</p>
        <script>alert(1)</script><br><table><tr><td>Total</td><td>1200</td></tr></table>
        </body></html>"""
        t = html_a_texto(html)
        self.assertNotIn("<", t)
        self.assertNotIn("color:red", t)
        self.assertNotIn("alert", t)
        self.assertNotIn("comentario", t)
        self.assertIn("Hola Isai", t)
        self.assertIn("$1,200", t)
        self.assertIn("aplicó", t)
        self.assertNotIn("\n\n\n", t)

    def test_vacio(self):
        self.assertEqual(html_a_texto(""), "")
        self.assertEqual(html_a_texto(None), "")


class MigracionTests(unittest.TestCase):
    def test_importa_base_vieja(self):
        import sqlite3
        d = tempfile.mkdtemp()
        vieja = os.path.join(d, "email_index.db")
        con = sqlite3.connect(vieja)
        con.execute("""CREATE TABLE emails (id INTEGER PRIMARY KEY, sender TEXT, recipient TEXT,
                       subject TEXT, body TEXT, date TEXT, pst_file TEXT, indexed_date TEXT)""")
        con.executemany("INSERT INTO emails (sender,recipient,subject,body,date,pst_file) "
                        "VALUES (?,?,?,?,?,?)",
                        [("Ana", "isai@x.com", "Hola", "Cuerpo viejo", "2024-01-01 08:00:00",
                          "Outlook:Bandeja de entrada")])
        con.commit(); con.close()

        nueva = os.path.join(d, "correos.db")
        db = BaseCorreos(nueva)
        n = importar_base_antigua(db, vieja)
        self.assertEqual(n, 1)
        self.assertEqual(db.total(), 1)
        self.assertEqual(len(db.buscar(texto="viejo")), 1)
        db.cerrar()

    def test_solo_importa_una_vez(self):
        """Borrar correos.db debe dejar la app limpia, no reimportar lo viejo."""
        import sqlite3, os as _os
        d = tempfile.mkdtemp()
        vieja = _os.path.join(d, "email_index.db")
        con = sqlite3.connect(vieja)
        con.execute("CREATE TABLE emails (id INTEGER PRIMARY KEY, sender TEXT, recipient TEXT,"
                    " subject TEXT, body TEXT, date TEXT, pst_file TEXT, indexed_date TEXT)")
        con.execute("INSERT INTO emails (sender,subject,body,date,pst_file)"
                    " VALUES ('Ana','Hola','cuerpo','2024-01-01 08:00:00','Outlook:Bandeja')")
        con.commit(); con.close()

        b1 = BaseCorreos(_os.path.join(d, "correos.db"))
        self.assertEqual(importar_base_antigua(b1, vieja), 1)
        b1.cerrar()
        _os.remove(_os.path.join(d, "correos.db"))

        b2 = BaseCorreos(_os.path.join(d, "correos.db"))
        self.assertEqual(importar_base_antigua(b2, vieja), 0,
                         "no debe reimportar tras borrar la base")
        self.assertEqual(b2.total(), 0)
        b2.cerrar()

    def test_sin_base_vieja_no_truena(self):
        db = BaseCorreos(os.path.join(tempfile.mkdtemp(), "correos.db"))
        self.assertEqual(importar_base_antigua(db, "no_existe_12345.db"), 0)
        db.cerrar()


class CarpetaDatosTests(unittest.TestCase):
    """Donde se guarda la base al empaquetar con PyInstaller."""

    def test_devuelve_carpeta_escribible(self):
        d = carpeta_datos()
        self.assertTrue(os.path.isdir(d))
        p = os.path.join(d, ".prueba_tmp")
        with open(p, "w") as f:
            f.write("x")
        os.remove(p)

    def test_congelado_usa_la_carpeta_del_ejecutable(self):
        import sys as _s, tempfile
        d = tempfile.mkdtemp()
        exe = os.path.join(d, "BuscadorCorreos.exe")
        with open(exe, "wb") as f:
            f.write(b"MZ")
        viejo_frozen = getattr(_s, "frozen", None)
        viejo_exe = _s.executable
        try:
            _s.frozen = True
            _s.executable = exe
            self.assertEqual(carpeta_datos(), d,
                             "la base debe quedar junto al .exe, no en la carpeta temporal")
        finally:
            _s.executable = viejo_exe
            if viejo_frozen is None:
                del _s.frozen
            else:
                _s.frozen = viejo_frozen

    def test_carpeta_de_solo_lectura_cae_a_datos_del_usuario(self):
        import sys as _s, tempfile, stat
        d = tempfile.mkdtemp()
        exe = os.path.join(d, "app.exe")
        with open(exe, "wb") as f:
            f.write(b"MZ")
        viejo_frozen = getattr(_s, "frozen", None)
        viejo_exe = _s.executable
        os.chmod(d, stat.S_IREAD | stat.S_IEXEC)
        try:
            _s.frozen = True
            _s.executable = exe
            destino = carpeta_datos()
            self.assertNotEqual(destino, d, "no debe intentar escribir donde no puede")
            self.assertTrue(os.path.isdir(destino))
        finally:
            os.chmod(d, stat.S_IRWXU)
            _s.executable = viejo_exe
            if viejo_frozen is None:
                del _s.frozen
            else:
                _s.frozen = viejo_frozen


class PorcentajeTests(unittest.TestCase):
    """El porcentaje no debe mentir por redondeo."""

    def test_no_redondea_a_100_si_falta_alguno(self):
        self.assertEqual(porcentaje(5713, 5728), "99.7%")
        self.assertEqual(porcentaje(9999, 10000), "99.9%")
        self.assertEqual(porcentaje(99999, 100000), "99.9%")

    def test_100_solo_cuando_estan_todos(self):
        self.assertEqual(porcentaje(5728, 5728), "100%")
        self.assertEqual(porcentaje(1, 1), "100%")

    def test_casos_limite(self):
        self.assertEqual(porcentaje(0, 0), "0%")
        self.assertEqual(porcentaje(0, 100), "0%")
        self.assertEqual(porcentaje(1, 3), "33.3%")

    def test_la_barra_de_estado_dice_cuantos_faltan(self):
        import tempfile as tf
        ruta = os.path.join(tf.mkdtemp(), "correos.db")
        db = BaseCorreos(ruta)
        db.guardar([dict(entry_id=f"X{i}", carpeta="B", remitente="A", correo_rem="",
                         destinatarios="", asunto=f"a{i}",
                         cuerpo="" if i < 15 else "texto",
                         fecha="2024-01-01 00:00:00", adjuntos=0)
                    for i in range(5728)])
        self.assertEqual(db.total(), 5728)
        self.assertEqual(db.total_con_cuerpo(), 5713)
        self.assertEqual(porcentaje(db.total_con_cuerpo(), db.total()), "99.7%")
        db.cerrar()


class EntornoTests(unittest.TestCase):
    def test_outlook_reporta_sin_reventar(self):
        ok, motivo = outlook_disponible()
        self.assertIsInstance(ok, bool)
        if not ok:
            self.assertTrue(motivo)


if __name__ == "__main__":
    unittest.main(verbosity=2)
