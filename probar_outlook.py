# -*- coding: utf-8 -*-
"""Prueba del indexador usando un Outlook simulado.
Verifica lo que no se puede probar sin Windows.
Ejecuta:  python probar_outlook.py"""
import os, sys, tempfile, types, unittest
from datetime import datetime

from motor_busqueda import BaseCorreos
import indexador_outlook as ix


# ----------------------------------------------------- Outlook de mentira
class FakeAttachments:
    def __init__(self, n): self.Count = n

class FakePA:
    def __init__(self, props): self.props = props
    def GetProperty(self, esquema):
        if esquema in self.props: return self.props[esquema]
        raise Exception("propiedad no disponible")

class FakeItem:
    def __init__(self, eid, asunto, remitente="Ana", correo="ana@x.com", para="isai@x.com",
                 body=None, html=None, preview="", fecha=None, adj=0, clase=43, props=None):
        self.EntryID = eid; self.Subject = asunto; self.SenderName = remitente
        self.SenderEmailAddress = correo; self.To = para
        self.ReceivedTime = fecha or datetime(2024, 5, 1, 9, 0, 0)
        self.Attachments = FakeAttachments(adj); self.Class = clase
        self.Preview = preview
        self.PropertyAccessor = FakePA(props or {})
        if body is not None: self.Body = body
        if html is not None: self.HTMLBody = html
    def __getattr__(self, n): raise AttributeError(n)

class FakeItems:
    def __init__(self, items): self._i = list(items); self._p = 0
    @property
    def Count(self): return len(self._i)
    def Sort(self, *a): pass
    def GetFirst(self):
        self._p = 0
        return self._i[0] if self._i else None
    def GetNext(self):
        self._p += 1
        return self._i[self._p] if self._p < len(self._i) else None

class FakeStore:
    def __init__(self, ruta): self.FilePath = ruta

class FakeFolder:
    def __init__(self, nombre, items=(), sub=(), archivo=None):
        self.Name = nombre; self.Items = FakeItems(items); self.Folders = list(sub)
        self.Store = FakeStore(archivo) if archivo else None

class FakeNS:
    """Simula AddStoreEx/RemoveStore de Outlook."""
    def __init__(self, folders): self.Folders = folders
    def AddStoreEx(self, ruta, tipo):
        if str(ruta).lower().endswith("roto.pst"):
            raise Exception("el archivo esta danado")
        self.Folders.append(FakeFolder(
            os.path.basename(str(ruta)),
            sub=[FakeFolder("Bandeja de entrada", [
                FakeItem("P1", "Correo archivado 2019", "Cliente Viejo", "viejo@x.com",
                         body="Contenido dentro del PST archivado.")])],
            archivo=str(ruta)))
    def RemoveStore(self, carpeta):
        self.Folders.remove(carpeta)
class FakeApp:
    def __init__(self, ns): self._ns = ns
    def GetNamespace(self, _): return self._ns


def outlook_de_prueba():
    entrada = FakeFolder("Bandeja de entrada", [
        FakeItem("E1", "Factura 4471", "María Gómez", "maria@x.com",
                 body="El monto es de $12,450 MXN.", adj=1,
                 fecha=datetime(2024, 3, 15, 10, 23, 45)),
        # solo HTML: debe convertirse a texto legible
        FakeItem("E2", "Boletín mensual", "Marketing", "mk@x.com",
                 html="<html><head><style>.a{color:red}</style></head><body>"
                      "<p>Hola&nbsp;Isai</p><script>x()</script>"
                      "<div>Promoci&oacute;n del mes</div></body></html>"),
        # sin Body ni HTML: cae al PropertyAccessor
        FakeItem("E3", "Aviso", "Sistema", "sys@x.com",
                 props={ix.PR_BODY: "Texto recuperado por propiedad MAPI"}),
        # ni body ni propiedad: usa Preview
        FakeItem("E4", "Solo imagen", "Publicidad", "p@x.com", preview="Vista previa breve"),
        # no es correo (una cita): debe ignorarse
        FakeItem("E5", "Junta de equipo", clase=26),
    ], sub=[
        FakeFolder("Proyectos", [
            FakeItem("E6", "Propuesta Norte", "Ana Ruiz", "ana@n.com",
                     body="Propuesta comercial 2025."),
        ], sub=[
            FakeFolder("Cliente Norte", [
                FakeItem("E7", "Contrato firmado", "Legal", "legal@n.com",
                         body="Se anexa el contrato firmado."),
            ])
        ])
    ])
    enviados = FakeFolder("Elementos enviados", [
        FakeItem("E8", "RE: Factura 4471", "Isai Carreto", "isai@x.com",
                 body="Adjunto comprobante.", adj=2),
    ])
    basura = FakeFolder("Elementos eliminados", [
        FakeItem("E9", "Correo borrado", body="No debe indexarse"),
    ])
    return FakeNS([FakeFolder("Buzón - Isai", sub=[entrada, enviados, basura])])


def instalar_falso(ns):
    """Suplanta pythoncom y win32com para poder probar sin Windows."""
    pc = types.ModuleType("pythoncom")
    pc.CoInitialize = lambda: None
    pc.CoUninitialize = lambda: None
    sys.modules["pythoncom"] = pc
    w = types.ModuleType("win32com"); c = types.ModuleType("win32com.client")
    c.Dispatch = lambda _: FakeApp(ns)
    w.client = c
    sys.modules["win32com"] = w; sys.modules["win32com.client"] = c
    ix.outlook_disponible = lambda: (True, "")


class IndexadorTests(unittest.TestCase):
    def setUp(self):
        instalar_falso(outlook_de_prueba())
        self.ruta = tempfile.mktemp(suffix=".db")
        self.db = BaseCorreos(self.ruta)
        self.avisos = []

    def tearDown(self):
        self.db.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.ruta + s)
            except OSError: pass

    def _indexar(self, **kw):
        idx = ix.Indexador(self.db, progreso=lambda t, h=0, n=0: self.avisos.append(t), **kw)
        return idx.ejecutar()

    def test_01_indexa_todo_lo_valido(self):
        n = self._indexar()
        self.assertEqual(n, 7, "deben entrar 7 correos (E1-E4, E6, E7, E8)")
        self.assertEqual(self.db.total(), 7)

    def test_02_recorre_subcarpetas_anidadas(self):
        """BUG 2 corregido: antes solo veia el primer nivel."""
        self._indexar()
        carpetas = {c for c, _ in self.db.carpetas()}
        self.assertIn("Buzón - Isai/Bandeja de entrada/Proyectos", carpetas)
        self.assertIn("Buzón - Isai/Bandeja de entrada/Proyectos/Cliente Norte", carpetas)
        self.assertEqual(len(self.db.buscar(texto="contrato firmado")), 1)

    def test_03_omite_elementos_eliminados(self):
        self._indexar()
        self.assertEqual(len(self.db.buscar(texto="No debe indexarse")), 0)

    def test_04_ignora_lo_que_no_es_correo(self):
        self._indexar()
        self.assertEqual(len(self.db.buscar(texto="Junta de equipo")), 0)

    def test_05_html_se_convierte_en_texto(self):
        self._indexar()
        r = self.db.buscar(texto="Promoción")
        self.assertEqual(len(r), 1)
        cuerpo = self.db.por_id(r[0]["id"])["cuerpo"]
        self.assertIn("Hola Isai", cuerpo)
        self.assertNotIn("<", cuerpo)
        self.assertNotIn("color:red", cuerpo)
        self.assertNotIn("x()", cuerpo)

    def test_06_recupera_cuerpo_por_propiedad_mapi(self):
        self._indexar()
        r = self.db.buscar(texto="propiedad MAPI")
        self.assertEqual(len(r), 1)

    def test_07_ultimo_recurso_vista_previa(self):
        self._indexar()
        self.assertEqual(len(self.db.buscar(texto="Vista previa breve")), 1)

    def test_08_todos_terminan_con_contenido(self):
        """Requisito del usuario: que se vea el contenido de todos."""
        self._indexar()
        self.assertEqual(self.db.total_con_cuerpo(), self.db.total(),
                         "ningun correo debe quedar sin contenido")

    def test_09_fechas_normalizadas_y_ordenables(self):
        self._indexar()
        f = [c["fecha"] for c in self.db.recientes(20)]
        self.assertTrue(all(len(x) == 19 and x[4] == "-" for x in f), f)
        self.assertEqual(f, sorted(f, reverse=True))

    def test_10_segunda_pasada_no_duplica(self):
        self._indexar()
        antes = self.db.total()
        self._indexar()
        self.assertEqual(self.db.total(), antes, "reindexar no debe duplicar")

    def test_11_segunda_pasada_es_incremental(self):
        self._indexar()
        self.avisos.clear()
        n = self._indexar()
        self.assertEqual(n, 0, "sin correos nuevos no debe reescribir nada")

    def test_12_detecta_correos_nuevos(self):
        self._indexar()
        ns = outlook_de_prueba()
        ns.Folders[0].Folders[0].Items._i.append(
            FakeItem("E99", "Correo nuevo de hoy", body="Contenido reciente"))
        instalar_falso(ns)
        n = self._indexar()
        self.assertEqual(n, 1)
        self.assertEqual(len(self.db.buscar(texto="Contenido reciente")), 1)

    def test_13_se_puede_cancelar(self):
        idx = ix.Indexador(self.db, cancelado=lambda: True)
        idx.ejecutar()
        self.assertEqual(self.db.total(), 0)

    def test_14_reporta_avance(self):
        self._indexar()
        self.assertTrue(any("carpetas" in a for a in self.avisos))
        self.assertTrue(any("Listo" in a for a in self.avisos))

    def test_15_guarda_adjuntos_y_destinatarios(self):
        self._indexar()
        r = self.db.buscar(texto="12,450")
        c = self.db.por_id(r[0]["id"])
        self.assertEqual(c["adjuntos"], 1)
        self.assertEqual(c["destinatarios"], "isai@x.com")
        self.assertEqual(c["correo_rem"], "maria@x.com")

    def test_16_entry_id_permite_abrir_en_outlook(self):
        self._indexar()
        r = self.db.buscar(texto="12,450")
        self.assertEqual(self.db.por_id(r[0]["id"])["entry_id"], "E1")

    def test_17_sin_outlook_da_mensaje_claro(self):
        ix.outlook_disponible = lambda: (False, "Esta función necesita Windows con Outlook instalado.")
        with self.assertRaises(RuntimeError) as e:
            ix.Indexador(self.db).ejecutar()
        self.assertIn("Windows", str(e.exception))


class EtiquetaTiendaTests(unittest.TestCase):
    """Dos .pst con el mismo nombre deben poder distinguirse."""

    def _tienda(self, nombre, archivo=None):
        return FakeFolder(nombre, archivo=archivo)

    def test_un_solo_almacen_deja_el_nombre_limpio(self):
        t = self._tienda("Buzón", r"C:\Datos\buzon.pst")
        ns = FakeNS([t])
        self.assertEqual(ix.etiqueta_tienda(t, ns), "Buzón")

    def test_dos_iguales_se_distinguen_por_su_carpeta(self):
        a = self._tienda("correo@x.com", r"C:\Users\I\Archivos de Outlook\correo.pst")
        b = self._tienda("correo@x.com", r"C:\Users\I\Indexer\pst-file\correo.pst")
        ns = FakeNS([a, b])
        self.assertEqual(ix.etiqueta_tienda(a, ns), "correo@x.com (Archivos de Outlook)")
        self.assertEqual(ix.etiqueta_tienda(b, ns), "correo@x.com (pst-file)")

    def test_nombres_distintos_no_se_tocan(self):
        a = self._tienda("Buzón principal", r"C:\a\uno.pst")
        b = self._tienda("Archivo 2019", r"C:\b\dos.pst")
        ns = FakeNS([a, b])
        self.assertEqual(ix.etiqueta_tienda(a, ns), "Buzón principal")
        self.assertEqual(ix.etiqueta_tienda(b, ns), "Archivo 2019")

    def test_sin_ruta_de_archivo_no_truena(self):
        a = self._tienda("Buzón"); b = self._tienda("Buzón")
        ns = FakeNS([a, b])
        self.assertEqual(ix.etiqueta_tienda(a, ns), "Buzón")

    def test_las_carpetas_indexadas_llevan_la_etiqueta(self):
        ruta = os.path.join(tempfile.mkdtemp(), "correos.db")
        db = BaseCorreos(ruta)
        ns = outlook_de_prueba()
        ns.Folders.append(FakeFolder("Buzón - Isai", sub=[
            FakeFolder("Bandeja de entrada", [
                FakeItem("Z1", "Correo del archivo", body="viene de la copia")])],
            archivo=r"C:\copia\buzon.pst"))
        ns.Folders[0].Store = FakeStore(r"C:\principal\buzon.pst")
        instalar_falso(ns)
        ix.Indexador(db).ejecutar()
        carpetas = {c for c, _ in db.carpetas()}
        self.assertTrue(any("(principal)" in c for c in carpetas), carpetas)
        self.assertTrue(any("(copia)" in c for c in carpetas), carpetas)
        db.cerrar()


class MontarPstTests(unittest.TestCase):
    """Montar un .pst en Outlook en vez de depender de libpff."""

    def setUp(self):
        self.ns = outlook_de_prueba()
        instalar_falso(self.ns)
        self.dir = tempfile.mkdtemp()
        self.pst = os.path.join(self.dir, "archivo_2019.pst")
        with open(self.pst, "wb") as f: f.write(b"!BDN" + b"\0" * 600)
        self.ruta = tempfile.mktemp(suffix=".db")
        self.db = BaseCorreos(self.ruta)

    def tearDown(self):
        self.db.cerrar()
        for s in ("", "-wal", "-shm"):
            try: os.remove(self.ruta + s)
            except OSError: pass

    def test_01_monta_un_pst_no_abierto(self):
        estaba, nombre = ix.montar_pst(self.ns, self.pst)
        self.assertFalse(estaba, "no estaba montado antes")
        self.assertEqual(nombre, "archivo_2019.pst")
        self.assertIn(os.path.normcase(self.pst), ix.pst_montados(self.ns))

    def test_02_montar_dos_veces_no_duplica(self):
        ix.montar_pst(self.ns, self.pst)
        antes = len(self.ns.Folders)
        estaba, _ = ix.montar_pst(self.ns, self.pst)
        self.assertTrue(estaba, "debe avisar que ya estaba montado")
        self.assertEqual(len(self.ns.Folders), antes)

    def test_03_su_contenido_se_indexa(self):
        ix.montar_pst(self.ns, self.pst)
        ix.Indexador(self.db).ejecutar()
        r = self.db.buscar(texto="PST archivado")
        self.assertEqual(len(r), 1)
        self.assertIn("Contenido dentro del PST", self.db.por_id(r[0]["id"])["cuerpo"])

    def test_04_desmontar_no_borra_el_archivo(self):
        ix.montar_pst(self.ns, self.pst)
        self.assertTrue(ix.desmontar_pst(self.ns, self.pst))
        self.assertNotIn(os.path.normcase(self.pst), ix.pst_montados(self.ns))
        self.assertTrue(os.path.exists(self.pst), "el .pst debe seguir en el disco")

    def test_05_archivo_inexistente(self):
        with self.assertRaises(RuntimeError) as e:
            ix.montar_pst(self.ns, os.path.join(self.dir, "no_existe.pst"))
        self.assertIn("No existe", str(e.exception))

    def test_06_extension_incorrecta(self):
        otro = os.path.join(self.dir, "cosa.txt")
        with open(otro, "w") as f: f.write("x")
        with self.assertRaises(RuntimeError) as e:
            ix.montar_pst(self.ns, otro)
        self.assertIn(".pst", str(e.exception))

    def test_07_archivo_danado_da_mensaje_util(self):
        roto = os.path.join(self.dir, "roto.pst")
        with open(roto, "wb") as f: f.write(b"basura")
        with self.assertRaises(RuntimeError) as e:
            ix.montar_pst(self.ns, roto)
        self.assertIn("Outlook no pudo abrir", str(e.exception))


if __name__ == "__main__":
    unittest.main(verbosity=2)
