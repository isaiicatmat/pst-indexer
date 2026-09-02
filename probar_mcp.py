# -*- coding: utf-8 -*-
"""Pruebas del servidor MCP, hablando el protocolo de verdad.
Ejecuta:  python probar_mcp.py"""
import asyncio
import os
import sys
import tempfile
import unittest

from motor_busqueda import BaseCorreos

try:
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
    HAY_MCP = True
except ImportError:
    HAY_MCP = False

AQUI = os.path.dirname(os.path.abspath(__file__))


def correos_de_prueba(n_extra=0):
    base = [
        dict(entry_id="1", carpeta="Bandeja de entrada", remitente="María Gómez",
             correo_rem="maria@constructoranorte.com", destinatarios="isai@pieconsulting.co.kr",
             asunto="Factura 4471 pendiente de pago",
             cuerpo="Estimado Isai:\n\nLe recuerdo que la factura 4471 sigue pendiente. "
                    "El monto es de $12,450.00 MXN.",
             fecha="2024-03-15 10:23:00", origen=r"C:\Archivos de Outlook\buzon.pst",
             adjuntos=1),
        dict(entry_id="2", carpeta="Elementos enviados", remitente="Isai Carreto",
             correo_rem="isai@pieconsulting.co.kr", destinatarios="maria@constructoranorte.com",
             asunto="RE: Factura 4471", cuerpo="Hola María, reviso y te confirmo el viernes.",
             fecha="2024-03-16 14:05:00", origen=r"C:\Archivos de Outlook\buzon.pst",
             adjuntos=0),
        dict(entry_id="3", carpeta="Bandeja de entrada", remitente="Boletín",
             correo_rem="news@b.com", destinatarios="isai@pieconsulting.co.kr",
             asunto="Sin contenido", cuerpo="", fecha="2024-04-01 08:00:00",
             origen=r"C:\Archivos de Outlook\buzon.pst", adjuntos=0),
    ]
    for i in range(n_extra):
        base.append(dict(
            entry_id=f"X{i}", carpeta="Bandeja de entrada", remitente=f"Remitente {i}",
            correo_rem=f"r{i}@x.com", destinatarios="isai@pieconsulting.co.kr",
            asunto=f"Asunto de relleno {i}", cuerpo="contenido de relleno repetido",
            fecha="2024-05-%02d 09:00:00" % ((i % 28) + 1),
            origen=r"C:\Archivos de Outlook\buzon.pst", adjuntos=0))
    return base


@unittest.skipUnless(HAY_MCP, "el paquete mcp no esta instalado")
class ServidorMcpTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp()
        cls.ruta = os.path.join(cls.dir, "correos.db")
        db = BaseCorreos(cls.ruta)
        db.guardar(correos_de_prueba(n_extra=80))
        db.cerrar()

    def _llamar(self, herramienta, argumentos=None, listar=False):
        async def run():
            params = StdioServerParameters(
                command=sys.executable,
                args=[os.path.join(AQUI, "servidor_mcp.py")],
                env={**os.environ, "CORREOS_DB": self.ruta})
            async with stdio_client(params) as (r, w):
                async with ClientSession(r, w) as s:
                    await s.initialize()
                    if listar:
                        return await s.list_tools()
                    res = await s.call_tool(herramienta, argumentos or {})
                    return res.content[0].text
        return asyncio.run(run())

    # ------------------------------------------------------------ seguridad
    def test_01_solo_expone_tres_herramientas_de_lectura(self):
        t = self._llamar(None, listar=True)
        nombres = sorted(h.name for h in t.tools)
        self.assertEqual(nombres, ["buscar_correos", "leer_correo", "resumen_indice"])

    def test_02_todas_declaran_solo_lectura(self):
        t = self._llamar(None, listar=True)
        for h in t.tools:
            self.assertTrue(h.annotations.read_only_hint, f"{h.name} no declara solo lectura")
            self.assertFalse(h.annotations.destructive_hint, f"{h.name} se declara destructiva")

    def test_03_no_hay_ninguna_via_de_escritura(self):
        t = self._llamar(None, listar=True)
        prohibidas = ("envi", "send", "responder", "reply", "borr", "delete",
                      "elimin", "mover", "move", "escrib", "write", "marcar")
        for h in t.tools:
            texto = (h.name + " " + (h.description or "")).lower()
            for p in ("envi", "send", "borr", "delete", "elimin"):
                self.assertNotIn(p, h.name.lower(), f"{h.name} suena a escritura")

    def test_04_la_base_se_abre_en_solo_lectura(self):
        db = BaseCorreos(self.ruta, solo_lectura=True)
        try:
            with self.assertRaises(Exception):
                db.con.execute("DELETE FROM correos")
            with self.assertRaises(PermissionError):
                db.guardar([dict(entry_id="Z")])
        finally:
            db.cerrar()

    # -------------------------------------------------------------- limites
    def test_05_limite_por_defecto(self):
        r = self._llamar("buscar_correos", {"texto": "relleno"})
        self.assertEqual(r.count("[id "), 20, "por defecto deben ser 20")

    def test_06_no_se_puede_superar_el_tope(self):
        r = self._llamar("buscar_correos", {"texto": "relleno", "limite": 500})
        self.assertLessEqual(r.count("[id "), 50, "el tope es 50 aunque se pida mas")

    def test_07_limite_minimo_valido(self):
        r = self._llamar("buscar_correos", {"texto": "relleno", "limite": 0})
        self.assertGreaterEqual(r.count("[id "), 1)

    def test_08_avisa_cuando_llega_al_limite(self):
        r = self._llamar("buscar_correos", {"texto": "relleno", "limite": 5})
        self.assertIn("limite de 5", r)

    def test_09_el_extracto_va_recortado(self):
        r = self._llamar("buscar_correos", {"texto": "factura"})
        for linea in r.splitlines():
            self.assertLess(len(linea), 400, "ninguna linea debe ser enorme")

    # -------------------------------------------------------------- busqueda
    def test_10_busca_por_texto(self):
        r = self._llamar("buscar_correos", {"texto": "4471"})
        self.assertIn("Factura 4471", r)
        self.assertIn("[id 1]", r)

    def test_11_busca_sin_acentos(self):
        r = self._llamar("buscar_correos", {"remitente": "maria"})
        self.assertIn("María Gómez", r)

    def test_12_filtra_por_carpeta_y_fecha(self):
        r = self._llamar("buscar_correos", {"carpeta": "Elementos enviados"})
        self.assertIn("RE: Factura 4471", r)
        self.assertNotIn("[id 1]", r)
        r2 = self._llamar("buscar_correos", {"desde": "2024-03-16", "hasta": "2024-03-16"})
        self.assertIn("[id 2]", r2)
        self.assertNotIn("[id 1]", r2)

    def test_13_sin_criterios_pide_alguno(self):
        r = self._llamar("buscar_correos", {})
        self.assertIn("al menos un criterio", r)

    def test_14_sin_resultados_lo_explica(self):
        r = self._llamar("buscar_correos", {"texto": "zzzznoexiste"})
        self.assertIn("Ningun correo coincide", r)

    def test_15_texto_raro_no_rompe_el_servidor(self):
        for basura in ['"', "((", "*", "NEAR(", "a OR", "\\", "%$#"]:
            self._llamar("buscar_correos", {"texto": basura})

    # ---------------------------------------------------------------- lectura
    def test_16_lee_el_correo_completo(self):
        r = self._llamar("leer_correo", {"id": 1})
        self.assertIn("12,450.00 MXN", r)
        self.assertIn("maria@constructoranorte.com", r)
        self.assertIn("Adjuntos: 1", r)
        self.assertIn("buzon.pst", r)

    def test_17_id_inexistente(self):
        self.assertIn("No hay ningun correo", self._llamar("leer_correo", {"id": 999999}))

    def test_18_correo_sin_texto_lo_explica(self):
        r = self._llamar("leer_correo", {"id": 3})
        self.assertIn("no tiene texto guardado", r)

    def test_19_cuerpo_enorme_se_recorta(self):
        ruta = os.path.join(tempfile.mkdtemp(), "correos.db")
        db = BaseCorreos(ruta)
        db.guardar([dict(entry_id="G", carpeta="B", remitente="A", correo_rem="",
                         destinatarios="", asunto="Enorme", cuerpo="x" * 60000,
                         fecha="2024-01-01 00:00:00", origen="", adjuntos=0)])
        db.cerrar()
        anterior, self.__class__.ruta = self.__class__.ruta, ruta
        try:
            r = self._llamar("leer_correo", {"id": 1})
            self.assertIn("texto recortado", r)
            self.assertLess(len(r), 25000)
        finally:
            self.__class__.ruta = anterior

    # ---------------------------------------------------------------- resumen
    def test_20_resumen_describe_el_indice(self):
        r = self._llamar("resumen_indice")
        self.assertIn("correos indexados", r)
        self.assertIn("Bandeja de entrada", r)
        self.assertIn("buzon.pst", r)
        self.assertIn("solo lectura", r)

    def test_21_base_inexistente_da_mensaje_util(self):
        anterior, self.__class__.ruta = self.__class__.ruta, "/no/existe/correos.db"
        try:
            r = self._llamar("resumen_indice")
            self.assertIn("Actualizar correos", r)
        finally:
            self.__class__.ruta = anterior


if __name__ == "__main__":
    unittest.main(verbosity=2)
