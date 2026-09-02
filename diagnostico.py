# -*- coding: utf-8 -*-
"""Compara lo que hay en Outlook contra lo que quedo indexado, carpeta por carpeta.
Ejecuta:  python diagnostico.py"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from motor_busqueda import BaseCorreos
from indexador_outlook import CARPETAS_OMITIDAS, CLASE_CORREO, outlook_disponible

ok, motivo = outlook_disponible()
if not ok:
    print(f"\n  {motivo}\n")
    sys.exit(1)

import pythoncom
import win32com.client

pythoncom.CoInitialize()
ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print("\n" + "=" * 74)
print("  QUE HAY EN OUTLOOK  vs  QUE QUEDO INDEXADO")
print("=" * 74)
print("\n  Recorriendo Outlook (puede tardar un minuto)...\n")

filas = []


def recorrer(carpetas, ruta=""):
    for c in carpetas:
        try:
            nombre = str(c.Name)
        except Exception:
            continue
        completo = f"{ruta}/{nombre}" if ruta else nombre
        omitida = nombre.lower() in CARPETAS_OMITIDAS
        total = correos = 0
        try:
            total = int(c.Items.Count)
        except Exception:
            pass
        if total and not omitida:
            try:
                it = c.Items.GetFirst()
                while it is not None:
                    try:
                        if int(it.Class) == CLASE_CORREO:
                            correos += 1
                    except Exception:
                        pass
                    it = c.Items.GetNext()
            except Exception:
                pass
        filas.append((completo, total, correos, omitida))
        etiqueta = "(omitida)" if omitida else f"{correos} correos"
        corta = completo if len(completo) <= 52 else "..." + completo[-49:]
        print(f"    {corta:<54} {etiqueta}")
        sys.stdout.flush()
        try:
            recorrer(c.Folders, completo)
        except Exception:
            pass


for tienda in ns.Folders:
    try:
        print(f"  ALMACEN: {tienda.Name}")
        try:
            print(f"           {tienda.Store.FilePath}")
        except Exception:
            pass
        recorrer(tienda.Folders, str(tienda.Name))
    except Exception as e:
        print(f"  (no se pudo leer un almacen: {e})")

db = os.path.join(os.path.dirname(os.path.abspath(__file__)), "correos.db")
indexadas = {}
if os.path.exists(db):
    b = BaseCorreos(db)
    indexadas = dict(b.carpetas())
    total_db, con_cuerpo = b.total(), b.total_con_cuerpo()
    b.cerrar()
else:
    total_db = con_cuerpo = 0

print(f"\n  {'CARPETA':<44} {'ELEMS':>7} {'CORREOS':>8} {'EN BASE':>8}")
print("  " + "-" * 70)

suma_correos = suma_base = 0
faltantes = []
for ruta, total, correos, omitida in filas:
    corta = ruta if len(ruta) <= 43 else "..." + ruta[-40:]
    if omitida:
        print(f"  {corta:<44} {total:>7} {'(omitida)':>17}")
        continue
    en_base = indexadas.get(ruta, 0)
    suma_correos += correos
    suma_base += en_base
    marca = "" if en_base >= correos else "  <-- FALTAN"
    if correos and en_base < correos:
        faltantes.append((ruta, correos, en_base))
    print(f"  {corta:<44} {total:>7} {correos:>8} {en_base:>8}{marca}")

print("  " + "-" * 70)
print(f"  {'TOTAL':<44} {'':>7} {suma_correos:>8} {suma_base:>8}")

print(f"\n  Base de datos: {total_db} correos, {con_cuerpo} con contenido "
      f"({100*con_cuerpo/total_db if total_db else 0:.0f}%)")

carpetas_solo_en_base = set(indexadas) - {r for r, _, _, o in filas if not o}
if carpetas_solo_en_base:
    print("\n  Carpetas que estan en la base pero ya no en Outlook:")
    for c in sorted(carpetas_solo_en_base):
        print(f"    {c}  ({indexadas[c]})")

print("\n" + "=" * 74)
if faltantes:
    print("\n  HAY CORREOS SIN INDEXAR:\n")
    for ruta, correos, en_base in faltantes:
        print(f"    {ruta}: {correos} en Outlook, {en_base} en la base")
    print("\n  Pulsa 'Actualizar correos' en la aplicacion.")
else:
    print("\n  Todo lo que Outlook ofrece esta indexado.")
print()
