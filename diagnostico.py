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
from motor_busqueda import BaseCorreos, carpeta_datos, porcentaje
from indexador_outlook import CARPETAS_OMITIDAS, etiqueta_tienda, outlook_disponible

ok, motivo = outlook_disponible()
if not ok:
    print(f"\n  {motivo}\n")
    sys.exit(1)

import pythoncom
import win32com.client

pythoncom.CoInitialize()
ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print("\n" + "=" * 78)
print("  QUE HAY EN OUTLOOK  vs  QUE QUEDO INDEXADO")
print("=" * 78)

filas = []


def recorrer(carpetas, ruta):
    for c in carpetas:
        try:
            nombre = str(c.Name)
        except Exception:
            continue
        completo = f"{ruta}/{nombre}"
        omitida = nombre.lower() in CARPETAS_OMITIDAS
        try:
            total = int(c.Items.Count)      # una sola llamada, sin recorrer
        except Exception:
            total = 0
        filas.append((completo, total, omitida))
        try:
            recorrer(c.Folders, completo)
        except Exception:
            pass


etiquetas = []
for tienda in ns.Folders:
    try:
        etiquetas.append((tienda, etiqueta_tienda(tienda, ns)))
    except Exception:
        pass

for tienda, etiqueta in etiquetas:
    ruta_archivo = ""
    try:
        ruta_archivo = str(tienda.Store.FilePath or "")
    except Exception:
        pass
    print(f"\n  ALMACEN: {etiqueta}")
    if ruta_archivo:
        print(f"           {ruta_archivo}")
    sys.stdout.flush()
    try:
        recorrer(tienda.Folders, etiqueta)
    except Exception as e:
        print(f"           (no se pudo recorrer: {e})")

db = os.path.join(carpeta_datos(), "correos.db")
indexadas, total_db, con_cuerpo = {}, 0, 0
if os.path.exists(db):
    b = BaseCorreos(db)
    indexadas = dict(b.carpetas())
    total_db, con_cuerpo = b.total(), b.total_con_cuerpo()
    b.cerrar()

print(f"\n  {'CARPETA':<50} {'OUTLOOK':>9} {'INDEXADO':>9}")
print("  " + "-" * 72)

suma_out = suma_db = 0
sospechosas = []
sobrantes = []
for ruta, total, omitida in filas:
    corta = ruta if len(ruta) <= 49 else "..." + ruta[-46:]
    if omitida:
        print(f"  {corta:<50} {total:>9} {'(omitida)':>9}")
        continue
    en_base = indexadas.get(ruta, 0)
    suma_out += total
    suma_db += en_base
    marca = ""
    if total and en_base == 0:
        marca = "  <-- NADA INDEXADO"
        sospechosas.append((ruta, total, en_base))
    elif en_base > total:
        marca = "  <-- SOBRAN"
        sobrantes.append((ruta, total, en_base))
    print(f"  {corta:<50} {total:>9} {en_base:>9}{marca}")

print("  " + "-" * 72)
print(f"  {'TOTAL':<50} {suma_out:>9} {suma_db:>9}")
print("\n  OUTLOOK incluye citas, contactos y tareas; INDEXADO solo correos,")
print("  asi que es normal que la primera columna sea algo mayor.")

print(f"\n  Base de datos: {total_db} correos, {con_cuerpo} con contenido "
      f"({porcentaje(con_cuerpo, total_db)})")

huerfanas = set(indexadas) - {r for r, _, o in filas if not o}
if huerfanas:
    print("\n  Carpetas en la base que Outlook ya no ofrece")
    print("  (normal si cerraste un archivo .pst):")
    for c in sorted(huerfanas):
        print(f"    {c}  ({indexadas[c]} correos)")

print("\n" + "=" * 78)
if sospechosas:
    print("\n  CARPETAS CON CORREOS QUE NO SE INDEXARON:\n")
    for ruta, total, _ in sospechosas:
        print(f"    {ruta}: {total} elementos en Outlook, 0 en la base")
    print("\n  Pulsa 'Actualizar correos' en la aplicacion.")
elif not sobrantes:
    print("\n  No hay carpetas con correos sin indexar.")

if sobrantes:
    de_mas = sum(b - o for _, o, b in sobrantes)
    print(f"\n  LA BASE TIENE {de_mas} CORREOS QUE OUTLOOK YA NO OFRECE:\n")
    for ruta, total, en_base in sobrantes:
        print(f"    {ruta}: {total} en Outlook, {en_base} en la base")
    print("\n  Suele significar que se indexo un archivo .pst que despues se")
    print("  cerro en Outlook. Los correos siguen siendo buscables, pero ya no")
    print("  se pueden abrir con 'Abrir en Outlook'.")
    print("\n  Para dejar solo lo que Outlook ofrece ahora:")
    print("  menu ... > 'Vaciar el indice y empezar de cero', y luego")
    print("  'Actualizar correos'.")
print()
