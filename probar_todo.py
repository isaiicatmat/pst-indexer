# -*- coding: utf-8 -*-
"""Ejecuta TODAS las pruebas del sistema.  python probar_todo.py

Cada modulo corre en su propio proceso: las pruebas de la interfaz crean un
bucle de eventos de Qt y las del servidor MCP uno de asyncio, y compartir
proceso los hace chocar. Ademas, si un modulo se cae no arrastra a los demas.
"""
import os
import re
import subprocess
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.*=false")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

AQUI = os.path.dirname(os.path.abspath(__file__))

MODULOS = [
    ("Motor de busqueda y base de datos", "probar_sistema"),
    ("Indexador de Outlook (simulado)",   "probar_outlook"),
    ("Interfaz de usuario",               "probar_interfaz"),
    ("Flujo completo de principio a fin", "probar_completo"),
    ("Servidor MCP (solo lectura)",       "probar_mcp"),
]

print("\n" + "=" * 64)
print("  PRUEBAS DEL BUSCADOR DE CORREOS")
print("=" * 64)

total = fallos = omitidas = 0
for titulo, modulo in MODULOS:
    print(f"\n  {titulo}")
    print("  " + "-" * len(titulo))
    sys.stdout.flush()

    r = subprocess.run([sys.executable, os.path.join(AQUI, modulo + ".py")],
                       capture_output=True, text=True, cwd=AQUI,
                       env={**os.environ}, errors="replace")
    salida = (r.stdout or "") + (r.stderr or "")

    m = re.search(r"^Ran (\d+) test", salida, re.MULTILINE)
    n = int(m.group(1)) if m else 0
    total += n
    saltadas = len(re.findall(r"\.\.\. skipped", salida))
    omitidas += saltadas

    if r.returncode == 0:
        extra = f", {saltadas} omitidas" if saltadas else ""
        print(f"    {n} pruebas ... OK{extra}")
    else:
        detalle = re.findall(r"^(?:FAIL|ERROR): (\S+)", salida, re.MULTILINE)
        fallos += max(len(detalle), 1)
        print(f"    {n} pruebas ... CON FALLOS")
        for d in detalle:
            print(f"      {d}")
        if not detalle:      # se cayo antes de poder informar
            for linea in salida.strip().splitlines()[-6:]:
                print(f"      {linea}")

print("\n" + "=" * 64)
if fallos:
    print(f"  RESULTADO: {total - fallos} de {total} pruebas pasaron, {fallos} con problemas")
else:
    print(f"  RESULTADO: las {total} pruebas pasaron correctamente"
          + (f"  ({omitidas} omitidas)" if omitidas else ""))
print("=" * 64 + "\n")
sys.exit(1 if fallos else 0)
