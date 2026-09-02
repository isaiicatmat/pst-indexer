# -*- coding: utf-8 -*-
"""Ejecuta TODAS las pruebas del sistema.  python probar_todo.py"""
import os, sys, unittest
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.*=false")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

MODULOS = [
    ("Motor de busqueda y base de datos", "probar_sistema"),
    ("Indexador de Outlook (simulado)",   "probar_outlook"),
    ("Interfaz de usuario",               "probar_interfaz"),
    ("Flujo completo de principio a fin", "probar_completo"),
]

print("\n" + "=" * 64)
print("  PRUEBAS DEL BUSCADOR DE CORREOS")
print("=" * 64)

total = fallos = 0
for titulo, mod in MODULOS:
    print(f"\n  {titulo}")
    print("  " + "-" * (len(titulo)))
    suite = unittest.defaultTestLoader.loadTestsFromName(mod)
    r = unittest.TextTestRunner(verbosity=0, stream=open(os.devnull, "w")).run(suite)
    total += r.testsRun
    fallos += len(r.failures) + len(r.errors)
    estado = "OK" if not (r.failures or r.errors) else "CON FALLOS"
    print(f"    {r.testsRun} pruebas ... {estado}")
    for caso, texto in list(r.failures) + list(r.errors):
        print(f"      FALLA: {caso}")
        print("        " + texto.strip().splitlines()[-1])

print("\n" + "=" * 64)
if fallos:
    print(f"  RESULTADO: {total - fallos} de {total} pruebas pasaron, {fallos} con problemas")
else:
    print(f"  RESULTADO: las {total} pruebas pasaron correctamente")
print("=" * 64 + "\n")
sys.exit(1 if fallos else 0)
