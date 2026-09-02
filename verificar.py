# -*- coding: utf-8 -*-
"""Revision rapida del sistema. Ejecuta:  python verificar.py
Dice en lenguaje claro que funciona y que falta."""
import os
import sys

# La consola de Windows usa cp1252 y truena con acentos: forzamos UTF-8.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OK, MAL, AVISO = "[OK]  ", "[FALLA]", "[AVISO]"
problemas = []


def linea(estado, texto):
    print(f"  {estado} {texto}")


print("\n" + "=" * 62)
print("  REVISION DEL BUSCADOR DE CORREOS")
print("=" * 62 + "\n")

v = sys.version_info
if v >= (3, 8):
    linea(OK, f"Python {v.major}.{v.minor}.{v.micro}")
else:
    linea(MAL, f"Python {v.major}.{v.minor} es muy antiguo (se necesita 3.8 o superior)")
    problemas.append("Instala una version reciente de Python.")

try:
    import PyQt5  # noqa: F401
    linea(OK, "PyQt5 (la ventana de la aplicacion)")
except ImportError:
    linea(MAL, "Falta PyQt5")
    problemas.append("Ejecuta:  python -m pip install PyQt5")

import sqlite3
try:
    sqlite3.connect(":memory:").execute("CREATE VIRTUAL TABLE t USING fts5(x)")
    linea(OK, f"Busqueda instantanea FTS5 (SQLite {sqlite3.sqlite_version})")
except sqlite3.OperationalError:
    linea(AVISO, "Sin FTS5: la busqueda funciona pero mas lenta")

if sys.platform == "win32":
    try:
        import win32com.client  # noqa: F401
        linea(OK, "pywin32 (conexion con Outlook)")
        try:
            import pythoncom
            pythoncom.CoInitialize()
            win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            linea(OK, "Outlook responde correctamente")
        except Exception as e:
            linea(AVISO, f"Outlook no respondio ({str(e)[:60]})")
            problemas.append("Abre Outlook antes de pulsar 'Actualizar correos'.")
    except ImportError:
        linea(MAL, "Falta pywin32")
        problemas.append("Ejecuta:  python -m pip install pywin32")
else:
    linea(AVISO, f"Sistema {sys.platform}: se puede buscar, pero solo Windows lee Outlook")

aqui = os.path.dirname(os.path.abspath(__file__))
for f in ("buscador_correos.py", "motor_busqueda.py", "indexador_outlook.py"):
    if os.path.exists(os.path.join(aqui, f)):
        linea(OK, f"Archivo {f}")
    else:
        linea(MAL, f"Falta el archivo {f}")
        problemas.append(f"Vuelve a copiar {f} a la carpeta.")

db = os.path.join(aqui, "correos.db")
if os.path.exists(db):
    try:
        sys.path.insert(0, aqui)
        from motor_busqueda import BaseCorreos
        b = BaseCorreos(db)
        t, c = b.total(), b.total_con_cuerpo()
        pct = (100 * c / t) if t else 0
        linea(OK, f"Base de datos: {t:,} correos, {c:,} con contenido ({pct:.0f}%)")
        if t and pct < 70:
            problemas.append("Muchos correos sin contenido: pulsa 'Actualizar correos'.")
        b.cerrar()
    except Exception as e:
        linea(MAL, f"La base de datos dio un error: {e}")
        problemas.append("Borra correos.db y vuelve a indexar.")
else:
    linea(AVISO, "Todavia no hay correos indexados (normal la primera vez)")

print("\n" + "-" * 62)
if problemas:
    print("\n  QUE HACER:\n")
    for i, p in enumerate(problemas, 1):
        print(f"    {i}. {p}")
else:
    print("\n  Todo en orden. Abre la aplicacion con INICIAR.bat\n")
print()
