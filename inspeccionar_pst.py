# -*- coding: utf-8 -*-
"""Inspecciona archivos .pst sin abrirlos con Outlook.
Ejecuta:  python inspeccionar_pst.py [carpeta]
Reporta formato, tamano y si Outlook ya lo tiene montado."""
import os
import struct
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# wVer en la cabecera PST -> formato del archivo
FORMATOS = {
    14: ("ANSI (Outlook 97-2002)", "limite de 2 GB, formato antiguo"),
    15: ("ANSI (Outlook 97-2002)", "limite de 2 GB, formato antiguo"),
    21: ("UNICODE (Outlook 2003+)", "formato moderno"),
    23: ("UNICODE (Outlook 2003+)", "formato moderno"),
    36: ("UNICODE con paginas de 4K", "Outlook 2013+"),
}


def analizar(ruta):
    info = {"ruta": ruta, "bytes": os.path.getsize(ruta)}
    with open(ruta, "rb") as f:
        cab = f.read(568)
    if len(cab) < 512 or cab[:4] != b"!BDN":
        info["valido"] = False
        info["nota"] = "No tiene la firma '!BDN': no parece un PST/OST valido."
        return info
    info["valido"] = True
    ver = struct.unpack_from("<H", cab, 10)[0]
    nombre, detalle = FORMATOS.get(ver, (f"desconocido (wVer={ver})", ""))
    info["formato"] = nombre
    info["detalle"] = detalle
    info["cifrado"] = {0: "sin cifrar", 1: "permutacion", 2: "ciclico"}.get(
        cab[513] if len(cab) > 513 else -1, "desconocido")
    return info


def main():
    carpeta = sys.argv[1] if len(sys.argv) > 1 else "."
    print("\n" + "=" * 66)
    print("  INSPECCION DE ARCHIVOS PST")
    print("=" * 66)
    print(f"\n  Carpeta: {os.path.abspath(carpeta)}\n")

    if not os.path.isdir(carpeta):
        print(f"  No existe la carpeta: {carpeta}\n")
        return

    encontrados = []
    for raiz, _, archivos in os.walk(carpeta):
        for a in archivos:
            if a.lower().endswith((".pst", ".ost")):
                encontrados.append(os.path.join(raiz, a))

    if not encontrados:
        print("  No se encontro ningun archivo .pst ni .ost aqui.\n")
        return

    for r in encontrados:
        try:
            i = analizar(r)
        except Exception as e:
            print(f"  {os.path.basename(r)}: no se pudo leer ({e})\n")
            continue
        mb = i["bytes"] / 1024 / 1024
        print(f"  Archivo : {os.path.basename(r)}")
        print(f"  Tamano  : {mb:,.1f} MB".replace(",", " "))
        if not i["valido"]:
            print(f"  AVISO   : {i['nota']}\n")
            continue
        print(f"  Formato : {i['formato']}  ({i['detalle']})")
        print(f"  Cifrado : {i['cifrado']}")
        print()

    print("-" * 66)
    print("\n  LECTURA DIRECTA (sin Outlook)\n")
    try:
        import pypff
        print("    libpff disponible: se puede leer el PST sin Outlook.")
        for r in encontrados:
            try:
                pst = pypff.file()
                pst.open(r)
                raiz = pst.get_root_folder()

                def contar(c, prof=0):
                    n = c.get_number_of_sub_messages()
                    for k in range(c.get_number_of_sub_folders()):
                        n += contar(c.get_sub_folder(k), prof + 1)
                    return n

                print(f"    {os.path.basename(r)}: {contar(raiz):,} mensajes"
                      .replace(",", " "))
                pst.close()
            except Exception as e:
                print(f"    {os.path.basename(r)}: error al abrir ({e})")
    except ImportError:
        print("    libpff NO instalado.")
        print("    Para probarlo:  python -m pip install libpff-python")

    print("\n" + "-" * 66)
    print("\n  QUE VE OUTLOOK AHORA MISMO\n")
    if sys.platform != "win32":
        print("    (solo se puede comprobar en Windows)\n")
        return
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        rutas_montadas = []
        for t in ns.Folders:
            try:
                ruta_archivo = ""
                try:
                    ruta_archivo = str(t.Store.FilePath or "")
                except Exception:
                    pass
                n = 0
                try:
                    n = t.Folders.Count
                except Exception:
                    pass
                print(f"    Almacen: {t.Name}   ({n} carpetas de primer nivel)")
                if ruta_archivo:
                    print(f"             archivo: {ruta_archivo}")
                    rutas_montadas.append(os.path.normcase(ruta_archivo))
            except Exception:
                continue
        print()
        for r in encontrados:
            montado = os.path.normcase(os.path.abspath(r)) in rutas_montadas
            estado = "SI esta montado en Outlook" if montado else "NO esta montado en Outlook"
            print(f"    {os.path.basename(r)}: {estado}")
    except Exception as e:
        print(f"    No se pudo consultar Outlook: {e}")
    print()


if __name__ == "__main__":
    main()
