"""
Lee los correos de Outlook y los guarda en la base de datos.
Recorre TODAS las carpetas (incluidas subcarpetas) y extrae el cuerpo completo.
"""
import os
import re
import sys
from datetime import datetime
from html import unescape

from motor_busqueda import normalizar_fecha

CLASE_CORREO = 43           # olMail
PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_BODY = "http://schemas.microsoft.com/mapi/proptag/0x1000001F"

CARPETAS_OMITIDAS = {"elementos eliminados", "deleted items",
                     "correo no deseado", "junk email", "junk e-mail",
                     "fuentes rss", "rss feeds", "conversation history",
                     "historial de conversaciones", "sync issues",
                     "problemas de sincronizacion", "problemas de sincronización"}


def outlook_disponible():
    if sys.platform != "win32":
        return False, "Esta función necesita Windows con Outlook instalado."
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "Falta la librería pywin32. Instálala con:  python -m pip install pywin32"
    return True, ""


def html_a_texto(html):
    """Convierte HTML de correo en texto legible (sin etiquetas ni basura)."""
    if not html:
        return ""
    t = re.sub(r"(?is)<(script|style|head)[^>]*>.*?</\1>", " ", html)
    t = re.sub(r"(?is)<!--.*?-->", " ", t)
    t = re.sub(r"(?i)<br\s*/?>", "\n", t)
    t = re.sub(r"(?i)</(p|div|tr|li|h[1-6]|table)\s*>", "\n", t)
    t = re.sub(r"(?i)</t[dh]\s*>", "\t", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = unescape(t)
    t = t.replace(" ", " ").replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n[ \t]+", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def _prop(item, esquema):
    try:
        v = item.PropertyAccessor.GetProperty(esquema)
        return str(v).strip() if v else ""
    except Exception:
        return ""


def extraer_cuerpo(item):
    """Cuatro intentos para no dejar ningun correo sin contenido."""
    try:
        b = item.Body
        if b and str(b).strip():
            return str(b).strip()
    except Exception:
        pass
    try:
        h = item.HTMLBody
        if h:
            t = html_a_texto(str(h))
            if t:
                return t
    except Exception:
        pass
    t = _prop(item, PR_BODY)
    if t:
        return t
    try:
        # Ultimo recurso: vista previa que Outlook ya tiene calculada
        return str(getattr(item, "Preview", "") or "").strip()
    except Exception:
        return ""


def extraer_fecha(item):
    for attr in ("ReceivedTime", "SentOn", "CreationTime", "LastModificationTime"):
        try:
            v = getattr(item, attr, None)
            if v is None:
                continue
            if isinstance(v, datetime):
                return v.strftime("%Y-%m-%d %H:%M:%S")
            f = normalizar_fecha(str(v))
            if f:
                return f
        except Exception:
            continue
    return ""


def _texto(item, attr):
    try:
        v = getattr(item, attr, "")
        return str(v).strip() if v else ""
    except Exception:
        return ""


def leer_correo(item, carpeta):
    if _texto(item, "Class") not in ("", str(CLASE_CORREO)):
        try:
            if int(item.Class) != CLASE_CORREO:
                return None
        except Exception:
            pass
    correo_rem = _prop(item, PR_SMTP) or _texto(item, "SenderEmailAddress")
    if correo_rem.startswith("/"):          # direccion interna de Exchange, poco util
        correo_rem = _prop(item, PR_SMTP) or ""
    try:
        adj = int(item.Attachments.Count)
    except Exception:
        adj = 0
    return {
        "entry_id": _texto(item, "EntryID") or None,
        "carpeta": carpeta,
        "remitente": _texto(item, "SenderName") or correo_rem,
        "correo_rem": correo_rem,
        "destinatarios": _texto(item, "To"),
        "asunto": _texto(item, "Subject") or "(sin asunto)",
        "cuerpo": extraer_cuerpo(item),
        "fecha": extraer_fecha(item),
        "adjuntos": adj,
    }


def _recolectar_carpetas(carpetas, acumulado, ruta=""):
    for c in carpetas:
        try:
            nombre = str(c.Name)
        except Exception:
            continue
        if nombre.lower() in CARPETAS_OMITIDAS:
            continue
        completo = f"{ruta}/{nombre}" if ruta else nombre
        acumulado.append((completo, c))
        try:
            _recolectar_carpetas(c.Folders, acumulado, completo)
        except Exception:
            pass
    return acumulado


OL_UNICODE = 3          # olStoreUnicode


def _ruta_de(tienda):
    try:
        return os.path.normcase(os.path.abspath(str(tienda.Store.FilePath or "")))
    except Exception:
        return ""


def pst_montados(ns):
    """Rutas de los archivos .pst que Outlook tiene abiertos ahora mismo."""
    rutas = {}
    for t in ns.Folders:
        r = _ruta_de(t)
        if r:
            rutas[r] = t
    return rutas


def montar_pst(ns, ruta):
    """Abre un archivo .pst dentro de Outlook para poder leerlo.
    Es lo mismo que hacer Archivo > Abrir > Archivo de datos de Outlook.
    Devuelve (ya_estaba_montado, nombre_del_almacen)."""
    ruta = os.path.abspath(ruta)
    if not os.path.exists(ruta):
        raise RuntimeError(f"No existe el archivo:\n{ruta}")
    if not ruta.lower().endswith(".pst"):
        raise RuntimeError("El archivo debe tener extensión .pst")

    clave = os.path.normcase(ruta)
    montados = pst_montados(ns)
    if clave in montados:
        return True, str(montados[clave].Name)

    try:
        ns.AddStoreEx(ruta, OL_UNICODE)
    except Exception as e:
        raise RuntimeError(
            "Outlook no pudo abrir el archivo.\n\n"
            "Suele deberse a que otro programa lo tiene en uso, a que está en "
            "una carpeta protegida, o a que el archivo está dañado.\n\n"
            f"Detalle: {e}")

    nuevos = pst_montados(ns)
    if clave not in nuevos:
        raise RuntimeError("Outlook aceptó el archivo pero no aparece en la lista.")
    return False, str(nuevos[clave].Name)


def desmontar_pst(ns, ruta):
    """Cierra el .pst en Outlook. No borra nada del disco."""
    clave = os.path.normcase(os.path.abspath(ruta))
    montados = pst_montados(ns)
    if clave not in montados:
        return False
    try:
        ns.RemoveStore(montados[clave])
        return True
    except Exception:
        return False


class Indexador:
    """Extrae correos de Outlook hacia la base. Reporta avance y se puede cancelar."""

    def __init__(self, base, progreso=None, cancelado=None):
        self.base = base
        self.progreso = progreso or (lambda *a, **k: None)
        self.cancelado = cancelado or (lambda: False)

    def _avisar(self, texto, hechos=0, total=0):
        self.progreso(texto, hechos, total)

    def ejecutar(self, solo_nuevos=True, lote=200):
        ok, motivo = outlook_disponible()
        if not ok:
            raise RuntimeError(motivo)

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()          # imprescindible al indexar en segundo plano
        try:
            self._avisar("Conectando con Outlook...")
            try:
                app = win32com.client.Dispatch("Outlook.Application")
                ns = app.GetNamespace("MAPI")
            except Exception as e:
                raise RuntimeError(
                    "No se pudo conectar con Outlook. Ábrelo y vuelve a intentar.\n\n"
                    f"Detalle: {e}")

            self._avisar("Buscando carpetas...")
            carpetas = []
            for tienda in ns.Folders:
                try:
                    _recolectar_carpetas(tienda.Folders, carpetas, str(tienda.Name))
                except Exception:
                    continue
            if not carpetas:
                raise RuntimeError("Outlook no devolvió ninguna carpeta de correo.")

            conocidos = self.base.entry_ids_existentes() if solo_nuevos else set()

            total = 0
            for c in carpetas:
                try:
                    total += int(c[1].Items.Count)
                except Exception:
                    pass
            self._avisar(f"{len(carpetas)} carpetas, {total} correos en total.", 0, total)

            hechos = guardados = 0
            for nombre, carpeta in carpetas:
                if self.cancelado():
                    break
                try:
                    items = carpeta.Items
                    items.Sort("[ReceivedTime]", True)
                except Exception:
                    continue

                buffer = []
                try:
                    item = items.GetFirst()
                except Exception:
                    continue

                while item is not None:
                    if self.cancelado():
                        break
                    hechos += 1
                    try:
                        eid = str(item.EntryID)
                        if not (solo_nuevos and eid in conocidos):
                            datos = leer_correo(item, nombre)
                            if datos and datos["entry_id"]:
                                buffer.append(datos)
                    except Exception:
                        pass

                    if len(buffer) >= lote:
                        guardados += self.base.guardar(buffer)
                        buffer.clear()
                    if hechos % 25 == 0:
                        self._avisar(f"{nombre}  —  {guardados} guardados", hechos, total)
                    try:
                        item = items.GetNext()
                    except Exception:
                        break

                if buffer:
                    guardados += self.base.guardar(buffer)
                self._avisar(f"{nombre}  —  {guardados} guardados", hechos, total)

            self._avisar(f"Listo. {guardados} correos actualizados.", hechos, total)
            return guardados
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
