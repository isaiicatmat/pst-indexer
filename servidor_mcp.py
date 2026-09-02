# -*- coding: utf-8 -*-
"""
Servidor MCP para consultar el correo ya indexado desde Claude Code.

Solo lectura: la base se abre en modo de solo lectura de SQLite, asi que la
garantia no depende de que este archivo se porte bien. No existe ninguna
herramienta para enviar, responder, mover ni borrar correo.

Arranque:  python servidor_mcp.py
La ruta de la base se toma de la variable CORREOS_DB, o de la carpeta habitual.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from motor_busqueda import BaseCorreos, carpeta_datos, porcentaje

# Cuanto se devuelve como maximo. Evita volcar el buzon entero en una consulta.
LIMITE_POR_DEFECTO = 20
LIMITE_MAXIMO = 50
EXTRACTO = 200
CUERPO_MAXIMO = 20000

SOLO_LECTURA = ToolAnnotations(read_only_hint=True, destructive_hint=False,
                               idempotent_hint=True, open_world_hint=False)

servidor = MCPServer(
    name="correos",
    title="Correos de Outlook",
    instructions=(
        "Consulta el correo de Outlook ya indexado en esta computadora. "
        "Es una copia local de solo lectura: no se puede enviar, responder, "
        "mover ni borrar nada.\n\n"
        "Empieza por `resumen_indice` para saber que hay disponible. "
        "Usa `buscar_correos` para localizar mensajes y `leer_correo` para "
        "obtener el texto completo de uno concreto.\n\n"
        "En `texto` se busca a la vez en remitente, destinatarios, asunto y "
        "cuerpo. Varias palabras exigen que aparezcan todas; entre comillas se "
        "busca la frase exacta. Los acentos no importan."),
)


def _ruta_base():
    return os.environ.get("CORREOS_DB") or os.path.join(carpeta_datos(), "correos.db")


def _abrir():
    ruta = _ruta_base()
    if not os.path.exists(ruta):
        raise FileNotFoundError(
            f"No se encontro la base de correos en:\n{ruta}\n\n"
            "Abre el Buscador de Correos y pulsa «Actualizar correos», o define "
            "la variable CORREOS_DB con la ruta correcta.")
    return BaseCorreos(ruta, solo_lectura=True)


def _con_base(funcion):
    """Ejecuta la consulta y convierte cualquier fallo en texto explicativo.

    Si se deja escapar la excepcion, el cliente solo recibe
    'Error executing tool', y se pierde la explicacion de que hacer.
    """
    try:
        base = _abrir()
    except Exception as e:
        return str(e)
    try:
        return funcion(base)
    except Exception as e:
        return f"No se pudo completar la consulta.\n\n{type(e).__name__}: {e}"
    finally:
        base.cerrar()


def _fecha(valor):
    return (valor or "")[:16]


@servidor.tool(
    title="Buscar correos",
    description=(
        "Busca en el correo indexado por texto libre, remitente, carpeta o rango "
        "de fechas. Devuelve una lista breve con un extracto de cada mensaje; "
        "para el texto completo usa despues `leer_correo` con el id."),
    annotations=SOLO_LECTURA,
)
def buscar_correos(
    texto: str = "",
    remitente: str = "",
    desde: str = "",
    hasta: str = "",
    carpeta: str = "",
    limite: int = LIMITE_POR_DEFECTO,
) -> str:
    """Args:
    texto: palabras a buscar en remitente, asunto y cuerpo. Entre comillas, frase exacta.
    remitente: filtra por nombre o direccion del remitente.
    desde: fecha inicial inclusive, formato AAAA-MM-DD.
    hasta: fecha final inclusive, formato AAAA-MM-DD.
    carpeta: nombre exacto de la carpeta, tal como aparece en `resumen_indice`.
    limite: cuantos correos devolver como maximo (tope 50).
    """
    limite = max(1, min(int(limite or LIMITE_POR_DEFECTO), LIMITE_MAXIMO))
    if not any([texto.strip(), remitente.strip(), desde.strip(),
                hasta.strip(), carpeta.strip()]):
        return ("Indica al menos un criterio: texto, remitente, carpeta o fechas. "
                "Para ver que hay disponible usa `resumen_indice`.")

    filas = _con_base(lambda base: base.buscar(
        texto=texto.strip(), remitente=remitente.strip(), desde=desde.strip(),
        hasta=hasta.strip(), carpeta=carpeta.strip(), limite=limite))
    if isinstance(filas, str):
        return filas

    if not filas:
        return ("Ningun correo coincide. Prueba con menos palabras, o revisa el "
                "rango de fechas y el nombre de la carpeta.")

    partes = [f"{len(filas)} correo(s). Usa `leer_correo` con el id para el texto completo.\n"]
    for f in filas:
        extracto = " ".join((f.get("extracto") or "").split())[:EXTRACTO]
        partes.append(
            f"[id {f['id']}] {_fecha(f['fecha'])}  ·  {f['remitente']}\n"
            f"  Asunto: {f['asunto']}\n"
            f"  Carpeta: {f['carpeta']}"
            + ("  ·  con adjuntos" if f.get("adjuntos") else "")
            + (f"\n  {extracto}…" if extracto else ""))
    if len(filas) == limite:
        partes.append(f"\n(Se alcanzo el limite de {limite}. Afina la busqueda "
                      f"o sube `limite` hasta {LIMITE_MAXIMO}.)")
    return "\n\n".join(partes)


@servidor.tool(
    title="Leer un correo",
    description="Devuelve el texto completo de un correo a partir del id que dio `buscar_correos`.",
    annotations=SOLO_LECTURA,
)
def leer_correo(id: int) -> str:
    """Args:
    id: identificador del correo, tal como aparece entre corchetes en los resultados.
    """
    c = _con_base(lambda base: base.por_id(int(id)))
    if isinstance(c, str):
        return c
    if not c:
        return f"No hay ningun correo con el id {id}."

    cuerpo = c["cuerpo"] or "(este correo no tiene texto guardado; suele ser una imagen o un adjunto)"
    recortado = ""
    if len(cuerpo) > CUERPO_MAXIMO:
        cuerpo = cuerpo[:CUERPO_MAXIMO]
        recortado = f"\n\n[…texto recortado en {CUERPO_MAXIMO} caracteres]"

    cabecera = [
        f"De: {c['remitente']}" + (f" <{c['correo_rem']}>" if c["correo_rem"] else ""),
        f"Para: {c['destinatarios'] or '-'}",
        f"Fecha: {_fecha(c['fecha'])}",
        f"Asunto: {c['asunto']}",
        f"Carpeta: {c['carpeta']}",
    ]
    if c["adjuntos"]:
        cabecera.append(f"Adjuntos: {c['adjuntos']}")
    if c.get("origen"):
        cabecera.append(f"Archivo: {c['origen']}")
    return "\n".join(cabecera) + "\n\n" + cuerpo + recortado


@servidor.tool(
    title="Resumen del indice",
    description=("Que correo hay disponible: cuantos mensajes, en que fechas, "
                 "que carpetas y de que archivos de Outlook salieron."),
    annotations=SOLO_LECTURA,
)
def resumen_indice() -> str:
    """Sin argumentos."""
    def consultar(base):
        total = base.total()
        if total == 0:
            return None
        return (total, base.total_con_cuerpo(),
                base.con.execute("SELECT MIN(fecha), MAX(fecha) FROM correos "
                                 "WHERE fecha <> ''").fetchone(),
                base.carpetas(), base.origenes())

    datos = _con_base(consultar)
    if isinstance(datos, str):
        return datos
    if datos is None:
        return ("La base existe pero esta vacia. Abre el Buscador de Correos "
                "y pulsa «Actualizar correos».")
    total, con_cuerpo, rango, carpetas, origenes = datos

    lineas = [
        f"{total} correos indexados, {con_cuerpo} con texto legible "
        f"({porcentaje(con_cuerpo, total)}).",
        f"Van del {_fecha(rango[0])} al {_fecha(rango[1])}.",
        "",
        f"Carpetas ({len(carpetas)}):",
    ]
    for nombre, n in carpetas[:25]:
        lineas.append(f"  {nombre}  —  {n}")
    if len(carpetas) > 25:
        lineas.append(f"  …y {len(carpetas) - 25} carpetas mas")
    if origenes:
        lineas += ["", "Archivos de Outlook de los que provienen:"]
        for origen, n, _ in origenes:
            lineas.append(f"  {origen or '(sin registrar)'}  —  {n}")
    lineas += ["", "Esta copia es de solo lectura: no se puede enviar ni modificar nada."]
    return "\n".join(lineas)


if __name__ == "__main__":
    servidor.run(transport="stdio")
