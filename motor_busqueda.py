"""
Motor de busqueda de correos.
Base de datos SQLite + indice de texto completo (FTS5) para busquedas instantaneas.
"""
import os
import re
import sqlite3
import sys
from datetime import datetime

DB_DEFAULT = "correos.db"


def porcentaje(parte, total):
    """Texto del porcentaje que NUNCA redondea a 100 si falta algo.

    5713 de 5728 es 99.74%, y mostrarlo como '100%' oculta que hay 15 correos
    sin contenido. Solo se dice 100% cuando de verdad estan todos.
    """
    if not total:
        return "0%"
    if parte >= total:
        return "100%"
    if parte <= 0:
        return "0%"
    p = 100.0 * parte / total
    return f"{min(p, 99.9):.1f}%"


def carpeta_datos():
    """Donde se guarda correos.db.

    Al empaquetar con PyInstaller, __file__ apunta a una carpeta temporal que
    Windows borra al cerrar: la base se perderia en cada sesion. Por eso se usa
    la carpeta del ejecutable. Si esa carpeta es de solo lectura (por ejemplo
    dentro de Archivos de programa), se cae a la carpeta de datos del usuario.
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    prueba = os.path.join(base, ".prueba_escritura")
    try:
        with open(prueba, "w") as f:
            f.write("x")
        os.remove(prueba)
        return base
    except OSError:
        alterna = os.path.join(
            os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
            "BuscadorCorreos")
        os.makedirs(alterna, exist_ok=True)
        return alterna
ESQUEMA_VERSION = 2


def _conectar(ruta):
    con = sqlite3.connect(ruta)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    return con


def hay_fts5(con):
    try:
        con.execute("CREATE VIRTUAL TABLE IF NOT EXISTS _prueba_fts USING fts5(x)")
        con.execute("DROP TABLE IF EXISTS _prueba_fts")
        return True
    except sqlite3.OperationalError:
        return False


class BaseCorreos:
    """Capa de datos. Guarda correos y los busca rapido."""

    def __init__(self, ruta=DB_DEFAULT):
        self.ruta = ruta
        self.cerrada = False
        self.con = _conectar(ruta)
        self.fts = hay_fts5(self.con)
        self._crear_esquema()

    # ---------------------------------------------------------------- esquema
    def _crear_esquema(self):
        c = self.con
        c.execute("""
            CREATE TABLE IF NOT EXISTS correos (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id      TEXT UNIQUE,
                carpeta       TEXT NOT NULL DEFAULT '',
                remitente     TEXT NOT NULL DEFAULT '',
                correo_rem    TEXT NOT NULL DEFAULT '',
                destinatarios TEXT NOT NULL DEFAULT '',
                asunto        TEXT NOT NULL DEFAULT '',
                cuerpo        TEXT NOT NULL DEFAULT '',
                fecha         TEXT NOT NULL DEFAULT '',
                adjuntos      INTEGER NOT NULL DEFAULT 0,
                indexado      TEXT NOT NULL DEFAULT ''
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS ix_fecha ON correos(fecha DESC)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_carpeta ON correos(carpeta)")
        c.execute("CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT)")
        c.execute("INSERT OR REPLACE INTO meta VALUES ('esquema', ?)", (str(ESQUEMA_VERSION),))

        if self.fts:
            c.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS correos_fts USING fts5(
                    remitente, destinatarios, asunto, cuerpo,
                    content='correos', content_rowid='id', tokenize="unicode61 remove_diacritics 2"
                )
            """)
            for nombre, cuerpo in (
                ("correos_ai", """AFTER INSERT ON correos BEGIN
                    INSERT INTO correos_fts(rowid, remitente, destinatarios, asunto, cuerpo)
                    VALUES (new.id, new.remitente, new.destinatarios, new.asunto, new.cuerpo);
                END"""),
                ("correos_ad", """AFTER DELETE ON correos BEGIN
                    INSERT INTO correos_fts(correos_fts, rowid, remitente, destinatarios, asunto, cuerpo)
                    VALUES ('delete', old.id, old.remitente, old.destinatarios, old.asunto, old.cuerpo);
                END"""),
                ("correos_au", """AFTER UPDATE ON correos BEGIN
                    INSERT INTO correos_fts(correos_fts, rowid, remitente, destinatarios, asunto, cuerpo)
                    VALUES ('delete', old.id, old.remitente, old.destinatarios, old.asunto, old.cuerpo);
                    INSERT INTO correos_fts(rowid, remitente, destinatarios, asunto, cuerpo)
                    VALUES (new.id, new.remitente, new.destinatarios, new.asunto, new.cuerpo);
                END"""),
            ):
                c.execute(f"CREATE TRIGGER IF NOT EXISTS {nombre} {cuerpo}")
        self.con.commit()

    # ---------------------------------------------------------------- escritura
    def entry_ids_existentes(self):
        cur = self.con.execute("SELECT entry_id FROM correos WHERE entry_id IS NOT NULL")
        return {f[0] for f in cur.fetchall()}

    def guardar(self, correos):
        """Inserta o ACTUALIZA (upsert por entry_id). Devuelve cuantos se escribieron."""
        ahora = datetime.now().isoformat(timespec="seconds")
        filas = [(c.get("entry_id"), c.get("carpeta", ""), c.get("remitente", ""),
                  c.get("correo_rem", ""), c.get("destinatarios", ""), c.get("asunto", ""),
                  c.get("cuerpo", ""), c.get("fecha", ""), int(c.get("adjuntos", 0)), ahora)
                 for c in correos]
        self.con.executemany("""
            INSERT INTO correos (entry_id, carpeta, remitente, correo_rem, destinatarios,
                                 asunto, cuerpo, fecha, adjuntos, indexado)
            VALUES (?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(entry_id) DO UPDATE SET
                carpeta=excluded.carpeta, remitente=excluded.remitente,
                correo_rem=excluded.correo_rem, destinatarios=excluded.destinatarios,
                asunto=excluded.asunto, cuerpo=excluded.cuerpo, fecha=excluded.fecha,
                adjuntos=excluded.adjuntos, indexado=excluded.indexado
        """, filas)
        self.con.commit()
        return len(filas)

    def borrar_legado(self):
        """Quita los correos heredados de la version anterior una vez que
        Outlook ya entrego los verdaderos. Evita verlos duplicados."""
        reales = self.con.execute(
            "SELECT COUNT(*) FROM correos WHERE entry_id NOT LIKE 'legacy:%'").fetchone()[0]
        if reales == 0:
            return 0
        cur = self.con.execute("DELETE FROM correos WHERE entry_id LIKE 'legacy:%'")
        self.con.commit()
        return cur.rowcount

    def vaciar(self):
        """Deja la base vacia sin tener que borrar archivos a mano."""
        self.con.execute("DELETE FROM correos")
        if self.fts:
            self.con.execute("INSERT INTO correos_fts(correos_fts) VALUES ('rebuild')")
        self.con.commit()
        try:
            self.con.execute("VACUUM")      # devuelve el espacio al disco
        except sqlite3.OperationalError:
            pass

    # ---------------------------------------------------------------- lectura
    def total(self):
        return self.con.execute("SELECT COUNT(*) FROM correos").fetchone()[0]

    def total_con_cuerpo(self):
        return self.con.execute(
            "SELECT COUNT(*) FROM correos WHERE cuerpo IS NOT NULL AND TRIM(cuerpo) <> ''"
        ).fetchone()[0]

    def carpetas(self):
        cur = self.con.execute(
            "SELECT carpeta, COUNT(*) n FROM correos GROUP BY carpeta ORDER BY n DESC")
        return [(f["carpeta"], f["n"]) for f in cur.fetchall()]

    def por_id(self, cid):
        f = self.con.execute("SELECT * FROM correos WHERE id=?", (cid,)).fetchone()
        return dict(f) if f else None

    # ---------------------------------------------------------------- busqueda
    @staticmethod
    def _a_consulta_fts(texto):
        """Convierte lo que escribe el usuario en una consulta FTS5 valida y segura.
        Cada palabra se busca como prefijo. Las frases entre comillas se respetan."""
        frases = re.findall(r'"([^"]+)"', texto)
        resto = re.sub(r'"[^"]*"', " ", texto)
        palabras = [p for p in re.split(r"[^\w@.\-]+", resto, flags=re.UNICODE) if p]
        partes = ['"%s"' % f.replace('"', "") for f in frases if f.strip()]
        partes += ['"%s"*' % p.replace('"', "") for p in palabras]
        return " AND ".join(partes)

    def buscar(self, texto="", remitente="", asunto="", contenido="",
               desde="", hasta="", carpeta="", limite=500):
        """Busqueda combinada. `texto` usa FTS5 (rapido); el resto son filtros."""
        where, params = [], []

        if remitente:
            where.append("(c.remitente LIKE ? OR c.correo_rem LIKE ?)")
            params += [f"%{remitente}%", f"%{remitente}%"]
        if asunto:
            where.append("c.asunto LIKE ?")
            params.append(f"%{asunto}%")
        if contenido:
            where.append("c.cuerpo LIKE ?")
            params.append(f"%{contenido}%")
        if carpeta:
            where.append("c.carpeta = ?")
            params.append(carpeta)
        if desde:
            where.append("c.fecha >= ?")
            params.append(f"{desde} 00:00:00")
        if hasta:
            where.append("c.fecha <= ?")
            params.append(f"{hasta} 23:59:59")

        consulta_fts = self._a_consulta_fts(texto) if texto.strip() else ""

        if consulta_fts and self.fts:
            sql = ("SELECT c.id, c.remitente, c.asunto, c.fecha, c.carpeta, c.adjuntos, "
                   "       substr(c.cuerpo, 1, 300) AS extracto "
                   "FROM correos_fts f JOIN correos c ON c.id = f.rowid "
                   "WHERE correos_fts MATCH ?")
            p = [consulta_fts] + params
            if where:
                sql += " AND " + " AND ".join(where)
            sql += " ORDER BY c.fecha DESC LIMIT ?"
            p.append(limite)
            try:
                return [dict(f) for f in self.con.execute(sql, p).fetchall()]
            except sqlite3.OperationalError:
                pass  # consulta FTS rara -> caemos al modo LIKE

        if texto.strip():
            where.append("(c.asunto LIKE ? OR c.cuerpo LIKE ? OR c.remitente LIKE ? "
                         "OR c.destinatarios LIKE ?)")
            params += [f"%{texto}%"] * 4

        sql = ("SELECT c.id, c.remitente, c.asunto, c.fecha, c.carpeta, c.adjuntos, "
               "       substr(c.cuerpo, 1, 300) AS extracto FROM correos c")
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY c.fecha DESC LIMIT ?"
        params.append(limite)
        return [dict(f) for f in self.con.execute(sql, params).fetchall()]

    def recientes(self, limite=200):
        cur = self.con.execute(
            "SELECT id, remitente, asunto, fecha, carpeta, adjuntos, "
            "       substr(cuerpo,1,300) AS extracto "
            "FROM correos ORDER BY fecha DESC LIMIT ?", (limite,))
        return [dict(f) for f in cur.fetchall()]

    def cerrar(self):
        self.cerrada = True
        try:
            self.con.close()
        except Exception:
            pass


MARCA_MIGRACION = ".version_anterior_importada"


def importar_base_antigua(base, ruta_antigua="email_index.db"):
    """Trae los correos de la version anterior para no empezar de cero.

    Solo ocurre UNA vez: despues se deja una marca. Sin ella, borrar
    correos.db para empezar limpio volvia a traer los correos viejos, que es
    justo lo contrario de lo que espera quien lo borra.
    """
    if not os.path.exists(ruta_antigua):
        return 0
    marca = os.path.join(os.path.dirname(os.path.abspath(base.ruta)), MARCA_MIGRACION)
    if os.path.exists(marca):
        return 0
    try:
        vieja = sqlite3.connect(ruta_antigua)
        vieja.row_factory = sqlite3.Row
        filas = vieja.execute(
            "SELECT sender, recipient, subject, body, date, pst_file FROM emails").fetchall()
        vieja.close()
    except Exception:
        return 0

    lote = []
    for i, f in enumerate(filas):
        lote.append({
            "entry_id": f"legacy:{i}:{(f['subject'] or '')[:40]}",
            "carpeta": (f["pst_file"] or "").replace("Outlook:", ""),
            "remitente": f["sender"] or "",
            "correo_rem": "",
            "destinatarios": f["recipient"] or "",
            "asunto": f["subject"] or "",
            "cuerpo": f["body"] or "",
            "fecha": normalizar_fecha(f["date"] or ""),
            "adjuntos": 0,
        })
    n = base.guardar(lote) if lote else 0
    try:
        with open(marca, "w", encoding="utf-8") as f:
            f.write("Los correos de email_index.db ya se importaron una vez.\n"
                    "Borra este archivo si quieres volver a importarlos.\n")
    except OSError:
        pass
    return n


def normalizar_fecha(valor):
    """Deja cualquier fecha como 'YYYY-MM-DD HH:MM:SS' para poder ordenar y filtrar."""
    if not valor:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d %H:%M:%S")
    s = str(valor).strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})", s)
    if m:
        return "%s-%s-%s %s:%s:%s" % m.groups()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return "%s-%s-%s 00:00:00" % m.groups()
    # formatos tipo 09/15/2024 10:23:45 AM  o  15/09/2024 10:23
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{2,4})[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([APap][Mm])?", s)
    if m:
        a, b, anio, hh, mm, ss, ampm = m.groups()
        anio = int(anio); anio += 2000 if anio < 100 else 0
        a, b, hh, mm = int(a), int(b), int(hh), int(mm)
        ss = int(ss or 0)
        if ampm:
            up = ampm.upper()
            if up == "PM" and hh < 12: hh += 12
            if up == "AM" and hh == 12: hh = 0
        mes, dia = (a, b) if a <= 12 else (b, a)   # heuristica MM/DD vs DD/MM
        try:
            return datetime(anio, mes, dia, hh, mm, ss).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            return ""
    for fmt in ("%Y-%m-%d %H:%M:%S%z", "%a %b %d %H:%M:%S %Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    return ""
