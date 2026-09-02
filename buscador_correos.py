"""
Buscador de Correos — aplicacion principal.
Ejecutar:  python buscador_correos.py
"""
import os
import sys
import html as _html
import re
from datetime import datetime

from PyQt5.QtCore import (Qt, QThread, pyqtSignal, QTimer, QSize, QRect, QDate)
from PyQt5.QtGui import (QFont, QColor, QPainter, QPen, QIcon, QKeySequence, QPalette)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
    QPushButton, QLabel, QListWidget, QListWidgetItem, QStyledItemDelegate,
    QStyle, QTextBrowser, QSplitter, QFrame, QMessageBox, QProgressDialog,
    QDateEdit, QComboBox, QToolButton, QSizePolicy, QStackedWidget, QShortcut,
    QFileDialog, QMenu)

from motor_busqueda import (BaseCorreos, importar_base_antigua, carpeta_datos,
                            porcentaje)
from indexador_outlook import (Indexador, outlook_disponible,
                               montar_pst, desmontar_pst)

APP = "Buscador de Correos"
MESES = ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sep", "oct", "nov", "dic"]

C_FONDO   = "#FFFFFF"
C_PANEL   = "#F7F8FA"
C_BORDE   = "#E3E6EA"
C_TEXTO   = "#1A1D21"
C_SUAVE   = "#6B7280"
C_ACENTO  = "#2563EB"
C_ACENTO2 = "#EFF4FF"


def fecha_amable(iso):
    """'2024-03-15 10:23:45' -> '15 mar 2024, 10:23'"""
    if not iso:
        return "sin fecha"
    try:
        d = datetime.strptime(iso[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return iso[:10] or "sin fecha"
    hoy = datetime.now()
    if d.date() == hoy.date():
        return f"hoy, {d:%H:%M}"
    if (hoy - d).days == 1:
        return f"ayer, {d:%H:%M}"
    return f"{d.day} {MESES[d.month - 1]} {d.year}, {d:%H:%M}"


# --------------------------------------------------------------------- indexado
class HiloIndexado(QThread):
    avance = pyqtSignal(str, int, int)
    termino = pyqtSignal(int, str)

    def __init__(self, ruta_db, solo_nuevos=True):
        super().__init__()
        self.ruta_db = ruta_db
        self.solo_nuevos = solo_nuevos
        self._cancelar = False

    def cancelar(self):
        self._cancelar = True

    def run(self):
        base = None
        try:
            base = BaseCorreos(self.ruta_db)      # conexion propia de este hilo
            idx = Indexador(base,
                            progreso=lambda t, h=0, n=0: self.avance.emit(t, h, n),
                            cancelado=lambda: self._cancelar)
            n = idx.ejecutar(solo_nuevos=self.solo_nuevos)
            self.termino.emit(n, "")
        except Exception as e:
            self.termino.emit(0, str(e))
        finally:
            if base:
                base.cerrar()


# ----------------------------------------------------------------- lista bonita
class DelegadoCorreo(QStyledItemDelegate):
    """Dibuja cada correo en tres lineas, estilo bandeja moderna."""
    ALTO = 76

    def sizeHint(self, option, index):
        return QSize(option.rect.width(), self.ALTO)

    def paint(self, p, option, index):
        d = index.data(Qt.UserRole) or {}
        r = option.rect
        sel = option.state & QStyle.State_Selected
        hov = option.state & QStyle.State_MouseOver

        p.save()
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(r, QColor(C_ACENTO2) if sel else
                   (QColor("#F2F4F7") if hov else QColor(C_FONDO)))
        if sel:
            p.fillRect(QRect(r.left(), r.top(), 3, r.height()), QColor(C_ACENTO))
        p.setPen(QPen(QColor(C_BORDE)))
        p.drawLine(r.left() + 14, r.bottom(), r.right() - 14, r.bottom())

        x, ancho = r.left() + 16, r.width() - 32
        fm_der = p.fontMetrics()
        fecha = fecha_amable(d.get("fecha", ""))
        f = QFont(option.font); f.setPointSize(max(8, option.font.pointSize() - 1))
        p.setFont(f)
        ancho_fecha = p.fontMetrics().horizontalAdvance(fecha) + 8
        p.setPen(QColor(C_SUAVE))
        p.drawText(QRect(r.right() - ancho_fecha - 16, r.top() + 12, ancho_fecha, 16),
                   Qt.AlignRight | Qt.AlignVCenter, fecha)

        f = QFont(option.font); f.setBold(True)
        p.setFont(f); p.setPen(QColor(C_TEXTO))
        rem = p.fontMetrics().elidedText(d.get("remitente") or "(sin remitente)",
                                         Qt.ElideRight, ancho - ancho_fecha - 10)
        p.drawText(QRect(x, r.top() + 10, ancho - ancho_fecha - 10, 18),
                   Qt.AlignLeft | Qt.AlignVCenter, rem)

        f = QFont(option.font); f.setBold(False)
        p.setFont(f); p.setPen(QColor(C_TEXTO))
        marca = "  \U0001F4CE" if d.get("adjuntos") else ""
        asunto = p.fontMetrics().elidedText((d.get("asunto") or "(sin asunto)") + marca,
                                            Qt.ElideRight, ancho)
        p.drawText(QRect(x, r.top() + 30, ancho, 18), Qt.AlignLeft | Qt.AlignVCenter, asunto)

        f = QFont(option.font); f.setPointSize(max(8, option.font.pointSize() - 1))
        p.setFont(f); p.setPen(QColor(C_SUAVE))
        extracto = re.sub(r"\s+", " ", (d.get("extracto") or "")).strip() or "(sin contenido)"
        carpeta = d.get("carpeta", "")
        linea = f"{carpeta}  ·  {extracto}" if carpeta else extracto
        p.drawText(QRect(x, r.top() + 50, ancho, 16), Qt.AlignLeft | Qt.AlignVCenter,
                   p.fontMetrics().elidedText(linea, Qt.ElideRight, ancho))
        p.restore()


# ------------------------------------------------------------------- ventana
class Ventana(QMainWindow):
    def __init__(self, ruta_db="correos.db"):
        super().__init__()
        self.ruta_db = ruta_db
        self.base = BaseCorreos(ruta_db)
        self.hilo = None
        self.dialogo = None
        self.terminos = []
        self._construir()
        self._primer_arranque()

    # ---------------------------------------------------------------- montaje
    def _construir(self):
        self.setWindowTitle(APP)
        self.resize(1280, 780)
        self.setMinimumSize(940, 560)
        self.setStyleSheet(self._estilos())

        raiz = QWidget(); self.setCentralWidget(raiz)
        col = QVBoxLayout(raiz); col.setContentsMargins(0, 0, 0, 0); col.setSpacing(0)

        col.addWidget(self._barra_superior())
        self.barra_busqueda = self._barra_busqueda()
        col.addWidget(self.barra_busqueda)
        self.panel_filtros = self._filtros()
        self.panel_filtros.setVisible(False)
        col.addWidget(self.panel_filtros)

        self.pilas = QStackedWidget()
        self.pilas.addWidget(self._vista_resultados())   # 0
        self.pilas.addWidget(self._vista_vacia())        # 1
        col.addWidget(self.pilas, 1)

        self.estado = QLabel("")
        self.estado.setObjectName("estado")
        self.estado.setContentsMargins(16, 6, 16, 6)
        col.addWidget(self.estado)

        QShortcut(QKeySequence("Ctrl+F"), self, lambda: self.caja.setFocus())
        QShortcut(QKeySequence("Ctrl+L"), self, lambda: self.caja.setFocus())
        QShortcut(QKeySequence("Escape"), self, self._limpiar)
        QShortcut(QKeySequence("F5"), self, self.actualizar_correos)

    def _barra_superior(self):
        b = QFrame(); b.setObjectName("superior")
        h = QHBoxLayout(b); h.setContentsMargins(16, 10, 16, 10); h.setSpacing(10)
        t = QLabel(APP); t.setObjectName("titulo")
        h.addWidget(t); h.addStretch()
        self.btn_actualizar = QPushButton("Actualizar correos")
        self.btn_actualizar.setObjectName("primario")
        self.btn_actualizar.setCursor(Qt.PointingHandCursor)
        self.btn_actualizar.setToolTip("Lee Outlook y trae los correos nuevos (F5)")
        self.btn_actualizar.clicked.connect(self.actualizar_correos)

        self.btn_pst = QPushButton("Añadir archivo .pst")
        self.btn_pst.setObjectName("secundario")
        self.btn_pst.setCursor(Qt.PointingHandCursor)
        self.btn_pst.setToolTip("Para buscar dentro de un archivo .pst archivado "
                                "que no está abierto en Outlook")
        self.btn_pst.clicked.connect(self.agregar_pst)
        self.btn_opciones = QToolButton()
        self.btn_opciones.setText("⋯")
        self.btn_opciones.setObjectName("opciones")
        self.btn_opciones.setCursor(Qt.PointingHandCursor)
        self.btn_opciones.setToolTip("Más opciones")
        self.btn_opciones.setPopupMode(QToolButton.InstantPopup)
        menu = QMenu(self)
        menu.addAction("Vaciar el índice y empezar de cero", self.vaciar_indice)
        menu.addSeparator()
        menu.addAction("Ver dónde se guardan los datos", self.mostrar_carpeta_datos)
        self.btn_opciones.setMenu(menu)

        h.addWidget(self.btn_pst)
        h.addWidget(self.btn_actualizar)
        h.addWidget(self.btn_opciones)
        return b

    def _barra_busqueda(self):
        b = QFrame(); b.setObjectName("busqueda")
        v = QVBoxLayout(b); v.setContentsMargins(16, 12, 16, 12); v.setSpacing(8)
        fila = QHBoxLayout(); fila.setSpacing(8)
        self.caja = QLineEdit()
        self.caja.setObjectName("caja")
        self.caja.setPlaceholderText("Escribe lo que buscas: un nombre, una palabra, un número de factura...")
        self.caja.setClearButtonEnabled(True)
        self.caja.textChanged.connect(self._al_teclear)
        self.caja.returnPressed.connect(self.buscar)
        fila.addWidget(self.caja, 1)

        self.btn_filtros = QToolButton()
        self.btn_filtros.setText("Filtros")
        self.btn_filtros.setCheckable(True)
        self.btn_filtros.setCursor(Qt.PointingHandCursor)
        self.btn_filtros.setObjectName("secundario")
        self.btn_filtros.toggled.connect(self._alternar_filtros)
        fila.addWidget(self.btn_filtros)
        v.addLayout(fila)

        self.pista = QLabel("Se busca en el remitente, el asunto y el contenido de todos los correos.")
        self.pista.setObjectName("pista")
        v.addWidget(self.pista)

        self.temporizador = QTimer(self)
        self.temporizador.setSingleShot(True)
        self.temporizador.setInterval(250)
        self.temporizador.timeout.connect(self.buscar)
        return b

    def _filtros(self):
        b = QFrame(); b.setObjectName("filtros")
        h = QHBoxLayout(b); h.setContentsMargins(16, 10, 16, 10); h.setSpacing(10)

        h.addWidget(QLabel("Remitente:"))
        self.f_remitente = QLineEdit(); self.f_remitente.setPlaceholderText("nombre o correo")
        self.f_remitente.setMaximumWidth(220)
        self.f_remitente.textChanged.connect(self._al_teclear)
        h.addWidget(self.f_remitente)

        h.addWidget(QLabel("Carpeta:"))
        self.f_carpeta = QComboBox(); self.f_carpeta.setMinimumWidth(190)
        self.f_carpeta.currentIndexChanged.connect(self.buscar)
        h.addWidget(self.f_carpeta)

        self.f_usar_fecha = QToolButton()
        self.f_usar_fecha.setText("Filtrar por fecha")
        self.f_usar_fecha.setCheckable(True)
        self.f_usar_fecha.setObjectName("secundario")
        self.f_usar_fecha.setCursor(Qt.PointingHandCursor)
        self.f_usar_fecha.toggled.connect(self._alternar_fechas)
        h.addWidget(self.f_usar_fecha)

        self.lbl_desde = QLabel("del")
        self.f_desde = QDateEdit(QDate.currentDate().addYears(-1))
        self.lbl_hasta = QLabel("al")
        self.f_hasta = QDateEdit(QDate.currentDate())
        for w in (self.f_desde, self.f_hasta):
            w.setCalendarPopup(True); w.setDisplayFormat("dd/MM/yyyy")
            w.dateChanged.connect(self.buscar)
        for w in (self.lbl_desde, self.f_desde, self.lbl_hasta, self.f_hasta):
            w.setVisible(False); h.addWidget(w)

        h.addStretch()
        b_limpiar = QPushButton("Limpiar filtros")
        b_limpiar.setObjectName("secundario"); b_limpiar.setCursor(Qt.PointingHandCursor)
        b_limpiar.clicked.connect(self._limpiar)
        h.addWidget(b_limpiar)
        return b

    def _vista_resultados(self):
        div = QSplitter(Qt.Horizontal)
        div.setObjectName("division")
        div.setHandleWidth(1)

        izq = QWidget(); vi = QVBoxLayout(izq)
        vi.setContentsMargins(0, 0, 0, 0); vi.setSpacing(0)
        self.conteo = QLabel(""); self.conteo.setObjectName("conteo")
        self.conteo.setContentsMargins(16, 10, 16, 8)
        vi.addWidget(self.conteo)
        self.lista = QListWidget()
        self.lista.setObjectName("lista")
        self.lista.setItemDelegate(DelegadoCorreo())
        self.lista.setMouseTracking(True)
        self.lista.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.lista.currentItemChanged.connect(self._mostrar_correo)
        vi.addWidget(self.lista, 1)
        div.addWidget(izq)

        der = QWidget(); vd = QVBoxLayout(der)
        vd.setContentsMargins(0, 0, 0, 0); vd.setSpacing(0)
        self.cabecera = QLabel(""); self.cabecera.setObjectName("cabecera")
        self.cabecera.setContentsMargins(20, 16, 20, 12)
        self.cabecera.setWordWrap(True)
        self.cabecera.setTextInteractionFlags(Qt.TextSelectableByMouse)
        vd.addWidget(self.cabecera)
        self.cuerpo = QTextBrowser(); self.cuerpo.setObjectName("cuerpo")
        self.cuerpo.setOpenExternalLinks(False)
        vd.addWidget(self.cuerpo, 1)

        acciones = QFrame(); acciones.setObjectName("acciones")
        ha = QHBoxLayout(acciones); ha.setContentsMargins(20, 10, 20, 10); ha.setSpacing(8)
        self.btn_outlook = QPushButton("Abrir en Outlook")
        self.btn_outlook.setObjectName("primario"); self.btn_outlook.setCursor(Qt.PointingHandCursor)
        self.btn_outlook.clicked.connect(self._abrir_en_outlook)
        self.btn_copiar = QPushButton("Copiar contenido")
        self.btn_copiar.setObjectName("secundario"); self.btn_copiar.setCursor(Qt.PointingHandCursor)
        self.btn_copiar.clicked.connect(self._copiar)
        ha.addWidget(self.btn_outlook); ha.addWidget(self.btn_copiar); ha.addStretch()
        vd.addWidget(acciones)
        div.addWidget(der)

        div.setStretchFactor(0, 4); div.setStretchFactor(1, 6)
        div.setSizes([460, 700])
        self.division = div
        return div

    def _vista_vacia(self):
        w = QWidget(); v = QVBoxLayout(w)
        v.setAlignment(Qt.AlignCenter); v.setSpacing(14)
        t = QLabel("Aún no hay correos indexados")
        t.setObjectName("vacio_titulo"); t.setAlignment(Qt.AlignCenter)
        s = QLabel("Pulsa el botón y la aplicación leerá tus correos de Outlook.\n"
                   "Solo hay que hacerlo una vez; después las búsquedas son instantáneas.")
        s.setObjectName("vacio_texto"); s.setAlignment(Qt.AlignCenter)
        b = QPushButton("Traer mis correos de Outlook")
        b.setObjectName("primario"); b.setCursor(Qt.PointingHandCursor)
        b.setMinimumHeight(42); b.setMaximumWidth(280)
        b.clicked.connect(self.actualizar_correos)
        b2 = QPushButton("…o abrir un archivo .pst archivado")
        b2.setObjectName("secundario"); b2.setCursor(Qt.PointingHandCursor)
        b2.setMaximumWidth(280); b2.setMinimumHeight(38)
        b2.clicked.connect(self.agregar_pst)
        v.addWidget(t); v.addWidget(s); v.addWidget(b, 0, Qt.AlignCenter)
        v.addWidget(b2, 0, Qt.AlignCenter)
        return w

    # ---------------------------------------------------------------- arranque
    def _primer_arranque(self):
        if self.base.total() == 0:
            antigua = os.path.join(os.path.dirname(os.path.abspath(self.ruta_db)),
                                   "email_index.db")
            n = importar_base_antigua(self.base, antigua)
            if n:
                self._nota(f"Se recuperaron {n} correos de la versión anterior.")
        self._cargar_carpetas()
        self._refrescar_estado()
        self._modo_vacio(self.base.total() == 0)
        if self.base.total():
            self.buscar()

    def _modo_vacio(self, vacio):
        """Sin correos no tiene sentido mostrar la busqueda: solo estorba."""
        self.barra_busqueda.setVisible(not vacio)
        if vacio:
            self.btn_filtros.setChecked(False)
        self.panel_filtros.setVisible(not vacio and self.btn_filtros.isChecked())
        self.pilas.setCurrentIndex(1 if vacio else 0)

    def _cargar_carpetas(self):
        actual = self.f_carpeta.currentText() if self.f_carpeta.count() else ""
        self.f_carpeta.blockSignals(True)
        self.f_carpeta.clear()
        self.f_carpeta.addItem("Todas las carpetas", "")
        for nombre, n in self.base.carpetas():
            if nombre:
                self.f_carpeta.addItem(f"{nombre}  ({n})", nombre)
        i = self.f_carpeta.findText(actual)
        if i >= 0:
            self.f_carpeta.setCurrentIndex(i)
        self.f_carpeta.blockSignals(False)

    def _refrescar_estado(self):
        t = self.base.total()
        c = self.base.total_con_cuerpo()
        texto = (f"{t:,} correos indexados   ·   "
                 f"{c:,} con contenido legible ({porcentaje(c, t)})").replace(",", " ")
        if t and c < t:
            texto += f"   ·   {t - c} sin texto"
        self.estado.setText(texto)

    def _nota(self, texto):
        self.pista.setText(texto)

    # ---------------------------------------------------------------- busqueda
    def _al_teclear(self, *_):
        self.temporizador.start()

    def _alternar_filtros(self, activo):
        self.panel_filtros.setVisible(activo)

    def _alternar_fechas(self, activo):
        for w in (self.lbl_desde, self.f_desde, self.lbl_hasta, self.f_hasta):
            w.setVisible(activo)
        self.buscar()

    def _limpiar(self):
        for w in (self.caja, self.f_remitente):
            w.blockSignals(True); w.clear(); w.blockSignals(False)
        self.f_carpeta.blockSignals(True); self.f_carpeta.setCurrentIndex(0)
        self.f_carpeta.blockSignals(False)
        self.f_usar_fecha.setChecked(False)
        self.buscar()
        self.caja.setFocus()

    def buscar(self):
        # El temporizador de 250 ms puede dispararse justo al cerrar la ventana.
        if getattr(self.base, "cerrada", False):
            return
        if self.base.total() == 0:
            self._modo_vacio(True)
            return
        self._modo_vacio(False)
        texto = self.caja.text().strip()
        usar_f = self.f_usar_fecha.isChecked()
        res = self.base.buscar(
            texto=texto,
            remitente=self.f_remitente.text().strip(),
            carpeta=self.f_carpeta.currentData() or "",
            desde=self.f_desde.date().toString("yyyy-MM-dd") if usar_f else "",
            hasta=self.f_hasta.date().toString("yyyy-MM-dd") if usar_f else "",
            limite=500)
        self.terminos = [p for p in re.split(r"[^\w@.\-]+", texto) if len(p) > 1]
        self._pintar(res, texto)

    def _pintar(self, res, texto):
        self.lista.blockSignals(True)
        self.lista.clear()
        for d in res:
            it = QListWidgetItem()
            it.setData(Qt.UserRole, d)
            self.lista.addItem(it)
        self.lista.blockSignals(False)

        if not res:
            self.conteo.setText("Sin resultados")
            self.cabecera.setText("")
            self.cuerpo.setHtml(
                f"<div style='color:{C_SUAVE};padding:40px;text-align:center;"
                f"font-family:sans-serif'><p style='font-size:15px'>"
                f"No se encontró ningún correo.</p>"
                f"<p>Prueba con menos palabras, revisa los filtros, "
                f"o pulsa <b>Actualizar correos</b> para traer los más recientes.</p></div>")
            return
        tope = "  (se muestran los 500 más recientes)" if len(res) >= 500 else ""
        etiqueta = "correo encontrado" if len(res) == 1 else "correos encontrados"
        self.conteo.setText(f"{len(res)} {etiqueta}{tope}" if texto
                            else f"{len(res)} correos más recientes")
        self.lista.setCurrentRow(0)

    # ---------------------------------------------------------------- detalle
    def _mostrar_correo(self, actual, _anterior=None):
        if actual is None:
            return
        d = actual.data(Qt.UserRole) or {}
        # Se recupera por ID: el contenido SIEMPRE corresponde al correo elegido.
        c = self.base.por_id(d.get("id"))
        if not c:
            self.cuerpo.setPlainText("No se pudo cargar este correo.")
            return
        self.correo_actual = c
        para = _html.escape(c["destinatarios"] or "-")
        rem = _html.escape(c["remitente"] or "(sin remitente)")
        if c["correo_rem"]:
            rem += " &lt;" + _html.escape(c["correo_rem"]) + "&gt;"
        n_adj = c["adjuntos"]
        adj = (f"  ·  {n_adj} adjunto" + ("s" if n_adj != 1 else "")) if n_adj else ""
        self.cabecera.setText(
            f"<div style='font-size:16px;font-weight:600;color:{C_TEXTO}'>"
            f"{_html.escape(c['asunto'] or '(sin asunto)')}</div>"
            f"<div style='color:{C_SUAVE};font-size:12px;margin-top:6px'>"
            f"<b>De:</b> {rem}<br>"
            f"<b>Para:</b> {para}<br>"
            f"{fecha_amable(c['fecha'])}  ·  {_html.escape(c['carpeta'] or '-')}{adj}</div>")

        texto = c["cuerpo"] or ""
        if texto.strip():
            cuerpo = _html.escape(texto).replace("\n", "<br>")
            # Qt no soporta <mark> (HTML5): se resalta con estilo en linea.
            for t in sorted(set(self.terminos), key=len, reverse=True):
                cuerpo = re.sub(
                    f"({re.escape(_html.escape(t))})",
                    r'<span style="background-color:#FFE9A8;color:#1A1D21;">\1</span>',
                    cuerpo, flags=re.IGNORECASE)
            self.cuerpo.setHtml(
                f"<div style='font-family:-apple-system,Segoe UI,sans-serif;"
                f"font-size:13.5px;line-height:1.55;color:{C_TEXTO};padding:4px 6px'>"
                f"{cuerpo}</div>")
        else:
            self.cuerpo.setHtml(
                f"<div style='color:{C_SUAVE};padding:30px;font-family:sans-serif'>"
                f"<p><b>Este correo no tiene texto guardado.</b></p>"
                f"<p>Suele pasar con correos que son solo una imagen o un adjunto. "
                f"Pulsa <b>Abrir en Outlook</b> para verlo completo.</p></div>")

    def _copiar(self):
        c = getattr(self, "correo_actual", None)
        if not c:
            return
        QApplication.clipboard().setText(
            f"De: {c['remitente']}\nPara: {c['destinatarios']}\n"
            f"Fecha: {fecha_amable(c['fecha'])}\nAsunto: {c['asunto']}\n\n{c['cuerpo']}")
        self._nota("Contenido copiado al portapapeles.")

    def _abrir_en_outlook(self):
        c = getattr(self, "correo_actual", None)
        if not c:
            return
        ok, motivo = outlook_disponible()
        if not ok:
            QMessageBox.information(self, APP, motivo)
            return
        eid = c.get("entry_id") or ""
        if eid.startswith("legacy:"):
            QMessageBox.information(
                self, APP,
                "Este correo viene de la versión anterior y no guarda el enlace a Outlook.\n\n"
                "Pulsa 'Actualizar correos' para volver a leerlos y poder abrirlos.")
            return
        try:
            import pythoncom, win32com.client
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            ns.GetItemFromID(eid).Display()
        except Exception as e:
            QMessageBox.warning(self, APP,
                                f"No se pudo abrir el correo en Outlook.\n\nDetalle: {e}")

    # ---------------------------------------------------------------- indexar
    def vaciar_indice(self):
        """Deja la app como recien instalada, sin borrar nada de Outlook."""
        n = self.base.total()
        if n == 0:
            QMessageBox.information(self, APP, "El índice ya está vacío.")
            return
        r = QMessageBox.question(
            self, APP,
            f"Se van a quitar los {n:,} correos del índice de esta aplicación.\n\n"
            "Tus correos de Outlook NO se tocan: siguen intactos. Solo se borra "
            "la copia que usa el buscador, y puedes volver a crearla con "
            "«Actualizar correos».\n\n¿Continuar?".replace(",", " "),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if r != QMessageBox.Yes:
            return
        self.base.vaciar()
        self._cargar_carpetas()
        self._refrescar_estado()
        self._modo_vacio(True)
        self._nota("Índice vaciado. Pulsa «Actualizar correos» para reconstruirlo.")

    def mostrar_carpeta_datos(self):
        carpeta = os.path.dirname(os.path.abspath(self.ruta_db))
        mb = 0
        try:
            mb = os.path.getsize(self.ruta_db) / 1024 / 1024
        except OSError:
            pass
        QMessageBox.information(
            self, APP,
            f"Los correos indexados se guardan aquí:\n\n{self.ruta_db}\n\n"
            f"Tamaño actual: {mb:,.1f} MB\n\n"
            "Es una copia local para poder buscar rápido. Si la borras, se "
            "reconstruye con «Actualizar correos».".replace(",", " "))

    def agregar_pst(self):
        """Abre un .pst archivado dentro de Outlook y lo indexa.
        Se usa el propio Outlook como lector: no hace falta nada mas."""
        ok, motivo = outlook_disponible()
        if not ok:
            QMessageBox.information(self, APP, motivo)
            return

        ruta, _ = QFileDialog.getOpenFileName(
            self, "Elige el archivo .pst", "", "Archivos de Outlook (*.pst)")
        if not ruta:
            return

        mb = os.path.getsize(ruta) / 1024 / 1024 if os.path.exists(ruta) else 0
        r = QMessageBox.question(
            self, APP,
            f"Se va a abrir este archivo dentro de Outlook para poder leerlo:\n\n"
            f"{os.path.basename(ruta)}  ({mb:,.0f} MB)\n\n"
            "Aparecerá en tu lista de carpetas de Outlook, igual que si lo "
            "abrieras con Archivo > Abrir. No se modifica ni se mueve nada, y "
            "puedes quitarlo después desde Outlook.\n\n¿Continuar?"
            .replace(",", " "),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if r != QMessageBox.Yes:
            return

        try:
            import pythoncom, win32com.client
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            ya_estaba, nombre = montar_pst(ns, ruta)
        except Exception as e:
            QMessageBox.warning(self, APP, str(e))
            return

        self._nota(f"Archivo «{nombre}» "
                   + ("ya estaba abierto en Outlook." if ya_estaba else "abierto en Outlook."))
        self.actualizar_correos()

    def actualizar_correos(self):
        if self.hilo and self.hilo.isRunning():
            return
        ok, motivo = outlook_disponible()
        if not ok:
            QMessageBox.information(self, APP, motivo)
            return

        self.dialogo = QProgressDialog("Conectando con Outlook…", "Cancelar", 0, 0, self)
        self.dialogo.setWindowTitle("Actualizando correos")
        self.dialogo.setWindowModality(Qt.WindowModal)
        self.dialogo.setMinimumWidth(460)
        self.dialogo.setMinimumDuration(0)
        self.dialogo.setAutoClose(False)
        self.dialogo.setAutoReset(False)
        self.btn_actualizar.setEnabled(False)
        self.btn_pst.setEnabled(False)
        self._cerrando_indexado = False

        self.hilo = HiloIndexado(self.ruta_db, solo_nuevos=True)
        self.dialogo.canceled.connect(self.hilo.cancelar)
        self.hilo.avance.connect(self._avance)
        self.hilo.termino.connect(self._fin_indexado)
        self.hilo.start()
        self.dialogo.show()

    def _avance(self, texto, hechos, total):
        # Se toma una referencia local: setValue() procesa eventos por dentro y
        # el indexado puede terminar (y anular self.dialogo) en ese instante.
        d = self.dialogo
        if d is None:
            return
        if total:
            d.setMaximum(total)
            d.setValue(min(hechos, total))
            d.setLabelText(f"{texto}\n\n{hechos:,} de {total:,} correos revisados"
                           .replace(",", " "))
        else:
            d.setLabelText(texto)

    def _fin_indexado(self, n, error):
        if getattr(self, "_cerrando_indexado", False):
            return
        self._cerrando_indexado = True
        if self.dialogo:
            self.dialogo.close(); self.dialogo = None
        self.btn_actualizar.setEnabled(True)
        self.btn_pst.setEnabled(True)
        if error:
            QMessageBox.warning(self, APP, f"No se pudo completar la actualización.\n\n{error}")
        self.base = BaseCorreos(self.ruta_db)   # reabrir para ver lo que escribió el hilo
        if not error:
            self.base.borrar_legado()           # fuera los heredados: ya hay reales
        self._cargar_carpetas()
        self._refrescar_estado()
        self._modo_vacio(self.base.total() == 0)
        self.buscar()
        if not error:
            self._nota(f"Actualización terminada: {n} correos nuevos o corregidos.")
        self._cerrando_indexado = False

    def closeEvent(self, e):
        self.temporizador.stop()
        if self.hilo and self.hilo.isRunning():
            self.hilo.cancelar(); self.hilo.wait(3000)
        self.base.cerrar()
        e.accept()

    # ---------------------------------------------------------------- estilos
    def _estilos(self):
        return f"""
        QMainWindow, QWidget {{ background:{C_FONDO}; color:{C_TEXTO};
            font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif; font-size:13px; }}
        #superior {{ background:{C_PANEL}; border-bottom:1px solid {C_BORDE}; }}
        #titulo {{ font-size:15px; font-weight:600; background:transparent;
            border:none; }}
        #busqueda {{ background:{C_FONDO}; border-bottom:1px solid {C_BORDE}; }}
        #filtros {{ background:{C_PANEL}; border-bottom:1px solid {C_BORDE}; }}
        #caja {{ padding:11px 14px; border:1px solid {C_BORDE}; border-radius:8px;
            font-size:14px; background:{C_FONDO}; }}
        #caja:focus {{ border:2px solid {C_ACENTO}; padding:10px 13px; }}
        #pista {{ color:{C_SUAVE}; font-size:11.5px; }}
        #conteo {{ color:{C_SUAVE}; font-size:12px; font-weight:600;
            background:{C_PANEL}; border-bottom:1px solid {C_BORDE}; }}
        #estado {{ color:{C_SUAVE}; font-size:11.5px;
            background:{C_PANEL}; border-top:1px solid {C_BORDE}; }}
        #lista {{ border:none; border-right:1px solid {C_BORDE}; outline:none; }}
        #cabecera {{ background:{C_PANEL}; border-bottom:1px solid {C_BORDE}; }}
        #cuerpo {{ border:none; padding:10px 14px; background:{C_FONDO}; }}
        #acciones {{ background:{C_PANEL}; border-top:1px solid {C_BORDE}; }}
        #vacio_titulo {{ font-size:19px; font-weight:600; }}
        #vacio_texto {{ color:{C_SUAVE}; font-size:13px; }}
        QPushButton#primario, QToolButton#primario {{ background:{C_ACENTO}; color:white;
            border:none; border-radius:7px; padding:9px 18px; font-weight:600; }}
        QPushButton#primario:hover {{ background:#1D4ED8; }}
        QPushButton#primario:disabled {{ background:#9CB4E8; }}
        QPushButton#secundario, QToolButton#secundario {{ background:{C_FONDO};
            color:{C_TEXTO}; border:1px solid {C_BORDE}; border-radius:7px; padding:8px 14px; }}
        QPushButton#secundario:hover, QToolButton#secundario:hover {{ background:{C_PANEL}; }}
        QToolButton#secundario:checked {{ background:{C_ACENTO2};
            border-color:{C_ACENTO}; color:{C_ACENTO}; font-weight:600; }}
        /* El indicador de menu descentraba el texto del boton "..." */
        QToolButton#opciones {{ background:{C_FONDO}; color:{C_TEXTO};
            border:1px solid {C_BORDE}; border-radius:7px;
            padding:8px 14px; font-size:15px; font-weight:700; }}
        QToolButton#opciones:hover {{ background:{C_PANEL}; }}
        QToolButton#opciones::menu-indicator {{ image:none; width:0; }}
        QLineEdit, QComboBox, QDateEdit {{ padding:6px 9px; border:1px solid {C_BORDE};
            border-radius:6px; background:{C_FONDO}; }}
        QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{ border-color:{C_ACENTO}; }}
        QSplitter::handle {{ background:{C_BORDE}; }}
        QScrollBar:vertical {{ background:transparent; width:11px; margin:0; }}
        QScrollBar::handle:vertical {{ background:#C9CED6; border-radius:5px; min-height:30px; }}
        QScrollBar::handle:vertical:hover {{ background:#AAB2BD; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ height:0; }}
        """


def main():
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName(APP)
    ruta = os.path.join(carpeta_datos(), "correos.db")
    try:
        v = Ventana(ruta)
    except Exception as e:
        QMessageBox.critical(
            None, APP,
            "No se pudo abrir la base de datos de correos.\n\n"
            f"Archivo: {ruta}\nDetalle: {e}\n\n"
            "Suele deberse a que la carpeta es de solo lectura o a que el archivo "
            "quedó dañado. Cierra la aplicación, borra el archivo 'correos.db' "
            "y vuelve a abrirla: se creará de nuevo.")
        sys.exit(1)
    v.show()
    v.caja.setFocus()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
