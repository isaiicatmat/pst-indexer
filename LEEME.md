# Buscador de Correos

Encuentra cualquier correo de Outlook en menos de un segundo, buscando por
remitente, asunto, contenido o fecha.

---

## Cómo usarlo (3 pasos)

1. **Abre `INICIAR.bat`** (doble clic).
2. Pulsa **«Traer mis correos de Outlook»**. Tarda unos minutos la primera vez.
   Deja Outlook abierto mientras tanto.
3. **Escribe lo que buscas.** Los resultados aparecen mientras tecleas y el
   contenido del correo se ve a la derecha, sin hacer doble clic.

Eso es todo. No hay que escribir comandos.

---

## Cómo buscar

| Si escribes… | Encuentra… |
|---|---|
| `factura` | correos con esa palabra en remitente, asunto o contenido |
| `factura pendiente` | los que tengan **las dos** palabras |
| `"propuesta comercial"` | esa frase exacta (con comillas) |
| `4471` | ese número de factura, folio o referencia |
| `maria` | también encuentra «María» (los acentos dan igual) |

El botón **Filtros** añade remitente, carpeta y rango de fechas.

### Atajos

- `Ctrl + F` — ir a la caja de búsqueda
- `Esc` — limpiar y empezar de nuevo
- `↑` `↓` — moverse entre resultados
- `F5` — traer los correos nuevos

---

## Preguntas frecuentes

**¿Tengo que volver a indexar cada vez?**
No. Pulsa «Actualizar correos» cuando quieras incorporar los nuevos: solo lee
los que faltan, así que tarda segundos.

**¿Toca mis correos?**
No. Solo los lee. Nunca borra, mueve ni envía nada.

**¿Se envía algo por internet?**
No. Todo se queda en tu computadora, en el archivo `correos.db`.

**Un correo aparece sin contenido.**
Suele ser un correo que es solo una imagen o un adjunto. Pulsa
**«Abrir en Outlook»** para verlo completo.

**Algo no funciona.**
Ejecuta `python verificar.py`: dice en lenguaje claro qué falta y cómo
arreglarlo.

---

## Para quien mantenga esto

```
buscador_correos.py    Interfaz (PyQt5). Punto de entrada.
motor_busqueda.py      Base SQLite + índice FTS5. Toda la búsqueda.
indexador_outlook.py   Lectura de Outlook por COM.
verificar.py           Diagnóstico del entorno.
probar_todo.py         Las 69 pruebas automáticas.
_version_anterior/     Código antiguo, ya no se usa.
```

Requisitos: Windows con Outlook, Python 3.8+, `PyQt5` y `pywin32`.
`INICIAR.bat` los instala si faltan.

### Decisiones de diseño

- **FTS5 en vez de `LIKE '%…%'`**: entre 200 y 1300 veces más rápido en
  50 000 correos (0.2 ms frente a 216 ms). Es lo que permite buscar mientras
  se teclea.
- **Clave por `EntryID` de Outlook**: reindexar actualiza los correos en vez de
  duplicarlos, y permite abrir el original en Outlook.
- **Cuerpo con cuatro alternativas** (texto plano → HTML limpio → propiedad
  MAPI → vista previa): ningún correo se queda sin contenido.
- **Fechas normalizadas a `YYYY-MM-DD HH:MM:SS`** al guardar: el orden y los
  filtros funcionan igual sea cual sea el idioma de Windows.

Ejecutar las pruebas:

```
python probar_todo.py
```

### Antes de entregar

`correos.db` contiene **todos los correos de la computadora donde se ejecutó**,
con su texto completo. Nunca debe salir de ahí.

Por eso `crear_ejecutable.bat` rehace la carpeta `ENTREGAR` desde cero en cada
compilación y se detiene si encuentra una base de datos dentro. Para probar el
programa usa `dist\BuscadorCorreos.exe`, no el de `ENTREGAR`: al abrirlo se
crea la base junto al ejecutable.

### Ejecutar los .bat desde PowerShell

PowerShell no ejecuta programas de la carpeta actual sin indicarlo. Hay que
anteponer `.\`:

```
.\crear_ejecutable.bat
```

En `cmd.exe` funciona sin el prefijo. Desde el Explorador, doble clic.
