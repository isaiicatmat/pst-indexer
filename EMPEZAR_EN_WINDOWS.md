# Arranque en la laptop con Outlook

Este archivo es el punto de partida para la sesión de Claude Code que corre en
Windows. Ábrela en la carpeta del proyecto y pídele que siga estos pasos.

## Contexto

La aplicación (`buscador_correos.py`) está terminada y pasa 69 pruebas
automáticas, **pero todas se ejecutaron en macOS contra un Outlook simulado**.
Lo que falta verificar es exactamente lo que no se puede probar sin Windows:

1. Que `win32com` conecte con el Outlook real.
2. Que la extracción del cuerpo funcione con correos reales (HTML, firmas,
   cadenas de respuestas, correos de Exchange).
3. Que las fechas de `pywintypes` se normalicen bien.
4. Que el rendimiento aguante el buzón completo.

## Pasos

Si usas PowerShell, los `.bat` necesitan el prefijo `.\` (`.\crear_ejecutable.bat`);
en `cmd.exe` no hace falta.

```bat
python -m pip install -r requirements.txt
python verificar.py
python probar_todo.py
python inspeccionar_pst.py "C:\Users\Isai\Documents\Indexer\pst-file"
```

- `verificar.py` — confirma Python, PyQt5, pywin32 y que Outlook responde.
- `probar_todo.py` — las 69 pruebas. Deben pasar igual en Windows.
- `inspeccionar_pst.py` — dice el formato del PST de prueba (ANSI o UNICODE),
  su tamaño y **si Outlook ya lo tiene montado**. De eso depende si hace falta
  añadir lectura directa con `libpff`.

## El PST de prueba

El archivo `isai.carreto@pieconsulting.co.kr.pst` (3 GB, UNICODE) **no está
montado** en Outlook: el que Outlook tiene abierto es otra copia, en
`Documents\Archivos de Outlook`.

Se descartó `libpff` porque solo publica ruedas precompiladas para macOS; en
Windows habría que compilar C con Visual C++ Build Tools. En su lugar la app
usa el propio Outlook como lector: el botón **«Añadir archivo .pst»** lo abre
con `AddStoreEx`, igual que Archivo > Abrir > Archivo de datos de Outlook.

Al probarlo hay que anotar **cuánto tarda** en indexar los 3 GB. Si la segunda
pasada resulta lenta, el siguiente paso sería filtrar por fecha con
`Items.Restrict` en lugar de recorrer todo. No se hizo por adelantado porque el
formato de fecha de `Restrict` depende del idioma de Windows y es una fuente
conocida de errores: primero conviene medir.

Después, la prueba de verdad:

```bat
python buscador_correos.py
```

Pulsar «Traer mis correos de Outlook» y comprobar:

- Que el contador de progreso avanza y se puede cancelar.
- Que al terminar, la barra inferior diga un porcentaje **alto** de correos
  «con contenido legible». Si sale bajo, la extracción del cuerpo falla con
  correos reales y hay que revisar `extraer_cuerpo()` en
  `indexador_outlook.py`.
- Que al buscar, el contenido aparezca a la derecha con los términos
  resaltados.
- Que «Abrir en Outlook» abra el correo original.
- Que pulsar «Actualizar correos» por segunda vez tarde segundos, no minutos
  (debe ser incremental).

## Qué reportar

- La salida completa de los tres scripts.
- El porcentaje de «contenido legible» que muestre la barra inferior.
- Cuánto tardó la primera indexación y cuántos correos entraron.
- Cualquier correo que aparezca vacío y que en Outlook sí tenga texto.

## Si algo falla

El error más probable es que `pywin32` no encuentre Outlook por una diferencia
de arquitectura (Python de 64 bits con Outlook de 32 bits o al revés).
`verificar.py` lo detecta. La solución es instalar el Python que coincida con
la versión de Office.
