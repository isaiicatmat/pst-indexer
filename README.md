# Buscador de Correos Outlook 🔍

Una aplicación de escritorio simple y potente para buscar en tus archivos PST de Outlook con indexación rápida.

## Características

✅ **Lectura de archivos PST** - Soporta múltiples archivos PST  
✅ **Indexación automática** - Crea una base de datos para búsquedas rápidas  
✅ **Búsqueda avanzada** - Por remitente, asunto, contenido y fecha  
✅ **Interfaz gráfica** - Fácil de usar, sin necesidad de terminal  
✅ **Actualización incremental** - Actualiza el índice cuando recibas nuevos correos  
✅ **Ver detalles** - Haz doble-click para ver el contenido completo del correo  

## Instalación rápida

### Paso 1: Instalar Python

1. Descarga Python desde: https://www.python.org/downloads/
2. **IMPORTANTE**: Marca la casilla "Add Python to PATH" durante la instalación
3. Haz click en "Install Now"

### Paso 2: Instalar dependencias

1. Descarga todos los archivos del buscador en una carpeta
2. Doble-click en `instalar.bat`
3. Espera a que se complete la instalación
4. Presiona cualquier tecla para cerrar

### Paso 3: Ejecutar la aplicación

1. Doble-click en `ejecutar.bat`
2. ¡Listo! La ventana de la aplicación se abrirá

## Uso

### Primera vez: Indexar tus correos

1. Haz click en **"📁 Seleccionar carpeta PST"**
2. Navega a la carpeta con tus archivos .pst:
   ```
   C:\Users\[TuUsuario]\OneDrive - Pie Consulting\Documentos\Archivos de Outlook
   ```
3. Selecciona la carpeta y espera a que se complete la indexación
4. Verás el número de correos indexados en la barra de estado

### Buscar correos

Rellena uno o más campos de búsqueda:

- **Remitente**: Nombre del que envió el correo
- **Asunto**: Parte del asunto del correo
- **Contenido**: Palabras dentro del correo
- **Desde/Hasta**: Rango de fechas (opcional)

Ejemplo:
```
Remitente: juan@empresa.com
Asunto: reunión
```

Haz click en **"🔍 Buscar"** para ejecutar la búsqueda.

### Ver detalles

Haz **doble-click** en cualquier resultado para ver el contenido completo del correo.

### Actualizar el índice

Cuando recibas nuevos correos y quieras buscar en ellos:

1. Haz click en **"🔄 Actualizar índice"**
2. Selecciona la misma carpeta PST
3. La aplicación agregará los nuevos correos al índice

### Ver estadísticas

Haz click en **"ℹ️ Estadísticas"** para ver:
- Total de correos indexados
- Número de archivos PST
- Rango de fechas

## Preguntas frecuentes

### ¿Dónde encuentro mis archivos PST?

Típicamente en:
```
C:\Users\[TuUsuario]\OneDrive - Pie Consulting\Documentos\Archivos de Outlook
```

O si usas OneDrive:
```
C:\Users\[TuUsuario]\OneDrive\Documentos\Archivos de Outlook
```

Para encontrarlo manualmente:
1. Abre Outlook
2. Archivo → Abrir y exportar → Abrir archivo de datos de Outlook
3. Esto te mostrará la ubicación

### ¿Qué pasa si falla la instalación?

Si al ejecutar `instalar.bat` aparecen errores:

1. Verifica que Python está en el PATH:
   - Abre cmd y escribe: `python --version`
   - Si funciona, el PATH está bien

2. Si Python no se encuentra:
   - Desinstala Python
   - Instálalo nuevamente marcando **"Add Python to PATH"**
   - Reinicia la computadora

3. Si aún falla:
   - Abre cmd en la carpeta de la aplicación
   - Ejecuta: `python -m pip install -r requirements.txt`

### ¿Es lenta la búsqueda?

No, la aplicación usa una base de datos SQLite indexada:
- Indexación inicial: ~2-5 segundos por archivo PST
- Búsquedas: instantáneas (< 100ms)

### ¿Qué pasa con mis correos?

- La aplicación **no modifica** tus archivos PST
- Los correos se copian a una base de datos local llamada `email_index.db`
- Puedes eliminar `email_index.db` en cualquier momento y se volverá a crear

### ¿Cuánta memoria usa?

Normalmente 50-200 MB dependiendo del número de correos indexados.

## Archivos incluidos

```
email_searcher/
├── email_searcher.py     # Aplicación principal
├── instalar.bat          # Script de instalación
├── ejecutar.bat          # Script para ejecutar
├── requirements.txt      # Dependencias
├── README.md            # Este archivo
└── email_index.db       # Base de datos (se crea automáticamente)
```

## Limitaciones conocidas

- Solo funciona con Python 3.7+
- Windows 10/11 (puede funcionar en otras versiones)
- Archivos PST en formato moderno (creados con Outlook 2003+)
- Máximo 1000 resultados por búsqueda (ajustable en código)

## Solución de problemas

### "No se encuentra python"
```
Solución: Instala Python y marca "Add Python to PATH"
```

### "ModuleNotFoundError: No module named 'libpst'"
```
Solución: Ejecuta instalar.bat nuevamente
O manualmente: python -m pip install libpst-python
```

### "El archivo PST no se indexa"
```
Solución: Verifica que el archivo PST no esté corrupto
Intenta exportar el PST a una nueva carpeta desde Outlook
```

### La interfaz se ve rara / fea

No te preocupes, eso depende de tu versión de Windows. La funcionalidad es lo importante.

## Mejoras futuras planeadas

- [ ] Búsqueda por destinatarios
- [ ] Exportar resultados a Excel
- [ ] Ver adjuntos
- [ ] Búsqueda por conversación
- [ ] Modo oscuro

## Soporte

Si encuentras problemas:

1. Verifica el archivo `email_searcher.log` para errores detallados
2. Intenta eliminar `email_index.db` y reindexa los PST
3. Actualiza Python a la versión más reciente

## Licencia

Libre para uso personal.

---

¡Espero que te sea útil! 😊
