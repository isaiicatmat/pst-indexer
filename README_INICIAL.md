# PST Indexer 🔍

[![Python 3.7+](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/Platform-Windows-green.svg)]()

Una aplicación de escritorio rápida y fácil de usar para indexar y buscar en archivos PST de Outlook. 

Perfecta cuando Outlook tiene problemas con su búsqueda nativa o necesitas búsquedas más rápidas y precisas.

## 🎯 Características

- **Lectura de archivos PST** - Soporta múltiples archivos .pst simultáneamente
- **Indexación rápida** - Base de datos SQLite para búsquedas instantáneas
- **Búsqueda avanzada** - Por remitente, asunto, contenido y rango de fechas
- **Interfaz gráfica** - GUI intuitiva, sin necesidad de terminal
- **Actualización incremental** - Agrega nuevos correos sin perder datos previos
- **Vista de detalles** - Doble-click para ver el contenido completo del correo
- **Estadísticas** - Visualiza el estado de tu índice

## 🚀 Inicio rápido

### Requisitos previos

- Python 3.7 o superior
- Windows 10/11 (también funciona en Linux/Mac)

### Instalación

1. **Descarga** los archivos del repositorio

2. **Ejecuta el instalador**
   ```bash
   instalar.bat
   ```
   O manualmente:
   ```bash
   pip install -r requirements.txt
   ```

3. **Ejecuta la aplicación**
   ```bash
   ejecutar.bat
   ```
   O:
   ```bash
   python email_searcher.py
   ```

## 📖 Uso

### Primera vez: Indexar tus correos

1. Abre la aplicación
2. Haz click en **"📁 Seleccionar carpeta PST"**
3. Navega a tu carpeta de archivos Outlook
4. Espera a que se complete la indexación

**Ejemplo de ruta típica:**
```
C:\Users\[TuUsuario]\OneDrive - Pie Consulting\Documentos\Archivos de Outlook
```

### Buscar correos

Rellena uno o más campos de búsqueda:

- **Remitente**: `juan@empresa.com`
- **Asunto**: `reunión importante`
- **Contenido**: `presupuesto`
- **Desde/Hasta**: `2024-01-01` a `2024-12-31`

Presiona **"🔍 Buscar"** y obtén resultados al instante.

### Ver detalles

Haz **doble-click** en cualquier resultado para ver el correo completo.

### Actualizar índice

Cuando recibas nuevos correos:

1. Haz click en **"🔄 Actualizar índice"**
2. Selecciona la misma carpeta PST
3. Los nuevos correos se agregarán automáticamente

## 📁 Estructura del proyecto

```
pst-indexer/
├── email_searcher.py      # Aplicación principal con GUI
├── instalar.bat           # Script de instalación (Windows)
├── ejecutar.bat           # Script para ejecutar (Windows)
├── requirements.txt       # Dependencias Python
├── config.json           # Configuración por defecto
├── setup.py              # Setup para instalación pip
├── LICENSE               # Licencia MIT
├── README.md             # Manual de usuario completo
└── .gitignore           # Archivos ignorados por git
```

## ⚙️ Configuración

Edita `config.json` para personalizar:

```json
{
  "default_pst_folder": "C:\\Users\\...",
  "database_path": "email_index.db",
  "max_results_per_search": 1000,
  "enable_logging": true,
  "auto_update_on_startup": false
}
```

## 🔧 Desarrollo

### Requisitos

- Python 3.7+
- `pip` para instalar dependencias

### Instalar dependencias de desarrollo

```bash
pip install -r requirements.txt
```

### Ejecutar en modo desarrollo

```bash
python email_searcher.py
```

### Estructura de código

- **EmailIndexer**: Clase que maneja la indexación y búsqueda en SQLite
- **EmailSearcherGUI**: Clase que maneja la interfaz gráfica con Tkinter
- **Threading**: Para no bloquear la UI durante operaciones largas

## 📊 Rendimiento

- **Indexación**: ~2-5 segundos por archivo PST
- **Búsquedas**: < 100ms para cualquier criterio
- **Memoria**: 50-200 MB dependiendo del volumen de correos
- **Base de datos**: SQLite con índices optimizados

## 🐛 Problemas conocidos

- Solo funciona con archivos PST modernos (Outlook 2003+)
- Si el PST está corrupto, no podrá ser indexado
- Máximo 1000 resultados por búsqueda (ajustable en código)

## 💡 Solución de problemas

### "ModuleNotFoundError: No module named 'libpst'"
```bash
pip install libpst-python
```

### "Python no encontrado"
Asegúrate de haber instalado Python y marcado "Add Python to PATH"

### El PST no se indexa
- Verifica que el archivo no esté corrupto
- Intenta exportar a una nueva carpeta desde Outlook
- Revisa el archivo `email_searcher.log` para errores

## 📋 Roadmap

- [ ] Búsqueda por destinatarios (To, Cc, Bcc)
- [ ] Exportar resultados a Excel
- [ ] Visualizar adjuntos
- [ ] Búsqueda por conversación/thread
- [ ] Modo oscuro
- [ ] Búsqueda con expresiones regulares
- [ ] Soporte para archivos .msg individuales

## 📝 Licencia

Este proyecto está bajo la licencia MIT - ver el archivo [LICENSE](LICENSE) para detalles.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Para cambios importantes, abre un issue primero para discutir qué te gustaría cambiar.

## 📧 Contacto

- Autor: Isaí Carrero Martínez
- Email: isaiicatmat@gmail.com
- GitHub: [@isaiicatmat](https://github.com/isaiicatmat)

---

**¿Te fue útil?** Considera dejar una ⭐ en el repositorio.
