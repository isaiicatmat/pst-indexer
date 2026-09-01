# Distribución del Buscador de Correos

## Para usuarios finales (tu compañero)

### Opción A: Ejecutable simple (Recomendado)

1. **Descargar archivos:**
   - `Buscador de Correos.exe`
   - `COMO_EXPORTAR_CORREOS.md`
   - `exportar_correos.vbs`

2. **Guardar en una carpeta** (ej: `C:\Buscador_Correos`)

3. **Doble-click en `Buscador de Correos.exe`**

✅ **Listo.** No necesita instalar nada más.

---

## Cómo generar el ejecutable (para ti)

### Paso 1: Preparar el proyecto

```bash
cd pst-indexer
git pull
```

### Paso 2: Instalar PyInstaller

```bash
pip install pyinstaller
```

### Paso 3: Generar ejecutable

```bash
pyinstaller --onefile --windowed --name "Buscador de Correos" email_searcher.py
```

### Paso 4: Encontrar el ejecutable

El archivo `.exe` estará en:
```
dist\Buscador de Correos.exe
```

### Paso 5: Distribuir

Copia estos archivos juntos en una carpeta:
```
Buscador de Correos\
├── Buscador de Correos.exe
├── exportar_correos.vbs
├── COMO_EXPORTAR_CORREOS.md
└── README.md (opcional)
```

Comprime la carpeta como `.zip` y envía a tu compañero.

---

## Instalación para tu compañero

1. **Descarga el .zip**
2. **Descomprime en una carpeta** (ej: `C:\Buscador_Correos`)
3. **Doble-click en `Buscador de Correos.exe`**
4. **Sigue el archivo `COMO_EXPORTAR_CORREOS.md`**

---

## Ventajas del ejecutable

✅ No requiere Python instalado  
✅ Una sola carpeta fácil de llevar  
✅ Funciona en cualquier Windows  
✅ Sin instalación complicada  
✅ Comprimido en ~50-100MB  

---

## Desventajas (mínimas)

- El .exe es más grande (~50MB)
- Algunos antivirus pueden alertar (falso positivo)
- Si cambia el código, hay que regenerar

---

## Si quieres incluir el script de exportación

El archivo `exportar_correos.vbs` funciona independientemente:

```bash
cscript.exe exportar_correos.vbs
```

Incluye ambos archivos juntos para máxima comodidad.

---

## Solución de problemas

### "El antivirus bloquea el .exe"

Esto es falso positivo. El ejecutable es generado por PyInstaller que es confiable.

**Solución:**
- Agregar excepción en el antivirus
- O distribuir el código Python + instrucciones de instalación

### "El .exe tarda en abrir"

Esto es normal (PyInstaller descomprime el archivo en memoria).

**Tarda ~5 segundos la primera vez**

### "El .exe no funciona en otra computadora"

Asegúrate de:
- ✅ Copiar también `exportar_correos.vbs`
- ✅ La otra computadora tiene Windows 10+ (o Win 7 SP1+)
- ✅ Hay espacio en disco

---

## Script automático (opcional)

Si quieres automatizar todo, usa:

```bash
crear_ejecutable.bat
```

Este script:
1. Instala PyInstaller
2. Genera el .exe automáticamente
3. Te dice dónde está

Luego copias la carpeta `dist\Buscador de Correos.exe` a tu compañero.

---

## Actualizaciones futuras

Si actualizas el código:

```bash
git pull
crear_ejecutable.bat
```

Y vuelve a generar el .exe.

---

¿Preguntas? Consulta el README.md principal.
