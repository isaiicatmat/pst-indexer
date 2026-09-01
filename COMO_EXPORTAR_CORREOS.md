# Cómo Exportar Correos desde Outlook

El buscador funciona mejor con archivos `.msg` individuales en lugar de archivos `.pst` contenedores.

Aquí hay dos formas de exportar tus correos:

## Opción 1: Exportar desde Outlook (Recomendado)

### Pasos:

1. **Abre Outlook**

2. **Selecciona la carpeta** que quieres exportar (Bandeja de entrada, Enviados, etc.)

3. **Click derecho → Copiar carpeta** (o usa Archivo → Guardar como)

4. **Elige guardar como:**
   - Formato: **Archivo de datos de Outlook (.pst)**
   - O mejor: **Exportar como .msg** (si tu versión lo permite)

5. **Selecciona una carpeta** (ej: `C:\Mis_Correos_Exportados`)

6. **Ejecuta el Buscador:**
   ```bash
   ejecutar.bat
   ```

7. **Click en "📁 Seleccionar carpeta PST"**

8. **Navega a la carpeta** donde guardaste los correos

---

## Opción 2: Usar Script VBA (Más fácil)

### Pasos:

1. **Descarga el archivo `exportar_correos.vbs`** (viene con la app)

2. **Doble-click en `exportar_correos.vbs`**

3. **Se abrirá una ventana** pidiendo seleccionar una carpeta de Outlook

4. **Elige la carpeta** que quieres exportar (Bandeja de entrada, etc.)

5. **Espera a que termine** (verás un mensaje de confirmación)

6. **Los correos se guardarán en:**
   ```
   C:\Users\[TuUsuario]\Desktop\Correos_Exportados
   ```

7. **En el Buscador, selecciona esa carpeta**

---

## Opción 3: Desde Outlook Web (outlook.com)

Si usas Outlook en línea:

1. **Selecciona correos** (Ctrl+A para todos)
2. **Más acciones (...) → Descargar**
3. **Se descargan como .msg** en tu carpeta Descargas
4. **Usa esa carpeta en el Buscador**

---

## Solución de problemas

### "El Buscador no encuentra los correos"

- ✅ Verifica que seleccionaste la carpeta correcta
- ✅ Comprueba que contiene archivos `.msg` o `.pst`
- ✅ Revisa `email_searcher.log` para ver errores

### "Tengo muchos correos, ¿tarda mucho?"

- La indexación es rápida (2-5 segundos por archivo)
- Una vez indexado, las búsquedas son instantáneas
- Puedes indexar múltiples carpetas de una vez

### "Quiero buscar en TODO Outlook"

1. **Exporta varias carpetas** a la misma ubicación
2. **El Buscador las indexará todas juntas**

---

## Notas

- Los archivos `.msg` son más confiables que `.pst`
- El Buscador **NO modifica** tus correos originales
- Todos los correos se indexan localmente (privado, sin internet)

¿Problemas? Revisa el archivo `email_searcher.log` para más detalles.
