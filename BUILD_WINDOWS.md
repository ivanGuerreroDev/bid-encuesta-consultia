# Guía para Crear Ejecutable de Windows

Esta guía explica cómo crear un archivo ejecutable (.exe) de la aplicación de escritorio para Windows.

## 📋 Requisitos Previos

- Python 3.8 o superior instalado en Windows
- Acceso a línea de comandos (CMD o PowerShell)
- Todas las dependencias del proyecto

## 🚀 Método 1: Usando el Script Automático (Recomendado)

### Paso 1: Preparar el entorno

1. Abre la terminal (CMD) en la carpeta del proyecto
2. Asegúrate de tener todas las dependencias instaladas:
   ```cmd
   pip install -r requirements.txt
   ```

### Paso 2: Ejecutar el script de construcción

Simplemente ejecuta el archivo batch:
```cmd
build_windows.bat
```

Este script automáticamente:
- ✅ Verifica e instala PyInstaller si no está presente
- ✅ Instala todas las dependencias necesarias
- ✅ Construye el ejecutable usando la configuración optimizada
- ✅ Muestra la ubicación del archivo final

### Paso 3: Encontrar el ejecutable

El archivo ejecutable se creará en:
```
dist/GeneradorReportesBID.exe
```

## 🛠️ Método 2: Manual (Paso a Paso)

### Paso 1: Instalar PyInstaller

```cmd
pip install pyinstaller
```

### Paso 2: Instalar dependencias

```cmd
pip install -r requirements.txt
```

### Paso 3: Construir el ejecutable

```cmd
pyinstaller build_windows.spec --clean
```

### Paso 4: Ubicar el ejecutable

Busca el archivo en la carpeta `dist/`:
```
dist/GeneradorReportesBID.exe
```

## 📦 Distribución

### Opción A: Ejecutable único (Portable)

El archivo `GeneradorReportesBID.exe` es completamente portable. Puedes:
1. Copiarlo a cualquier carpeta
2. Ejecutarlo directamente sin instalación
3. El archivo `config.json` se creará en el mismo directorio donde esté el .exe

### Opción B: Crear un instalador (Opcional)

Para crear un instalador profesional, puedes usar:
- **NSIS** (Nullsoft Scriptable Install System)
- **Inno Setup**
- **WiX Toolset**

## 📝 Notas Importantes

### Tamaño del ejecutable
- El ejecutable puede pesar entre 50-150 MB debido a que incluye:
  - Python runtime
  - Todas las bibliotecas (pandas, requests, tkinter, etc.)
  - Dependencias de sistema

### Antivirus
- Algunos antivirus pueden marcar el ejecutable como sospechoso
- Esto es normal con ejecutables creados por PyInstaller
- Solución: Agregar excepción en el antivirus o firmar digitalmente el ejecutable

### Primera ejecución
- La primera vez puede tardar un poco más en cargar
- Se creará automáticamente el archivo `config.json` en el mismo directorio

### Modo Debug
Si el ejecutable tiene problemas, puedes compilar en modo debug:

1. Edita `build_windows.spec`
2. Cambia `console=False` a `console=True`
3. Recompila con `pyinstaller build_windows.spec --clean`

Esto mostrará una ventana de consola con mensajes de depuración.

## 🎨 Personalización

### Agregar un ícono personalizado

1. Consigue un archivo `.ico` (ícono de Windows)
2. Colócalo en la carpeta del proyecto
3. Edita `build_windows.spec`:
   ```python
   icon='mi_icono.ico'  # Reemplaza None con el nombre de tu ícono
   ```
4. Recompila

### Cambiar el nombre del ejecutable

Edita `build_windows.spec` y cambia:
```python
name='GeneradorReportesBID',  # Cambia este nombre
```

## 🐛 Solución de Problemas

### Error: "PyInstaller no encontrado"
```cmd
pip install --upgrade pyinstaller
```

### Error: "Module not found"
Asegúrate de que todas las dependencias estén instaladas:
```cmd
pip install -r requirements.txt
```

### El ejecutable no inicia
1. Compila en modo debug (`console=True`)
2. Revisa los mensajes de error en la consola
3. Verifica que Python sea 64-bit si estás en Windows 64-bit

### Error de Tkinter
Si hay problemas con Tkinter:
1. Reinstala Python asegurándote de marcar "tcl/tk and IDLE"
2. Verifica que tkinter funcione: `python -m tkinter`

## 📂 Estructura de archivos después de compilar

```
bid-encuesta-consultia/
├── app_desktop.py           # Código fuente
├── build_windows.spec       # Configuración PyInstaller
├── build_windows.bat        # Script de construcción
├── requirements.txt         # Dependencias
├── build/                   # Archivos temporales (puedes eliminar)
└── dist/                    # Carpeta con el ejecutable
    └── GeneradorReportesBID.exe  # ⭐ EJECUTABLE FINAL
```

## 🚀 Distribución a usuarios

Para entregar la aplicación a otros usuarios:

1. **Solo el ejecutable**:
   - Envía únicamente `GeneradorReportesBID.exe`
   - El usuario solo necesita ejecutarlo
   - No requiere Python instalado

2. **Con documentación**:
   ```
   GeneradorReportesBID/
   ├── GeneradorReportesBID.exe
   ├── README.txt (instrucciones de uso)
   └── config.json (opcional, con configuración pre-cargada)
   ```

## 💡 Consejos

- ✅ Compila en una máquina limpia para asegurar compatibilidad
- ✅ Prueba el ejecutable en diferentes versiones de Windows
- ✅ Considera usar modo `console=True` para la primera versión (facilita debug)
- ✅ Documenta la versión de Python usada para compilar
- ✅ Mantén backups del código fuente

## 🔄 Actualizar el ejecutable

Cuando hagas cambios en el código:

1. Modifica `app_desktop.py`
2. Ejecuta nuevamente `build_windows.bat`
3. El nuevo ejecutable estará en `dist/`

## 📞 Soporte

Si tienes problemas:
1. Revisa los logs en modo debug
2. Verifica que todas las dependencias estén actualizadas
3. Consulta la documentación de PyInstaller: https://pyinstaller.org/
