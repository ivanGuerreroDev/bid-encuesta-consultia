# Aplicación de Escritorio - Generador de Reportes BID Ciberseguridad

## Descripción

Aplicación de escritorio desarrollada con Tkinter para procesar encuestas sobre brechas digitales en ciberseguridad en PYMEs. La aplicación descarga archivos de SharePoint, procesa los datos y genera reportes en formato Excel.

## Características

✅ **Interfaz gráfica intuitiva** con Tkinter  
✅ **Gestión de configuración** con persistencia en archivo JSON  
✅ **Integración con Microsoft SharePoint** mediante Graph API  
✅ **Procesamiento de datos** de encuestas y puntajes  
✅ **Generación de reportes** en formato Excel con múltiples hojas  
✅ **Modo debug** para guardar archivos localmente  
✅ **Log de actividad** en tiempo real  
✅ **Ejecución en segundo plano** sin bloquear la interfaz  

## Instalación

### Requisitos previos

- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Pasos de instalación

1. **Clonar o descargar el proyecto**

2. **Instalar dependencias**:
```bash
pip install -r requirements.txt
```

## Configuración

Al ejecutar la aplicación por primera vez, necesitarás configurar las credenciales de Microsoft Graph API:

1. Haz clic en el botón **⚙ Configuración**
2. Completa los siguientes campos:

   - **Tenant ID**: ID del tenant de Azure AD
   - **Client ID**: ID de la aplicación registrada en Azure AD
   - **Client Secret**: Secreto de la aplicación
   - **Site URL**: URL del sitio de SharePoint (formato: `dominio.sharepoint.com:/sites/NombreSitio`)
   - **Ruta Encuesta**: Ruta del archivo de encuesta en SharePoint
   - **Ruta Puntajes**: Ruta del archivo de puntajes en SharePoint
   - **Nombre Archivo Salida**: Nombre del archivo Excel resultante

3. Opcionalmente, activa el **Modo Debug** para guardar archivos localmente en la carpeta `debug_files/`

4. Haz clic en **Guardar**

### Archivo de configuración

La configuración se guarda en `config.json` en el mismo directorio de la aplicación. Este archivo incluye:

```json
{
    "tenant_id": "tu-tenant-id",
    "client_id": "tu-client-id",
    "client_secret": "tu-client-secret",
    "site_url": "marketingconsultia.sharepoint.com:/sites/BIDCiberseguridad",
    "encuesta_path": "Documentos compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx",
    "puntajes_path": "Documentos compartidos/puntajes.xlsx",
    "debug_mode": false,
    "output_filename": "tabla_radar.xlsx"
}
```

## Uso

### Ejecutar la aplicación

```bash
python app_desktop.py
```

### Generar un reporte

1. Asegúrate de que la configuración esté completa
2. Haz clic en el botón **▶ Generar Reporte**
3. La aplicación mostrará el progreso en el área de log
4. Al finalizar, recibirás una notificación con el resultado

### Proceso de generación

El proceso incluye los siguientes pasos:

1. ✓ Obtención de token de acceso de Microsoft Graph API
2. ✓ Descarga del archivo de encuesta desde SharePoint
3. ✓ Descarga del archivo de puntajes desde SharePoint
4. ✓ Procesamiento de datos de empresas
5. ✓ Cálculo de puntajes por sección
6. ✓ Generación de archivo Excel con múltiples hojas:
   - Resultados agrupados por empresa, tamaño, país y sección
   - General por empresas (con porcentaje total)
   - General por países
7. ✓ Subida del archivo generado a SharePoint

## Estructura del proyecto

```
bid-encuesta-consultia/
├── app.py                  # Aplicación Flask original
├── app_desktop.py          # Aplicación de escritorio Tkinter
├── requirements.txt        # Dependencias del proyecto
├── README_DESKTOP.md       # Esta documentación
├── config.json            # Configuración (se crea automáticamente)
└── debug_files/           # Archivos de debug (si está activado)
```

## Diferencias con la versión Flask

| Característica | Flask (app.py) | Desktop (app_desktop.py) |
|----------------|----------------|--------------------------|
| Interfaz | API REST | GUI con Tkinter |
| Configuración | Variables de entorno (.env) | Archivo JSON con GUI |
| Ejecución | Servidor web | Aplicación local |
| Logs | Terminal/consola | Ventana integrada |
| Uso | Múltiples usuarios remotos | Usuario local |

## Solución de problemas

### Error: "Faltan credenciales de configuración"
- Verifica que hayas configurado `tenant_id`, `client_id` y `client_secret` en la ventana de configuración

### Error: "No se pudo encontrar un drive válido"
- Verifica que la URL del sitio de SharePoint sea correcta
- Asegúrate de tener permisos en el sitio de SharePoint

### Error al descargar archivos
- Verifica que las rutas de los archivos en SharePoint sean correctas
- Asegúrate de que los archivos existan en la ubicación especificada

### La aplicación se congela
- El procesamiento se ejecuta en segundo plano, pero operaciones largas pueden tardar
- Revisa el log para ver el progreso

## Modo Debug

Cuando activas el modo debug:
- Los archivos descargados de SharePoint se guardan en `debug_files/`
- El archivo Excel generado también se guarda localmente
- Útil para verificar el contenido de los archivos sin necesidad de acceder a SharePoint

## Seguridad

⚠️ **Importante**: El archivo `config.json` contiene credenciales sensibles. 

- No compartas este archivo
- Añádelo a `.gitignore` si usas control de versiones
- Considera cifrar el archivo en producción

## Soporte

Para problemas o preguntas:
1. Revisa el log de actividad en la aplicación
2. Verifica la configuración
3. Activa el modo debug para más detalles

## Licencia

[Especifica tu licencia aquí]
