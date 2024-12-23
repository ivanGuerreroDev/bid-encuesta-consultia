import pandas as pd
import re
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import os
from dotenv import load_dotenv
load_dotenv()

# Configuración de conexión a SharePoint
sharepoint_site = "https://marketingconsultia.sharepoint.com/sites/BIDCiberseguridad"
sharepoint_url_encuesta = "/sites/BIDCiberseguridad/Documentos%20compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx"
sharepoint_url_puntajes = "/sites/BIDCiberseguridad/Documentos%20compartidos/puntajes.xlsx"
username =  os.getenv("SHAREPOINT_USER")
password = os.getenv("SHAREPOINT_PASSWORD")
# Conectar a SharePoint
ctx = ClientContext(sharepoint_site).with_credentials(UserCredential(username, password))
#print folders of site
web = ctx.web
folderDocumentosCompartidos = web.get_folder_by_server_relative_url("/sites/BIDCiberseguridad/Documentos%20compartidos")
# get file encuesta
file_encuesta = web.get_file_by_server_relative_url(sharepoint_url_encuesta)
if not file_encuesta.execute_query():
    print("File 'df_encuesta.xlsx' not found in SharePoint.")
    exit()

# get file puntajes
file_puntajes = web.get_file_by_server_relative_url(sharepoint_url_puntajes)
if not file_puntajes.execute_query():
    print("File 'df_puntajes.xlsx' not found in SharePoint.")
    exit()

    
# Descargar df_encuesta
with open("df_encuesta.xlsx", "wb") as file_encuesta:
    response_encuesta = ctx.web.get_file_by_server_relative_url(sharepoint_url_encuesta).download(file_encuesta).execute_query()

# Descargar df_puntajes
with open("df_puntajes.xlsx", "wb") as file_puntajes:
    response_puntajes = ctx.web.get_file_by_server_relative_url(sharepoint_url_puntajes).download(file_puntajes).execute_query()

# Cargar los archivos descargados como DataFrames
df_encuesta = pd.read_excel("df_encuesta.xlsx", sheet_name="Form1")
df_puntajes = pd.read_excel("df_puntajes.xlsx")

# Limpiar nombres de columnas para evitar problemas
df_puntajes.columns = df_puntajes.columns.str.strip()

# Crear una lista para almacenar los resultados
resultados = []
empresas = {}
secciones_puntaje = {}

# Obtener nombre de la empresa
for index, row in df_encuesta.iterrows():
    for pregunta in df_encuesta.columns:
        respuesta =  row[pregunta]
        
        if isinstance(pregunta, str):
            current_pregunta_code = re.search(r"\[([A-Za-z0-9_.]+)\]", pregunta)
            pregunta_code = current_pregunta_code.group(1) if current_pregunta_code else ""
        else:
            pregunta_code = ""
        if isinstance(respuesta, str):
            current_respuesta_code = re.search(r"\[([A-Za-z0-9_.]+)\]", respuesta)
            respuesta_code = current_respuesta_code.group(1) if current_respuesta_code else ""
        else:
            respuesta_code = ""
        if( row['ID'] not in empresas):
            empresas[row['ID']] = {}
        if pregunta_code == "Pg001":
            empresas[row['ID']]['Empresa'] = respuesta
        if respuesta_code == "Pg011.01":
            empresas[row['ID']]['Pais'] = 'Costa Rica'
        if respuesta_code == "Pg011.02":
            empresas[row['ID']]['Pais'] = 'Panamá'

# Sumar puntajes por secciones
for index, row in df_puntajes.iterrows():
    if row['Seccion'] not in secciones_puntaje:
        secciones_puntaje[row['Seccion']] = 0
    secciones_puntaje[row['Seccion']] += row['Puntaje']

df_empresas = pd.DataFrame(empresas.items(), columns=['ID', 'Empresa'])

# Iterar sobre cada fila en el DataFrame de Encuesta
for index, row in df_encuesta.iterrows():
    for pregunta in df_encuesta.columns:
        current_id = row['ID']
        current_respuesta = row[pregunta]
        
        # Asegurarse de que current_respuesta es una cadena
        if isinstance(current_respuesta, str):
            current_respuesta_code = re.search(r"\[([A-Za-z0-9_.]+)\]", current_respuesta)
            respuesta_code = current_respuesta_code.group(1) if current_respuesta_code else ""
        else:
            respuesta_code = ""

        # Calcular el puntaje correspondiente
        puntaje_fila = 0
        seccion = ''
        tamano = ''
        for index_puntaje, row_puntaje in df_puntajes.iterrows():
            if (row_puntaje['Respuesta Pequeña'] == respuesta_code or 
                row_puntaje['Respuesta Mediana'] == respuesta_code):
                puntaje_fila = row_puntaje['Puntaje']
                seccion = row_puntaje['Seccion']
                if row_puntaje['Respuesta Pequeña'] == respuesta_code:
                    tamano = 'Pequeña'
                if row_puntaje['Respuesta Mediana'] == respuesta_code:
                    tamano = 'Mediana'
                break

        # Almacenar solo ID, Empresa y Puntaje en la lista
        empresa_data = empresas.get(current_id, {})
        if isinstance(empresa_data, dict):
            nombre_empresa = empresa_data.get('Empresa', '')
        else:
            nombre_empresa = empresa_data  # Si es una cadena, úsala directamente
        
        #get pais
        pais = empresas[current_id].get('Pais', '')
        if nombre_empresa != '' and seccion != '':
            resultados.append({
                'ID': current_id,
                'Empresa': nombre_empresa,
                'Tamaño': tamano,
                'Pais': pais,
                'Puntaje': puntaje_fila,
                'Seccion': seccion,
                'Puntaje Seccion': secciones_puntaje[seccion],
            })
# Crear un nuevo DataFrame con los resultados
df_resultados = pd.DataFrame(resultados)

# Asegurarse de que 'Puntaje' es numérico
df_resultados['Puntaje'] = pd.to_numeric(df_resultados['Puntaje'], errors='coerce')

# Agrupar por ID y Seccion y sumar los puntajes
df_resultados_agrupados = df_resultados.groupby(['ID', 'Empresa', 'Tamaño', 'Pais', 'Seccion'], as_index=False)['Puntaje'].sum()

# Incrustar el puntaje de la sección
df_resultados_agrupados['Puntaje Seccion'] = df_resultados_agrupados['Seccion'].map(secciones_puntaje)

# Crear un archivo Excel con los resultados
output = BytesIO()
df_resultados_agrupados.to_excel(output, index=False, engine='openpyxl')
# guardar archivo excel resultado.xlsx 
#with open("tabla_radar.xlsx", "wb") as file_resultado:
#    file_resultado.write(output.getvalue())

# Subir el archivo Excel a SharePoint
folderDocumentosCompartidos.upload_file("tabla_radar.xlsx", output.getvalue()).execute_query()

print("Archivo 'tabla_radar.xlsx' subido con éxito a SharePoint.")