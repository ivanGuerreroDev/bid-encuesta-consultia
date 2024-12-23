import pandas as pd
import re
from flask import Flask, jsonify
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import os
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor
import numpy as np

load_dotenv()
app = Flask(__name__)

def download_sharepoint_file(ctx, url, filename):
    """Helper function to download files from SharePoint"""
    with open(filename, "wb") as file:
        ctx.web.get_file_by_server_relative_url(url).download(file).execute_query()

def process_empresa_data(df_encuesta):
    """Process company data efficiently using vectorized operations"""
    empresas = {}
    
    # Extraer códigos de preguntas una sola vez
    df_encuesta['pregunta_code'] = df_encuesta.columns.map(
        lambda x: re.search(r"\[([A-Za-z0-9_.]+)\]", str(x)).group(1) if isinstance(x, str) and re.search(r"\[([A-Za-z0-9_.]+)\]", str(x)) else ""
    )
    
    # Extraer códigos de respuestas una sola vez
    df_encuesta['respuesta_code'] = df_encuesta.apply(
        lambda x: re.search(r"\[([A-Za-z0-9_.]+)\]", str(x)).group(1) if isinstance(x, str) and re.search(r"\[([A-Za-z0-9_.]+)\]", str(x)) else "",
        axis=1
    )
    
    for _, row in df_encuesta.iterrows():
        id_empresa = row['ID']
        if id_empresa not in empresas:
            empresas[id_empresa] = {}
            
        # Asignar empresa
        mask_pg001 = row['pregunta_code'] == "Pg001"
        if any(mask_pg001):
            empresas[id_empresa]['Empresa'] = row[mask_pg001].iloc[0]
            
        # Asignar país
        if "Pg011.01" in row['respuesta_code'].values:
            empresas[id_empresa]['Pais'] = 'Costa Rica'
        elif "Pg011.02" in row['respuesta_code'].values:
            empresas[id_empresa]['Pais'] = 'Panamá'
            
    return empresas

@app.route('/generate-excel', methods=['GET'])
def generate_excel():
    try:
        # Configuración de conexión a SharePoint
        sharepoint_site = "https://marketingconsultia.sharepoint.com/sites/BIDCiberseguridad"
        sharepoint_urls = {
            'encuesta': "/sites/BIDCiberseguridad/Documentos%20compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx",
            'puntajes': "/sites/BIDCiberseguridad/Documentos%20compartidos/puntajes.xlsx"
        }
        
        # Conectar a SharePoint
        ctx = ClientContext(sharepoint_site).with_credentials(
            UserCredential(os.getenv("SHAREPOINT_USER"), os.getenv("SHAREPOINT_PASSWORD"))
        )
        
        # Descargar archivos en paralelo
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = {
                executor.submit(download_sharepoint_file, ctx, sharepoint_urls['encuesta'], "df_encuesta.xlsx"): 'encuesta',
                executor.submit(download_sharepoint_file, ctx, sharepoint_urls['puntajes'], "df_puntajes.xlsx"): 'puntajes'
            }
        
        # Cargar DataFrames
        df_encuesta = pd.read_excel("df_encuesta.xlsx", sheet_name="Form1")
        df_puntajes = pd.read_excel("df_puntajes.xlsx")
        df_puntajes.columns = df_puntajes.columns.str.strip()
        
        # Procesar datos de empresas
        empresas = process_empresa_data(df_encuesta)
        
        # Calcular secciones_puntaje usando groupby
        secciones_puntaje = df_puntajes.groupby('Seccion')['Puntaje'].sum().to_dict()
        
        # Preparar resultados usando vectorización
        resultados = []
        for index, row in df_encuesta.iterrows():
            empresa_data = empresas.get(row['ID'], {})
            if not isinstance(empresa_data, dict) or 'Empresa' not in empresa_data:
                continue
                
            for pregunta in df_encuesta.columns:
                respuesta = row[pregunta]
                if not isinstance(respuesta, str):
                    continue
                    
                respuesta_code = re.search(r"\[([A-Za-z0-9_.]+)\]", respuesta)
                if not respuesta_code:
                    continue
                    
                respuesta_code = respuesta_code.group(1)
                
                # Buscar puntaje eficientemente
                puntaje_match = df_puntajes[
                    (df_puntajes['Respuesta Pequeña'] == respuesta_code) |
                    (df_puntajes['Respuesta Mediana'] == respuesta_code)
                ]
                
                if not puntaje_match.empty:
                    resultados.append({
                        'ID': row['ID'],
                        'Empresa': empresa_data['Empresa'],
                        'Tamaño': 'Pequeña' if respuesta_code == puntaje_match['Respuesta Pequeña'].iloc[0] else 'Mediana',
                        'Pais': empresa_data.get('Pais', ''),
                        'Puntaje': puntaje_match['Puntaje'].iloc[0],
                        'Seccion': puntaje_match['Seccion'].iloc[0],
                        'Puntaje Seccion': secciones_puntaje[puntaje_match['Seccion'].iloc[0]]
                    })
        
        # Crear DataFrame final y agrupar resultados
        df_resultados = pd.DataFrame(resultados)
        df_resultados['Puntaje'] = pd.to_numeric(df_resultados['Puntaje'], errors='coerce')
        df_resultados_agrupados = df_resultados.groupby(
            ['ID', 'Empresa', 'Tamaño', 'Pais', 'Seccion'],
            as_index=False
        )['Puntaje'].sum()
        
        # Agregar Puntaje Seccion
        df_resultados_agrupados['Puntaje Seccion'] = df_resultados_agrupados['Seccion'].map(secciones_puntaje)
        
        # Guardar y subir resultado
        output = BytesIO()
        df_resultados_agrupados.to_excel(output, index=False, engine='openpyxl')
        
        # Subir archivo a SharePoint
        ctx.web.get_folder_by_server_relative_url("/sites/BIDCiberseguridad/Documentos%20compartidos") \
           .upload_file("tabla_radar.xlsx", output.getvalue()) \
           .execute_query()
        
        # Limpiar archivos temporales
        for filename in ["df_encuesta.xlsx", "df_puntajes.xlsx"]:
            if os.path.exists(filename):
                os.remove(filename)
        
        return jsonify({"message": "Excel generado correctamente"}), 200
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=80)