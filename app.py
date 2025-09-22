import pandas as pd
import re
from flask import Flask, jsonify
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from io import BytesIO
import os
from dotenv import load_dotenv
import tempfile
from collections import defaultdict

load_dotenv()
app = Flask(__name__)

def download_sharepoint_file(ctx, url, temp_dir):
    """Helper function to download files from SharePoint using temporary files"""
    try:
        file_name = url.split('/')[-1]
        temp_path = os.path.join(temp_dir, file_name)
        
        file_obj = ctx.web.get_file_by_server_relative_url(url)
        ctx.load(file_obj)
        ctx.execute_query()
        
        with open(temp_path, 'wb') as local_file:
            file_response = file_obj.download(local_file)
            ctx.execute_query()
            
        return temp_path
        
    except Exception as e:
        print(f"Error downloading file from {url}: {str(e)}")
        raise

def process_empresa_data(df_encuesta):
    """Process company data efficiently using vectorized operations"""
    empresas = {}
    
    for _, row in df_encuesta.iterrows():
        id_empresa = row['ID']
        if id_empresa not in empresas:
            empresas[id_empresa] = {'Empresa': '', 'Pais': ''}
            
        for columna, valor in row.items():
            if isinstance(columna, str) and isinstance(valor, str):
                # Buscar código de pregunta para empresa
                if 'Pg001' in columna:
                    empresas[id_empresa]['Empresa'] = valor
                # Buscar códigos de respuesta para país
                if '[Pg011.01]' in valor:
                    empresas[id_empresa]['Pais'] = 'Costa Rica'
                elif '[Pg011.02]' in valor:
                    empresas[id_empresa]['Pais'] = 'Panamá'
            
    return empresas

@app.route('/generate-excel', methods=['GET'])
def generate_excel():
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Configuración de conexión a SharePoint
            sharepoint_site = "https://marketingconsultia.sharepoint.com/sites/BIDCiberseguridad"
            sharepoint_urls = {
                'encuesta': "/sites/BIDCiberseguridad/Documentos%20compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx",
                'puntajes': "/sites/BIDCiberseguridad/Documentos%20compartidos/puntajes.xlsx"
            }
            
            ctx = ClientContext(sharepoint_site).with_credentials(
                ClientCredential(os.getenv("CLIENT_ID"), os.getenv("CLIENT_SECRET"))
            )
            
            try:
                encuesta_path = download_sharepoint_file(ctx, sharepoint_urls['encuesta'], temp_dir)
                puntajes_path = download_sharepoint_file(ctx, sharepoint_urls['puntajes'], temp_dir)
            except Exception as e:
                print(f"Error en la descarga de archivos: {str(e)}")
                raise
            
            try:
                df_encuesta = pd.read_excel(encuesta_path, sheet_name="Form1", engine='openpyxl')
                df_puntajes = pd.read_excel(puntajes_path, engine='openpyxl')
            except Exception as e:
                print(f"Error leyendo archivos Excel: {str(e)}")
                raise
                
            df_puntajes.columns = df_puntajes.columns.str.strip()
            
            # Procesar datos de empresas
            empresas = process_empresa_data(df_encuesta)
            # Calcular secciones_puntaje usando groupby
            secciones_puntaje = df_puntajes.groupby('Seccion')['Puntaje'].sum().to_dict()
            # Calcula puntaje por tamaño suma si la fila tiene valor en la columna Pregunta Pequeña o Pregunta Mediana.
            secciones_puntaje_pequena = defaultdict(int)
            secciones_puntaje_mediana = defaultdict(int)

            for _, row in df_puntajes.iterrows():
                if row['Respuesta Pequeña'] and not pd.isna(row['Respuesta Pequeña']):
                    secciones_puntaje_pequena[row['Seccion']] += row['Puntaje']
                if row['Respuesta Mediana'] and not pd.isna(row['Respuesta Mediana']):
                    secciones_puntaje_mediana[row['Seccion']] += row['Puntaje']
            # Preparar resultados
            resultados = []
            # Procesar cada respuesta de la encuesta
            for _, row_encuesta in df_encuesta.iterrows():
                id_empresa = row_encuesta['ID']
                empresa_info = empresas.get(id_empresa, {})
                if not empresa_info.get('Empresa'):
                    continue
                
                # Procesar cada respuesta de la fila
                for columna, respuesta in row_encuesta.items():
                    if not isinstance(respuesta, str):
                        continue
                        
                    respuesta_match = re.search(r"\[([A-Za-z0-9_.]+)\]", str(respuesta))
                    if not respuesta_match:
                        continue
                        
                    respuesta_code = respuesta_match.group(1)
                    
                    # Buscar en df_puntajes
                    puntaje_match = df_puntajes[
                        (df_puntajes['Respuesta Pequeña'] == respuesta_code) |
                        (df_puntajes['Respuesta Mediana'] == respuesta_code)
                    ]
                    
                    if not puntaje_match.empty:
                        tamano = 'Pequeña' if respuesta_code == puntaje_match['Respuesta Pequeña'].iloc[0] else 'Mediana'
                        seccion = puntaje_match['Seccion'].iloc[0]
                        
                        resultados.append({
                            'ID': id_empresa,
                            'Empresa': empresa_info.get('Empresa', ''),
                            'Tamaño': tamano,
                            'Pais': empresa_info.get('Pais', ''),
                            'Puntaje': float(puntaje_match['Puntaje'].iloc[0]),
                            'Seccion': seccion,
                            'Puntaje Seccion': secciones_puntaje_pequena[seccion] if tamano == 'Pequeña' else secciones_puntaje_mediana[seccion]
                        })
            
            # Crear DataFrame y agrupar resultados
            if not resultados:
                raise ValueError("No se encontraron resultados para procesar")
                
            df_resultados = pd.DataFrame(resultados)
            
            # Agrupar resultados por las columnas necesarias y sumar puntajes
            df_resultados_agrupados = df_resultados.groupby(
                ['ID', 'Empresa', 'Tamaño', 'Pais', 'Seccion'],
                as_index=False
            ).agg({
                'Puntaje': 'sum',
                'Puntaje Seccion': 'first'  # Tomamos el primer valor ya que es el mismo para cada sección
            })

            # Calcular puntaje total por empresa
            df_puntaje_total = df_resultados_agrupados.groupby(['ID', 'Empresa'], as_index=False).agg({
                'Puntaje': 'sum',
                'Puntaje Seccion': 'sum'
            })
            # Calcular porcentaje total
            df_puntaje_total['Porcentaje Total'] = df_puntaje_total['Puntaje'] / df_puntaje_total['Puntaje Seccion']
            
            # Calcular puntaje por pais promedia Puntaje
            df_puntaje_total_pais = df_resultados_agrupados.groupby(['Pais', 'Seccion'], as_index=False).agg({
                'Puntaje': 'mean',
                'Puntaje Seccion': 'first'
            })
            
            # Guardar resultado en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resultados_agrupados.to_excel(writer, index=False)
                df_puntaje_total[['ID', 'Empresa', 'Porcentaje Total']].to_excel(writer, index=False, sheet_name='General por empresas')
                df_puntaje_total_pais.to_excel(writer, index=False, sheet_name='General por paises')
            
            output.seek(0)
            
            # Subir archivo a SharePoint
            folder = ctx.web.get_folder_by_server_relative_url("/sites/BIDCiberseguridad/Documentos%20compartidos")
            folder.upload_file("tabla_radar.xlsx", output.getvalue()).execute_query()
            
            return jsonify({"message": "Excel generado correctamente"}), 200
            
        except Exception as e:
            return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8090)