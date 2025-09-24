import pandas as pd
import re
from flask import Flask, jsonify
import requests
from io import BytesIO
import os
from dotenv import load_dotenv
import tempfile
from collections import defaultdict

load_dotenv()
app = Flask(__name__)

def get_access_token():
    """Obtener token de acceso usando Client Credentials Flow para Microsoft Graph"""
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    
    # URL para obtener el token
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    # Datos para la solicitud
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    # Solicitar el token
    response = requests.post(token_url, data=token_data)
    response.raise_for_status()
    
    return response.json()['access_token']

def download_sharepoint_file(access_token, site_id, file_path, temp_dir):
    """Descargar archivo desde SharePoint usando Microsoft Graph API"""
    try:
        # Construir la URL correcta para Microsoft Graph
        # Usar /sites/{site-id}/drives/{drive-id}/root:/{item-path}:/content
        # Primero necesitamos obtener el drive ID del sitio
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }
        
        # Obtener información de los drives
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives_data = drives_response.json()
        
        # Buscar el drive principal (Documents)
        drive_id = None
        for drive in drives_data.get('value', []):
            if drive.get('name') == 'Documents' or 'document' in drive.get('name', '').lower():
                drive_id = drive['id']
                break
        
        if not drive_id and drives_data.get('value'):
            # Si no encontramos el drive de Documents, usar el primero disponible
            drive_id = drives_data['value'][0]['id']
        
        if not drive_id:
            raise Exception("No se pudo encontrar un drive válido en el sitio de SharePoint")
        
        print(f"Debug - Drive ID: {drive_id}")
        
        # Construir la URL del archivo
        # Remover "Documentos compartidos/" del path ya que es parte del drive
        clean_file_path = file_path.replace('Documentos compartidos/', '')
        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{clean_file_path}:/content"
        
        print(f"Debug - File URL: {file_url}")
        
        # Descargar el archivo
        file_response = requests.get(file_url, headers=headers)
        file_response.raise_for_status()
        
        # Guardar el archivo en el directorio temporal
        filename = os.path.basename(file_path)
        local_path = os.path.join(temp_dir, filename)
        
        with open(local_path, 'wb') as f:
            f.write(file_response.content)
        
        print(f"Debug - Archivo descargado: {local_path}")
        return local_path
        
    except Exception as e:
        print(f"Error downloading file from {file_path}: {str(e)}")
        raise

def upload_sharepoint_file(access_token, site_id, file_content, filename, folder_path=""):
    """Subir archivo a SharePoint usando Microsoft Graph API"""
    try:
        # Obtener el drive ID del sitio
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }
        
        # Obtener información de los drives
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives_data = drives_response.json()
        
        # Buscar el drive principal (Documents)
        drive_id = None
        for drive in drives_data.get('value', []):
            if drive.get('name') == 'Documents' or 'document' in drive.get('name', '').lower():
                drive_id = drive['id']
                break
        
        if not drive_id and drives_data.get('value'):
            # Si no encontramos el drive de Documents, usar el primero disponible
            drive_id = drives_data['value'][0]['id']
        
        if not drive_id:
            raise Exception("No se pudo encontrar un drive válido en el sitio de SharePoint")
        
        print(f"Debug - Upload Drive ID: {drive_id}")
        
        # Construir la URL para subir el archivo
        # Si hay folder_path, incluirlo en la ruta
        if folder_path:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{filename}:/content"
        else:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{filename}:/content"
        
        print(f"Debug - Upload URL: {upload_url}")
        
        # Headers para la subida
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        
        # Subir el archivo
        upload_response = requests.put(upload_url, headers=upload_headers, data=file_content)
        upload_response.raise_for_status()
        
        print(f"Debug - Archivo subido exitosamente: {filename}")
        return upload_response.json()
        
    except Exception as e:
        print(f"Error uploading file {filename}: {str(e)}")
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
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # URLs de los archivos en SharePoint
            sharepoint_files = {
                'encuesta': 'Documentos compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx',
                'puntajes': 'Documentos compartidos/puntajes.xlsx'
            }
            
            # Obtener token de acceso para Microsoft Graph
            access_token = get_access_token()
            
            # Site ID de SharePoint (necesitamos obtenerlo primero)
            site_url = "marketingconsultia.sharepoint.com:/sites/BIDCiberseguridad"
            site_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Accept': 'application/json'
            }
            
            # Obtener información del sitio para conseguir el site_id
            site_response = requests.get(site_info_url, headers=headers)
            site_response.raise_for_status()
            site_id = site_response.json()['id']
            
            print(f"Debug - Site ID: {site_id}")
            
            try:
                # Descargar archivos desde SharePoint usando Microsoft Graph
                encuesta_path = download_sharepoint_file(access_token, site_id, sharepoint_files['encuesta'], temp_dir)
                puntajes_path = download_sharepoint_file(access_token, site_id, sharepoint_files['puntajes'], temp_dir)
                
                # Leer los archivos Excel
                df_encuesta = pd.read_excel(encuesta_path, sheet_name="Form1")
                df_puntajes = pd.read_excel(puntajes_path)
                
                print(f"Debug - Columnas de encuesta: {list(df_encuesta.columns)}")
                print(f"Debug - Columnas de puntajes: {list(df_puntajes.columns)}")
                print(f"Debug - Primeras filas de encuesta:")
                print(df_encuesta.head())
                
                # Verificar si existe la columna ID, si no, usar la primera columna como ID
                if 'ID' not in df_encuesta.columns:
                    # Usar la primera columna como ID
                    primera_columna = df_encuesta.columns[0]
                    df_encuesta = df_encuesta.rename(columns={primera_columna: 'ID'})
                    print(f"Debug - Renombrada columna '{primera_columna}' a 'ID'")
                
                # Procesar los datos
                df_empresas = process_empresa_data(df_encuesta)
                
                # Crear el archivo Excel final
                output_path = os.path.join(temp_dir, 'resultado_final.xlsx')
                output = BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_empresas_df = pd.DataFrame.from_dict(df_empresas, orient='index').reset_index()
                    df_empresas_df.columns = ['ID'] + list(df_empresas_df.columns[1:])
                    df_empresas_df.to_excel(writer, sheet_name='Empresas', index=False)
                    df_puntajes.to_excel(writer, sheet_name='Puntajes', index=False)
                
                # Obtener el contenido del archivo en memoria
                output.seek(0)
                file_content = output.getvalue()
                
                # Subir archivo a SharePoint usando Microsoft Graph
                upload_result = upload_sharepoint_file(
                    access_token, 
                    site_id, 
                    file_content, 
                    "tabla_radar.xlsx"
                )
                
                return jsonify({
                    "message": "Archivo Excel generado y subido exitosamente a SharePoint",
                    "empresas_procesadas": len(df_empresas),
                    "total_puntajes": len(df_puntajes),
                    "archivo_subido": "tabla_radar.xlsx",
                    "upload_info": upload_result.get('name', 'tabla_radar.xlsx')
                })
                
            except Exception as download_error:
                print(f"Error en la descarga de archivos: {str(download_error)}")
                return jsonify({"error": str(download_error)}), 500
                
    except Exception as e:
        print(f"Error general: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8090)