import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import re
import requests
from io import BytesIO
import os
import json
import tempfile
from collections import defaultdict
import threading
from datetime import datetime

class ConfigManager:
    """Gestiona la configuración de la aplicación con persistencia"""
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self):
        """Cargar configuración desde archivo JSON"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error cargando configuración: {e}")
                return self.get_default_config()
        return self.get_default_config()
    
    def save_config(self, config):
        """Guardar configuración en archivo JSON"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            self.config = config
            return True
        except Exception as e:
            print(f"Error guardando configuración: {e}")
            return False
    
    def get_default_config(self):
        """Configuración por defecto"""
        return {
            'tenant_id': '',
            'client_id': '',
            'client_secret': '',
            'site_url': 'marketingconsultia.sharepoint.com:/sites/BIDCiberseguridad',
            'encuesta_path': 'Documentos compartidos/Encuesta sobre brechas digitales en ciberseguridad en PYMEs.xlsx',
            'puntajes_path': 'Documentos compartidos/puntajes.xlsx',
            'debug_mode': False,
            'output_filename': 'tabla_radar.xlsx'
        }
    
    def get(self, key, default=None):
        """Obtener valor de configuración"""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """Establecer valor de configuración"""
        self.config[key] = value


class SharePointProcessor:
    """Procesador de datos de SharePoint"""
    
    def __init__(self, config_manager):
        self.config = config_manager
        self.debug_dir = None
        
        # Crear directorio debug si está activado
        if self.config.get('debug_mode'):
            self.debug_dir = os.path.join(os.getcwd(), "debug_files")
            if not os.path.exists(self.debug_dir):
                os.makedirs(self.debug_dir)
    
    def get_access_token(self):
        """Obtener token de acceso usando Client Credentials Flow"""
        tenant_id = self.config.get('tenant_id')
        client_id = self.config.get('client_id')
        client_secret = self.config.get('client_secret')
        
        if not all([tenant_id, client_id, client_secret]):
            raise ValueError("Faltan credenciales de configuración")
        
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        
        token_data = {
            'grant_type': 'client_credentials',
            'client_id': client_id,
            'client_secret': client_secret,
            'scope': 'https://graph.microsoft.com/.default'
        }
        
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        
        return response.json()['access_token']
    
    def download_sharepoint_file(self, access_token, site_id, file_path, temp_dir):
        """Descargar archivo desde SharePoint"""
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }
        
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives_data = drives_response.json()
        
        drive_id = None
        for drive in drives_data.get('value', []):
            if drive.get('name') == 'Documents' or 'document' in drive.get('name', '').lower():
                drive_id = drive['id']
                break
        
        if not drive_id and drives_data.get('value'):
            drive_id = drives_data['value'][0]['id']
        
        if not drive_id:
            raise Exception("No se pudo encontrar un drive válido")
        
        clean_file_path = file_path.replace('Documentos compartidos/', '')
        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{clean_file_path}:/content"
        
        file_response = requests.get(file_url, headers=headers)
        file_response.raise_for_status()
        
        filename = os.path.basename(file_path)
        local_path = os.path.join(temp_dir, filename)
        
        with open(local_path, 'wb') as f:
            f.write(file_response.content)
        
        if self.debug_dir:
            debug_path = os.path.join(self.debug_dir, filename)
            with open(debug_path, 'wb') as f:
                f.write(file_response.content)
        
        return local_path
    
    def upload_sharepoint_file(self, access_token, site_id, file_content, filename):
        """Subir archivo a SharePoint"""
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }
        
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives_data = drives_response.json()
        
        drive_id = None
        for drive in drives_data.get('value', []):
            if drive.get('name') == 'Documents' or 'document' in drive.get('name', '').lower():
                drive_id = drive['id']
                break
        
        if not drive_id and drives_data.get('value'):
            drive_id = drives_data['value'][0]['id']
        
        if not drive_id:
            raise Exception("No se pudo encontrar un drive válido")
        
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{filename}:/content"
        
        upload_headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        
        upload_response = requests.put(upload_url, headers=upload_headers, data=file_content)
        upload_response.raise_for_status()
        
        return upload_response.json()
    
    def process_empresa_data(self, df_encuesta):
        """Procesar datos de empresas"""
        empresas = {}
        
        for _, row in df_encuesta.iterrows():
            id_empresa = row['ID']
            if id_empresa not in empresas:
                empresas[id_empresa] = {'Empresa': '', 'Pais': '', 'tamano_empresa': 'Desconocido'}
                
            for columna, valor in row.items():
                if isinstance(columna, str) and isinstance(valor, str):
                    if 'Pg001' in columna:
                        empresas[id_empresa]['Empresa'] = valor
                    if '[Pg011.01]' in valor:
                        empresas[id_empresa]['Pais'] = 'Costa Rica'
                    elif '[Pg011.02]' in valor:
                        empresas[id_empresa]['Pais'] = 'Panamá'
                    
                    # Procesar tamaño de empresa
                    if empresas[id_empresa]['Pais'] == 'Panamá':
                        if '[Pa012.01]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Micro'
                        elif '[Pa012.02]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Pequeña'
                        elif '[Pa012.03]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Mediana'
                        elif '[Pa012.04]' in valor or '[Pa012.05]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Grande'
                    elif empresas[id_empresa]['Pais'] == 'Costa Rica':
                        if '[Pc012.01]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Micro'
                        elif '[Pc012.02]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Pequeña'
                        elif '[Pc012.03]' in valor or '[Pc012.04]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Mediana'
                        elif '[Pc012.05]' in valor or '[Pc012.06]' in valor:
                            empresas[id_empresa]['tamano_empresa'] = 'Grande'
        
        return empresas
    
    def process_data(self, log_callback=None):
        """Procesar datos y generar Excel"""
        def log(message):
            if log_callback:
                log_callback(message)
            print(message)
        
        with tempfile.TemporaryDirectory() as temp_dir:
            log("Obteniendo token de acceso...")
            access_token = self.get_access_token()
            
            log("Obteniendo información del sitio de SharePoint...")
            site_url = self.config.get('site_url')
            site_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Accept': 'application/json'
            }
            
            site_response = requests.get(site_info_url, headers=headers)
            site_response.raise_for_status()
            site_id = site_response.json()['id']
            
            log(f"Site ID obtenido: {site_id}")
            
            log("Descargando archivo de encuesta...")
            encuesta_path = self.download_sharepoint_file(
                access_token, site_id, 
                self.config.get('encuesta_path'), 
                temp_dir
            )
            
            log("Descargando archivo de puntajes...")
            puntajes_path = self.download_sharepoint_file(
                access_token, site_id, 
                self.config.get('puntajes_path'), 
                temp_dir
            )
            
            log("Leyendo archivos Excel...")
            df_encuesta = pd.read_excel(encuesta_path, sheet_name="Form1")
            df_puntajes = pd.read_excel(puntajes_path)
            
            if 'ID' not in df_encuesta.columns:
                primera_columna = df_encuesta.columns[0]
                df_encuesta = df_encuesta.rename(columns={primera_columna: 'ID'})
                log(f"Renombrada columna '{primera_columna}' a 'ID'")
            
            log("Procesando datos de empresas...")
            empresas = self.process_empresa_data(df_encuesta)
            
            log("Calculando puntajes por sección...")
            secciones_puntaje = df_puntajes.groupby('Seccion')['Puntaje'].sum().to_dict()
            
            secciones_puntaje_pequena = defaultdict(int)
            secciones_puntaje_mediana = defaultdict(int)
            
            for _, row in df_puntajes.iterrows():
                if row['Respuesta Pequeña'] and not pd.isna(row['Respuesta Pequeña']):
                    secciones_puntaje_pequena[row['Seccion']] += row['Puntaje']
                if row['Respuesta Mediana'] and not pd.isna(row['Respuesta Mediana']):
                    secciones_puntaje_mediana[row['Seccion']] += row['Puntaje']
            
            log("Procesando respuestas...")
            resultados = []
            
            for _, row_encuesta in df_encuesta.iterrows():
                id_empresa = row_encuesta['ID']
                empresa_info = empresas.get(id_empresa, {})
                if not empresa_info.get('Empresa'):
                    continue
                
                for columna, respuesta in row_encuesta.items():
                    if not isinstance(respuesta, str):
                        continue
                        
                    respuesta_match = re.search(r"\[([A-Za-z0-9_.]+)\]", str(respuesta))
                    if not respuesta_match:
                        continue
                        
                    respuesta_code = respuesta_match.group(1)
                    
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
                            'Tamaño de empresa': empresa_info.get('tamano_empresa', 'Desconocido'),
                            'Pais': empresa_info.get('Pais', ''),
                            'Puntaje': float(puntaje_match['Puntaje'].iloc[0]),
                            'Seccion': seccion,
                            'Puntaje Seccion': secciones_puntaje_pequena[seccion] if tamano == 'Pequeña' else secciones_puntaje_mediana[seccion]
                        })
            
            if not resultados:
                raise ValueError("No se encontraron resultados para procesar")
            
            log(f"Generando archivo Excel con {len(resultados)} resultados...")
            df_resultados = pd.DataFrame(resultados)
            
            df_resultados_agrupados = df_resultados.groupby(
                ['ID', 'Empresa', 'Tamaño', 'Pais', 'Seccion', 'Tamaño de empresa'],
                as_index=False
            ).agg({
                'Puntaje': 'sum',
                'Puntaje Seccion': 'first'
            })
            
            df_puntaje_total = df_resultados_agrupados.groupby(['ID', 'Empresa'], as_index=False).agg({
                'Puntaje': 'sum',
                'Puntaje Seccion': 'sum'
            })
            df_puntaje_total['Porcentaje Total'] = df_puntaje_total['Puntaje'] / df_puntaje_total['Puntaje Seccion']
            
            df_puntaje_total_pais = df_resultados_agrupados.groupby(['Pais', 'Seccion'], as_index=False).agg({
                'Puntaje': 'mean',
                'Puntaje Seccion': 'first'
            })
            
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resultados_agrupados.to_excel(writer, index=False)
                df_puntaje_total[['ID', 'Empresa', 'Porcentaje Total']].to_excel(writer, index=False, sheet_name='General por empresas')
                df_puntaje_total_pais.to_excel(writer, index=False, sheet_name='General por paises')
            
            output.seek(0)
            file_content = output.getvalue()
            
            if self.debug_dir:
                debug_excel_path = os.path.join(self.debug_dir, self.config.get('output_filename'))
                with open(debug_excel_path, 'wb') as f:
                    f.write(file_content)
                log(f"Archivo guardado localmente en: {debug_excel_path}")
            
            log("Subiendo archivo a SharePoint...")
            upload_result = self.upload_sharepoint_file(
                access_token, 
                site_id, 
                file_content, 
                self.config.get('output_filename')
            )
            
            return {
                "success": True,
                "empresas_procesadas": len(empresas),
                "total_resultados": len(df_resultados_agrupados),
                "archivo_subido": self.config.get('output_filename')
            }


class ConfigWindow(tk.Toplevel):
    """Ventana de configuración"""
    
    def __init__(self, parent, config_manager):
        super().__init__(parent)
        self.config_manager = config_manager
        self.title("Configuración")
        self.geometry("600x500")
        self.resizable(True, True)
        
        # Hacer la ventana modal
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.load_current_config()
    
    def create_widgets(self):
        """Crear widgets de la ventana"""
        # Frame principal con scroll
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar expansión del frame principal
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Configuración de SharePoint", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Campos de configuración
        self.entries = {}
        
        fields = [
            ('tenant_id', 'Tenant ID:', False),
            ('client_id', 'Client ID:', False),
            ('client_secret', 'Client Secret:', True),
            ('site_url', 'Site URL:', False),
            ('encuesta_path', 'Ruta Encuesta:', False),
            ('puntajes_path', 'Ruta Puntajes:', False),
            ('output_filename', 'Nombre Archivo Salida:', False)
        ]
        
        row = 1
        for field_name, label_text, is_password in fields:
            label = ttk.Label(main_frame, text=label_text)
            label.grid(row=row, column=0, sticky=tk.W, pady=5)
            
            entry = ttk.Entry(main_frame, width=50, show='*' if is_password else '')
            entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
            self.entries[field_name] = entry
            
            row += 1
        
        # Checkbox para modo debug
        self.debug_var = tk.BooleanVar()
        debug_check = ttk.Checkbutton(main_frame, text="Modo Debug (guardar archivos localmente)", 
                                     variable=self.debug_var)
        debug_check.grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row+1, column=0, columnspan=2, pady=20)
        
        save_btn = ttk.Button(button_frame, text="Guardar", command=self.save_config)
        save_btn.grid(row=0, column=0, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancelar", command=self.destroy)
        cancel_btn.grid(row=0, column=1, padx=5)
        
        # Configurar expansión de columnas en main_frame
        main_frame.columnconfigure(1, weight=1)
    
    def load_current_config(self):
        """Cargar configuración actual en los campos"""
        for field_name, entry in self.entries.items():
            value = self.config_manager.get(field_name, '')
            entry.delete(0, tk.END)
            entry.insert(0, value)
        
        self.debug_var.set(self.config_manager.get('debug_mode', False))
    
    def save_config(self):
        """Guardar configuración"""
        new_config = {}
        for field_name, entry in self.entries.items():
            new_config[field_name] = entry.get()
        
        new_config['debug_mode'] = self.debug_var.get()
        
        if self.config_manager.save_config(new_config):
            messagebox.showinfo("Éxito", "Configuración guardada correctamente")
            self.destroy()
        else:
            messagebox.showerror("Error", "No se pudo guardar la configuración")


class MainApplication(tk.Tk):
    """Aplicación principal"""
    
    def __init__(self):
        super().__init__()
        
        self.title("Generador de Reportes - BID Ciberseguridad")
        self.geometry("800x600")
        
        self.config_manager = ConfigManager()
        self.processor = SharePointProcessor(self.config_manager)
        
        self.create_widgets()
        self.check_config()
    
    def create_widgets(self):
        """Crear widgets de la interfaz principal"""
        # Frame superior con título y botones
        top_frame = ttk.Frame(self, padding="10")
        top_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Configurar expansión de columna en top_frame
        top_frame.columnconfigure(0, weight=1)
        
        title_label = ttk.Label(top_frame, text="Generador de Reportes de Encuestas", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Botones de acción
        button_frame = ttk.Frame(top_frame)
        button_frame.grid(row=0, column=1, sticky=tk.E)
        
        config_btn = ttk.Button(button_frame, text="⚙ Configuración", 
                               command=self.open_config)
        config_btn.grid(row=0, column=0, padx=5)
        
        self.process_btn = ttk.Button(button_frame, text="▶ Generar Reporte", 
                                     command=self.start_processing,
                                     style='Accent.TButton')
        self.process_btn.grid(row=0, column=1, padx=5)
        
        # Separador
        ttk.Separator(self, orient='horizontal').grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Frame de información
        info_frame = ttk.LabelFrame(self, text="Estado", padding="10")
        info_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=5)
        
        # Área de log
        log_label = ttk.Label(info_frame, text="Registro de actividad:")
        log_label.grid(row=0, column=0, sticky=tk.W)
        
        self.log_text = scrolledtext.ScrolledText(info_frame, height=20, width=90)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        
        # Barra de progreso
        self.progress_var = tk.StringVar(value="Listo")
        progress_label = ttk.Label(self, textvariable=self.progress_var)
        progress_label.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
        
        self.progress_bar = ttk.Progressbar(self, mode='indeterminate')
        self.progress_bar.grid(row=4, column=0, sticky=(tk.W, tk.E), padx=10, pady=(0, 10))
        
        # Configurar pesos de las filas/columnas para que se expandan
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(1, weight=1)
        
        # Log inicial
        self.log("Aplicación iniciada")
        self.log(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    def check_config(self):
        """Verificar si la configuración está completa"""
        required_fields = ['tenant_id', 'client_id', 'client_secret']
        missing_fields = [f for f in required_fields if not self.config_manager.get(f)]
        
        if missing_fields:
            self.log("⚠ Configuración incompleta. Por favor, configura las credenciales.")
            messagebox.showwarning("Configuración incompleta", 
                                  "Por favor, configura las credenciales antes de generar reportes.")
    
    def log(self, message):
        """Agregar mensaje al log"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.update_idletasks()
    
    def open_config(self):
        """Abrir ventana de configuración"""
        ConfigWindow(self, self.config_manager)
    
    def start_processing(self):
        """Iniciar procesamiento en un hilo separado"""
        self.process_btn.config(state='disabled')
        self.progress_bar.start()
        self.progress_var.set("Procesando...")
        
        # Ejecutar en un hilo separado para no bloquear la UI
        thread = threading.Thread(target=self.process_data, daemon=True)
        thread.start()
    
    def process_data(self):
        """Procesar datos (ejecutado en hilo separado)"""
        try:
            self.log("\n" + "="*50)
            self.log("Iniciando generación de reporte...")
            
            result = self.processor.process_data(log_callback=self.log)
            
            self.log("="*50)
            self.log("✓ Proceso completado exitosamente")
            self.log(f"  - Empresas procesadas: {result['empresas_procesadas']}")
            self.log(f"  - Total resultados: {result['total_resultados']}")
            self.log(f"  - Archivo subido: {result['archivo_subido']}")
            
            self.after(0, lambda: messagebox.showinfo("Éxito", 
                f"Reporte generado exitosamente\n\n"
                f"Empresas procesadas: {result['empresas_procesadas']}\n"
                f"Archivo: {result['archivo_subido']}"))
            
        except Exception as e:
            error_msg = f"✗ Error: {str(e)}"
            self.log(error_msg)
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        finally:
            self.after(0, self.finish_processing)
    
    def finish_processing(self):
        """Finalizar procesamiento (ejecutado en hilo principal)"""
        self.progress_bar.stop()
        self.process_btn.config(state='normal')
        self.progress_var.set("Listo")


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
