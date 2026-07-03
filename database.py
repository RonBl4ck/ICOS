# -*- coding: utf-8 -*-
"""
Módulo de Base de Datos
Maneja la conexión a Google Sheets a través de gspread.
Busca LCLs en 4 hojas diferentes del libro de datos externo y
registra el historial en la hoja de Registro.
Si no hay credenciales, cae en modo local automático (db_local.xlsx).
"""

import os
import json
import pandas as pd
import streamlit as st
import datetime
from google.oauth2.service_account import Credentials
import gspread

# IDs reales de las hojas de cálculo compartidos en tu script
DATA_SPREADSHEET_ID = "1D2DDxIYMpxUaxJSfsHaNsX3tsvoLMow05astu2zIAV4"
HISTORY_SPREADSHEET_ID = "1YgdZK_aFa1EDAHJZHAuCHLBNcz076_326WJCPUWeFY4"
LOCAL_DB_FILE = "db_local.xlsx"

# Configuración de mapeo de columnas para la extracción de LCL y datos de las 4 hojas
HOJAS_CONFIG = [
    {
        "nombre": "ODM2026",
        "lcl_cols": [25],            # Z
        "supervisor_col": 0,         # A
        "cliente_col": 4,            # E
        "contratista_col": 31,       # AF
        "distrito_col": 20           # U
    },
    {
        "nombre": "OV_2026",
        "lcl_cols": [9, 45],         # J, AT
        "supervisor_col": 0,         # A
        "cliente_col": 5,            # F
        "contratista_col": 12,       # M
        "distrito_col": 2            # C
    },
    {
        "nombre": "OV_2025",
        "lcl_cols": [9, 36],         # J, AK
        "supervisor_col": 0,         # A
        "cliente_col": 5,            # F
        "contratista_col": 10,       # K
        "distrito_col": 2            # C
    },
    {
        "nombre": "OV_2024",
        "lcl_cols": [9, 35],         # J, AJ
        "supervisor_col": 0,         # A
        "cliente_col": 5,            # F
        "contratista_col": 10,       # K
        "distrito_col": 2            # C
    }
]

# Datos semilla locales para desarrollo/pruebas offline
MOCK_PROYECTOS = [
    {"LCL": "LCL-001", "Cliente": "Enel Distribución S.A.", "Distrito": "San Miguel", "Contratista": "Luz y Fuerza S.A.", "Supervisor": "Ing. Carlos Mendoza"},
    {"LCL": "LCL-002", "Cliente": "Pluz Energía Perú", "Distrito": "Miraflores", "Contratista": "ElecNor S.A.", "Supervisor": "Ing. Ana Gamarra"},
    {"LCL": "LCL-003", "Cliente": "Inmobiliaria Sagitario", "Distrito": "Santiago de Surco", "Contratista": "Proyectos Eléctricos SAC", "Supervisor": "Ing. Roberto Solís"},
    {"LCL": "LCL-004", "Cliente": "Saga Falabella S.A.", "Distrito": "San Isidro", "Contratista": "Instalaciones del Sur", "Supervisor": "Ing. Lucía Torres"},
    {"LCL": "LCL-005", "Cliente": "Distribuidora del Norte", "Distrito": "Los Olivos", "Contratista": "Consorcio Eléctrico Alfa", "Supervisor": "Ing. Manuel Espinoza"},
]

def verificar_modo_conexion():
    """
    Verifica si las credenciales de Google Sheets están configuradas en st.secrets.
    Retorna 'gsheets' si están listas o 'local' en caso contrario.
    """
    try:
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            gs_conf = st.secrets["connections"]["gsheets"]
            if "spreadsheet" in gs_conf or "project_id" in gs_conf:
                return "gsheets"
    except Exception:
        pass
    return "local"

def obtener_cliente_gspread():
    """
    Inicializa y retorna un cliente de gspread autorizado con las credenciales de secrets.
    """
    if verificar_modo_conexion() == "gsheets":
        try:
            gs_conf = st.secrets["connections"]["gsheets"]
            creds_info = {
                "type": "service_account",
                "project_id": gs_conf.get("project_id"),
                "private_key_id": gs_conf.get("private_key_id"),
                "private_key": gs_conf.get("private_key").replace("\\n", "\n") if gs_conf.get("private_key") else None,
                "client_email": gs_conf.get("client_email"),
                "client_id": gs_conf.get("client_id"),
                "auth_uri": gs_conf.get("auth_uri", "https://accounts.google.com/o/oauth2/auth"),
                "token_uri": gs_conf.get("token_uri", "https://oauth2.googleapis.com/token"),
                "auth_provider_x509_cert_url": gs_conf.get("auth_provider_x509_cert_url", "https://www.googleapis.com/oauth2/v1/certs"),
                "client_x509_cert_url": gs_conf.get("client_x509_cert_url")
            }
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Error al autenticar gspread: {e}")
            return None
    return None

def inicializar_db_local():
    """
    Crea el archivo db_local.xlsx con datos semilla si no existe.
    """
    if not os.path.exists(LOCAL_DB_FILE):
        df_proyectos = pd.DataFrame(MOCK_PROYECTOS)
        df_historial = pd.DataFrame(columns=[
            "ID_Revision", "Fecha", "LCL", "Cliente", "Distrito", 
            "Contratista", "Supervisor", "Tipo_Atencion", 
            "Numero_Revision", "Estado", "Observaciones", "Respuestas"
        ])
        
        with pd.ExcelWriter(LOCAL_DB_FILE, engine="openpyxl") as writer:
            df_proyectos.to_excel(writer, sheet_name="Proyectos", index=False)
            df_historial.to_excel(writer, sheet_name="Registro", index=False)

def buscar_proyecto_por_lcl(lcl):
    """
    Busca un proyecto por LCL consultando las 4 hojas en Google Sheets.
    Retorna un diccionario con los datos del proyecto o None si no existe.
    """
    modo = verificar_modo_conexion()
    lcl_clean = str(lcl).strip().upper()
    
    if modo == "gsheets":
        client = obtener_cliente_gspread()
        if client:
            try:
                # Abrir el libro de datos externo
                sh = client.open_by_key(DATA_SPREADSHEET_ID)
                
                # Iterar por cada hoja configurada
                for cfg in HOJAS_CONFIG:
                    try:
                        worksheet = sh.worksheet(cfg["nombre"])
                        data = worksheet.get_all_values()
                        
                        if not data:
                            continue
                            
                        # Recorrer filas buscando coincidencia de LCL
                        for row in data:
                            # Calcular el máximo índice de columna requerido
                            max_idx = max(
                                cfg["supervisor_col"], 
                                cfg["cliente_col"], 
                                cfg["contratista_col"], 
                                cfg["distrito_col"], 
                                *cfg["lcl_cols"]
                            )
                            if len(row) <= max_idx:
                                continue
                                
                            # Comparar en las columnas designadas
                            for lcl_col in cfg["lcl_cols"]:
                                row_lcl = str(row[lcl_col]).strip().upper()
                                if row_lcl and row_lcl == lcl_clean:
                                    # Encontrado: Extraer los datos mapeados
                                    return {
                                        "LCL": row[lcl_col].strip(),
                                        "Cliente": row[cfg["cliente_col"]].strip(),
                                        "Distrito": row[cfg["distrito_col"]].strip(),
                                        "Contratista": row[cfg["contratista_col"]].strip(),
                                        "Supervisor": row[cfg["supervisor_col"]].strip()
                                    }
                    except Exception as e:
                        # Si falla una pestaña (ej. no existe o no hay acceso), continuamos con las demás
                        continue
            except Exception as e:
                st.warning(f"Error al buscar en Google Sheets: {e}. Cambiando a base de datos local.")
                modo = "local"
                
    if modo == "local":
        inicializar_db_local()
        try:
            df = pd.read_excel(LOCAL_DB_FILE, sheet_name="Proyectos")
            df["LCL"] = df["LCL"].astype(str).str.strip().str.upper()
            resultado = df[df["LCL"] == lcl_clean]
            if not resultado.empty:
                # Retorna el proyecto original
                df_orig = pd.read_excel(LOCAL_DB_FILE, sheet_name="Proyectos")
                return df_orig.iloc[resultado.index[0]].to_dict()
        except Exception as e:
            st.error(f"Error al leer base de datos local: {e}")
            
    return None

def registrar_revision(datos_revision):
    """
    Guarda los datos de la revisión en Google Sheets (Hoja Registro) o en la base local.
    """
    modo = verificar_modo_conexion()
    
    if modo == "gsheets":
        client = obtener_cliente_gspread()
        if client:
            try:
                # Abrir el libro de historial
                sh = client.open_by_key(HISTORY_SPREADSHEET_ID)
                # Abrir pestaña Registro, o crearla si no existe
                try:
                    worksheet = sh.worksheet("Registro")
                except gspread.WorksheetNotFound:
                    worksheet = sh.add_worksheet(title="Registro", rows=1000, cols=10)
                    
                # Escribir cabecera si está vacía
                if not worksheet.get_all_values():
                    worksheet.append_row([
                        "Fecha", "LCL", "Cliente", "Nº Revisión", "Distrito", 
                        "Contratista", "Tipo de atención", "Supervisor", "Estado",
                        "ID_Revision", "Observaciones", "Respuestas"
                    ])
                
                # Escribir la fila respetando el orden original (iniciando con Fecha en columna A)
                worksheet.append_row([
                    datos_revision["Fecha"],
                    datos_revision["LCL"],
                    datos_revision["Cliente"],
                    int(datos_revision["Numero_Revision"]),
                    datos_revision["Distrito"],
                    datos_revision["Contratista"],
                    datos_revision["Tipo_Atencion"],
                    datos_revision["Supervisor"],
                    datos_revision["Estado"],
                    datos_revision["ID_Revision"],
                    datos_revision["Observaciones"],
                    datos_revision["Respuestas"]
                ])
                return True
            except Exception as e:
                st.warning(f"Error al escribir en Google Sheets: {e}. Guardando localmente.")
                modo = "local"
                
    if modo == "local":
        inicializar_db_local()
        try:
            # Estructurar fila local con el orden de columnas unificado
            nueva_fila = pd.DataFrame([{
                "Fecha": datos_revision["Fecha"],
                "LCL": datos_revision["LCL"],
                "Cliente": datos_revision["Cliente"],
                "Numero_Revision": int(datos_revision["Numero_Revision"]),
                "Distrito": datos_revision["Distrito"],
                "Contratista": datos_revision["Contratista"],
                "Tipo_Atencion": datos_revision["Tipo_Atencion"],
                "Supervisor": datos_revision["Supervisor"],
                "Estado": datos_revision["Estado"],
                "ID_Revision": datos_revision["ID_Revision"],
                "Observaciones": datos_revision["Observaciones"],
                "Respuestas": datos_revision["Respuestas"]
            }])
            
            # Determinar dinámicamente si la pestaña se llama 'Registro' o 'Historial'
            try:
                xl = pd.ExcelFile(LOCAL_DB_FILE)
                sheets = xl.sheet_names
                sheet_historial = "Registro" if "Registro" in sheets else ("Historial" if "Historial" in sheets else "Registro")
            except Exception:
                sheet_historial = "Registro"
                
            df_actual = pd.read_excel(LOCAL_DB_FILE, sheet_name=sheet_historial)
            df_actual = pd.concat([df_actual, nueva_fila], ignore_index=True)
            df_proyectos = pd.read_excel(LOCAL_DB_FILE, sheet_name="Proyectos")
            
            with pd.ExcelWriter(LOCAL_DB_FILE, engine="openpyxl") as writer:
                df_proyectos.to_excel(writer, sheet_name="Proyectos", index=False)
                df_actual.to_excel(writer, sheet_name=sheet_historial, index=False)
            return True
        except Exception as e:
            st.error(f"Error al escribir localmente: {e}")
            
    return False

def normalizar_df_historial(df):
    columnas_requeridas = {
        "ID_Revision": "",
        "Fecha": "",
        "LCL": "",
        "Cliente": "",
        "Distrito": "",
        "Contratista": "",
        "Supervisor": "",
        "Tipo_Atencion": "",
        "Numero_Revision": 1,
        "Estado": "",
        "Observaciones": "",
        "Respuestas": "{}"
    }
    
    # Asegurar que todas las columnas existan
    for col, default_val in columnas_requeridas.items():
        if col not in df.columns:
            df[col] = default_val
            
    # Rellenar nulos
    for col, default_val in columnas_requeridas.items():
        df[col] = df[col].fillna(default_val)
        
    # Saneamiento de ID_Revision vacío para filas antiguas
    if not df.empty:
        df["ID_Revision"] = df["ID_Revision"].astype(str).str.strip()
        mask_vacio = (df["ID_Revision"] == "") | (df["ID_Revision"] == "None")
        if mask_vacio.any():
            # Generar un ID pseudo-único para filas antiguas
            df.loc[mask_vacio, "ID_Revision"] = df[mask_vacio].apply(
                lambda r: f"REV-ANT-{str(r['LCL']).strip()}-{str(r.name)}", axis=1
            )
            
    return df

def obtener_historial():
    """
    Obtiene todas las revisiones del historial.
    """
    modo = verificar_modo_conexion()
    df_resultado = None
    
    if modo == "gsheets":
        client = obtener_cliente_gspread()
        if client:
            try:
                sh = client.open_by_key(HISTORY_SPREADSHEET_ID)
                worksheet = sh.worksheet("Registro")
                
                # Leer usando get_all_values para evitar errores de cabeceras no únicas
                data = worksheet.get_all_values()
                if data:
                    headers = [str(h).strip() for h in data[0]]
                    rows = data[1:]
                    
                    # Sanear cabeceras duplicadas o vacías para evitar colisiones en pandas/gspread
                    cleaned_headers = []
                    for idx, h in enumerate(headers):
                        if not h:
                            cleaned_headers.append(f"Col_Vacia_{idx}")
                        elif h in cleaned_headers:
                            cleaned_headers.append(f"{h}_{idx}")
                        else:
                            cleaned_headers.append(h)
                            
                    df = pd.DataFrame(rows, columns=cleaned_headers)
                    
                    # Normalizar nombres de columnas internos
                    df.rename(columns={
                        "Nº Revisión": "Numero_Revision",
                        "Tipo de atención": "Tipo_Atencion"
                    }, inplace=True)
                    df_resultado = df
            except Exception as e:
                st.warning(f"Error al leer historial desde Google Sheets: {e}. Usando local.")
                modo = "local"
                
    if modo == "local" or df_resultado is None:
        inicializar_db_local()
        try:
            xl = pd.ExcelFile(LOCAL_DB_FILE)
            sheets = xl.sheet_names
            sheet_historial = "Registro" if "Registro" in sheets else ("Historial" if "Historial" in sheets else "Registro")
            
            df = pd.read_excel(LOCAL_DB_FILE, sheet_name=sheet_historial)
            # Normalizar columnas por compatibilidad
            df.rename(columns={
                "Nº Revisión": "Numero_Revision",
                "Tipo de atención": "Tipo_Atencion"
            }, inplace=True)
            df_resultado = df
        except Exception as e:
            st.error(f"Error al leer historial local: {e}")
            df_resultado = pd.DataFrame()
            
    return normalizar_df_historial(df_resultado)
