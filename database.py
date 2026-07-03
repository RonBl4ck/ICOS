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
                        "ID_Revision", "Fecha", "LCL", "Cliente", "Nº Revisión", 
                        "Distrito", "Contratista", "Tipo de atención", "Supervisor", "Estado", "Observaciones", "Respuestas"
                    ])
                
                # Escribir la fila
                worksheet.append_row([
                    datos_revision["ID_Revision"],
                    datos_revision["Fecha"],
                    datos_revision["LCL"],
                    datos_revision["Cliente"],
                    int(datos_revision["Numero_Revision"]),
                    datos_revision["Distrito"],
                    datos_revision["Contratista"],
                    datos_revision["Tipo_Atencion"],
                    datos_revision["Supervisor"],
                    datos_revision["Estado"],
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
            # Estructurar fila local
            nueva_fila = pd.DataFrame([{
                "ID_Revision": datos_revision["ID_Revision"],
                "Fecha": datos_revision["Fecha"],
                "LCL": datos_revision["LCL"],
                "Cliente": datos_revision["Cliente"],
                "Distrito": datos_revision["Distrito"],
                "Contratista": datos_revision["Contratista"],
                "Supervisor": datos_revision["Supervisor"],
                "Tipo_Atencion": datos_revision["Tipo_Atencion"],
                "Numero_Revision": int(datos_revision["Numero_Revision"]),
                "Estado": datos_revision["Estado"],
                "Observaciones": datos_revision["Observaciones"],
                "Respuestas": datos_revision["Respuestas"]
            }])
            
            df_actual = pd.read_excel(LOCAL_DB_FILE, sheet_name="Registro")
            df_actual = pd.concat([df_actual, nueva_fila], ignore_index=True)
            df_proyectos = pd.read_excel(LOCAL_DB_FILE, sheet_name="Proyectos")
            
            with pd.ExcelWriter(LOCAL_DB_FILE, engine="openpyxl") as writer:
                df_proyectos.to_excel(writer, sheet_name="Proyectos", index=False)
                df_actual.to_excel(writer, sheet_name="Registro", index=False)
            return True
        except Exception as e:
            st.error(f"Error al escribir localmente: {e}")
            
    return False

def obtener_historial():
    """
    Obtiene todas las revisiones del historial.
    """
    modo = verificar_modo_conexion()
    
    if modo == "gsheets":
        client = obtener_cliente_gspread()
        if client:
            try:
                sh = client.open_by_key(HISTORY_SPREADSHEET_ID)
                worksheet = sh.worksheet("Registro")
                records = worksheet.get_all_records()
                if records:
                    df = pd.DataFrame(records)
                    # Normalizar nombres de columnas a español de tu script si es necesario
                    # Pero usaremos los del dataframe directamente.
                    # Mapear nombres si tienen discrepancia:
                    df.rename(columns={
                        "Nº Revisión": "Numero_Revision",
                        "Tipo de atención": "Tipo_Atencion"
                    }, inplace=True)
                    return df
            except Exception as e:
                st.warning(f"Error al leer historial desde Google Sheets: {e}. Usando local.")
                modo = "local"
                
    if modo == "local":
        inicializar_db_local()
        try:
            df = pd.read_excel(LOCAL_DB_FILE, sheet_name="Registro")
            return df
        except Exception as e:
            st.error(f"Error al leer historial local: {e}")
            
    return pd.DataFrame(columns=[
        "ID_Revision", "Fecha", "LCL", "Cliente", "Distrito", 
        "Contratista", "Supervisor", "Tipo_Atencion", 
        "Numero_Revision", "Estado", "Observaciones", "Respuestas"
    ])
