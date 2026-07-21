# -*- coding: utf-8 -*-
"""
Módulo Principal de la Aplicación (app.py)
Orquesta la interfaz de usuario en Streamlit, maneja el flujo de trabajo
de revisión, los estados de sesión y el enrutamiento de páginas.
"""

import json
import datetime
import streamlit as st
import pandas as pd

# Importar módulos locales
import checklists
import database
import pdf_generator

# Configuración de la página web
st.set_page_config(
    page_title="Formato ICO",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para lograr un diseño empresarial moderno
st.markdown("""
<style>
    /* Reducir el padding superior predeterminado de Streamlit */
    .block-container {
        padding-top: 2.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    
    /* Eliminar márgenes del título principal */
    h1 {
        margin-top: 0px !important;
        padding-top: 0px !important;
        margin-bottom: 10px !important;
    }

    /* Estilo para las tarjetas de indicadores */
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    .metric-val {
        font-size: 32px;
        font-weight: 700;
        color: #3c5b9e;
        line-height: 1.2;
        margin-bottom: 4px;
    }
    .metric-label {
        font-size: 13px;
        font-weight: 500;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Personalización de los Radio Buttons horizontales */
    div[role="radiogroup"] {
        gap: 15px;
    }
    
    /* Separador sutil para las filas del checklist (muy compacto) */
    .checklist-row {
        border-bottom: 1px solid #F3F4F6;
        margin-top: 1px !important;
        margin-bottom: 1px !important;
        padding: 0px !important;
    }
    
    /* Reducir brecha de radio buttons y tamaño */
    div[role="radiogroup"] {
        gap: 6px !important;
    }
    div[role="radiogroup"] label {
        font-size: 11px !important;
        padding: 0px 3px !important;
        margin: 0px !important;
    }

    /* Comprimir el bloque horizontal (st.columns) */
    div[data-testid="stHorizontalBlock"] {
        margin-bottom: -16px !important; /* Estruja las filas hacia arriba */
        padding: 0px !important;
    }

    /* Comprimir el radio widget en general */
    div.stRadio {
        margin: 0px !important;
        padding: 0px !important;
    }

    /* Reducir márgenes de textos en markdown */
    div[data-testid="stMarkdownContainer"] p {
        margin: 0px !important;
        padding: 0px !important;
    }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# MENÚ LATERAL Y NAVEGACIÓN
# ------------------------------------------------------------------
with st.sidebar:
    import os
    logo_path = "MARCA/Logo a colores.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("<h2 style='color:#3c5b9e; text-align:center; margin-top:10px;'>PLUZ</h2>", unsafe_allow_html=True)
    st.markdown("<h3 style='color:#3c5b9e; margin-top:5px; text-align:center;'>ICOS Contratista</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; font-size:12px; color:#6B7280;'>Auditoría Técnica y Control de Calidad</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    # Navegación entre páginas
    page = st.radio(
        "Navegación",
        ["📝 Nuevo Formato ICO", "📊 Historial de Formatos ICO"],
        key="navigation_page"
    )
    
    st.markdown("---")
    # Indicador de estado de conexión
    conexion_modo = database.verificar_modo_conexion()
    if conexion_modo == "gsheets":
        st.success("🟢 Conectado a Google Sheets")
    else:
        st.info("🟡 Corriendo en Modo Local\n(db_local.xlsx)")
        
    st.markdown("---")
    st.caption("v1.0.0 | Desarrollado en Python & Streamlit")

# ------------------------------------------------------------------
# PÁGINA 1: NUEVA REVISIÓN
# ------------------------------------------------------------------
def resetear_checklist_callback():
    # Callback ejecutado por Streamlit únicamente cuando el usuario cambia el tipo de atención
    nuevo_tipo = st.session_state.tipo_atencion_widget
    st.session_state.tipo_atencion_actual = nuevo_tipo
    # Inicializar las respuestas del nuevo tipo de atención a sus valores por defecto
    checklist_items = checklists.obtener_checklist_completo(nuevo_tipo)
    for item in checklist_items:
        st.session_state[f"widget_resp_{item['id']}"] = item["respuesta_defecto"]

if page == "📝 Nuevo Formato ICO":
    st.markdown("<h1 style='color:#3c5b9e;'>FORMATO ICO</h1>", unsafe_allow_html=True)
    st.markdown("Complete la información general e inspeccione los requisitos del checklist.")
    
    # Inicializar estados de la sesión para persistir búsquedas e inicializaciones
    if "proyecto_encontrado" not in st.session_state:
        st.session_state.proyecto_encontrado = None
    if "lcl_buscado" not in st.session_state:
        st.session_state.lcl_buscado = None
    if "tipo_atencion_actual" not in st.session_state:
        st.session_state.tipo_atencion_actual = "Reforma"

    # --------------------------------------------------------------
    # COLUMNADO PRINCIPAL: IZQUIERDA (Búsqueda, Info y Dictamen) / DERECHA (Checklist)
    # --------------------------------------------------------------
    col_izq, col_der = st.columns([5, 7])
    
    # --- COLUMNA IZQUIERDA: BÚSQUEDA Y METADATOS ---
    with col_izq:
        st.subheader("1. Buscar Proyecto")
        col_lcl, col_btn = st.columns([3, 1])
        
        with col_lcl:
            lcl_input = st.text_input(
                "Ingrese el código LCL del proyecto:",
                value=st.session_state.lcl_buscado if st.session_state.lcl_buscado else "",
                placeholder="Ej. LCL-001, LCL-002",
                label_visibility="collapsed",
                key="lcl_search_input_widget"
            ).strip()
            
        with col_btn:
            buscar_clicked = st.button("🔍 Buscar", use_container_width=True, type="primary", key="btn_lcl_search_submit")
            
        # Acción de búsqueda
        if buscar_clicked:
            if lcl_input:
                with st.spinner("Buscando proyecto..."):
                    proyecto = database.buscar_proyecto_por_lcl(lcl_input)
                    if proyecto:
                        st.session_state.proyecto_encontrado = proyecto
                        st.session_state.lcl_buscado = lcl_input
                        # Limpiar observaciones de la búsqueda previa
                        if "txt_observaciones" in st.session_state:
                            del st.session_state["txt_observaciones"]
                        # Resetear respuestas del checklist para el nuevo proyecto
                        checklist_items = checklists.obtener_checklist_completo(st.session_state.get("tipo_atencion_actual", "Reforma"))
                        for item in checklist_items:
                            k_name = f"widget_resp_{item['id']}"
                            st.session_state[k_name] = item["respuesta_defecto"]
                        st.success(f"¡Proyecto {lcl_input} cargado!")
                        st.rerun()
                    else:
                        st.session_state.proyecto_encontrado = None
                        st.session_state.lcl_buscado = None
                        st.error(f"Error: LCL '{lcl_input}' no existe en Google Sheets ni en la BD local.")
            else:
                st.warning("Por favor, ingrese un código LCL.")
                
        # Si hay un proyecto cargado, mostramos el resto de la columna izquierda
        if st.session_state.proyecto_encontrado:
            proyecto = st.session_state.proyecto_encontrado
            
            st.markdown("---")
            st.subheader("2. Información General")
            
            st.text_input("Cliente", value=proyecto.get("Cliente", ""), disabled=True)
            st.text_input("Distrito", value=proyecto.get("Distrito", ""), disabled=True)
            st.text_input("Contratista", value=proyecto.get("Contratista", ""), disabled=True)
            st.text_input("Supervisor Revisor", value=proyecto.get("Supervisor", ""), disabled=True)
            
            fecha_hoy = datetime.date.today().strftime("%Y-%m-%d")
            st.text_input("Fecha de Auditoría", value=fecha_hoy, disabled=True)
            
            num_revision = st.number_input(
                "Número de Revisión",
                min_value=1,
                value=1,
                step=1,
                key="num_revision_val"
            )
            
            # Recuperar respuestas para calcular el estado en tiempo real
            tipo_atencion = st.session_state.tipo_atencion_actual
            checklist_items = checklists.obtener_checklist_completo(tipo_atencion)
            
            # Asegurar la inicialización inicial del checklist
            for item in checklist_items:
                k_name = f"widget_resp_{item['id']}"
                if k_name not in st.session_state:
                    st.session_state[k_name] = item["respuesta_defecto"]
            
            respuestas_actuales = {}
            tiene_no = False
            for item in checklist_items:
                q_id = item["id"]
                val = st.session_state.get(f"widget_resp_{q_id}", item["respuesta_defecto"])
                respuestas_actuales[q_id] = val
                if val == "No":
                    tiene_no = True
            
            estado_calculado = "OBSERVADO" if tiene_no else "CONFORME"
            
            st.markdown("---")
            st.subheader("4. Evaluación y Dictamen")
            
            if estado_calculado == "CONFORME":
                st.markdown(
                    "<div style='background-color:#E8F5E9; border: 1.5px solid #6dab5e; border-radius:8px; padding:10px; text-align:center; color:#2E7D32; font-weight:bold; font-size:16px;'>"
                    "✅ CONFORME (Requisitos aprobados o N.A.)"
                    "</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    "<div style='background-color:#FEE2E2; border: 1.5px solid #EF4444; border-radius:8px; padding:10px; text-align:center; color:#991B1B; font-weight:bold; font-size:16px;'>"
                    "❌ OBSERVADO (Existe incumplimiento)"
                    "</div>",
                    unsafe_allow_html=True
                )
                
            st.write("")
            observaciones = st.text_area(
                "Observaciones Generales:",
                placeholder="Escriba aquí los detalles...",
                height=100,
                key="txt_observaciones"
            )
            
            st.markdown("---")
            st.subheader("5. Finalizar Auditoría")
            btn_generar = st.button("📋 Guardar y Generar PDF", type="primary", use_container_width=True, key="btn_guardar_pdf")
            
            if btn_generar:
                time_stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
                id_revision = f"REV-{proyecto.get('LCL', 'NOLCL')}-{time_stamp}"
                fecha_registro = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                datos_revision = {
                    "ID_Revision": id_revision,
                    "Fecha": fecha_registro,
                    "LCL": proyecto.get("LCL", ""),
                    "Cliente": proyecto.get("Cliente", ""),
                    "Distrito": proyecto.get("Distrito", ""),
                    "Contratista": proyecto.get("Contratista", ""),
                    "Supervisor": proyecto.get("Supervisor", ""),
                    "Tipo_Atencion": tipo_atencion,
                    "Numero_Revision": int(num_revision),
                    "Estado": estado_calculado,
                    "Observaciones": observaciones,
                    "Respuestas": json.dumps(respuestas_actuales)
                }
                
                checklist_para_pdf = []
                for item in checklist_items:
                    q_id = item["id"]
                    checklist_para_pdf.append({
                        "id": q_id,
                        "descripcion": item["descripcion"],
                        "respuesta": respuestas_actuales[q_id]
                    })
                
                with st.spinner("Guardando y compilando informe PDF..."):
                    guardado_ok = database.registrar_revision(datos_revision)
                    pdf_buffer = pdf_generator.generar_pdf(datos_revision, checklist_para_pdf)
                    
                    if guardado_ok:
                        st.success("🎉 Auditoría guardada con éxito.")
                        st.download_button(
                            label="📥 Descargar Reporte Técnico PDF",
                            data=pdf_buffer,
                            file_name=f"Revision_{proyecto.get('LCL', '')}_{time_stamp}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key="btn_descarga_ok"
                        )
                    else:
                        st.error("Error al guardar en el historial.")

    # --- COLUMNA DERECHA: TIPO DE ATENCIÓN Y CHECKLIST COMPACTO ---
    with col_der:
        if st.session_state.proyecto_encontrado:
            tipo_atencion = st.session_state.tipo_atencion_actual
            checklist_items = checklists.obtener_checklist_completo(tipo_atencion)
            
            options_atencion = [
                "Reforma",
                "Reforma Sustancial",
                "Informe de Factibilidad",
                "Informe de Pre Factibilidad",
                "Obra Civil",
                "Informe de Área"
            ]
            default_sel_idx = options_atencion.index(tipo_atencion)

            st.subheader("3. Tipo de Atención y Checklist")
            
            # Selector de Tipo de Atención (sincronizado con key y callback)
            st.selectbox(
                "Seleccione el Tipo de Atención:",
                options_atencion,
                index=default_sel_idx,
                key="tipo_atencion_widget",
                on_change=resetear_checklist_callback,
                label_visibility="collapsed"
            )
            
            st.write("") # Pequeño espacio de separación
            
            # Mostrar los 39 requisitos en tamaño compacto
            for item in checklist_items:
                q_id = item["id"]
                descripcion = item["descripcion"]
                
                r_col1, r_col2, r_col3 = st.columns([1, 7, 4])
                
                with r_col1:
                    st.markdown(f"<span style='font-size:12px; font-weight:bold;'>{q_id}</span>", unsafe_allow_html=True)
                with r_col2:
                    st.markdown(f"<span style='font-size:11px; line-height:1.2; display:block;'>{descripcion}</span>", unsafe_allow_html=True)
                with r_col3:
                    # Inicializar en la sesión si no existe
                    if f"widget_resp_{q_id}" not in st.session_state:
                        st.session_state[f"widget_resp_{q_id}"] = item["respuesta_defecto"]
                    
                    st.radio(
                        f"Respuesta_{q_id}",
                        options=["Sí", "No", "N.A."],
                        key=f"widget_resp_{q_id}",
                        horizontal=True,
                        label_visibility="collapsed"
                    )
                st.markdown("<div class='checklist-row'></div>", unsafe_allow_html=True)
        else:
            st.info("💡 Ingrese un código LCL a la izquierda y presione Buscar para cargar el Formato ICO.")
            
# ------------------------------------------------------------------
# PÁGINA 2: HISTORIAL DE REVISIONES
# ------------------------------------------------------------------
elif page == "📊 Historial de Formatos ICO":
    st.markdown("# HISTORIAL DE REVISIONES TÉCNICAS")
    st.markdown("Consulte el histórico de auditorías, aplique filtros y re-genere reportes en PDF de revisiones pasadas.")
    
    # Cargar datos desde la base de datos
    with st.spinner("Cargando base de datos de auditorías..."):
        df_historial = database.obtener_historial()
        
    if df_historial.empty:
        st.info("No se han registrado auditorías todavía. Realiza una nueva revisión en la pantalla principal.")
    else:
        # ----------------------------------------------------------
        # INDICADORES (TARJETAS KPI)
        # ----------------------------------------------------------
        total_rev = len(df_historial)
        conformes = len(df_historial[df_historial["Estado"] == "CONFORME"])
        observados = len(df_historial[df_historial["Estado"] == "OBSERVADO"])
        
        st.write("")
        kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
        
        with kpi_col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-val">{total_rev}</div>
                <div class="metric-label">Total Revisiones</div>
            </div>
            """, unsafe_allow_html=True)
            
        with kpi_col2:
            st.markdown(f"""
            <div class="metric-card" style="border-left: 5px solid #10B981;">
                <div class="metric-val" style="color: #10B981;">{conformes}</div>
                <div class="metric-label">Conformes</div>
            </div>
            """, unsafe_allow_html=True)
            
        with kpi_col3:
            st.markdown(f"""
            <div class="metric-card" style="border-left: 5px solid #EF4444;">
                <div class="metric-val" style="color: #EF4444;">{observados}</div>
                <div class="metric-label">Observados</div>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown("---")
        
        # ----------------------------------------------------------
        # SECCIÓN DE FILTROS
        # ----------------------------------------------------------
        with st.expander("🔍 Filtros de Búsqueda Avanzada", expanded=True):
            f_col1, f_col2, f_col3 = st.columns(3)
            
            with f_col1:
                filtro_lcl = st.text_input("Filtrar por LCL (Proyecto):", placeholder="Ej. LCL-001")
                filtro_estado = st.selectbox("Filtrar por Estado:", ["Todos", "CONFORME", "OBSERVADO"])
                
            with f_col2:
                filtro_cliente = st.text_input("Filtrar por Cliente:")
                filtro_fecha = st.text_input("Filtrar por Fecha (AÑO-MES-DÍA):", placeholder="Ej. 2026-07")
                
            with f_col3:
                filtro_contratista = st.text_input("Filtrar por Contratista:")
        
        # Aplicar filtros al DataFrame
        df_filtrado = df_historial.copy()
        
        if filtro_lcl:
            df_filtrado = df_filtrado[df_filtrado["LCL"].astype(str).str.contains(filtro_lcl, case=False, na=False)]
            
        if filtro_cliente:
            df_filtrado = df_filtrado[df_filtrado["Cliente"].astype(str).str.contains(filtro_cliente, case=False, na=False)]
            
        if filtro_contratista:
            df_filtrado = df_filtrado[df_filtrado["Contratista"].astype(str).str.contains(filtro_contratista, case=False, na=False)]
            
        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Estado"] == filtro_estado]
            
        if filtro_fecha:
            df_filtrado = df_filtrado[df_filtrado["Fecha"].astype(str).str.contains(filtro_fecha, na=False)]
            
        st.markdown(f"**Resultados encontrados:** {len(df_filtrado)} registros")
        
        # ----------------------------------------------------------
        # TABLA DE REGISTROS
        # ----------------------------------------------------------
        if not df_filtrado.empty:
            # Seleccionar y reordenar columnas para visualización clara
            df_mostrar = df_filtrado[[
                "ID_Revision", "Fecha", "LCL", "Cliente", "Distrito", 
                "Contratista", "Supervisor", "Tipo_Atencion", 
                "Numero_Revision", "Estado"
            ]].sort_values(by="Fecha", ascending=False)
            
            st.dataframe(
                df_mostrar,
                use_container_width=True,
                hide_index=True
            )
            
            # ----------------------------------------------------------
            # RE-GENERACIÓN Y DESCARGA DE PDF DESDE EL HISTORIAL
            # ----------------------------------------------------------
            st.markdown("### Re-generación de Reporte PDF del Historial")
            st.markdown("Seleccione una revisión de la lista para volver a compilar y descargar su reporte técnico en PDF.")
            
            col_sel, col_dl = st.columns([3, 1])
            
            with col_sel:
                seleccion_id = st.selectbox(
                    "Seleccione la revisión (por ID):",
                    options=df_mostrar["ID_Revision"].tolist(),
                    help="Lista de IDs del historial actualmente filtrado"
                )
                
            with col_dl:
                st.write("") # Alineación
                st.write("")
                
                if seleccion_id:
                    # Extraer fila correspondiente
                    fila_sel = df_filtrado[df_filtrado["ID_Revision"] == seleccion_id].iloc[0]
                    
                    try:
                        # Reconstruir las respuestas desde la cadena JSON
                        respuestas_json = fila_sel["Respuestas"]
                        if isinstance(respuestas_json, str):
                            respuestas_dict = json.loads(respuestas_json)
                        else:
                            # Por si pandas lo lee directo como dict
                            respuestas_dict = respuestas_json
                            
                        # Obtener las preguntas del tipo de atención asociado
                        tipo_atencion_sel = fila_sel["Tipo_Atencion"]
                        preguntas_tipo = checklists.obtener_checklist_completo(tipo_atencion_sel)
                        
                        # Armar estructura de checklist completa para el PDF
                        checklist_reconst = []
                        for item in preguntas_tipo:
                            q_id = item["id"]
                            # Obtener respuesta guardada (compatibilidad con llave string/int)
                            ans_guardada = respuestas_dict.get(str(q_id), respuestas_dict.get(q_id, "N.A."))
                            checklist_reconst.append({
                                "id": q_id,
                                "descripcion": item["descripcion"],
                                "respuesta": ans_guardada
                            })
                            
                        # Armar diccionario de datos de revisión
                        revision_reconst = {
                            "LCL": fila_sel["LCL"],
                            "Fecha": fila_sel["Fecha"],
                            "Cliente": fila_sel["Cliente"],
                            "Distrito": fila_sel["Distrito"],
                            "Contratista": fila_sel["Contratista"],
                            "Supervisor": fila_sel["Supervisor"],
                            "Tipo_Atencion": fila_sel["Tipo_Atencion"],
                            "Numero_Revision": int(fila_sel["Numero_Revision"]),
                            "Estado": fila_sel["Estado"],
                            "Observaciones": fila_sel.get("Observaciones", "")
                        }
                        
                        # Generar el PDF
                        pdf_reconst_buf = pdf_generator.generar_pdf(revision_reconst, checklist_reconst)
                        
                        # Botón de descarga
                        st.download_button(
                            label="📥 Volver a descargar PDF",
                            data=pdf_reconst_buf,
                            file_name=f"Reporte_Recompilado_{fila_sel['LCL']}_{seleccion_id}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"btn_dl_reconst_{seleccion_id}"
                        )
                    except Exception as e:
                        st.error(f"Error al reconstruir el reporte técnico: {e}")
        else:
            st.warning("No hay registros que coincidan con los filtros aplicados.")
