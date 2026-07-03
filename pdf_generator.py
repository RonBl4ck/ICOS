# -*- coding: utf-8 -*-
"""
Módulo de Generación de PDF
Genera un informe PDF profesional y limpio con ReportLab a partir de los
datos de la revisión y las respuestas del checklist.
"""

import os
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image as RLImage
)
from reportlab.pdfgen import canvas

class NumberedCanvas(canvas.Canvas):
    """
    Canvas personalizado para calcular el total de páginas dinámicamente
    y dibujar el encabezado y pie de página en cada hoja.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        
        # Margen izquierdo = 36 pt, Derecho = 576 pt (para ancho de 612 pt)
        # Margen inferior = 54 pt, Superior = 738 pt (para alto de 792 pt)
        
        # Dibujar encabezado en páginas mayores a 1
        if self._pageNumber > 1:
            self.setStrokeColor(colors.HexColor("#E5E7EB"))
            self.setLineWidth(0.5)
            self.line(36, 750, 576, 750)
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#4B5563"))
            self.drawString(36, 755, "FORMATO ICO - REVISIÓN TÉCNICA")
            
        # Dibujar barra tricolor en la parte inferior de todas las hojas
        # Ancho total = 540 pt (de 36 a 576). 180 pt por color: Azul (#3c5b9e), Amarillo (#fbc241), Verde (#6dab5e)
        self.setFillColor(colors.HexColor("#3c5b9e"))
        self.rect(36, 45, 180, 4, fill=True, stroke=False)
        self.setFillColor(colors.HexColor("#fbc241"))
        self.rect(216, 45, 180, 4, fill=True, stroke=False)
        self.setFillColor(colors.HexColor("#6dab5e"))
        self.rect(396, 45, 180, 4, fill=True, stroke=False)
        
        # Dibujar pie de página en todas las hojas por debajo de la barra
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#4B5563"))
        page_text = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(576, 30, page_text)
        self.drawString(36, 30, "Confidencial - Uso Interno - Pluz Energía")
        
        self.restoreState()

def generar_pdf(revision, checklist):
    """
    Genera el archivo PDF a partir de los datos de la revisión.
    
    Parámetros:
    - revision: Diccionario con llaves (LCL, Cliente, Distrito, Contratista,
                Supervisor, Fecha, Tipo_Atencion, Numero_Revision, Estado, Observaciones)
    - checklist: Lista de diccionarios con las respuestas
                 [{'id': 1, 'descripcion': '...', 'respuesta': 'Sí'/'No'/'N.A.'}]
                 
    Retorna:
    - BytesIO: Buffer con el contenido del PDF para descarga en memoria.
    """
    buffer = BytesIO()
    
    # Configurar el documento: márgenes de 0.5 pulgadas (36 puntos)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=54,
        bottomMargin=64
    )
    
    story = []
    
    # ----------------------------------------------------
    # Estilos
    # ----------------------------------------------------
    styles = getSampleStyleSheet()
    
    # Color principal: Azul corporativo Pluz (#3c5b9e)
    # Color secundario: Gris oscuro (#374151)
    
    style_titulo = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#3c5b9e"),
        alignment=TA_CENTER,
        spaceAfter=15
    )
    
    style_titulo_derecha = ParagraphStyle(
        'DocTitleRight',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#3c5b9e"),
        alignment=TA_RIGHT
    )
    
    style_subseccion = ParagraphStyle(
        'SubSection',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#1F2937"),
        spaceBefore=10,
        spaceAfter=6,
        keepWithNext=True
    )
    
    style_body_bold = ParagraphStyle(
        'BodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#374151")
    )
    
    style_body_val = ParagraphStyle(
        'BodyVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#4B5563")
    )
    
    style_checklist_num = ParagraphStyle(
        'ChecklistNum',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=TA_CENTER
    )
    
    style_checklist_desc = ParagraphStyle(
        'ChecklistDesc',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#1F2937")
    )
    
    style_resp_si = ParagraphStyle('RespSi', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor("#6dab5e"), alignment=TA_CENTER)
    style_resp_no = ParagraphStyle('RespNo', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor("#B91C1C"), alignment=TA_CENTER)
    style_resp_na = ParagraphStyle('RespNA', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#6B7280"), alignment=TA_CENTER)

    # ----------------------------------------------------
    # Membrete y Título con Logo de Pluz
    # ----------------------------------------------------
    logo_path = os.path.join(os.path.dirname(__file__), "MARCA", "Logo a colores.jpg")
    
    if os.path.exists(logo_path):
        # Escalar manteniendo proporción 16:9 (110 de ancho, 62 de alto)
        logo_img = RLImage(logo_path, width=110, height=62)
    else:
        logo_img = Paragraph("<b>PLUZ</b>", ParagraphStyle('PluzTxt', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor("#3c5b9e")))
        
    title_p = Paragraph("FORMATO ICO", style_titulo_derecha)
    
    # Tabla de cabecera: Logo (150 pt) | Título (390 pt)
    tabla_cabecera = Table([[logo_img, title_p]], colWidths=[150, 390])
    tabla_cabecera.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    
    story.append(tabla_cabecera)
    story.append(Spacer(1, 10))
    
    # ----------------------------------------------------
    # Tabla de Datos Generales (4 columnas)
    # Ancho total disponible = 540 pt
    # Col 1: Label (100 pt) | Col 2: Val (170 pt) | Col 3: Label (100 pt) | Col 4: Val (170 pt)
    # ----------------------------------------------------
    datos_tabla = [
        [
            Paragraph("LCL:", style_body_bold), Paragraph(str(revision.get('LCL', '')), style_body_val),
            Paragraph("Fecha de Revisión:", style_body_bold), Paragraph(str(revision.get('Fecha', '')), style_body_val)
        ],
        [
            Paragraph("Cliente:", style_body_bold), Paragraph(str(revision.get('Cliente', '')), style_body_val),
            Paragraph("Nº de Revisión:", style_body_bold), Paragraph(str(revision.get('Numero_Revision', '1')), style_body_val)
        ],
        [
            Paragraph("Distrito:", style_body_bold), Paragraph(str(revision.get('Distrito', '')), style_body_val),
            Paragraph("Tipo de Atención:", style_body_bold), Paragraph(str(revision.get('Tipo_Atencion', '')), style_body_val)
        ],
        [
            Paragraph("Contratista:", style_body_bold), Paragraph(str(revision.get('Contratista', '')), style_body_val),
            Paragraph("Supervisor:", style_body_bold), Paragraph(str(revision.get('Supervisor', '')), style_body_val)
        ]
    ]
    
    t_generales = Table(datos_tabla, colWidths=[100, 170, 100, 170])
    t_generales.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F9FAFB")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F9FAFB")),
    ]))
    
    story.append(t_generales)
    story.append(Spacer(1, 15))
    
    # ----------------------------------------------------
    # Estado de la Revisión (Destacado en un banner)
    # ----------------------------------------------------
    estado = str(revision.get('Estado', 'OBSERVADO')).upper()
    if estado == "CONFORME":
        bg_color = colors.HexColor("#D1FAE5")
        text_color = "#065F46"
        border_color = colors.HexColor("#10B981")
        mensaje = "CUMPLE CON LOS REQUISITOS ESTABLECIDOS - ESTADO: CONFORME"
    else:
        bg_color = colors.HexColor("#FEE2E2")
        text_color = "#991B1B"
        border_color = colors.HexColor("#EF4444")
        mensaje = "CONTIENE OBSERVACIONES TÉCNICAS - ESTADO: OBSERVADO"
        
    style_estado_txt = ParagraphStyle(
        'EstadoTxt',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor(text_color),
        alignment=TA_CENTER
    )
    
    t_estado = Table([[Paragraph(mensaje, style_estado_txt)]], colWidths=[540])
    t_estado.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg_color),
        ('BOX', (0,0), (-1,-1), 1, border_color),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    
    story.append(t_estado)
    story.append(Spacer(1, 15))
    
    # ----------------------------------------------------
    # Checklist de Preguntas
    # Ancho total disponible = 540 pt
    # Col 1: Nº (30 pt) | Col 2: Descripción (420 pt) | Col 3: Respuesta (90 pt)
    # ----------------------------------------------------
    story.append(Paragraph("DETALLE DEL CHECKLIST DE INSPECCIÓN", style_subseccion))
    
    encabezado_checklist = [
        Paragraph("<b>Nº</b>", ParagraphStyle('H1', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=TA_CENTER)),
        Paragraph("<b>Requisito a Inspeccionar</b>", ParagraphStyle('H2', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white)),
        Paragraph("<b>Resultado</b>", ParagraphStyle('H3', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=TA_CENTER))
    ]
    
    filas_checklist = [encabezado_checklist]
    
    for idx, item in enumerate(checklist, 1):
        num_p = Paragraph(str(item.get('id', idx)), style_checklist_num)
        desc_p = Paragraph(item.get('descripcion', ''), style_checklist_desc)
        
        resp_val = item.get('respuesta', 'N.A.')
        if resp_val == "Sí":
            resp_p = Paragraph("SÍ", style_resp_si)
        elif resp_val == "No":
            resp_p = Paragraph("NO", style_resp_no)
        else:
            resp_p = Paragraph("N.A.", style_resp_na)
            
        filas_checklist.append([num_p, desc_p, resp_p])
        
    t_checklist = Table(filas_checklist, colWidths=[30, 420, 90], repeatRows=1)
    
    # Estilo de la tabla
    estilos_tabla = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#3c5b9e")),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]
    
    # Filas alternas (Cebra)
    for i in range(1, len(filas_checklist)):
        if i % 2 == 0:
            estilos_tabla.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F9FAFB")))
            
    t_checklist.setStyle(TableStyle(estilos_tabla))
    story.append(t_checklist)
    story.append(Spacer(1, 15))
    
    # ----------------------------------------------------
    # Observaciones
    # ----------------------------------------------------
    observaciones_txt = revision.get('Observaciones', '')
    if not observaciones_txt or observaciones_txt.strip() == "":
        observaciones_txt = "Sin observaciones particulares registradas."
        
    style_obs_body = ParagraphStyle(
        'ObsBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#374151")
    )
    
    bloque_obs = []
    bloque_obs.append(Paragraph("OBSERVACIONES GENERALES", style_subseccion))
    
    # Caja para observaciones
    t_obs = Table([[Paragraph(observaciones_txt, style_obs_body)]], colWidths=[540])
    t_obs.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F3F4F6")),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#D1D5DB")),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    bloque_obs.append(t_obs)
    
    story.append(KeepTogether(bloque_obs))
    story.append(Spacer(1, 10))
    
    # ----------------------------------------------------
    # Construir PDF
    # ----------------------------------------------------
    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer
