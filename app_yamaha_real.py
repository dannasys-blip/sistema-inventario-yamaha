from flask import Flask, render_template, request, redirect, url_for, make_response
import pandas as pd
import os
import json
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io

app = Flask(__name__)

CSV_FILE = 'base_datos_yamaha_final.csv'

def inicializar_csv():
    if not os.path.exists(CSV_FILE):
        columnas = [
            'id', 'tipo_equipo', 'codigo_patrimonial', 'marca', 'modelo', 
            'serie', 'procesador', 'memoria_ram', 'almacenamiento', 
            'usuario_asignado', 'area', 'sucursal', 'estado', 'perifericos', 'observaciones', 'fotos', 'firma'
        ]
        df = pd.DataFrame(columns=columnas)
        df.to_csv(CSV_FILE, index=False)

inicializar_csv()

@app.route('/')
def index():
    df = pd.read_csv(CSV_FILE)
    registros = []
    for _, row in df.iterrows():
        reg = row.to_dict()
        try:
            reg['perifericos_lista'] = json.loads(reg['perifericos']) if pd.notna(reg['perifericos']) else []
        except:
            reg['perifericos_lista'] = []
        try:
            reg['fotos_lista'] = json.loads(reg['fotos']) if pd.notna(reg['fotos']) else []
        except:
            reg['fotos_lista'] = []
        registros.append(reg)
    return render_template('index.html', registros=registros)

@app.route('/guardar', methods=['POST'])
def guardar():
    df = pd.read_csv(CSV_FILE)
    
    nuevo_id = int(df['id'].max()) + 1 if not df.empty and not pd.isna(df['id'].max()) else 1
    tipo_equipo = request.form.get('tipo_equipo', '').strip()
    almacenamiento = request.form.get('almacenamiento', '').strip()
    
    # Si es impresora, switch o TV, limpiamos campos de PC/Laptop
    if tipo_equipo.lower() in ['impresora', 'switch', 'tv']:
        almacenamiento = ''

    # Periféricos dinámicos
    tipos_p = request.form.getlist('periferico_tipo[]')
    marcas_p = request.form.getlist('periferico_marca[]')
    modelos_p = request.form.getlist('periferico_modelo[]')
    series_p = request.form.getlist('periferico_serie[]')
    estados_p = request.form.getlist('periferico_estado[]')
    obs_p = request.form.getlist('periferico_obs[]')

    lista_perifericos = []
    for i in range(len(tipos_p)):
        if tipos_p[i].strip():
            lista_perifericos.append({
                'tipo': tipos_p[i].strip(),
                'marca': marcas_p[i].strip() if i < len(marcas_p) else '',
                'modelo': modelos_p[i].strip() if i < len(modelos_p) else '',
                'serie': series_p[i].strip() if i < len(series_p) else '',
                'estado': estados_p[i].strip() if i < len(estados_p) else 'Operativo',
                'observaciones': obs_p[i].strip() if i < len(obs_p) else ''
            })

    # Fotos (URLs o enlaces de Cloudinary)
    fotos_input = request.form.get('fotos', '').strip()
    lista_fotos = [f.strip() for f in fotos_input.split('\n') if f.strip()] if fotos_input else []

    nuevo_registro = {
        'id': nuevo_id,
        'tipo_equipo': tipo_equipo,
        'codigo_patrimonial': request.form.get('codigo_patrimonial', '').strip(),
        'marca': request.form.get('marca', '').strip(),
        'modelo': request.form.get('modelo', '').strip(),
        'serie': request.form.get('serie', '').strip(),
        'procesador': request.form.get('procesador', '').strip(),
        'memoria_ram': request.form.get('memoria_ram', '').strip(),
        'almacenamiento': almacenamiento,
        'usuario_asignado': request.form.get('usuario_asignado', '').strip(),
        'area': request.form.get('area', '').strip(),
        'sucursal': request.form.get('sucursal', '').strip(),
        'estado': request.form.get('estado', 'Activo').strip(),
        'perifericos': json.dumps(lista_perifericos),
        'observaciones': request.form.get('observaciones_generales', '').strip(),
        'fotos': json.dumps(lista_fotos),
        'firma': request.form.get('firma', '').strip()
    }
    
    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    df.to_csv(CSV_FILE, index=False)
    
    return redirect(url_for('index'))

@app.route('/pdf/<int:id>')
def generar_pdf(id):
    df = pd.read_csv(CSV_FILE)
    reg = df[df['id'] == id]
    if reg.empty:
        return "Registro no encontrado", 404
    
    row = reg.iloc[0].to_dict()
    try:
        perifericos = json.loads(row['perifericos']) if pd.notna(row['perifericos']) else []
    except:
        perifericos = []

    try:
        fotos = json.loads(row['fotos']) if pd.notna(row['fotos']) else []
    except:
        fotos = []

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=13, textColor=colors.HexColor('#003366'), alignment=1, spaceAfter=8)
    section_style = ParagraphStyle('SectionStyle', parent=styles['Heading2'], fontSize=9, textColor=colors.white, backColor=colors.HexColor('#003366'), spaceAfter=4, leftIndent=4)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontSize=8, textColor=colors.black)
    bold_style = ParagraphStyle('BoldStyle', parent=normal_style, fontName='Helvetica-Bold')

    # Encabezado Oficial
    elements.append(Paragraph("<b>YAMAHA MOTOR / VASSOSP TRADING EIRL</b><br>FICHA DE INVENTARIO DE EQUIPOS TECNOLOGICOS", title_style))
    elements.append(Spacer(1, 5))

    # 1. Datos del Trabajador / Responsable[cite: 3]
    elements.append(Paragraph("1. DATOS DEL TRABAJADOR / RESPONSABLE", section_style))
    data_collab = [
        [Paragraph("<b>N° Ficha:</b>", normal_style), Paragraph(f"INV-2026-{row['id']:03d}", normal_style), Paragraph("<b>Fecha:</b>", normal_style), Paragraph("2026-11-08", normal_style)],
        [Paragraph("<b>Colaborador:</b>", normal_style), Paragraph(row['usuario_asignado'], normal_style), Paragraph("<b>Área / Dpto:</b>", normal_style), Paragraph(row['area'], normal_style)],
        [Paragraph("<b>Sucursal:</b>", normal_style), Paragraph(row['sucursal'], normal_style), Paragraph("<b>Estado:</b>", normal_style), Paragraph(row['estado'], normal_style)]
    ]
    t_collab = Table(data_collab, colWidths=[70, 200, 70, 210])
    t_collab.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 3)]))
    elements.append(t_collab)
    elements.append(Spacer(1, 6))

    # 2. Detalle de Equipo Principal[cite: 3]
    elements.append(Paragraph(f"2. DETALLE DE EQUIPO PRINCIPAL ({row['tipo_equipo']})", section_style))
    data_eq = [
        [Paragraph("<b>Tipo de Equipo:</b>", normal_style), Paragraph(row['tipo_equipo'], normal_style), Paragraph("<b>Patrimonial:</b>", normal_style), Paragraph(row['codigo_patrimonial'], normal_style)],
        [Paragraph("<b>Marca / Modelo:</b>", normal_style), Paragraph(f"{row['marca']} {row['modelo']}", normal_style), Paragraph("<b>N° Serie:</b>", normal_style), Paragraph(row['serie'], normal_style)],
        [Paragraph("<b>Procesador:</b>", normal_style), Paragraph(row['procesador'], normal_style), Paragraph("<b>Memoria RAM:</b>", normal_style), Paragraph(row['memoria_ram'], normal_style)],
        [Paragraph("<b>Almacenamiento:</b>", normal_style), Paragraph(row['almacenamiento'], normal_style), Paragraph("<b>Sucursal:</b>", normal_style), Paragraph(row['sucursal'], normal_style)]
    ]
    t_eq = Table(data_eq, colWidths=[90, 180, 90, 190])
    t_eq.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 3)]))
    elements.append(t_eq)
    elements.append(Spacer(1, 6))

    # 5. Periféricos, Accesorios y Otros Activos[cite: 2, 3]
    elements.append(Paragraph("5. PERIFÉRICOS, ACCESORIOS Y OTROS ACTIVOS ASIGNADOS", section_style))
    periph_data = [[Paragraph("<b>Tipo de Activo</b>", bold_style), Paragraph("<b>Marca</b>", bold_style), Paragraph("<b>Modelo</b>", bold_style), Paragraph("<b>N° Serie</b>", bold_style), Paragraph("<b>Estado</b>", bold_style), Paragraph("<b>Observaciones</b>", bold_style)]]
    
    if perifericos:
        for p in perifericos:
            periph_data.append([
                Paragraph(p.get('tipo', ''), normal_style),
                Paragraph(p.get('marca', ''), normal_style),
                Paragraph(p.get('modelo', ''), normal_style),
                Paragraph(p.get('serie', ''), normal_style),
                Paragraph(p.get('estado', ''), normal_style),
                Paragraph(p.get('observaciones', ''), normal_style)
            ])
    else:
        periph_data.append([Paragraph("Ninguno", normal_style), Paragraph("-", normal_style), Paragraph("-", normal_style), Paragraph("-", normal_style), Paragraph("-", normal_style), Paragraph("-", normal_style)])

    t_periph = Table(periph_data, colWidths=[80, 75, 85, 75, 65, 170])
    t_periph.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f2f2f2')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3)
    ]))
    elements.append(t_periph)
    elements.append(Spacer(1, 6))

    # 7. Observaciones[cite: 3]
    elements.append(Paragraph("7. OBSERVACIONES Y SOPORTE", section_style))
    obs_text = row['observaciones'] if row['observaciones'] else "El equipo se entrega en óptimas condiciones operativas para el desempeño de funciones."
    data_obs = [[Paragraph(obs_text, normal_style)]]
    t_obs = Table(data_obs, colWidths=[550])
    t_obs.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('PADDING', (0,0), (-1,-1), 5)]))
    elements.append(t_obs)
    elements.append(Spacer(1, 6))

    # Registro Fotográfico (Cloudinary)[cite: 3]
    elements.append(Paragraph("REGISTRO FOTOGRÁFICO (CLOUDINARY)", section_style))
    foto_rows = []
    if fotos:
        for f in fotos:
            foto_rows.append([Paragraph(f"• Ver Adjunto Cloudinary: <a href='{f}' color='blue'>{f}</a>", normal_style)])
    else:
        foto_rows.append([Paragraph("• Ver Adjunto Ficha Cloudinary (Sin enlaces adicionales registrados)", normal_style)])
    
    t_fotos = Table(foto_rows, colWidths=[550])
    t_fotos.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('PADDING', (0,0), (-1,-1), 4)]))
    elements.append(t_fotos)
    elements.append(Spacer(1, 12))

    # Conformidad y Firmas[cite: 1, 3]
    data_firmas = [
        [Paragraph("________________________________________", normal_style), Paragraph("________________________________________", normal_style)],
        [Paragraph("<b>Sandra Elizabeth Hinojosa Rodriguez</b><br>Vassosp Trading EIRL<br>Responsable Soporte TI", normal_style), 
         Paragraph(f"<b>{row['firma']}</b><br>Trabajador / Responsable<br><b>Recí Conforme</b>", normal_style)]
    ]
    t_firmas = Table(data_firmas, colWidths=[275, 275])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(t_firmas)

    doc.build(elements)
    buffer.seek(0)
    
    response = make_response(buffer.getvalue())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'inline; filename=Ficha_INV_{row["id"]:03d}.pdf'
    return response

@app.route('/baja/<int:id>', methods=['POST'])
def baja(id):
    df = pd.read_csv(CSV_FILE)
    motivo_baja = request.form.get('motivo_baja', 'Obsolescencia / Daño')
    if id in df['id'].values:
        df.loc[df['id'] == id, 'estado'] = f"BAJA - {motivo_baja}"
        df.to_csv(CSV_FILE, index=False)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)