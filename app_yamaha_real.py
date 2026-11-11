from flask import Flask, render_template, request, redirect, url_for, make_response, send_file
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
    columnas = [
        'id', 'tipo_equipo', 'codigo_patrimonial', 'marca', 'modelo', 
        'serie', 'procesador', 'memoria_ram', 'almacenamiento', 
        'usuario_asignado', 'area', 'sucursal', 'estado', 'perifericos', 'observaciones', 'fotos', 'firma'
    ]
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=columnas)
        df.to_csv(CSV_FILE, index=False)
    else:
        # Si el archivo ya existe pero le faltan columnas (como Estado_Ficha antiguo), las autocompleta de manera segura
        df = pd.read_csv(CSV_FILE)
        cambios = False
        for col in columnas:
            if col not in df.columns:
                df[col] = ''
                cambios = True
        if cambios:
            df.to_csv(CSV_FILE, index=False)

inicializar_csv()

@app.route('/')
def index():
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    
    # Filtros de búsqueda opcionales
    busqueda = request.args.get('q', '').lower()
    sucursal_filtro = request.args.get('sucursal', '')
    
    if busqueda:
        df = df[df.astype(str).apply(lambda x: x.str.lower().str.contains(busqueda).any(), axis=1)]
    if sucursal_filtro:
        df = df[df['sucursal'] == sucursal_filtro]

    registros = []
    for _, row in df.iterrows():
        reg = row.to_dict()
        
        # Valores por defecto seguros
        if 'estado' not in reg or pd.isna(reg['estado']):
            reg['estado'] = 'Activo'
        if 'firma' not in reg or pd.isna(reg['firma']):
            reg['firma'] = 'Sin Firma'
        if 'tipo_equipo' not in reg or pd.isna(reg['tipo_equipo']):
            reg['tipo_equipo'] = 'Laptop'
            
        for campo in ['codigo_patrimonial', 'marca', 'modelo', 'serie', 'usuario_asignado', 'area', 'sucursal', 'procesador', 'memoria_ram', 'almacenamiento', 'observaciones']:
            if campo not in reg or pd.isna(reg[campo]):
                reg[campo] = ''

        try:
            reg['perifericos_lista'] = json.loads(str(reg.get('perifericos', ''))) if pd.notna(reg.get('perifericos')) else []
        except:
            reg['perifericos_lista'] = []
            
        try:
            reg['fotos_lista'] = json.loads(str(reg.get('fotos', ''))) if pd.notna(reg.get('fotos')) else []
        except:
            reg['fotos_lista'] = []
            
        registros.append(reg)
        
    return render_template('index.html', registros=registros, busqueda=busqueda, sucursal_filtro=sucursal_filtro)

@app.route('/guardar', methods=['POST'])
def guardar():
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    
    registro_id = request.form.get('id', '')
    tipo_equipo = request.form.get('tipo_equipo', '').strip()
    almacenamiento = request.form.get('almacenamiento', '').strip()
    procesador = request.form.get('procesador', '').strip()
    memoria_ram = request.form.get('memoria_ram', '').strip()
    
    if tipo_equipo.lower() in ['impresora', 'switch', 'tv']:
        almacenamiento = ''
        procesador = ''
        memoria_ram = ''

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

    fotos_input = request.form.get('fotos', '').strip()
    lista_fotos = [f.strip() for f in fotos_input.split('\n') if f.strip()] if fotos_input else []

    datos_form = {
        'tipo_equipo': tipo_equipo,
        'codigo_patrimonial': request.form.get('codigo_patrimonial', '').strip(),
        'marca': request.form.get('marca', '').strip(),
        'modelo': request.form.get('modelo', '').strip(),
        'serie': request.form.get('serie', '').strip(),
        'procesador': procesador,
        'memoria_ram': memoria_ram,
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

    if registro_id:  # Editar existente
        idx = df[df['id'] == int(registro_id)].index
        if not idx.empty:
            for k, v in datos_form.items():
                df.loc[idx, k] = v
    else:  # Crear nuevo
        nuevo_id = int(df['id'].max()) + 1 if not df.empty and not pd.isna(df['id'].max()) else 1
        datos_form['id'] = nuevo_id
        df = pd.concat([df, pd.DataFrame([datos_form])], ignore_index=True)
    
    df.to_csv(CSV_FILE, index=False)
    return redirect(url_for('index'))

@app.route('/editar/<int:id>')
def editar(id):
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    reg = df[df['id'] == id]
    if reg.empty:
        return redirect(url_for('index'))
    
    row = reg.iloc[0].to_dict()
    try:
        row['perifericos_lista'] = json.loads(str(row.get('perifericos', ''))) if pd.notna(row.get('perifericos')) else []
    except:
        row['perifericos_lista'] = []
        
    try:
        row['fotos_lista'] = json.loads(str(row.get('fotos', ''))) if pd.notna(row.get('fotos')) else []
    except:
        row['fotos_lista'] = []
        
    registros = []
    for _, r in df.iterrows():
        rd = r.to_dict()
        try:
            rd['perifericos_lista'] = json.loads(str(rd.get('perifericos', ''))) if pd.notna(rd.get('perifericos')) else []
        except:
            rd['perifericos_lista'] = []
        try:
            rd['fotos_lista'] = json.loads(str(rd.get('fotos', ''))) if pd.notna(rd.get('fotos')) else []
        except:
            rd['fotos_lista'] = []
        registros.append(rd)

    return render_template('index.html', registros=registros, edit_reg=row)

@app.route('/pdf/<int:id>')
def generar_pdf(id):
    inicializar_csv()
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

    elements.append(Paragraph("<b>YAMAHA MOTOR / VASSOSP TRADING EIRL</b><br>ACTA Y FICHA DE INVENTARIO DE EQUIPOS TECNOLOGICOS", title_style))
    elements.append(Spacer(1, 5))

    elements.append(Paragraph("1. DATOS DEL TRABAJADOR / RESPONSABLE", section_style))
    data_collab = [
        [Paragraph("<b>N° Ficha / Acta:</b>", normal_style), Paragraph(f"INV-2026-{row['id']:03d}", normal_style), Paragraph("<b>Fecha:</b>", normal_style), Paragraph("2026-11-10", normal_style)],
        [Paragraph("<b>Colaborador:</b>", normal_style), Paragraph(str(row.get('usuario_asignado', '')), normal_style), Paragraph("<b>Área / Dpto:</b>", normal_style), Paragraph(str(row.get('area', '')), normal_style)],
        [Paragraph("<b>Sucursal:</b>", normal_style), Paragraph(str(row.get('sucursal', '')), normal_style), Paragraph("<b>Estado:</b>", normal_style), Paragraph(str(row.get('estado', '')), normal_style)]
    ]
    t_collab = Table(data_collab, colWidths=[70, 200, 70, 210])
    t_collab.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 3)]))
    elements.append(t_collab)
    elements.append(Spacer(1, 6))

    elements.append(Paragraph(f"2. DETALLE DE EQUIPO PRINCIPAL ({row.get('tipo_equipo', '')})", section_style))
    data_eq = [
        [Paragraph("<b>Tipo de Equipo:</b>", normal_style), Paragraph(str(row.get('tipo_equipo', '')), normal_style), Paragraph("<b>Patrimonial:</b>", normal_style), Paragraph(str(row.get('codigo_patrimonial', '')), normal_style)],
        [Paragraph("<b>Marca / Modelo:</b>", normal_style), Paragraph(f"{row.get('marca', '')} {row.get('modelo', '')}", normal_style), Paragraph("<b>N° Serie:</b>", normal_style), Paragraph(str(row.get('serie', '')), normal_style)],
        [Paragraph("<b>Procesador:</b>", normal_style), Paragraph(str(row.get('procesador', '')), normal_style), Paragraph("<b>Memoria RAM:</b>", normal_style), Paragraph(str(row.get('memoria_ram', '')), normal_style)],
        [Paragraph("<b>Almacenamiento:</b>", normal_style), Paragraph(str(row.get('almacenamiento', '')), normal_style), Paragraph("<b>Sucursal:</b>", normal_style), Paragraph(str(row.get('sucursal', '')), normal_style)]
    ]
    t_eq = Table(data_eq, colWidths=[90, 180, 90, 190])
    t_eq.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('PADDING', (0,0), (-1,-1), 3)]))
    elements.append(t_eq)
    elements.append(Spacer(1, 6))

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

    elements.append(Paragraph("7. OBSERVACIONES Y SOPORTE", section_style))
    obs_text = row.get('observaciones', '') if str(row.get('observaciones', '')).strip() else "El equipo se entrega en óptimas condiciones operativas para el desempeño de funciones."
    data_obs = [[Paragraph(obs_text, normal_style)]]
    t_obs = Table(data_obs, colWidths=[550])
    t_obs.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.grey), ('PADDING', (0,0), (-1,-1), 5)]))
    elements.append(t_obs)
    elements.append(Spacer(1, 6))

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

    data_firmas = [
        [Paragraph("________________________________________", normal_style), Paragraph("________________________________________", normal_style)],
        [Paragraph("<b>Sandra Elizabeth Hinojosa Rodriguez</b><br>Vassosp Trading EIRL<br><b>Responsable Soporte TI (Técnico)</b>", normal_style), 
         Paragraph(f"<b>{row.get('firma', '')}</b><br>Trabajador / Responsable (Empleado)<br><b>Recí Conforme</b>", normal_style)]
    ]
    t_firmas = Table(data_firmas, colWidths=[275, 275])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(t_firmas)

    doc.build(elements)
    buffer.seek(0)
    
    response = make_response(buffer.getvalue())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'inline; filename=Acta_Ficha_INV_{row["id"]:03d}.pdf'
    return response

@app.route('/exportar_excel')
def exportar_excel():
    inicializar_csv()
    if os.path.exists(CSV_FILE):
        df = pd.read_csv(CSV_FILE)
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='inventario_yamaha.xlsx')
    return redirect(url_for('index'))

@app.route('/baja/<int:id>', methods=['POST'])
def baja(id):
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    motivo_baja = request.form.get('motivo_baja', 'Obsolescencia / Daño')
    if id in df['id'].values:
        df.loc[df['id'] == id, 'estado'] = f"BAJA - {motivo_baja}"
        df.to_csv(CSV_FILE, index=False)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)