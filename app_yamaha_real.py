from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import os
import json
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)

CSV_FILE = 'base_datos_yamaha_final.csv'
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def inicializar_csv():
    columnas = [
        'id', 'tipo_equipo', 'codigo_patrimonial', 'marca', 'modelo', 
        'serie', 'procesador', 'memoria_ram', 'almacenamiento', 
        'usuario_asignado', 'area', 'sucursal', 'estado', 'perifericos', 
        'observaciones', 'fotos', 'firma', 'tecnico_soporte', 
        'celular_asignado', 'celular_numero', 'celular_modelo'
    ]
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=columnas)
        df.to_csv(CSV_FILE, index=False)
    else:
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
    
    busqueda = request.args.get('q', '').lower()
    sucursal_filtro = request.args.get('sucursal', '')
    
    if busqueda:
        df = df[df.astype(str).apply(lambda x: x.str.lower().str.contains(busqueda).any(), axis=1)]
    if sucursal_filtro:
        df = df[df['sucursal'] == sucursal_filtro]

    registros = []
    total_celulares = 0
    total_activos = 0

    for _, row in df.iterrows():
        reg = row.to_dict()
        
        if 'estado' not in reg or pd.isna(reg['estado']):
            reg['estado'] = 'Activo'
        if 'firma' not in reg or pd.isna(reg['firma']):
            reg['firma'] = 'Sin Firma'
        if 'tipo_equipo' not in reg or pd.isna(reg['tipo_equipo']):
            reg['tipo_equipo'] = 'Laptop'
        if 'tecnico_soporte' not in reg or pd.isna(reg['tecnico_soporte']) or not str(reg['tecnico_soporte']).strip():
            reg['tecnico_soporte'] = 'Sandra Elizabeth Hinojosa Rodriguez'
            
        for campo in ['codigo_patrimonial', 'marca', 'modelo', 'serie', 'usuario_asignado', 'area', 'sucursal', 'procesador', 'memoria_ram', 'almacenamiento', 'observaciones', 'celular_asignado', 'celular_numero', 'celular_modelo']:
            if campo not in reg or pd.isna(reg[campo]):
                reg[campo] = ''

        if 'BAJA' not in str(reg['estado']).upper():
            total_activos += 1
        if str(reg.get('celular_asignado', '')).strip().lower() in ['si', 'sí', 'true', '1']:
            total_celulares += 1

        try:
            reg['perifericos_lista'] = json.loads(str(reg.get('perifericos', ''))) if pd.notna(reg.get('perifericos')) else []
        except:
            reg['perifericos_lista'] = []
            
        try:
            reg['fotos_lista'] = json.loads(str(reg.get('fotos', ''))) if pd.notna(reg.get('fotos')) else []
        except:
            reg['fotos_lista'] = []
            
        registros.append(reg)
        
    metricas = {
        'total': len(registros),
        'activos': total_activos,
        'celulares': total_celulares
    }
        
    return render_template('index.html', registros=registros, busqueda=busqueda, sucursal_filtro=sucursal_filtro, metricas=metricas)

@app.route('/guardar', methods=['POST'])
def guardar():
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    
    registro_id = request.form.get('id', '')
    tipo_equipo = request.form.get('tipo_equipo', '').strip()
    almacenamiento = request.form.get('almacenamiento', '').strip()
    procesador = request.form.get('procesador', '').strip()
    memoria_ram = request.form.get('memoria_ram', '').strip()
    
    if tipo_equipo.lower() in ['impresora', 'switch', 'tv', 'scanner-pda']:
        almacenamiento = ''
        procesador = ''
        memoria_ram = ''

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

    lista_fotos = []
    if registro_id:
        reg_old = df[df['id'] == int(registro_id)]
        if not reg_old.empty:
            try:
                lista_fotos = json.loads(str(reg_old.iloc[0].get('fotos', '')))
            except:
                lista_fotos = []

    archivos_fotos = request.files.getlist('fotos_archivos')
    for archivo in archivos_fotos:
        if archivo and archivo.filename != '':
            filename = secure_filename(archivo.filename)
            filename_final = f"eq_{registro_id}_{filename}" if registro_id else f"eq_new_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_final)
            archivo.save(filepath)
            lista_fotos.append(f"/static/uploads/{filename_final}")

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
        'firma': request.form.get('firma', '').strip(),
        'tecnico_soporte': request.form.get('tecnico_soporte', 'Sandra Elizabeth Hinojosa Rodriguez').strip(),
        'celular_asignado': request.form.get('celular_asignado', 'No').strip(),
        'celular_numero': request.form.get('celular_numero', '').strip(),
        'celular_modelo': request.form.get('celular_modelo', '').strip()
    }

    if registro_id:
        idx = df[df['id'] == int(registro_id)].index
        if not idx.empty:
            for k, v in datos_form.items():
                df.loc[idx, k] = v
    else:
        nuevo_id = int(df['id'].max()) + 1 if not df.empty and not pd.isna(df['id'].max()) else 1
        datos_form['id'] = nuevo_id
        df = pd.concat([df, pd.DataFrame([datos_form])], ignore_index=True)
    
    df.to_csv(CSV_FILE, index=False)
    return redirect(url_for('index'))

@app.route('/importar_masivo', methods=['POST'])
def importar_masivo():
    inicializar_csv()
    archivo = request.files.get('archivo_excel')
    if not archivo or archivo.filename == '':
        return redirect(url_for('index'))
    
    try:
        df_subido = pd.read_excel(archivo)
        df_actual = pd.read_csv(CSV_FILE)
        
        nuevo_id_base = int(df_actual['id'].max()) + 1 if not df_actual.empty and not pd.isna(df_actual['id'].max()) else 1
        
        nuevos_registros = []
        for idx, row in df_subido.iterrows():
            reg = {
                'id': nuevo_id_base + idx,
                'tipo_equipo': str(row.get('tipo_equipo', 'Laptop')),
                'codigo_patrimonial': str(row.get('codigo_patrimonial', 'PAT-000')),
                'marca': str(row.get('marca', 'Yamaha')),
                'modelo': str(row.get('modelo', '')),
                'serie': str(row.get('serie', 'SN-000')),
                'procesador': str(row.get('procesador', '')),
                'memoria_ram': str(row.get('memoria_ram', '')),
                'almacenamiento': str(row.get('almacenamiento', '')),
                'usuario_asignado': str(row.get('usuario_asignado', 'Sin Asignar')),
                'area': str(row.get('area', 'General')),
                'sucursal': str(row.get('sucursal', 'Lima')),
                'estado': 'Activo',
                'perifericos': '[]',
                'observaciones': str(row.get('observaciones', 'Carga masiva por bloques')),
                'fotos': '[]',
                'firma': str(row.get('usuario_asignado', 'Validado')),
                'tecnico_soporte': 'Sandra Elizabeth Hinojosa Rodriguez',
                'celular_asignado': str(row.get('celular_asignado', 'No')),
                'celular_numero': str(row.get('celular_numero', '')),
                'celular_modelo': str(row.get('celular_modelo', ''))
            }
            nuevos_registros.append(reg)
            
        if nuevos_registros:
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_final = pd.concat([df_actual, df_nuevos], ignore_index=True)
            df_final.to_csv(CSV_FILE, index=False)
            
    except Exception as e:
        print("Error en carga masiva:", e)
        
    return redirect(url_for('index'))

@app.route('/plantilla_masiva')
def plantilla_masiva():
    columnas_plantilla = [
        'tipo_equipo', 'codigo_patrimonial', 'marca', 'modelo', 'serie', 
        'procesador', 'memoria_ram', 'almacenamiento', 'usuario_asignado', 
        'area', 'sucursal', 'celular_asignado', 'celular_numero', 'celular_modelo', 'observaciones'
    ]
    df_plantilla = pd.DataFrame(columns=columnas_plantilla)
    df_plantilla.loc[0] = [
        'Laptop', 'PAT-10293', 'HP', '250 G8', 'CND8392', 
        'Intel i3', '8GB', '256GB SSD', 'Carlos Mendoza', 
        'Logística', 'Lima', 'No', '', '', 'Entrega conforme'
    ]
    
    output = io.BytesIO()
    df_plantilla.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='plantilla_carga_masiva_yamaha.xlsx')

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

    metricas = {'total': len(registros), 'activos': len(registros), 'celulares': 0}
    return render_template('index.html', registros=registros, edit_reg=row, metricas=metricas)

@app.route('/pdf/<int:id>')
def generar_pdf(id):
    inicializar_csv()
    df = pd.read_csv(CSV_FILE)
    reg = df[df['id'] == id]
    if reg.empty:
        return "Registro no encontrado", 404
    
    row = reg.iloc[0].to_dict()
    try:
        row['perifericos_lista'] = json.loads(str(row.get('perifericos', ''))) if pd.notna(row.get('perifericos')) else []
    except:
        row['perifericos_lista'] = []
        
    try:
        row['fotos_lista'] = json.loads(str(row.get('fotos', ''))) if pd.notna(row.get('fotos')) else []
    except:
        row['fotos_lista'] = []

    # Renderiza directamente tu archivo acta.html pasándole todos los datos del registro (row)
    return render_template('acta.html', row=row)

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