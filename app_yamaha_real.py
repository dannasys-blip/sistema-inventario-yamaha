import os
import re
import datetime
import pandas as pd
import pypdf
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
import cloudinary
import cloudinary.uploader

app = Flask(__name__)
app.secret_key = 'yamaha_inventario_personal_separado'

# Configuración de Cloudinary (lee automáticamente las variables de entorno de Render)
cloudinary.config(
    cloud_name = os.environ.get('CLOUDINARY_CLOUD_NAME'),
    api_key = os.environ.get('CLOUDINARY_API_KEY'),
    api_secret = os.environ.get('CLOUDINARY_API_SECRET')
)

UPLOAD_FOLDER = 'reportes_pdf'
OUTPUT_EXCEL = 'Inventario_Corporativo_Yamaha_Final.xlsx'
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'webp'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True) # Asegura que exista la carpeta templates

DB_EDICIONES = 'base_datos_yamaha_final.csv'
DB_TECNICOS = 'personal_tecnicos.csv'
DB_EMPLEADOS = 'personal_empleados.csv'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cargar_tecnicos():
    if os.path.exists(DB_TECNICOS):
        return pd.read_csv(DB_TECNICOS, dtype={'DNI': str})
    df = pd.DataFrame([
        {"DNI": "45612378", "Nombre": "Carlos Mendoza", "Cargo": "Especialista Soporte TI Sede Central"},
        {"DNI": "48923145", "Nombre": "Luis Alberto Torres", "Cargo": "Analista de Infraestructura y Redes"},
        {"DNI": "41236547", "Nombre": "Marco Antonio Ruiz", "Cargo": "Técnico de Campo - Provincias"}
    ])
    df.to_csv(DB_TECNICOS, index=False)
    return df

def cargar_empleados():
    if os.path.exists(DB_EMPLEADOS):
        return pd.read_csv(DB_EMPLEADOS, dtype={'DNI': str})
    df = pd.DataFrame([
        {"DNI": "40123456", "Nombre": "Ana María Gamarra", "Cargo": "Coordinadora de Administración"},
        {"DNI": "42365478", "Nombre": "Jorge Luis Huamán", "Cargo": "Supervisor de Operaciones"},
        {"DNI": "43981234", "Nombre": "Rosa Elena Paredes", "Cargo": "Especialista de Logística y Almacén"},
        {"DNI": "47852361", "Nombre": "David Alejandro Silva", "Cargo": "Asistente de Finanzas y Contabilidad"}
    ])
    df.to_csv(DB_EMPLEADOS, index=False)
    return df

def cargar_registros_base():
    if os.path.exists(DB_EDICIONES):
        df = pd.read_csv(DB_EDICIONES, dtype={'ID_Unico': str})
        if 'Estado' not in df.columns:
            df['Estado'] = 'Operativo'
        if 'Firma_Tecnico' not in df.columns:
            df['Firma_Tecnico'] = 'Soporte TI Yamaha'
        if 'Firma_Empleado' not in df.columns:
            df['Firma_Empleado'] = '-'
        if 'Foto_Equipo' not in df.columns:
            df['Foto_Equipo'] = 'Sin foto'
        return df
    return pd.DataFrame(columns=[
        "ID_Unico", "Archivo_PDF", "N_Ficha", "Fecha", "Sede", "Clasificacion",
        "ID_Trabajador", "Responsable", "Cargo", "Tipo_Equipo", "Modelo",
        "N_Serie", "Atributo_1", "Atributo_2", "Atributo_3", "IP", "Estado",
        "Firma_Tecnico", "Firma_Empleado", "Foto_Equipo"
    ])

def guardar_registros_base(df):
    df['ID_Unico'] = df['ID_Unico'].astype(str)
    df.to_csv(DB_EDICIONES, index=False)
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Inventario General', index=False)
        if not df.empty and 'Tipo_Equipo' in df.columns:
            resumen = df[df['Estado'] != 'Dado de baja']['Tipo_Equipo'].value_counts().reset_index()
            resumen.columns = ['Tipo de Equipo', 'Cantidad']
            resumen.to_excel(writer, sheet_name='Resumen por Activo', index=False)

def generar_siguiente_ficha():
    anio_actual = datetime.datetime.now().year
    df = cargar_registros_base()
    if not df.empty and 'N_Ficha' in df.columns:
        fichas_anio = df[df['N_Ficha'].str.contains(f'INV-{anio_actual}', na=False)]
        siguiente_num = len(fichas_anio) + 1
    else:
        siguiente_num = 1
    return f"INV-{anio_actual}-{siguiente_num:04d}"

def procesar_nuevos_archivos(carpeta_pdf):
    df_actual = cargar_registros_base()
    archivos_existentes = set(df_actual['Archivo_PDF'].values) if not df_actual.empty else set()
    
    nuevos_registros = []
    if os.path.exists(carpeta_pdf):
        for file_name in os.listdir(carpeta_pdf):
            if file_name.lower().endswith('.pdf') and file_name not in archivos_existentes:
                file_path = os.path.join(carpeta_pdf, file_name)
                try:
                    reader = pypdf.PdfReader(file_path)
                    text = ""
                    for page in reader.pages:
                        text += page.extract_text() or ""
                    
                    sede = re.search(r'Sede/ Ubicación:\s*([^\n]+)', text, re.IGNORECASE)
                    trabajador = re.search(r'Nombre Completo:\s*([^\n]+)', text, re.IGNORECASE)
                    cargo = re.search(r'Cargo:\s*([^\n]+)', text, re.IGNORECASE)
                    modelo = re.search(r'Marca / Modelo:\s*([^\n]+)', text, re.IGNORECASE)
                    serie = re.search(r'N° Serie:\s*([^\n]+)', text, re.IGNORECASE)
                    ip = re.search(r'Dirección IP:\s*([^\n]+)', text, re.IGNORECASE)
                    estado_pdf = re.search(r'Estado:\s*([^\n]+)', text, re.IGNORECASE)

                    sede_val = sede.group(1).strip() if sede else "Sede Central - San Isidro"
                    tipo_val = "Laptop"
                    text_l = text.lower()
                    if any(k in text_l for k in ['tv', 'televisor', 'pantalla', 'smart tv', 'pulgadas']):
                        tipo_val = 'Televisor (TV)'
                    elif any(k in text_l for k in ['switch', 'router', 'access point', 'vlan', 'puertos']):
                        tipo_val = 'Switch / Red'
                    elif any(k in text_l for k in ['impresora', 'printer', 'epson', 'laserjet', 'toner']):
                        tipo_val = 'Impresora'
                    elif any(k in text_l for k in ['celular', 'smartphone', 'imei', 'línea']):
                        tipo_val = 'Celular'
                    elif any(k in text_l for k in ['pc', 'desktop', 'optiplex', 'thinkcentre']):
                        tipo_val = 'PC de Escritorio'

                    tipo_ubicacion = "Local" if any(l in sede_val.lower() for l in ["lima", "san isidro", "local", "central"]) else "Provincia"
                    
                    next_id = str(int(df_actual['ID_Unico'].astype(int).max() + 1)) if not df_actual.empty and df_actual['ID_Unico'].astype(str).str.isdigit().any() else str(len(df_actual) + len(nuevos_registros) + 1)
                    nuevo_codigo_ficha = generar_siguiente_ficha()

                    nuevos_registros.append({
                        "ID_Unico": next_id,
                        "Archivo_PDF": file_name,
                        "N_Ficha": nuevo_codigo_ficha,
                        "Fecha": datetime.datetime.now().strftime('%Y-%m-%d'),
                        "Sede": sede_val,
                        "Clasificacion": tipo_ubicacion,
                        "ID_Trabajador": "-",
                        "Responsable": trabajador.group(1).strip() if trabajador else "Sin Asignar",
                        "Cargo": cargo.group(1).strip() if cargo else "-",
                        "Tipo_Equipo": tipo_val,
                        "Modelo": modelo.group(1).strip() if modelo else "Desconocido",
                        "N_Serie": serie.group(1).strip() if serie else "S/N",
                        "Atributo_1": "-",
                        "Atributo_2": "-",
                        "Atributo_3": "-",
                        "IP": ip.group(1).strip() if ip else "-",
                        "Estado": estado_pdf.group(1).strip() if estado_pdf else "Operativo",
                        "Firma_Tecnico": "Carlos Mendoza",
                        "Firma_Empleado": trabajador.group(1).strip() if trabajador else "Pendiente",
                        "Foto_Equipo": "Sin foto"
                    })
                except Exception as e:
                    print(f"Error procesando {file_name}: {e}")

    if nuevos_registros:
        df_nuevo = pd.DataFrame(nuevos_registros)
        df_combined = pd.concat([df_actual, df_nuevo], ignore_index=True) if not df_actual.empty else df_nuevo
        guardar_registros_base(df_combined)

@app.route('/')
def index():
    procesar_nuevos_archivos(UPLOAD_FOLDER)
    df = cargar_registros_base()
    if not df.empty:
        df_vigentes = df[df['Estado'] != 'Dado de baja']
        total_equipos = len(df_vigentes)
        total_local = len(df_vigentes[df_vigentes['Clasificacion'] == 'Local'])
        total_provincia = len(df_vigentes[df_vigentes['Clasificacion'] == 'Provincia'])
        registros = df_vigentes.to_dict(orient='records')
    else:
        total_equipos = 0
        total_local = 0
        total_provincia = 0
        registros = []

    df_tec = cargar_tecnicos()
    tecnicos_dict = df_tec.set_index('DNI').to_dict(orient='index')
    tecnicos_json = df_tec.set_index('DNI').to_dict(orient='index')

    df_emp = cargar_empleados()
    empleados_dict = df_emp.set_index('DNI').to_dict(orient='index')
    empleados_json = df_emp.set_index('DNI').to_dict(orient='index')

    return render_template(
        'index.html',
        registros=registros,
        total_equipos=total_equipos,
        total_local=total_local,
        total_provincia=total_provincia,
        tecnicos=tecnicos_dict,
        tecnicos_json=tecnicos_json,
        empleados=empleados_dict,
        empleados_json=empleados_json,
        personal_tecnicos=df_tec.to_dict(orient='records'),
        personal_empleados=df_emp.to_dict(orient='records')
    )

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return redirect(url_for('index'))
    files = request.files.getlist('files')
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))
    return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)

@app.route('/download_excel')
def download_excel():
    if os.path.exists(OUTPUT_EXCEL):
        return send_from_directory('.', OUTPUT_EXCEL, as_attachment=True)
    return redirect(url_for('index'))

@app.route('/delete/<id_unico>')
def delete_registro(id_unico):
    df = cargar_registros_base()
    if not df.empty and id_unico in df['ID_Unico'].values:
        df.loc[df['ID_Unico'] == id_unico, 'Estado'] = 'Dado de baja'
        guardar_registros_base(df)
    return redirect(url_for('index'))

@app.route('/update', methods=['POST'])
def update_registro():
    id_unico = request.form.get('id_unico')
    df = cargar_registros_base()
    if not df.empty and id_unico in df['ID_Unico'].values:
        idx = df[df['ID_Unico'] == id_unico].index[0]
        df.at[idx, 'Tipo_Equipo'] = request.form.get('tipo_equipo', df.at[idx, 'Tipo_Equipo'])
        df.at[idx, 'Modelo'] = request.form.get('modelo', df.at[idx, 'Modelo'])
        df.at[idx, 'N_Serie'] = request.form.get('n_serie', df.at[idx, 'N_Serie'])
        df.at[idx, 'Sede'] = request.form.get('sede', df.at[idx, 'Sede'])
        df.at[idx, 'IP'] = request.form.get('ip', df.at[idx, 'IP'])
        df.at[idx, 'Firma_Tecnico'] = request.form.get('firma_tecnico', df.at[idx, 'Firma_Tecnico'])
        df.at[idx, 'Responsable'] = request.form.get('responsable', df.at[idx, 'Responsable'])
        df.at[idx, 'Cargo'] = request.form.get('cargo', df.at[idx, 'Cargo'])
        df.at[idx, 'Firma_Empleado'] = request.form.get('firma_empleado', df.at[idx, 'Firma_Empleado'])
        df.at[idx, 'Atributo_1'] = request.form.get('atributo_1', '-')
        df.at[idx, 'Atributo_2'] = request.form.get('atributo_2', '-')
        df.at[idx, 'Atributo_3'] = request.form.get('atributo_3', '-')

        sede_val = str(df.at[idx, 'Sede'])
        df.at[idx, 'Clasificacion'] = "Local" if any(l in sede_val.lower() for l in ["lima", "san isidro", "local", "central"]) else "Provincia"

        # Subida de fotos directamente a Cloudinary para persistencia en Render
        files = request.files.getlist('fotos_equipo')
        urls_fotos = []
        for file in files:
            if file and allowed_file(file.filename):
                try:
                    upload_result = cloudinary.uploader.upload(file)
                    secure_url = upload_result.get('secure_url')
                    if secure_url:
                        urls_fotos.append(secure_url)
                except Exception as e:
                    print(f"Error al subir imagen a Cloudinary: {e}")

        if urls_fotos:
            actuales = str(df.at[idx, 'Foto_Equipo'])
            if actuales and actuales != 'Sin foto':
                lista_act = [f.strip() for f in actuales.split(',')]
                urls_fotos = list(set(lista_act + urls_fotos))
            df.at[idx, 'Foto_Equipo'] = ','.join(urls_fotos)

        guardar_registros_base(df)
    return redirect(url_for('index'))

@app.route('/agregar_manual', methods=['POST'])
def agregar_manual():
    df = cargar_registros_base()
    next_id = str(int(df['ID_Unico'].astype(int).max() + 1)) if not df.empty and df['ID_Unico'].astype(str).str.isdigit().any() else str(len(df) + 1)
    ficha = generar_siguiente_ficha()

    sede_val = request.form.get('sede', 'Sede Central - San Isidro')
    tipo_ubicacion = "Local" if any(l in sede_val.lower() for l in ["lima", "san isidro", "local", "central"]) else "Provincia"

    nuevo_reg = {
        "ID_Unico": next_id,
        "Archivo_PDF": "Manual",
        "N_Ficha": ficha,
        "Fecha": datetime.datetime.now().strftime('%Y-%m-%d'),
        "Sede": sede_val,
        "Clasificacion": tipo_ubicacion,
        "ID_Trabajador": "-",
        "Responsable": request.form.get('responsable', 'Sin Asignar'),
        "Cargo": request.form.get('cargo', '-'),
        "Tipo_Equipo": request.form.get('tipo_equipo', 'Laptop'),
        "Modelo": request.form.get('modelo', 'Desconocido'),
        "N_Serie": request.form.get('n_serie', 'S/N'),
        "Atributo_1": request.form.get('atributo_1', '-'),
        "Atributo_2": request.form.get('atributo_2', '-'),
        "Atributo_3": request.form.get('atributo_3', '-'),
        "IP": request.form.get('ip', '-'),
        "Estado": "Operativo",
        "Firma_Tecnico": request.form.get('firma_tecnico', 'Carlos Mendoza'),
        "Firma_Empleado": request.form.get('firma_empleado', 'Pendiente'),
        "Foto_Equipo": "Sin foto"
    }

    df_nuevo = pd.DataFrame([nuevo_reg])
    df_combined = pd.concat([df, df_nuevo], ignore_index=True) if not df.empty else df_nuevo
    guardar_registros_base(df_combined)
    return redirect(url_for('index'))

@app.route('/imprimir/<id_unico>')
def imprimir(id_unico):
    df = cargar_registros_base()
    if not df.empty and id_unico in df['ID_Unico'].values:
        item = df[df['ID_Unico'] == id_unico].to_dict(orient='records')[0]
        return render_template('acta.html', item=item)
    return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)