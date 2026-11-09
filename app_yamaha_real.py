import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename
import cloudinary
import cloudinary.uploader
import cloudinary.api
from pypdf import PdfReader

app = Flask(__name__)

# Configuración de Cloudinary (lee las variables de entorno de Render)
cloudinary.config(
    cloud_name = os.environ.get('CLOUDINARY_CLOUD_NAME'),
    api_key = os.environ.get('CLOUDINARY_API_KEY'),
    api_secret = os.environ.get('CLOUDINARY_API_SECRET'),
    secure = True
)

DB_CSV = 'base_datos_yamaha_final.csv'
TECNICOS_CSV = 'personal_tecnicos.csv'
EMPLEADOS_CSV = 'personal_empleados.csv'
EXCEL_OUTPUT = 'Inventario_Corporativo_Yamaha_Final.xlsx'

def inicializar_archivos():
    if not os.path.exists(DB_CSV):
        df_init = pd.DataFrame(columns=[
            'ID_Unico', 'N_Ficha', 'Fecha', 'Sede', 'Responsable', 'Cargo', 
            'Tipo_Equipo', 'Modelo', 'N_Serie', 'IP', 'Sistema_Operativo', 
            'IMEI_o_Detalle', 'Procesador_RAM', 'Perifericos', 'Observaciones', 
            'Firma_Tecnico', 'Firma_Empleado', 'Firma_Digital_Empleado', 'Foto_Equipo', 'Estado'
        ])
        df_init.to_csv(DB_CSV, index=False)
    
    if not os.path.exists(TECNICOS_CSV):
        pd.DataFrame(columns=['Nombre', 'Cargo']).to_csv(TECNICOS_CSV, index=False)
        
    if not os.path.exists(EMPLEADOS_CSV):
        pd.DataFrame(columns=['Nombre', 'Cargo']).to_csv(EMPLEADOS_CSV, index=False)

inicializar_archivos()

@app.route('/')
def index():
    registros = []
    total_equipos = total_local = total_provincia = 0
    tecnicos = []
    empleados = []

    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV).fillna('')
        registros = df.to_dict(orient='records')
        total_equipos = len(df)
        total_local = len(df[df['Sede'].astype(str).str.contains('Lima|Central', case=False, na=False)])
        total_provincia = total_equipos - total_local

    if os.path.exists(TECNICOS_CSV):
        df_tec = pd.read_csv(TECNICOS_CSV).fillna('')
        tecnicos = df_tec.to_dict(orient='records')

    if os.path.exists(EMPLEADOS_CSV):
        df_emp = pd.read_csv(EMPLEADOS_CSV).fillna('')
        empleados = df_emp.to_dict(orient='records')

    return render_template('index.html', 
                           registros=registros, 
                           total_equipos=total_equipos, 
                           total_local=total_local, 
                           total_provincia=total_provincia,
                           tecnicos=tecnicos,
                           empleados=empleados)

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return redirect(url_for('index'))

    os.makedirs('reportes_pdf', exist_ok=True)
    
    for file in files:
        if file and file.filename.endswith('.pdf'):
            filename = secure_filename(file.filename)
            filepath = os.path.join('reportes_pdf', filename)
            file.save(filepath)
            
            try:
                reader = PdfReader(filepath)
                id_unico = str(uuid.uuid4())[:8]
                n_ficha_gen = filename.replace('.pdf', '')
                
                nuevo_registro = {
                    'ID_Unico': id_unico,
                    'N_Ficha': n_ficha_gen,
                    'Fecha': pd.Timestamp.today().strftime('%Y-%m-%d'),
                    'Sede': 'Lima / Central',
                    'Responsable': 'Por actualizar',
                    'Cargo': 'Colaborador',
                    'Tipo_Equipo': 'PC / Laptop',
                    'Modelo': 'Genérico por PDF',
                    'N_Serie': 'S/N',
                    'IP': '0.0.0.0',
                    'Sistema_Operativo': 'Windows 11 Pro',
                    'IMEI_o_Detalle': 'N/A',
                    'Procesador_RAM': 'Por definir',
                    'Perifericos': 'Mouse, Teclado',
                    'Observaciones': 'Ingresado vía PDF',
                    'Firma_Tecnico': 'Soporte TI',
                    'Firma_Empleado': 'Por actualizar',
                    'Firma_Digital_Empleado': '',
                    'Foto_Equipo': 'Sin foto',
                    'Estado': 'Activo'
                }
                
                df = pd.read_csv(DB_CSV)
                df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                df.to_csv(DB_CSV, index=False)
            except Exception as e:
                print(f"Error procesando PDF {filename}: {e}")

    return redirect(url_for('index'))

@app.route('/add_manual', methods=['POST'])
def add_manual():
    n_ficha = request.form.get('n_ficha')
    fecha = request.form.get('fecha')
    sede = request.form.get('sede')
    responsable = request.form.get('responsable')
    cargo = request.form.get('cargo', '')
    tipo_equipo = request.form.get('tipo_equipo')
    modelo = request.form.get('modelo')
    n_serie = request.form.get('n_serie')
    ip = request.form.get('ip', '')
    sistema_operativo = request.form.get('sistema_operativo', '')
    imei_o_detalle = request.form.get('imei_o_detalle', '')
    procesador_ram = request.form.get('procesador_ram', '')
    perifericos = request.form.get('perifericos', '')
    observaciones = request.form.get('observaciones', '')
    firma_tecnico = request.form.get('firma_tecnico', '')
    firma_digital_empleado = request.form.get('firma_digital_empleado', '')
    id_unico = str(uuid.uuid4())[:8]

    # Cloudinary: Organizar fotos en carpetas independientes por cada N_Ficha
    fotos_urls = []
    archivos = request.files.getlist('fotos_equipo')
    folder_path = f"Yamaha_Inventario/Ficha_{n_ficha}"
    
    for archivo in archivos:
        if archivo and archivo.filename != '':
            try:
                upload_result = cloudinary.uploader.upload(archivo, folder=folder_path)
                fotos_urls.append(upload_result.get('secure_url'))
            except Exception as e:
                print(f"Error subiendo a Cloudinary: {e}")

    foto_equipo_str = ",".join(fotos_urls) if fotos_urls else "Sin foto"

    nuevo_registro = {
        'ID_Unico': id_unico,
        'N_Ficha': n_ficha,
        'Fecha': fecha,
        'Sede': sede,
        'Responsable': responsable,
        'Cargo': cargo,
        'Tipo_Equipo': tipo_equipo,
        'Modelo': modelo,
        'N_Serie': n_serie,
        'IP': ip,
        'Sistema_Operativo': sistema_operativo,
        'IMEI_o_Detalle': imei_o_detalle,
        'Procesador_RAM': procesador_ram,
        'Perifericos': perifericos,
        'Observaciones': observaciones,
        'Firma_Tecnico': firma_tecnico,
        'Firma_Empleado': responsable,
        'Firma_Digital_Empleado': firma_digital_empleado,
        'Foto_Equipo': foto_equipo_str,
        'Estado': 'Activo'
    }

    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    else:
        df = pd.DataFrame([nuevo_registro])
    
    df.to_csv(DB_CSV, index=False)
    return redirect(url_for('index'))

@app.route('/add_personal', methods=['POST'])
def add_personal():
    tipo = request.form.get('tipo_personal')
    nombre = request.form.get('nombre_personal')
    cargo = request.form.get('cargo_personal', '')

    nuevo_dato = {'Nombre': nombre, 'Cargo': cargo}
    file_path = TECNICOS_CSV if tipo == 'tecnico' else EMPLEADOS_CSV

    if os.path.exists(file_path):
        df = pd.read_csv(file_path)
        df = pd.concat([df, pd.DataFrame([nuevo_dato])], ignore_index=True)
    else:
        df = pd.DataFrame([nuevo_dato])

    df.to_csv(file_path, index=False)
    return redirect(url_for('index'))

@app.route('/delete/<id_unico>')
def delete(id_unico):
    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        df = df[df['ID_Unico'].astype(str) != str(id_unico)]
        df.to_csv(DB_CSV, index=False)
    return redirect(url_for('index'))

@app.route('/imprimir/<id_unico>')
def imprimir(id_unico):
    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV).fillna('')
        registro = df[df['ID_Unico'].astype(str) == str(id_unico)]
        if not registro.empty:
            item = registro.iloc[0].to_dict()
            return render_template('acta.html', item=item)
    return "Registro no encontrado", 404

@app.route('/download_excel')
def download_excel():
    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        df.to_excel(EXCEL_OUTPUT, index=False)
        return send_file(EXCEL_OUTPUT, as_attachment=True)
    return "No hay datos para exportar", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)