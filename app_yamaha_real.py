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
            'ID_Unico', 'N_Ficha', 'Fecha', 'Sede', 'ID_Trabajador', 'Responsable', 'Area', 'Cargo', 'Correo',
            'Tipo_Equipo', 'Marca_Modelo', 'N_Serie', 'Sistema_Operativo', 'Version_SO', 'Procesador',
            'Memoria_RAM', 'Part_Number', 'Hostname', 'Direccion_IP', 'Direccion_MAC', 'Tipo_Conexion', 
            'Punto_Red', 'Estado', 'Disco_Almacenamiento', 'Impresora_Instalada', 'Perifericos_Json',
            'Celular_IMEI', 'Celular_MarcaModelo', 'Celular_RAM', 'Celular_Almacenamiento', 'Celular_Procesador',
            'Celular_Telefono', 'Celular_SO', 'Celular_Accesorios', 'Celular_EstadoFisico', 'Observaciones',
            'Firma_Tecnico', 'Firma_Digital_Tecnico', 'Firma_Empleado', 'Firma_Digital_Empleado', 'Foto_Equipo'
        ])
        df_init.to_csv(DB_CSV, index=False)
    
    if not os.path.exists(TECNICOS_CSV):
        pd.DataFrame(columns=['Nombre', 'Especialidad', 'Firma_Digital']).to_csv(TECNICOS_CSV, index=False)
        
    if not os.path.exists(EMPLEADOS_CSV):
        pd.DataFrame(columns=['ID_Empleado', 'Nombre', 'Area', 'Cargo', 'Correo']).to_csv(EMPLEADOS_CSV, index=False)

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
        total_local = len(df[df['Sede'].astype(str).str.contains('Lima|Central|San Isidro', case=False, na=False)])
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

@app.route('/add_manual', methods=['POST'])
def add_manual():
    n_ficha = request.form.get('n_ficha')
    fecha = request.form.get('fecha')
    sede = request.form.get('sede')
    id_trabajador = request.form.get('id_trabajador', '')
    responsable = request.form.get('responsable', '')
    area = request.form.get('area', '')
    cargo = request.form.get('cargo', '')
    correo = request.form.get('correo', '')
    tipo_equipo = request.form.get('tipo_equipo', 'Laptop')
    
    # Datos de equipo principal / laptop
    marca_modelo = request.form.get('marca_modelo', '')
    n_serie = request.form.get('n_serie', '')
    sistema_operativo = request.form.get('sistema_operativo', '')
    version_so = request.form.get('version_so', '')
    procesador = request.form.get('procesador', '')
    memoria_ram = request.form.get('memoria_ram', '')
    part_number = request.form.get('part_number', '')
    hostname = request.form.get('hostname', '')
    direccion_ip = request.form.get('direccion_ip', '')
    direccion_mac = request.form.get('direccion_mac', '')
    tipo_conexion = request.form.get('tipo_conexion', '')
    punto_red = request.form.get('punto_red', '')
    estado = request.form.get('estado', 'Operativo')
    
    # Almacenamiento e Impresoras
    disco_almacenamiento = request.form.get('disco_almacenamiento', '')
    impresora_instalada = request.form.get('impresora_instalada', '')
    
    # Celular
    celular_imei = request.form.get('celular_imei', '')
    celular_marcamodelo = request.form.get('celular_marcamodelo', '')
    celular_ram = request.form.get('celular_ram', '')
    celular_almacenamiento = request.form.get('celular_almacenamiento', '')
    celular_procesador = request.form.get('celular_procesador', '')
    celular_telefono = request.form.get('celular_telefono', '')
    celular_so = request.form.get('celular_so', '')
    celular_accesorios = request.form.get('celular_accesorios', '')
    celular_estadofisico = request.form.get('celular_estadofisico', '')
    
    observaciones = request.form.get('observaciones', '')
    firma_tecnico = request.form.get('firma_tecnico', '')
    firma_digital_empleado = request.form.get('firma_digital_empleado', '')
    id_unico = str(uuid.uuid4())[:8]

    # Buscar firma digital del técnico automáticamente desde su registro
    firma_digital_tecnico = ""
    if os.path.exists(TECNICOS_CSV):
        df_tec = pd.read_csv(TECNICOS_CSV).fillna('')
        match_tec = df_tec[df_tec['Nombre'] == firma_tecnico]
        if not match_tec.empty:
            firma_digital_tecnico = match_tec.iloc[0]['Firma_Digital']

    # Subir fotos a Cloudinary organizadas por Ficha exacta
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
        'ID_Trabajador': id_trabajador,
        'Responsable': responsable,
        'Area': area,
        'Cargo': cargo,
        'Correo': correo,
        'Tipo_Equipo': tipo_equipo,
        'Marca_Modelo': marca_modelo,
        'N_Serie': n_serie,
        'Sistema_Operativo': sistema_operativo,
        'Version_SO': version_so,
        'Procesador': procesador,
        'Memoria_RAM': memoria_ram,
        'Part_Number': part_number,
        'Hostname': hostname,
        'Direccion_IP': direccion_ip,
        'Direccion_MAC': direccion_mac,
        'Tipo_Conexion': tipo_conexion,
        'Punto_Red': punto_red,
        'Estado': estado,
        'Disco_Almacenamiento': disco_almacenamiento,
        'Impresora_Instalada': impresora_instalada,
        'Perifericos_Json': '',
        'Celular_IMEI': celular_imei,
        'Celular_MarcaModelo': celular_marcamodelo,
        'Celular_RAM': celular_ram,
        'Celular_Almacenamiento': celular_almacenamiento,
        'Celular_Procesador': celular_procesador,
        'Celular_Telefono': celular_telefono,
        'Celular_SO': celular_so,
        'Celular_Accesorios': celular_accesorios,
        'Celular_EstadoFisico': celular_estadofisico,
        'Observaciones': observaciones,
        'Firma_Tecnico': firma_tecnico,
        'Firma_Digital_Tecnico': firma_digital_tecnico,
        'Firma_Empleado': responsable,
        'Firma_Digital_Empleado': firma_digital_empleado,
        'Foto_Equipo': foto_equipo_str
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
    
    if tipo == 'tecnico':
        nombre = request.form.get('nombre_tecnico')
        especialidad = request.form.get('especialidad_tecnico', '')
        firma_digital = request.form.get('firma_digital_tecnico_reg', '')
        
        nuevo_dato = {'Nombre': nombre, 'Especialidad': especialidad, 'Firma_Digital': firma_digital}
        file_path = TECNICOS_CSV
    else:
        id_emp = request.form.get('id_empleado_reg')
        nombre = request.form.get('nombre_empleado_reg')
        area = request.form.get('area_empleado_reg')
        cargo = request.form.get('cargo_empleado_reg')
        correo = request.form.get('correo_empleado_reg', '')
        
        nuevo_dato = {'ID_Empleado': id_emp, 'Nombre': nombre, 'Area': area, 'Cargo': cargo, 'Correo': correo}
        file_path = EMPLEADOS_CSV

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