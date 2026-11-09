import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file
import cloudinary
import cloudinary.uploader
import cloudinary.api
from datetime import datetime

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
        pd.DataFrame(columns=[
            'ID_Unico', 'N_Ficha', 'Fecha', 'Sede', 'ID_Trabajador', 'Responsable', 'Area', 'Cargo', 'Correo',
            'Tipo_Equipo', 'Marca_Modelo', 'N_Serie', 'Sistema_Operativo', 'Version_SO', 'Procesador',
            'Memoria_RAM', 'Part_Number', 'Hostname', 'Direccion_IP', 'Direccion_MAC', 'Tipo_Conexion', 
            'Punto_Red', 'Estado', 'Disco_Almacenamiento', 'Impresora_Instalada',
            'Switch_Puertos', 'Switch_Gestionable', 'Impresora_PuertoIP', 'TV_Pulgadas', 'TV_Resolucion',
            'Incluye_Celular', 'Celular_MarcaModelo', 'Celular_IMEI', 'Celular_Telefono', 
            'Celular_RAM', 'Celular_Almacenamiento', 'Celular_Procesador', 'Celular_SO', 
            'Celular_Accesorios', 'Celular_EstadoFisico',
            'Observaciones', 'Firma_Tecnico', 'Firma_Digital_Tecnico', 'Firma_Empleado', 'Firma_Digital_Empleado', 'Foto_Equipo',
            # Campos de Baja estilo SUNAT
            'Estado_Ficha', 'Motivo_Baja'
        ]).to_csv(DB_CSV, index=False)
    
    if not os.path.exists(TECNICOS_CSV):
        pd.DataFrame(columns=['Nombre', 'Especialidad', 'Firma_Digital']).to_csv(TECNICOS_CSV, index=False)
        
    if not os.path.exists(EMPLEADOS_CSV):
        pd.DataFrame(columns=['ID_Empleado', 'Nombre', 'Area', 'Cargo', 'Correo']).to_csv(EMPLEADOS_CSV, index=False)

inicializar_archivos()

def generar_codigo_ficha():
    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        siguiente_num = len(df) + 1
    else:
        siguiente_num = 1
    anio_actual = datetime.now().strftime('%Y')
    return f"INV-{anio_actual}-{siguiente_num:03d}"

@app.route('/')
def index():
    registros = []
    total_equipos = total_local = total_provincia = 0
    tecnicos = []
    empleados = []
    siguiente_ficha = generar_codigo_ficha()

    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV).fillna('')
        registros = df.to_dict(orient='records')
        total_equipos = len(df[df['Estado_Ficha'] != 'ANULADO'])
        total_local = len(df[(df['Sede'].astype(str).str.contains('Lima|Central|San Isidro', case=False, na=False)) & (df['Estado_Ficha'] != 'ANULADO')])
        total_provincia = total_equipos - total_local

    if os.path.exists(TECNICOS_CSV):
        tecnicos = pd.read_csv(TECNICOS_CSV).fillna('').to_dict(orient='records')

    if os.path.exists(EMPLEADOS_CSV):
        empleados = pd.read_csv(EMPLEADOS_CSV).fillna('').to_dict(orient='records')

    return render_template('index.html', 
                           registros=registros, 
                           total_equipos=total_equipos, 
                           total_local=total_local, 
                           total_provincia=total_provincia,
                           tecnicos=tecnicos,
                           empleados=empleados,
                           siguiente_ficha=siguiente_ficha)

@app.route('/add_manual', methods=['POST'])
def add_manual():
    n_ficha = generar_codigo_ficha()
    fecha = request.form.get('fecha')
    sede = request.form.get('sede')
    id_trabajador = request.form.get('id_trabajador', '')
    responsable = request.form.get('responsable', '')
    area = request.form.get('area', '')
    cargo = request.form.get('cargo', '')
    correo = request.form.get('correo', '')
    
    tipo_equipo = request.form.get('tipo_equipo', 'Laptop')
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
    
    disco_almacenamiento = request.form.get('disco_almacenamiento', '')
    impresora_instalada = request.form.get('impresora_instalada', '')
    
    switch_puertos = request.form.get('switch_puertos', '')
    switch_gestionable = request.form.get('switch_gestionable', '')
    impresora_puertoip = request.form.get('impresora_puertoip', '')
    tv_pulgadas = request.form.get('tv_pulgadas', '')
    tv_resolucion = request.form.get('tv_resolucion', '')

    incluye_celular = request.form.get('incluye_celular') == 'on'
    celular_marcamodelo = request.form.get('celular_marcamodelo', '') if incluye_celular else ''
    celular_imei = request.form.get('celular_imei', '') if incluye_celular else ''
    celular_telefono = request.form.get('celular_telefono', '') if incluye_celular else ''
    celular_ram = request.form.get('celular_ram', '') if incluye_celular else ''
    celular_almacenamiento = request.form.get('celular_almacenamiento', '') if incluye_celular else ''
    celular_procesador = request.form.get('celular_procesador', '') if incluye_celular else ''
    celular_so = request.form.get('celular_so', '') if incluye_celular else ''
    celular_accesorios = request.form.get('celular_accesorios', '') if incluye_celular else ''
    celular_estadofisico = request.form.get('celular_estadofisico', '') if incluye_celular else ''
    
    observaciones = request.form.get('observaciones', '')
    firma_tecnico = request.form.get('firma_tecnico', '')
    firma_digital_empleado = request.form.get('firma_digital_empleado', '')
    id_unico = str(uuid.uuid4())[:8]

    # Resolver firma digital del técnico (Buscando por nombre o por imagen subida en el registro de técnicos)
    firma_digital_tecnico = ""
    if os.path.exists(TECNICOS_CSV):
        df_tec = pd.read_csv(TECNICOS_CSV).fillna('')
        match_tec = df_tec[df_tec['Nombre'] == firma_tecnico]
        if not match_tec.empty:
            firma_digital_tecnico = match_tec.iloc[0]['Firma_Digital']

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
        'ID_Unico': id_unico, 'N_Ficha': n_ficha, 'Fecha': fecha, 'Sede': sede,
        'ID_Trabajador': id_trabajador, 'Responsable': responsable, 'Area': area, 'Cargo': cargo, 'Correo': correo,
        'Tipo_Equipo': tipo_equipo, 'Marca_Modelo': marca_modelo, 'N_Serie': n_serie,
        'Sistema_Operativo': sistema_operativo, 'Version_SO': version_so, 'Procesador': procesador,
        'Memoria_RAM': memoria_ram, 'Part_Number': part_number, 'Hostname': hostname,
        'Direccion_IP': direccion_ip, 'Direccion_MAC': direccion_mac, 'Tipo_Conexion': tipo_conexion,
        'Punto_Red': punto_red, 'Estado': estado, 'Disco_Almacenamiento': disco_almacenamiento,
        'Impresora_Instalada': impresora_instalada,
        'Switch_Puertos': switch_puertos, 'Switch_Gestionable': switch_gestionable,
        'Impresora_PuertoIP': impresora_puertoip, 'TV_Pulgadas': tv_pulgadas, 'TV_Resolucion': tv_resolucion,
        'Incluye_Celular': 'Sí' if incluye_celular else 'No',
        'Celular_MarcaModelo': celular_marcamodelo, 'Celular_IMEI': celular_imei, 'Celular_Telefono': celular_telefono,
        'Celular_RAM': celular_ram, 'Celular_Almacenamiento': celular_almacenamiento, 'Celular_Procesador': celular_procesador,
        'Celular_SO': celular_so, 'Celular_Accesorios': celular_accesorios, 'Celular_EstadoFisico': celular_estadofisico,
        'Observaciones': observaciones, 'Firma_Tecnico': firma_tecnico, 'Firma_Digital_Tecnico': firma_digital_tecnico,
        'Firma_Empleado': responsable, 'Firma_Digital_Empleado': firma_digital_empleado, 'Foto_Equipo': foto_equipo_str,
        'Estado_Ficha': 'ACTIVO', 'Motivo_Baja': ''
    }

    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    else:
        df = pd.DataFrame([nuevo_registro])
    
    df.to_csv(DB_CSV, index=False)
    return redirect(url_for('index'))

@app.route('/dar_baja/<id_unico>', methods=['POST'])
def dar_baja(id_unico):
    motivo = request.form.get('motivo_baja', 'Sin motivo especificado')
    if os.path.exists(DB_CSV):
        df = pd.read_csv(DB_CSV)
        # Actualiza el estado respetando el código de la ficha original
        df.loc[df['ID_Unico'].astype(str) == str(id_unico), 'Estado_Ficha'] = 'ANULADO'
        df.loc[df['ID_Unico'].astype(str) == str(id_unico), 'Motivo_Baja'] = motivo
        df.to_csv(DB_CSV, index=False)
    return redirect(url_for('index'))

@app.route('/add_personal', methods=['POST'])
def add_personal():
    tipo = request.form.get('tipo_personal')
    if tipo == 'tecnico':
        nombre_tec = request.form.get('nombre_tecnico')
        especialidad = request.form.get('especialidad_tecnico', '')
        firma_digital_tec = request.form.get('firma_digital_tecnico_reg', '')

        # O si subió un archivo PNG para la firma del técnico
        archivo_firma = request.files.get('archivo_firma_tecnico')
        if archivo_firma and archivo_firma.filename != '':
            try:
                upload_res = cloudinary.uploader.upload(archivo_firma, folder="Yamaha_Firmas_Tecnicos")
                firma_digital_tec = upload_res.get('secure_url')
            except Exception as e:
                print(f"Error subiendo firma PNG: {e}")

        nuevo_dato = {
            'Nombre': nombre_tec,
            'Especialidad': especialidad,
            'Firma_Digital': firma_digital_tec
        }
        file_path = TECNICOS_CSV
    else:
        nuevo_dato = {
            'ID_Empleado': request.form.get('id_empleado_reg'),
            'Nombre': request.form.get('nombre_empleado_reg'),
            'Area': request.form.get('area_empleado_reg'),
            'Cargo': request.form.get('cargo_empleado_reg'),
            'Correo': request.form.get('correo_empleado_reg', '')
        }
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