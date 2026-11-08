import os
import re
import datetime
import pandas as pd
import pypdf
from flask import Flask, render_template_string, request, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'yamaha_inventario_personal_separado'

UPLOAD_FOLDER = 'reportes_pdf'
OUTPUT_EXCEL = 'Inventario_Corporativo_Yamaha_Final.xlsx'
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'webp'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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

def guardar_tecnicos(df):
    df['DNI'] = df['DNI'].astype(str)
    df.to_csv(DB_TECNICOS, index=False)

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

def guardar_empleados(df):
    df['DNI'] = df['DNI'].astype(str)
    df.to_csv(DB_EMPLEADOS, index=False)

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

def obtener_etiquetas_atributos(tipo):
    if tipo in ['Laptop', 'PC de Escritorio']:
        return "Procesador", "Memoria RAM", "Almacenamiento / Hostname"
    elif tipo == 'Impresora':
        return "Tipo de Tóner / Tinta", "Conectividad (USB/Red)", "ShareName / Driver"
    elif tipo == 'Switch / Red':
        return "N° de Puertos", "Capacidad (L2/L3)", "VLAN / Configuración"
    elif tipo == 'Celular':
        return "N° IMEI", "Línea / Operador", "Almacenamiento"
    elif tipo == 'Televisor (TV)':
        return "Pulgadas de Pantalla", "Resolución", "Tecnología / Smart TV"
    else:
        return "Atributo 1", "Atributo 2", "Atributo 3"

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

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Inventario Yamaha - Gestión de Personal y Activos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        body { background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .navbar-custom { background-color: #1F4E78; }
        .card-custom { border-radius: 10px; border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        .badge-local { background-color: #0d6efd; }
        .badge-provincia { background-color: #fd7e14; }
        .img-thumb { width: 35px; height: 35px; object-fit: cover; border-radius: 4px; border: 1px solid #ddd; margin-right: 2px; }
    </style>
    <script>
        const tecnicosDB = {{ tecnicos_json | safe }};
        const empleadosDB = {{ empleados_json | safe }};

        function seleccionarTecnico(selectEl, nombreInputId) {
            let dni = selectEl.value;
            if (tecnicosDB[dni]) {
                document.getElementById(nombreInputId).value = tecnicosDB[dni].Nombre;
            } else {
                document.getElementById(nombreInputId).value = "";
            }
        }

        function seleccionarEmpleado(selectEl, nombreInputId, cargoInputId, firmaInputId) {
            let dni = selectEl.value;
            if (empleadosDB[dni]) {
                let emp = empleadosDB[dni];
                document.getElementById(nombreInputId).value = emp.Nombre;
                document.getElementById(cargoInputId).value = emp.Cargo;
                if(firmaInputId) document.getElementById(firmaInputId).value = dni;
            } else {
                document.getElementById(nombreInputId).value = "";
                document.getElementById(cargoInputId).value = "";
            }
        }

        function actualizarCampos(tipoSelect, prefix) {
            let tipo = tipoSelect.value;
            let lbl1 = document.getElementById(prefix + '_lbl1');
            let lbl2 = document.getElementById(prefix + '_lbl2');
            let lbl3 = document.getElementById(prefix + '_lbl3');

            if (tipo === 'Laptop' || tipo === 'PC de Escritorio') {
                lbl1.innerText = "Procesador";
                lbl2.innerText = "Memoria RAM";
                lbl3.innerText = "Almacenamiento / Hostname";
            } else if (tipo === 'Impresora') {
                lbl1.innerText = "Tipo de Tóner / Tinta";
                lbl2.innerText = "Conectividad (USB/Red)";
                lbl3.innerText = "ShareName / Driver";
            } else if (tipo === 'Switch / Red') {
                lbl1.innerText = "N° de Puertos";
                lbl2.innerText = "Capacidad (L2/L3)";
                lbl3.innerText = "VLAN / Configuración";
            } else if (tipo === 'Celular') {
                lbl1.innerText = "N° IMEI";
                lbl2.innerText = "Línea / Operador";
                lbl3.innerText = "Almacenamiento";
            } else if (tipo === 'Televisor (TV)') {
                lbl1.innerText = "Pulgadas de Pantalla";
                lbl2.innerText = "Resolución";
                lbl3.innerText = "Tecnología / Smart TV";
            }
        }
    </script>
</head>
<body>
    <nav class="navbar navbar-dark navbar-custom mb-4">
        <div class="container d-flex justify-content-between">
            <span class="navbar-brand mb-0 h1"><i class="bi bi-shield-check me-2"></i> YAMAHA - SISTEMA INTEGRADO DE INVENTARIO Y PERSONAL</span>
            <span class="text-white small">Gestión Local y Provincias</span>
        </div>
    </nav>

    <div class="container pb-5">
        <!-- Pestañas de Navegación -->
        <ul class="nav nav-tabs mb-4" id="inventarioTabs" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active fw-bold" id="activos-tab" data-bs-toggle="tab" data-bs-target="#activos-content" type="button" role="tab"><i class="bi bi-laptop me-1"></i> Inventario de Activos</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link fw-bold text-success" id="personal-tab" data-bs-toggle="tab" data-bs-target="#personal-content" type="button" role="tab"><i class="bi bi-people-fill me-1"></i> Módulo de Personal (Técnicos y Empleados)</button>
            </li>
        </ul>

        <div class="tab-content" id="inventarioTabContent">
            <!-- CONTENIDO DE ACTIVOS -->
            <div class="tab-pane fade show active" id="activos-content" role="tabpanel">
                <div class="row mb-4">
                    <div class="col-md-6 mb-3 mb-md-0">
                        <div class="card card-custom p-4 h-100">
                            <h5 class="fw-bold mb-2"><i class="bi bi-cloud-upload text-primary me-2"></i>Cargar Fichas PDF</h5>
                            <p class="text-muted small">Sube archivos PDF para generar código correlativo automático (ej: <code>INV-2026-0001</code>).</p>
                            <form action="/upload" method="post" enctype="multipart/form-data" class="row g-2 mt-auto">
                                <div class="col-12">
                                    <input type="file" name="files" class="form-control form-control-sm" multiple accept=".pdf" required>
                                </div>
                                <div class="col-12 text-end">
                                    <button type="submit" class="btn btn-primary btn-sm"><i class="bi bi-upload me-1"></i> Subir PDFs</button>
                                </div>
                            </form>
                        </div>
                    </div>

                    <div class="col-md-6">
                        <div class="card card-custom p-4 h-100 bg-light border border-primary-subtle">
                            <h5 class="fw-bold mb-2"><i class="bi bi-plus-circle-fill text-success me-2"></i>Registro Manual de Activos</h5>
                            <p class="text-muted small">Selecciona personal previamente registrado por DNI para asociarlo al activo.</p>
                            <div class="mt-auto text-end">
                                <button class="btn btn-success btn-sm" data-bs-toggle="modal" data-bs-target="#modalNuevoActivo">
                                    <i class="bi bi-plus-lg me-1"></i> Nuevo Activo Manual
                                </button>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="row mb-4">
                    <div class="col-md-4">
                        <div class="card card-custom p-3 bg-white border-start border-primary border-4">
                            <span class="text-muted small fw-bold">TOTAL ACTIVOS</span>
                            <h3 class="fw-bold text-primary mb-0">{{ total_equipos }}</h3>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card card-custom p-3 bg-white border-start border-info border-4">
                            <span class="text-muted small fw-bold">SEDE LOCAL (LIMA)</span>
                            <h3 class="fw-bold text-info mb-0">{{ total_local }}</h3>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card card-custom p-3 bg-white border-start border-warning border-4">
                            <span class="text-muted small fw-bold">PROVINCIAS</span>
                            <h3 class="fw-bold text-warning mb-0">{{ total_provincia }}</h3>
                        </div>
                    </div>
                </div>

                <div class="card card-custom p-4">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h5 class="fw-bold mb-0"><i class="bi bi-table me-2"></i>Consolidado de Activos Vigentes</h5>
                        {% if total_equipos > 0 %}
                        <a href="/download_excel" class="btn btn-success btn-sm"><i class="bi bi-file-earmark-excel me-1"></i> Descargar Excel Oficial</a>
                        {% endif %}
                    </div>
                    <div class="table-responsive">
                        <table class="table table-hover align-middle small">
                            <thead class="table-light">
                                <tr>
                                    <th>Fotos</th>
                                    <th>N° Ficha (ID)</th>
                                    <th>Perfil / Tipo</th>
                                    <th>Marca / Modelo</th>
                                    <th>N° Serie</th>
                                    <th>Responsable & Personal</th>
                                    <th>Sede</th>
                                    <th>Acciones</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for item in registros %}
                                <tr>
                                    <td>
                                        {% if item.Foto_Equipo and item.Foto_Equipo != 'Sin foto' %}
                                            {% set fotos = item.Foto_Equipo.split(',') %}
                                            {% for f in fotos %}
                                                <a href="/uploads/{{ f.strip() }}" target="_blank">
                                                    <img src="/uploads/{{ f.strip() }}" class="img-thumb" title="Ver foto">
                                                </a>
                                            {% endfor %}
                                        {% else %}
                                        <span class="badge bg-light text-muted border">Sin fotos</span>
                                        {% endif %}
                                    </td>
                                    <td><code class="text-primary fw-bold">{{ item.N_Ficha }}</code></td>
                                    <td><span class="badge bg-secondary">{{ item.Tipo_Equipo }}</span></td>
                                    <td><strong>{{ item.Modelo }}</strong></td>
                                    <td><code>{{ item.N_Serie }}</code></td>
                                    <td>
                                        <strong>{{ item.Responsable }}</strong><br>
                                        <small class="text-muted">Cargo: {{ item.Cargo }}</small><br>
                                        <span class="badge bg-light text-dark border mt-1"><i class="bi bi-shield-lock text-primary"></i> TI: {{ item.Firma_Tecnico }}</span>
                                    </td>
                                    <td>{{ item.Sede }}<br>
                                        {% if item.Clasificacion == 'Local' %}
                                        <span class="badge badge-local">Local</span>
                                        {% else %}
                                        <span class="badge badge-provincia">Provincia</span>
                                        {% endif %}
                                    </td>
                                    <td>
                                        <div class="btn-group-vertical gap-1" role="group">
                                            <button class="btn btn-sm btn-outline-primary" data-bs-toggle="modal" data-bs-target="#editModal{{ item.ID_Unico }}">
                                                <i class="bi bi-pencil-square"></i> Editar
                                            </button>
                                            <a href="/imprimir/{{ item.ID_Unico }}" target="_blank" class="btn btn-sm btn-outline-secondary">
                                                <i class="bi bi-printer"></i> Acta & Ficha
                                            </a>
                                            <a href="/delete/{{ item.ID_Unico }}" class="btn btn-sm btn-outline-danger" onclick="return confirm('¿Estás seguro de dar de baja este activo?');">
                                                <i class="bi bi-trash"></i> Dar Baja
                                            </a>
                                        </div>
                                    </td>
                                </tr>

                                <!-- Modal de Edición -->
                                <div class="modal fade" id="editModal{{ item.ID_Unico }}" tabindex="-1">
                                    <div class="modal-dialog modal-lg">
                                        <div class="modal-content">
                                            <form action="/update" method="POST" enctype="multipart/form-data">
                                                <div class="modal-header bg-primary text-white">
                                                    <h5 class="modal-title">Editar Activo y Personal: {{ item.N_Ficha }}</h5>
                                                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                                                </div>
                                                <div class="modal-body row g-3">
                                                    <input type="hidden" name="id_unico" value="{{ item.ID_Unico }}">
                                                    
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">N° Ficha (ID)</label>
                                                        <input type="text" class="form-control bg-light" name="n_ficha" value="{{ item.N_Ficha }}" readonly>
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">Perfil / Tipo</label>
                                                        <select class="form-select" name="tipo_equipo" id="tipo_{{ item.ID_Unico }}" onchange="actualizarCampos(this, 'edit_{{ item.ID_Unico }}')">
                                                            <option value="Laptop" {% if item.Tipo_Equipo == 'Laptop' %}selected{% endif %}>Laptop</option>
                                                            <option value="PC de Escritorio" {% if item.Tipo_Equipo == 'PC de Escritorio' %}selected{% endif %}>PC de Escritorio</option>
                                                            <option value="Impresora" {% if item.Tipo_Equipo == 'Impresora' %}selected{% endif %}>Impresora</option>
                                                            <option value="Switch / Red" {% if item.Tipo_Equipo == 'Switch / Red' %}selected{% endif %}>Switch / Red</option>
                                                            <option value="Celular" {% if item.Tipo_Equipo == 'Celular' %}selected{% endif %}>Celular</option>
                                                            <option value="Televisor (TV)" {% if item.Tipo_Equipo == 'Televisor (TV)' %}selected{% endif %}>Televisor (TV)</option>
                                                        </select>
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">Marca / Modelo</label>
                                                        <input type="text" class="form-control" name="modelo" value="{{ item.Modelo }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">N° Serie</label>
                                                        <input type="text" class="form-control" name="n_serie" value="{{ item.N_Serie }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">Sede / Ubicación</label>
                                                        <input type="text" class="form-control" name="sede" value="{{ item.Sede }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">Dirección IP</label>
                                                        <input type="text" class="form-control" name="ip" value="{{ item.IP }}">
                                                    </div>

                                                    <hr class="mt-2">
                                                    <h6 class="text-primary fw-bold"><i class="bi bi-person-badge"></i> Seleccionar Técnico Emisor (por DNI)</h6>
                                                    <div class="col-md-6">
                                                        <label class="form-label small fw-bold">Seleccionar Técnico Registrado</label>
                                                        <select class="form-select form-select-sm" onchange="seleccionarTecnico(this, 'firma_tec_{{ item.ID_Unico }}')">
                                                            <option value="">-- Seleccionar Técnico --</option>
                                                            {% for dni, tec in tecnicos.items() %}
                                                            <option value="{{ dni }}">{{ dni }} - {{ tec.Nombre }}</option>
                                                            {% endfor %}
                                                        </select>
                                                    </div>
                                                    <div class="col-md-6">
                                                        <label class="form-label small fw-bold">Técnico Emisor (TI)</label>
                                                        <input type="text" id="firma_tec_{{ item.ID_Unico }}" class="form-control form-control-sm" name="firma_tecnico" value="{{ item.Firma_Tecnico }}">
                                                    </div>

                                                    <hr class="mt-2">
                                                    <h6 class="text-primary fw-bold"><i class="bi bi-person-check"></i> Seleccionar Empleado / Responsable (por DNI)</h6>
                                                    <div class="col-md-4">
                                                        <label class="form-label small fw-bold">Seleccionar Empleado</label>
                                                        <select class="form-select form-select-sm" onchange="seleccionarEmpleado(this, 'resp_{{ item.ID_Unico }}', 'cargo_{{ item.ID_Unico }}', 'firma_emp_{{ item.ID_Unico }}')">
                                                            <option value="">-- Seleccionar Empleado --</option>
                                                            {% for dni, emp in empleados.items() %}
                                                            <option value="{{ dni }}">{{ dni }} - {{ emp.Nombre }}</option>
                                                            {% endfor %}
                                                        </select>
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label small fw-bold">Nombre Completo</label>
                                                        <input type="text" id="resp_{{ item.ID_Unico }}" class="form-control form-control-sm" name="responsable" value="{{ item.Responsable }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label small fw-bold">Cargo</label>
                                                        <input type="text" id="cargo_{{ item.ID_Unico }}" class="form-control form-control-sm" name="cargo" value="{{ item.Cargo }}">
                                                    </div>
                                                    <input type="hidden" id="firma_emp_{{ item.ID_Unico }}" name="firma_empleado" value="{{ item.Firma_Empleado }}">

                                                    <hr class="mt-2">
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold">Subir Fotos (Máx. 3)</label>
                                                        <input type="file" name="fotos_equipo" class="form-control form-control-sm" accept=".png, .jpg, .jpeg, .webp" multiple>
                                                    </div>

                                                    <hr class="mt-2">
                                                    <h6 class="text-primary fw-bold">Características Específicas</h6>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold" id="edit_{{ item.ID_Unico }}_lbl1">Atributo 1</label>
                                                        <input type="text" class="form-control" name="atributo_1" value="{{ item.Atributo_1 }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold" id="edit_{{ item.ID_Unico }}_lbl2">Atributo 2</label>
                                                        <input type="text" class="form-control" name="atributo_2" value="{{ item.Atributo_2 }}">
                                                    </div>
                                                    <div class="col-md-4">
                                                        <label class="form-label fw-bold" id="edit_{{ item.ID_Unico }}_lbl3">Atributo 3</label>
                                                        <input type="text" class="form-control" name="atributo_3" value="{{ item.Atributo_3 }}">
                                                    </div>
                                                </div>
                                                <div class="modal-footer">
                                                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                                                    <button type="submit" class="btn btn-primary">Guardar Cambios</button>
                                                </div>
                                            </form>
                                        </div>
                                    </div>
                                </div>
                                <script>
                                    window.addEventListener('DOMContentLoaded', () => {
                                        let sel = document.getElementById('tipo_{{ item.ID_Unico }}');
                                        if(sel) actualizarCampos(sel, 'edit_{{ item.ID_Unico }}');
                                    });
                                </script>
                                {% else %}
                                <tr>
                                    <td colspan="8" class="text-center text-muted py-4">No hay activos registrados.</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- CONTENIDO DE GESTIÓN DE PERSONAL (SECCIÓN APARTE) -->
            <div class="tab-pane fade" id="personal-content" role="tabpanel">
                <div class="row">
                    <!-- Columna Técnicos -->
                    <div class="col-md-6 mb-4">
                        <div class="card card-custom p-4 h-100 border border-success-subtle">
                            <h5 class="fw-bold text-success mb-3"><i class="bi bi-shield-shaded me-2"></i>Catálogo de Técnicos (Soporte TI)</h5>
                            
                            <form action="/add_tecnico" method="POST" class="row g-2 mb-4 p-3 bg-light rounded">
                                <div class="col-12"><h6 class="small fw-bold text-muted">Registrar Nuevo Técnico</h6></div>
                                <div class="col-md-4">
                                    <input type="text" class="form-control form-control-sm" name="dni" placeholder="DNI (8 dígitos)" maxlength="8" required>
                                </div>
                                <div class="col-md-8">
                                    <input type="text" class="form-control form-control-sm" name="nombre" placeholder="Nombre completo" required>
                                </div>
                                <div class="col-12">
                                    <input type="text" class="form-control form-control-sm" name="cargo" placeholder="Cargo / Especialidad" required>
                                </div>
                                <div class="col-12 text-end">
                                    <button type="submit" class="btn btn-success btn-sm"><i class="bi bi-plus-circle me-1"></i> Guardar Técnico</button>
                                </div>
                            </form>

                            <div class="table-responsive">
                                <table class="table table-sm table-striped align-middle small">
                                    <thead class="table-dark">
                                        <tr>
                                            <th>DNI</th>
                                            <th>Nombre</th>
                                            <th>Cargo</th>
                                            <th>Acción</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {% for dni, tec in tecnicos.items() %}
                                        <tr>
                                            <td><code>{{ dni }}</code></td>
                                            <td><strong>{{ tec.Nombre }}</strong></td>
                                            <td>{{ tec.Cargo }}</td>
                                            <td>
                                                <a href="/delete_tecnico/{{ dni }}" class="btn btn-outline-danger btn-sm py-0 px-1" onclick="return confirm('¿Eliminar técnico?');"><i class="bi bi-trash"></i></a>
                                            </td>
                                        </tr>
                                        {% endfor %}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Columna Empleados -->
                    <div class="col-md-6 mb-4">
                        <div class="card card-custom p-4 h-100 border border-primary-subtle">
                            <h5 class="fw-bold text-primary mb-3"><i class="bi bi-people me-2"></i>Catálogo de Empleados (Responsables)</h5>
                            
                            <form action="/add_empleado" method="POST" class="row g-2 mb-4 p-3 bg-light rounded">
                                <div class="col-12"><h6 class="small fw-bold text-muted">Registrar Nuevo Empleado</h6></div>
                                <div class="col-md-4">
                                    <input type="text" class="form-control form-control-sm" name="dni" placeholder="DNI (8 dígitos)" maxlength="8" required>
                                </div>
                                <div class="col-md-8">
                                    <input type="text" class="form-control form-control-sm" name="nombre" placeholder="Nombre completo" required>
                                </div>
                                <div class="col-12">
                                    <input type="text" class="form-control form-control-sm" name="cargo" placeholder="Cargo / Área" required>
                                </div>
                                <div class="col-12 text-end">
                                    <button type="submit" class="btn btn-primary btn-sm"><i class="bi bi-plus-circle me-1"></i> Guardar Empleado</button>
                                </div>
                            </form>

                            <div class="table-responsive">
                                <table class="table table-sm table-striped align-middle small">
                                    <thead class="table-dark">
                                        <tr>
                                            <th>DNI</th>
                                            <th>Nombre</th>
                                            <th>Cargo</th>
                                            <th>Acción</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {% for dni, emp in empleados.items() %}
                                        <tr>
                                            <td><code>{{ dni }}</code></td>
                                            <td><strong>{{ emp.Nombre }}</strong></td>
                                            <td>{{ emp.Cargo }}</td>
                                            <td>
                                                <a href="/delete_empleado/{{ dni }}" class="btn btn-outline-danger btn-sm py-0 px-1" onclick="return confirm('¿Eliminar empleado?');"><i class="bi bi-trash"></i></a>
                                            </td>
                                        </tr>
                                        {% endfor %}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal Nuevo Activo -->
    <div class="modal fade" id="modalNuevoActivo" tabindex="-1">
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <form action="/add_manual" method="POST" enctype="multipart/form-data">
                    <div class="modal-header bg-success text-white">
                        <h5 class="modal-title">Registrar Nuevo Activo (Selección de Personal por DNI)</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body row g-3">
                        <div class="col-md-4">
                            <label class="form-label fw-bold">Perfil / Tipo</label>
                            <select class="form-select" name="tipo_equipo" id="nuevo_tipo" onchange="actualizarCampos(this, 'nuevo')">
                                <option value="Laptop" selected>Laptop</option>
                                <option value="PC de Escritorio">PC de Escritorio</option>
                                <option value="Impresora">Impresora</option>
                                <option value="Switch / Red">Switch / Red</option>
                                <option value="Celular">Celular</option>
                                <option value="Televisor (TV)">Televisor (TV)</option>
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold">Marca / Modelo</label>
                            <input type="text" class="form-control" name="modelo" placeholder="Modelo del equipo" required>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold">N° Serie</label>
                            <input type="text" class="form-control" name="n_serie" placeholder="N° de Serie" required>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold">Sede / Ubicación</label>
                            <input type="text" class="form-control" name="sede" value="Sede Central - San Isidro">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold">Dirección IP</label>
                            <input type="text" class="form-control" name="ip" placeholder="192.168.1.X">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold">Fotos (Máx. 3)</label>
                            <input type="file" name="fotos_equipo" class="form-control form-control-sm" accept=".png, .jpg, .jpeg, .webp" multiple>
                        </div>

                        <hr class="mt-2">
                        <h6 class="text-success fw-bold"><i class="bi bi-person-badge"></i> Seleccionar Técnico Emisor</h6>
                        <div class="col-md-6">
                            <label class="form-label small fw-bold">Seleccionar Técnico</label>
                            <select class="form-select" onchange="seleccionarTecnico(this, 'firma_tec_nuevo')">
                                <option value="">-- Seleccionar Técnico --</option>
                                {% for dni, tec in tecnicos.items() %}
                                <option value="{{ dni }}">{{ dni }} - {{ tec.Nombre }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label small fw-bold">Técnico Emisor (TI)</label>
                            <input type="text" id="firma_tec_nuevo" class="form-control" name="firma_tecnico" value="Carlos Mendoza" required>
                        </div>

                        <hr class="mt-2">
                        <h6 class="text-success fw-bold"><i class="bi bi-person-check"></i> Seleccionar Empleado / Responsable</h6>
                        <div class="col-md-4">
                            <label class="form-label small fw-bold">Seleccionar Empleado</label>
                            <select class="form-select" onchange="seleccionarEmpleado(this, 'resp_nuevo', 'cargo_nuevo', 'firma_emp_nuevo')">
                                <option value="">-- Seleccionar Empleado --</option>
                                {% for dni, emp in empleados.items() %}
                                <option value="{{ dni }}">{{ dni }} - {{ emp.Nombre }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small fw-bold">Nombre Completo</label>
                            <input type="text" id="resp_nuevo" class="form-control" name="responsable" placeholder="Nombre completo" required>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small fw-bold">Cargo</label>
                            <input type="text" id="cargo_nuevo" class="form-control" name="cargo" placeholder="Cargo">
                        </div>
                        <input type="hidden" id="firma_emp_nuevo" name="firma_empleado" value="-">

                        <hr class="mt-3">
                        <h6 class="text-success fw-bold">Características Específicas</h6>
                        <div class="col-md-4">
                            <label class="form-label fw-bold" id="nuevo_lbl1">Procesador</label>
                            <input type="text" class="form-control" name="atributo_1">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold" id="nuevo_lbl2">Memoria RAM</label>
                            <input type="text" class="form-control" name="atributo_2">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label fw-bold" id="nuevo_lbl3">Almacenamiento</label>
                            <input type="text" class="form-control" name="atributo_3">
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancelar</button>
                        <button type="submit" class="btn btn-success">Guardar Activo</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        window.addEventListener('DOMContentLoaded', () => {
            let sel = document.getElementById('nuevo_tipo');
            if(sel) actualizarCampos(sel, 'nuevo');
        });
    </script>
</body>
</html>
"""

PRINT_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Acta de Entrega - Ficha {{ item.N_Ficha }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #fff; color: #000; font-family: Arial, sans-serif; }
        .ficha-box { border: 2px solid #1F4E78; padding: 30px; border-radius: 8px; margin-top: 30px; }
        .header-title { background-color: #1F4E78; color: white; padding: 10px 15px; border-radius: 5px; }
        .signature-line { border-top: 1px solid #000; margin-top: 50px; padding-top: 5px; text-align: center; }
        .table-specs td { padding: 8px 12px; }
        @media print {
            .no-print { display: none !important; }
            body { background: none; }
            .ficha-box { border: none; margin: 0; padding: 0; }
        }
    </style>
</head>
<body>
    <div class="container my-4">
        <div class="text-end no-print mb-3">
            <button onclick="window.print();" class="btn btn-primary btn-sm"><i class="bi bi-printer"></i> Imprimir o Guardar PDF</button>
            <a href="/" class="btn btn-secondary btn-sm">Volver al Inventario</a>
        </div>

        <div class="ficha-box">
            <div class="d-flex justify-content-between align-items-center header-title mb-4">
                <h4 class="mb-0 fw-bold">YAMAHA - ACTA DE ASIGNACIÓN DE ACTIVOS</h4>
                <h5 class="mb-0">Ficha: {{ item.N_Ficha }}</h5>
            </div>

            <p class="text-muted small">Fecha de emisión: {{ item.Fecha }} | Sede: {{ item.Sede }} ({{ item.Clasificacion }})</p>

            <h5 class="text-primary fw-bold mt-4 border-bottom pb-2">1. Datos del Colaborador (Responsable)</h5>
            <div class="row g-2 mb-3">
                <div class="col-md-6"><strong>Colaborador:</strong> {{ item.Responsable }}</div>
                <div class="col-md-6"><strong>Cargo:</strong> {{ item.Cargo }}</div>
            </div>

            <h5 class="text-primary fw-bold mt-4 border-bottom pb-2">2. Especificaciones del Activo</h5>
            <div class="row g-2 mb-3">
                <div class="col-md-4"><strong>Tipo de Componente:</strong> {{ item.Tipo_Equipo }}</div>
                <div class="col-md-4"><strong>Marca / Modelo:</strong> {{ item.Modelo }}</div>
                <div class="col-md-4"><strong>N° de Serie / ID:</strong> {{ item.N_Serie }}</div>
                <div class="col-md-4"><strong>Dirección IP:</strong> {{ item.IP }}</div>
                <div class="col-md-4"><strong>Estado Operativo:</strong> {{ item.Estado }}</div>
                <div class="col-md-4">
                    <strong>Evidencias Fotográficas:</strong> 
                    {% if item.Foto_Equipo and item.Foto_Equipo != 'Sin foto' %}
                        {% set fotos = item.Foto_Equipo.split(',') %}
                        {% for f in fotos %}
                            <a href="/uploads/{{ f.strip() }}" target="_blank" class="no-print text-decoration-none me-1 badge bg-primary">[ Ver Foto {{ loop.index }} ]</a>
                        {% endfor %}
                        <span class="d-none d-print-inline">Registradas ({{ item.Foto_Equipo }})</span>
                    {% else %}
                    <span class="text-muted">No adjuntas</span>
                    {% endif %}
                </div>
            </div>

            {% set lbl1, lbl2, lbl3 = obtener_etiquetas_atributos(item.Tipo_Equipo) %}
            <h5 class="text-primary fw-bold mt-4 border-bottom pb-2">3. Características Técnicas Específicas</h5>
            <table class="table table-bordered table-specs mt-2 small">
                <tbody>
                    <tr>
                        <td width="33%"><strong>{{ lbl1 }}:</strong> {{ item.Atributo_1 }}</td>
                        <td width="33%"><strong>{{ lbl2 }}:</strong> {{ item.Atributo_2 }}</td>
                        <td width="33%"><strong>{{ lbl3 }}:</strong> {{ item.Atributo_3 }}</td>
                    </tr>
                </tbody>
            </table>

            <div class="alert alert-secondary small mt-4">
                <strong>Términos de Recepción:</strong> El colaborador confirma haber recibido el equipo en óptimas condiciones operativas y se compromete a su cuidado y uso estrictamente laboral conforme a las normativas de Yamaha.
            </div>

            <div class="row mt-5 pt-4">
                <div class="col-md-6 text-center px-4">
                    <div class="signature-line">
                        <strong>{{ item.Firma_Tecnico }}</strong><br>
                        <span class="text-muted small">Técnico Emisor (Soporte TI)</span>
                    </div>
                </div>
                <div class="col-md-6 text-center px-4">
                    <div class="signature-line">
                        <strong>{{ item.Responsable }}</strong><br>
                        <span class="text-muted small">Colaborador / Receptor (Firma Digital / DNI)</span>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    procesar_nuevos_archivos(app.config['UPLOAD_FOLDER'])
    df = cargar_registros_base()
    
    if not df.empty and 'Estado' in df.columns:
        df_activos = df[df['Estado'] != 'Dado de baja']
    else:
        df_activos = df

    registros = df_activos.to_dict(orient='records') if not df_activos.empty else []
    total_equipos = len(registros)
    total_local = sum(1 for r in registros if r['Clasificacion'] == 'Local')
    total_provincia = sum(1 for r in registros if r['Clasificacion'] == 'Provincia')

    df_tec = cargar_tecnicos()
    tecnicos_dict = df_tec.set_index('DNI')[['Nombre', 'Cargo']].to_dict('index')

    df_emp = cargar_empleados()
    empleados_dict = df_emp.set_index('DNI')[['Nombre', 'Cargo']].to_dict('index')

    import json
    return render_template_string(
        HTML_TEMPLATE, 
        registros=registros, 
        total_equipos=total_equipos, 
        total_local=total_local, 
        total_provincia=total_provincia,
        tecnicos=tecnicos_dict,
        empleados=empleados_dict,
        tecnicos_json=json.dumps(tecnicos_dict),
        empleados_json=json.dumps(empleados_dict)
    )

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return redirect(request.url)
    files = request.files.getlist('files')
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    return redirect(url_for('index'))

@app.route('/add_tecnico', methods=['POST'])
def add_tecnico():
    dni = request.form.get('dni', '').strip()
    nombre = request.form.get('nombre', '').strip()
    cargo = request.form.get('cargo', '').strip()
    if dni and nombre:
        df = cargar_tecnicos()
        if dni in df['DNI'].values:
            df.loc[df['DNI'] == dni, ['Nombre', 'Cargo']] = [nombre, cargo]
        else:
            nuevo = pd.DataFrame([{"DNI": dni, "Nombre": nombre, "Cargo": cargo}])
            df = pd.concat([df, nuevo], ignore_index=True)
        guardar_tecnicos(df)
    return redirect(url_for('index'))

@app.route('/delete_tecnico/<dni>')
def delete_tecnico(dni):
    df = cargar_tecnicos()
    df = df[df['DNI'] != str(dni)]
    guardar_tecnicos(df)
    return redirect(url_for('index'))

@app.route('/add_empleado', methods=['POST'])
def add_empleado():
    dni = request.form.get('dni', '').strip()
    nombre = request.form.get('nombre', '').strip()
    cargo = request.form.get('cargo', '').strip()
    if dni and nombre:
        df = cargar_empleados()
        if dni in df['DNI'].values:
            df.loc[df['DNI'] == dni, ['Nombre', 'Cargo']] = [nombre, cargo]
        else:
            nuevo = pd.DataFrame([{"DNI": dni, "Nombre": nombre, "Cargo": cargo}])
            df = pd.concat([df, nuevo], ignore_index=True)
        guardar_empleados(df)
    return redirect(url_for('index'))

@app.route('/delete_empleado/<dni>')
def delete_empleado(dni):
    df = cargar_empleados()
    df = df[df['DNI'] != str(dni)]
    guardar_empleados(df)
    return redirect(url_for('index'))

@app.route('/add_manual', methods=['POST'])
def add_manual():
    df = cargar_registros_base()
    if not df.empty and 'ID_Unico' in df.columns:
        try:
            numeric_ids = pd.to_numeric(df['ID_Unico'], errors='coerce')
            next_id = str(int(numeric_ids.max() + 1)) if not numeric_ids.isna().all() else str(len(df) + 1)
        except:
            next_id = str(len(df) + 1)
    else:
        next_id = "1"
    
    sede_val = request.form.get('sede', 'Sede Central - San Isidro')
    tipo_ubicacion = "Local" if any(l in sede_val.lower() for l in ["lima", "san isidro", "local", "central"]) else "Provincia"
    nuevo_codigo_ficha = generar_siguiente_ficha()
    responsable_val = request.form.get('responsable', 'Sin Asignar')

    archivos_fotos = request.files.getlist('fotos_equipo')
    nombres_guardados = []
    
    for i, file_foto in enumerate(archivos_fotos[:3]):
        if file_foto and allowed_file(file_foto.filename):
            ext = file_foto.filename.rsplit('.', 1)[1].lower()
            foto_filename = f"{nuevo_codigo_ficha}_{i+1}.{ext}"
            file_foto.save(os.path.join(app.config['UPLOAD_FOLDER'], foto_filename))
            nombres_guardados.append(foto_filename)

    foto_equipo_str = ",".join(nombres_guardados) if nombres_guardados else "Sin foto"

    nuevo_reg = {
        "ID_Unico": next_id,
        "Archivo_PDF": "Registro Manual",
        "N_Ficha": nuevo_codigo_ficha,
        "Fecha": datetime.datetime.now().strftime('%Y-%m-%d'),
        "Sede": sede_val,
        "Clasificacion": tipo_ubicacion,
        "ID_Trabajador": "-",
        "Responsable": responsable_val,
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
        "Firma_Empleado": responsable_val,
        "Foto_Equipo": foto_equipo_str
    }

    df_nuevo = pd.DataFrame([nuevo_reg])
    df_combined = pd.concat([df, df_nuevo], ignore_index=True) if not df.empty else df_nuevo
    guardar_registros_base(df_combined)
    return redirect(url_for('index'))

@app.route('/update', methods=['POST'])
def update_registro():
    id_unico = str(request.form.get('id_unico'))
    df = cargar_registros_base()
    
    if not df.empty:
        df['ID_Unico'] = df['ID_Unico'].astype(str)
        if id_unico in df['ID_Unico'].values:
            idx = df.index[df['ID_Unico'] == id_unico][0]
            codigo_ficha = df.at[idx, 'N_Ficha']
            
            df.at[idx, 'Tipo_Equipo'] = request.form.get('tipo_equipo')
            df.at[idx, 'Modelo'] = request.form.get('modelo')
            df.at[idx, 'N_Serie'] = request.form.get('n_serie')
            df.at[idx, 'Responsable'] = request.form.get('responsable')
            df.at[idx, 'Cargo'] = request.form.get('cargo')
            
            nueva_sede = request.form.get('sede')
            df.at[idx, 'Sede'] = nueva_sede
            df.at[idx, 'Clasificacion'] = "Local" if any(l in nueva_sede.lower() for l in ["lima", "san isidro", "local", "central"]) else "Provincia"
            
            df.at[idx, 'IP'] = request.form.get('ip')
            df.at[idx, 'Atributo_1'] = request.form.get('atributo_1')
            df.at[idx, 'Atributo_2'] = request.form.get('atributo_2')
            df.at[idx, 'Atributo_3'] = request.form.get('atributo_3')
            df.at[idx, 'Firma_Tecnico'] = request.form.get('firma_tecnico', 'Carlos Mendoza')
            df.at[idx, 'Firma_Empleado'] = request.form.get('firma_empleado', '-')

            archivos_fotos = request.files.getlist('fotos_equipo')
            if archivos_fotos and archivos_fotos[0].filename != '':
                nombres_guardados = []
                for i, file_foto in enumerate(archivos_fotos[:3]):
                    if file_foto and allowed_file(file_foto.filename):
                        ext = file_foto.filename.rsplit('.', 1)[1].lower()
                        foto_filename = f"{codigo_ficha}_{i+1}.{ext}"
                        file_foto.save(os.path.join(app.config['UPLOAD_FOLDER'], foto_filename))
                        nombres_guardados.append(foto_filename)
                if nombres_guardados:
                    df.at[idx, 'Foto_Equipo'] = ",".join(nombres_guardados)
            
            guardar_registros_base(df)
        
    return redirect(url_for('index'))

@app.route('/delete/<id_unico>')
def delete_registro(id_unico):
    id_unico_str = str(id_unico).strip()
    df = cargar_registros_base()
    if not df.empty:
        df['ID_Unico'] = df['ID_Unico'].astype(str).str.strip()
        if id_unico_str in df['ID_Unico'].values:
            idx = df.index[df['ID_Unico'] == id_unico_str][0]
            df.at[idx, 'Estado'] = 'Dado de baja'
            guardar_registros_base(df)
    return redirect(url_for('index'))

@app.route('/imprimir/<id_unico>')
def imprimir_ficha(id_unico):
    df = cargar_registros_base()
    if not df.empty:
        df['ID_Unico'] = df['ID_Unico'].astype(str)
        match = df[df['ID_Unico'] == str(id_unico)]
        if not match.empty:
            item = match.iloc[0].to_dict()
            return render_template_string(PRINT_TEMPLATE, item=item, obtener_etiquetas_atributos=obtener_etiquetas_atributos)
    return "Ficha no encontrada", 404

@app.route('/download_excel')
def download_excel():
    if os.path.exists(OUTPUT_EXCEL):
        return send_from_directory('.', OUTPUT_EXCEL, as_attachment=True)
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("Iniciando servidor Yamaha en http://127.0.0.1:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)