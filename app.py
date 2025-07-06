import os
import csv
import uuid
import msal
from io import BytesIO
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, abort, send_file
)
from docxtpl import DocxTemplate

# ─── Configuración ────────────────────────────────────────────────────────────

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

# Azure AD (variables de entorno definidas en Render)
CLIENT_ID     = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
TENANT_ID     = os.environ.get('TENANT_ID')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.environ.get('REDIRECT_URI')
SCOPE         = ['User.Read']

# Rutas
BASE_PATH      = os.path.dirname(__file__)
CSV_DIR        = os.environ.get('CSV_DIR', os.path.join(BASE_PATH, 'csv'))
PLANTILLAS_DIR = os.path.join(BASE_PATH, 'plantillas')

# Asegurarnos de que existan
os.makedirs(CSV_DIR, exist_ok=True)

# Archivos CSV
USUARIOS_CSV   = os.path.join(CSV_DIR, 'usuarios.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos.csv')

# Plantillas .docx
TEMPLATE_SOLICITUD = os.path.join(PLANTILLAS_DIR, 'ITA-VI-SS-FO-01 Solicitud de Servicio Social.docx')
TEMPLATE_BIMESTRAL = os.path.join(PLANTILLAS_DIR, 'Reporte_Bimestral_Plantilla.docx')
TEMPLATE_FINAL     = os.path.join(PLANTILLAS_DIR, 'Reporte_Final_Plantilla.docx')

# Campos completos para alumnos (debe coincidir con el encabezado de tu alumnos.csv)
CAMPOS_ALUMNOS = [
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio','No_Control',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia','Titular_Dependencia',
    'Cargo_Responsable','Responsable_Proyecto','Nombre_Programa','Modalidad_Externa','Modalidad_Interna',
    'Fecha_Inicio','Fecha_Terminacion','Actividades',
    # TP_...
    'TP_Edu_Adultos','TP_Deportivo','TP_Civico','TP_Salud','TP_Otros','TP_Desarrollo','TP_Cultural',
    'TP_Sustentable','TP_Medio_Amb',
    # Fecha de solicitud
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Campos para bimestral
    'día1','mes1','año1','dia2','mes2','año2',
    'Reporte_No','Nombre_supervisor','Puesto_supervisor',
    'Actividad_1','Actividad_2','Actividad_3','Actividad_4',
    'Actividad_5','Actividad_6','Actividad_7','Actividad_8',
    # Campos para final
    'Actividad1','Logro1','Actividad2','Logro2','Actividad3','Logro3',
    'Actividad4','Logro4','Actividad5','Logro5','Actividad6','Logro6',
    'Actividad7','Logro7','Actividad8','Logro8',
    'Aprendizaje1','Beneficio1','Aprendizaje2','Beneficio2','Aprendizaje3','Beneficio3',
    'Aprendizaje4','Beneficio4','Aprendizaje5','Beneficio5','Aprendizaje6','Beneficio6',
    'Aprendizaje7','Beneficio7','Aprendizaje8','Beneficio8',
    # Y otros que quieras
]

# ─── Funciones utilitarias CSV ─────────────────────────────────────────────────

def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

# ─── Autenticación / Sesión ────────────────────────────────────────────────────

def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get('correo','').strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get('user')
    return u and u.get('rol') in roles

@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if 'error' in result:
        flash(result.get('error_description','Error autenticación'),'danger')
        return redirect(url_for('dashboard'))
    claims = result.get('id_token_claims', {})
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales permitidas", 403
    perfil = usuario_desde_csv(email) or {'correo':email,'rol':'Alumno','area':''}
    session['user'] = {
        'correo': perfil['correo'],
        'rol'   : perfil['rol'],
        'area'  : perfil.get('area',''),
        'name'  : claims.get('name','')
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login',_external=True)}"
    )

# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

# ─── CRUD Genérico para entidades ─────────────────────────────────────────────

def obtener_mapa(tipo):
    return {
        'usuarios':   (USUARIOS_CSV,   ['correo','rol','area']),
        'profesores': (PROFESORES_CSV, ['correo','nombre','area']),
        'areas':      (AREAS_CSV,      ['id','nombre','encargado']),
        'alumnos':    (ALUMNOS_CSV,    CAMPOS_ALUMNOS)
    }.get(tipo)

@app.route('/entidad/<tipo>')
def entidad_list(tipo):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'), 404
    path, campos = m
    registros = cargar_csv(path)
    return render_template('entidad_list.html', tipo=tipo,
                           campos=campos, registros=registros)

@app.route('/entidad/<tipo>/new', methods=['GET','POST'])
def entidad_new(tipo):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'),404
    path, campos = m
    if request.method=='POST':
        datos = {c: request.form.get(c,'').strip() for c in campos}
        # Generar nueva ID si corresponde
        if 'id' in campos and not datos.get('id'):
            datos['id'] = str(uuid.uuid4())
        registros = cargar_csv(path)
        registros.append(datos)
        guardar_csv(path, registros, campos)
        return redirect(url_for('entidad_list', tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo,
                           campos=campos, valores={})

@app.route('/entidad/<tipo>/edit/<pk>', methods=['GET','POST'])
def entidad_edit(tipo, pk):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'),404
    path, campos = m
    registros = cargar_csv(path)
    pk_field = campos[0]
    item = next((r for r in registros if r.get(pk_field)==pk), None)
    if not item:
        flash('Registro no encontrado','danger')
        return redirect(url_for('entidad_list', tipo=tipo))
    if request.method=='POST':
        for c in campos:
            item[c] = request.form.get(c,'').strip()
        guardar_csv(path, registros, campos)
        return redirect(url_for('entidad_list', tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo,
                           campos=campos, valores=item)

@app.route('/entidad/<tipo>/delete/<pk>', methods=['POST'])
def entidad_delete(tipo, pk):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'),404
    path, campos = m
    pk_field = campos[0]
    registros = [r for r in cargar_csv(path) if r.get(pk_field)!=pk]
    guardar_csv(path, registros, campos)
    return redirect(url_for('entidad_list', tipo=tipo))

# ─── Generación de documentos ─────────────────────────────────────────────────

def get_alumno(no_control):
    return next((a for a in cargar_csv(ALUMNOS_CSV)
                 if a.get('No_Control')==no_control), None)

def puede_acceder(alumno):
    rol    = session['user']['rol']
    correo = session['user']['correo']
    area   = session['user']['area']
    if rol=='Administrador':
        return True
    if rol=='Encargado' and alumno.get('area')==area:
        return True
    if rol=='Maestro' and alumno.get('profesor','').lower()==correo.lower():
        return True
    if rol=='Alumno' and alumno.get('Correo','').lower()==correo.lower():
        return True
    return False

def _render_docx(template_path, contexto, nombre_descarga):
    """Renderiza plantilla y envía DOCX en memoria."""
    doc = DocxTemplate(template_path)
    doc.render(contexto)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf,
                     download_name=nombre_descarga,
                     as_attachment=True)

@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    alumno = get_alumno(no_control)
    if not alumno or not puede_acceder(alumno):
        abort(403)
    contexto = {k: alumno.get(k,'') for k in alumno}
    return _render_docx(TEMPLATE_SOLICITUD,
                        contexto,
                        f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<bimestre>/<no_control>')
def generar_bimestral(bimestre, no_control):
    alumno = get_alumno(no_control)
    if not alumno or not puede_acceder(alumno):
        abort(403)
    # Campos comunes
    contexto = {
        'AP': alumno.get('Apellido_Paterno',''),
        'AM': alumno.get('Apellido_Materno',''),
        'Nombre': alumno.get('Nombre',''),
        'carrera': alumno.get('Carrera',''),
        'no_de_control': no_control,
        'día1': alumno.get('día1',''),
        'mes1': alumno.get('mes1',''),
        'año1': alumno.get('año1',''),
        'dia2': alumno.get('dia2',''),
        'mes2': alumno.get('mes2',''),
        'año2': alumno.get('año2',''),
        'Nombre_Programa': alumno.get('Nombre_Programa',''),
        'Reporte_No': bimestre,
        'Nombre_supervisor': alumno.get('Nombre_supervisor',''),
        'Puesto_supervisor': alumno.get('Puesto_supervisor',''),
    }
    for i in range(1,9):
        contexto[f'Actividad_{i}'] = alumno.get(f'Actividad_{i}','')
    return _render_docx(TEMPLATE_BIMESTRAL,
                        contexto,
                        f"Reporte_Bimestral_{no_control}_B{bimestre}.docx")

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    alumno = get_alumno(no_control)
    if not alumno or not puede_acceder(alumno):
        abort(403)
    contexto = {
        'Nombre': alumno.get('Nombre',''),
        'AP': alumno.get('Apellido_Paterno',''),
        'AM': alumno.get('Apellido_Materno',''),
        'no_de_control': no_control,
        'Carrera': alumno.get('Carrera',''),
        'Dependencia': alumno.get('Dependencia',''),
        'Nombre_Proyecto': alumno.get('Nombre_Programa',''),
        'Responsable_Proyecto': alumno.get('Responsable_Proyecto',''),
        'Municipio': alumno.get('Municipio',''),
        'Estado': alumno.get('Estado',''),
        'Fecha': alumno.get('Fecha_Terminacion',''),
        'Periodo': alumno.get('Periodo',''),
    }
    for i in range(1,9):
        contexto[f'Actividad{i}']   = alumno.get(f'Actividad{i}','')
        contexto[f'Logro{i}']       = alumno.get(f'Logro{i}','')
        contexto[f'Aprendizaje{i}'] = alumno.get(f'Aprendizaje{i}','')
        contexto[f'Beneficio{i}']   = alumno.get(f'Beneficio{i}','')
    return _render_docx(TEMPLATE_FINAL,
                        contexto,
                        f"Reporte_Final_{no_control}.docx")

# ─── Manejador 404 ─────────────────────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

# ─── Inicio ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',5000)), debug=True)
