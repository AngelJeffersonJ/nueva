from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import csv, uuid, os, msal, zipfile
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

# Azure AD
CLIENT_ID     = os.environ.get('CLIENT_ID', 'c306c8d3-68dc-4110-b5fb-771b942c10db')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET', 'XS_8Q~s3G7CwKdbOF9fcwsFBdLt1NFRiG0XvBdrL')
TENANT_ID     = os.environ.get('TENANT_ID', '63de1475-1a48-4463-aff2-b2581f2a972e')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.environ.get('REDIRECT_URI', 'http://localhost:5000/getAToken')
SCOPE         = ['User.Read']

# Paths
BASE_PATH       = os.path.dirname(__file__)
CSV_PATH        = os.path.join(BASE_PATH, 'csv')
OUTPUT_PATH     = os.path.join(BASE_PATH, 'documentos_generados')
TEMPLATE_PATH   = os.path.join(BASE_PATH, 'plantillas', 'plantilla_solicitud_completa.docx')

os.makedirs(CSV_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

# CSV archivos
USUARIOS_CSV   = os.path.join(CSV_PATH, 'usuarios.csv')
PROFESORES_CSV = os.path.join(CSV_PATH, 'profesores.csv')
AREAS_CSV      = os.path.join(CSV_PATH, 'areas.csv')
ALUMNOS_CSV    = os.path.join(CSV_PATH, 'alumnos.csv')

# Campos de alumnos
CAMPOS_ALUMNOS = [
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio','No_Control',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia','Titular_Dependencia',
    'Director_Dependencia','Responsable_Proyecto','Cargo_Responsable','Nombre_Programa','Modalidad_Externa',
    'Modalidad_Interna','Fecha_Inicio','Fecha_Terminacion','Actividades','TP_Edu_Adultos','TP_Deportivo',
    'TP_Civico','TP_Salud','TP_Otros','TP_Desarrollo','TP_Cultural','TP_Sustentable','TP_Medio_Amb',
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'
]

# --- CSV Utilities ---
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

# --- Autenticación y Sesión ---
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get('correo', '').strip().lower() == email:
            return u
    return None

def validar_acceso(roles):
    u = session.get('user')
    return u and u.get('rol') in roles

# --- MSAL Login Flow ---
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
        flash(result.get('error_description', 'Auth error'), 'danger')
        return redirect(url_for('login'))

    claims = result.get('id_token_claims', {})
    email = claims.get('preferred_username', '').lower()

    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {'correo': email, 'rol': 'Alumno', 'area': ''}
    session['user'] = {
        'correo': perfil['correo'],
        'rol': perfil['rol'],
        'area': perfil.get('area', ''),
        'name': claims.get('name', '')
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri=" +
        url_for('login', _external=True)
    )

# --- Dashboard ---
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

# --- Mapa dinámico ---
def obtener_mapa(tipo):
    return {
        'usuarios':   (USUARIOS_CSV,   ['correo', 'rol', 'area', 'profesor']),
        'profesores': (PROFESORES_CSV, ['correo', 'nombre', 'area']),
        'areas':      (AREAS_CSV,      ['id', 'nombre', 'encargado']),
        'alumnos':    (ALUMNOS_CSV,    CAMPOS_ALUMNOS)
    }.get(tipo)

# --- CRUD Genérico ---
@app.route('/entidad/<tipo>')
def entidad_list(tipo):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'), 404
    archivo, campos = m
    registros = cargar_csv(archivo)
    return render_template('entidad_list.html', tipo=tipo, campos=campos, registros=registros)

@app.route('/entidad/<tipo>/new', methods=['GET', 'POST'])
def entidad_new(tipo):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'), 404
    archivo, campos = m

    if request.method == 'POST':
        row = {c: request.form.get(c, '').strip() for c in campos if not c.startswith('TP_')}

        if any(v == '' for v in row.values()):
            flash('Todos los campos son obligatorios', 'danger')
            return redirect(url_for('entidad_new', tipo=tipo))

        tp_fields = [c for c in campos if c.startswith('TP_')]
        sel = request.form.get('tipo_participacion')
        for c in tp_fields:
            row[c] = 'X' if c == sel else ''

        registros = cargar_csv(archivo)
        pk = campos[0]
        if pk == 'id':
            row['id'] = str(uuid.uuid4())

        if any(r.get(pk, '') == row.get(pk) for r in registros):
            flash(f'El valor {row.get(pk)} ya existe', 'danger')
            return redirect(url_for('entidad_new', tipo=tipo))

        registros.append(row)
        guardar_csv(archivo, registros, campos)
        return redirect(url_for('entidad_list', tipo=tipo))

    tp_fields = [c for c in campos if c.startswith('TP_')]
    return render_template('entidad_form.html', tipo=tipo, campos=campos, valores={}, tp_fields=tp_fields)

@app.route('/entidad/<tipo>/edit/<pk>', methods=['GET', 'POST'])
def entidad_edit(tipo, pk):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'), 404
    archivo, campos = m
    pk_field = campos[0]

    registros = cargar_csv(archivo)
    item = next((r for r in registros if r.get(pk_field) == pk), None)
    if not item:
        flash('Registro no encontrado', 'danger')
        return redirect(url_for('entidad_list', tipo=tipo))

    if request.method == 'POST':
        for c in campos:
            if not c.startswith('TP_'):
                v = request.form.get(c, '').strip()
                if v == '':
                    flash(f'El campo {c} no puede quedar vacío', 'danger')
                    return redirect(url_for('entidad_edit', tipo=tipo, pk=pk))
                item[c] = v

        tp_fields = [c for c in campos if c.startswith('TP_')]
        sel = request.form.get('tipo_participacion')
        for c in tp_fields:
            item[c] = 'X' if c == sel else ''

        guardar_csv(archivo, registros, campos)
        return redirect(url_for('entidad_list', tipo=tipo))

    tp_fields = [c for c in campos if c.startswith('TP_')]
    return render_template('entidad_form.html', tipo=tipo, campos=campos, valores=item, tp_fields=tp_fields)

@app.route('/entidad/<tipo>/delete/<pk>', methods=['POST'])
def entidad_delete(tipo, pk):
    m = obtener_mapa(tipo)
    if not m:
        return render_template('404.html'), 404
    archivo, campos = m
    pk_field = campos[0]
    registros = [r for r in cargar_csv(archivo) if r.get(pk_field) != pk]
    guardar_csv(archivo, registros, campos)
    return redirect(url_for('entidad_list', tipo=tipo))

# --- Generación de documentos Word ---
@app.route('/generar-documentos')
def generar_documentos():
    try:
        df = pd.read_csv(ALUMNOS_CSV)
        for _, alumno in df.iterrows():
            doc = DocxTemplate(TEMPLATE_PATH)
            ctx = {k: str(v) for k, v in alumno.items()}
            doc.render(ctx)
            doc.save(os.path.join(OUTPUT_PATH, f"Solicitud_{alumno['No_Control']}.docx"))
        flash('Documentos generados exitosamente', 'success')
    except Exception as e:
        flash(f'Error generando documentos: {e}', 'danger')
    return redirect(url_for('dashboard'))

@app.route('/descargar-documentos')
def descargar_documentos():
    memory = BytesIO()
    with zipfile.ZipFile(memory, 'w') as zf:
        for fn in os.listdir(OUTPUT_PATH):
            zf.write(os.path.join(OUTPUT_PATH, fn), arcname=fn)
    memory.seek(0)
    return send_file(memory, download_name='documentos_alumnos.zip', as_attachment=True)

# --- Error 404 ---
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

# --- Main ---
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
