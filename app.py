import os
import csv
import uuid
from io import BytesIO
from functools import wraps

from flask import Flask, render_template, redirect, url_for, session, request, flash, send_file
from dotenv import load_dotenv
import msal
from docxtpl import DocxTemplate

load_dotenv()

app = Flask(__name__)
# Usar la variable FLASK_SECRET de tu .env
app.secret_key = os.getenv('FLASK_SECRET', 'supersecret')

# ── Configuración Azure AD ─────────────────────────────────────────────────────
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# ── Rutas de CSV ────────────────────────────────────────────────────────────────
CSV_DIR        = os.path.join(app.root_path, 'csv')
USERS_CSV      = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
# Usar los CSV completos con todos los campos
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos_completo.csv')

# ── Plantillas DOCX ────────────────────────────────────────────────────────────
TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR, 'plantilla_reporte_final.docx')

# ── Campos para documentos.csv ─────────────────────────────────────────────────
DOC_FIELDS = [
    'No_Control',
    # Solicitud
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos',
    'Dependencia','Domicilio_Dependencia','Titular_Dependencia','Cargo_Responsable',
    'Responsable_Proyecto','Nombre_Programa','Modalidad_Externa','Modalidad_Interna',
    'Fecha_Inicio','Fecha_Terminacion','Actividades',
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros',
    # Solicitud fechas
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Bimestral
    'Dia1_1','Mes1_1','Anio1_1','Dia2_1','Mes2_1','Anio2_1',
    'Dia1_2','Mes1_2','Anio1_2','Dia2_2','Mes2_2','Anio2_2',
    'Dia1_3','Mes1_3','Anio1_3','Dia2_3','Mes2_3','Anio2_3',
    'Nombre_supervisor','Puesto_supervisor',
    'Actividad_1','Actividad_2','Actividad_3','Actividad_4',
    'Actividad_5','Actividad_6','Actividad_7','Actividad_8',
    # Evaluaciones (resp/estu)
] + [
    f'{pref}{i}' for pref in (
        'resp3_','estu3_','resp4_','resp5_','estu5_','resp6_','resp7_','estu7_',
        'resp8_','resp9_','estu9_','resp10_','resp11_','estu11_','resp12_',
        'resp13_','estu13_','resp14_','resp15_','estu15_','resp16_','estu17_'
    ) for i in range(5)
] + [
    # Final
    'Municipio','Estado',
    'Actividad1','Logro1','Actividad2','Logro2','Actividad3','Logro3','Actividad4','Logro4',
    'Actividad5','Logro5','Actividad6','Logro6','Actividad7','Logro7','Actividad8','Logro8',
    'Aprendizaje1','Beneficio1','Aprendizaje2','Beneficio2','Aprendizaje3','Beneficio3',
    'Aprendizaje4','Beneficio4','Aprendizaje5','Beneficio5','Aprendizaje6','Beneficio6',
    'Aprendizaje7','Beneficio7','Aprendizaje8','Beneficio8',
    'Nombre_Responsable','Cargo_Responsable'
]

# ── Helpers ────────────────────────────────────────────────────────────────────
def cargar_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def guardar_csv(path, rows, fieldnames):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper

def roles_required(*roles):
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if session.get('role') not in roles:
                flash('Acceso denegado','danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return decorator

# ── Autenticación ─────────────────────────────────────────────────────────────
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
    if 'error' in result:
        flash(result.get('error_description','Error al autenticar'),'danger')
        return redirect(url_for('login'))
    claims = result['id_token_claims']
    email = claims['preferred_username'].lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales",403

    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u['correo'].lower()==email), None)
    if not perfil:
        perfil = {'correo':email,'rol':'Alumno','area':''}

    session.update(user=perfil, role=perfil['rol'], name=claims.get('name',''))
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login', _external=True)}"
    )

# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# ── CRUD Usuarios ─────────────────────────────────────────────────────────────
@app.route('/usuarios')
@roles_required('Administrador')
def usuarios_list():
    return render_template('usuarios_list.html',
                           usuarios=cargar_csv(USERS_CSV))

@app.route('/usuarios/new', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_new():
    if request.method=='POST':
        u = cargar_csv(USERS_CSV)
        u.append({
            'correo': request.form['correo'].strip().lower(),
            'rol':    request.form['rol'],
            'area':   request.form['area']
        })
        guardar_csv(USERS_CSV, u, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    u = cargar_csv(USERS_CSV)
    user = next((x for x in u if x['correo']==correo), None)
    if not user:
        flash('No encontrado','danger')
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        user['rol']=request.form['rol']
        user['area']=request.form['area']
        guardar_csv(USERS_CSV, u, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=user)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    u = [x for x in cargar_csv(USERS_CSV) if x['correo']!=correo]
    guardar_csv(USERS_CSV, u, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ── CRUD Áreas ───────────────────────────────────────────────────────────────
@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    return render_template('areas_list.html',
                           areas=cargar_csv(AREAS_CSV))

@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    if request.method=='POST':
        a = cargar_csv(AREAS_CSV)
        a.append({
            'id':      str(uuid.uuid4()),
            'nombre':  request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    a = cargar_csv(AREAS_CSV)
    area = next((x for x in a if x['id']==id), None)
    if not area:
        flash('No encontrado','danger')
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        area['nombre']=request.form['nombre']
        area['encargado']=request.form['encargado']
        guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    a = [x for x in cargar_csv(AREAS_CSV) if x['id']!=id]
    guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ── CRUD Profesores ──────────────────────────────────────────────────────────
@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    return render_template('profesores_list.html',
                           profesores=cargar_csv(PROFESORES_CSV))

@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method=='POST':
        p = cargar_csv(PROFESORES_CSV)
        p.append({
            'correo': request.form['correo'].strip().lower(),
            'nombre': request.form['nombre'],
            'area':   request.form['area']
        })
        guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    p = cargar_csv(PROFESORES_CSV)
    prof = next((x for x in p if x['correo']==correo), None)
    if not prof:
        flash('No encontrado','danger')
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        prof['nombre']=request.form['nombre']
        prof['area']=request.form['area']
        guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    p = [x for x in cargar_csv(PROFESORES_CSV) if x['correo']!=correo]
    guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ── CRUD Alumnos ─────────────────────────────────────────────────────────────
@app.route('/alumnos')
@login_required
def alumnos_list():
    al = cargar_csv(ALUMNOS_CSV)
    role = session['role']
    user = session['user']
    if role == 'Maestro':
        al = [x for x in al if x['profesor'].lower() == user['correo']]
    if role == 'Encargado':
        al = [x for x in al if x['area'] == user['area']]
    return render_template('alumnos_list.html', alumnos=al)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method=='POST':
        a = cargar_csv(ALUMNOS_CSV)
        a.append({
            'No_Control': request.form['No_Control'],
            'Apellido_Paterno': request.form.get('Apellido_Paterno',''),
            'Apellido_Materno': request.form.get('Apellido_Materno',''),
            'Nombre': request.form.get('Nombre',''),
            'Sexo': request.form.get('Sexo',''),
            'Telefono': request.form.get('Telefono',''),
            'Correo': request.form.get('Correo',''),
            'Domicilio': request.form.get('Domicilio',''),
            'Carrera': request.form.get('Carrera',''),
            'Periodo': request.form.get('Periodo',''),
            'Semestre': request.form.get('Semestre',''),
            'Creditos': request.form.get('Creditos',''),
            'area': request.form['area'],
            'profesor': request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, a, [*DOC_FIELDS, 'profesor','area'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(no_control):
    a = cargar_csv(ALUMNOS_CSV)
    alum = next((x for x in a if x['No_Control']==no_control), None)
    if not alum:
        flash('No encontrado','danger')
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        for k in ('Apellido_Paterno','Apellido_Materno','Nombre','Sexo',
                  'Telefono','Correo','Domicilio','Carrera','Periodo',
                  'Semestre','Creditos','area','profesor'):
            if k in request.form:
                alum[k] = request.form.get(k,'')
        guardar_csv(ALUMNOS_CSV, a, [*DOC_FIELDS, 'profesor','area'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<no_control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(no_control):
    a = [x for x in cargar_csv(ALUMNOS_CSV) if x['No_Control']!=no_control]
    guardar_csv(ALUMNOS_CSV, a, [*DOC_FIELDS, 'profesor','area'])
    return redirect(url_for('alumnos_list'))

# ── CRUD Documentos ───────────────────────────────────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    return render_template('documentos_list.html',
                           documentos=cargar_csv(DOCUMENTOS_CSV),
                           alumnos=cargar_csv(ALUMNOS_CSV))

@app.route('/documentos/new', methods=['GET','POST'])
@login_required
def documentos_new():
    alumnos = cargar_csv(ALUMNOS_CSV)
    if request.method=='POST':
        docs = cargar_csv(DOCUMENTOS_CSV)
        nuevo = {f: request.form.get(f,'') for f in DOC_FIELDS}
        # autocompleta campos de alumnos
        alum = next((x for x in alumnos if x['No_Control']==nuevo['No_Control']), {})
        for k in ('Apellido_Paterno','Apellido_Materno','Nombre','Carrera'):
            nuevo[k] = alum.get(k,nuevo.get(k,''))
        docs.append(nuevo)
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html',
                           alumnos=alumnos, documento={})

@app.route('/documentos/edit/<no_control>', methods=['GET','POST'])
@login_required
def documentos_edit(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((x for x in docs if x['No_Control']==no_control), None)
    if not doc:
        flash('No encontrado','danger')
        return redirect(url_for('documentos_list'))
    if request.method=='POST':
        for f in DOC_FIELDS:
            if f != 'No_Control':
                doc[f] = request.form.get(f,'')
        # autocompleta de alumnos
        alum = next((x for x in alumnos if x['No_Control']==no_control), {})
        for k in ('Apellido_Paterno','Apellido_Materno','Nombre','Carrera'):
            doc[k] = alum.get(k, doc[k])
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html',
                           alumnos=alumnos, documento=doc)

@app.route('/documentos/delete/<no_control>')
@login_required
def documentos_delete(no_control):
    docs = [x for x in cargar_csv(DOCUMENTOS_CSV) if x['No_Control']!=no_control]
    guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
    return redirect(url_for('documentos_list'))

# ── Generación de DOCX ───────────────────────────────────────────────────────
def _render_docx(template_path, context, filename):
    doc = DocxTemplate(template_path)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alum = next((x for x in cargar_csv(ALUMNOS_CSV) if x['No_Control']==no_control), {})
    doci = next((x for x in cargar_csv(DOCUMENTOS_CSV) if x['No_Control']==no_control), {})
    ctx = {**alum, **doci}
    return _render_docx(TPL_SOLICITUD, ctx, f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    if num < 1 or num > 3:
        flash('Bimestre inválido','warning')
        return redirect(url_for('dashboard'))
    alum = next((x for x in cargar_csv(ALUMNOS_CSV) if x['No_Control']==no_control), {})
    doci = next((x for x in cargar_csv(DOCUMENTOS_CSV) if x['No_Control']==no_control), {})
    ctx = {**alum, **doci, 'reporte_no': num}
    return _render_docx(TPL_BIMESTRAL, ctx, f"Bimestral_{no_control}_B{num}.docx")

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alum = next((x for x in cargar_csv(ALUMNOS_CSV) if x['No_Control']==no_control), {})
    doci = next((x for x in cargar_csv(DOCUMENTOS_CSV) if x['No_Control']==no_control), {})
    ctx = {**alum, **doci}
    return _render_docx(TPL_FINAL, ctx, f"Final_{no_control}.docx")

# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(error):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0',
            port=int(os.getenv('PORT', 5000)))
