import os
import csv
import uuid
from io import BytesIO
from zipfile import ZipFile
from functools import wraps

from flask import (
    Flask, render_template, redirect, url_for,
    session, request, flash, send_file
)
from dotenv import load_dotenv
import msal
from docxtpl import DocxTemplate

load_dotenv()

app = Flask(__name__)
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
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos_completo.csv')

# ── Plantillas DOCX ────────────────────────────────────────────────────────────
TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

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
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Bimestrales
    'Dia1_1','Mes1_1','Anio1_1','Dia2_1','Mes2_1','Anio2_1',
    'Dia1_2','Mes1_2','Anio1_2','Dia2_2','Mes2_2','Anio2_2',
    'Dia1_3','Mes1_3','Anio1_3','Dia2_3','Mes2_3','Anio2_3',
    'Nombre_supervisor','Puesto_supervisor',
    'Actividad_1','Actividad_2','Actividad_3','Actividad_4',
    'Actividad_5','Actividad_6','Actividad_7','Actividad_8',
]
# Evaluaciones dinámicas
for n in range(3, 18):
    for i in range(5):
        DOC_FIELDS.append(f'resp{n}_{i}')
    for i in range(5):
        DOC_FIELDS.append(f'estu{n}_{i}')

# Reporte final
DOC_FIELDS += [
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
    try:
        if not os.path.isfile(path):
            return []
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    except Exception as e:
        app.logger.error(f"Error cargando CSV {path}: {e}")
        return []

def guardar_csv(path, rows, fieldnames):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        app.logger.error(f"Error guardando CSV {path}: {e}")
        flash("No se pudo guardar el CSV", 'danger')

def login_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapped

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if session.get('role') not in roles:
                flash('Acceso denegado','danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapped
    return deco

def _render_docx(template_path, context, filename):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype=(
              'application/vnd.openxmlformats-'
              'officedocument.wordprocessingml.document'
            )
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento", 'danger')
        return redirect(url_for('dashboard'))

# ── Autenticación ─────────────────────────────────────────────────────────────
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if 'error' in result:
        flash(result.get('error_description','Error al autenticar'),'danger')
        return redirect(url_for('login'))
    claims = result['id_token_claims']
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403
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
        f"{AUTHORITY}/oauth2/v2.0/logout?"
        f"post_logout_redirect_uri={url_for('login', _external=True)}"
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
        usuarios = cargar_csv(USERS_CSV)
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'rol':    request.form.get('rol','').strip(),
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("El correo es obligatorio","warning")
            return redirect(url_for('usuarios_new'))
        usuarios.append(nuevo)
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u['correo']==correo), None)
    if not perfil:
        flash("Usuario no encontrado","danger")
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        perfil['rol']  = request.form.get('rol','').strip()
        perfil['area'] = request.form.get('area','').strip()
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=perfil)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    usuarios = [u for u in cargar_csv(USERS_CSV) if u['correo']!=correo]
    guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
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
        areas = cargar_csv(AREAS_CSV)
        nueva = {
            'id':        str(uuid.uuid4()),
            'nombre':    request.form.get('nombre','').strip(),
            'encargado': request.form.get('encargado','').strip()
        }
        if not nueva['nombre']:
            flash("El nombre es obligatorio","warning")
            return redirect(url_for('areas_new'))
        areas.append(nueva)
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    areas = cargar_csv(AREAS_CSV)
    area = next((a for a in areas if a['id']==id), None)
    if not area:
        flash("Área no encontrada","danger")
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        area['nombre']    = request.form.get('nombre','').strip()
        area['encargado'] = request.form.get('encargado','').strip()
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    areas = [a for a in cargar_csv(AREAS_CSV) if a['id']!=id]
    guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
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
        profs = cargar_csv(PROFESORES_CSV)
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'nombre': request.form.get('nombre','').strip(),
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("El correo es obligatorio","warning")
            return redirect(url_for('profesores_new'))
        profs.append(nuevo)
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    profs = cargar_csv(PROFESORES_CSV)
    prof = next((p for p in profs if p['correo']==correo), None)
    if not prof:
        flash("Profesor no encontrado","danger")
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        prof['nombre'] = request.form.get('nombre','').strip()
        prof['area']   = request.form.get('area','').strip()
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    profs = [p for p in cargar_csv(PROFESORES_CSV) if p['correo']!=correo]
    guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))


# ── CRUD Alumnos ─────────────────────────────────────────────────────────────
@app.route('/alumnos')
@login_required
def alumnos_list():
    alumnos = cargar_csv(ALUMNOS_CSV)
    role = session.get('role')
    usr  = session.get('user',{})
    if role=='Maestro':
        alumnos=[a for a in alumnos if a['profesor'].lower()==usr['correo'].lower()]
    elif role=='Encargado':
        alumnos=[a for a in alumnos if a['area']==usr['area']]
    return render_template('alumnos_list.html', alumnos=alumnos)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method=='POST':
        no_ctl = request.form.get('no_control','').strip()
        if not no_ctl:
            flash("No_Control es obligatorio","warning")
            return redirect(url_for('alumnos_new'))
        lst = cargar_csv(ALUMNOS_CSV)
        lst.append({
            'No_Control': no_ctl,
            'nombre':     request.form.get('nombre','').strip(),
            'area':       request.form.get('area','').strip(),
            'profesor':   request.form.get('profesor','').strip()
        })
        guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<No_Control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(No_Control):
    lst = cargar_csv(ALUMNOS_CSV)
    alum = next((a for a in lst if a['No_Control']==No_Control), None)
    if not alum:
        flash("Alumno no encontrado","danger")
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        alum['nombre']   = request.form.get('nombre','').strip()
        alum['area']     = request.form.get('area','').strip()
        alum['profesor'] = request.form.get('profesor','').strip()
        guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<No_Control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(No_Control):
    lst = [a for a in cargar_csv(ALUMNOS_CSV) if a['No_Control']!=No_Control]
    guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))


# ── CRUD Documentos: lista, nueva y eliminación ───────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    return render_template(
        'documentos_list.html',
        documentos=cargar_csv(DOCUMENTOS_CSV),
        alumnos=cargar_csv(ALUMNOS_CSV)
    )

# Stub para que exista la ruta /documentos/new usada en la plantilla
@app.route('/documentos/new')
@login_required
def documentos_new():
    flash("Seleccione en la tabla el tipo de documento que desea crear para cada alumno.", "info")
    return redirect(url_for('documentos_list'))

@app.route('/documentos/delete/<no_control>')
@login_required
def documentos_delete(no_control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control'] != no_control]
    guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
    return redirect(url_for('documentos_list'))


# ── Formulario y guardado: Solicitud ──────────────────────────────────────────
SOL_FIELDS = [
  'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
  'Carrera','Periodo','Semestre','Creditos',
  'Dependencia','Domicilio_Dependencia','Titular_Dependencia','Cargo_Responsable',
  'Responsable_Proyecto','Nombre_Programa','Modalidad_Externa','Modalidad_Interna',
  'Fecha_Inicio','Fecha_Terminacion','Actividades'
] + [f'TP_{x}' for x in [
  'Edu_Adultos','Desarrollo','Deportivo','Cultural','Civico',
  'Sustentable','Salud','Medio_Amb','Otros'
]]

@app.route('/documentos/solicitud/<no_control>', methods=['GET','POST'])
@login_required
def documentos_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method=='POST':
        if not doc:
            doc = {f: '' for f in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        for field in SOL_FIELDS:
            if field.startswith('TP_'):
                doc[field] = '1' if request.form.get(field) else '0'
            else:
                doc[field] = request.form.get(field, '')
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Solicitud guardada","success")
        return redirect(url_for('documentos_list'))
    return render_template(
        'solicitud_form.html',
        no_control=no_control,
        documento=doc or {}
    )


# ── Formulario y guardado: Bimestral ─────────────────────────────────────────
BIM_FIELDS = [
    'Dia1_{n}','Mes1_{n}','Anio1_{n}',
    'Dia2_{n}','Mes2_{n}','Anio2_{n}',
    'Nombre_supervisor','Puesto_supervisor'
] + [f'Actividad_{i}' for i in range(1,9)] \
  + [f'resp{n}_{i}' for n in range(3,18) for i in range(5)] \
  + [f'estu{n}_{i}' for n in range(3,18) for i in range(5)]

@app.route('/documentos/bimestral/<no_control>/<int:num>', methods=['GET','POST'])
@login_required
def documentos_bimestral(no_control, num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method=='POST':
        if not doc:
            doc = {f: '' for f in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        # fechas dinámicas
        for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2'):
            field = f"{part}_{num}"
            doc[field] = request.form.get(field,'')
        # supervisor y actividades
        doc['Nombre_supervisor'] = request.form.get('Nombre_supervisor','')
        doc['Puesto_supervisor'] = request.form.get('Puesto_supervisor','')
        for i in range(1,9):
            doc[f'Actividad_{i}'] = request.form.get(f'Actividad_{i}','')
        # evaluaciones
        for n in range(3,18):
            for i in range(5):
                doc[f'resp{n}_{i}'] = request.form.get(f'resp{n}_{i}','')
                doc[f'estu{n}_{i}'] = request.form.get(f'estu{n}_{i}','')
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash(f"Bimestral {num} guardado","success")
        return redirect(url_for('documentos_list'))
    return render_template(
        'bimestral_form.html',
        no_control=no_control,
        num=num,
        documento=doc or {}
    )


# ── Formulario y guardado: Final ─────────────────────────────────────────────
FINAL_FIELDS = [
    'Municipio','Estado'
] + [f'Actividad{i}' for i in range(1,9)] + [f'Logro{i}' for i in range(1,9)] \
  + [f'Aprendizaje{i}' for i in range(1,9)] + [f'Beneficio{i}' for i in range(1,9)] \
  + ['Nombre_Responsable','Cargo_Responsable']

@app.route('/documentos/final/<no_control>', methods=['GET','POST'])
@login_required
def documentos_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method=='POST':
        if not doc:
            doc = {f: '' for f in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        # final fields
        for f in FINAL_FIELDS:
            doc[f] = request.form.get(f,'')
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Reporte final guardado","success")
        return redirect(url_for('documentos_list'))
    return render_template(
        'final_form.html',
        no_control=no_control,
        documento=doc or {}
    )


# ── Generación de DOCX ───────────────────────────────────────────────────────
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_SOLICITUD,
        {**alum,**doci},
        f"Solicitud_{no_control}.docx"
    )

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_BIMESTRAL,
        {**alum,**doci,'reporte_no': num},
        f"Bimestral_{no_control}_B{num}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_FINAL,
        {**alum,**doci},
        f"Final_{no_control}.docx"
    )


# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404


if __name__ == '__main__':
    app.run(
        debug=True,
        host='0.0.0.0',
        port=int(os.getenv('PORT', 5000))
    )
