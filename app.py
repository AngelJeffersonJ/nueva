import os
import csv
import uuid
from io import BytesIO
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
app.secret_key = os.getenv('FLASK_SECRET_KEY','supersecret')

# ── Configuración Azure AD ─────────────────────────────────────────────────────
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# ── Rutas de CSV ────────────────────────────────────────────────────────────────
CSV_DIR        = os.path.join(app.root_path,'csv')
USERS_CSV      = os.path.join(CSV_DIR,'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR,'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR,'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR,'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR,'documentos_completo.csv')

# ── Plantillas DOCX ────────────────────────────────────────────────────────────
TPL_DIR        = os.path.join(app.root_path,'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR,'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR,'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR,'Reporte_Final_Lleno.docx')

# ── Sinónimos de campos (minúscula) → canónicos ────────────────────────────────
FIELD_SYNONYMS = {
    'ap':                    'Apellido_Paterno',
    'am':                    'Apellido_Materno',
    'no_de_control':         'No_Control',
    'carrera':               'Carrera',
    'carrera>':              'Carrera',
    'municipio':             'Municipio',
    'estado':                'Estado',
    'num_control':           'Num_Control',
    'nombre_estudiante':     'Nombre',
    'observaciones_encargado':'Observaciones_Encargado',
    'observaciones_estudiante':'Observaciones_Estudiante',
}

# ── Campos esperados ───────────────────────────────────────────────────────────
ALUMNOS_FIELDS = ['No_Control','nombre','area','profesor']

DOC_FIELDS = [
    'No_Control',
    # Solicitud
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia',
    'Titular_Dependencia','Cargo_Responsable','Responsable_Proyecto','Nombre_Programa',
    'Modalidad_Externa','Modalidad_Interna','Fecha_Inicio','Fecha_Terminacion','Actividades',
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros',
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
] + [f'{part}_{n}' for n in (1,2,3) for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2')] + [
    'Nombre_supervisor','Puesto_supervisor'
] + [f'Actividad_{i}' for i in range(1,9)] + [
    'x1','x2','x3'
] + [f'resp{n}_{i}' for n in range(3,18) for i in range(5)] + [
    f'estu{n}_{i}' for n in range(3,18) for i in range(5)
] + [
    # Final
    'Municipio','Estado'
] + [f'Actividad{i}' for i in range(1,9)] + [f'Logro{i}' for i in range(1,9)] + [
    f'Aprendizaje{i}' for i in range(1,9)
] + [f'Beneficio{i}' for i in range(1,9)] + [
    'Nombre_Responsable','Cargo_Responsable','Num_Control',
    'Observaciones_Encargado','Observaciones_Estudiante'
]

# ── Carga CSV normalizando sinónimos y rellenando campos faltantes ────────────
def cargar_csv(path, expected_fields=None):
    rows = []
    if os.path.isfile(path):
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for raw in reader:
                norm = {}
                for key,val in raw.items():
                    canon = FIELD_SYNONYMS.get(key.strip().lower(), key.strip())
                    norm[canon] = val.strip()
                if expected_fields:
                    for fld in expected_fields:
                        norm.setdefault(fld,'')
                rows.append(norm)
    return rows

# ── Guarda CSV con campos canónicos ────────────────────────────────────────────
def guardar_csv(path, rows, fieldnames):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'w',newline='',encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader(); w.writerows(rows)
    except Exception as e:
        app.logger.error(f"Error guardando CSV: {e}")
        flash("No se pudo guardar el CSV","danger")

# ── Decoradores de autenticación/roles ─────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def w(*args,**kw):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args,**kw)
    return w

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*args,**kw):
            if session.get('role') not in roles:
                flash("Acceso denegado","danger")
                return redirect(url_for('dashboard'))
            return f(*args,**kw)
        return w
    return deco

# ── Renderiza y envía DOCX ─────────────────────────────────────────────────────
def _render_docx(template_path, context, filename):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return send_file(
            buf, as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento","danger")
        return redirect(url_for('dashboard'))

# ── Autenticación Azure AD ────────────────────────────────────────────────────
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
    email  = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales",403
    usuarios = cargar_csv(USERS_CSV)
    perfil   = next((u for u in usuarios if u.get('correo','').lower()==email), None)
    if not perfil:
        perfil = {'correo':email,'rol':'Alumno','area':''}
    session.update(user=perfil, role=perfil['rol'], name=claims.get('name',''))
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login',_external=True)}"
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
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'rol':    request.form.get('rol','').strip(),
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("Correo obligatorio","warning")
            return redirect(url_for('usuarios_new'))
        u.append(nuevo)
        guardar_csv(USERS_CSV,u,['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    u = cargar_csv(USERS_CSV)
    usr = next((x for x in u if x['correo']==correo), None)
    if not usr:
        flash("No encontrado","danger")
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        usr['rol']  = request.form.get('rol','').strip()
        usr['area'] = request.form.get('area','').strip()
        guardar_csv(USERS_CSV,u,['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=usr)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    u = [x for x in cargar_csv(USERS_CSV) if x['correo']!=correo]
    guardar_csv(USERS_CSV,u,['correo','rol','area'])
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
        nueva = {
            'id':        str(uuid.uuid4()),
            'nombre':    request.form.get('nombre','').strip(),
            'encargado': request.form.get('encargado','').strip()
        }
        if not nueva['nombre']:
            flash("Nombre obligatorio","warning")
            return redirect(url_for('areas_new'))
        a.append(nueva)
        guardar_csv(AREAS_CSV,a,['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    a = cargar_csv(AREAS_CSV)
    ar = next((x for x in a if x['id']==id), None)
    if not ar:
        flash("No encontrado","danger")
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        ar['nombre']    = request.form.get('nombre','').strip()
        ar['encargado'] = request.form.get('encargado','').strip()
        guardar_csv(AREAS_CSV,a,['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=ar)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    a = [x for x in cargar_csv(AREAS_CSV) if x['id']!=id]
    guardar_csv(AREAS_CSV,a,['id','nombre','encargado'])
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
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'nombre': request.form.get('nombre','').strip(),
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("Correo obligatorio","warning")
            return redirect(url_for('profesores_new'))
        p.append(nuevo)
        guardar_csv(PROFESORES_CSV,p,['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    p = cargar_csv(PROFESORES_CSV)
    prof = next((x for x in p if x['correo']==correo), None)
    if not prof:
        flash("No encontrado","danger")
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        prof['nombre'] = request.form.get('nombre','').strip()
        prof['area']   = request.form.get('area','').strip()
        guardar_csv(PROFESORES_CSV,p,['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    p = [x for x in cargar_csv(PROFESORES_CSV) if x['correo']!=correo]
    guardar_csv(PROFESORES_CSV,p,['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ── CRUD Alumnos ─────────────────────────────────────────────────────────────
@app.route('/alumnos')
@login_required
def alumnos_list():
    al    = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    role  = session['role']; usr = session['user']
    if role=='Maestro':
        al = [x for x in al if x['profesor'].lower()==usr['correo'].lower()]
    if role=='Encargado':
        al = [x for x in al if x['area']==usr['area']]
    return render_template('alumnos_list.html', alumnos=al)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method=='POST':
        no_ctl = request.form.get('No_Control','').strip()
        if not no_ctl:
            flash("No_Control obligatorio","warning")
            return redirect(url_for('alumnos_new'))
        lst = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
        lst.append({
            'No_Control':no_ctl,
            'nombre':     request.form.get('nombre','').strip(),
            'area':       request.form.get('area','').strip(),
            'profesor':   request.form.get('profesor','').strip
        })
        guardar_csv(ALUMNOS_CSV,lst,ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<No_Control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(No_Control):
    lst  = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    alum = next((x for x in lst if x['No_Control']==No_Control), None)
    if not alum:
        flash("No encontrado","danger")
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        alum['nombre']   = request.form.get('nombre','').strip()
        alum['area']     = request.form.get('area','').strip()
        alum['profesor'] = request.form.get('profesor','').strip()
        guardar_csv(ALUMNOS_CSV,lst,ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<No_Control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(No_Control):
    lst = [x for x in cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS) if x['No_Control']!=No_Control]
    guardar_csv(ALUMNOS_CSV,lst,ALUMNOS_FIELDS)
    return redirect(url_for('alumnos_list'))

# ── CRUD Documentos ───────────────────────────────────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    return render_template('documentos_list.html',
                           documentos=docs, alumnos=alumnos)

@app.route('/documentos/new')
@login_required
def documentos_new():
    flash("Para crear o editar documentos usa los botones de Solicitud, Bimestral o Final.","info")
    return redirect(url_for('documentos_list'))

@app.route('/documentos/edit/<no_control>')
@login_required
def documentos_edit(no_control):
    return redirect(url_for('documentos_solicitud', no_control=no_control))

@app.route('/documentos/delete/<no_control>')
@login_required
def documentos_delete(no_control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS) if d['No_Control']!=no_control]
    guardar_csv(DOCUMENTOS_CSV,docs,DOC_FIELDS)
    return redirect(url_for('documentos_list'))

# ── Solicitud ─────────────────────────────────────────────────────────────────
SOL_TEXT_FIELDS = [
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia',
    'Titular_Dependencia','Cargo_Responsable','Responsable_Proyecto','Nombre_Programa',
    'Fecha_Inicio','Fecha_Terminacion','Actividades'
]
SOL_TP_FIELDS = [
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros'
]

@app.route('/documentos/solicitud/<no_control>', methods=['GET','POST'])
@login_required
def documentos_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    doc  = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method=='POST':
        faltan = [f for f in SOL_TEXT_FIELDS if not request.form.get(f)]
        tp_sel = [f for f in SOL_TP_FIELDS if request.form.get(f)]
        if len(tp_sel)!=1:
            faltan.append('Tipo de Proyecto')
        for part in ('Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'):
            if not request.form.get(part):
                faltan.append(part)
        if faltan:
            flash("Completa todos los campos: "+", ".join(faltan),"warning")
            return redirect(request.url)
        if not doc:
            doc = {k:'' for k in DOC_FIELDS}
            doc['No_Control']=no_control; docs.append(doc)
        for f in SOL_TEXT_FIELDS:
            doc[f]=request.form[f].strip()
        for f in SOL_TP_FIELDS:
            doc[f]='X' if f in tp_sel else ''
        for part in ('Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'):
            doc[part]=request.form.get(part).strip()
        guardar_csv(DOCUMENTOS_CSV,docs,DOC_FIELDS)
        flash("Solicitud guardada","success")
        return redirect(url_for('documentos_list'))
    return render_template('solicitud_form.html', no_control=no_control, documento=doc or {})

# ── Bimestral ─────────────────────────────────────────────────────────────────
@app.route('/documentos/bimestral/<no_control>/<int:num>', methods=['GET','POST'])
@login_required
def documentos_bimestral(no_control,num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    docs = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    doc  = next((d for d in docs if d['No_Control']==no_control),None)
    if request.method=='POST':
        faltan=[]
        for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2'):
            if not request.form.get(f"{part}_{num}"):
                faltan.append(f"{part}_{num}")
        if not request.form.get('Nombre_supervisor'):
            faltan.append('Nombre_supervisor')
        if not request.form.get('Puesto_supervisor'):
            faltan.append('Puesto_supervisor')
        for i in range(1,9):
            if not request.form.get(f"Actividad_{i}"):
                faltan.append(f"Actividad_{i}")
        if not request.form.get('bim'):
            faltan.append('bim')
        for n in range(3,18):
            if not request.form.get(f"resp{n}") or not request.form.get(f"estu{n}"):
                faltan.append(f"resp{n}/estu{n}")
        if faltan:
            flash("Completa todos los campos: "+", ".join(faltan),"warning")
            return redirect(request.url)
        if not doc:
            doc={k:'' for k in DOC_FIELDS}
            doc['No_Control']=no_control; docs.append(doc)
        for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2'):
            doc[f"{part}_{num}"]=request.form.get(f"{part}_{num}")
        doc['Nombre_supervisor']=request.form.get('Nombre_supervisor').strip()
        doc['Puesto_supervisor']=request.form.get('Puesto_supervisor').strip()
        for i in range(1,9):
            doc[f"Actividad_{i}"]=request.form.get(f"Actividad_{i}").strip()
        sel=request.form.get('bim')
        doc['x1'],doc['x2'],doc['x3']=('X' if sel=='1' else '',
                                       'X' if sel=='2' else '',
                                       'X' if sel=='3' else '')
        for n in range(3,18):
            for i in range(5):
                doc[f"resp{n}_{i}"]=''
                doc[f"estu{n}_{i}"]=''
            dr=int(request.form.get(f"resp{n}"))
            de=int(request.form.get(f"estu{n}"))
            doc[f"resp{n}_{dr}"]='X'
            doc[f"estu{n}_{de}"]='X'
        guardar_csv(DOCUMENTOS_CSV,docs,DOC_FIELDS)
        flash(f"Bimestral {num} guardado","success")
        return redirect(url_for('documentos_list'))
    return render_template('bimestral_form.html',
                           no_control=no_control, num=num, documento=doc or {})

# ── Final ─────────────────────────────────────────────────────────────────────
FINAL_FIELDS = (
    ['Municipio','Estado'] +
    [f'Actividad{i}' for i in range(1,9)] +
    [f'Logro{i}'     for i in range(1,9)] +
    [f'Aprendizaje{i}' for i in range(1,9)] +
    [f'Beneficio{i}'   for i in range(1,9)] +
    ['Nombre_Responsable','Cargo_Responsable','Num_Control',
     'Observaciones_Encargado','Observaciones_Estudiante']
)

@app.route('/documentos/final/<no_control>', methods=['GET','POST'])
@login_required
def documentos_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    doc  = next((d for d in docs if d['No_Control']==no_control),None)
    if request.method=='POST':
        faltan=[f for f in FINAL_FIELDS if not request.form.get(f)]
        if faltan:
            flash("Completa todos los campos: "+", ".join(faltan),"warning")
            return redirect(request.url)
        if not doc:
            doc={k:'' for k in DOC_FIELDS}
            doc['No_Control']=no_control; docs.append(doc)
        for f in FINAL_FIELDS:
            doc[f]=request.form.get(f).strip()
        guardar_csv(DOCUMENTOS_CSV,docs,DOC_FIELDS)
        flash("Reporte final guardado","success")
        return redirect(url_for('documentos_list'))
    return render_template('final_form.html',
                           no_control=no_control, documento=doc or {})

# ── Generación de DOCX ───────────────────────────────────────────────────────
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(TPL_SOLICITUD,{**alum,**doci},f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control,num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_BIMESTRAL,
        {**alum,**doci,'reporte_no':num},
        f"Bimestral_{no_control}_B{num}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    doci['Nombre_Estudiante'] = doci.get('Nombre',alum.get('Nombre',''))
    return _render_docx(TPL_FINAL,{**alum,**doci},f"Final_{no_control}.docx")

# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'),404

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0',
            port=int(os.getenv('PORT',5000)))
