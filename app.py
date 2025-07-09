import os
import csv
import uuid
from io import BytesIO
from zipfile import ZipFile
from functools import wraps

from flask import Flask, render_template, redirect, url_for, session, request, flash, send_file
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
    # Fechas solicitud
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Fechas bimestrales (3 bimestres)
    'Dia1_1','Mes1_1','Anio1_1','Dia2_1','Mes2_1','Anio2_1',
    'Dia1_2','Mes1_2','Anio1_2','Dia2_2','Mes2_2','Anio2_2',
    'Dia1_3','Mes1_3','Anio1_3','Dia2_3','Mes2_3','Anio2_3',
    'Nombre_supervisor','Puesto_supervisor',
    'Actividad_1','Actividad_2','Actividad_3','Actividad_4',
    'Actividad_5','Actividad_6','Actividad_7','Actividad_8'
] + [
    f'{pref}{i}' for pref in (
        'resp3_','estu3_','resp4_','resp5_','estu5_','resp6_','resp7_','estu7_',
        'resp8_','resp9_','estu9_','resp10_','resp11_','estu11_','resp12_',
        'resp13_','estu13_','resp14_','resp15_','estu15_','resp16_','estu17_'
    ) for i in range(5)
] + [
    # Reporte final
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
    """Carga un CSV y retorna lista de dicts; en caso de error, retorna lista vacía."""
    try:
        if not os.path.isfile(path):
            app.logger.warning(f"CSV no encontrado: {path}")
            return []
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    except Exception as e:
        app.logger.error(f"Error cargando CSV {path}: {e}")
        return []

def guardar_csv(path, rows, fieldnames):
    """Guarda lista de dicts en CSV; en error, registra log y muestra flash."""
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
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return decorator

def _render_docx(template_path, context, filename):
    """Renderiza plantilla DOCX con contexto y la envía; maneja errores."""
    try:
        if not os.path.isfile(template_path):
            flash("Plantilla no encontrada", 'danger')
            return redirect(url_for('dashboard'))
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO()
        doc.save(buf); buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento", 'danger')
        return redirect(url_for('dashboard'))

# ── Autenticación ─────────────────────────────────────────────────────────────
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
        flash(result.get('error_description', 'Error al autenticar'), 'danger')
        return redirect(url_for('login'))
    claims = result['id_token_claims']
    email  = claims.get('preferred_username', '').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403
    usuarios = cargar_csv(USERS_CSV)
    perfil   = next((u for u in usuarios if u.get('correo','').lower() == email), None)
    if not perfil:
        perfil = {'correo': email, 'rol': 'Alumno', 'area': ''}
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
    return render_template('usuarios_list.html', usuarios=cargar_csv(USERS_CSV))

@app.route('/usuarios/new', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_new():
    if request.method == 'POST':
        usuarios = cargar_csv(USERS_CSV)
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'rol':    request.form.get('rol','').strip(),  
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("El correo es obligatorio", 'warning')
            return redirect(url_for('usuarios_new'))
        usuarios.append(nuevo)
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    usuarios = cargar_csv(USERS_CSV)
    perfil   = next((u for u in usuarios if u.get('correo') == correo), None)
    if not perfil:
        flash("Usuario no encontrado", 'danger')
        return redirect(url_for('usuarios_list'))
    if request.method == 'POST':
        perfil['rol']  = request.form.get('rol','').strip()
        perfil['area'] = request.form.get('area','').strip()
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=perfil)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    usuarios = [u for u in cargar_csv(USERS_CSV) if u.get('correo') != correo]
    guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ── CRUD Áreas ───────────────────────────────────────────────────────────────
@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    return render_template('areas_list.html', areas=cargar_csv(AREAS_CSV))

@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    if request.method == 'POST':
        areas = cargar_csv(AREAS_CSV)
        nueva = {
            'id':        str(uuid.uuid4()),
            'nombre':    request.form.get('nombre','').strip(),
            'encargado': request.form.get('encargado','').strip()
        }
        if not nueva['nombre']:
            flash("El nombre es obligatorio", 'warning')
            return redirect(url_for('areas_new'))
        areas.append(nueva)
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    areas = cargar_csv(AREAS_CSV)
    area  = next((a for a in areas if a.get('id') == id), None)
    if not area:
        flash("Área no encontrada", 'danger')
        return redirect(url_for('areas_list'))
    if request.method == 'POST':
        area['nombre']    = request.form.get('nombre','').strip()
        area['encargado'] = request.form.get('encargado','').strip()
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    areas = [a for a in cargar_csv(AREAS_CSV) if a.get('id') != id]
    guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ── CRUD Profesores ──────────────────────────────────────────────────────────
@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    return render_template('profesores_list.html', profesores=cargar_csv(PROFESORES_CSV))

@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method == 'POST':
        profs = cargar_csv(PROFESORES_CSV)
        nuevo = {
            'correo': request.form.get('correo','').strip().lower(),
            'nombre': request.form.get('nombre','').strip(),
            'area':   request.form.get('area','').strip()
        }
        if not nuevo['correo']:
            flash("El correo es obligatorio", 'warning')
            return redirect(url_for('profesores_new'))
        profs.append(nuevo)
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    profs = cargar_csv(PROFESORES_CSV)
    prof  = next((p for p in profs if p.get('correo') == correo), None)
    if not prof:
        flash("Profesor no encontrado", 'danger')
        return redirect(url_for('profesores_list'))
    if request.method == 'POST':
        prof['nombre'] = request.form.get('nombre','').strip()
        prof['area']   = request.form.get('area','').strip()
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    profs = [p for p in cargar_csv(PROFESORES_CSV) if p.get('correo') != correo]
    guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ── CRUD Alumnos ─────────────────────────────────────────────────────────────
@app.route('/alumnos')
@login_required
def alumnos_list():
    lista = cargar_csv(ALUMNOS_CSV)
    role  = session.get('role')
    usr   = session.get('user', {})
    if role == 'Maestro':
        lista = [a for a in lista if a.get('profesor','').lower() == usr.get('correo','').lower()]
    elif role == 'Encargado':
        lista = [a for a in lista if a.get('area','') == usr.get('area','')]
    return render_template('alumnos_list.html', alumnos=lista)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method == 'POST':
        nueva = {
            'No_Control': request.form.get('No_Control','').strip(),
            'nombre':     request.form.get('nombre','').strip(),
            'area':       request.form.get('area','').strip(),
            'profesor':   request.form.get('profesor','').strip()
        }
        if not nueva['No_Control']:
            flash("No_Control es obligatorio", 'warning')
            return redirect(url_for('alumnos_new'))
        lst = cargar_csv(ALUMNOS_CSV)
        lst.append(nueva)
        guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(no_control):
    lst  = cargar_csv(ALUMNOS_CSV)
    alum = next((a for a in lst if a.get('No_Control') == no_control), None)
    if not alum:
        flash("Alumno no encontrado", 'danger')
        return redirect(url_for('alumnos_list'))
    if request.method == 'POST':
        alum['nombre']   = request.form.get('nombre','').strip()
        alum['area']     = request.form.get('area','').strip()
        alum['profesor'] = request.form.get('profesor','').strip()
        guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
        return	redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<no_control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(no_control):
    lst = [a for a in cargar_csv(ALUMNOS_CSV) if a.get('No_Control') != no_control]
    guardar_csv(ALUMNOS_CSV, lst, ['No_Control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))

# ── CRUD Documentos ───────────────────────────────────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    return render_template(
        'documentos_list.html',
        documentos=cargar_csv(DOCUMENTOS_CSV),
        alumnos=cargar_csv(ALUMNOS_CSV)
    )

@app.route('/documentos/new', methods=['GET','POST'])
@login_required
def documentos_new():
    alumnos = cargar_csv(ALUMNOS_CSV)
    if request.method == 'POST':
        nuevo = {f: request.form.get(f,'').strip() for f in DOC_FIELDS}
        if not nuevo.get('No_Control'):
            flash("No_Control es obligatorio", 'warning')
            return redirect(url_for('documentos_new'))
        alumno = next((a for a in alumnos if a.get('No_Control') == nuevo['No_Control']), {})
        for campo in ('Apellido_Paterno','Apellido_Materno','Nombre','Carrera'):
            nuevo[campo] = alumno.get(campo, nuevo.get(campo,''))
        docs = cargar_csv(DOCUMENTOS_CSV)
        docs.append(nuevo)
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html', alumnos=alumnos, documento={})

@app.route('/documentos/edit/<no_control>', methods=['GET','POST'])
@login_required
def documentos_edit(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d.get('No_Control') == no_control), None)
    if not doc:
        flash("Documento no encontrado", 'danger')
        return redirect(url_for('documentos_list'))
    if request.method == 'POST':
        for f in DOC_FIELDS:
            if f != 'No_Control':
                doc[f] = request.form.get(f,'').strip()
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html', alumnos=cargar_csv(ALUMNOS_CSV), documento=doc)

@app.route('/documentos/delete/<no_control>')
@login_required
def documentos_delete(no_control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV) if d.get('No_Control') != no_control]
    guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
    return redirect(url_for('documentos_list'))

# ── Generación de DOCX ───────────────────────────────────────────────────────
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a.get('No_Control') == no_control), None)
    doci    = next((d for d in docs    if d.get('No_Control') == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos para generar solicitud", 'warning')
        return redirect(url_for('dashboard'))
    ctx = {**alum, **doci}
    return _render_docx(TPL_SOLICITUD, ctx, f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    if num not in (1, 2, 3):
        flash("Bimestre inválido", 'warning')
        return redirect(url_for('dashboard'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a.get('No_Control') == no_control), None)
    doci    = next((d for d in docs    if d.get('No_Control') == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos para generar reporte bimestral", 'warning')
        return redirect(url_for('dashboard'))
    ctx = {**alum, **doci, 'reporte_no': num}
    return _render_docx(TPL_BIMESTRAL, ctx, f"Reporte_Bimestral_{no_control}_B{num}.docx")

@app.route('/generar_todos_bimestrales/<no_control>')
@login_required
def generar_todos_bimestrales(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a.get('No_Control') == no_control), None)
    doci    = next((d for d in docs    if d.get('No_Control') == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos para generar reportes", 'warning')
        return redirect(url_for('dashboard'))
    zip_buf = BytesIO()
    with ZipFile(zip_buf, 'w') as z:
        for num in (1, 2, 3):
            ctx = {**alum, **doci, 'reporte_no': num}
            doc = DocxTemplate(TPL_BIMESTRAL)
            doc.render(ctx)
            b = BytesIO()
            doc.save(b); b.seek(0)
            z.writestr(f"Reporte_Bimestral_{no_control}_B{num}.docx", b.read())
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        as_attachment=True,
        download_name=f"Reportes_Bimestrales_{no_control}.zip",
        mimetype='application/zip'
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a.get('No_Control') == no_control), None)
    doci    = next((d for d in docs    if d.get('No_Control') == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos para generar reporte final", 'warning')
        return redirect(url_for('dashboard'))
    ctx = {**alum, **doci}
    return _render_docx(TPL_FINAL, ctx, f"Reporte_Final_{no_control}.docx")

# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(error):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
