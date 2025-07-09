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
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecret')

# Azure AD configuration
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# CSV file paths (in a csv/ folder at project root)
CSV_DIR        = os.path.join(app.root_path, 'csv')
USERS_CSV      = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos.csv')

# Docx templates (in a plantillas/ folder at project root)
TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TEMPLATE_SOLI  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TEMPLATE_BIM   = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TEMPLATE_FINAL = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

# ── Helpers ────────────────────────────────────────────────────────────────────

def cargar_csv(path):
    """Load a CSV into a list of dicts, normalizing all keys to lowercase."""
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = []
        for row in reader:
            norm = {k.strip().lower(): v for k, v in row.items()}
            rows.append(norm)
        return rows

def guardar_csv(path, rows, fieldnames):
    """Save list of lowercase-keyed dicts to CSV with the given fieldnames."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, '') for k in fieldnames})

def normalize_context(row):
    """
    Given a dict with lowercase_underscore keys, return a new dict containing:
    - the original keys,
    - plus Pascal_Case keys for each (e.g. 'apellido_paterno' → 'Apellido_Paterno').
    """
    ctx = {}
    for k, v in row.items():
        ctx[k] = v
        parts = k.split('_')
        pascal = '_'.join(p.capitalize() for p in parts)
        ctx[pascal] = v
    return ctx

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

# ── Authentication ─────────────────────────────────────────────────────────────

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
    claims = result.get('id_token_claims', {})
    email = claims.get('preferred_username', '').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403

    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u.get('correo') == email), None)
    if not perfil:
        perfil = {'correo': email, 'rol': 'Alumno', 'area': ''}
    session['user'] = perfil
    session['role'] = perfil['rol']
    session['name'] = claims.get('name', '')
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', user=session['user'], role=session['role'])

# ── Usuarios CRUD (Administrador) ─────────────────────────────────────────────

@app.route('/usuarios')
@roles_required('Administrador')
def usuarios_list():
    usuarios = cargar_csv(USERS_CSV)
    return render_template('usuarios_list.html', usuarios=usuarios)

@app.route('/usuarios/new', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_new():
    if request.method == 'POST':
        usuarios = cargar_csv(USERS_CSV)
        usuarios.append({
            'correo': request.form['correo'].strip().lower(),
            'rol':    request.form['rol'],
            'area':   request.form['area']
        })
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    usuarios = cargar_csv(USERS_CSV)
    user = next((u for u in usuarios if u['correo'] == correo), None)
    if not user:
        flash('Usuario no encontrado','danger')
        return redirect(url_for('usuarios_list'))
    if request.method == 'POST':
        user['rol']  = request.form['rol']
        user['area'] = request.form['area']
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=user)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    usuarios = [u for u in cargar_csv(USERS_CSV) if u['correo'] != correo]
    guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ── Áreas CRUD (Administrador, Encargado) ────────────────────────────────────

@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    areas = cargar_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=areas)

@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    if request.method == 'POST':
        areas = cargar_csv(AREAS_CSV)
        areas.append({
            'id':        str(uuid.uuid4()),
            'nombre':    request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    areas = cargar_csv(AREAS_CSV)
    area = next((a for a in areas if a['id'] == id), None)
    if not area:
        flash('Área no encontrada','danger')
        return redirect(url_for('areas_list'))
    if request.method == 'POST':
        area['nombre']    = request.form['nombre']
        area['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    areas = [a for a in cargar_csv(AREAS_CSV) if a['id'] != id]
    guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ── Profesores CRUD (Administrador, Encargado) ───────────────────────────────

@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    profesores = cargar_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=profesores)

@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method == 'POST':
        profs = cargar_csv(PROFESORES_CSV)
        profs.append({
            'correo': request.form['correo'].strip().lower(),
            'nombre': request.form['nombre'],
            'area':   request.form['area']
        })
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    profs = cargar_csv(PROFESORES_CSV)
    prof = next((p for p in profs if p['correo'] == correo), None)
    if not prof:
        flash('Profesor no encontrado','danger')
        return redirect(url_for('profesores_list'))
    if request.method == 'POST':
        prof['nombre'] = request.form['nombre']
        prof['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    profs = [p for p in cargar_csv(PROFESORES_CSV) if p['correo'] != correo]
    guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ── Alumnos CRUD (Administrador, Encargado, Maestro) ────────────────────────

@app.route('/alumnos')
@login_required
def alumnos_list():
    alumnos = cargar_csv(ALUMNOS_CSV)
    role = session['role']
    user = session['user']
    if role == 'Maestro':
        alumnos = [a for a in alumnos if a.get('profesor','').lower() == user['correo']]
    elif role == 'Encargado':
        alumnos = [a for a in alumnos if a.get('area','') == user.get('area','')]
    return render_template('alumnos_list.html', alumnos=alumnos)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method == 'POST':
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            'no_control': request.form['no_control'],
            'nombre':     request.form['nombre'],
            'area':       request.form['area'],
            'profesor':   request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, rows, ['no_control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(no_control):
    rows = cargar_csv(ALUMNOS_CSV)
    alumno = next((a for a in rows if a['no_control'] == no_control), None)
    if not alumno:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('alumnos_list'))
    if request.method == 'POST':
        alumno['nombre']   = request.form['nombre']
        alumno['area']     = request.form['area']
        alumno['profesor'] = request.form['profesor']
        guardar_csv(ALUMNOS_CSV, rows, ['no_control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alumno)

@app.route('/alumnos/delete/<no_control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(no_control):
    rows = [a for a in cargar_csv(ALUMNOS_CSV) if a['no_control'] != no_control]
    guardar_csv(ALUMNOS_CSV, rows, ['no_control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))

# ── Document Generation ───────────────────────────────────────────────────────

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumno_raw  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['no_control'] == no_control), {})
    docinfo_raw = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control'] == no_control), {})
    context = {}
    context.update(normalize_context(alumno_raw))
    context.update(normalize_context(docinfo_raw))
    doc = DocxTemplate(TEMPLATE_SOLI)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf,
                     download_name=f"Solicitud_{no_control}.docx",
                     as_attachment=True)

@app.route('/generar_bimestral/<no_control>/<int:numero>')
@login_required
def generar_bimestral(no_control, numero):
    alumno_raw  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['no_control'] == no_control), {})
    docinfo_raw = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control'] == no_control), {})
    context = {}
    context.update(normalize_context(alumno_raw))
    context.update(normalize_context(docinfo_raw))
    context['reporte_no'] = numero
    doc = DocxTemplate(TEMPLATE_BIM)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Bimestral_{no_control}_Bim{numero}.docx",
                     as_attachment=True)

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumno_raw  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['no_control'] == no_control), {})
    docinfo_raw = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control'] == no_control), {})
    context = {}
    context.update(normalize_context(alumno_raw))
    context.update(normalize_context(docinfo_raw))
    doc = DocxTemplate(TEMPLATE_FINAL)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Final_{no_control}.docx",
                     as_attachment=True)

# ── Error handlers ────────────────────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

# ── Run ────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
