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

# ─── Setup ─────────────────────────────────────────────────────────────────────

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecret')

# Azure AD
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# CSV paths
BASE_DIR        = app.root_path
CSV_DIR         = os.path.join(BASE_DIR, 'csv')
USERS_CSV       = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV       = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV  = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV     = os.path.join(CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV  = os.path.join(CSV_DIR, 'documentos.csv')

# DOCX templates
TPL_DIR        = os.path.join(BASE_DIR, 'plantillas')
TEMPLATE_SOLI  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TEMPLATE_BIM   = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TEMPLATE_FINAL = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')


# ─── Helpers ───────────────────────────────────────────────────────────────────

def cargar_csv(path):
    """Load a CSV into a list of dicts."""
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))


def guardar_csv(path, rows, fieldnames):
    """Save list of dicts to CSV with the given fieldnames."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper


def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if session.get('role') not in roles:
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return deco


# ─── Auth ──────────────────────────────────────────────────────────────────────

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

    users = cargar_csv(USERS_CSV)
    perfil = next((u for u in users if u['correo']==email), None)
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
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )


# ─── Dashboard ─────────────────────────────────────────────────────────────────

@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')


# ─── Usuarios CRUD ─────────────────────────────────────────────────────────────

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
    user = next((u for u in usuarios if u['correo']==correo), None)
    if not user:
        flash('Usuario no encontrado','danger')
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        user['rol']  = request.form['rol']
        user['area'] = request.form['area']
        guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=user)


@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    usuarios = [u for u in cargar_csv(USERS_CSV) if u['correo']!=correo]
    guardar_csv(USERS_CSV, usuarios, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))


# ─── Áreas CRUD ───────────────────────────────────────────────────────────────

@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    areas = cargar_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=areas)


@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    if request.method=='POST':
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
    area = next((a for a in areas if a['id']==id), None)
    if not area:
        flash('Área no encontrada','danger')
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        area['nombre']    = request.form['nombre']
        area['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)


@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    areas = [a for a in cargar_csv(AREAS_CSV) if a['id']!=id]
    guardar_csv(AREAS_CSV, areas, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))


# ─── Profesores CRUD ──────────────────────────────────────────────────────────

@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    profesores = cargar_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=profesores)


@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method=='POST':
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
    prof = next((p for p in profs if p['correo']==correo), None)
    if not prof:
        flash('Profesor no encontrado','danger')
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        prof['nombre'] = request.form['nombre']
        prof['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)


@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    profs = [p for p in cargar_csv(PROFESORES_CSV) if p['correo']!=correo]
    guardar_csv(PROFESORES_CSV, profs, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))


# ─── Alumnos CRUD ─────────────────────────────────────────────────────────────

@app.route('/alumnos')
@login_required
def alumnos_list():
    alumnos = cargar_csv(ALUMNOS_CSV)
    role    = session['role']
    user    = session['user']
    if role=='Maestro':
        alumnos = [a for a in alumnos if a['profesor']==user['correo']]
    elif role=='Encargado':
        alumnos = [a for a in alumnos if a['area']==user['area']]
    return render_template('alumnos_list.html', alumnos=alumnos)


@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method=='POST':
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            'No_Control': request.form['No_Control'],
            'nombre':     request.form['nombre'],
            'area':       request.form['area'],
            'profesor':   request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, rows, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})


@app.route('/alumnos/edit/<No_Control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(No_Control):
    rows = cargar_csv(ALUMNOS_CSV)
    alum = next((a for a in rows if a['No_Control']==No_Control), None)
    if not alum:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        alum['nombre']   = request.form['nombre']
        alum['area']     = request.form['area']
        alum['profesor'] = request.form['profesor']
        guardar_csv(ALUMNOS_CSV, rows, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)


@app.route('/alumnos/delete/<No_Control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(No_Control):
    rows = [a for a in cargar_csv(ALUMNOS_CSV) if a['No_Control']!=No_Control]
    guardar_csv(ALUMNOS_CSV, rows, ['No_Control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))


# ─── Documentos CRUD ──────────────────────────────────────────────────────────

@app.route('/documentos')
@login_required
def documentos_list():
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_list.html', documentos=docs, alumnos=alumnos)


@app.route('/documentos/new', methods=['GET','POST'])
@login_required
def documentos_new():
    if request.method=='POST':
        docs = cargar_csv(DOCUMENTOS_CSV)
        data = request.form.to_dict(flat=True)
        docs.append(data)
        # infer header from first record
        guardar_csv(DOCUMENTOS_CSV, docs, list(docs[0].keys()))
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_form.html', documento={}, alumnos=alumnos)


@app.route('/documentos/edit/<No_Control>', methods=['GET','POST'])
@login_required
def documentos_edit(No_Control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d['No_Control']==No_Control), {})
    if request.method=='POST':
        data = request.form.to_dict(flat=True)
        for i, d in enumerate(docs):
            if d['No_Control']==No_Control:
                docs[i] = data
                break
        guardar_csv(DOCUMENTOS_CSV, docs, list(docs[0].keys()))
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_form.html', documento=doc, alumnos=alumnos)


@app.route('/documentos/delete/<No_Control>')
@login_required
def documentos_delete(No_Control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']!=No_Control]
    if docs:
        guardar_csv(DOCUMENTOS_CSV, docs, list(docs[0].keys()))
    else:
        # remove file if empty
        os.remove(DOCUMENTOS_CSV)
    return redirect(url_for('documentos_list'))


# ─── Generación de DOCX ──────────────────────────────────────────────────────

@app.route('/generar_solicitud/<No_Control>')
@login_required
def generar_solicitud(No_Control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==No_Control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==No_Control), {})
    context = {**alumno, **docinfo}
    doc = DocxTemplate(TEMPLATE_SOLI)
    doc.render(context)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Solicitud_{No_Control}.docx", as_attachment=True)


@app.route('/generar_bimestral/<No_Control>/<int:num>')
@login_required
def generar_bimestral(No_Control, num):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==No_Control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==No_Control), {})
    context = {**alumno, **docinfo, 'reporte_no': num}
    doc = DocxTemplate(TEMPLATE_BIM)
    doc.render(context)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Reporte_Bimestral_{No_Control}_Bim{num}.docx", as_attachment=True)


@app.route('/generar_final/<No_Control>')
@login_required
def generar_final(No_Control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==No_Control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==No_Control), {})
    context = {**alumno, **docinfo}
    doc = DocxTemplate(TEMPLATE_FINAL)
    doc.render(context)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Reporte_Final_{No_Control}.docx", as_attachment=True)


# ─── Error handler ────────────────────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404


# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
