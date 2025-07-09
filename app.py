import os
import csv
import uuid
from io import BytesIO
from functools import wraps

from flask import Flask, render_template, redirect, url_for, session, request, flash, send_file
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
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        return [{k.strip().lower(): v for k, v in row.items()} for row in reader]

def guardar_csv(path, rows, fieldnames):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*a, **k)
    return w

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            if session.get('role') not in roles:
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return f(*a, **k)
        return w
    return deco

# ─── Authentication ─────────────────────────────────────────────────────────────

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
        flash(result.get('error_description','Error al autenticar'),'danger')
        return redirect(url_for('login'))
    claims = result.get('id_token_claims',{})
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403
    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u['correo']==email), None)
    if not perfil:
        perfil = {'correo': email, 'rol': 'Alumno', 'area': ''}
    session['user'] = perfil
    session['role'] = perfil['rol']
    session['name'] = claims.get('name','')
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login',_external=True)}"
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
    return render_template('usuarios_list.html', usuarios=cargar_csv(USERS_CSV))

@app.route('/usuarios/new', methods=['GET','POST'])
@roles_required('Administrador')
def usuarios_new():
    if request.method=='POST':
        u = cargar_csv(USERS_CSV)
        u.append({
            'correo': request.form['correo'].lower().strip(),
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
        flash('Usuario no encontrado','danger')
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        user['rol']  = request.form['rol']
        user['area'] = request.form['area']
        guardar_csv(USERS_CSV, u, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=user)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    u = [x for x in cargar_csv(USERS_CSV) if x['correo']!=correo]
    guardar_csv(USERS_CSV, u, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ─── Áreas CRUD ───────────────────────────────────────────────────────────────

@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    return render_template('areas_list.html', areas=cargar_csv(AREAS_CSV))

@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    if request.method=='POST':
        a = cargar_csv(AREAS_CSV)
        a.append({
            'id':        str(uuid.uuid4()),
            'nombre':    request.form['nombre'],
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
        flash('Área no encontrada','danger')
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        area['nombre']    = request.form['nombre']
        area['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    a = [x for x in cargar_csv(AREAS_CSV) if x['id']!=id]
    guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ─── Profesores CRUD ──────────────────────────────────────────────────────────

@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    return render_template('profesores_list.html', profesores=cargar_csv(PROFESORES_CSV))

@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method=='POST':
        p = cargar_csv(PROFESORES_CSV)
        p.append({
            'correo': request.form['correo'].lower().strip(),
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
        flash('Profesor no encontrado','danger')
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        prof['nombre'] = request.form['nombre']
        prof['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    p = [x for x in cargar_csv(PROFESORES_CSV) if x['correo']!=correo]
    guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
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
        a = cargar_csv(ALUMNOS_CSV)
        a.append({
            'no_control': request.form['no_control'],
            'nombre':     request.form['nombre'],
            'area':       request.form['area'],
            'profesor':   request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, a, ['no_control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(no_control):
    a = cargar_csv(ALUMNOS_CSV)
    alum = next((x for x in a if x['no_control']==no_control), None)
    if not alum:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        alum['nombre']   = request.form['nombre']
        alum['area']     = request.form['area']
        alum['profesor'] = request.form['profesor']
        guardar_csv(ALUMNOS_CSV, a, ['no_control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<no_control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(no_control):
    a = [x for x in cargar_csv(ALUMNOS_CSV) if x['no_control']!=no_control]
    guardar_csv(ALUMNOS_CSV, a, ['no_control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))

# ─── Documentos CRUD ──────────────────────────────────────────────────────────

@app.route('/documentos')
@login_required
def documentos_list():
    docs    = cargar_csv(DOCUMENTOS_CSV)
    # mapeo para que tu template use d.No_Control
    for d in docs:
        d['No_Control'] = d.get('no_control','')
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_list.html',
                           documentos=docs,
                           alumnos=alumnos)

@app.route('/documentos/new', methods=['GET','POST'])
@login_required
def documentos_new():
    if request.method=='POST':
        docs = cargar_csv(DOCUMENTOS_CSV)
        data = {k.lower(): v for k, v in request.form.to_dict(flat=True).items()}
        docs.append(data)
        guardar_csv(DOCUMENTOS_CSV, docs, list(data.keys()))
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_form.html',
                           documento={},
                           alumnos=alumnos)

@app.route('/documentos/edit/<No_Control>', methods=['GET','POST'])
@login_required
def documentos_edit(No_Control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d.get('no_control')==No_Control), None)
    if not doc:
        flash('Documento no encontrado','danger')
        return redirect(url_for('documentos_list'))
    if request.method=='POST':
        updated = {k.lower(): v for k, v in request.form.to_dict(flat=True).items()}
        for i, d in enumerate(docs):
            if d.get('no_control')==No_Control:
                docs[i] = updated
                break
        guardar_csv(DOCUMENTOS_CSV, docs, list(updated.keys()))
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    # vuelve a exponer No_Control para el form
    doc['No_Control'] = doc.get('no_control','')
    return render_template('documentos_form.html',
                           documento=doc,
                           alumnos=alumnos)

@app.route('/documentos/delete/<No_Control>')
@login_required
def documentos_delete(No_Control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV)
            if d.get('no_control')!=No_Control]
    if docs:
        guardar_csv(DOCUMENTOS_CSV, docs, list(docs[0].keys()))
    else:
        os.remove(DOCUMENTOS_CSV)
    return redirect(url_for('documentos_list'))

# ─── Generación de DOCX ──────────────────────────────────────────────────────

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)
                    if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV)
                    if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo}
    doc = DocxTemplate(TEMPLATE_SOLI); buf=BytesIO()
    doc.render(ctx); doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Solicitud_{no_control}.docx",
                     as_attachment=True)

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control,num):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)
                    if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV)
                    if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo, 'reporte_no': num}
    doc = DocxTemplate(TEMPLATE_BIM); buf=BytesIO()
    doc.render(ctx); doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Bimestral_{no_control}_Bim{num}.docx",
                     as_attachment=True)

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)
                    if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV)
                    if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo}
    doc = DocxTemplate(TEMPLATE_FINAL); buf=BytesIO()
    doc.render(ctx); doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Final_{no_control}.docx",
                     as_attachment=True)

# ─── Error handler ────────────────────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(debug=True,
            host='0.0.0.0',
            port=int(os.getenv('PORT', 5000)))
