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
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos.csv')

# ── Plantillas DOCX ────────────────────────────────────────────────────────────
TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

# ── Helpers ────────────────────────────────────────────────────────────────────
def cargar_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def guardar_csv(path, rows, fieldnames):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, '') for k in fieldnames})

def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*a, **k)
    return w

def roles_required(*roles):
    def dec(f):
        @wraps(f)
        def w(*a, **k):
            if session.get('role') not in roles:
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return f(*a, **k)
        return w
    return dec

# ── Autenticación ─────────────────────────────────────────────────────────────
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
    if 'error' in result:
        flash(result.get('error_description','Error al autenticar'), 'danger')
        return redirect(url_for('login'))
    email = result['id_token_claims']['preferred_username'].lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403

    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u['correo'].lower()==email), None)
    if not perfil:
        perfil = {'correo': email, 'rol':'Alumno', 'area':''}

    session.update(user=perfil, role=perfil['rol'], name=result['id_token_claims'].get('name',''))
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login', _external=True)}")

# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# ── CRUD Usuarios, Áreas, Profesores y Alumnos (omitido para brevedad) ─────────

# ── CRUD Documentos ───────────────────────────────────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    documentos = cargar_csv(DOCUMENTOS_CSV)
    alumnos     = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_list.html', documentos=documentos, alumnos=alumnos)

@app.route('/documentos/new', methods=['GET','POST'])
@login_required
def documentos_new():
    alumnos = cargar_csv(ALUMNOS_CSV)
    if request.method=='POST':
        docs = cargar_csv(DOCUMENTOS_CSV)
        campos = []
        # Cargar encabezados actuales o definir por primera vez
        if docs:
            campos = list(docs[0].keys())
        else:
            campos = list(request.form.keys())
        nuevo = {k: request.form.get(k, '') for k in campos}
        # autocompletar datos alumno
        no = nuevo.get('No_Control')
        alumno = next((a for a in alumnos if a['No_Control']==no), {})
        for k in ['Apellido_Paterno','Apellido_Materno','Nombre','Carrera']:
            if k in alumno:
                nuevo[k] = alumno[k]
        docs.append(nuevo)
        guardar_csv(DOCUMENTOS_CSV, docs, campos)
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html', alumnos=alumnos, documento={})

@app.route('/documentos/edit/<no_control>', methods=['GET','POST'])
@login_required
def documentos_edit(no_control):
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alumnos = cargar_csv(ALUMNOS_CSV)
    doc     = next((d for d in docs if d['No_Control']==no_control), None)
    if not doc:
        flash('Documento no encontrado','danger')
        return redirect(url_for('documentos_list'))
    if request.method=='POST':
        for k in doc.keys():
            if k!='No_Control':
                doc[k] = request.form.get(k, '')
        # volver a autocompletar datos alumno
        alumno = next((a for a in alumnos if a['No_Control']==no_control), {})
        for k in ['Apellido_Paterno','Apellido_Materno','Nombre','Carrera']:
            if k in alumno:
                doc[k] = alumno[k]
        guardar_csv(DOCUMENTOS_CSV, docs, list(doc.keys()))
        return redirect(url_for('documentos_list'))
    return render_template('documentos_form.html', alumnos=alumnos, documento=doc)

@app.route('/documentos/delete/<no_control>')
@login_required
def documentos_delete(no_control):
    docs = [d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']!=no_control]
    campos = list(docs[0].keys()) if docs else ['No_Control']
    guardar_csv(DOCUMENTOS_CSV, docs, campos)
    return redirect(url_for('documentos_list'))

# ── Generación DOCX ──────────────────────────────────────────────────────────
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==no_control), {})
    ctx = {**alumno, **docinfo}
    doc = DocxTemplate(TPL_SOLICITUD); doc.render(ctx)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Solicitud_{no_control}.docx", as_attachment=True)

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==no_control), {})
    ctx = {**alumno, **docinfo, 'reporte_no': num}
    doc = DocxTemplate(TPL_BIMESTRAL); doc.render(ctx)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Bimestral_{no_control}_B{num}.docx", as_attachment=True)

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumno  = next((a for a in cargar_csv(ALUMNOS_CSV)    if a['No_Control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['No_Control']==no_control), {})
    ctx = {**alumno, **docinfo}
    doc = DocxTemplate(TPL_FINAL); doc.render(ctx)
    buf = BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(buf, download_name=f"Final_{no_control}.docx", as_attachment=True)

# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT',5000)))
