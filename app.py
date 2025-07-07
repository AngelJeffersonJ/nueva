import os
import uuid
import csv
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import msal
from docxtpl import DocxTemplate

# ── App setup ──────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = os.getenv('FLASK_SECRET', 'super-secret-key')

# ── Azure AD Config ───────────────────────────────────────────────────────────
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.getenv('REDIRECT_URI')  # e.g. 'https://tu-dominio.com/getAToken'
SCOPE         = ['User.Read']

# ── CSV paths ─────────────────────────────────────────────────────────────────
BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
CSV_DIR         = os.path.join(BASE_DIR, 'csv')
USUARIOS_CSV    = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV       = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV  = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV     = os.path.join(CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV  = os.path.join(CSV_DIR, 'documentos.csv')

# ── Plantillas ────────────────────────────────────────────────────────────────
PLANTILLAS_DIR      = os.path.join(BASE_DIR, 'plantillas')
TPL_SOLICITUD       = os.path.join(PLANTILLAS_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL       = os.path.join(PLANTILLAS_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL           = os.path.join(PLANTILLAS_DIR, 'Reporte_Final_Plantilla.docx')

# ── Helpers para CSV ───────────────────────────────────────────────────────────
def cargar_csv(path):
    """Carga un CSV y retorna lista de dicts; recarga en cada llamada."""
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def guardar_csv(path, rows, fieldnames):
    """Guarda lista de dicts en CSV, sobreescribiendo."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

def find_user(email):
    """Busca en usuarios.csv, devuelve dict o None."""
    for u in cargar_csv(USUARIOS_CSV):
        if u['correo'].strip().lower() == email.lower():
            return u
    return None

# ── Decorators ─────────────────────────────────────────────────────────────────
def login_required(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper

def role_required(*roles):
    from functools import wraps
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if session.get('role') not in roles:
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return decorator

# ── Autenticación ──────────────────────────────────────────────────────────────
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
        flash(result.get('error_description','Error de autenticación'), 'danger')
        return redirect(url_for('dashboard'))

    email = result.get('id_token_claims',{}).get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403

    perfil = find_user(email) or {'correo': email, 'rol': 'Alumno', 'area': ''}
    session['user'] = perfil['correo']
    session['role'] = perfil['rol']
    session['area'] = perfil.get('area','')
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# ── CRUD Alumnos ───────────────────────────────────────────────────────────────
@app.route('/alumnos')
@login_required
def alumnos_list():
    rows = cargar_csv(ALUMNOS_CSV)
    return render_template('alumnos_list.html', alumnos=rows)

@app.route('/alumnos/new', methods=['GET','POST'])
@login_required
@role_required('Administrador')
def alumnos_new():
    if request.method == 'POST':
        rows = cargar_csv(ALUMNOS_CSV)
        nuevo = {
            'no_control': request.form['no_control'],
            'nombre':     request.form['nombre'],
            'correo':     request.form['correo'],
            'carrera':    request.form['carrera']
        }
        rows.append(nuevo)
        guardar_csv(ALUMNOS_CSV, rows,
                    fieldnames=['no_control','nombre','correo','carrera'])
        flash('Alumno agregado', 'success')
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
@login_required
@role_required('Administrador')
def alumnos_edit(no_control):
    rows = cargar_csv(ALUMNOS_CSV)
    alumno = next((a for a in rows if a['no_control']==no_control), None)
    if not alumno:
        flash('Alumno no encontrado', 'danger')
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        alumno['nombre']  = request.form['nombre']
        alumno['correo']  = request.form['correo']
        alumno['carrera'] = request.form['carrera']
        guardar_csv(ALUMNOS_CSV, rows,
                    fieldnames=['no_control','nombre','correo','carrera'])
        flash('Alumno actualizado', 'success')
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alumno)

@app.route('/alumnos/delete/<no_control>')
@login_required
@role_required('Administrador')
def alumnos_delete(no_control):
    rows = [a for a in cargar_csv(ALUMNOS_CSV) if a['no_control']!=no_control]
    guardar_csv(ALUMNOS_CSV, rows,
                fieldnames=['no_control','nombre','correo','carrera'])
    flash('Alumno eliminado', 'success')
    return redirect(url_for('alumnos_list'))

# ── Generación de documentos ──────────────────────────────────────────────────
def generate_doc(template, context, filename):
    doc = DocxTemplate(template)
    doc.render(context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return send_file(bio, download_name=filename, as_attachment=True)

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumno = next((a for a in cargar_csv(ALUMNOS_CSV) if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo}
    return generate_doc(TPL_SOLICITUD, ctx, f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:numero>')
@login_required
def generar_bimestral(no_control, numero):
    alumno = next((a for a in cargar_csv(ALUMNOS_CSV) if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo, 'numero': numero}
    return generate_doc(
        TPL_BIMESTRAL,
        ctx,
        f"Reporte_Bimestral_{numero}_{no_control}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumno = next((a for a in cargar_csv(ALUMNOS_CSV) if a['no_control']==no_control), {})
    docinfo = next((d for d in cargar_csv(DOCUMENTOS_CSV) if d['no_control']==no_control), {})
    ctx = {**alumno, **docinfo}
    return generate_doc(TPL_FINAL, ctx, f"Reporte_Final_{no_control}.docx")

# ── Run ────────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
