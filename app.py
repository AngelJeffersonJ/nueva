import os
import csv
from io import BytesIO
from functools import wraps
from flask import Flask, render_template, redirect, url_for, session, request, flash, send_file
from dotenv import load_dotenv
import msal
from docxtpl import DocxTemplate

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecret')

# ── Configuración Azure AD ─────────────────────────────
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = os.getenv('REDIRECT_URI')
SCOPE = ['User.Read']

# ── Rutas de CSV ───────────────────────────────────────
CSV_DIR = os.path.join(app.root_path, 'csv')
USERS_CSV = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV = os.path.join(CSV_DIR, 'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos_completo.csv')

# ── Plantillas DOCX ────────────────────────────────────
TPL_DIR = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

# ── Utilidades ─────────────────────────────────────
def cargar_csv(path):
    rows = []
    if os.path.isfile(path):
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append({k.strip(): v.strip() for k, v in row.items()})
    return rows

def guardar_csv(path, rows, fieldnames):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

# ── Decoradores ────────────────────────────────────
def login_required(f):
    @wraps(f)
    def w(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return w

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*args, **kwargs):
            if session.get('role') not in roles:
                flash("Acceso denegado", "danger")
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return w
    return deco

# ── Autenticación Azure AD ─────────────────────────
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
        flash(result.get('error_description', 'Error al autenticar'), 'danger')
        return redirect(url_for('login'))
    claims = result['id_token_claims']
    email = claims.get('preferred_username', '').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403
    usuarios = cargar_csv(USERS_CSV)
    perfil = next((u for u in usuarios if u.get('correo', '').lower() == email), None)
    if not perfil:
        perfil = {'correo': email, 'rol': 'Alumno', 'area': ''}
    session.update(user=perfil, role=perfil['rol'], name=claims.get('name', ''))
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login', _external=True)}"
    )

# ── Dashboard ──────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# ── CRUD Usuarios/Áreas/Profesores/Alumnos (omitir por brevedad; igual que antes) ──
# ...

# ── Documentos ─────────────────────────────────────
@app.route('/documentos')
@login_required
def documentos_list():
    docs = cargar_csv(DOCUMENTOS_CSV)
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('documentos_list.html', documentos=docs, alumnos=alumnos)

@app.route('/documentos/solicitud/<no_control>', methods=['GET', 'POST'])
@login_required
def documentos_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control'] == no_control), None)
    if not doc:
        flash("Documento no encontrado", "danger")
        return redirect(url_for('documentos_list'))
    # lógica de formulario
    return render_template('solicitud_form.html', documento=doc, no_control=no_control)

@app.route('/documentos/bimestral/<no_control>/<int:num>', methods=['GET', 'POST'])
@login_required
def documentos_bimestral(no_control, num):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control'] == no_control), None)
    if not doc:
        flash("Documento no encontrado", "danger")
        return redirect(url_for('documentos_list'))
    # lógica de formulario
    return render_template('bimestral_form.html', documento=doc, no_control=no_control, num=num)

@app.route('/documentos/final/<no_control>', methods=['GET', 'POST'])
@login_required
def documentos_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc = next((d for d in docs if d['No_Control'] == no_control), None)
    if not doc:
        flash("Documento no encontrado", "danger")
        return redirect(url_for('documentos_list'))
    # lógica de formulario
    return render_template('final_form.html', documento=doc, no_control=no_control)

# ── Generar DOCX ──────────────────────────────────
def _render_docx(template_path, context, filename):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf, as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento", "danger")
        return redirect(url_for('dashboard'))

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs = cargar_csv(DOCUMENTOS_CSV)
    alum = next((a for a in alumnos if a['No_Control'] == no_control), None)
    doci = next((d for d in docs if d['No_Control'] == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(TPL_SOLICITUD, {**alum, **doci}, f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs = cargar_csv(DOCUMENTOS_CSV)
    alum = next((a for a in alumnos if a['No_Control'] == no_control), None)
    doci = next((d for d in docs if d['No_Control'] == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_BIMESTRAL, {**alum, **doci, 'reporte_no': num},
        f"Bimestral_{no_control}_B{num}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs = cargar_csv(DOCUMENTOS_CSV)
    alum = next((a for a in alumnos if a['No_Control'] == no_control), None)
    doci = next((d for d in docs if d['No_Control'] == no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    doci['Nombre_Estudiante'] = doci.get('Nombre', alum.get('Nombre', ''))
    return _render_docx(TPL_FINAL, {**alum, **doci}, f"Final_{no_control}.docx")

# ── Error 404 ─────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
