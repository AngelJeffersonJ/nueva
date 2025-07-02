from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import os, msal, csv, uuid, zipfile
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

# ─── Azure AD Config (Render injecta estas) ─────────────────────────────
CLIENT_ID     = os.environ.get('CLIENT_ID', 'tu_client_id')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET', 'tu_client_secret')
TENANT_ID     = os.environ.get('TENANT_ID', 'tu_tenant_id')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.environ.get('REDIRECT_URI', 'http://localhost:5000/getAToken')
SCOPE         = ['User.Read']

# ─── Rutas y constantes ───────────────────────────────────────────────────
BASE_PATH       = os.path.dirname(__file__)
CSV_DIR         = os.path.join(BASE_PATH, os.environ.get('CSV_DIR', 'csv'))
os.makedirs(CSV_DIR, exist_ok=True)

USUARIOS_CSV    = os.path.join(CSV_DIR, 'usuarios.csv')
PROFESORES_CSV  = os.path.join(CSV_DIR, 'profesores.csv')
AREAS_CSV       = os.path.join(CSV_DIR, 'areas.csv')
ALUMNOS_CSV     = os.path.join(CSV_DIR, 'alumnos.csv')

TEMPLATE_SOLICITUD   = os.path.join(BASE_PATH, 'plantillas', 'plantilla_solicitud_completa.docx')
TEMPLATE_BIMESTRAL   = os.path.join(BASE_PATH, 'plantillas', 'Reporte_Bimestral_Plantilla.docx')
TEMPLATE_FINAL       = os.path.join(BASE_PATH, 'plantillas', 'Reporte_Final_Lleno.docx')
OUTPUT_DIR           = os.path.join(BASE_PATH, 'documentos_generados')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── Utilidades CSV ──────────────────────────────────────────────────────
def cargar_csv(path):
    """Devuelve lista de dicts (puede tener campos extra)."""
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fieldnames):
    """
    Escribe sólo las columnas en `fieldnames`,
    ignorando cualquier clave extra en cada row dict.
    """
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for row in rows:
            filtered = {k: row.get(k, '') for k in fieldnames}
            w.writerow(filtered)

# ─── Helpers de autenticación y roles ────────────────────────────────────
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get('correo','').strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get('user')
    return bool(u and u.get('rol') in roles)

# ─── Flujo MSAL / Login / Logout ─────────────────────────────────────────
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    if not code:
        flash("Código de autorización faltante", "danger")
        return redirect(url_for('login'))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if 'error' in result:
        return f"Error MSAL: { result.get('error_description') }", 500

    claims = result.get('id_token_claims', {})
    email  = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {'correo':email,'rol':'Alumno','area':''}
    session['user'] = {
        'correo': perfil['correo'],
        'rol':    perfil['rol'],
        'area':   perfil.get('area',''),
        'name':   claims.get('name','')
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ─── Dashboard ───────────────────────────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    # Para la tabla de generación masiva
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template('dashboard.html',
                           usuario=session['user'],
                           alumnos=alumnos)

# ─── CRUD de Usuarios ─────────────────────────────────────────────────────
@app.route('/usuarios')
def usuarios_list():
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(USUARIOS_CSV)
    return render_template('usuarios_list.html', usuarios=filas)

@app.route('/usuarios/new', methods=['GET','POST'])
def usuarios_new():
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        filas = cargar_csv(USUARIOS_CSV)
        nueva = {
            'correo': request.form['correo'],
            'rol':    request.form['rol'],
            'area':   request.form['area']
        }
        filas.append(nueva)
        guardar_csv(USUARIOS_CSV, filas, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
def usuarios_edit(correo):
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(USUARIOS_CSV)
    u = next((r for r in filas if r['correo']==correo), None)
    if not u: return "No existe", 404
    if request.method=='POST':
        u['rol']  = request.form['rol']
        u['area'] = request.form['area']
        guardar_csv(USUARIOS_CSV, filas, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=u)

@app.route('/usuarios/delete/<correo>', methods=['POST'])
def usuarios_delete(correo):
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    filas = [r for r in cargar_csv(USUARIOS_CSV) if r['correo']!=correo]
    guardar_csv(USUARIOS_CSV, filas, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ─── CRUD de Profesores ───────────────────────────────────────────────────
@app.route('/profesores')
def profesores_list():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=filas)

@app.route('/profesores/new', methods=['GET','POST'])
def profesores_new():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        filas = cargar_csv(PROFESORES_CSV)
        filas.append({
            'correo': request.form['correo'],
            'nombre': request.form['nombre'],
            'area':   request.form['area']
        })
        guardar_csv(PROFESORES_CSV, filas, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
def profesores_edit(correo):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(PROFESORES_CSV)
    p = next((r for r in filas if r['correo']==correo), None)
    if not p: return "No existe", 404
    if request.method=='POST':
        p['nombre'], p['area'] = request.form['nombre'], request.form['area']
        guardar_csv(PROFESORES_CSV, filas, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=p)

@app.route('/profesores/delete/<correo>', methods=['POST'])
def profesores_delete(correo):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = [r for r in cargar_csv(PROFESORES_CSV) if r['correo']!=correo]
    guardar_csv(PROFESORES_CSV, filas, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ─── CRUD de Áreas ───────────────────────────────────────────────────────
@app.route('/areas')
def areas_list():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=filas)

@app.route('/areas/new', methods=['GET','POST'])
def areas_new():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        filas = cargar_csv(AREAS_CSV)
        filas.append({
            'id':        str(uuid.uuid4()),
            'nombre':    request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, filas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
def areas_edit(id):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(AREAS_CSV)
    a = next((r for r in filas if r['id']==id), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre'], a['encargado'] = request.form['nombre'], request.form['encargado']
        guardar_csv(AREAS_CSV, filas, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=a)

@app.route('/areas/delete/<id>', methods=['POST'])
def areas_delete(id):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    filas = [r for r in cargar_csv(AREAS_CSV) if r['id']!=id]
    guardar_csv(AREAS_CSV, filas, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ─── CRUD de Alumnos ──────────────────────────────────────────────────────
# Sólo los 4 campos: correo, nombre, area, profesor
ALUMNOS_FIELDS = ['correo','nombre','area','profesor']

@app.route('/alumnos')
def alumnos_list():
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(ALUMNOS_CSV)
    # Si eres Maestro, filtras por tu correo:
    if session['user']['rol']=='Maestro':
        filas = [r for r in filas if r.get('profesor')==session['user']['correo']]
    # Si eres Encargado, filtras por tu área:
    if session['user']['rol']=='Encargado':
        filas = [r for r in filas if r.get('area')==session['user']['area']]
    return render_template('alumnos_list.html', alumnos=filas)

@app.route('/alumnos/new', methods=['GET','POST'])
def alumnos_new():
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        filas = cargar_csv(ALUMNOS_CSV)
        nueva = {
            'correo':   request.form['correo'],
            'nombre':   request.form['nombre'],
            'area':     request.form['area'],
            'profesor': request.form['profesor']
        }
        filas.append(nueva)
        # Aquí filtramos para que sólo se escriban las 4 columnas
        guardar_csv(ALUMNOS_CSV, filas, ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<correo>', methods=['GET','POST'])
def alumnos_edit(correo):
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(ALUMNOS_CSV)
    a = next((r for r in filas if r['correo']==correo), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre'], a['area'], a['profesor'] = (
            request.form['nombre'],
            request.form['area'],
            request.form['profesor']
        )
        guardar_csv(ALUMNOS_CSV, filas, ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=a)

@app.route('/alumnos/delete/<correo>', methods=['POST'])
def alumnos_delete(correo):
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    filas = [r for r in cargar_csv(ALUMNOS_CSV) if r['correo']!=correo]
    guardar_csv(ALUMNOS_CSV, filas, ALUMNOS_FIELDS)
    return redirect(url_for('alumnos_list'))

# ─── Generación de documentos ─────────────────────────────────────────────
def _make_doc(template_path, context, output_name):
    doc = DocxTemplate(template_path)
    doc.render(context)
    out_path = os.path.join(OUTPUT_DIR, output_name)
    doc.save(out_path)
    return out_path

@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    df = pd.read_csv(ALUMNOS_CSV)
    alumno = df[df['No_Control']==int(no_control)].squeeze().to_dict()
    out = _make_doc(TEMPLATE_SOLICITUD, alumno, f"Solicitud_{no_control}.docx")
    return send_file(out, as_attachment=True)

@app.route('/generar_bimestral/<int:periodo>/<no_control>')
def generar_bimestral(periodo, no_control):
    df = pd.read_csv(ALUMNOS_CSV)
    row = df[df['No_Control']==int(no_control)].squeeze().to_dict()
    # Ajusta nombres de campos según plantilla bim. (ejemplo para periodo 1)
    # ... aquí podrías inyectar día1, mes1, año1, Actividad_1..8, etc.
    out = _make_doc(TEMPLATE_BIMESTRAL, row, f"Reporte_Bim{periodo}_{no_control}.docx")
    return send_file(out, as_attachment=True)

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    df = pd.read_csv(ALUMNOS_CSV)
    row = df[df['No_Control']==int(no_control)].squeeze().to_dict()
    out = _make_doc(TEMPLATE_FINAL, row, f"Reporte_Final_{no_control}.docx")
    return send_file(out, as_attachment=True)

# ─── Manejo 404 ───────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',5000)), debug=True)
