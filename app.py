# app.py
# -*- coding: utf-8 -*-
from flask import Flask, render_template, session, redirect, url_for, request, flash, send_file
import os, csv, uuid, msal, zipfile
from io import BytesIO
from docxtpl import DocxTemplate

app = Flask(__name__)
# Lee la clave secreta de Flask desde la env var, o usa 'clave_segura' por defecto
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

# ─── Configuración de Azure AD (ya definidas en Render o en .env) ───────────
CLIENT_ID     = os.environ.get('CLIENT_ID',     'c306c8d3-68dc-4110-b5fb-771b942c10db')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET', 'CEX8Q~txlNOmiN4PAZKbV7matISmRV1HREkvbcMS')
TENANT_ID     = os.environ.get('TENANT_ID',     '63de1475-1a48-4463-aff2-b2581f2a972e')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.environ.get('REDIRECT_URI',  'https://sistema-documental.onrender.com/getAToken')
SCOPE         = ['User.Read']

# ─── Rutas a los CSV ─────────────────────────────────────────────────────────
CSV_DIR        = os.environ.get('CSV_DIR', 'csv')
USUARIOS_CSV   = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos.csv')

# ─── Plantillas Word y carpeta de salida ─────────────────────────────────────
BASE_PATH            = os.path.dirname(__file__)
PLANTILLA_SOLICITUD  = os.path.join(BASE_PATH, 'plantillas', 'plantilla_solicitud_completa.docx')
PLANTILLA_BIMESTRAL  = os.path.join(BASE_PATH, 'plantillas', 'Reporte_Bimestral_Plantilla.docx')
PLANTILLA_FINAL      = os.path.join(BASE_PATH, 'plantillas', 'Reporte_Final_Plantilla.docx')
OUTPUT_PATH          = os.path.join(BASE_PATH, 'documentos_generados')
os.makedirs(OUTPUT_PATH, exist_ok=True)

# ─── Funciones genéricas de CSV ───────────────────────────────────────────────
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

# ─── Helpers de autenticación y autorización ─────────────────────────────────
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get('correo','').strip().lower() == email.lower():
            return {'correo': u['correo'], 'rol': u['rol'], 'area': u.get('area','')}
    return None

def validar_acceso(roles):
    u = session.get('user')
    return bool(u and u.get('rol') in roles)

# ─── Flujo de login con MSAL ─────────────────────────────────────────────────
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
    if not code:
        flash('Código de autorización faltante', 'danger')
        return redirect(url_for('login'))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if 'error' in result:
        flash(result.get('error_description','Error MSAL'), 'danger')
        return redirect(url_for('login'))

    claims = result.get('id_token_claims', {})
    email  = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas @aguascalientes.tecnm.mx permitidas", 403

    perfil = usuario_desde_csv(email) or {'correo':email,'rol':'Alumno','area':''}
    session['user'] = {
        'correo': perfil['correo'],
        'rol'   : perfil['rol'],
        'area'  : perfil['area'],
        'name'  : claims.get('name','')
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ─── Dashboard ────────────────────────────────────────────────────────────────
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

# ─── CRUD Usuarios ────────────────────────────────────────────────────────────
@app.route('/usuarios')
def usuarios_list():
    if not validar_acceso(['Administrador']):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    return render_template('usuarios_list.html', usuarios=cargar_csv(USUARIOS_CSV))

@app.route('/usuarios/new', methods=['GET','POST'])
def usuarios_new():
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(USUARIOS_CSV)
        rows.append({
            'correo': request.form['correo'],
            'rol'   : request.form['rol'],
            'area'  : request.form['area']
        })
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html')

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
def usuarios_edit(correo):
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(USUARIOS_CSV)
    u = next((x for x in rows if x['correo']==correo), None)
    if not u: return "No existe",404
    if request.method=='POST':
        u['rol']  = request.form['rol']
        u['area'] = request.form['area']
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=u)

@app.route('/usuarios/delete/<correo>', methods=['POST'])
def usuarios_delete(correo):
    if not validar_acceso(['Administrador']):
        return redirect(url_for('dashboard'))
    rows = [x for x in cargar_csv(USUARIOS_CSV) if x['correo']!=correo]
    guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ─── CRUD Áreas ───────────────────────────────────────────────────────────────
@app.route('/areas')
def areas_list():
    if not validar_acceso(['Administrador','Encargado']):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    return render_template('areas_list.html', areas=cargar_csv(AREAS_CSV))

@app.route('/areas/new', methods=['GET','POST'])
def areas_new():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            'id'       : str(uuid.uuid4()),
            'nombre'   : request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html')

@app.route('/areas/edit/<id>', methods=['GET','POST'])
def areas_edit(id):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(AREAS_CSV)
    a = next((x for x in rows if x['id']==id), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre']    = request.form['nombre']
        a['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=a)

@app.route('/areas/delete/<id>', methods=['POST'])
def areas_delete(id):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    rows = [x for x in cargar_csv(AREAS_CSV) if x['id']!=id]
    guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ─── CRUD Profesores ──────────────────────────────────────────────────────────
@app.route('/profesores')
def profesores_list():
    if not validar_acceso(['Administrador','Encargado']):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    return render_template('profesores_list.html', profesores=cargar_csv(PROFESORES_CSV))

@app.route('/profesores/new', methods=['GET','POST'])
def profesores_new():
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(PROFESORES_CSV)
        rows.append({
            'correo': request.form['correo'],
            'nombre': request.form['nombre'],
            'area'  : request.form['area']
        })
        guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html')

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
def profesores_edit(correo):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(PROFESORES_CSV)
    p = next((x for x in rows if x['correo']==correo), None)
    if not p: return "No existe",404
    if request.method=='POST':
        p['nombre'] = request.form['nombre']
        p['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=p)

@app.route('/profesores/delete/<correo>', methods=['POST'])
def profesores_delete(correo):
    if not validar_acceso(['Administrador','Encargado']):
        return redirect(url_for('dashboard'))
    rows = [x for x in cargar_csv(PROFESORES_CSV) if x['correo']!=correo]
    guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ─── CRUD Alumnos ─────────────────────────────────────────────────────────────
@app.route('/alumnos')
def alumnos_list():
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    rows = cargar_csv(ALUMNOS_CSV)
    u    = session['user']
    if u['rol']=='Maestro':
        rows = [x for x in rows if x['profesor']==u['correo']]
    if u['rol']=='Encargado':
        rows = [x for x in rows if x['area']==u['area']]
    return render_template('alumnos_list.html', alumnos=rows)

@app.route('/alumnos/new', methods=['GET','POST'])
def alumnos_new():
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            'correo'  : request.form['correo'],
            'nombre'  : request.form['nombre'],
            'area'    : request.form['area'],
            'profesor': request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, rows, ['correo','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html')

@app.route('/alumnos/edit/<correo>', methods=['GET','POST'])
def alumnos_edit(correo):
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(ALUMNOS_CSV)
    a    = next((x for x in rows if x['correo']==correo), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre']   = request.form['nombre']
        a['area']     = request.form['area']
        a['profesor'] = request.form['profesor']
        guardar_csv(ALUMNOS_CSV, rows, ['correo','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=a)

@app.route('/alumnos/delete/<correo>', methods=['POST'])
def alumnos_delete(correo):
    if not validar_acceso(['Administrador','Encargado','Maestro']):
        return redirect(url_for('dashboard'))
    rows = [x for x in cargar_csv(ALUMNOS_CSV) if x['correo']!=correo]
    guardar_csv(ALUMNOS_CSV, rows, ['correo','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))

# ─── Generación de documentos Word ────────────────────────────────────────────
def _find_row(no_control):
    for r in cargar_csv(DOCUMENTOS_CSV):
        if r.get('No_Control','') == str(no_control):
            return r
    return None

def _make_doc(template_path, context, filename):
    doc = DocxTemplate(template_path)
    doc.render(context)
    mem = BytesIO()
    doc.save(mem)
    mem.seek(0)
    return send_file(mem, download_name=filename, as_attachment=True)

@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    row = _find_row(no_control)
    if not row:
        flash('No existe registro para solicitud','danger')
        return redirect(url_for('dashboard'))
    context = row.copy()
    return _make_doc(PLANTILLA_SOLICITUD, context, f'Solicitud_{no_control}.docx')

@app.route('/generar_bimestral/<reporte_no>/<no_control>')
def generar_bimestral(reporte_no, no_control):
    row = _find_row(no_control)
    if not row:
        flash('No existe registro para bimestral','danger')
        return redirect(url_for('dashboard'))
    context = row.copy()
    context['Reporte_No'] = reporte_no
    for i in ['1','2','3']:
        context[f'x{i}'] = 'X' if i == reporte_no else ''
    return _make_doc(PLANTILLA_BIMESTRAL, context, f'Reporte_Bimestral_{reporte_no}_{no_control}.docx')

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    row = _find_row(no_control)
    if not row:
        flash('No existe registro para reporte final','danger')
        return redirect(url_for('dashboard'))
    context = row.copy()
    return _make_doc(PLANTILLA_FINAL, context, f'Reporte_Final_{no_control}.docx')

# ─── Manejo de error 404 ───────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    # En Render usará la env VAR PORT, en local por defecto 5000
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
