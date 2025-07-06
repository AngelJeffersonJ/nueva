import os
import csv
import uuid
import msal
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from io import BytesIO
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET', 'clave_segura')

# --- Azure AD (MSAL) ---
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ["User.Read"]

# --- CSV paths ---
CSV_DIR        = os.getenv('CSV_DIR', 'csv')
USUARIOS_CSV   = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos.csv')
DATOS_CSV      = os.path.join(CSV_DIR, 'documentos.csv')

# --- DOCX templates ---
PLANTILLA_SOL = 'plantillas/plantilla_solicitud_completa.docx'
PLANTILLA_BIM = 'plantillas/Reporte_Bimestral_Plantilla.docx'
PLANTILLA_FIN = 'plantillas/Reporte_Final_Plantilla.docx'

# ────────────────── Helpers CSV ───────────────────────────────────────────────

def cargar_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def guardar_csv(path, rows, fieldnames):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

# ────────────────── Helper DOCX ──────────────────────────────────────────────

def _make_doc(template_path, contexto):
    doc = DocxTemplate(template_path)
    doc.render(contexto)
    mem = BytesIO()
    doc.save(mem)
    mem.seek(0)
    return mem

# ────────────────── Autenticación MSAL ───────────────────────────────────────

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
        flash(result.get('error_description','Error en autenticación'),'danger')
        return redirect(url_for('dashboard'))

    claims = result.get('id_token_claims',{})
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403

    usuarios = cargar_csv(USUARIOS_CSV)
    perfil = next((u for u in usuarios if u['correo'].lower()==email), None)
    if not perfil:
        perfil = {'correo': email, 'rol':'Alumno', 'area':''}
        usuarios.append(perfil)
        guardar_csv(USUARIOS_CSV, usuarios, ['correo','rol','area'])

    session['user'] = {
        'correo': perfil['correo'],
        'rol':     perfil['rol'],
        'area':    perfil.get('area',''),
        'name':    claims.get('name','')
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    post = url_for('login', _external=True)
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={post}")

# ────────────────── Dashboard ─────────────────────────────────────────────────

@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

def validar_rol(*roles):
    return session.get('user',{}).get('rol') in roles

# ────────────────── CRUD Usuarios ─────────────────────────────────────────────

@app.route('/usuarios')
def usuarios_list():
    if not validar_rol('Administrador'):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    rows = cargar_csv(USUARIOS_CSV)
    return render_template('usuarios_list.html', usuarios=rows)

@app.route('/usuarios/new', methods=['GET','POST'])
def usuarios_new():
    if not validar_rol('Administrador'):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(USUARIOS_CSV)
        rows.append({
            'correo': request.form['correo'],
            'rol':    request.form['rol'],
            'area':   request.form['area']
        })
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
def usuarios_edit(correo):
    if not validar_rol('Administrador'):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(USUARIOS_CSV)
    u = next((r for r in rows if r['correo']==correo), None)
    if not u: return "No existe",404
    if request.method=='POST':
        u['rol']  = request.form['rol']
        u['area'] = request.form['area']
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=u)

@app.route('/usuarios/delete/<correo>', methods=['POST'])
def usuarios_delete(correo):
    if not validar_rol('Administrador'):
        return redirect(url_for('dashboard'))
    rows = [r for r in cargar_csv(USUARIOS_CSV) if r['correo']!=correo]
    guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ────────────────── CRUD Áreas ────────────────────────────────────────────────

@app.route('/areas')
def areas_list():
    if not validar_rol('Administrador','Encargado'):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    rows = cargar_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=rows)

@app.route('/areas/new', methods=['GET','POST'])
def areas_new():
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            'id':        str(uuid.uuid4()),
            'nombre':    request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
def areas_edit(id):
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(AREAS_CSV)
    a = next((r for r in rows if r['id']==id), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre']    = request.form['nombre']
        a['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=a)

@app.route('/areas/delete/<id>', methods=['POST'])
def areas_delete(id):
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    rows = [r for r in cargar_csv(AREAS_CSV) if r['id']!=id]
    guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ────────────────── CRUD Profesores ───────────────────────────────────────────

@app.route('/profesores')
def profesores_list():
    if not validar_rol('Administrador','Encargado'):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    rows = cargar_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=rows)

@app.route('/profesores/new', methods=['GET','POST'])
def profesores_new():
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        rows = cargar_csv(PROFESORES_CSV)
        rows.append({
            'correo': request.form['correo'],
            'nombre': request.form['nombre'],
            'area':   request.form['area']
        })
        guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
def profesores_edit(correo):
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    rows = cargar_csv(PROFESORES_CSV)
    p = next((r for r in rows if r['correo']==correo), None)
    if not p: return "No existe",404
    if request.method=='POST':
        p['nombre'] = request.form['nombre']
        p['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=p)

@app.route('/profesores/delete/<correo>', methods=['POST'])
def profesores_delete(correo):
    if not validar_rol('Administrador','Encargado'):
        return redirect(url_for('dashboard'))
    rows = [r for r in cargar_csv(PROFESORES_CSV) if r['correo']!=correo]
    guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ────────────────── CRUD Alumnos (clave = No_Control) ─────────────────────────

@app.route('/alumnos')
def alumnos_list():
    if not validar_rol('Administrador','Encargado','Maestro'):
        flash('Acceso denegado','danger')
        return redirect(url_for('dashboard'))
    filas = cargar_csv(ALUMNOS_CSV)
    u = session['user']
    if u['rol']=='Maestro':
        filas = [r for r in filas if r['profesor']==u['correo']]
    if u['rol']=='Encargado':
        filas = [r for r in filas if r['area']==u['area']]
    return render_template('alumnos_list.html', alumnos=filas)

@app.route('/alumnos/new', methods=['GET','POST'])
def alumnos_new():
    if not validar_rol('Administrador','Encargado','Maestro'):
        return redirect(url_for('dashboard'))
    if request.method=='POST':
        filas = cargar_csv(ALUMNOS_CSV)
        filas.append({
            'No_Control': request.form['No_Control'],
            'nombre':     request.form['nombre'],
            'area':       request.form['area'],
            'profesor':   request.form['profesor']
        })
        guardar_csv(ALUMNOS_CSV, filas, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
def alumnos_edit(no_control):
    if not validar_rol('Administrador','Encargado','Maestro'):
        return redirect(url_for('dashboard'))
    filas = cargar_csv(ALUMNOS_CSV)
    a = next((r for r in filas if r['No_Control']==no_control), None)
    if not a: return "No existe",404
    if request.method=='POST':
        a['nombre']   = request.form['nombre']
        a['area']     = request.form['area']
        a['profesor'] = request.form['profesor']
        guardar_csv(ALUMNOS_CSV, filas, ['No_Control','nombre','area','profesor'])
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=a)

@app.route('/alumnos/delete/<no_control>', methods=['POST'])
def alumnos_delete(no_control):
    if not validar_rol('Administrador','Encargado','Maestro'):
        return redirect(url_for('dashboard'))
    filas = [r for r in cargar_csv(ALUMNOS_CSV) if r['No_Control']!=no_control]
    guardar_csv(ALUMNOS_CSV, filas, ['No_Control','nombre','area','profesor'])
    return redirect(url_for('alumnos_list'))

# ────────────────── Generación de Documentos ──────────────────────────────────

def _datos_docs():
    return cargar_csv(DATOS_CSV)

# Índices
@app.route('/generar_solicitud/')
def generar_solicitud_index():
    return render_template('gen_index.html',
        titulo="Generar Solicitud",
        ruta_base='generar_solicitud',
        alumnos=_datos_docs()
    )

@app.route('/generar_bimestral/')
def generar_bimestral_index():
    return render_template('gen_index.html',
        titulo="Generar Reporte Bimestral",
        ruta_base='generar_bimestral',
        alumnos=_datos_docs()
    )

@app.route('/generar_final/')
def generar_final_index():
    return render_template('gen_index.html',
        titulo="Generar Reporte Final",
        ruta_base='generar_final',
        alumnos=_datos_docs()
    )

# Acciones
@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    fila = next((r for r in _datos_docs() if r['No_Control']==no_control), None)
    if not fila:
        flash("Alumno no encontrado","danger")
        return redirect(url_for('generar_solicitud_index'))
    mem = _make_doc(PLANTILLA_SOL, fila)
    return send_file(mem, download_name=f"Solicitud_{no_control}.docx", as_attachment=True)

@app.route('/generar_bimestral/<int:bimestre>/<no_control>')
def generar_bimestral(bimestre, no_control):
    fila = next((r for r in _datos_docs() if r['No_Control']==no_control), None)
    if not fila or bimestre not in (1,2,3):
        flash("Parámetros inválidos","danger")
        return redirect(url_for('generar_bimestral_index'))
    contexto = {**fila, 'Reporte_No': bimestre}
    mem = _make_doc(PLANTILLA_BIM, contexto)
    return send_file(mem, download_name=f"Reporte_Bimestral_{bimestre}_{no_control}.docx", as_attachment=True)

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    fila = next((r for r in _datos_docs() if r['No_Control']==no_control), None)
    if not fila:
        flash("Alumno no encontrado","danger")
        return redirect(url_for('generar_final_index'))
    mem = _make_doc(PLANTILLA_FIN, fila)
    return send_file(mem, download_name=f"Reporte_Final_{no_control}.docx", as_attachment=True)

# ────────────────── 404 Handler ───────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True)
