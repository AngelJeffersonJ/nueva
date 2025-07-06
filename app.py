import os
import csv
import io
import uuid
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import msal
from docxtpl import DocxTemplate

app = Flask(__name__)
# ---------------------- Configuración ----------------------
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

CLIENT_ID     = os.environ.get('CLIENT_ID', '')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET', '')
TENANT_ID     = os.environ.get('TENANT_ID', '')
AUTHORITY     = f'https://login.microsoftonline.com/{TENANT_ID}'
REDIRECT_URI  = os.environ.get('REDIRECT_URI', '')
SCOPE         = ['User.Read']

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
CSV_DIR       = os.environ.get('CSV_DIR', 'csv')
USUARIOS_CSV   = os.path.join(BASE_DIR, CSV_DIR, 'usuarios.csv')
PROFESORES_CSV = os.path.join(BASE_DIR, CSV_DIR, 'profesores.csv')
AREAS_CSV      = os.path.join(BASE_DIR, CSV_DIR, 'areas.csv')
ALUMNOS_CSV    = os.path.join(BASE_DIR, CSV_DIR, 'alumnos.csv')

PLANTILLA_SOLICITUD = os.path.join(BASE_DIR, 'plantillas', 'plantilla_solicitud_completa.docx')
PLANTILLA_BIMESTRAL = os.path.join(BASE_DIR, 'plantillas', 'Reporte_Bimestral_Plantilla.docx')
PLANTILLA_FINAL     = os.path.join(BASE_DIR, 'plantillas', 'Reporte_Final_Lleno.docx')

# ---------------------- Helpers CSV ----------------------
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

# Obtiene el perfil (rol, área) desde usuarios.csv
def obtener_perfil(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u['correo'].strip().lower() == email.lower():
            return u
    return None

# Decorador simple para chequear roles
def requiere_rol(*roles):
    def decorator(fn):
        def wrapper(*args, **kwargs):
            u = session.get('user')
            if not u or u.get('rol') not in roles:
                flash('Acceso denegado', 'danger')
                return redirect(url_for('dashboard'))
            return fn(*args, **kwargs)
        wrapper.__name__ = fn.__name__
        return wrapper
    return decorator

# ---------------------- Autenticación ----------------------
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
        flash(result.get('error_description','Error autenticando'), 'danger')
        return redirect(url_for('login'))

    claims = result.get('id_token_claims', {})
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales", 403

    perfil = obtener_perfil(email) or {'correo': email, 'rol': 'Estudiante', 'area': ''}
    session['user'] = {
        'correo': perfil['correo'],
        'rol': perfil['rol'],
        'area': perfil.get('area',''),
        'name': claims.get('name', email)
    }
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    post = url_for('login', _external=True)
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={post}")

# ---------------------- Dashboard ----------------------
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))

    rol = session['user']['rol']
    filas = cargar_csv(ALUMNOS_CSV)

    # Filtrar alumnos según rol
    if rol == 'Profesor':
        filas = [a for a in filas if a.get('profesor','').lower() == session['user']['correo'].lower()]
    elif rol == 'Jefe':
        filas = [a for a in filas if a.get('area') == session['user']['area']]
    # Estudiante: verá solo su propio No_Control (si lo deseas puedes filtrar)

    return render_template('dashboard.html',
                           usuario=session['user'],
                           alumnos=filas)

# ---------------------- CRUD Usuarios ----------------------
@app.route('/usuarios')
@requiere_rol('Administrador')
def usuarios_list():
    rows = cargar_csv(USUARIOS_CSV)
    return render_template('usuarios_list.html', usuarios=rows)

@app.route('/usuarios/new', methods=['GET','POST'])
@requiere_rol('Administrador')
def usuarios_new():
    if request.method=='POST':
        rows = cargar_csv(USUARIOS_CSV)
        nuevo = {
            'correo': request.form['correo'],
            'rol':    request.form['rol'],
            'area':   request.form['area']
        }
        rows.append(nuevo)
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
@requiere_rol('Administrador')
def usuarios_edit(correo):
    rows = cargar_csv(USUARIOS_CSV)
    u = next((r for r in rows if r['correo']==correo), None)
    if not u:
        flash('No existe ese usuario','danger')
        return redirect(url_for('usuarios_list'))
    if request.method=='POST':
        u['rol']  = request.form['rol']
        u['area'] = request.form['area']
        guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=u)

@app.route('/usuarios/delete/<correo>', methods=['POST'])
@requiere_rol('Administrador')
def usuarios_delete(correo):
    rows = [r for r in cargar_csv(USUARIOS_CSV) if r['correo']!=correo]
    guardar_csv(USUARIOS_CSV, rows, ['correo','rol','area'])
    return redirect(url_for('usuarios_list'))

# ---------------------- CRUD Áreas ----------------------
@app.route('/areas')
@requiere_rol('Administrador','Jefe')
def areas_list():
    rows = cargar_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=rows)

@app.route('/areas/new', methods=['GET','POST'])
@requiere_rol('Administrador','Jefe')
def areas_new():
    if request.method=='POST':
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            'id': str(uuid.uuid4()),
            'nombre':    request.form['nombre'],
            'encargado': request.form['encargado']
        })
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@requiere_rol('Administrador','Jefe')
def areas_edit(id):
    rows = cargar_csv(AREAS_CSV)
    a = next((r for r in rows if r['id']==id), None)
    if not a:
        flash('Área no encontrada','danger')
        return redirect(url_for('areas_list'))
    if request.method=='POST':
        a['nombre']    = request.form['nombre']
        a['encargado'] = request.form['encargado']
        guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=a)

@app.route('/areas/delete/<id>', methods=['POST'])
@requiere_rol('Administrador','Jefe')
def areas_delete(id):
    rows = [r for r in cargar_csv(AREAS_CSV) if r['id']!=id]
    guardar_csv(AREAS_CSV, rows, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ---------------------- CRUD Profesores ----------------------
@app.route('/profesores')
@requiere_rol('Administrador','Jefe')
def profesores_list():
    rows = cargar_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=rows)

@app.route('/profesores/new', methods=['GET','POST'])
@requiere_rol('Administrador','Jefe')
def profesores_new():
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
@requiere_rol('Administrador','Jefe')
def profesores_edit(correo):
    rows = cargar_csv(PROFESORES_CSV)
    p = next((r for r in rows if r['correo']==correo), None)
    if not p:
        flash('No encontrado','danger')
        return redirect(url_for('profesores_list'))
    if request.method=='POST':
        p['nombre'] = request.form['nombre']
        p['area']   = request.form['area']
        guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=p)

@app.route('/profesores/delete/<correo>', methods=['POST'])
@requiere_rol('Administrador','Jefe')
def profesores_delete(correo):
    rows = [r for r in cargar_csv(PROFESORES_CSV) if r['correo']!=correo]
    guardar_csv(PROFESORES_CSV, rows, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ---------------------- CRUD Alumnos ----------------------
@app.route('/alumnos')
@requiere_rol('Administrador','Jefe','Profesor','Estudiante')
def alumnos_list():
    rows = cargar_csv(ALUMNOS_CSV)
    rol = session['user']['rol']
    if rol=='Profesor':
        rows = [r for r in rows if r.get('profesor','').lower()==session['user']['correo'].lower()]
    if rol=='Jefe':
        rows = [r for r in rows if r.get('area')==session['user']['area']]
    if rol=='Estudiante':
        rows = [r for r in rows if r.get('No_Control')==session['user']['correo'].split('@')[0]]
    return render_template('alumnos_list.html', alumnos=rows)

@app.route('/alumnos/new', methods=['GET','POST'])
@requiere_rol('Administrador','Jefe','Profesor')
def alumnos_new():
    if request.method=='POST':
        rows = cargar_csv(ALUMNOS_CSV)
        nuevo = {
            'No_Control': request.form['No_Control'],
            'Nombre':     request.form['Nombre'],
            'Apellido_Paterno': request.form['Apellido_Paterno'],
            'Apellido_Materno': request.form['Apellido_Materno'],
            'Carrera':    request.form['Carrera'],
            'area':       request.form['area'],
            'profesor':   request.form['profesor'],
            # aquí puedes añadir más campos según tu CSV
        }
        rows.append(nuevo)
        guardar_csv(ALUMNOS_CSV, rows, list(nuevo.keys()))
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<No_Control>', methods=['GET','POST'])
@requiere_rol('Administrador','Jefe','Profesor')
def alumnos_edit(No_Control):
    rows = cargar_csv(ALUMNOS_CSV)
    a = next((r for r in rows if r['No_Control']==No_Control), None)
    if not a:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('alumnos_list'))
    if request.method=='POST':
        for campo in ['Nombre','Apellido_Paterno','Apellido_Materno','Carrera','area','profesor']:
            a[campo] = request.form[campo]
        guardar_csv(ALUMNOS_CSV, rows, list(a.keys()))
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=a)

@app.route('/alumnos/delete/<No_Control>', methods=['POST'])
@requiere_rol('Administrador','Jefe','Profesor')
def alumnos_delete(No_Control):
    rows = [r for r in cargar_csv(ALUMNOS_CSV) if r['No_Control']!=No_Control]
    if rows:
        guardar_csv(ALUMNOS_CSV, rows, list(rows[0].keys()))
    return redirect(url_for('alumnos_list'))

# ---------------------- Generación de Documentos ----------------------
@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    alumno = next((a for a in alumnos if a['No_Control']==no_control), None)
    if not alumno:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('dashboard'))
    doc = DocxTemplate(PLANTILLA_SOLICITUD)
    doc.render(alumno)
    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Solicitud_{no_control}.docx",
                     as_attachment=True)

@app.route('/generar_bimestral/<int:bimestre>/<no_control>')
def generar_bimestral(bimestre, no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    alumno = next((a for a in alumnos if a['No_Control']==no_control), None)
    if not alumno:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('dashboard'))
    ctx = alumno.copy()
    ctx['Reporte_No'] = bimestre
    doc = DocxTemplate(PLANTILLA_BIMESTRAL)
    doc.render(ctx)
    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Bimestral_{bimestre}_{no_control}.docx",
                     as_attachment=True)

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    alumno = next((a for a in alumnos if a['No_Control']==no_control), None)
    if not alumno:
        flash('Alumno no encontrado','danger')
        return redirect(url_for('dashboard'))
    doc = DocxTemplate(PLANTILLA_FINAL)
    doc.render(alumno)
    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return send_file(buf,
                     download_name=f"Reporte_Final_{no_control}.docx",
                     as_attachment=True)

# ---------------------- Error 404 ----------------------
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__=='__main__':
    app.run(debug=True)
