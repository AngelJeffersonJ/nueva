import os
from flask import Flask, render_template, redirect, url_for, session, request, flash, send_file
from dotenv import load_dotenv

from models import cargar_csv, guardar_csv, ALUMNOS_FIELDS, DOCUMENTOS_FIELDS
from auth import build_msal_app, get_auth_url, acquire_token_by_code
from decorators import login_required, roles_required

from docxtpl import DocxTemplate
from io import BytesIO

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecret')

# --- Rutas de archivos ---
CSV_DIR        = os.path.join(app.root_path, 'csv')
USERS_CSV      = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos_completo.csv')

TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

# ============ AUTENTICACIÓN ====================
@app.route('/login')
def login():
    msal_app = build_msal_app()
    auth_url = get_auth_url(msal_app)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = build_msal_app()
    result = acquire_token_by_code(msal_app, code)
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
    from auth import AUTHORITY
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login', _external=True)}"
    )

# ============ DASHBOARD ========================
@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# ============ CRUD USUARIOS ====================
@app.route('/usuarios')
@roles_required('Administrador')
def usuarios_list():
    return render_template('usuarios_list.html', usuarios=cargar_csv(USERS_CSV))

@app.route('/usuarios/new', methods=['GET', 'POST'])
@roles_required('Administrador')
def usuarios_new():
    if request.method == 'POST':
        u = cargar_csv(USERS_CSV)
        nuevo = {
            'correo': request.form.get('correo', '').strip().lower(),
            'rol':    request.form.get('rol', '').strip(),
            'area':   request.form.get('area', '').strip()
        }
        if not nuevo['correo']:
            flash("Correo obligatorio", "warning")
            return redirect(url_for('usuarios_new'))
        u.append(nuevo)
        guardar_csv(USERS_CSV, u, ['correo', 'rol', 'area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET', 'POST'])
@roles_required('Administrador')
def usuarios_edit(correo):
    u = cargar_csv(USERS_CSV)
    usr = next((x for x in u if x['correo'] == correo), None)
    if not usr:
        flash("No encontrado", "danger")
        return redirect(url_for('usuarios_list'))
    if request.method == 'POST':
        usr['rol']  = request.form.get('rol', '').strip()
        usr['area'] = request.form.get('area', '').strip()
        guardar_csv(USERS_CSV, u, ['correo', 'rol', 'area'])
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=usr)

@app.route('/usuarios/delete/<correo>')
@roles_required('Administrador')
def usuarios_delete(correo):
    u = [x for x in cargar_csv(USERS_CSV) if x['correo'] != correo]
    guardar_csv(USERS_CSV, u, ['correo', 'rol', 'area'])
    return redirect(url_for('usuarios_list'))

# ============ CRUD AREAS =======================
@app.route('/areas')
@roles_required('Administrador','Encargado')
def areas_list():
    return render_template('areas_list.html', areas=cargar_csv(AREAS_CSV))

@app.route('/areas/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_new():
    import uuid
    if request.method == 'POST':
        a = cargar_csv(AREAS_CSV)
        nueva = {
            'id':        str(uuid.uuid4()),
            'nombre':    request.form.get('nombre', '').strip(),
            'encargado': request.form.get('encargado', '').strip()
        }
        if not nueva['nombre']:
            flash("Nombre obligatorio", "warning")
            return redirect(url_for('areas_new'))
        a.append(nueva)
        guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def areas_edit(id):
    a = cargar_csv(AREAS_CSV)
    ar = next((x for x in a if x['id'] == id), None)
    if not ar:
        flash("No encontrado", "danger")
        return redirect(url_for('areas_list'))
    if request.method == 'POST':
        ar['nombre']    = request.form.get('nombre', '').strip()
        ar['encargado'] = request.form.get('encargado', '').strip()
        guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=ar)

@app.route('/areas/delete/<id>')
@roles_required('Administrador','Encargado')
def areas_delete(id):
    a = [x for x in cargar_csv(AREAS_CSV) if x['id'] != id]
    guardar_csv(AREAS_CSV, a, ['id','nombre','encargado'])
    return redirect(url_for('areas_list'))

# ============ CRUD PROFESORES ==================
@app.route('/profesores')
@roles_required('Administrador','Encargado')
def profesores_list():
    return render_template('profesores_list.html', profesores=cargar_csv(PROFESORES_CSV))

@app.route('/profesores/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_new():
    if request.method == 'POST':
        p = cargar_csv(PROFESORES_CSV)
        nuevo = {
            'correo': request.form.get('correo', '').strip().lower(),
            'nombre': request.form.get('nombre', '').strip(),
            'area':   request.form.get('area', '').strip()
        }
        if not nuevo['correo']:
            flash("Correo obligatorio", "warning")
            return redirect(url_for('profesores_new'))
        p.append(nuevo)
        guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
@roles_required('Administrador','Encargado')
def profesores_edit(correo):
    p = cargar_csv(PROFESORES_CSV)
    prof = next((x for x in p if x['correo'] == correo), None)
    if not prof:
        flash("No encontrado", "danger")
        return redirect(url_for('profesores_list'))
    if request.method == 'POST':
        prof['nombre'] = request.form.get('nombre', '').strip()
        prof['area']   = request.form.get('area', '').strip()
        guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
@roles_required('Administrador','Encargado')
def profesores_delete(correo):
    p = [x for x in cargar_csv(PROFESORES_CSV) if x['correo'] != correo]
    guardar_csv(PROFESORES_CSV, p, ['correo','nombre','area'])
    return redirect(url_for('profesores_list'))

# ============ CRUD ALUMNOS =====================
@app.route('/alumnos')
@login_required
def alumnos_list():
    al    = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    role  = session['role']; usr = session['user']
    if role=='Maestro':
        al = [x for x in al if x['profesor'].lower()==usr['correo'].lower()]
    if role=='Encargado':
        al = [x for x in al if x['area']==usr['area']]
    return render_template('alumnos_list.html', alumnos=al)

@app.route('/alumnos/new', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_new():
    if request.method == 'POST':
        no_ctl = request.form.get('No_Control', '').strip()
        if not no_ctl:
            flash("No_Control obligatorio", "warning")
            return redirect(url_for('alumnos_new'))
        lst = cargar_csv(ALUMNOS_CSV, ALUMNOS_FIELDS)
        lst.append({
            'No_Control': no_ctl,
            'nombre':     request.form.get('nombre', '').strip(),
            'area':       request.form.get('area', '').strip(),
            'profesor':   request.form.get('profesor', '').strip()
        })
        guardar_csv(ALUMNOS_CSV, lst, ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<No_Control>', methods=['GET','POST'])
@roles_required('Administrador','Encargado','Maestro')
def alumnos_edit(No_Control):
    lst  = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    alum = next((x for x in lst if x['No_Control']==No_Control), None)
    if not alum:
        flash("No encontrado", "danger")
        return redirect(url_for('alumnos_list'))
    if request.method == 'POST':
        alum['nombre']   = request.form.get('nombre', '').strip()
        alum['area']     = request.form.get('area', '').strip()
        alum['profesor'] = request.form.get('profesor', '').strip()
        guardar_csv(ALUMNOS_CSV, lst, ALUMNOS_FIELDS)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alum)

@app.route('/alumnos/delete/<No_Control>')
@roles_required('Administrador','Encargado','Maestro')
def alumnos_delete(No_Control):
    lst = [x for x in cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS) if x['No_Control']!=No_Control]
    guardar_csv(ALUMNOS_CSV, lst, ALUMNOS_FIELDS)
    return redirect(url_for('alumnos_list'))

# ============ CRUD DOCUMENTOS ==================
@app.route('/documentos')
@login_required
def documentos_list():
    docs    = cargar_csv(DOCUMENTOS_CSV,DOCUMENTOS_FIELDS)
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    return render_template('documentos_list.html', documentos=docs, alumnos=alumnos)

# ============ GENERACIÓN DE DOCX ===============
def render_docx(template_path, context, filename):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return send_file(
            buf, as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento","danger")
        return redirect(url_for('dashboard'))

@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOCUMENTOS_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return render_docx(TPL_SOLICITUD,{**alum,**doci},f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control,num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOCUMENTOS_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return render_docx(
        TPL_BIMESTRAL,
        {**alum,**doci,'reporte_no':num},
        f"Bimestral_{no_control}_B{num}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOCUMENTOS_FIELDS)
    alum    = next((a for a in alumnos if a['No_Control']==no_control),None)
    doci    = next((d for d in docs    if d['No_Control']==no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    doci['Nombre_Estudiante'] = doci.get('Nombre',alum.get('Nombre',''))
    return render_docx(TPL_FINAL,{**alum,**doci},f"Final_{no_control}.docx")

# =========== ERROR 404 =============
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'),404

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT',5000)))
