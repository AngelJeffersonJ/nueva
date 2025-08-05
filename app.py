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
app.secret_key = os.getenv('FLASK_SECRET_KEY','supersecret')

# Configuración de Azure AD
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# Rutas de CSV y plantillas
CSV_DIR        = os.path.join(app.root_path,'csv')
USERS_CSV      = os.path.join(CSV_DIR,'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR,'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR,'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR,'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR,'documentos_completo.csv')

TPL_DIR        = os.path.join(app.root_path,'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR,'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR,'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR,'Reporte_Final_Lleno.docx')

FIELD_SYNONYMS = {
    'ap': 'Apellido_Paterno',
    'am': 'Apellido_Materno',
    'no_de_control': 'No_Control',
    'num_control': 'Num_Control',
    'carrera>': 'Carrera',
    'nombre_estudiante': 'Nombre',
}
ALUMNOS_FIELDS = ['No_Control','nombre','area','profesor']
DOC_FIELDS = [
    'No_Control','Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio','Carrera',
    'Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia','Titular_Dependencia','Cargo_Responsable',
    'Responsable_Proyecto','Nombre_Programa','Modalidad_Externa','Modalidad_Interna','Fecha_Inicio','Fecha_Terminacion',
    'Actividades','TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico','TP_Sustentable',
    'TP_Salud','TP_Medio_Amb','TP_Otros','Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'
]
DOC_FIELDS += [f'{x}_{i}' for x in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2') for i in range(1,4)]
DOC_FIELDS += ['Nombre_supervisor','Puesto_supervisor']
DOC_FIELDS += [f'Actividad_{i}' for i in range(1,9)]
DOC_FIELDS += ['x1','x2','x3']
DOC_FIELDS += [f'resp{n}_{i}' for n in range(3,18) for i in range(5)]
DOC_FIELDS += [f'estu{n}_{i}' for n in range(3,18) for i in range(5)]
DOC_FIELDS += ['Municipio','Estado']
DOC_FIELDS += [f'Actividad{i}' for i in range(1,9)]
DOC_FIELDS += [f'Logro{i}' for i in range(1,9)]
DOC_FIELDS += [f'Aprendizaje{i}' for i in range(1,9)]
DOC_FIELDS += [f'Beneficio{i}' for i in range(1,9)]
DOC_FIELDS += ['Nombre_Responsable','Cargo_Responsable','Num_Control','Observaciones_Encargado','Observaciones_Estudiante']

def cargar_csv(path, expected_fields=None):
    rows = []
    if os.path.isfile(path):
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for raw in reader:
                norm = {}
                for key,val in raw.items():
                    canon = FIELD_SYNONYMS.get(key.strip().lower(), key.strip())
                    norm[canon] = val.strip()
                if expected_fields:
                    for fld in expected_fields:
                        norm.setdefault(fld,'')
                rows.append(norm)
    return rows

def guardar_csv(path, rows, fieldnames):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'w',newline='',encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader(); w.writerows(rows)
    except Exception as e:
        app.logger.error(f"Error guardando CSV: {e}")
        flash("No se pudo guardar el CSV","danger")

def login_required(f):
    @wraps(f)
    def w(*args,**kw):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args,**kw)
    return w

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def w(*args,**kw):
            if session.get('role') not in roles:
                flash("Acceso denegado","danger")
                return redirect(url_for('dashboard'))
            return f(*args,**kw)
        return w
    return deco

def _render_docx(template_path, context, filename):
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
        flash(result.get('error_description','Error al autenticar'),'danger')
        return redirect(url_for('login'))
    claims = result['id_token_claims']
    email  = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas institucionales",403
    usuarios = cargar_csv(USERS_CSV)
    perfil   = next((u for u in usuarios if u.get('correo','').lower()==email), None)
    if not perfil:
        perfil = {'correo':email,'rol':'Alumno','area':''}
    session.update(user=perfil, role=perfil['rol'], name=claims.get('name',''))
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri="
        f"{url_for('login',_external=True)}"
    )

@app.route('/')
@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# CRUD Usuarios/Áreas/Profesores/Alumnos
@app.route('/usuarios')
@login_required
def usuarios_list():
    return render_template('usuarios_list.html', usuarios=cargar_csv(USERS_CSV))

@app.route('/areas')
@login_required
def areas_list():
    return render_template('areas_list.html', areas=cargar_csv(AREAS_CSV))

@app.route('/profesores')
@login_required
def profesores_list():
    return render_template('profesores_list.html', profesores=cargar_csv(PROFESORES_CSV))

@app.route('/alumnos')
@login_required
def alumnos_list():
    return render_template('alumnos_list.html', alumnos=cargar_csv(ALUMNOS_CSV))

# --- CRUD DOCUMENTOS ---
@app.route('/documentos')
@login_required
def documentos_list():
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    return render_template('documentos_list.html', documentos=docs, alumnos=alumnos)

# --- SOLICITUD ---
@app.route('/documentos/solicitud/<no_control>', methods=['GET','POST'])
@login_required
def documentos_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV, DOC_FIELDS)
    doc  = next((d for d in docs if d.get('No_Control','') == no_control), None)
    if not doc:
        doc = {k: '' for k in DOC_FIELDS}
        doc['No_Control'] = no_control
        docs.append(doc)
    if request.method == 'POST':
        for f in DOC_FIELDS:
            doc[f] = request.form.get(f, doc.get(f, ''))
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Solicitud actualizada correctamente", "success")
        return redirect(url_for('documentos_list'))
    return render_template('solicitud_form.html', documento=doc)

# --- BIMESTRAL ---
@app.route('/documentos/bimestral/<no_control>/<int:num>', methods=['GET','POST'])
@login_required
def documentos_bimestral(no_control,num):
    if num not in (1,2,3):
        flash("Bimestre inválido","warning")
        return redirect(url_for('documentos_list'))
    docs = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    doc  = next((d for d in docs if d.get('No_Control','') == no_control), None)
    if not doc:
        doc = {k: '' for k in DOC_FIELDS}
        doc['No_Control'] = no_control
        docs.append(doc)
    if request.method == 'POST':
        for f in DOC_FIELDS:
            doc[f] = request.form.get(f, doc.get(f, ''))
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash(f"Bimestral {num} guardado correctamente", "success")
        return redirect(url_for('documentos_list'))
    return render_template('bimestral_form.html', documento=doc, num=num)

# --- FINAL ---
@app.route('/documentos/final/<no_control>', methods=['GET','POST'])
@login_required
def documentos_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV, DOC_FIELDS)
    doc  = next((d for d in docs if d.get('No_Control','') == no_control), None)
    if not doc:
        doc = {k: '' for k in DOC_FIELDS}
        doc['No_Control'] = no_control
        docs.append(doc)
    if request.method == 'POST':
        for f in DOC_FIELDS:
            doc[f] = request.form.get(f, doc.get(f, ''))
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Reporte final guardado correctamente", "success")
        return redirect(url_for('documentos_list'))
    return render_template('final_form.html', documento=doc)

# --- GENERACIÓN DE DOCX ---
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a.get('No_Control','') == no_control),None)
    doci    = next((d for d in docs    if d.get('No_Control','') == no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(TPL_SOLICITUD,{**alum,**doci},f"Solicitud_{no_control}.docx")

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control,num):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a.get('No_Control','') == no_control),None)
    doci    = next((d for d in docs    if d.get('No_Control','') == no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(TPL_BIMESTRAL, {**alum,**doci,'reporte_no':num}, f"Bimestral_{no_control}_B{num}.docx")

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV,ALUMNOS_FIELDS)
    docs    = cargar_csv(DOCUMENTOS_CSV,DOC_FIELDS)
    alum    = next((a for a in alumnos if a.get('No_Control','') == no_control),None)
    doci    = next((d for d in docs    if d.get('No_Control','') == no_control),None)
    if not alum or not doci:
        flash("Datos incompletos","warning")
        return redirect(url_for('documentos_list'))
    doci['Nombre_Estudiante'] = doci.get('Nombre',alum.get('Nombre',''))
    return _render_docx(TPL_FINAL,{**alum,**doci},f"Final_{no_control}.docx")

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'),404

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT',5000)))
