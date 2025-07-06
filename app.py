import os
import csv
import uuid
import zipfile
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import msal
import pandas as pd
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'clave_segura')

# --- Configuración Azure AD ---
CLIENT_ID     = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
TENANT_ID     = os.environ.get('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get('REDIRECT_URI')
SCOPE         = ['User.Read']

# --- Rutas y paths ---
BASE_DIR        = os.path.dirname(__file__)
CSV_DIR         = os.environ.get('CSV_DIR', 'csv')
USUARIOS_CSV    = os.path.join(BASE_DIR, CSV_DIR, 'usuarios.csv')
AREAS_CSV       = os.path.join(BASE_DIR, CSV_DIR, 'areas.csv')
PROFESORES_CSV  = os.path.join(BASE_DIR, CSV_DIR, 'profesores.csv')
ALUMNOS_CSV     = os.path.join(BASE_DIR, CSV_DIR, 'alumnos.csv')
DOCUMENTOS_CSV  = os.path.join(BASE_DIR, CSV_DIR, 'documentos.csv')

PLANTILLA_SOLICITUD = os.path.join(BASE_DIR, 'plantillas','ITA-VI-SS-FO-01 Solicitud de Servicio Social.docx')
PLANTILLA_BIMESTRAL = os.path.join(BASE_DIR, 'plantillas','ITA-VI-SS-FO-02 Reporte Bimestral.docx')
PLANTILLA_FINAL     = os.path.join(BASE_DIR, 'plantillas','ITA-VI-SS-FO-03 Reporte Final.docx')

OUTPUT_DIR      = os.path.join(BASE_DIR, 'documentos_generados')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Campos para CRUD genérico ---
MAPA_ENTIDADES = {
    'usuarios'  : (USUARIOS_CSV,   ['correo','rol','area']),
    'areas'     : (AREAS_CSV,      ['id','nombre','encargado']),
    'profesores': (PROFESORES_CSV, ['correo','nombre','area']),
    'alumnos'   : (ALUMNOS_CSV,    ['correo','nombre','area','profesor','No_Control'])
}

# --- Helper CSV ---
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, campos):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(rows)

# --- Usuarios y permisos ---
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u['correo'].lower() == email.lower():
            return u
    return None

def validar_acceso(permitidos):
    u = session.get('user')
    return u and u.get('rol') in permitidos

# --- Azure AD Auth ---
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(scopes=SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route('/getAToken')
def authorized():
    code = request.args.get('code')
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
    if 'error' in result:
        flash(result.get('error_description','Error de autenticación'),'danger')
        return redirect(url_for('dashboard'))
    claims = result.get('id_token_claims',{})
    email = claims.get('preferred_username','').lower()
    if not email.endswith('@aguascalientes.tecnm.mx'):
        return "Solo cuentas @aguascalientes.tecnm.mx",403
    perfil = usuario_desde_csv(email) or {'correo':email,'rol':'Alumno','area':''}
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
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login',_external=True)}")

# --- Dashboard ---
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

# --- CRUD genérico ---
@app.route('/entidad/<tipo>')
def entidad_list(tipo):
    if tipo not in MAPA_ENTIDADES:
        return "Tipo inválido",404
    # permisos
    roles = {
        'usuarios': ['Administrador'],
        'areas':    ['Administrador','Encargado'],
        'profesores':['Administrador','Encargado'],
        'alumnos':  ['Administrador','Encargado','Maestro']
    }[tipo]
    if not validar_acceso(roles):
        flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    csv_path, campos = MAPA_ENTIDADES[tipo]
    registros = cargar_csv(csv_path)
    # filtro para Maestro/Encargado en alumnos
    if tipo=='alumnos':
        u=session['user']
        if u['rol']=='Maestro':
            registros=[r for r in registros if r['profesor']==u['correo']]
        if u['rol']=='Encargado':
            registros=[r for r in registros if r['area']==u['area']]
    return render_template('entidad_list.html', tipo=tipo, campos=campos, registros=registros)

@app.route('/entidad/<tipo>/new', methods=['GET','POST'])
def entidad_new(tipo):
    if tipo not in MAPA_ENTIDADES: return "Tipo inválido",404
    roles = {
        'usuarios': ['Administrador'],
        'areas':    ['Administrador','Encargado'],
        'profesores':['Administrador','Encargado'],
        'alumnos':  ['Administrador','Encargado','Maestro']
    }[tipo]
    if not validar_acceso(roles):
        flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    csv_path, campos = MAPA_ENTIDADES[tipo]
    if request.method=='POST':
        row={c: request.form.get(c,'').strip() for c in campos if c!='id'}
        if tipo=='areas': row['id']=str(uuid.uuid4())
        filas=cargar_csv(csv_path)
        filas.append(row)
        guardar_csv(csv_path,filas,campos)
        return redirect(url_for('entidad_list',tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo, campos=campos, valores={})

@app.route('/entidad/<tipo>/edit/<pk>', methods=['GET','POST'])
def entidad_edit(tipo,pk):
    if tipo not in MAPA_ENTIDADES: return "Tipo inválido",404
    roles = {
        'usuarios': ['Administrador'],
        'areas':    ['Administrador','Encargado'],
        'profesores':['Administrador','Encargado'],
        'alumnos':  ['Administrador','Encargado','Maestro']
    }[tipo]
    if not validar_acceso(roles):
        flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    csv_path, campos = MAPA_ENTIDADES[tipo]
    filas=cargar_csv(csv_path)
    llave=campos[0]
    item = next((r for r in filas if r[llave]==pk), None)
    if not item: flash('No encontrado','danger'); return redirect(url_for('entidad_list',tipo=tipo))
    if request.method=='POST':
        for c in campos:
            if c!='id':
                item[c]=request.form.get(c,'').strip()
        guardar_csv(csv_path,filas,campos)
        return redirect(url_for('entidad_list',tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo, campos=campos, valores=item)

@app.route('/entidad/<tipo>/delete/<pk>', methods=['POST'])
def entidad_delete(tipo,pk):
    if tipo not in MAPA_ENTIDADES: return "Tipo inválido",404
    roles = {
        'usuarios': ['Administrador'],
        'areas':    ['Administrador','Encargado'],
        'profesores':['Administrador','Encargado'],
        'alumnos':  ['Administrador','Encargado','Maestro']
    }[tipo]
    if not validar_acceso(roles):
        flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    csv_path, campos = MAPA_ENTIDADES[tipo]
    filas=[r for r in cargar_csv(csv_path) if r[campos[0]]!=pk]
    guardar_csv(csv_path,filas,campos)
    return redirect(url_for('entidad_list',tipo=tipo))

# --- Generación de documentos ---
def _find_doc(no_control):
    for r in cargar_csv(DOCUMENTOS_CSV):
        if r.get('No_Control')==no_control:
            return r
    return None

def _make_doc(template, contexto, filename):
    doc = DocxTemplate(template)
    doc.render(contexto)
    out_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(out_path)
    return send_file(out_path, as_attachment=True, download_name=filename)

@app.route('/generar/solicitud/<no_control>')
def generar_solicitud(no_control):
    if 'user' not in session: return redirect(url_for('login'))
    row=_find_doc(no_control)
    if not row: flash('Alumno no encontrado','danger'); return redirect(url_for('dashboard'))
    # permiso
    rol, area, correo = session['user']['rol'], session['user']['area'], session['user']['correo']
    if rol=='Encargado' and row.get('Área')!=area: flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    if rol=='Maestro' and row.get('Profesor')!=correo: flash('Acceso denegado','danger'); return redirect(url_for('dashboard'))
    if rol=='Alumno':
        # buscar su control
        alum = next((a for a in cargar_csv(ALUMNOS_CSV) if a['correo']==correo), None)
        if not alum or alum['No_Control']!=no_control:
            flash('Solo tu propio documento','danger'); return redirect(url_for('dashboard'))
    return _make_doc(PLANTILLA_SOLICITUD, row, f"Solicitud_{no_control}.docx")

@app.route('/generar/bimestral/<bimestre>/<no_control>')
def generar_bimestral(bimestre,no_control):
    if 'user' not in session: return redirect(url_for('login'))
    row=_find_doc(no_control)
    if not row: flash('Alumno no encontrado','danger'); return redirect(url_for('dashboard'))
    # mismo control de acceso que arriba...
    contexto={**row, 'Reporte_No':bimestre}
    return _make_doc(PLANTILLA_BIMESTRAL, contexto, f"Reporte_Bimestral_{bimestre}_{no_control}.docx")

@app.route('/generar/final/<no_control>')
def generar_final(no_control):
    if 'user' not in session: return redirect(url_for('login'))
    row=_find_doc(no_control)
    if not row: flash('Alumno no encontrado','danger'); return redirect(url_for('dashboard'))
    return _make_doc(PLANTILLA_FINAL, row, f"Reporte_Final_{no_control}.docx")

# --- Error 404 ---
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'),404

if __name__=='__main__':
    app.run(debug=True)
