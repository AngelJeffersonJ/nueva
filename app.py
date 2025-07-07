import os
import csv
import uuid
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, send_from_directory, flash
)
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = 'clave_segura'

# Directorios y rutas
CSV_DIR = 'csv'
ALUMNOS_CSV      = os.path.join(CSV_DIR, 'alumnos.csv')
AREAS_CSV        = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV   = os.path.join(CSV_DIR, 'profesores.csv')
USUARIOS_CSV     = os.path.join(CSV_DIR, 'usuarios.csv')
DOCUMENTOS_DIR   = 'documentos_generados'
PLANTILLAS_DIR   = 'plantillas'

# ── Helpers CSV ────────────────────────────────────────────────────────────────

def leer_csv(ruta):
    if not os.path.exists(ruta):
        return []
    with open(ruta, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def escribir_csv(ruta, campos, filas):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        writer.writerows(filas)

# ── Helper búsqueda ───────────────────────────────────────────────────────────

def buscar_por(ruta, clave, valor):
    for fila in leer_csv(ruta):
        if fila.get(clave) == valor:
            return fila
    return None

# ── Generación de DOCX ────────────────────────────────────────────────────────

def generar_docx(plantilla, datos, salida):
    os.makedirs(DOCUMENTOS_DIR, exist_ok=True)
    tpl = DocxTemplate(os.path.join(PLANTILLAS_DIR, plantilla))
    tpl.render(datos)
    path = os.path.join(DOCUMENTOS_DIR, salida)
    tpl.save(path)
    return path

# ── Autenticación Ficticia ────────────────────────────────────────────────────

@app.route('/login')
def login():
    # Login simulado: siempre admin
    session['user'] = {'correo':'admin@inst.mx','nombre':'Administrador'}
    session['role'] = 'Administrador'
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.route('/')
@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html', usuario=session['user'])

# ── CRUD ALUMNOS ───────────────────────────────────────────────────────────────

@app.route('/alumnos')
def alumnos_list():
    filas = leer_csv(ALUMNOS_CSV)
    return render_template('alumnos_list.html', alumnos=filas)

@app.route('/alumnos/new', methods=['GET','POST'])
def alumnos_new():
    if request.method=='POST':
        datos = request.form.to_dict()
        filas = leer_csv(ALUMNOS_CSV)
        filas.append(datos)
        escribir_csv(ALUMNOS_CSV, filas[0].keys(), filas)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno={})

@app.route('/alumnos/edit/<no_control>', methods=['GET','POST'])
def alumnos_edit(no_control):
    filas = leer_csv(ALUMNOS_CSV)
    alumno = buscar_por(ALUMNOS_CSV, 'no_control', no_control)
    if not alumno:
        return "No encontrado",404
    if request.method=='POST':
        nuevos = request.form.to_dict()
        for i,f in enumerate(filas):
            if f['no_control']==no_control:
                filas[i]=nuevos
                break
        escribir_csv(ALUMNOS_CSV, filas[0].keys(), filas)
        return redirect(url_for('alumnos_list'))
    return render_template('alumnos_form.html', alumno=alumno)

@app.route('/alumnos/delete/<no_control>')
def alumnos_delete(no_control):
    filas = [f for f in leer_csv(ALUMNOS_CSV) if f['no_control']!=no_control]
    if filas:
        escribir_csv(ALUMNOS_CSV, filas[0].keys(), filas)
    else:
        os.remove(ALUMNOS_CSV)
    return redirect(url_for('alumnos_list'))

# ── CRUD ÁREAS ─────────────────────────────────────────────────────────────────

@app.route('/areas')
def areas_list():
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(AREAS_CSV)
    return render_template('areas_list.html', areas=filas)

@app.route('/areas/new', methods=['GET','POST'])
def areas_new():
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    if request.method=='POST':
        datos = request.form.to_dict()
        datos['id'] = uuid.uuid4().hex
        filas = leer_csv(AREAS_CSV)
        filas.append(datos)
        campos = filas[0].keys() if filas else ['id','nombre','encargado']
        escribir_csv(AREAS_CSV, campos, filas)
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area={})

@app.route('/areas/edit/<id>', methods=['GET','POST'])
def areas_edit(id):
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(AREAS_CSV)
    area = buscar_por(AREAS_CSV,'id',id)
    if not area: return "No encontrado",404
    if request.method=='POST':
        nuevos = request.form.to_dict(); nuevos['id']=id
        for i,a in enumerate(filas):
            if a['id']==id:
                filas[i]=nuevos; break
        escribir_csv(AREAS_CSV, filas[0].keys(), filas)
        return redirect(url_for('areas_list'))
    return render_template('areas_form.html', area=area)

@app.route('/areas/delete/<id>')
def areas_delete(id):
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = [a for a in leer_csv(AREAS_CSV) if a['id']!=id]
    if filas:
        escribir_csv(AREAS_CSV, filas[0].keys(), filas)
    else:
        os.remove(AREAS_CSV)
    return redirect(url_for('areas_list'))

# ── CRUD PROFESORES ────────────────────────────────────────────────────────────

@app.route('/profesores')
def profesores_list():
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(PROFESORES_CSV)
    return render_template('profesores_list.html', profesores=filas)

@app.route('/profesores/new', methods=['GET','POST'])
def profesores_new():
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    if request.method=='POST':
        datos = request.form.to_dict()
        filas = leer_csv(PROFESORES_CSV)
        filas.append(datos)
        campos = filas[0].keys() if filas else ['correo','nombre','area']
        escribir_csv(PROFESORES_CSV, campos, filas)
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor={})

@app.route('/profesores/edit/<correo>', methods=['GET','POST'])
def profesores_edit(correo):
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(PROFESORES_CSV)
    prof = buscar_por(PROFESORES_CSV,'correo',correo)
    if not prof: return "No encontrado",404
    if request.method=='POST':
        nuevos = request.form.to_dict()
        for i,p in enumerate(filas):
            if p['correo']==correo:
                filas[i]=nuevos; break
        escribir_csv(PROFESORES_CSV, filas[0].keys(), filas)
        return redirect(url_for('profesores_list'))
    return render_template('profesores_form.html', profesor=prof)

@app.route('/profesores/delete/<correo>')
def profesores_delete(correo):
    if session.get('role') not in ['Administrador','Jefe']:
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = [p for p in leer_csv(PROFESORES_CSV) if p['correo']!=correo]
    if filas:
        escribir_csv(PROFESORES_CSV, filas[0].keys(), filas)
    else:
        os.remove(PROFESORES_CSV)
    return redirect(url_for('profesores_list'))

# ── CRUD USUARIOS ──────────────────────────────────────────────────────────────

@app.route('/usuarios')
def usuarios_list():
    if session.get('role')!='Administrador':
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(USUARIOS_CSV)
    return render_template('usuarios_list.html', usuarios=filas)

@app.route('/usuarios/new', methods=['GET','POST'])
def usuarios_new():
    if session.get('role')!='Administrador':
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    if request.method=='POST':
        datos = request.form.to_dict()
        filas = leer_csv(USUARIOS_CSV)
        filas.append(datos)
        campos = filas[0].keys() if filas else ['correo','rol','area']
        escribir_csv(USUARIOS_CSV, campos, filas)
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario={})

@app.route('/usuarios/edit/<correo>', methods=['GET','POST'])
def usuarios_edit(correo):
    if session.get('role')!='Administrador':
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = leer_csv(USUARIOS_CSV)
    usr = buscar_por(USUARIOS_CSV,'correo',correo)
    if not usr: return "No encontrado",404
    if request.method=='POST':
        nuevos = request.form.to_dict()
        for i,u in enumerate(filas):
            if u['correo']==correo:
                filas[i]=nuevos; break
        escribir_csv(USUARIOS_CSV, filas[0].keys(), filas)
        return redirect(url_for('usuarios_list'))
    return render_template('usuarios_form.html', usuario=usr)

@app.route('/usuarios/delete/<correo>')
def usuarios_delete(correo):
    if session.get('role')!='Administrador':
        flash('Acceso denegado'); return redirect(url_for('dashboard'))
    filas = [u for u in leer_csv(USUARIOS_CSV) if u['correo']!=correo]
    if filas:
        escribir_csv(USUARIOS_CSV, filas[0].keys(), filas)
    else:
        os.remove(USUARIOS_CSV)
    return redirect(url_for('usuarios_list'))

# ── Generación de Documentos específicos ───────────────────────────────────────

@app.route('/generar_solicitud/<no_control>')
def generar_solicitud(no_control):
    alumno = buscar_por(ALUMNOS_CSV,'no_control',no_control)
    if not alumno: return "No encontrado",404
    generar_docx(
        'plantilla_solicitud_completa.docx',
        alumno,
        f"Solicitud_{no_control}.docx"
    )
    flash('Solicitud generada.')
    return redirect(url_for('alumnos_list'))

@app.route('/generar_bimestral/<no_control>/<int:bi>')
def generar_bimestral(no_control,bi):
    alumno = buscar_por(ALUMNOS_CSV,'no_control',no_control)
    if not alumno: return "No encontrado",404
    generar_docx(
        'Reporte_Bimestral_Plantilla.docx',
        alumno,
        f"Reporte_Bim{bi}_{no_control}.docx"
    )
    flash(f'Bimestral {bi} generado.')
    return redirect(url_for('alumnos_list'))

@app.route('/generar_final/<no_control>')
def generar_final(no_control):
    alumno = buscar_por(ALUMNOS_CSV,'no_control',no_control)
    if not alumno: return "No encontrado",404
    generar_docx(
        'Reporte_Final_Lleno.docx',
        alumno,
        f"Reporte_Final_{no_control}.docx"
    )
    flash('Reporte final generado.')
    return redirect(url_for('alumnos_list'))

@app.route('/descargar/<path:fname>')
def descargar(fname):
    return send_from_directory(DOCUMENTOS_DIR, fname, as_attachment=True)

# ── Arranque ───────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    os.makedirs(CSV_DIR, exist_ok=True)
    os.makedirs(DOCUMENTOS_DIR, exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
