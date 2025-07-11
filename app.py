import os
import csv
import uuid
from io import BytesIO
from functools import wraps

from flask import (
    Flask, render_template, redirect, url_for,
    session, request, flash, send_file
)
from dotenv import load_dotenv
import msal
from docxtpl import DocxTemplate

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'supersecret')

# ── Configuración Azure AD ─────────────────────────────────────────────────────
CLIENT_ID     = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID     = os.getenv('TENANT_ID')
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv('REDIRECT_URI')
SCOPE         = ['User.Read']

# ── Rutas de CSV ────────────────────────────────────────────────────────────────
CSV_DIR        = os.path.join(app.root_path, 'csv')
USERS_CSV      = os.path.join(CSV_DIR, 'usuarios.csv')
AREAS_CSV      = os.path.join(CSV_DIR, 'areas.csv')
PROFESORES_CSV = os.path.join(CSV_DIR, 'profesores.csv')
ALUMNOS_CSV    = os.path.join(CSV_DIR, 'alumnos_completo.csv')
DOCUMENTOS_CSV = os.path.join(CSV_DIR, 'documentos_completo.csv')

# ── Plantillas DOCX ────────────────────────────────────────────────────────────
TPL_DIR        = os.path.join(app.root_path, 'plantillas')
TPL_SOLICITUD  = os.path.join(TPL_DIR, 'plantilla_solicitud_completa.docx')
TPL_BIMESTRAL  = os.path.join(TPL_DIR, 'Reporte_Bimestral_Plantilla.docx')
TPL_FINAL      = os.path.join(TPL_DIR, 'Reporte_Final_Lleno.docx')

# ── Campos para documentos_completo.csv ────────────────────────────────────────
DOC_FIELDS = [
    'No_Control',
    # Solicitud
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos',
    'Dependencia','Domicilio_Dependencia','Titular_Dependencia','Cargo_Responsable',
    'Responsable_Proyecto','Nombre_Programa',
    # TP_ radio: solo uno debe marcarse con 'X'
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros',
    'Fecha_Inicio','Fecha_Terminacion','Actividades',
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Bimestrales fechas
] + [f'{part}_{n}' for n in (1,2,3) for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2')] + [
    'Nombre_supervisor','Puesto_supervisor'
] + [f'Actividad_{i}' for i in range(1,9)] + [
    'x1','x2','x3'
] + [f'resp{n}_{i}' for n in range(3,18) for i in range(5)] + [
] + [f'estu{n}_{i}' for n in range(3,18) for i in range(5)] + [
    # Final
    'Municipio','Estado'
] + [f'Actividad{i}' for i in range(1,9)] + [f'Logro{i}' for i in range(1,9)] + [
] + [f'Aprendizaje{i}' for i in range(1,9)] + [f'Beneficio{i}' for i in range(1,9)] + [
    'Nombre_Responsable','Cargo_Responsable','Num_Control'
]

# ── Funciones auxiliares ───────────────────────────────────────────────────────
def cargar_csv(path):
    if not os.path.isfile(path):
        return []
    try:
        with open(path, newline='', encoding='utf-8') as f:
            return list(csv.DictReader(f))
    except Exception as e:
        app.logger.error(f"Error cargando CSV {path}: {e}")
        return []

def guardar_csv(path, rows, fieldnames):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows)
    except Exception as e:
        app.logger.error(f"Error guardando CSV {path}: {e}")
        flash("No se pudo guardar el CSV", "danger")

def login_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapped

def roles_required(*roles):
    def deco(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if session.get('role') not in roles:
                flash("Acceso denegado", "danger")
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapped
    return deco

def _render_docx(template_path, context, filename):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generando DOCX: {e}")
        flash("Error generando el documento", "danger")
        return redirect(url_for('dashboard'))

# ── Autenticación y CRUD Usuarios/Áreas/Profesores/Alumnos (sin cambios) ──
# Omite aquí por brevedad; usa tu código existente.

# ── Solicitud (GET/POST) ───────────────────────────────────────────────────────
SOL_TEXT_FIELDS = [
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia',
    'Titular_Dependencia','Cargo_Responsable','Responsable_Proyecto','Nombre_Programa',
    'Fecha_Inicio','Fecha_Terminacion','Actividades'
]
SOL_TP_FIELDS = [
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros'
]

@app.route('/documentos/solicitud/<no_control>', methods=['GET','POST'])
@login_required
def documentos_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method == 'POST':
        # validar todos los textos
        faltan = [f for f in SOL_TEXT_FIELDS if not request.form.get(f)]
        # validar TP_ radio: exactamente 1 marcada
        tp_sel = [f for f in SOL_TP_FIELDS if request.form.get(f)]
        if len(tp_sel) != 1:
            faltan += ['TP (un solo tipo de proyecto)']
        # validar fecha solicitud
        for part in ('Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'):
            if not request.form.get(part):
                faltan.append(part)
        if faltan:
            flash("Completa todos los campos: " + ", ".join(faltan), "warning")
            return redirect(request.url)
        # crear o actualizar
        if not doc:
            doc = {k: '' for k in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        # guardar textos
        for f in SOL_TEXT_FIELDS:
            doc[f] = request.form[f].strip()
        # guardar TP_
        for f in SOL_TP_FIELDS:
            doc[f] = 'X' if f in tp_sel else ''
        # guardar fecha solicitud
        for part in ('Dia_Solicitud','Mes_Solicitud','Anio_Solicitud'):
            doc[part] = request.form[part].strip()
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Solicitud guardada", "success")
        return redirect(url_for('documentos_list'))
    return render_template('solicitud_form.html', no_control=no_control, documento=doc or {})

# ── Bimestral (GET/POST) ──────────────────────────────────────────────────────
@app.route('/documentos/bimestral/<no_control>/<int:num>', methods=['GET','POST'])
@login_required
def documentos_bimestral(no_control, num):
    if num not in (1,2,3):
        flash("Bimestre inválido", "warning")
        return redirect(url_for('documentos_list'))
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method == 'POST':
        faltan = []
        # fechas
        for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2'):
            if not request.form.get(f"{part}_{num}"):
                faltan.append(f"{part}_{num}")
        # supervisor y actividades
        if not request.form.get('Nombre_supervisor'):
            faltan.append('Nombre_supervisor')
        if not request.form.get('Puesto_supervisor'):
            faltan.append('Puesto_supervisor')
        for i in range(1,9):
            if not request.form.get(f"Actividad_{i}"):
                faltan.append(f"Actividad_{i}")
        # bimestre radio
        if not request.form.get('bim'):
            faltan.append('bim')
        # evaluaciones
        for n in range(3,18):
            if not request.form.get(f"resp{n}"):
                faltan.append(f"resp{n}")
            if not request.form.get(f"estu{n}"):
                faltan.append(f"estu{n}")
        if faltan:
            flash("Completa todos los campos: " + ", ".join(faltan), "warning")
            return redirect(request.url)
        # crear/actualizar
        if not doc:
            doc = {k: '' for k in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        # fechas
        for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2'):
            doc[f"{part}_{num}"] = request.form[f"{part}_{num}"]
        # supervisor y actividades
        doc['Nombre_supervisor'] = request.form['Nombre_supervisor']
        doc['Puesto_supervisor'] = request.form['Puesto_supervisor']
        for i in range(1,9):
            doc[f"Actividad_{i}"] = request.form[f"Actividad_{i}"]
        # bimestre
        sel = request.form['bim']
        doc['x1'], doc['x2'], doc['x3'] = (
            'X' if sel=='1' else '',
            'X' if sel=='2' else '',
            'X' if sel=='3' else ''
        )
        # evaluaciones
        for n in range(3,18):
            for i in range(5):
                doc[f"resp{n}_{i}"] = ''
                doc[f"estu{n}_{i}"] = ''
            dr = int(request.form[f"resp{n}"])
            de = int(request.form[f"estu{n}"])
            doc[f"resp{n}_{dr}"] = 'X'
            doc[f"estu{n}_{de}"] = 'X'
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash(f"Bimestral {num} guardado", "success")
        return redirect(url_for('documentos_list'))
    return render_template(
        'bimestral_form.html',
        no_control=no_control,
        num=num,
        documento=doc or {}
    )

# ── Final (GET/POST) ──────────────────────────────────────────────────────────
FINAL_FIELDS = (
    ['Municipio','Estado'] +
    [f'Actividad{i}' for i in range(1,9)] +
    [f'Logro{i}'     for i in range(1,9)] +
    [f'Aprendizaje{i}' for i in range(1,9)] +
    [f'Beneficio{i}'   for i in range(1,9)] +
    ['Nombre_Responsable','Cargo_Responsable','Num_Control']
)

@app.route('/documentos/final/<no_control>', methods=['GET','POST'])
@login_required
def documentos_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    doc  = next((d for d in docs if d['No_Control']==no_control), None)
    if request.method == 'POST':
        faltan = [f for f in FINAL_FIELDS if not request.form.get(f)]
        if faltan:
            flash("Completa todos los campos: " + ", ".join(faltan), "warning")
            return redirect(request.url)
        if not doc:
            doc = {k: '' for k in DOC_FIELDS}
            doc['No_Control'] = no_control
            docs.append(doc)
        for f in FINAL_FIELDS:
            doc[f] = request.form[f].strip()
        guardar_csv(DOCUMENTOS_CSV, docs, DOC_FIELDS)
        flash("Reporte final guardado", "success")
        return redirect(url_for('documentos_list'))
    return render_template(
        'final_form.html',
        no_control=no_control,
        documento=doc or {}
    )

# ── Generación de DOCX ───────────────────────────────────────────────────────
@app.route('/generar_solicitud/<no_control>')
@login_required
def generar_solicitud(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_SOLICITUD,
        {**alum, **doci},
        f"Solicitud_{no_control}.docx"
    )

@app.route('/generar_bimestral/<no_control>/<int:num>')
@login_required
def generar_bimestral(no_control, num):
    if num not in (1,2,3):
        flash("Bimestre inválido", "warning")
        return redirect(url_for('documentos_list'))
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_BIMESTRAL,
        {**alum, **doci, 'reporte_no': num},
        f"Bimestral_{no_control}_B{num}.docx"
    )

@app.route('/generar_final/<no_control>')
@login_required
def generar_final(no_control):
    alumnos = cargar_csv(ALUMNOS_CSV)
    docs    = cargar_csv(DOCUMENTOS_CSV)
    alum    = next((a for a in alumnos if a['No_Control']==no_control), None)
    doci    = next((d for d in docs    if d['No_Control']==no_control), None)
    if not alum or not doci:
        flash("Datos incompletos", "warning")
        return redirect(url_for('documentos_list'))
    return _render_docx(
        TPL_FINAL,
        {**alum, **doci},
        f"Final_{no_control}.docx"
    )

# ── Error 404 ─────────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__=='__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT',5000)))
