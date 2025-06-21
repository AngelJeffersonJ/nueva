from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from deta import Deta
import os, zipfile, uuid
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate
from werkzeug.security import generate_password_hash, check_password_hash

# Init App & Deta
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET', 'clave_segura')

deta = Deta(os.getenv('DETA_PROJECT_KEY'))
usuarios_db   = deta.Base('usuarios')
profesores_db = deta.Base('profesores')
areas_db      = deta.Base('areas')
alumnos_db    = deta.Base('alumnos')

# Templates & Output
TEMPLATE_PATH = 'plantillas/plantilla_solicitud_completa.docx'
OUTPUT_PATH   = 'documentos_generados'
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Auth helpers
def usuario_desde_db(email):
    res = usuarios_db.fetch({'correo': email}).items
    return res[0] if res else None

# Login local (reemplazo de MSAL)
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['correo'].lower()
        password = request.form['password']
        user = usuario_desde_db(email)
        if user and check_password_hash(user.get('password_hash', ''), password):
            session['user'] = user
            return redirect(url_for('dashboard'))
        else:
            flash('Correo o contraseña incorrectos', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# Dashboard
@app.route('/')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', usuario=session['user'])

# CRUD for any entity
def get_base(tipo):
    return {
        'usuarios'  : usuarios_db,
        'profesores': profesores_db,
        'areas'     : areas_db,
        'alumnos'   : alumnos_db
    }.get(tipo)

@app.route('/entidad/<tipo>')
def entidad_list(tipo):
    base = get_base(tipo)
    if not base:
        return render_template('404.html'), 404
    items = base.fetch().items
    campos = items[0].keys() if items else []
    return render_template('entidad_list.html', tipo=tipo, campos=campos, registros=items)

@app.route('/entidad/<tipo>/new', methods=['GET','POST'])
def entidad_new(tipo):
    base = get_base(tipo)
    if not base:
        return render_template('404.html'), 404
    if request.method == 'POST':
        data = request.form.to_dict()
        data['key'] = request.form.get('key') or str(uuid.uuid4())
        if tipo == 'usuarios' and 'password' in data:
            data['password_hash'] = generate_password_hash(data['password'])
            del data['password']
        base.put(data)
        return redirect(url_for('entidad_list', tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo, valores={})

@app.route('/entidad/<tipo>/edit/<key>', methods=['GET','POST'])
def entidad_edit(tipo, key):
    base = get_base(tipo)
    if not base:
        return render_template('404.html'), 404
    item = base.get(key)
    if not item:
        flash('No encontrado', 'danger')
        return redirect(url_for('entidad_list', tipo=tipo))
    if request.method == 'POST':
        data = request.form.to_dict()
        data['key'] = key
        if tipo == 'usuarios' and 'password' in data and data['password']:
            data['password_hash'] = generate_password_hash(data['password'])
            del data['password']
        else:
            data['password_hash'] = item.get('password_hash')  # conservar el hash actual
        base.put(data)
        return redirect(url_for('entidad_list', tipo=tipo))
    return render_template('entidad_form.html', tipo=tipo, valores=item)

@app.route('/entidad/<tipo>/delete/<key>', methods=['POST'])
def entidad_delete(tipo, key):
    base = get_base(tipo)
    if not base:
        return render_template('404.html'), 404
    base.delete(key)
    return redirect(url_for('entidad_list', tipo=tipo))

# Generate & Download documents
@app.route('/generar-documentos')
def generar():
    items = alumnos_db.fetch().items
    os.makedirs(OUTPUT_PATH, exist_ok=True)
    for alum in items:
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(alum)
        doc.save(os.path.join(OUTPUT_PATH, f"Solicitud_{alum['No_Control']}.docx"))
    flash('Documentos generados', 'success')
    return redirect(url_for('dashboard'))

@app.route('/descargar-documentos')
def descargar():
    mem = BytesIO()
    with zipfile.ZipFile(mem, 'w') as zf:
        for fn in os.listdir(OUTPUT_PATH):
            zf.write(os.path.join(OUTPUT_PATH, fn), fn)
    mem.seek(0)
    return send_file(mem, download_name='documentos.zip', as_attachment=True)

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
