import os
import csv
import uuid
import zipfile
import logging
from io import BytesIO
from datetime import datetime

from flask import (
    Flask, render_template, session, redirect, url_for,
    request, flash, send_file
)
from dotenv import load_dotenv
import msal
from docxtpl import DocxTemplate

# Carga .env
load_dotenv()

# Configuración básica
app = Flask(__name__, static_folder="static")
app.secret_key = os.getenv("FLASK_SECRET", "clave_segura")

# Logging
logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s %(levelname)s %(message)s")
app.logger.setLevel(logging.DEBUG)

# Azure AD (MSAL)
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv("REDIRECT_URI")  # e.g. https://tu-dominio/getAToken
SCOPE         = ["User.Read"]

# Rutas a CSV
USUARIOS_CSV   = "csv/usuarios.csv"
AREAS_CSV      = "csv/areas.csv"
PROFESORES_CSV = "csv/profesores.csv"
ALUMNOS_CSV    = "csv/alumnos.csv"

# Plantillas
TEMPLATES = {
    "solicitud":       "plantillas/plantilla_solicitud_completa.docx",
    "bimestral":       "plantillas/Reporte_Bimestral_Plantilla.docx",
    "final":           "plantillas/Reporte_Final_Plantilla.docx",
    # añade más si tienes otras plantillas
}

# Helper para CSV
def load_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def save_csv(path, rows, fieldnames):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

def find_by_key(rows, key_name, key_val):
    return next((r for r in rows if r.get(key_name)==key_val), None)

# Antes de cada petición: saludos, HEAD y protección de sesión
@app.before_request
def before():
    # 1) HEAD de health-check
    if request.method == "HEAD":
        return "", 200
    # 2) Rutas públicas
    public = {"/login", "/getAToken", "/static"}
    if any(request.path.startswith(p) for p in public):
        return
    # 3) Sesión requerida
    if not session.get("user"):
        app.logger.debug("No session, redirigiendo a login")
        return redirect(url_for("login"))

# --- Autenticación MSAL ---
@app.route("/login")
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if "error" in result:
        flash(result.get("error_description", "Error de autenticación"), "danger")
        return redirect(url_for("login"))

    claims = result.get("id_token_claims", {})
    email = claims.get("preferred_username", "").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales", 403

    # Carga perfil desde CSV o crea uno por defecto
    usuarios = load_csv(USUARIOS_CSV)
    perfil = find_by_key(usuarios, "correo", email) or {
        "correo": email, "rol": "Alumno", "area": ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol":     perfil["rol"],
        "area":    perfil.get("area",""),
        "name":    claims.get("name", perfil["correo"])
    }
    app.logger.info("Usuario %s ha iniciado sesión", email)
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    user = session.pop("user", None)
    app.logger.info("Usuario desconectado: %s", user and user.get("correo"))
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# --- Dashboard ---
@app.route("/")
@app.route("/dashboard")
def dashboard():
    user = session.get("user")
    return render_template("dashboard.html", usuario=user)

# --- CRUD Usuarios ---
@app.route("/usuarios")
def usuarios_list():
    if session["user"]["rol"]!="Administrador":
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = load_csv(USUARIOS_CSV)
    return render_template("usuarios_list.html", usuarios=rows)

@app.route("/usuarios/new", methods=["GET","POST"])
def usuarios_new():
    if session["user"]["rol"]!="Administrador":
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = load_csv(USUARIOS_CSV)
        nuevo = {
            "correo": request.form["correo"],
            "rol":    request.form["rol"],
            "area":   request.form["area"]
        }
        rows.append(nuevo)
        save_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario={})

@app.route("/usuarios/edit/<correo>", methods=["GET","POST"])
def usuarios_edit(correo):
    if session["user"]["rol"]!="Administrador":
        return redirect(url_for("dashboard"))
    rows = load_csv(USUARIOS_CSV)
    u = find_by_key(rows, "correo", correo)
    if not u: return "No existe", 404
    if request.method=="POST":
        u["rol"]  = request.form["rol"]
        u["area"] = request.form["area"]
        save_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario=u)

@app.route("/usuarios/delete/<correo>", methods=["POST"])
def usuarios_delete(correo):
    if session["user"]["rol"]!="Administrador":
        return redirect(url_for("dashboard"))
    rows = [r for r in load_csv(USUARIOS_CSV) if r["correo"]!=correo]
    save_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
    return redirect(url_for("usuarios_list"))

# --- CRUD Áreas ---
@app.route("/areas")
def areas_list():
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = load_csv(AREAS_CSV)
    return render_template("areas_list.html", areas=rows)

@app.route("/areas/new", methods=["GET","POST"])
def areas_new():
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = load_csv(AREAS_CSV)
        nuevo = {
            "id": str(uuid.uuid4()),
            "nombre":    request.form["nombre"],
            "encargado": request.form["encargado"]
        }
        rows.append(nuevo)
        save_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area={})

@app.route("/areas/edit/<id>", methods=["GET","POST"])
def areas_edit(id):
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    rows = load_csv(AREAS_CSV)
    area = find_by_key(rows, "id", id)
    if not area: return "No existe",404
    if request.method=="POST":
        area["nombre"]    = request.form["nombre"]
        area["encargado"] = request.form["encargado"]
        save_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area=area)

@app.route("/areas/delete/<id>", methods=["POST"])
def areas_delete(id):
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    rows = [r for r in load_csv(AREAS_CSV) if r["id"]!=id]
    save_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
    return redirect(url_for("areas_list"))

# --- CRUD Profesores ---
@app.route("/profesores")
def profesores_list():
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = load_csv(PROFESORES_CSV)
    return render_template("profesores_list.html", profesores=rows)

@app.route("/profesores/new", methods=["GET","POST"])
def profesores_new():
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = load_csv(PROFESORES_CSV)
        nuevo = {
            "correo": request.form["correo"],
            "nombre": request.form["nombre"],
            "area":   request.form["area"]
        }
        rows.append(nuevo)
        save_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor={})

@app.route("/profesores/edit/<correo>", methods=["GET","POST"])
def profesores_edit(correo):
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    rows = load_csv(PROFESORES_CSV)
    prof = find_by_key(rows, "correo", correo)
    if not prof: return "No existe",404
    if request.method=="POST":
        prof["nombre"] = request.form["nombre"]
        prof["area"]   = request.form["area"]
        save_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor=prof)

@app.route("/profesores/delete/<correo>", methods=["POST"])
def profesores_delete(correo):
    if session["user"]["rol"] not in ("Administrador","Encargado"):
        return redirect(url_for("dashboard"))
    rows = [r for r in load_csv(PROFESORES_CSV) if r["correo"]!=correo]
    save_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
    return redirect(url_for("profesores_list"))

# --- CRUD Alumnos + Generación de documentos ---
@app.route("/alumnos")
def alumnos_list():
    if session["user"]["rol"] not in ("Administrador","Encargado","Maestro"):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = load_csv(ALUMNOS_CSV)
    # Filtrado según rol
    rol = session["user"]["rol"]
    if rol=="Maestro":
        rows = [r for r in rows if r.get("profesor")==session["user"]["correo"]]
    if rol=="Encargado":
        rows = [r for r in rows if r.get("area")==session["user"]["area"]]
    return render_template("alumnos_list.html", alumnos=rows)

@app.route("/alumnos/new", methods=["GET","POST"])
def alumnos_new():
    if session["user"]["rol"] not in ("Administrador","Encargado","Maestro"):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = load_csv(ALUMNOS_CSV)
        nuevo = {
            "no_control": request.form["no_control"],
            "apellido_paterno": request.form["apellido_paterno"],
            "apellido_materno": request.form["apellido_materno"],
            "nombre": request.form["nombre"],
            "carrera": request.form["carrera"],
            "periodo": request.form["periodo"],
            "semestre": request.form["semestre"],
            "creditos": request.form["creditos"],
            "area": request.form["area"],
            "profesor": request.form["profesor"],
            # agrega más campos según tu CSV y plantillas
        }
        rows.append(nuevo)
        save_csv(ALUMNOS_CSV, rows, nuevo.keys())
        return redirect(url_for("alumnos_list"))
    # Para GET, pasamos listas para selects
    return render_template("alumnos_form.html", alumno={})

@app.route("/alumnos/edit/<no_control>", methods=["GET","POST"])
def alumnos_edit(no_control):
    if session["user"]["rol"] not in ("Administrador","Encargado","Maestro"):
        return redirect(url_for("dashboard"))
    rows = load_csv(ALUMNOS_CSV)
    alum = find_by_key(rows, "no_control", no_control)
    if not alum: return "No existe",404
    if request.method=="POST":
        for k in alum.keys():
            if k in request.form:
                alum[k] = request.form[k]
        save_csv(ALUMNOS_CSV, rows, alum.keys())
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno=alum)

@app.route("/alumnos/delete/<no_control>", methods=["POST"])
def alumnos_delete(no_control):
    if session["user"]["rol"] not in ("Administrador","Encargado","Maestro"):
        return redirect(url_for("dashboard"))
    rows = [r for r in load_csv(ALUMNOS_CSV) if r["no_control"]!=no_control]
    save_csv(ALUMNOS_CSV, rows, rows[0].keys() if rows else [])
    return redirect(url_for("alumnos_list"))

# --- Generación de documentos individuales ---
def generate_doc(template_key, context):
    tpl_path = TEMPLATES.get(template_key)
    if not tpl_path or not os.path.exists(tpl_path):
        return None
    doc = DocxTemplate(tpl_path)
    doc.render(context)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

@app.route("/alumnos/<no_control>/generate/<tipo>")
def alumnos_generate(no_control, tipo):
    # tipo en {"solicitud","bimestral","final",...}
    rows = load_csv(ALUMNOS_CSV)
    alum = find_by_key(rows, "no_control", no_control)
    if not alum:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("alumnos_list"))

    # añade campos comunes al contexto
    ctx = dict(alum)
    ctx.setdefault("Fecha_Inicio", datetime.now().strftime("%d/%m/%Y"))
    ctx.setdefault("Fecha_Terminacion", (datetime.now()).strftime("%d/%m/%Y"))
    # ... más defaults si quieres

    bio = generate_doc(tipo, ctx)
    if not bio:
        flash("Plantilla no encontrada", "danger")
        return redirect(url_for("alumnos_list"))
    filename = f"{no_control}_{tipo}.docx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# --- Descargar todos los docs de una vez ---
@app.route("/alumnos/download_all")
def alumnos_download_all():
    rows = load_csv(ALUMNOS_CSV)
    mem = BytesIO()
    with zipfile.ZipFile(mem, "w") as zf:
        for alum in rows:
            for tipo in TEMPLATES:
                ctx = dict(alum)
                bio = generate_doc(tipo, ctx)
                if bio:
                    zf.writestr(f"{alum['no_control']}_{tipo}.docx", bio.read())
                    bio.seek(0)
    mem.seek(0)
    return send_file(
        mem, as_attachment=True,
        download_name="todos_documentos.zip",
        mimetype="application/zip"
    )

# --- Error 404 genérico ---
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    # Para desarrollo: auto-reload; en producción usa Gunicorn
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=True)
