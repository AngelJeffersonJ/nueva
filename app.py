import os
import csv
import uuid
from io import BytesIO

from flask import (
    Flask, render_template, redirect, url_for,
    request, session, flash, send_file, abort, g
)
import msal
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "clave_segura")

# --- Configuración Azure AD ---
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv("REDIRECT_URI")
SCOPE         = ["User.Read"]

# --- Rutas a CSV ---
USUARIOS_CSV    = "csv/usuarios.csv"
AREAS_CSV       = "csv/areas.csv"
PROFESORES_CSV  = "csv/profesores.csv"
ALUMNOS_CSV     = "csv/alumnos.csv"
DOCUMENTOS_CSV  = "csv/documentos.csv"

# --- Plantillas DOCX ---
TEMPLATE_SOLICITUD   = "plantillas/plantilla_solicitud_completa.docx"
TEMPLATE_BIMESTRAL   = "plantillas/Reporte_Bimestral_Plantilla.docx"
TEMPLATE_FINAL       = "plantillas/Reporte_Final_Plantilla.docx"


def cargar_csv(path):
    """Lee todo el CSV y devuelve lista de dicts."""
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def guardar_csv(path, rows, fields):
    """Reescribe el CSV con rows (lista de dict) y cabecera fields."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower() == email.lower():
            return u
    return None


# ─── Autenticación ─────────────────────────────────────────────────────────────

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
    if not code:
        return "Parámetro code faltante", 400
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if "error" in result:
        return f"Error autenticación: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email = claims.get("preferred_username", "").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {
        "correo": email,
        "rol":    "Alumno",
        "area":   ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol":     perfil["rol"],
        "area":    perfil.get("area",""),
        "name":    claims.get("name", perfil["correo"])
    }
    return redirect(url_for("dashboard"))


@app.route("/logout")
def logout():
    session.clear()
    post = url_for("login", _external=True)
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={post}")


def validar_acceso(roles):
    u = session.get("user")
    return u and u.get("rol") in roles


# ─── Dashboard ────────────────────────────────────────────────────────────────

@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template(
        "dashboard.html",
        usuario=session["user"]
    )


# ─── CRUD Usuarios ────────────────────────────────────────────────────────────

@app.route("/usuarios")
def usuarios_list():
    if not validar_acceso(["Administrador"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = cargar_csv(USUARIOS_CSV)
    return render_template("usuarios_list.html", usuarios=rows)


@app.route("/usuarios/new", methods=["GET","POST"])
def usuarios_new():
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("usuarios_list"))
    if request.method == "POST":
        rows = cargar_csv(USUARIOS_CSV)
        rows.append({
            "correo": request.form["correo"],
            "rol":    request.form["rol"],
            "area":   request.form["area"]
        })
        guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        flash("Usuario creado", "success")
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario={})


@app.route("/usuarios/edit/<correo>", methods=["GET","POST"])
def usuarios_edit(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("usuarios_list"))
    rows = cargar_csv(USUARIOS_CSV)
    u = next((r for r in rows if r["correo"]==correo), None)
    if not u:
        flash("No existe", "danger")
        return redirect(url_for("usuarios_list"))
    if request.method == "POST":
        u["rol"]  = request.form["rol"]
        u["area"] = request.form["area"]
        guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        flash("Usuario actualizado", "success")
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario=u)


@app.route("/usuarios/delete/<correo>", methods=["POST"])
def usuarios_delete(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("usuarios_list"))
    rows = [r for r in cargar_csv(USUARIOS_CSV) if r["correo"]!=correo]
    guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
    flash("Usuario eliminado", "success")
    return redirect(url_for("usuarios_list"))


# ─── CRUD Áreas ───────────────────────────────────────────────────────────────

@app.route("/areas")
def areas_list():
    if not validar_acceso(["Administrador","Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = cargar_csv(AREAS_CSV)
    return render_template("areas_list.html", areas=rows)


@app.route("/areas/new", methods=["GET","POST"])
def areas_new():
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("areas_list"))
    if request.method=="POST":
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            "id":        str(uuid.uuid4()),
            "nombre":    request.form["nombre"],
            "encargado": request.form["encargado"]
        })
        guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        flash("Área creada", "success")
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area={})


@app.route("/areas/edit/<id>", methods=["GET","POST"])
def areas_edit(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("areas_list"))
    rows = cargar_csv(AREAS_CSV)
    a = next((r for r in rows if r["id"]==id), None)
    if not a:
        flash("No existe", "danger")
        return redirect(url_for("areas_list"))
    if request.method=="POST":
        a["nombre"]    = request.form["nombre"]
        a["encargado"] = request.form["encargado"]
        guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        flash("Área actualizada", "success")
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area=a)


@app.route("/areas/delete/<id>", methods=["POST"])
def areas_delete(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("areas_list"))
    rows = [r for r in cargar_csv(AREAS_CSV) if r["id"]!=id]
    guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
    flash("Área eliminada", "success")
    return redirect(url_for("areas_list"))


# ─── CRUD Profesores ──────────────────────────────────────────────────────────

@app.route("/profesores")
def profesores_list():
    if not validar_acceso(["Administrador","Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = cargar_csv(PROFESORES_CSV)
    return render_template("profesores_list.html", profesores=rows)


@app.route("/profesores/new", methods=["GET","POST"])
def profesores_new():
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("profesores_list"))
    if request.method=="POST":
        rows = cargar_csv(PROFESORES_CSV)
        rows.append({
            "correo": request.form["correo"],
            "nombre": request.form["nombre"],
            "area":   request.form["area"]
        })
        guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        flash("Profesor creado", "success")
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor={})


@app.route("/profesores/edit/<correo>", methods=["GET","POST"])
def profesores_edit(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("profesores_list"))
    rows = cargar_csv(PROFESORES_CSV)
    p = next((r for r in rows if r["correo"]==correo), None)
    if not p:
        flash("No existe", "danger")
        return redirect(url_for("profesores_list"))
    if request.method=="POST":
        p["nombre"] = request.form["nombre"]
        p["area"]   = request.form["area"]
        guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        flash("Profesor actualizado", "success")
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor=p)


@app.route("/profesores/delete/<correo>", methods=["POST"])
def profesores_delete(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("profesores_list"))
    rows = [r for r in cargar_csv(PROFESORES_CSV) if r["correo"]!=correo]
    guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
    flash("Profesor eliminado", "success")
    return redirect(url_for("profesores_list"))


# ─── CRUD Alumnos ─────────────────────────────────────────────────────────────

@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))

    rows = cargar_csv(ALUMNOS_CSV)
    rol  = session["user"]["rol"]
    if rol=="Maestro":
        rows = [r for r in rows if r["profesor"]==session["user"]["correo"]]
    if rol=="Encargado":
        rows = [r for r in rows if r["area"]==session["user"]["area"]]

    return render_template("alumnos_list.html", alumnos=rows)


@app.route("/alumnos/new", methods=["GET","POST"])
def alumnos_new():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("alumnos_list"))
    if request.method=="POST":
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            "no_control": request.form["no_control"],
            "nombre":     request.form["nombre"],
            "area":       request.form["area"],
            "profesor":   request.form["profesor"]
        })
        guardar_csv(ALUMNOS_CSV, rows, ["no_control","nombre","area","profesor"])
        flash("Alumno creado", "success")
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno={})


@app.route("/alumnos/edit/<no_control>", methods=["GET","POST"])
def alumnos_edit(no_control):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("alumnos_list"))
    rows = cargar_csv(ALUMNOS_CSV)
    a = next((r for r in rows if r["no_control"]==no_control), None)
    if not a:
        flash("No existe", "danger")
        return redirect(url_for("alumnos_list"))
    if request.method=="POST":
        a["nombre"]   = request.form["nombre"]
        a["area"]     = request.form["area"]
        a["profesor"] = request.form["profesor"]
        guardar_csv(ALUMNOS_CSV, rows, ["no_control","nombre","area","profesor"])
        flash("Alumno actualizado", "success")
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno=a)


@app.route("/alumnos/delete/<no_control>", methods=["POST"])
def alumnos_delete(no_control):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("alumnos_list"))
    rows = [r for r in cargar_csv(ALUMNOS_CSV) if r["no_control"]!=no_control]
    guardar_csv(ALUMNOS_CSV, rows, ["no_control","nombre","area","profesor"])
    flash("Alumno eliminado", "success")
    return redirect(url_for("alumnos_list"))


# ─── Lista y generación de documentos ─────────────────────────────────────────

@app.route("/documentos")
def documentos_list():
    if "user" not in session:
        return redirect(url_for("login"))
    docs = cargar_csv(DOCUMENTOS_CSV)
    return render_template("documentos_list.html", documentos=docs)


def _make_doc(template_path, context, filename):
    doc = DocxTemplate(template_path)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, download_name=filename, as_attachment=True)


@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    row = next((d for d in docs if d.get("No_Control")==no_control), None)
    if not row:
        flash("Datos de documento no encontrados", "danger")
        return redirect(url_for("documentos_list"))
    return _make_doc(
        TEMPLATE_SOLICITUD,
        row,
        f"Solicitud_{no_control}.docx"
    )


@app.route("/generar_bimestral/<bimestre>/<no_control>")
def generar_bimestral(bimestre, no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    row = next((d for d in docs if d.get("No_Control")==no_control), None)
    if not row:
        flash("Datos de documento no encontrados", "danger")
        return redirect(url_for("documentos_list"))
    row["Reporte_No"] = bimestre
    return _make_doc(
        TEMPLATE_BIMESTRAL,
        row,
        f"Reporte_Bimestral_{no_control}_{bimestre}.docx"
    )


@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    docs = cargar_csv(DOCUMENTOS_CSV)
    row = next((d for d in docs if d.get("No_Control")==no_control), None)
    if not row:
        flash("Datos de documento no encontrados", "danger")
        return redirect(url_for("documentos_list"))
    return _make_doc(
        TEMPLATE_FINAL,
        row,
        f"Reporte_Final_{no_control}.docx"
    )


# ─── Manejador 404 ─────────────────────────────────────────────────────────────

@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=True)
