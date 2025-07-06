import os
import csv
import uuid
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from docxtpl import DocxTemplate
import msal

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# --- Configuración Azure AD ---
CLIENT_ID     = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
TENANT_ID     = os.environ.get("TENANT_ID", "")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI", "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# --- Rutas a CSV ---
USUARIOS_CSV   = "csv/usuarios.csv"
AREAS_CSV      = "csv/areas.csv"
PROFESORES_CSV = "csv/profesores.csv"
ALUMNOS_CSV    = "csv/alumnos.csv"
DOCUMENTOS_CSV = "csv/documentos.csv"

# --- Helpers CSV ---
def cargar_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def guardar_csv(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u["correo"].strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    user = session.get("user")
    return user and user.get("rol") in roles

# --- Autenticación ---
@app.route("/login")
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        SCOPE, redirect_uri=REDIRECT_URI
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
        return redirect(url_for("dashboard"))

    claims = result.get("id_token_claims", {})
    email = claims.get("preferred_username", "").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales", 403

    perfil = usuario_desde_csv(email) or {
        "correo": email, "rol": "Alumno", "area": ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol":     perfil["rol"],
        "area":    perfil.get("area", ""),
        "name":    claims.get("name", perfil["correo"])
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    post = url_for("login", _external=True)
    return redirect(f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={post}")

# --- Dashboard ---
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template(
        "dashboard.html",
        usuario=session["user"]
    )

# --- CRUD Usuarios (solo Admin) ---
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
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        rows = cargar_csv(USUARIOS_CSV)
        rows.append({
            "correo": request.form["correo"],
            "rol":    request.form["rol"],
            "area":   request.form["area"]
        })
        guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario={})

@app.route("/usuarios/edit/<correo>", methods=["GET","POST"])
def usuarios_edit(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(USUARIOS_CSV)
    user = next((r for r in rows if r["correo"]==correo), None)
    if not user: return "No existe",404
    if request.method=="POST":
        user["rol"]  = request.form["rol"]
        user["area"] = request.form["area"]
        guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario=user)

@app.route("/usuarios/delete/<correo>", methods=["POST"])
def usuarios_delete(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(USUARIOS_CSV) if r["correo"]!=correo]
    guardar_csv(USUARIOS_CSV, rows, ["correo","rol","area"])
    return redirect(url_for("usuarios_list"))

# --- CRUD Áreas (Admin, Encargado) ---
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
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            "id":        str(uuid.uuid4()),
            "nombre":    request.form["nombre"],
            "encargado": request.form["encargado"]
        })
        guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area={})

@app.route("/areas/edit/<id>", methods=["GET","POST"])
def areas_edit(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(AREAS_CSV)
    area = next((r for r in rows if r["id"]==id), None)
    if not area: return "No existe",404
    if request.method=="POST":
        area["nombre"]    = request.form["nombre"]
        area["encargado"] = request.form["encargado"]
        guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area=area)

@app.route("/areas/delete/<id>", methods=["POST"])
def areas_delete(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(AREAS_CSV) if r["id"]!=id]
    guardar_csv(AREAS_CSV, rows, ["id","nombre","encargado"])
    return redirect(url_for("areas_list"))

# --- CRUD Profesores (Admin, Encargado) ---
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
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = cargar_csv(PROFESORES_CSV)
        rows.append({
            "correo": request.form["correo"],
            "nombre": request.form["nombre"],
            "area":   request.form["area"]
        })
        guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor={})

@app.route("/profesores/edit/<correo>", methods=["GET","POST"])
def profesores_edit(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(PROFESORES_CSV)
    prof = next((r for r in rows if r["correo"]==correo), None)
    if not prof: return "No existe",404
    if request.method=="POST":
        prof["nombre"] = request.form["nombre"]
        prof["area"]   = request.form["area"]
        guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor=prof)

@app.route("/profesores/delete/<correo>", methods=["POST"])
def profesores_delete(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(PROFESORES_CSV) if r["correo"]!=correo]
    guardar_csv(PROFESORES_CSV, rows, ["correo","nombre","area"])
    return redirect(url_for("profesores_list"))

# --- CRUD Alumnos (Admin, Encargado, Maestro) ---
@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = cargar_csv(ALUMNOS_CSV)
    u = session["user"]
    if u["rol"]=="Maestro":
        rows = [r for r in rows if r["profesor"]==u["correo"]]
    if u["rol"]=="Encargado":
        rows = [r for r in rows if r["area"]==u["area"]]
    return render_template("alumnos_list.html", alumnos=rows)

@app.route("/alumnos/new", methods=["GET","POST"])
def alumnos_new():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            "No_Control": request.form["No_Control"],
            "nombre":     request.form["nombre"],
            "area":       request.form["area"],
            "profesor":   request.form["profesor"]
        })
        guardar_csv(ALUMNOS_CSV, rows, ["No_Control","nombre","area","profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno={})

@app.route("/alumnos/edit/<no_control>", methods=["GET","POST"])
def alumnos_edit(no_control):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(ALUMNOS_CSV)
    alum = next((r for r in rows if r["No_Control"]==no_control), None)
    if not alum: return "No existe",404
    if request.method=="POST":
        alum["nombre"]   = request.form["nombre"]
        alum["area"]     = request.form["area"]
        alum["profesor"] = request.form["profesor"]
        guardar_csv(ALUMNOS_CSV, rows, ["No_Control","nombre","area","profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno=alum)

@app.route("/alumnos/delete/<no_control>", methods=["POST"])
def alumnos_delete(no_control):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(ALUMNOS_CSV) if r["No_Control"]!=no_control]
    guardar_csv(ALUMNOS_CSV, rows, ["No_Control","nombre","area","profesor"])
    return redirect(url_for("alumnos_list"))

# --- Generación de documentos ---
TEMPLATE_SOL  = "plantillas/plantilla_solicitud_completa.docx"
TEMPLATE_BIM  = "plantillas/Reporte_Bimestral_Plantilla.docx"
TEMPLATE_FIN  = "plantillas/Reporte_Final_Lleno.docx"

def _make_doc(template_path, context, filename):
    doc = DocxTemplate(template_path)
    doc.render(context)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

def obtener_datos_documento(no_control):
    rows = cargar_csv(DOCUMENTOS_CSV)
    return next((r for r in rows if r["No_Control"]==no_control), None)

@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    row = obtener_datos_documento(no_control)
    if not row:
        flash("No se encontró al alumno en documentos.csv", "danger")
        return redirect(url_for("dashboard"))
    return _make_doc(TEMPLATE_SOL, row, f"Solicitud_{no_control}.docx")

@app.route("/generar_bimestral/<bim>/<no_control>")
def generar_bimestral(bim, no_control):
    row = obtener_datos_documento(no_control)
    if not row:
        flash("No se encontró al alumno en documentos.csv", "danger")
        return redirect(url_for("dashboard"))
    # puedes inyectar el bimestre en el contexto si la plantilla lo usa
    row["Bimestre"] = bim
    return _make_doc(TEMPLATE_BIM, row, f"Reporte_Bimestral_{bim}_{no_control}.docx")

@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    row = obtener_datos_documento(no_control)
    if not row:
        flash("No se encontró al alumno en documentos.csv", "danger")
        return redirect(url_for("dashboard"))
    return _make_doc(TEMPLATE_FIN, row, f"Reporte_Final_{no_control}.docx")

# --- Error 404 genérico ---
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
