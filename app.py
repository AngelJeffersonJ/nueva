import os
import uuid
import csv
import zipfile

from io import BytesIO
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file
)
import msal
import pandas as pd
from docxtpl import DocxTemplate

# ─── 1. Configuración de la aplicación ─────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# ─── 2. Azure AD (MSAL) ─────────────────────────────────────────────────────
CLIENT_ID     = os.environ.get("CLIENT_ID",     "tu_client_id")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "tu_client_secret")
TENANT_ID     = os.environ.get("TENANT_ID",     "tu_tenant_id")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# ─── 3. Rutas de CSV y plantillas ───────────────────────────────────────────
BASE_PATH      = os.path.dirname(__file__)
CSV_DIR        = os.environ.get("CSV_DIR", "csv")
USUARIOS_CSV   = os.path.join(BASE_PATH, CSV_DIR, "usuarios.csv")
AREAS_CSV      = os.path.join(BASE_PATH, CSV_DIR, "areas.csv")
PROFESORES_CSV = os.path.join(BASE_PATH, CSV_DIR, "profesores.csv")
ALUMNOS_CSV    = os.path.join(BASE_PATH, CSV_DIR, "alumnos.csv")

PLANTILLA_SOLICITUD  = os.path.join(BASE_PATH, "plantillas", "plantilla_solicitud_completa.docx")
PLANTILLA_BIMESTRAL1 = os.path.join(BASE_PATH, "plantillas", "Reporte_Bimestral_Plantilla.docx")
PLANTILLA_FINAL      = os.path.join(BASE_PATH, "plantillas", "Reporte_Final_Plantilla.docx")

DOCUMENTOS_DIR = os.path.join(BASE_PATH, "documentos_generados")
os.makedirs(DOCUMENTOS_DIR, exist_ok=True)

# ─── 4. Campos de CSV de alumnos (único CSV compartido) ────────────────────
CAMPOS_ALUMNOS = [
    # Datos básicos
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo","Domicilio","No_Control",
    "Carrera","Periodo","Semestre","Creditos","Dependencia","Domicilio_Dependencia","Titular_Dependencia",
    "Cargo_Responsable","Responsable_Proyecto","Nombre_Programa",
    "Modalidad_Externa","Modalidad_Interna","Fecha_Inicio","Fecha_Terminacion","Actividades",
    # Casillas tipo participación
    "TP_Edu_Adultos","TP_Deportivo","TP_Civico","TP_Salud","TP_Otros",
    "TP_Desarrollo","TP_Cultural","TP_Sustentable","TP_Medio_Amb",
    # Fechas de solicitud
    "Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
    # Campos para reporte bimestral
    "día1","mes1","año1","dia2","mes2","año2",
    "Nombre_supervisor","Puesto_supervisor",
    "Actividad_1","Actividad_2","Actividad_3","Actividad_4",
    "Actividad_5","Actividad_6","Actividad_7","Actividad_8",
    # Marcas de bimestre
    "x1","x2","x3",
    # Campos para reporte final
    "Actividad1","Logro1","Actividad2","Logro2","Actividad3","Logro3","Actividad4","Logro4",
    "Actividad5","Logro5","Actividad6","Logro6","Actividad7","Logro7","Actividad8","Logro8",
    "Aprendizaje1","Beneficio1","Aprendizaje2","Beneficio2","Aprendizaje3","Beneficio3",
    "Aprendizaje4","Beneficio4","Aprendizaje5","Beneficio5","Aprendizaje6","Beneficio6",
    "Aprendizaje7","Beneficio7","Aprendizaje8","Beneficio8"
]

# ─── 5. Funciones auxiliares CSV ────────────────────────────────────────────
def cargar_csv(ruta):
    if os.path.exists(ruta):
        with open(ruta, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(ruta, filas, campos):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(filas)

# ─── 6. Sesión y roles ──────────────────────────────────────────────────────
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# ─── 7. Login / Logout ─────────────────────────────────────────────────────
@app.route("/login")
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    if not code:
        flash("Falta código de autorización", "danger")
        return redirect(url_for("login"))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )
    if "error" in result:
        return f"Error MSAL: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {
        "correo": email, "rol": "Alumno", "area": ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol":    perfil["rol"],
        "area":   perfil.get("area",""),
        "name":   claims.get("name","")
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ─── 8. Dashboard ─────────────────────────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    rol    = session["user"]["rol"]
    correo = session["user"]["correo"]
    area   = session["user"]["area"]
    alumnos = cargar_csv(ALUMNOS_CSV)

    if rol == "Maestro":
        alumnos = [a for a in alumnos if a.get("profesor")==correo]
    elif rol == "Encargado":
        alumnos = [a for a in alumnos if a.get("area")==area]
    elif rol == "Alumno":
        alumnos = [a for a in alumnos if a.get("Correo","").lower()==correo.lower()]

    return render_template(
        "dashboard.html",
        usuario=session["user"],
        alumnos=alumnos
    )

# ─── 9. CRUD Usuarios ─────────────────────────────────────────────────────
@app.route("/usuarios")
def usuarios_list():
    if not validar_acceso(["Administrador"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    usuarios = cargar_csv(USUARIOS_CSV)
    return render_template("usuarios_list.html", usuarios=usuarios)

@app.route("/usuarios/new", methods=["GET","POST"])
def usuarios_new():
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        filas = cargar_csv(USUARIOS_CSV)
        filas.append({
            "correo": request.form["correo"],
            "rol":    request.form["rol"],
            "area":   request.form["area"]
        })
        guardar_csv(USUARIOS_CSV, filas, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario={})

@app.route("/usuarios/edit/<correo>", methods=["GET","POST"])
def usuarios_edit(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    filas = cargar_csv(USUARIOS_CSV)
    u = next((r for r in filas if r["correo"]==correo), None)
    if not u: return "No existe",404
    if request.method=="POST":
        u["rol"], u["area"] = request.form["rol"], request.form["area"]
        guardar_csv(USUARIOS_CSV, filas, ["correo","rol","area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario=u)

@app.route("/usuarios/delete/<correo>", methods=["POST"])
def usuarios_delete(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    filas = [r for r in cargar_csv(USUARIOS_CSV) if r["correo"]!=correo]
    guardar_csv(USUARIOS_CSV, filas, ["correo","rol","area"])
    return redirect(url_for("usuarios_list"))

# ─── 10. CRUD Áreas ───────────────────────────────────────────────────────
@app.route("/areas")
def areas_list():
    if not validar_acceso(["Administrador","Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    areas = cargar_csv(AREAS_CSV)
    return render_template("areas_list.html", areas=areas)

@app.route("/areas/new", methods=["GET","POST"])
def areas_new():
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        filas = cargar_csv(AREAS_CSV)
        filas.append({
            "id":        str(uuid.uuid4()),
            "nombre":    request.form["nombre"],
            "encargado": request.form["encargado"]
        })
        guardar_csv(AREAS_CSV, filas, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area={})

@app.route("/areas/edit/<id>", methods=["GET","POST"])
def areas_edit(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    filas = cargar_csv(AREAS_CSV)
    a = next((r for r in filas if r["id"]==id), None)
    if not a: return "No existe",404
    if request.method=="POST":
        a["nombre"], a["encargado"] = request.form["nombre"], request.form["encargado"]
        guardar_csv(AREAS_CSV, filas, ["id","nombre","encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area=a)

@app.route("/areas/delete/<id>", methods=["POST"])
def areas_delete(id):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    filas = [r for r in cargar_csv(AREAS_CSV) if r["id"]!=id]
    guardar_csv(AREAS_CSV, filas, ["id","nombre","encargado"])
    return redirect(url_for("areas_list"))

# ─── 11. CRUD Profesores ───────────────────────────────────────────────────
@app.route("/profesores")
def profesores_list():
    if not validar_acceso(["Administrador","Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    profs = cargar_csv(PROFESORES_CSV)
    return render_template("profesores_list.html", profesores=profs)

@app.route("/profesores/new", methods=["GET","POST"])
def profesores_new():
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        filas = cargar_csv(PROFESORES_CSV)
        filas.append({
            "correo": request.form["correo"],
            "nombre": request.form["nombre"],
            "area":   request.form["area"]
        })
        guardar_csv(PROFESORES_CSV, filas, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor={})

@app.route("/profesores/edit/<correo>", methods=["GET","POST"])
def profesores_edit(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    filas = cargar_csv(PROFESORES_CSV)
    p = next((r for r in filas if r["correo"]==correo), None)
    if not p: return "No existe",404
    if request.method=="POST":
        p["nombre"], p["area"] = request.form["nombre"], request.form["area"]
        guardar_csv(PROFESORES_CSV, filas, ["correo","nombre","area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor=p)

@app.route("/profesores/delete/<correo>", methods=["POST"])
def profesores_delete(correo):
    if not validar_acceso(["Administrador","Encargado"]):
        return redirect(url_for("dashboard"))
    filas = [r for r in cargar_csv(PROFESORES_CSV) if r["correo"]!=correo]
    guardar_csv(PROFESORES_CSV, filas, ["correo","nombre","area"])
    return redirect(url_for("profesores_list"))

# ─── 12. CRUD Alumnos ──────────────────────────────────────────────────────
@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    filas = cargar_csv(ALUMNOS_CSV)
    rol, correo, area = session["user"]["rol"], session["user"]["correo"], session["user"]["area"]
    if rol=="Maestro":
        filas = [a for a in filas if a.get("profesor")==correo]
    elif rol=="Encargado":
        filas = [a for a in filas if a.get("area")==area]
    return render_template("alumnos_list.html", alumnos=filas)

@app.route("/alumnos/new", methods=["GET","POST"])
def alumnos_new():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        filas = cargar_csv(ALUMNOS_CSV)
        filas.append({
            "Correo":   request.form["correo"],
            "Nombre":   request.form["nombre"],
            "area":     request.form["area"],
            "profesor": request.form["profesor"]
        })
        guardar_csv(ALUMNOS_CSV, filas, ["Correo","Nombre","area","profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno={})

@app.route("/alumnos/edit/<correo>", methods=["GET","POST"])
def alumnos_edit(correo):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    filas = cargar_csv(ALUMNOS_CSV)
    a = next((r for r in filas if r.get("Correo")==correo), None)
    if not a: return "No existe",404
    if request.method=="POST":
        a["Nombre"], a["area"], a["profesor"] = (
            request.form["nombre"],
            request.form["area"],
            request.form["profesor"]
        )
        guardar_csv(ALUMNOS_CSV, filas, ["Correo","Nombre","area","profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno=a)

@app.route("/alumnos/delete/<correo>", methods=["POST"])
def alumnos_delete(correo):
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    filas = [r for r in cargar_csv(ALUMNOS_CSV) if r.get("Correo")!=correo]
    guardar_csv(ALUMNOS_CSV, filas, ["Correo","Nombre","area","profesor"])
    return redirect(url_for("alumnos_list"))

# ─── 13. Generación de documentos Word ────────────────────────────────────
def _render_docs(plantilla, sufijo, df, extra=None):
    for _, fila in df.iterrows():
        tpl = DocxTemplate(plantilla)
        ctx = fila.to_dict()
        if extra:
            ctx.update(extra)
        tpl.render(ctx)
        nombre = f"{ctx.get('No_Control','sin_control')}_{sufijo}.docx"
        tpl.save(os.path.join(DOCUMENTOS_DIR, nombre))

@app.route("/generar/solicitud")
def generar_solicitud():
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"]==no]
    _render_docs(PLANTILLA_SOLICITUD, "Solicitud", df)
    flash("Solicitud(s) generada(s)", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/bimestral/<int:bimestre>")
def generar_bimestral(bimestre):
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"]==no]
    extra = {"Reporte_No": bimestre}
    # marcas X
    extra.update({"x1":"X" if bimestre==1 else "",
                  "x2":"X" if bimestre==2 else "",
                  "x3":"X" if bimestre==3 else ""})
    _render_docs(PLANTILLA_BIMESTRAL1, f"Bim{bimestre}", df, extra)
    flash(f"Bimestre {bimestre} generado", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/final")
def generar_final():
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"]==no]
    _render_docs(PLANTILLA_FINAL, "Final", df)
    flash("Reporte final generado", "success")
    return redirect(url_for("dashboard"))

# ─── 14. Error 404 ────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

# ─── 15. Inicio ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=False
    )
