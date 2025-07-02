import os
import csv
import uuid
import zipfile
import msal
import pandas as pd

from io import BytesIO
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file
)
from docxtpl import DocxTemplate

# ─── 1. Configuración de Flask y entorno ──────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# ─── 2. Azure AD (MSAL) ───────────────────────────────────────────────────────
CLIENT_ID     = os.environ.get("CLIENT_ID",     "tu_client_id")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "tu_client_secret")
TENANT_ID     = os.environ.get("TENANT_ID",     "tu_tenant_id")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# ─── 3. Rutas de ficheros ─────────────────────────────────────────────────────
BASE_PATH     = os.path.dirname(__file__)
CSV_DIR       = os.environ.get("CSV_DIR", "csv")
USUARIOS_CSV   = os.path.join(BASE_PATH, CSV_DIR, "usuarios.csv")
PROFESORES_CSV = os.path.join(BASE_PATH, CSV_DIR, "profesores.csv")
AREAS_CSV      = os.path.join(BASE_PATH, CSV_DIR, "areas.csv")
ALUMNOS_CSV    = os.path.join(BASE_PATH, CSV_DIR, "alumnos.csv")

PLANTILLA_SOLICITUD = os.path.join(BASE_PATH, "plantillas", "plantilla_solicitud_completa.docx")
PLANTILLA_BIMESTRAL = os.path.join(BASE_PATH, "plantillas", "Reporte_Bimestral_Plantilla.docx")
PLANTILLA_FINAL     = os.path.join(BASE_PATH, "plantillas", "Reporte_Final_Plantilla.docx")

OUTPUT_PATH   = os.path.join(BASE_PATH, "documentos_generados")
os.makedirs(OUTPUT_PATH, exist_ok=True)

# ─── 4. Helpers CSV ───────────────────────────────────────────────────────────
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# ─── 5. Flujo de login MSAL ───────────────────────────────────────────────────
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
        flash("Código de autorización faltante", "danger")
        return redirect(url_for("login"))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )

    if "error" in result:
        return f"Error al autenticarse: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username", "").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas @aguascalientes.tecnm.mx permitidas", 403

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

# ─── 6. Dashboard con lista filtrada de alumnos ──────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    # 1) Cargando todos los alumnos
    alumnos = cargar_csv(ALUMNOS_CSV)

    # 2) Filtrado según rol
    rol    = session["user"]["rol"]
    correo = session["user"]["correo"]
    area   = session["user"]["area"]

    if rol == "Maestro":
        alumnos = [a for a in alumnos if a.get("profesor")==correo]
    elif rol == "Encargado":
        alumnos = [a for a in alumnos if a.get("area")==area]
    elif rol == "Alumno":
        alumnos = [a for a in alumnos if a.get("correo")==correo]

    # 3) Render con la lista filtrada
    return render_template(
        "dashboard.html",
        usuario=session["user"],
        alumnos=alumnos
    )

# ─── 7. Función genérica para renderizar y salvar .docx ──────────────────────
def _render_docs(template_path, sufijo, dataframe, extra_ctx=None):
    for _, alumno in dataframe.iterrows():
        tpl = DocxTemplate(template_path)
        ctx = alumno.to_dict()
        if extra_ctx:
            ctx.update(extra_ctx)
        tpl.render(ctx)
        nombre = f"{ctx.get('No_Control','sin_control')}_{sufijo}.docx"
        tpl.save(os.path.join(OUTPUT_PATH, nombre))

# ─── 8. Rutas de generación de documentos ────────────────────────────────────
@app.route("/generar/solicitud")
def generar_solicitud():
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"] == no]
    _render_docs(PLANTILLA_SOLICITUD, "Solicitud", df)
    flash("Solicitudes generadas.", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/bimestral/<int:periodo>")
def generar_bimestral(periodo):
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"] == no]
    extra = {
        "Reporte_No": periodo,
        "x1": "X" if periodo==1 else "",
        "x2": "X" if periodo==2 else "",
        "x3": "X" if periodo==3 else ""
    }
    _render_docs(PLANTILLA_BIMESTRAL, f"Bimestral{periodo}", df, extra)
    flash(f"Reportes bimestrales {periodo} generados.", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/final")
def generar_final():
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    no = request.args.get("no_control")
    if no:
        df = df[df["No_Control"] == no]
    _render_docs(PLANTILLA_FINAL, "Final", df)
    flash("Reportes finales generados.", "success")
    return redirect(url_for("dashboard"))

# ─── 9. Error 404 ────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

# ─── 10. Arranque ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
