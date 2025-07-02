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

# ─── 1. Configuración de la aplicación ───────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# ─── 2. Azure AD (MSAL) ──────────────────────────────────────────────────────
CLIENT_ID     = os.environ.get("CLIENT_ID",     "c306c8d3-68dc-4110-b5fb-771b942c10db")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "CEX8Q~s3G7CwKdbOF9fcwsFBdLt1NFRiG0XvBdrL")
TENANT_ID     = os.environ.get("TENANT_ID",     "63de1475-1a48-4463-aff2-b2581f2a972e")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# ─── 3. Rutas de CSV y plantillas ─────────────────────────────────────────────
BASE_PATH           = os.path.dirname(__file__)
CSV_DIR             = os.environ.get("CSV_DIR", "csv")
USUARIOS_CSV        = os.path.join(BASE_PATH, CSV_DIR, "usuarios.csv")
AREAS_CSV           = os.path.join(BASE_PATH, CSV_DIR, "areas.csv")
PROFESORES_CSV      = os.path.join(BASE_PATH, CSV_DIR, "profesores.csv")
ALUMNOS_CSV         = os.path.join(BASE_PATH, CSV_DIR, "alumnos.csv")

PLANTILLA_SOLICITUD = os.path.join(BASE_PATH, "plantillas", "plantilla_solicitud_completa.docx")
PLANTILLA_BIMESTRAL = os.path.join(BASE_PATH, "plantillas", "Reporte_Bimestral_Plantilla.docx")
PLANTILLA_FINAL     = os.path.join(BASE_PATH, "plantillas", "Reporte_Final_Plantilla.docx")

DOCUMENTOS_DIR      = os.path.join(BASE_PATH, "documentos_generados")
os.makedirs(DOCUMENTOS_DIR, exist_ok=True)

# ─── 4. Funciones auxiliares ──────────────────────────────────────────────────
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

def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# ─── 5. Login / Logout ───────────────────────────────────────────────────────
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
        flash("Falta código de autorización", "danger")
        return redirect(url_for("login"))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
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

    perfil = usuario_desde_csv(email) or {"correo":email, "rol":"Alumno", "area":""}
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
        f"?post_logout_redirect_uri={url_for('login',_external=True)}"
    )

# ─── 6. Dashboard ───────────────────────────────────────────────────────────
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
        alumnos = [a for a in alumnos if a.get("profesor","").lower()==correo.lower()]
    elif rol == "Encargado":
        alumnos = [a for a in alumnos if a.get("area","")==area]
    elif rol == "Alumno":
        alumnos = [a for a in alumnos if a.get("correo","").lower()==correo.lower()]

    return render_template("dashboard.html",
        usuario=session["user"],
        alumnos=alumnos
    )

# ─── 7. CRUD Genérico Usuarios/Áreas/Profesores/Alumnos ────────────────────
# (Implementa aquí las rutas /usuarios, /areas, /profesores, /alumnos como antes)
# Ejemplo para alumnos:
@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        flash("Acceso denegado","danger")
        return redirect(url_for("dashboard"))
    filas = cargar_csv(ALUMNOS_CSV)
    rol, correo, area = session["user"]["rol"], session["user"]["correo"], session["user"]["area"]
    if rol=="Maestro":
        filas = [a for a in filas if a.get("profesor","").lower()==correo.lower()]
    if rol=="Encargado":
        filas = [a for a in filas if a.get("area","")==area]
    return render_template("alumnos_list.html", alumnos=filas)

# ─── 8. Generación de documentos y ZIP ──────────────────────────────────────
def _render_docs(plantilla, sufijo, df, extra=None):
    for _, fila in df.iterrows():
        tpl = DocxTemplate(plantilla)
        ctx = fila.to_dict()
        if extra:
            ctx.update(extra)
        tpl.render(ctx)
        nombre = f"{ctx.get('No_Control','sin_control')}_{sufijo}.docx"
        tpl.save(os.path.join(DOCUMENTOS_DIR, nombre))

@app.route("/generar/individual/<tipo>")
def descargar_doc_individual(tipo):
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    correo = session["user"]["correo"]
    df = df[df["correo"].str.lower()==correo.lower()]
    suf = {"solicitud":"Solicitud","b1":"Bim1","b2":"Bim2","b3":"Bim3","final":"Final"}[tipo]
    _render_docs({
        "solicitud":PLANTILLA_SOLICITUD,
        "b1":PLANTILLA_BIMESTRAL,
        "b2":PLANTILLA_BIMESTRAL,
        "b3":PLANTILLA_BIMESTRAL,
        "final":PLANTILLA_FINAL
    }[tipo], suf, df, extra={
        "Reporte_No": 1 if tipo=="b1" else (2 if tipo=="b2" else (3 if tipo=="b3" else "")),
        **({f"x{i}":"X" for i in (1,2,3) if tipo==f"b{i}"} or {})
    })
    return send_file(
        os.path.join(DOCUMENTOS_DIR, f"{df.iloc[0]['No_Control']}_{suf}.docx"),
        as_attachment=True
    )

@app.route("/generar/zip/<tipo>")
def generar_zip(tipo):
    df = pd.read_csv(ALUMNOS_CSV, dtype=str)
    user = session["user"]
    rol = user["rol"]; correo = user["correo"]; area = user["area"]
    # Filtrar según rol...
    if rol=="Maestro":
        df = df[df["profesor"].str.lower()==correo.lower()]
    if rol=="Encargado":
        df = df[df["area"]==area]
    if rol=="Alumno":
        df = df[df["correo"].str.lower()==correo.lower()]

    suf = {"solicitud":"Solicitud","b1":"Bim1","b2":"Bim2","b3":"Bim3","final":"Final"}[tipo]
    _render_docs({
        "solicitud":PLANTILLA_SOLICITUD,
        "b1":PLANTILLA_BIMESTRAL,
        "b2":PLANTILLA_BIMESTRAL,
        "b3":PLANTILLA_BIMESTRAL,
        "final":PLANTILLA_FINAL
    }[tipo], suf, df, extra={
        "Reporte_No": {"b1":1,"b2":2,"b3":3}.get(tipo,""),
        **({f"x{i}":"X" for i in (1,2,3) if tipo==f"b{i}"} or {})
    })

    # Crear ZIP con todos los .docx generados
    memoria = BytesIO()
    with zipfile.ZipFile(memoria,"w") as zf:
        for fn in os.listdir(DOCUMENTOS_DIR):
            if fn.endswith(f"_{suf}.docx"):
                zf.write(os.path.join(DOCUMENTOS_DIR,fn), arcname=fn)
    memoria.seek(0)
    return send_file(memoria,
        download_name=f"todos_{suf}.zip",
        as_attachment=True
    )

# ─── 9. Error 404 ───────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

# ─── 10. Inicio ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT",5000)),
        debug=True
    )
