import os
import csv
import uuid
import msal
import zipfile
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from io import BytesIO
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# --- Azure AD Config (Render env vars) ---
CLIENT_ID     = os.environ.get("CLIENT_ID",     "c306c8d3-68dc-4110-b5fb-771b942c10db")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "CEX8Q~txlNOmiN4PAZKbV7matISmRV1HREkvbcMS")
TENANT_ID     = os.environ.get("TENANT_ID",     "63de1475-1a48-4463-aff2-b2581f2a972e")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "https://sistema-documental.onrender.com/getAToken")
SCOPE         = ["User.Read"]

# --- Rutas y directorios ---
BASE_PATH       = os.path.dirname(__file__)
CSV_DIR         = os.environ.get("CSV_DIR", "csv")
PLANTILLAS_DIR  = os.path.join(BASE_PATH, "plantillas")
OUTPUT_DIR      = os.path.join(BASE_PATH, "documentos_generados")

os.makedirs(CSV_DIR,        exist_ok=True)
os.makedirs(PLANTILLAS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR,     exist_ok=True)

ALUMNOS_CSV        = os.path.join(CSV_DIR, "alumnos.csv")
SOLICITUD_PLANTILLA= os.path.join(PLANTILLAS_DIR, "plantilla_solicitud_completa.docx")
BIMESTRAL_PLANTILLA= os.path.join(PLANTILLAS_DIR, "Reporte_Bimestral_Plantilla.docx")
FINAL_PLANTILLA    = os.path.join(PLANTILLAS_DIR, "Reporte_Final_Plantilla.docx")

# --- Utilerías para CSV ---
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

# --- Helpers de autenticación ---
def usuario_desde_csv(email):
    for u in cargar_csv(ALUMNOS_CSV):  # opcional: otro CSV de usuarios
        if u.get("Correo","").strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# --- Flujo MSAL ---
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
        flash(result.get("error_description","Auth error"), "danger")
        return redirect(url_for("dashboard"))

    claims = result["id_token_claims"]
    email  = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {"Correo":email,"rol":"Alumno","area":""}
    session["user"] = {
        "correo": perfil["Correo"],
        "rol":    perfil.get("rol","Alumno"),
        "area":   perfil.get("area",""),
        "name":   claims.get("name","")
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout?post_logout_redirect_uri={url_for('login',_external=True)}"
    )

# --- Dashboard ---
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# --- CRUD de Alumnos (ejemplo) ---
@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        flash("Acceso denegado","danger")
        return redirect(url_for("dashboard"))
    filas = cargar_csv(ALUMNOS_CSV)
    return render_template("alumnos_list.html", alumnos=filas)

@app.route("/alumnos/new", methods=["GET","POST"])
def alumnos_new():
    if not validar_acceso(["Administrador","Encargado","Maestro"]):
        return redirect(url_for("dashboard"))
    if request.method=="POST":
        filas = cargar_csv(ALUMNOS_CSV)
        nueva = {
            "Correo":   request.form["correo"].strip(),
            "Nombre":   request.form["nombre"].strip(),
            "area":     request.form["area"].strip(),
            "profesor": request.form["profesor"].strip(),
            # el CSV unificado tiene muchas más columnas, pero aquí sólo editamos éstas cuatro
        }
        filas.append(nueva)
        guardar_csv(ALUMNOS_CSV, filas, nueva.keys())
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html")

# --- Función auxiliar para generar documentos ---
def _make_doc(template_path, fila, nombre_salida):
    context = {k: fila.get(k,"") for k in fila.keys()}
    doc = DocxTemplate(template_path)
    doc.render(context)
    destino = os.path.join(OUTPUT_DIR, nombre_salida)
    doc.save(destino)
    return destino

# --- Generar Solicitud ---
@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV) if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No se encontró el No. de control","danger")
        return redirect(url_for("dashboard"))
    path = _make_doc(SOLICITUD_PLANTILLA, fila, f"Solicitud_{no_control}.docx")
    return send_file(path, as_attachment=True)

# --- Generar Reporte Bimestral (1,2,3) ---
@app.route("/generar_bimestral/<rep>/<no_control>")
def generar_bimestral(rep, no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV) if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No se encontró el No. de control","danger")
        return redirect(url_for("dashboard"))
    copy = fila.copy()
    copy["Reporte_No"] = rep
    path = _make_doc(BIMESTRAL_PLANTILLA, copy, f"Reporte_Bimestral_{rep}_{no_control}.docx")
    return send_file(path, as_attachment=True)

# --- Generar Reporte Final ---
@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV) if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No se encontró el No. de control","danger")
        return redirect(url_for("dashboard"))
    copy = fila.copy()
    # Ajustes de nombre de campo en template
    copy["Fecha"]   = copy.get("Fecha_Terminacion","")
    copy["Periodo"] = copy.get("Periodo","")
    path = _make_doc(FINAL_PLANTILLA, copy, f"Reporte_Final_{no_control}.docx")
    return send_file(path, as_attachment=True)

# --- Error 404 ---
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
