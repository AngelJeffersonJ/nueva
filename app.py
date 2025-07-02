import os, csv, uuid, zipfile
from io import BytesIO
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file
)
import msal
import pandas as pd
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# ─── Azure AD / MSAL ─────────────────────────────────────────────
CLIENT_ID     = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID     = os.environ.get("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI")
SCOPE         = ["User.Read"]

# ─── Paths ────────────────────────────────────────────────────────
BASE_PATH       = os.path.dirname(__file__)
CSV_PATH        = os.path.join(BASE_PATH, os.environ.get("CSV_DIR","csv"))
PLANTILLA_SOL   = os.path.join(BASE_PATH, "plantillas", "plantilla_solicitud_completa.docx")
PLANTILLA_BIM   = os.path.join(BASE_PATH, "plantillas", "Reporte_Bimestral_Plantilla.docx")
PLANTILLA_FIN   = os.path.join(BASE_PATH, "plantillas", "Reporte_Final_Plantilla.docx")
OUTPUT_PATH     = os.path.join(BASE_PATH, "documentos_generados")

for p in (CSV_PATH, OUTPUT_PATH):
    os.makedirs(p, exist_ok=True)

USUARIOS_CSV   = os.path.join(CSV_PATH, "usuarios.csv")
PROFESORES_CSV = os.path.join(CSV_PATH, "profesores.csv")
AREAS_CSV      = os.path.join(CSV_PATH, "areas.csv")
ALUMNOS_CSV    = os.path.join(CSV_PATH, "alumnos.csv")

# ─── CSV Utilities ───────────────────────────────────────────────
def cargar_csv(path):
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path, rows, fields):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

# ─── Sesión & Roles ──────────────────────────────────────────────
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower()==email:
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return u and u.get("rol") in roles

# ─── Login / Logout ──────────────────────────────────────────────
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
        return f"Error MSAL: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas institucionales permitidas", 403

    perfil = usuario_desde_csv(email) or {
        "correo": email, "rol": "Alumno", "area": ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol":     perfil["rol"],
        "area":    perfil.get("area",""),
        "name":    claims.get("name","")
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ─── Dashboard ───────────────────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    # cargamos todos los alumnos para los botones
    alumnos = cargar_csv(ALUMNOS_CSV)
    return render_template("dashboard.html",
                            usuario=session["user"],
                            alumnos=alumnos)

# ─── Generación Dinámica de DOCX ────────────────────────────────
def _make_doc(plantilla_path, fila, nombre_salida):
    doc = DocxTemplate(plantilla_path)
    # docxtpl acepta valores nulos, preparamos contexto
    ctx = {k: (v if v is not None else "") for k,v in fila.items()}
    doc.render(ctx)
    mem = BytesIO()
    doc.save(mem)
    mem.seek(0)
    return send_file(mem,
                     download_name=nombre_salida,
                     as_attachment=True)

@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("dashboard"))
    return _make_doc(PLANTILLA_SOL,
                     fila,
                     f"Solicitud_{no_control}.docx")

@app.route("/generar_bimestral/<int:bimestre>/<no_control>")
def generar_bimestral(bimestre, no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila or bimestre not in (1,2,3):
        flash("Datos inválidos", "danger")
        return redirect(url_for("dashboard"))
    # agregamos al contexto el número de reporte
    fila["Reporte_No"] = bimestre
    return _make_doc(PLANTILLA_BIM,
                     fila,
                     f"Reporte_Bimestral_{bimestre}_{no_control}.docx")

@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("dashboard"))
    return _make_doc(PLANTILLA_FIN,
                     fila,
                     f"Reporte_Final_{no_control}.docx")

# ─── CRUD Genérico CSV (areas, profesores, usuarios, alumnos) ─────
def mapa_entidades(tipo):
    return {
        "usuarios"  : (USUARIOS_CSV,   ["correo","rol","area","profesor"]),
        "profesores": (PROFESORES_CSV, ["correo","nombre","area"]),
        "areas"     : (AREAS_CSV,      ["id","nombre","encargado"]),
        "alumnos"   : (ALUMNOS_CSV,    cargar_csv(ALUMNOS_CSV)[0].keys() if cargar_csv(ALUMNOS_CSV) else [])
    }.get(tipo)

@app.route("/<tipo>", methods=["GET"])
def lista_entidad(tipo):
    entidad = mapa_entidades(tipo)
    if not entidad:
        return render_template("404.html"), 404
    archivo, campos = entidad
    filas = cargar_csv(archivo)
    return render_template(f"{tipo}_list.html", **{tipo: filas})

@app.route("/<tipo>/new", methods=["GET","POST"])
def nueva_entidad(tipo):
    entidad = mapa_entidades(tipo)
    if not entidad:
        return render_template("404.html"), 404
    archivo, campos = entidad
    if request.method=="POST":
        row = {c: request.form.get(c,"").strip() for c in campos}
        if campos[0]=="id":
            row["id"] = str(uuid.uuid4())
        filas = cargar_csv(archivo)
        filas.append(row)
        guardar_csv(archivo, filas, campos)
        return redirect(url_for("lista_entidad", tipo=tipo))
    return render_template(f"{tipo}_form.html", **{tipo.rstrip("s"): {}})

@app.route("/<tipo>/edit/<pk>", methods=["GET","POST"])
def editar_entidad(tipo, pk):
    entidad = mapa_entidades(tipo)
    if not entidad:
        return render_template("404.html"), 404
    archivo, campos = entidad
    filas = cargar_csv(archivo)
    pk_field = campos[0]
    item = next((r for r in filas if r.get(pk_field)==pk), None)
    if not item:
        flash("No encontrado","danger")
        return redirect(url_for("lista_entidad", tipo=tipo))
    if request.method=="POST":
        for c in campos:
            item[c] = request.form.get(c,"").strip()
        guardar_csv(archivo, filas, campos)
        return redirect(url_for("lista_entidad", tipo=tipo))
    return render_template(f"{tipo}_form.html", **{tipo.rstrip("s"): item})

@app.route("/<tipo>/delete/<pk>", methods=["POST"])
def borrar_entidad(tipo, pk):
    entidad = mapa_entidades(tipo)
    if not entidad:
        return render_template("404.html"), 404
    archivo, campos = entidad
    filas = [r for r in cargar_csv(archivo) if r.get(campos[0])!=pk]
    guardar_csv(archivo, filas, campos)
    return redirect(url_for("lista_entidad", tipo=tipo))

# ─── Error 404 ──────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
