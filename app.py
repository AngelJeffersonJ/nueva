# app.py
from flask import (
    Flask, render_template, request, redirect, url_for, session, flash, send_file
)
import os, csv, uuid, msal, zipfile
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate
from dotenv import load_dotenv

# ────────────────────────────────────────────────
# 1. Load .env + Flask setup
# ────────────────────────────────────────────────
load_dotenv()  # loads CLIENT_ID, CLIENT_SECRET, etc from .env

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "💥-cambia-este-secreto-💥")

# ────────────────────────────────────────────────
# 2. Azure AD / MSAL config
# ────────────────────────────────────────────────
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv("REDIRECT_URI", "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# sanity check
for name, val in [("CLIENT_ID", CLIENT_ID), ("CLIENT_SECRET", CLIENT_SECRET), ("TENANT_ID", TENANT_ID)]:
    if not val:
        raise RuntimeError(f"Env var {name} missing")

# ────────────────────────────────────────────────
# 3. Paths for CSVs & Templates & Output
# ────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(__file__)
CSV_DIR     = os.getenv("CSV_DIR", "csv")
TPL_DIR     = os.path.join(BASE_DIR, "plantillas")
OUTPUT_DIR  = os.path.join(BASE_DIR, "documentos_generados")

os.makedirs(os.path.join(BASE_DIR, CSV_DIR), exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

USUARIOS_CSV   = os.path.join(BASE_DIR, CSV_DIR, "usuarios.csv")
PROFESORES_CSV = os.path.join(BASE_DIR, CSV_DIR, "profesores.csv")
AREAS_CSV      = os.path.join(BASE_DIR, CSV_DIR, "areas.csv")
ALUMNOS_CSV    = os.path.join(BASE_DIR, CSV_DIR, "alumnos.csv")

SOLICITUD_TEMPLATE    = os.path.join(TPL_DIR, "plantilla_solicitud_completa.docx")
BIMESTRAL_TEMPLATE    = os.path.join(TPL_DIR, "Reporte_Bimestral_Plantilla.docx")
FINAL_TEMPLATE        = os.path.join(TPL_DIR, "Reporte_Final_Plantilla.docx")

# This master field list must cover *all* placeholders across the 3 templates:
MASTER_FIELDS = [
    # solicitud fields
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo","Domicilio","No_Control",
    "Carrera","Periodo","Semestre","Creditos","Dependencia","Domicilio_Dependencia","Titular_Dependencia",
    "Cargo_Responsable","Responsable_Proyecto","Nombre_Programa","Modalidad_Externa","Modalidad_Interna",
    "Fecha_Inicio","Fecha_Terminacion","Actividades","Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
    # bimestral fields
    "AP","AM","carrera","día1","mes1","año1","dia2","mes2","año2",
    "Actividad_1","Actividad_2","Actividad_3","Actividad_4","Actividad_5","Actividad_6",
    "Actividad_7","Actividad_8","Reporte_No","Nombre_supervisor","Puesto_supervisor",
    # final fields
    "Municipio","Estado","Fecha","Periodo",
    "Actividad1","Logro1","Actividad2","Logro2","Actividad3","Logro3","Actividad4","Logro4",
    "Actividad5","Logro5","Actividad6","Logro6","Actividad7","Logro7","Actividad8","Logro8",
    "Aprendizaje1","Beneficio1","Aprendizaje2","Beneficio2","Aprendizaje3","Beneficio3",
    "Aprendizaje4","Beneficio4","Aprendizaje5","Beneficio5","Aprendizaje6","Beneficio6",
    "Aprendizaje7","Beneficio7","Aprendizaje8","Beneficio8"
]

# ────────────────────────────────────────────────
# 4. CSV Helpers & Auth helpers
# ────────────────────────────────────────────────
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

# ────────────────────────────────────────────────
# 5. MSAL: login / callback / logout
# ────────────────────────────────────────────────
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
    if not code:
        flash("Missing auth code", "danger")
        return redirect(url_for("login"))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )

    if "error" in result:
        return f"MSAL Error: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Only @aguascalientes.tecnm.mx allowed", 403

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

# ────────────────────────────────────────────────
# 6. Dashboard
# ────────────────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ────────────────────────────────────────────────
# 7. Generic CRUD for CSV Entities
# ────────────────────────────────────────────────
def obtener_mapa(tipo):
    return {
        "usuarios":   (USUARIOS_CSV,   ["correo","rol","area"]),
        "profesores": (PROFESORES_CSV, ["correo","nombre","area"]),
        "areas":      (AREAS_CSV,      ["id","nombre","encargado"]),
        "alumnos":    (ALUMNOS_CSV,    ["correo","nombre","area","profesor"])
    }.get(tipo)

@app.route("/entidad/<tipo>")
def entidad_list(tipo):
    m = obtener_mapa(tipo)
    if not m: return render_template("404.html"), 404
    archivo, campos = m
    regs = cargar_csv(archivo)
    return render_template("entidad_list.html", tipo=tipo, campos=campos, registros=regs)

@app.route("/entidad/<tipo>/new", methods=["GET","POST"])
def entidad_new(tipo):
    m = obtener_mapa(tipo)
    if not m: return render_template("404.html"), 404
    archivo, campos = m

    if request.method=="POST":
        row = {c: request.form.get(c,"").strip() for c in campos}
        if any(v=="" for v in row.values()):
            flash("All fields required","danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        if campos[0]=="id":
            row["id"] = str(uuid.uuid4())
        regs = cargar_csv(archivo)
        if any(r[campos[0]]==row[campos[0]] for r in regs):
            flash("Primary key exists","danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        regs.append(row)
        guardar_csv(archivo, regs, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    return render_template("entidad_form.html", tipo=tipo, campos=campos, valores={})

@app.route("/entidad/<tipo>/edit/<pk>", methods=["GET","POST"])
def entidad_edit(tipo, pk):
    m = obtener_mapa(tipo)
    if not m: return render_template("404.html"), 404
    archivo, campos = m
    regs = cargar_csv(archivo)
    item = next((r for r in regs if r[campos[0]]==pk), None)
    if not item:
        flash("Not found","danger")
        return redirect(url_for("entidad_list", tipo=tipo))

    if request.method=="POST":
        for c in campos:
            v = request.form.get(c,"").strip()
            if v=="": 
                flash(f"{c} cannot be empty","danger")
                return redirect(url_for("entidad_edit", tipo=tipo, pk=pk))
            item[c] = v
        guardar_csv(archivo, regs, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    return render_template("entidad_form.html", tipo=tipo, campos=campos, valores=item)

@app.route("/entidad/<tipo>/delete/<pk>", methods=["POST"])
def entidad_delete(tipo, pk):
    m = obtener_mapa(tipo)
    if not m: return render_template("404.html"), 404
    archivo, campos = m
    regs = [r for r in cargar_csv(archivo) if r[campos[0]]!=pk]
    guardar_csv(archivo, regs, campos)
    return redirect(url_for("entidad_list", tipo=tipo))

# ────────────────────────────────────────────────
# 8. Document generation
# ────────────────────────────────────────────────
def _render_docs(template_path, out_prefix, df, extra_ctx=None):
    """Internal: render one doc per row of df using DocxTemplate"""
    extra_ctx = extra_ctx or {}
    for idx, row in df.iterrows():
        ctx = {k: str(row.get(k,"")) for k in MASTER_FIELDS}
        ctx.update(extra_ctx)
        doc = DocxTemplate(template_path)
        doc.render(ctx)
        fn = f"{out_prefix}_{row.get('No_Control',idx)}.docx"
        doc.save(os.path.join(OUTPUT_DIR, fn))

@app.route("/generar/solicitud")
def generar_solicitud():
    df = pd.read_csv(ALUMNOS_CSV)
    _render_docs(SOLICITUD_TEMPLATE, "Solicitud", df)
    flash("Solicitudes generadas","success")
    return redirect(url_for("dashboard"))

@app.route("/generar/bimestral/<int:periodo>")
def generar_bimestral(periodo):
    df = pd.read_csv(ALUMNOS_CSV)
    _render_docs(BIMESTRAL_TEMPLATE, f"Bimestral{periodo}", df, {"Reporte_No": periodo})
    flash(f"Bimestral {periodo} generados","success")
    return redirect(url_for("dashboard"))

@app.route("/generar/final")
def generar_final():
    df = pd.read_csv(ALUMNOS_CSV)
    _render_docs(FINAL_TEMPLATE, "Final", df)
    flash("Reportes finales generados","success")
    return redirect(url_for("dashboard"))

@app.route("/descargar")
def descargar():
    mem = BytesIO()
    with zipfile.ZipFile(mem, "w") as zf:
        for fn in os.listdir(OUTPUT_DIR):
            zf.write(os.path.join(OUTPUT_DIR, fn), fn)
    mem.seek(0)
    return send_file(mem, download_name="docs.zip", as_attachment=True)

# ────────────────────────────────────────────────
# 9. 404 handler & run
# ────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__=="__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT",5000)), debug=False)
