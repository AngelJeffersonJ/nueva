# app.py
from flask import (
    Flask, render_template, request, redirect, url_for, session, flash, send_file
)
import os, csv, uuid, msal, zipfile
import pandas as pd
from io import BytesIO
from docxtpl import DocxTemplate
from dotenv import load_dotenv

# ─── 1. Carga de .env y configuración de Flask ──────────────────────────────
load_dotenv()  # lee el archivo .env en local
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "💥-cambia-este-secreto-💥")

# ─── 2. Configuración de Azure AD / MSAL ────────────────────────────────────
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv("REDIRECT_URI", "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# Verificar que existan las variables críticas
for nombre, valor in [("CLIENT_ID", CLIENT_ID), ("CLIENT_SECRET", CLIENT_SECRET), ("TENANT_ID", TENANT_ID)]:
    if not valor:
        raise RuntimeError(f"Falta la variable de entorno {nombre}")

# ─── 3. Rutas de archivos ───────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(__file__)
CSV_DIR     = os.getenv("CSV_DIR", "csv")
TPL_DIR     = os.path.join(BASE_DIR, "plantillas")
OUTPUT_DIR  = os.path.join(BASE_DIR, "documentos_generados")

os.makedirs(os.path.join(BASE_DIR, CSV_DIR), exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

USUARIOS_CSV     = os.path.join(BASE_DIR, CSV_DIR, "usuarios.csv")
PROFESORES_CSV   = os.path.join(BASE_DIR, CSV_DIR, "profesores.csv")
AREAS_CSV        = os.path.join(BASE_DIR, CSV_DIR, "areas.csv")
ALUMNOS_CSV      = os.path.join(BASE_DIR, CSV_DIR, "alumnos.csv")

SOLICITUD_TPL    = os.path.join(TPL_DIR, "plantilla_solicitud_completa.docx")
BIMESTRAL_TPL    = os.path.join(TPL_DIR, "Reporte_Bimestral_Plantilla.docx")
FINAL_TPL        = os.path.join(TPL_DIR, "Reporte_Final_Plantilla.docx")

# Campos maestros que cubren todas las plantillas
MASTER_FIELDS = [
    # solicitud
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo","Domicilio",
    "No_Control","Carrera","Periodo","Semestre","Creditos","Dependencia","Domicilio_Dependencia",
    "Titular_Dependencia","Cargo_Responsable","Responsable_Proyecto","Nombre_Programa",
    "Modalidad_Externa","Modalidad_Interna","Fecha_Inicio","Fecha_Terminacion","Actividades",
    "Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
    # bimestral
    "AP","AM","carrera","día1","mes1","año1","dia2","mes2","año2",
    "Actividad_1","Actividad_2","Actividad_3","Actividad_4","Actividad_5","Actividad_6",
    "Actividad_7","Actividad_8","Reporte_No","Nombre_supervisor","Puesto_supervisor",
    "x1","x2","x3",
    # final
    "Municipio","Estado","Fecha","Periodo",
    "Actividad1","Logro1","Actividad2","Logro2","Actividad3","Logro3","Actividad4","Logro4",
    "Actividad5","Logro5","Actividad6","Logro6","Actividad7","Logro7","Actividad8","Logro8",
    "Aprendizaje1","Beneficio1","Aprendizaje2","Beneficio2","Aprendizaje3","Beneficio3",
    "Aprendizaje4","Beneficio4","Aprendizaje5","Beneficio5","Aprendizaje6","Beneficio6",
    "Aprendizaje7","Beneficio7","Aprendizaje8","Beneficio8"
]

# ─── 4. Funciones de ayuda para CSV y autenticación ─────────────────────────
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

# ─── 5. Flujo MSAL: login / callback / logout ────────────────────────────────
@app.route("/login")
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    if not code:
        flash("Falta el código de autorización", "danger")
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
        return "Solo cuentas @aguascalientes.tecnm.mx permitidas", 403

    perfil = usuario_desde_csv(email) or {"correo":email,"rol":"Alumno","area":""}
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

# ─── 6. Dashboard ────────────────────────────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ─── 7. CRUD genérico para CSV ───────────────────────────────────────────────
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
        fila = {c: request.form.get(c,"").strip() for c in campos}
        if any(v=="" for v in fila.values()):
            flash("Todos los campos son obligatorios.", "danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        if campos[0]=="id":
            fila["id"] = str(uuid.uuid4())
        regs = cargar_csv(archivo)
        if any(r[campos[0]]==fila[campos[0]] for r in regs):
            flash("La clave primaria ya existe.", "danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        regs.append(fila)
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
        flash("Registro no encontrado.", "danger")
        return redirect(url_for("entidad_list", tipo=tipo))
    if request.method=="POST":
        for c in campos:
            v = request.form.get(c,"").strip()
            if v=="":
                flash(f"El campo {c} no puede estar vacío.", "danger")
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

# ─── 8. Generación de documentos ─────────────────────────────────────────────
def _render_docs(tpl, prefijo, df, extra=None):
    extra = extra or {}
    for _, fila in df.iterrows():
        ctx = {k: str(fila.get(k,"")) for k in MASTER_FIELDS}
        ctx.update(extra)
        doc = DocxTemplate(tpl)
        doc.render(ctx)
        nombre = f"{prefijo}_{fila.get('No_Control','')}.docx"
        doc.save(os.path.join(OUTPUT_DIR, nombre))

@app.route("/generar/solicitud")
def generar_solicitud():
    df = pd.read_csv(ALUMNOS_CSV)
    _render_docs(SOLICITUD_TPL, "Solicitud", df)
    flash("Solicitudes generadas.", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/bimestral/<int:periodo>")
def generar_bimestral(periodo):
    df = pd.read_csv(ALUMNOS_CSV)
    extra = {
        "Reporte_No": periodo,
        "x1": "X" if periodo==1 else "",
        "x2": "X" if periodo==2 else "",
        "x3": "X" if periodo==3 else ""
    }
    _render_docs(BIMESTRAL_TPL, f"Bimestral{periodo}", df, extra)
    flash(f"Reportes bimestrales {periodo} generados.", "success")
    return redirect(url_for("dashboard"))

@app.route("/generar/final")
def generar_final():
    df = pd.read_csv(ALUMNOS_CSV)
    _render_docs(FINAL_TPL, "Final", df)
    flash("Reportes finales generados.", "success")
    return redirect(url_for("dashboard"))

@app.route("/descargar")
def descargar():
    mem = BytesIO()
    with zipfile.ZipFile(mem, "w") as zf:
        for fn in os.listdir(OUTPUT_DIR):
            zf.write(os.path.join(OUTPUT_DIR, fn), fn)
    mem.seek(0)
    return send_file(mem, download_name="documentos.zip", as_attachment=True)

# ─── 9. Manejador 404 y arranque ─────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", 5000)),
        debug=False
    )
