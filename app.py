import os
import uuid
import csv
import zipfile
from io import BytesIO

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file
)
import msal
import pandas as pd
from docxtpl import DocxTemplate

# ──────────────────── 1. Configuración de Flask ────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "clave_segura")

# ──────────────────── 2. Configuración de Azure AD / MSAL ──────────────
CLIENT_ID     = os.environ.get("CLIENT_ID",     "tu_client_id")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "tu_client_secret")
TENANT_ID     = os.environ.get("TENANT_ID",     "tu_tenant_id")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "https://tu-app.onrender.com/getAToken")
SCOPE         = ["User.Read"]

# ──────────────────── 3. Rutas de archivos ────────────────────────────
BASE_PATH      = os.path.dirname(__file__)
CSV_DIR        = os.environ.get("CSV_DIR", "csv")
CSV_PATH       = os.path.join(BASE_PATH, CSV_DIR)
PLANTILLAS_DIR = os.path.join(BASE_PATH, "plantillas")
OUTPUT_PATH    = os.path.join(BASE_PATH, "documentos_generados")

os.makedirs(CSV_PATH, exist_ok=True)
os.makedirs(PLANTILLAS_DIR, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

USUARIOS_CSV    = os.path.join(CSV_PATH, "usuarios.csv")
PROFESORES_CSV  = os.path.join(CSV_PATH, "profesores.csv")
AREAS_CSV       = os.path.join(CSV_PATH, "areas.csv")
ALUMNOS_CSV     = os.path.join(CSV_PATH, "alumnos.csv")

TEMPLATE_SOLICITUD = os.path.join(PLANTILLAS_DIR, "plantilla_solicitud_completa.docx")
TEMPLATE_BIMESTRAL = os.path.join(PLANTILLAS_DIR, "Reporte_Bimestral_Plantilla.docx")
TEMPLATE_FINAL     = os.path.join(PLANTILLAS_DIR, "Reporte_Final_Plantilla.docx")

# ──────────────────── 4. Lista de campos de alumnos ───────────────────
CAMPOS_ALUMNOS = [
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo","Domicilio","No_Control",
    "Carrera","Periodo","Semestre","Creditos","Dependencia","Domicilio_Dependencia","Titular_Dependencia",
    "Cargo_Responsable","Responsable_Proyecto","Nombre_Programa",
    "Modalidad_Externa","Modalidad_Interna","Fecha_Inicio","Fecha_Terminacion","Actividades",
    "TP_Edu_Adultos","TP_Deportivo","TP_Civico","TP_Salud","TP_Otros",
    "TP_Desarrollo","TP_Cultural","TP_Sustentable","TP_Medio_Amb",
    "Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
    # los siguientes son para bimestral y final, si los incluyes en el mismo CSV
    "AP","AM","carrera","día1","mes1","año1","dia2","mes2","año2",
    "Nombre_supervisor","Puesto_supervisor",
    "Actividad_1","Actividad_2","Actividad_3","Actividad_4",
    "Actividad_5","Actividad_6","Actividad_7","Actividad_8",
    "x1","x2","x3",
    # autoevaluaciones y aprendizajes para final
    "Actividad1","Logro1","Actividad2","Logro2","Actividad3","Logro3","Actividad4","Logro4",
    "Actividad5","Logro5","Actividad6","Logro6","Actividad7","Logro7","Actividad8","Logro8",
    "Aprendizaje1","Beneficio1","Aprendizaje2","Beneficio2","Aprendizaje3","Beneficio3",
    "Aprendizaje4","Beneficio4","Aprendizaje5","Beneficio5","Aprendizaje6","Beneficio6",
    "Aprendizaje7","Beneficio7","Aprendizaje8","Beneficio8"
]

# ──────────────────── 5. Utilidades CSV ───────────────────────────────
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

# ──────────────────── 6. Sesión y roles ───────────────────────────────
def usuario_desde_csv(email):
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo","").strip().lower() == email.lower():
            return u
    return None

def validar_acceso(roles):
    u = session.get("user")
    return u and u.get("rol") in roles

# ──────────────────── 7. Login / Logout ───────────────────────────────
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
        return f"Error MSAL: {result.get('error_description')}", 400

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas @aguascalientes.tecnm.mx", 403

    perfil = usuario_desde_csv(email) or {
        "correo": email, "rol": "Alumno", "area": ""
    }
    session["user"] = {
        "correo": perfil["correo"],
        "rol"   : perfil["rol"],
        "area"  : perfil.get("area",""),
        "name"  : claims.get("name","")
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ──────────────────── 8. Dashboard ────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ──────────────────── 9. CRUD Genérico (Áreas, Profesores, Usuarios, Alumnos) ─
def mapa_entidades(tipo):
    return {
        "usuarios"  : (USUARIOS_CSV,   ["correo","rol","area","profesor"]),
        "profesores": (PROFESORES_CSV, ["correo","nombre","area"]),
        "areas"     : (AREAS_CSV,      ["id","nombre","encargado"]),
        "alumnos"   : (ALUMNOS_CSV,    CAMPOS_ALUMNOS)
    }.get(tipo)

@app.route("/entidad/<tipo>")
def entidad_list(tipo):
    m = mapa_entidades(tipo)
    if not m:
        return render_template("404.html"), 404
    archivo, campos = m
    registros = cargar_csv(archivo)
    return render_template("entidad_list.html",
                           tipo=tipo, campos=campos, registros=registros)

@app.route("/entidad/<tipo>/new", methods=["GET","POST"])
def entidad_new(tipo):
    m = mapa_entidades(tipo)
    if not m:
        return render_template("404.html"), 404
    archivo, campos = m

    if request.method == "POST":
        # tomo campos no TP_
        fila = {c: request.form.get(c,"").strip()
                for c in campos if not c.startswith("TP_")}
        # validación
        if any(v=="" for v in fila.values()):
            flash("Todos los campos obligatorios", "danger")
            return redirect(url_for("entidad_new", tipo=tipo))

        # casillas TP_
        tp_campos = [c for c in campos if c.startswith("TP_")]
        sel = request.form.get("tipo_participacion")
        for c in tp_campos:
            fila[c] = "X" if c==sel else ""

        regs = cargar_csv(archivo)
        pk = campos[0]
        if pk=="id":
            fila["id"] = str(uuid.uuid4())
        else:
            if any(r.get(pk)==fila[pk] for r in regs):
                flash(f"Ya existe {fila[pk]}", "danger")
                return redirect(url_for("entidad_new", tipo=tipo))

        regs.append(fila)
        guardar_csv(archivo, regs, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    tp_campos = [c for c in campos if c.startswith("TP_")]
    return render_template("entidad_form.html",
                           tipo=tipo, campos=campos,
                           valores={}, tp_fields=tp_campos)

@app.route("/entidad/<tipo>/edit/<pk>", methods=["GET","POST"])
def entidad_edit(tipo, pk):
    m = mapa_entidades(tipo)
    if not m:
        return render_template("404.html"), 404
    archivo, campos = m
    pk_field = campos[0]
    regs = cargar_csv(archivo)
    item = next((r for r in regs if r.get(pk_field)==pk), None)
    if not item:
        flash("No encontrado", "danger")
        return redirect(url_for("entidad_list", tipo=tipo))

    if request.method=="POST":
        for c in campos:
            if not c.startswith("TP_"):
                v = request.form.get(c,"").strip()
                if v=="":
                    flash(f"El campo {c} no puede quedar vacío", "danger")
                    return redirect(url_for("entidad_edit", tipo=tipo, pk=pk))
                item[c] = v
        tp_campos = [c for c in campos if c.startswith("TP_")]
        sel = request.form.get("tipo_participacion")
        for c in tp_campos:
            item[c] = "X" if c==sel else ""

        guardar_csv(archivo, regs, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    tp_campos = [c for c in campos if c.startswith("TP_")]
    return render_template("entidad_form.html",
                           tipo=tipo, campos=campos,
                           valores=item, tp_fields=tp_campos)

@app.route("/entidad/<tipo>/delete/<pk>", methods=["POST"])
def entidad_delete(tipo, pk):
    m = mapa_entidades(tipo)
    if not m:
        return render_template("404.html"), 404
    archivo, campos = m
    pk_field = campos[0]
    regs = [r for r in cargar_csv(archivo) if r.get(pk_field)!=pk]
    guardar_csv(archivo, regs, campos)
    return redirect(url_for("entidad_list", tipo=tipo))

# ──────────────────── 10. Generación de documentos Word ────────────────
def _make_doc(template_path, contexto, nombre_salida):
    doc = DocxTemplate(template_path)
    # convierto todo a cadena
    ctx = {k: str(v) for k,v in contexto.items()}
    doc.render(ctx)
    out_path = os.path.join(OUTPUT_PATH, nombre_salida)
    doc.save(out_path)
    return out_path

@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("dashboard"))
    out = _make_doc(TEMPLATE_SOLICITUD, fila,
                    f"Solicitud_{no_control}.docx")
    return send_file(out, as_attachment=True)

@app.route("/generar_bimestral/<int:bimestre>/<no_control>")
def generar_bimestral(bimestre, no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("dashboard"))
    # añado número de reporte
    fila["Reporte_No"] = bimestre
    out = _make_doc(TEMPLATE_BIMESTRAL, fila,
                    f"Reporte_Bimestral_{bimestre}_{no_control}.docx")
    return send_file(out, as_attachment=True)

@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    filas = cargar_csv(ALUMNOS_CSV)
    fila = next((r for r in filas if r.get("No_Control")==no_control), None)
    if not fila:
        flash("Alumno no encontrado", "danger")
        return redirect(url_for("dashboard"))
    out = _make_doc(TEMPLATE_FINAL, fila,
                    f"Reporte_Final_{no_control}.docx")
    return send_file(out, as_attachment=True)

# ──────────────────── 11. Descarga masiva (Solicitudes) ───────────────
@app.route("/generar_documentos_masivos")
def generar_documentos_masivos():
    rows = cargar_csv(ALUMNOS_CSV)
    memoria = BytesIO()
    with zipfile.ZipFile(memoria, "w") as zf:
        for r in rows:
            nc = r.get("No_Control")
            path = _make_doc(TEMPLATE_SOLICITUD, r, f"Solicitud_{nc}.docx")
            zf.write(path, arcname=os.path.basename(path))
    memoria.seek(0)
    return send_file(memoria, download_name="solicitudes.zip", as_attachment=True)

# ──────────────────── 12. Error 404 ────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

# ──────────────────── 13. Ejecutar ────────────────────────────────────
if __name__ == "__main__":
    puerto = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=puerto, debug=True)
