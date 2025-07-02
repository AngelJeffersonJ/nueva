import os
import csv
import uuid
import msal
import zipfile
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from io import BytesIO
from docxtpl import DocxTemplate

# ── Configuración de Flask y secretos ───────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "💥-cambia-este-secreto-💥")

# ── Azure AD (MSAL) ─────────────────────────────────────────────────────────
CLIENT_ID     = os.environ.get("CLIENT_ID",     "c306c8d3-68dc-4110-b5fb-771b942c10db")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "CEX8Q~txlNOmiN4PAZKbV7matISmRV1HREkvbcMS")
TENANT_ID     = os.environ.get("TENANT_ID",     "63de1475-1a48-4463-aff2-b2581f2a972e")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.environ.get("REDIRECT_URI",  "https://sistema-documental.onrender.com/getAToken")
SCOPE         = ["User.Read"]

# ── Rutas de archivos y directorios ────────────────────────────────────────
BASE_DIR          = os.path.dirname(__file__)
CSV_DIR           = os.environ.get("CSV_DIR", os.path.join(BASE_DIR, "csv"))
PLANTILLAS_DIR    = os.path.join(BASE_DIR, "plantillas")
OUTPUT_DIR        = os.path.join(BASE_DIR, "documentos_generados")

os.makedirs(CSV_DIR,        exist_ok=True)
os.makedirs(PLANTILLAS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR,     exist_ok=True)

USUARIOS_CSV        = os.path.join(CSV_DIR, "usuarios.csv")
PROFESORES_CSV      = os.path.join(CSV_DIR, "profesores.csv")
AREAS_CSV           = os.path.join(CSV_DIR, "areas.csv")
ALUMNOS_CSV         = os.path.join(CSV_DIR, "alumnos.csv")

SOLICITUD_PLANTILLA = os.path.join(PLANTILLAS_DIR, "plantilla_solicitud_completa.docx")
BIMESTRAL_PLANTILLA = os.path.join(PLANTILLAS_DIR, "Reporte_Bimestral_Plantilla.docx")
FINAL_PLANTILLA     = os.path.join(PLANTILLAS_DIR, "Reporte_Final_Plantilla.docx")

# ── Campos para cada CSV ─────────────────────────────────────────────────────
CAMPOS_USUARIOS   = ["correo","rol","area"]
CAMPOS_PROFESORES = ["correo","nombre","area"]
CAMPOS_AREAS      = ["id","nombre","encargado"]

# CAMPOS_ALUMNOS es la unión de todos los campos que usan las 3 plantillas
CAMPOS_ALUMNOS = [
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo","Domicilio","No_Control",
    "Carrera","Periodo","Semestre","Creditos","Dependencia","Domicilio_Dependencia","Titular_Dependencia",
    "Cargo_Responsable","Responsable_Proyecto","Nombre_Programa","Modalidad_Externa","Modalidad_Interna",
    "Fecha_Inicio","Fecha_Terminacion","Actividades",
    # Casillas de tipo de proyecto
    "TP_Edu_Adultos","TP_Deportivo","TP_Civico","TP_Salud","TP_Otros","TP_Desarrollo","TP_Cultural","TP_Sustentable","TP_Medio_Amb",
    # Fecha de solicitud
    "Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
    # Campos para reporte bimestral
    "AP","AM","carrera","día1","mes1","año1","dia2","mes2","año2",
    "Actividad_1","Actividad_2","Actividad_3","Actividad_4","Actividad_5","Actividad_6","Actividad_7","Actividad_8",
    "Reporte_No","Nombre_supervisor","Puesto_supervisor",
    # Campos para reporte final
    "Municipio","Estado","Fecha","Aprendizaje1","Beneficio1","Aprendizaje2","Beneficio2",
    "Aprendizaje3","Beneficio3","Aprendizaje4","Beneficio4","Aprendizaje5","Beneficio5",
    "Aprendizaje6","Beneficio6","Aprendizaje7","Beneficio7","Aprendizaje8","Beneficio8",
    # Autoevaluaciones para bimestres (ejemplo)
    "x1","x2","x3"
]

# ── Mapa genérico de entidades para CRUD ────────────────────────────────────
MAPA_ENTIDADES = {
    "usuarios":   (USUARIOS_CSV,   CAMPOS_USUARIOS),
    "profesores": (PROFESORES_CSV, CAMPOS_PROFESORES),
    "areas":      (AREAS_CSV,      CAMPOS_AREAS),
    "alumnos":    (ALUMNOS_CSV,    CAMPOS_ALUMNOS),
}

# ── Utilerías de CSV ────────────────────────────────────────────────────────
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

# ── Autenticación y roles ───────────────────────────────────────────────────
def usuario_desde_csv(email):
    # Busca en el CSV de alumnos un perfil (puedes usar otro CSV si lo prefieres)
    for u in cargar_csv(ALUMNOS_CSV):
        if u.get("Correo","").strip().lower() == email.lower():
            return {"correo":u["Correo"], "rol":u.get("rol","Alumno"), "area":u.get("area","")}
    return None

def validar_acceso(roles:list[str]) -> bool:
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# ── Flujo de login con MSAL ─────────────────────────────────────────────────
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
        flash(result.get("error_description","Error de autenticación"), "danger")
        return redirect(url_for("dashboard"))

    claims = result["id_token_claims"]
    email  = claims.get("preferred_username","").lower()
    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo cuentas @aguascalientes.tecnm.mx permitidas", 403

    perfil = usuario_desde_csv(email) or {"correo":email,"rol":"Alumno","area":""}
    session["user"] = {
        "correo": perfil["correo"],
        "rol":    perfil["rol"],
        "area":   perfil["area"],
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

# ── Dashboard ───────────────────────────────────────────────────────────────
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ── CRUD genérico de entidades ─────────────────────────────────────────────
@app.route("/entidad/<tipo>")
def entidad_list(tipo):
    m = MAPA_ENTIDADES.get(tipo)
    if not m:
        return render_template("404.html"), 404
    path, fields = m
    filas = cargar_csv(path)
    return render_template("entidad_list.html", tipo=tipo, campos=fields, registros=filas)

@app.route("/entidad/<tipo>/new", methods=["GET","POST"])
def entidad_new(tipo):
    m = MAPA_ENTIDADES.get(tipo)
    if not m:
        return render_template("404.html"), 404
    path, fields = m
    if request.method=="POST":
        filas = cargar_csv(path)
        nueva = {c: request.form.get(c,"").strip() for c in fields if not c.startswith("TP_")}
        # validación básica
        if any(v=="" for v in nueva.values()):
            flash("Todos los campos son obligatorios","danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        # casillas TP_
        tp = [c for c in fields if c.startswith("TP_")]
        sel = request.form.get("tipo_participacion","")
        for c in tp:
            nueva[c] = "X" if c==sel else ""
        # PK y unicidad
        pk = fields[0]
        if pk=="id":
            nueva["id"] = str(uuid.uuid4())
        clave = nueva.get(pk,"")
        if any(r.get(pk,"")==clave for r in filas):
            flash(f"Ya existe un registro con {pk}={clave}","danger")
            return redirect(url_for("entidad_new", tipo=tipo))
        filas.append(nueva)
        guardar_csv(path, filas, fields)
        return redirect(url_for("entidad_list", tipo=tipo))
    tp_fields = [c for c in fields if c.startswith("TP_")]
    return render_template("entidad_form.html",
                           tipo=tipo, campos=fields,
                           valores={}, tp_fields=tp_fields)

@app.route("/entidad/<tipo>/edit/<pk>", methods=["GET","POST"])
def entidad_edit(tipo, pk):
    m = MAPA_ENTIDADES.get(tipo)
    if not m:
        return render_template("404.html"), 404
    path, fields = m
    filas = cargar_csv(path)
    item = next((r for r in filas if r.get(fields[0])==pk), None)
    if not item:
        flash("Registro no encontrado","danger")
        return redirect(url_for("entidad_list", tipo=tipo))
    if request.method=="POST":
        # actualizar campos
        for c in fields:
            if not c.startswith("TP_"):
                v = request.form.get(c,"").strip()
                if v=="":
                    flash(f"El campo {c} no puede quedar vacío","danger")
                    return redirect(url_for("entidad_edit", tipo=tipo, pk=pk))
                item[c] = v
        tp_fields = [c for c in fields if c.startswith("TP_")]
        sel = request.form.get("tipo_participacion","")
        for c in tp_fields:
            item[c] = "X" if c==sel else ""
        guardar_csv(path, filas, fields)
        return redirect(url_for("entidad_list", tipo=tipo))
    tp_fields = [c for c in fields if c.startswith("TP_")]
    return render_template("entidad_form.html",
                           tipo=tipo, campos=fields,
                           valores=item, tp_fields=tp_fields)

@app.route("/entidad/<tipo>/delete/<pk>", methods=["POST"])
def entidad_delete(tipo, pk):
    m = MAPA_ENTIDADES.get(tipo)
    if not m:
        return render_template("404.html"), 404
    path, fields = m
    filas = [r for r in cargar_csv(path) if r.get(fields[0])!=pk]
    guardar_csv(path, filas, fields)
    return redirect(url_for("entidad_list", tipo=tipo))

# ── Generación de documentos ───────────────────────────────────────────────
def _make_doc(template, fila, salida):
    ctx = {k: fila.get(k,"") for k in fila.keys()}
    doc = DocxTemplate(template)
    doc.render(ctx)
    dest = os.path.join(OUTPUT_DIR, salida)
    doc.save(dest)
    return dest

@app.route("/generar_solicitud/<no_control>")
def generar_solicitud(no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV)
                 if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No. de control no encontrado","danger")
        return redirect(url_for("dashboard"))
    path = _make_doc(SOLICITUD_PLANTILLA, fila, f"Solicitud_{no_control}.docx")
    return send_file(path, as_attachment=True)

@app.route("/generar_bimestral/<rep>/<no_control>")
def generar_bimestral(rep, no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV)
                 if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No. de control no encontrado","danger")
        return redirect(url_for("dashboard"))
    copy = fila.copy()
    copy["Reporte_No"] = rep
    path = _make_doc(BIMESTRAL_PLANTILLA, copy,
                     f"Reporte_Bimestral_{rep}_{no_control}.docx")
    return send_file(path, as_attachment=True)

@app.route("/generar_final/<no_control>")
def generar_final(no_control):
    fila = next((r for r in cargar_csv(ALUMNOS_CSV)
                 if r.get("No_Control")==no_control), None)
    if not fila:
        flash("No. de control no encontrado","danger")
        return redirect(url_for("dashboard"))
    copy = fila.copy()
    # alias para campos de la plantilla final
    copy["Fecha"]   = copy.get("Fecha_Terminacion","")
    copy["Periodo"] = copy.get("Periodo","")
    path = _make_doc(FINAL_PLANTILLA, copy,
                     f"Reporte_Final_{no_control}.docx")
    return send_file(path, as_attachment=True)

# ── Descarga en ZIP (opcional) ─────────────────────────────────────────────
@app.route("/descargar_todos")
def descargar_todos():
    mem = BytesIO()
    with zipfile.ZipFile(mem, "w") as zf:
        for fn in os.listdir(OUTPUT_DIR):
            zf.write(os.path.join(OUTPUT_DIR, fn), fn)
    mem.seek(0)
    return send_file(mem, download_name="docs_servicio_social.zip", as_attachment=True)

# ── Error 404 ───────────────────────────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

if __name__ == "__main__":
    app.run(host="0.0.0.0",
            port=int(os.environ.get("PORT", 5000)),
            debug=True)
