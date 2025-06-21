"""
app.py – Flask + MSAL + CSV  (producción)

• Login Microsoft (@aguascalientes.tecnm.mx)
• CRUD genérico sobre CSV (usuarios, profesores, áreas, alumnos)
• Generación y descarga de solicitudes de servicio social (DOCX → ZIP)
"""

from __future__ import annotations

import csv
import os
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file,
)
import msal
import pandas as pd
from docxtpl import DocxTemplate

# ╔══════════════════════╗
# ║ 1. Cargar .env       ║
# ╚══════════════════════╝
load_dotenv()

# ─── Variables de entorno obligatorias ───
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")

for k, v in {"CLIENT_ID": CLIENT_ID, "CLIENT_SECRET": CLIENT_SECRET,
             "TENANT_ID": TENANT_ID}.items():
    if not v:
        raise RuntimeError(f"Variable de entorno {k} faltante")

# ╔══════════════════════╗
# ║ 2. Flask             ║
# ╚══════════════════════╝
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "REEMPLAZA-ESTE-SECRET")

# ╔══════════════════════╗
# ║ 3. Azure AD (MSAL)   ║
# ╚══════════════════════╝
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI = (
    os.getenv("REDIRECT_URI")
    or (f"https://{os.getenv('RENDER_EXTERNAL_HOSTNAME')}/getAToken"
        if os.getenv("RENDER_EXTERNAL_HOSTNAME")
        else "http://localhost:5000/getAToken")
)
SCOPE = ["User.Read"]

# ╔══════════════════════╗
# ║ 4. Paths & CSV       ║
# ╚══════════════════════╝
BASE_DIR      = Path(__file__).resolve().parent
CSV_DIR       = BASE_DIR / "csv"
TPL_DIR       = BASE_DIR / "plantillas"
OUT_DIR       = BASE_DIR / "documentos_generados"

CSV_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

USUARIOS_CSV   = CSV_DIR / "usuarios.csv"
PROFESORES_CSV = CSV_DIR / "profesores.csv"
AREAS_CSV      = CSV_DIR / "areas.csv"
ALUMNOS_CSV    = CSV_DIR / "alumnos.csv"

CAMPOS_ALUMNOS = [
    "Apellido_Paterno","Apellido_Materno","Nombre","Sexo","Telefono","Correo",
    "Domicilio","No_Control","Carrera","Periodo","Semestre","Creditos",
    "Dependencia","Domicilio_Dependencia","Titular_Dependencia",
    "Director_Dependencia","Responsable_Proyecto","Cargo_Responsable",
    "Nombre_Programa","Modalidad_Externa","Modalidad_Interna",
    "Fecha_Inicio","Fecha_Terminacion","Actividades",
    "TP_Edu_Adultos","TP_Deportivo","TP_Civico","TP_Salud","TP_Otros",
    "TP_Desarrollo","TP_Cultural","TP_Sustentable","TP_Medio_Amb",
    "Dia_Solicitud","Mes_Solicitud","Anio_Solicitud",
]

# ╔══════════════════════╗
# ║ 5. Utilidades CSV    ║
# ╚══════════════════════╝
def cargar_csv(path: Path) -> list[dict]:
    if path.exists():
        with path.open(newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path: Path, rows: list[dict], campos: list[str]) -> None:
    path.parent.mkdir(exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(rows)

# ╔══════════════════════╗
# ║ 6. Helpers sesión    ║
# ╚══════════════════════╝
def usuario_desde_csv(email: str) -> dict | None:
    email = email.lower().strip()
    for u in cargar_csv(USUARIOS_CSV):
        if u.get("correo", "").lower().strip() == email:
            return u
    return None

def validar_acceso(roles: list[str]) -> bool:
    u = session.get("user")
    return bool(u and u.get("rol") in roles)

# ╔══════════════════════╗
# ║ 7.  MSAL - Login     ║
# ╚══════════════════════╝
@app.route("/login")
def login() -> str:
    auth_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = auth_app.get_authorization_request_url(
        SCOPE, redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route("/getAToken")
def get_token() -> str:
    code = request.args.get("code")
    if not code:
        flash("Código de autorización faltante", "danger")
        return redirect(url_for("login"))

    auth_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = auth_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )

    if "error" in result:
        return f"Error MSAL: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username", "").lower()

    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo se permiten cuentas institucionales", 403

    perfil = usuario_desde_csv(email) or {"correo": email, "rol": "Alumno", "area": ""}

    session["user"] = {
        "correo": perfil["correo"],
        "rol":    perfil["rol"],
        "area":   perfil["area"],
        "name":   claims.get("name", ""),
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout() -> str:
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ╔══════════════════════╗
# ║ 8.  Dashboard        ║
# ╚══════════════════════╝
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ╔══════════════════════╗
# ║ 9. CRUD genérico     ║
# ╚══════════════════════╝
def mapa_csv(tipo: str):
    return {
        "usuarios":   (USUARIOS_CSV,   ["correo","rol","area","profesor"]),
        "profesores": (PROFESORES_CSV, ["correo","nombre","area"]),
        "areas":      (AREAS_CSV,      ["id","nombre","encargado"]),
        "alumnos":    (ALUMNOS_CSV,    CAMPOS_ALUMNOS),
    }.get(tipo)

@app.route("/entidad/<tipo>")
def entidad_list(tipo):
    m = mapa_csv(tipo)
    if not m:
        return render_template("404.html"), 404
    path, campos = m
    return render_template(
        "entidad_list.html",
        tipo=tipo, campos=campos, registros=cargar_csv(path)
    )

@app.route("/entidad/<tipo>/new", methods=["GET", "POST"])
def entidad_new(tipo):
    m = mapa_csv(tipo)
    if not m:
        return render_template("404.html"), 404
    path, campos = m

    if request.method == "POST":
        row = {c: request.form.get(c, "").strip() for c in campos if not c.startswith("TP_")}
        if any(v == "" for v in row.values()):
            flash("Todos los campos son obligatorios", "danger")
            return redirect(url_for("entidad_new", tipo=tipo))

        # casillas TP_
        tp_fields = [c for c in campos if c.startswith("TP_")]
        selected  = request.form.get("tipo_participacion")
        for c in tp_fields:
            row[c] = "X" if c == selected else ""

        registros = cargar_csv(path)
        pk = campos[0]
        if pk == "id":
            row["id"] = str(uuid.uuid4())
        elif any(r.get(pk) == row[pk] for r in registros):
            flash(f"El valor {row[pk]} ya existe", "danger")
            return redirect(url_for("entidad_new", tipo=tipo))

        registros.append(row)
        guardar_csv(path, registros, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    return render_template(
        "entidad_form.html", tipo=tipo,
        campos=campos, valores={}, tp_fields=[c for c in campos if c.startswith("TP_")]
    )

@app.route("/entidad/<tipo>/edit/<pk>", methods=["GET", "POST"])
def entidad_edit(tipo, pk):
    m = mapa_csv(tipo)
    if not m:
        return render_template("404.html"), 404
    path, campos = m
    pk_field = campos[0]

    registros = cargar_csv(path)
    item = next((r for r in registros if r.get(pk_field) == pk), None)
    if not item:
        flash("Registro no encontrado", "danger")
        return redirect(url_for("entidad_list", tipo=tipo))

    if request.method == "POST":
        for c in campos:
            if c.startswith("TP_"):
                continue
            val = request.form.get(c, "").strip()
            if val == "":
                flash(f"El campo {c} no puede quedar vacío", "danger")
                return redirect(url_for("entidad_edit", tipo=tipo, pk=pk))
            item[c] = val

        tp_fields = [c for c in campos if c.startswith("TP_")]
        selected = request.form.get("tipo_participacion")
        for c in tp_fields:
            item[c] = "X" if c == selected else ""

        guardar_csv(path, registros, campos)
        return redirect(url_for("entidad_list", tipo=tipo))

    return render_template(
        "entidad_form.html", tipo=tipo,
        campos=campos, valores=item, tp_fields=[c for c in campos if c.startswith("TP_")]
    )

@app.route("/entidad/<tipo>/delete/<pk>", methods=["POST"])
def entidad_delete(tipo, pk):
    m = mapa_csv(tipo)
    if not m:
        return render_template("404.html"), 404
    path, campos = m
    pk_field = campos[0]

    registros = [r for r in cargar_csv(path) if r.get(pk_field) != pk]
    guardar_csv(path, registros, campos)
    return redirect(url_for("entidad_list", tipo=tipo))

# ╔══════════════════════╗
# ║10. DOCX / ZIP        ║
# ╚══════════════════════╝
TEMPLATE_FILE = TPL_DIR / "plantilla_solicitud_completa.docx"

@app.route("/generar-documentos")
def generar_documentos():
    try:
        df = pd.read_csv(ALUMNOS_CSV)
        for _, alumno in df.iterrows():
            doc = DocxTemplate(TEMPLATE_FILE)
            context = {k: str(v) for k, v in alumno.items()}
            doc.render(context)
            doc.save(OUT_DIR / f"Solicitud_{alumno['No_Control']}.docx")
        flash("Documentos generados satisfactoriamente", "success")
    except Exception as exc:
        flash(f"Error generando documentos: {exc}", "danger")
    return redirect(url_for("dashboard"))

@app.route("/descargar-documentos")
def descargar_documentos():
    mem = BytesIO()
    with zipfile.ZipFile(mem, "w") as zf:
        for fn in OUT_DIR.iterdir():
            zf.write(fn, arcname=fn.name)
    mem.seek(0)
    return send_file(mem, download_name="documentos_alumnos.zip", as_attachment=True)

# ╔══════════════════════╗
# ║11. Error 404         ║
# ╚══════════════════════╝
@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

# ╔══════════════════════╗
# ║12. Arranque          ║
# ╚══════════════════════╝
if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))  # Render define PORT automáticamente
    app.run(host="0.0.0.0", port=port, debug=False)
