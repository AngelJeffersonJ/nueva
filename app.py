"""
app.py – Flask + MSAL + CSV (producción)

• Login con cuentas Microsoft (@aguascalientes.tecnm.mx)
• CRUD completo sobre CSV
• Control de acceso por rol
"""

from flask import Flask, render_template, session, redirect, url_for, request, flash
from dotenv import load_dotenv
import msal, csv, os, uuid

# ── 1. Cargar variables de entorno ------------------------------------------------
load_dotenv()  # lee el .env si existe

# ── 2. Instancia Flask ------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "clave_segura")  # ¡cámbialo en producción!

# ── 3. Configuración Azure AD -----------------------------------------------------
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID     = os.getenv("TENANT_ID")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
REDIRECT_URI  = os.getenv("REDIRECT_URI", "http://localhost:5000/getAToken")
SCOPE         = ["User.Read"]

# ── 4. Rutas a CSV ----------------------------------------------------------------
CSV_DIR         = os.getenv("CSV_DIR", "csv")
USUARIOS_CSV    = f"{CSV_DIR}/usuarios.csv"
AREAS_CSV       = f"{CSV_DIR}/areas.csv"
PROFESORES_CSV  = f"{CSV_DIR}/profesores.csv"
ALUMNOS_CSV     = f"{CSV_DIR}/alumnos.csv"

# ── 5. Helpers CSV ---------------------------------------------------------------
def cargar_csv(path: str) -> list[dict]:
    """Devuelve una lista de dicts leída del CSV, o [] si no existe."""
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    return []

def guardar_csv(path: str, rows: list[dict], campos: list[str]) -> None:
    """Sobrescribe el CSV con los registros indicados."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(rows)

def usuario_desde_csv(email: str) -> dict | None:
    """Busca al usuario en usuarios.csv; si existe devuelve su perfil."""
    for u in cargar_csv(USUARIOS_CSV):
        if u["correo"].strip().lower() == email.lower():
            return {"correo": u["correo"], "rol": u["rol"], "area": u.get("area", "")}
    return None

def validar_acceso(roles: list[str]) -> bool:
    """True si el usuario actual está logueado y tiene un rol permitido."""
    u = session.get("user")
    return u and u.get("rol") in roles

# ── 6. MSAL: login / callback / logout -------------------------------------------
@app.route("/login")
def login():
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route("/getAToken")
def authorized():
    code = request.args.get("code")
    if not code:
        flash("Código de autorización faltante", "danger")
        return redirect(url_for("login"))

    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=SCOPE, redirect_uri=REDIRECT_URI
    )

    # ── Manejo de error MSAL
    if "error" in result:
        return f"Error MSAL: {result.get('error_description')}", 500

    claims = result.get("id_token_claims", {})
    email  = claims.get("preferred_username", "").lower()

    if not email.endswith("@aguascalientes.tecnm.mx"):
        return "Solo se permiten cuentas institucionales", 403

    perfil = usuario_desde_csv(email) or {"correo": email, "rol": "Alumno", "area": ""}

    # Guardar sesión
    session["user"] = {
        "correo": perfil["correo"],
        "rol":    perfil["rol"],
        "area":   perfil["area"],
        "name":   claims.get("name", "")
    }
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('login', _external=True)}"
    )

# ── 7. Dashboard -----------------------------------------------------------------
@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", usuario=session["user"])

# ── 8. CRUD: Usuarios ------------------------------------------------------------
@app.route("/usuarios")
def usuarios_list():
    if not validar_acceso(["Administrador"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    return render_template("usuarios_list.html", usuarios=cargar_csv(USUARIOS_CSV))

@app.route("/usuarios/new", methods=["GET", "POST"])
def usuarios_new():
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        rows = cargar_csv(USUARIOS_CSV)
        rows.append({
            "correo": request.form["correo"],
            "rol":    request.form["rol"],
            "area":   request.form["area"]
        })
        guardar_csv(USUARIOS_CSV, rows, ["correo", "rol", "area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario={})

@app.route("/usuarios/edit/<correo>", methods=["GET", "POST"])
def usuarios_edit(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(USUARIOS_CSV)
    u    = next((r for r in rows if r["correo"] == correo), None)
    if not u:
        return "No existe", 404
    if request.method == "POST":
        u["rol"], u["area"] = request.form["rol"], request.form["area"]
        guardar_csv(USUARIOS_CSV, rows, ["correo", "rol", "area"])
        return redirect(url_for("usuarios_list"))
    return render_template("usuarios_form.html", usuario=u)

@app.route("/usuarios/delete/<correo>", methods=["POST"])
def usuarios_delete(correo):
    if not validar_acceso(["Administrador"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(USUARIOS_CSV) if r["correo"] != correo]
    guardar_csv(USUARIOS_CSV, rows, ["correo", "rol", "area"])
    return redirect(url_for("usuarios_list"))

# ── 9. CRUD: Áreas ---------------------------------------------------------------
@app.route("/areas")
def areas_list():
    if not validar_acceso(["Administrador", "Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    return render_template("areas_list.html", areas=cargar_csv(AREAS_CSV))

@app.route("/areas/new", methods=["GET", "POST"])
def areas_new():
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        rows = cargar_csv(AREAS_CSV)
        rows.append({
            "id":        str(uuid.uuid4()),
            "nombre":    request.form["nombre"],
            "encargado": request.form["encargado"]
        })
        guardar_csv(AREAS_CSV, rows, ["id", "nombre", "encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area={})

@app.route("/areas/edit/<id>", methods=["GET", "POST"])
def areas_edit(id):
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(AREAS_CSV)
    area = next((r for r in rows if r["id"] == id), None)
    if not area:
        return "No existe", 404
    if request.method == "POST":
        area["nombre"], area["encargado"] = request.form["nombre"], request.form["encargado"]
        guardar_csv(AREAS_CSV, rows, ["id", "nombre", "encargado"])
        return redirect(url_for("areas_list"))
    return render_template("areas_form.html", area=area)

@app.route("/areas/delete/<id>", methods=["POST"])
def areas_delete(id):
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(AREAS_CSV) if r["id"] != id]
    guardar_csv(AREAS_CSV, rows, ["id", "nombre", "encargado"])
    return redirect(url_for("areas_list"))

# ── 10. CRUD: Profesores ---------------------------------------------------------
@app.route("/profesores")
def profesores_list():
    if not validar_acceso(["Administrador", "Encargado"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    return render_template("profesores_list.html", profesores=cargar_csv(PROFESORES_CSV))

@app.route("/profesores/new", methods=["GET", "POST"])
def profesores_new():
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        rows = cargar_csv(PROFESORES_CSV)
        rows.append({
            "correo": request.form["correo"],
            "nombre": request.form["nombre"],
            "area":   request.form["area"]
        })
        guardar_csv(PROFESORES_CSV, rows, ["correo", "nombre", "area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor={})

@app.route("/profesores/edit/<correo>", methods=["GET", "POST"])
def profesores_edit(correo):
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(PROFESORES_CSV)
    p    = next((r for r in rows if r["correo"] == correo), None)
    if not p:
        return "No existe", 404
    if request.method == "POST":
        p["nombre"], p["area"] = request.form["nombre"], request.form["area"]
        guardar_csv(PROFESORES_CSV, rows, ["correo", "nombre", "area"])
        return redirect(url_for("profesores_list"))
    return render_template("profesores_form.html", profesor=p)

@app.route("/profesores/delete/<correo>", methods=["POST"])
def profesores_delete(correo):
    if not validar_acceso(["Administrador", "Encargado"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(PROFESORES_CSV) if r["correo"] != correo]
    guardar_csv(PROFESORES_CSV, rows, ["correo", "nombre", "area"])
    return redirect(url_for("profesores_list"))

# ── 11. CRUD: Alumnos ------------------------------------------------------------
@app.route("/alumnos")
def alumnos_list():
    if not validar_acceso(["Administrador", "Encargado", "Maestro"]):
        flash("Acceso denegado", "danger")
        return redirect(url_for("dashboard"))
    rows = cargar_csv(ALUMNOS_CSV)
    u    = session["user"]
    if u["rol"] == "Maestro":
        rows = [r for r in rows if r["profesor"] == u["correo"]]
    if u["rol"] == "Encargado":
        rows = [r for r in rows if r["area"] == u["area"]]
    return render_template("alumnos_list.html", alumnos=rows)

@app.route("/alumnos/new", methods=["GET", "POST"])
def alumnos_new():
    if not validar_acceso(["Administrador", "Encargado", "Maestro"]):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        rows = cargar_csv(ALUMNOS_CSV)
        rows.append({
            "correo":   request.form["correo"],
            "nombre":   request.form["nombre"],
            "area":     request.form["area"],
            "profesor": request.form["profesor"]
        })
        guardar_csv(ALUMNOS_CSV, rows, ["correo", "nombre", "area", "profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno={})

@app.route("/alumnos/edit/<correo>", methods=["GET", "POST"])
def alumnos_edit(correo):
    if not validar_acceso(["Administrador", "Encargado", "Maestro"]):
        return redirect(url_for("dashboard"))
    rows = cargar_csv(ALUMNOS_CSV)
    a    = next((r for r in rows if r["correo"] == correo), None)
    if not a:
        return "No existe", 404
    if request.method == "POST":
        a["nombre"], a["area"], a["profesor"] = (
            request.form["nombre"],
            request.form["area"],
            request.form["profesor"]
        )
        guardar_csv(ALUMNOS_CSV, rows, ["correo", "nombre", "area", "profesor"])
        return redirect(url_for("alumnos_list"))
    return render_template("alumnos_form.html", alumno=a)

@app.route("/alumnos/delete/<correo>", methods=["POST"])
def alumnos_delete(correo):
    if not validar_acceso(["Administrador", "Encargado", "Maestro"]):
        return redirect(url_for("dashboard"))
    rows = [r for r in cargar_csv(ALUMNOS_CSV) if r["correo"] != correo]
    guardar_csv(ALUMNOS_CSV, rows, ["correo", "nombre", "area", "profesor"])
    return redirect(url_for("alumnos_list"))

# ── 12. Error 404 ---------------------------------------------------------------
@app.errorhandler(404)
def page_not_found(e):
    return render_template("404.html"), 404

# ── 13. Arranque local ----------------------------------------------------------
if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))   # Render define PORT automáticamente
    app.run(host="0.0.0.0", port=port, debug=False)
