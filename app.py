# app.py
"""Entry-point Flask.

Nuova nomenclatura pagine:
- /MedicairGeek                -> Home (templates/home.html)
- /MedicairGeek/reportIncidente -> Report incidenti (Blueprint: report_incidente_bp)
- /MedicairGeek/reportIntervento -> Report intervento (placeholder, templates/reportIntervento.html)
- /MedicairGeek/storicoReport  -> Storico Report (Blueprint: consulta_report_bp)
"""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, render_template, redirect, request, session, url_for

from log_utils import log_activity
from report_incidente import report_incidente_bp
from consulta_report import consulta_report_bp
from info_impianto import ensure_info_impianto_cache, info_impianto_bp


BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.txt"
AUTH_ENABLED = False  # Imposta True per abilitare il login, False per bypassare l'autenticazione

LOGIN_EXEMPT = [
    "/MedicairGeek/login",
    "/MedicairGeek/logout",
    "/MedicairGeek/static/",
]


def load_credentials() -> dict[str, str]:
    if not CREDENTIALS_FILE.exists():
        return {}

    credentials: dict[str, str] = {}
    text = CREDENTIALS_FILE.read_text(encoding="utf-8")
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            continue
        username, password = line.split(":", 1)
        credentials[username.strip()] = password.strip()
    return credentials


def is_exempt_path(path: str) -> bool:
    if path.startswith("/MedicairGeek/static/"):
        return True
    return any(path == exempt or path.startswith(exempt) for exempt in LOGIN_EXEMPT)


def create_app() -> Flask:
    app = Flask(__name__, static_folder="static", static_url_path="/MedicairGeek/static")
    app.config["SECRET_KEY"] = "CHANGE_ME__report_medicair_secret"
    app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300MB
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)

    @app.before_request
    def require_login():
        if not AUTH_ENABLED:
            return None
        if is_exempt_path(request.path):
            return None
        if session.get("user"):
            return None
        return redirect(url_for("login", next=request.path))

    @app.get("/MedicairGeek/login")
    def login():
        if session.get("user"):
            return redirect(url_for("home"))
        return render_template("login.html", error=None, next=request.args.get("next", ""))

    @app.post("/MedicairGeek/login")
    def login_post():
        username = (request.form.get("username") or "").strip()
        password = (request.form.get("password") or "").strip()
        next_url = request.form.get("next") or request.args.get("next") or url_for("home")

        credentials = load_credentials()
        if username and password and credentials.get(username) == password:
            session.permanent = True
            session["user"] = username
            log_activity(f"login_success | user={username} | ip={request.remote_addr} | next={next_url}")
            return redirect(next_url)

        log_activity(f"login_failure | user={username or 'unknown'} | ip={request.remote_addr} | next={next_url}")
        return render_template(
            "login.html",
            error="Utente o password non validi.",
            next=request.form.get("next", ""),
        )

    @app.get("/MedicairGeek/logout")
    def logout():
        user = session.get("user")
        session.clear()
        if user:
            log_activity(f"logout | user={user} | ip={request.remote_addr}")
        return redirect(url_for("login"))

    try:
        ensure_info_impianto_cache()
    except Exception as exc:
        print(f"[InfoImpianto] Cache init warning: {exc}")

    # Blueprint: report incidente
    app.register_blueprint(report_incidente_bp)

    # Blueprint: consultazione / storico
    app.register_blueprint(consulta_report_bp)

    # Blueprint: info impianto
    app.register_blueprint(info_impianto_bp)

    # Home
    @app.get("/MedicairGeek")
    @app.get("/MedicairGeek/")
    def home():
        return render_template("home.html", title="MedicairGeek", now=datetime.now())

    # Report Intervento (placeholder)
    @app.get("/MedicairGeek/reportIntervento")
    def report_intervento_form():
        return render_template("reportIntervento.html", title="Report Intervento", now=datetime.now())

    return app


app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3570, debug=False)