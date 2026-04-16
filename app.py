# app.py
"""Entry-point Flask.

Nuova nomenclatura pagine:
- /MedicairGeek                -> Home (templates/home.html)
- /MedicairGeek/reportIncidente -> Report incidenti (Blueprint: report_incidente_bp)
- /MedicairGeek/reportIntervento -> Report intervento (placeholder, templates/reportIntervento.html)
- /MedicairGeek/storicoReport  -> Storico Report (Blueprint: consulta_report_bp)
- /MedicairGeek/disallineamentoQr -> Disallineamento QR (Blueprint: disallineamento_qr_bp)
"""

from __future__ import annotations

import base64
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, render_template, redirect, request, session, url_for

from log_utils import log_activity
from report_incidente import report_incidente_bp
from consulta_report import consulta_report_bp
from info_impianto import ensure_info_impianto_cache, info_impianto_bp
from disallineamento_qr import disallineamento_qr_bp


BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.txt"
AUTH_ENABLED = False  # Imposta True per abilitare il login, False per bypassare l'autenticazione
#Operatore:ReportGeek --> http://localhost:3570/MedicairGeek/quick-login?token=T3BlcmF0b3JlOlJlcG9ydEdlZWs=  
LOGIN_EXEMPT = [
    "/MedicairGeek/login",
    "/MedicairGeek/logout",
    "/MedicairGeek/quick-login",
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

    @app.get("/MedicairGeek/quick-login")
    def quick_login():
        """Quick login con credenziali codificate in base64.
        
        Parametro: token (base64 di username:password)
        Esempio: /MedicairGeek/quick-login?token=T3BlcmF0b3JlOlJlcG9ydEdlZWs=
        """
        token = request.args.get("token", "").strip()
        redirect_to = request.args.get("next", url_for("home"))

        if not token:
            log_activity(f"quick_login_failure | reason=no_token | ip={request.remote_addr}")
            return redirect(url_for("login"))

        try:
            decoded = base64.b64decode(token).decode("utf-8")
            if ":" not in decoded:
                raise ValueError("Invalid token format")
            
            username, password = decoded.split(":", 1)
            credentials = load_credentials()
            
            if username and password and credentials.get(username) == password:
                session.permanent = True
                session["user"] = username
                log_activity(f"quick_login_success | user={username} | ip={request.remote_addr} | next={redirect_to}")
                return redirect(redirect_to)
            else:
                log_activity(f"quick_login_failure | user={username or 'unknown'} | reason=invalid_credentials | ip={request.remote_addr}")
                return redirect(url_for("login"))
        except Exception as exc:
            log_activity(f"quick_login_failure | reason=decode_error | error={str(exc)} | ip={request.remote_addr}")
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

    # Blueprint: disallineamento QR
    app.register_blueprint(disallineamento_qr_bp)

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
