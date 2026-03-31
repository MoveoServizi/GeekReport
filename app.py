# app.py
"""Entry-point Flask.

Nuova nomenclatura pagine:
- /MedicairGeek                -> Home (templates/home.html)
- /MedicairGeek/reportIncidente -> Report incidenti (Blueprint: report_incidente_bp)
- /MedicairGeek/reportIntervento -> Report intervento (placeholder, templates/reportIntervento.html)
- /MedicairGeek/storicoReport  -> Storico Report (Blueprint: consulta_report_bp)
"""

from __future__ import annotations

from datetime import datetime

from flask import Flask, render_template

from report_incidente import report_incidente_bp
from consulta_report import consulta_report_bp
from info_impianto import ensure_info_impianto_cache, info_impianto_bp


def create_app() -> Flask:
    app = Flask(__name__, static_folder="static", static_url_path="/MedicairGeek/static")
    app.config["SECRET_KEY"] = "CHANGE_ME__report_medicair_secret"
    app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300MB

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