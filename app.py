"""Entry-point Flask.

Nuova nomenclatura pagine:
- /MedicairGeek        -> Home (templates/home.html)
- /reportIncidente     -> Report incidenti (Blueprint: report_incidente_bp)
- /reportIntervento    -> Report intervento (placeholder, templates/reportIntervento.html)

Nota: manteniamo la static_url_path storica /Geekplus/static per compatibilità.
"""

from __future__ import annotations

from datetime import datetime

from flask import Flask, render_template

from report_incidente import report_incidente_bp


def create_app() -> Flask:
    app = Flask(__name__, static_folder="static", static_url_path="/MedicairGeek/static")
    app.config["SECRET_KEY"] = "CHANGE_ME__report_medicair_secret"
    app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300MB

    # Blueprint: report incidente
    app.register_blueprint(report_incidente_bp)

    # Home
    @app.get("/MedicairGeek")
    @app.get("/MedicairGeek/")
    def home():
        # Template atteso: templates/home.html
        return render_template("home.html", title="MedicairGeek", now=datetime.now())

    # Report Intervento (placeholder)
    @app.get("/MedicairGeek/reportIntervento")
    def report_intervento_form():
        # Template atteso: templates/reportIntervento.html
        # Quando lo implementerai, conviene spostare anche questo in un blueprint dedicato.
        return render_template("reportIntervento.html", title="Report Intervento", now=datetime.now())

    return app


app = create_app()


if __name__ == "__main__":
    # Non avvio ensure_report_assets qui: è responsabilità del blueprint quando serve.
    app.run(host="0.0.0.0", port=3570, debug=False)
