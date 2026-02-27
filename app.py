import os
import re
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

APP_TITLE = "Report Medicair - Incidenti Robot"
BASE_DIR = Path(__file__).resolve().parent
REPORT_DIR = BASE_DIR / "report"
EXCEL_PATH = REPORT_DIR / "Incidenti_robot.xlsx"

# Upload
ALLOWED_EXT = {
    # immagini
    "jpg", "jpeg", "png", "gif", "webp", "heic",
    # video
    "mp4", "mov", "m4v", "avi", "mkv", "webm"
}
MAX_CONTENT_LENGTH = 300 * 1024 * 1024  # 300MB (regola se vuoi)

ROBOT_LIST = [
    "16278", "16279", "16292", "16294", "16302", "16306", "16314",
    "16325", "16337", "16339", "16340", "16348", "16349", "16350"
]

ZONA_LIST = [
    "Avvio Robot",
    "Corridoio principale",
    "corridoi",
    "Station 1",
    "Station 2",
    "station 3",
    "ingresso Ws1",
    "ingresso WS1/2"
]

LUCI_CAMPO1 = ["Fisse", "Lampeggianti"]
LUCI_CAMPO2 = ["bianca", "blu", "verde", "rossa", "gialla", "altro"]

HEADERS = [
    "id", "data", "ora", "robot", "scaffale", "cella", "zona",
    "luci robot", "errore", "note", "rimosso", "risoluzione"
]


def ensure_report_assets():
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    if not EXCEL_PATH.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Incidenti"
        ws.append(HEADERS)
        wb.save(EXCEL_PATH)


def sanitize_folder_part(s: str) -> str:
    s = s.strip()
    s = re.sub(r"[^\w\-]+", "_", s)
    return s[:80] if s else "NA"


def allowed_file(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXT


def get_next_id() -> int:
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    # cerca l'ultima riga utile guardando la colonna A (id)
    last_id = 0
    for row in range(ws.max_row, 1, -1):
        val = ws.cell(row=row, column=1).value
        if val is not None and str(val).strip() != "":
            try:
                last_id = int(val)
                break
            except Exception:
                pass
    wb.close()
    return last_id + 1


def append_row(row_values: list):
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.append(row_values)
    wb.save(EXCEL_PATH)
    wb.close()


app = Flask(
    __name__,
    static_folder="static",
    static_url_path="/Geekplus/static"
)
app.config["SECRET_KEY"] = "CHANGE_ME__report_medicair_secret"
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.get("/geekplus/reportMedicair")
def report_form():
    now = datetime.now()
    # datetime-local vuole formato: YYYY-MM-DDTHH:MM
    dt_local = now.strftime("%Y-%m-%dT%H:%M")
    return render_template(
        "reportMedicair.html",
        title=APP_TITLE,
        dt_local=dt_local,
        robot_list=ROBOT_LIST,
        zona_list=ZONA_LIST,
        luci_campo1 = LUCI_CAMPO1,
        luci_campo2=LUCI_CAMPO2,
    )


@app.post("/geekplus/reportMedicair")
def submit_report():
    ensure_report_assets()

    # Required (*)
    dt_local = request.form.get("dt_local", "").strip()
    robots = request.form.getlist("robots")  # multi-select
    scaffale = request.form.get("scaffale", "senza scaffale").strip()
    zona = request.form.get("zona", "").strip()
    errore = request.form.get("errore", "").strip()
    descrizione = request.form.get("descrizione", "").strip()
    luci_c1 = request.form.get("luci_c1", "").strip()
    luci_c2 = request.form.get("luci_c2", "").strip()

    # Optional
    cella = request.form.get("cella", "").strip()
    rimosso = request.form.get("rimosso", "no").strip().lower()
    risoluzione = request.form.get("risoluzione", "").strip()

    # Validazioni
    errors = []
    if not dt_local:
        errors.append("Data e ora sono obbligatori.")
    if not robots:
        errors.append("Seleziona almeno un robot.")
    if not zona:
        errors.append("Zona è obbligatoria.")
    if not luci_c2:
        errors.append("Luci robot è obbligatorio.")
    
    if not descrizione:
        errors.append("Descrizione è obbligatoria.")
    
    # Robots validi
    invalid = [r for r in robots if r not in ROBOT_LIST]
    if invalid:
        errors.append(f"Robot non validi: {', '.join(invalid)}")

    if zona and zona not in ZONA_LIST:
        errors.append("Zona non valida.")
    if luci_c2 and luci_c2 not in LUCI_CAMPO2:
        errors.append("Luci campo 2 non valide.")
    if rimosso not in ("si", "no"):
        rimosso = "no"

    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("report_form"))

    # Parsing data/ora per Excel
    # dt_local: "YYYY-MM-DDTHH:MM"
    try:
        dt = datetime.strptime(dt_local, "%Y-%m-%dT%H:%M")
    except Exception:
        flash("Formato data/ora non valido.", "error")
        return redirect(url_for("report_form"))

    next_id = get_next_id()

    # Cartella: id_giorno_ora
    folder_stamp = dt.strftime("%d-%m-%Y_%H-%M")
    folder_name = f"{next_id}_{folder_stamp}"
    folder_path = REPORT_DIR / folder_name
    folder_path.mkdir(parents=True, exist_ok=True)

    # Upload file (può essere 0..N)
    saved_files = []
    files = request.files.getlist("media")
    for f in files:
        if not f or not f.filename:
            continue
        if not allowed_file(f.filename):
            continue
        safe = secure_filename(f.filename)
        if not safe:
            continue
        dest = folder_path / safe
        # evita overwrite
        i = 1
        while dest.exists():
            stem = dest.stem
            suffix = dest.suffix
            dest = folder_path / f"{stem}_{i}{suffix}"
            i += 1
        f.save(dest)
        saved_files.append(dest.name)

    # Luci robot: campo1 fisso "lampeggiante" + campo2 selezionato
    luci_robot = f"{luci_c1}-{luci_c2}".strip()

    row = [
        next_id,
        dt.strftime("%d/%m/%Y"),
        dt.strftime("%H:%M"),
        ", ".join(robots),
        scaffale if scaffale else "senza scaffale",
        cella,
        zona,
        luci_robot,
        errore,
        descrizione,
        rimosso,
        risoluzione
    ]
    append_row(row)

    return render_template(
        "success.html",
        title=APP_TITLE,
        report_id=next_id,
        folder_name=folder_name,
        saved_files=saved_files,
    )


if __name__ == "__main__":
    ensure_report_assets()
    # Se hai reverse proxy/rewriting a monte, spesso basta ascoltare su 0.0.0.0
    app.run(host="0.0.0.0", port=3570, debug=False)