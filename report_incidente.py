from __future__ import annotations

import shutil
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

from flask import Blueprint, jsonify, render_template, request, abort
from openpyxl import Workbook, load_workbook
from werkzeug.utils import secure_filename

from config import DESTINATARI, REPORT_BASE_DIR
from email_sender import EmailSender
from modelli_latex import cleanup_latex_tmp, crea_report


# =====================
# Config / Costanti
# =====================

APP_TITLE = "Report Medicair - Incidenti Robot"
BASE_DIR = Path(__file__).resolve().parent

REPORT_DIR = REPORT_BASE_DIR
EXCEL_PATH = REPORT_DIR / "Incidenti_robot.xlsx"

ALLOWED_EXT = {
    "jpg", "jpeg", "png", "gif", "webp", "heic", "bmp", "tiff",
    "mp4", "mov", "m4v", "avi", "mkv", "webm",
    "pdf", "doc", "docx", "xls", "xlsx",
}

MAX_TITOLO_LEN = 30

CATEGORIE_LIST = [
    "Incidente",
    "Problema Software",
    "Problema Hardware",
    "Altro",
]

ROBOT_LIST = [
    "1216278", "1216279", "1216292", "1216294", "1216302", "1216306", "1216313", "1216314",
    "1216325", "1216337", "1216339", "1216340", "1216348", "1216349", "1216350",
    "Tutti", "Pavimento", "WorkingStation1", "WorkingStation2", "WorkingStation3",
    "ChargingStation", "Altro", "Scaffale",
]

ZONA_LIST = [
    "Avvio Robot",
    "Corridoio principale",
    "Corridoi",
    "Station 1",
    "Station 2",
    "station 3",
    "Manutenzione",
    "Altro",
]

LUCI_CAMPO1 = ["Fisse", "Lampeggianti", "Non applicabile"]
LUCI_CAMPO2 = ["Bianca", "Rossa", "Verde", "Blu", "Gialla", "Viola", "Altro"]

HEADERS = [
    "id",
    "data",
    "ora",
    "Categoria",
    "Titolo",
    "robot",
    "scaffale",
    "cella",
    "zona",
    "luci robot",
    "errore",
    "note",
    "rimosso",
    "risoluzione",
]


# =====================
# Job store (in-memory)
# =====================

JOBS: Dict[str, Dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()
JOB_TTL_SECONDS = 6 * 60 * 60  # 6 ore


def _job_set(job_id: str, **kwargs) -> None:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job.update(kwargs)


def _job_get(job_id: str) -> Optional[Dict[str, Any]]:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        return dict(job) if job else None


def _jobs_gc() -> None:
    now = time.time()
    to_delete = []

    with JOBS_LOCK:
        for jid, job in JOBS.items():
            created = job.get("created_ts", now)
            if now - created > JOB_TTL_SECONDS:
                to_delete.append(jid)

        for jid in to_delete:
            del JOBS[jid]


# =====================
# Helpers
# =====================

def ensure_report_assets() -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    if not EXCEL_PATH.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Incidenti"
        ws.append(HEADERS)
        wb.save(EXCEL_PATH)
        wb.close()


def allowed_file(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXT


def normalize_spaces(value: str) -> str:
    return " ".join((value or "").strip().split())


def sanitize_titolo(value: str) -> str:
    return normalize_spaces(value)[:MAX_TITOLO_LEN]


def get_next_id() -> int:
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

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


def append_row(row_values: list) -> None:
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.append(row_values)
    wb.save(EXCEL_PATH)
    wb.close()


def _send_report_email(
    report_id: int,
    dt: datetime,
    titolo: str,
    categoria: str,
    robots: list[str],
    note: str,
    destinatari: list[str],
    allegati: list[Path],
) -> dict:
    if not destinatari:
        return {"ok": True, "sent": 0, "errors": []}

    sender = EmailSender()
    template_name = "REPORT INCIDENTE"

    fields = {
        "data": dt.strftime("%d/%m/%Y %H:%M"),
        "robots": ", ".join(robots),
        "note": (note or "").strip(),
        "titolo": titolo,
        "categoria": categoria,
        "report_id": str(report_id),
    }

    attachments = [str(p) for p in (allegati or []) if p and p.exists() and p.is_file()]

    sent = 0
    errors: list[str] = []

    for dest in destinatari:
        try:
            res = sender.send_template(dest, template_name, fields, attachments)
            if res.ok:
                sent += 1
            else:
                errors.append(f"{dest}: {res.error}")
        except Exception as e:
            errors.append(f"{dest}: {e}")

    return {
        "ok": sent > 0 and not errors,
        "sent": sent,
        "errors": errors,
    }


# =====================
# Worker
# =====================

def _run_job(job_id: str, payload: Dict[str, Any]) -> None:
    """
    Worker che esegue le fasi pesanti:
    - Excel
    - PDF
    - Email
    - Cleanup latex/tmp
    """
    try:
        ensure_report_assets()

        dt: datetime = payload["dt"]
        dt_local: str = payload["dt_local"]
        titolo: str = payload["titolo"]
        categoria: str = payload["categoria"]
        robots: list[str] = payload["robots"]
        scaffale: str = payload["scaffale"]
        cella: str = payload["cella"]
        zona: str = payload["zona"]
        errore: str = payload["errore"]
        descrizione: str = payload["descrizione"]
        luci_robot: str = payload["luci_robot"]
        rimosso: str = payload["rimosso"]
        risoluzione: str = payload["risoluzione"]

        next_id: int = payload["next_id"]
        folder_name: str = payload["folder_name"]
        folder_path: Path = payload["folder_path"]
        saved_file_paths: list[Path] = payload["saved_file_paths"]

        saved_files = [p.name for p in saved_file_paths]

        # ---- EXCEL
        _job_set(job_id, phase="EXCEL", percent=45, message="Scrittura registro Excel…")

        row = [
            next_id,
            dt.strftime("%d/%m/%Y"),
            dt.strftime("%H:%M"),
            categoria,
            titolo,
            ", ".join(robots),
            scaffale if scaffale else "senza scaffale",
            cella,
            zona,
            luci_robot,
            errore,
            descrizione,
            rimosso,
            risoluzione,
        ]
        append_row(row)

        # ---- PDF
        pdf_report_path: Optional[Path] = None
        _job_set(job_id, phase="PDF", percent=62, message="Generazione PDF (LaTeX)…")

        try:
            allegati_list_tex = "\n".join([rf"\item {p.name}" for p in saved_file_paths])

            campi_report = {
                "ReportID": str(next_id),
                "DataIncidente": dt.strftime("%d/%m/%Y"),
                "OraIncidente": dt.strftime("%H:%M"),
                "Titolo": titolo,
                "Categoria": categoria,
                "Robot": ", ".join(robots),
                "Scaffale": scaffale if scaffale else "senza scaffale",
                "Cella": cella,
                "Zona": zona,
                "LuciRobot": luci_robot,
                "Errore": errore,
                "Rimosso": rimosso,
                "Risoluzione": risoluzione,
                "Descrizione": descrizione,
                "AllegatiList": allegati_list_tex,
            }

            nome_pdf = f"REPORT_{next_id}_{dt.strftime('%d-%m-%Y_%H-%M')}"
            res = crea_report("modello", campi_report, nome_file=nome_pdf)

            pdf_report_path = folder_path / res.pdf_path.name
            shutil.copy2(res.pdf_path, pdf_report_path)

            saved_files.append(pdf_report_path.name)
            saved_file_paths.append(pdf_report_path)

        except Exception as e:
            _job_set(job_id, message=f"PDF non generato: {e}")

        # ---- EMAIL
        _job_set(job_id, phase="EMAIL", percent=80, message="Invio email…")

        try:
            email_res = _send_report_email(
                report_id=next_id,
                dt=dt,
                titolo=titolo,
                categoria=categoria,
                robots=robots,
                note=descrizione,
                destinatari=list(DESTINATARI),
                allegati=list(saved_file_paths),
            )
        except Exception as e:
            email_res = {"ok": False, "sent": 0, "errors": [str(e)]}

        # ---- CLEANUP
        _job_set(job_id, phase="CLEANUP", percent=92, message="Pulizia file temporanei…")
        try:
            cleanup_latex_tmp()
        except Exception:
            pass

        # ---- DONE
        _job_set(
            job_id,
            phase="DONE",
            percent=100,
            message="Completato.",
            done=True,
            result={
                "report_id": next_id,
                "folder_name": folder_name,
                "saved_files": saved_files,
                "pdf_created": bool(pdf_report_path and pdf_report_path.exists()),
                "email": email_res,
                "dt_local": dt_local,
                "titolo": titolo,
                "categoria": categoria,
                "robots": robots,
            },
        )

    except Exception as e:
        _job_set(
            job_id,
            phase="ERROR",
            done=True,
            error=str(e),
            message="Errore durante l'elaborazione.",
            percent=100,
        )


# =====================
# Blueprint
# =====================

report_incidente_bp = Blueprint("report_incidente", __name__)


@report_incidente_bp.get("/MedicairGeek/reportIncidente")
def report_form():
    _jobs_gc()
    now = datetime.now()
    dt_local = now.strftime("%Y-%m-%dT%H:%M")

    return render_template(
        "reportIncidente.html",
        title=APP_TITLE,
        dt_local=dt_local,
        robot_list=ROBOT_LIST,
        zona_list=ZONA_LIST,
        luci_campo1=LUCI_CAMPO1,
        luci_campo2=LUCI_CAMPO2,
        categorie_list=CATEGORIE_LIST,
        max_titolo_len=MAX_TITOLO_LEN,
    )


@report_incidente_bp.post("/MedicairGeek/reportIncidente/start")
def start_job():
    """
    Riceve form + file (multipart), salva gli upload nella cartella report,
    crea job e avvia thread worker.
    Gli allegati sono facoltativi.
    """
    _jobs_gc()
    ensure_report_assets()

    # ---- parse form
    dt_local = normalize_spaces(request.form.get("dt_local", ""))
    titolo_raw = request.form.get("titolo", "") or ""
    titolo = sanitize_titolo(titolo_raw)
    categoria = normalize_spaces(request.form.get("categoria", ""))
    robots = request.form.getlist("robots")
    scaffale = normalize_spaces(request.form.get("scaffale", "senza scaffale"))
    zona = normalize_spaces(request.form.get("zona", ""))
    errore = normalize_spaces(request.form.get("errore", ""))
    descrizione = (request.form.get("descrizione", "") or "").strip()
    luci_c1 = normalize_spaces(request.form.get("luci_c1", ""))
    luci_c2 = normalize_spaces(request.form.get("luci_c2", ""))
    cella = normalize_spaces(request.form.get("cella", ""))
    rimosso = normalize_spaces(request.form.get("rimosso", "no")).lower()
    risoluzione = normalize_spaces(request.form.get("risoluzione", ""))

    # Allegati facoltativi, ma vanno comunque letti
    files = request.files.getlist("media")

    # ---- validation
    errors: list[str] = []

    if not dt_local:
        errors.append("Data e ora sono obbligatori.")

    if not titolo_raw.strip():
        errors.append("Titolo è obbligatorio.")
    elif len(normalize_spaces(titolo_raw)) > MAX_TITOLO_LEN:
        errors.append(f"Titolo troppo lungo: massimo {MAX_TITOLO_LEN} caratteri.")

    if not categoria:
        errors.append("Categoria è obbligatoria.")
    elif categoria not in CATEGORIE_LIST:
        errors.append("Categoria non valida.")

    if not robots:
        errors.append("Seleziona almeno un robot.")

    if not zona:
        errors.append("Zona è obbligatoria.")

    if not luci_c1 or not luci_c2:
        errors.append("Luci robot è obbligatorio (seleziona entrambi).")

    if not descrizione.strip():
        errors.append("Descrizione è obbligatoria.")

    invalid = [r for r in robots if r not in ROBOT_LIST]
    if invalid:
        errors.append(f"Robot non validi: {', '.join(invalid)}")

    if zona and zona not in ZONA_LIST:
        errors.append("Zona non valida.")

    if luci_c1 and luci_c1 not in LUCI_CAMPO1:
        errors.append("Luci campo 1 non valide.")

    if luci_c2 and luci_c2 not in LUCI_CAMPO2:
        errors.append("Luci campo 2 non valide.")

    if rimosso not in ("si", "no"):
        rimosso = "no"

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    # ---- dt parse
    try:
        dt = datetime.strptime(dt_local, "%Y-%m-%dT%H:%M")
    except Exception:
        return jsonify({"ok": False, "errors": ["Formato data/ora non valido."]}), 400

    # ---- create IDs + folder
    next_id = get_next_id()
    folder_stamp = dt.strftime("%d-%m-%Y_%H-%M")
    folder_name = f"{next_id}_{folder_stamp}"
    folder_path = REPORT_DIR / folder_name
    folder_path.mkdir(parents=True, exist_ok=True)

    # ---- create job
    job_id = uuid.uuid4().hex
    with JOBS_LOCK:
        JOBS[job_id] = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "created_ts": time.time(),
            "phase": "UPLOAD_SAVING",
            "percent": 10,
            "message": "Salvataggio allegati…",
            "done": False,
            "error": None,
            "result": None,
        }

    # ---- save uploads (facoltativi)
    saved_file_paths: list[Path] = []

    for f in files:
        if not f or not f.filename:
            continue
        if not allowed_file(f.filename):
            continue

        safe = secure_filename(f.filename)
        if not safe:
            continue

        dest = folder_path / safe
        i = 1
        while dest.exists():
            dest = folder_path / f"{dest.stem}_{i}{dest.suffix}"
            i += 1

        f.save(dest)
        saved_file_paths.append(dest)

    luci_robot = f"{luci_c1} - {luci_c2}".strip()

    # ---- prepare payload for worker
    payload = {
        "dt": dt,
        "dt_local": dt_local,
        "titolo": titolo,
        "categoria": categoria,
        "robots": robots,
        "scaffale": scaffale,
        "cella": cella,
        "zona": zona,
        "errore": errore,
        "descrizione": descrizione.strip(),
        "luci_robot": luci_robot,
        "rimosso": rimosso,
        "risoluzione": risoluzione,
        "next_id": next_id,
        "folder_name": folder_name,
        "folder_path": folder_path,
        "saved_file_paths": saved_file_paths,
    }

    _job_set(
        job_id,
        phase="UPLOAD_SAVED",
        percent=35,
        message="Upload completato. Avvio elaborazione…",
    )

    # ---- start worker thread
    t = threading.Thread(target=_run_job, args=(job_id, payload), daemon=True)
    t.start()

    return jsonify({"ok": True, "job_id": job_id}), 200


@report_incidente_bp.get("/MedicairGeek/reportIncidente/status/<job_id>")
def job_status(job_id: str):
    _jobs_gc()
    job = _job_get(job_id)
    if not job:
        return jsonify({"ok": False, "error": "Job non trovato o scaduto."}), 404
    return jsonify({"ok": True, "job": job}), 200


@report_incidente_bp.get("/MedicairGeek/reportIncidente/success/<job_id>")
def job_success(job_id: str):
    _jobs_gc()
    job = _job_get(job_id)
    if not job:
        abort(404)

    if not job.get("done") or job.get("phase") != "DONE" or not job.get("result"):
        abort(404)

    r = job["result"]
    return render_template(
        "send_report_success.html",
        title=APP_TITLE,
        report_id=r.get("report_id"),
        folder_name=r.get("folder_name"),
        saved_files=r.get("saved_files", []),
        destinatari=r.get("destinatari", []),
    )