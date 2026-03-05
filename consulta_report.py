# consulta_report.py
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from flask import Blueprint, jsonify, render_template, request, abort, send_from_directory
from openpyxl import load_workbook


consulta_report_bp = Blueprint("consulta_report", __name__)

BASE_DIR = Path(__file__).resolve().parent
REPORT_DIR = BASE_DIR / "report"
EXCEL_PATH = REPORT_DIR / "Incidenti_robot.xlsx"

# Estensioni media (coerenti con report_incidente.py)
IMG_EXT = {"jpg", "jpeg", "png", "gif", "webp", "heic"}
VID_EXT = {"mp4", "mov", "m4v", "avi", "mkv", "webm"}


def _safe_exists_excel() -> bool:
    return EXCEL_PATH.exists() and EXCEL_PATH.is_file()


def _normalize_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _truncate_2lines(text: str, max_chars: int = 140) -> str:
    t = (text or "").strip().replace("\r", "\n")
    # “2 righe” lato UI è più affidabile via CSS, qui mettiamo solo un limite sensato per API.
    if len(t) <= max_chars:
        return t
    return t[: max_chars - 1].rstrip() + "…"


def _find_folder_for_report_id(report_id: int) -> Optional[str]:
    """
    Il report_id NON è salvato nell’Excel come folder_name.
    Convenzione cartelle: report/<id>_<dd-mm-YYYY_HH-MM>
    Quindi cerchiamo la cartella che inizia con "{id}_".
    Se ce ne sono più di una (rarissimo), scegliamo la più recente (mtime).
    """
    if not REPORT_DIR.exists():
        return None

    prefix = f"{report_id}_"
    candidates: List[Path] = []
    for p in REPORT_DIR.iterdir():
        if p.is_dir() and p.name.startswith(prefix):
            candidates.append(p)

    if not candidates:
        return None

    candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return candidates[0].name


def _list_attachments(folder_name: str) -> List[Dict[str, Any]]:
    folder = REPORT_DIR / folder_name
    if not folder.exists() or not folder.is_dir():
        return []

    out: List[Dict[str, Any]] = []
    for p in sorted(folder.iterdir(), key=lambda x: x.name.lower()):
        if not p.is_file():
            continue

        ext = p.suffix.lower().lstrip(".")
        kind = "file"
        if ext in IMG_EXT:
            kind = "image"
        elif ext in VID_EXT:
            kind = "video"
        elif ext == "pdf":
            kind = "pdf"

        out.append(
            {
                "name": p.name,
                "kind": kind,
                "url": f"/MedicairGeek/ConsultaReport/media/{folder_name}/{p.name}",
            }
        )
    return out


def _read_excel_rows(limit: int = 5000) -> List[Dict[str, Any]]:
    """
    Legge l’Excel "Incidenti_robot.xlsx".
    Header previsto (report_incidente.py): id,data,ora,robot,scaffale,cella,zona,luci robot,errore,note,rimosso,risoluzione
    :contentReference[oaicite:1]{index=1}
    """
    if not _safe_exists_excel():
        return []

    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    rows: List[Dict[str, Any]] = []
    # riga 1 = header
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not r or r[0] is None:
            continue

        try:
            rid = int(str(r[0]).strip())
        except Exception:
            continue

        data = _normalize_str(r[1])
        ora = _normalize_str(r[2])
        robot = _normalize_str(r[3])
        scaffale = _normalize_str(r[4])
        cella = _normalize_str(r[5])
        zona = _normalize_str(r[6])
        luci = _normalize_str(r[7])
        errore = _normalize_str(r[8])
        note = _normalize_str(r[9])  # in realtà è la descrizione lunga

        rows.append(
            {
                "id": rid,
                "data": data,
                "ora": ora,
                "dt_label": f"{data} {ora}".strip(),
                "robot": robot,
                "scaffale": scaffale,
                "cella": cella,
                "zona": zona,
                "luci": luci,
                "errore": errore,
                "note": note,
                "note_preview": _truncate_2lines(note),
            }
        )

        if len(rows) >= limit:
            break

    wb.close()

    # Ordine: più recenti in alto (se data/ora non parseabili, fallback su id)
    def _sort_key(x: Dict[str, Any]) -> Tuple[int, str]:
        # Provo a creare una data “ordinabile” dd/mm/YYYY + HH:MM
        try:
            dt = datetime.strptime(f"{x.get('data','')} {x.get('ora','')}", "%d/%m/%Y %H:%M")
            return (0, dt.isoformat())
        except Exception:
            return (1, f"{x.get('id',0):010d}")

    rows.sort(key=_sort_key, reverse=True)
    return rows


@consulta_report_bp.get("/MedicairGeek/storicoReport")
def storico_report_page():
    return render_template("storicoReport.html", title="Storico Report")


@consulta_report_bp.get("/MedicairGeek/ConsultaReport/list")
def api_list_reports():
    """
    Query params:
    - q: ricerca testo (robot / note / zona / errore)
    - robot: filtro robot (substring sul campo robot che può contenere più robot)
    - limit: default 300
    """
    q = (request.args.get("q", "") or "").strip().lower()
    robot_filter = (request.args.get("robot", "") or "").strip().lower()
    try:
        limit = int(request.args.get("limit", "300"))
    except Exception:
        limit = 300
    limit = max(1, min(2000, limit))

    rows = _read_excel_rows(limit=5000)

    if q:
        def match_q(x: Dict[str, Any]) -> bool:
            blob = " ".join(
                [
                    _normalize_str(x.get("robot")),
                    _normalize_str(x.get("note")),
                    _normalize_str(x.get("zona")),
                    _normalize_str(x.get("errore")),
                    _normalize_str(x.get("scaffale")),
                    _normalize_str(x.get("cella")),
                ]
            ).lower()
            return q in blob

        rows = [r for r in rows if match_q(r)]

    if robot_filter:
        rows = [r for r in rows if robot_filter in _normalize_str(r.get("robot")).lower()]

    # aggiungo folder (se esiste) per indicare se ha allegati consultabili
    out: List[Dict[str, Any]] = []
    for r in rows[:limit]:
        folder = _find_folder_for_report_id(int(r["id"]))
        out.append(
            {
                "id": r["id"],
                "dt_label": r["dt_label"],
                "robot": r["robot"],
                "note_preview": r["note_preview"],
                "has_folder": bool(folder),
                "folder_name": folder or "",
            }
        )

    return jsonify({"ok": True, "reports": out})


@consulta_report_bp.get("/MedicairGeek/ConsultaReport/report/<int:report_id>")
def api_get_report(report_id: int):
    rows = _read_excel_rows(limit=5000)
    item = next((r for r in rows if int(r.get("id", -1)) == int(report_id)), None)
    if not item:
        return jsonify({"ok": False, "error": "Report non trovato."}), 404

    folder = _find_folder_for_report_id(report_id)
    attachments = _list_attachments(folder) if folder else []

    return jsonify(
        {
            "ok": True,
            "report": {
                "id": item["id"],
                "data": item["data"],
                "ora": item["ora"],
                "dt_label": item["dt_label"],
                "robot": item["robot"],
                "scaffale": item["scaffale"],
                "cella": item["cella"],
                "zona": item["zona"],
                "luci": item["luci"],
                "errore": item["errore"],
                "note": item["note"],
                "folder_name": folder or "",
                "attachments": attachments,
            },
        }
    )


@consulta_report_bp.get("/MedicairGeek/ConsultaReport/media/<path:folder_name>/<path:filename>")
def serve_report_media(folder_name: str, filename: str):
    """
    Serve in modo sicuro i file dentro report/<folder_name>/<filename>.
    Blocca path traversal.
    """
    base = REPORT_DIR.resolve()
    folder = (REPORT_DIR / folder_name).resolve()
    if base not in folder.parents and folder != base:
        abort(404)

    file_path = (folder / filename).resolve()
    if folder not in file_path.parents:
        abort(404)

    if not file_path.exists() or not file_path.is_file():
        abort(404)

    return send_from_directory(directory=str(folder), path=filename, as_attachment=False)

# dentro consulta_report.py

@consulta_report_bp.get("/MedicairGeek/storicoReport/<int:report_id>")
def storico_report_detail_page(report_id: int):
    # pagina dedicata (JS chiama l'API e renderizza)
    return render_template("reportDetail.html", title=f"Dettaglio Report #{report_id}", report_id=report_id)