from __future__ import annotations

from typing import Any, Dict, List, Optional

from flask import Blueprint, abort, render_template, request

from consulta_report import _read_excel_rows


info_impianto_bp = Blueprint("info_impianto", __name__)

ROBOT_REGISTRY: List[Dict[str, str]] = [
    {"id": "16278", "full_id": "1216278", "ip": "10.1.70.57/24", "label": "Robot 16278"},
    {"id": "16279", "full_id": "1216279", "ip": "10.1.70.55/24", "label": "Robot 16279"},
    {"id": "16292", "full_id": "1216292", "ip": "10.1.70.60/24", "label": "Robot 16292"},
    {"id": "16294", "full_id": "1216294", "ip": "10.1.70.62/24", "label": "Robot 16294"},
    {"id": "16302", "full_id": "1216302", "ip": "10.1.70.53/24", "label": "Robot 16302"},
    {"id": "16306", "full_id": "1216306", "ip": "10.1.70.56/24", "label": "Robot 16306"},
    {"id": "16313", "full_id": "1216313", "ip": "10.1.70.54/24", "label": "Robot 16313"},
    {"id": "16314", "full_id": "1216314", "ip": "10.1.70.51/24", "label": "Robot 16314"},
    {"id": "16325", "full_id": "1216325", "ip": "10.1.70.61/24", "label": "Robot 16325"},
    {"id": "16337", "full_id": "1216337", "ip": "10.1.70.52/24", "label": "Robot 16337"},
    {"id": "16339", "full_id": "1216339", "ip": "10.1.70.63/24", "label": "Robot 16339"},
    {"id": "16340", "full_id": "1216340", "ip": "10.1.70.64/24", "label": "Robot 16340"},
    {"id": "16348", "full_id": "1216348", "ip": "10.1.70.58/24", "label": "Robot 16348"},
    {"id": "16349", "full_id": "1216349", "ip": "10.1.70.50/24", "label": "Robot 16349"},
    {"id": "16350", "full_id": "1216350", "ip": "10.1.70.59/24", "label": "Robot 16350"},
]

HARDWARE_ITEMS: List[Dict[str, str]] = [
    {"label": "Station 1", "slug": "station-1"},
    {"label": "Station 2", "slug": "station-2"},
    {"label": "Station 3", "slug": "station-3"},
    {"label": "Pavimento", "slug": "pavimento"},
    {"label": "Scaffali", "slug": "scaffali"},
    {"label": "Ceste", "slug": "ceste"},
]

SOFTWARE_ITEMS: List[Dict[str, str]] = [
    {"label": "RMS", "slug": "rms"},
    {"label": "WMS", "slug": "wms"},
]


def _normalize_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _truncate(text: str, max_chars: int = 120) -> str:
    clean = " ".join(_normalize_str(text).replace("\r", " ").replace("\n", " ").split())
    if len(clean) <= max_chars:
        return clean
    return clean[: max_chars - 1].rstrip() + "…"


def _pick_first_non_empty(*values: Any) -> str:
    for value in values:
        text = _normalize_str(value)
        if text:
            return text
    return ""


def _clean_ip(value: Any) -> str:
    return _normalize_str(value).split("/", 1)[0]


def _is_yes_value(value: Any) -> bool:
    return _normalize_str(value).lower() in {"si", "sì", "yes", "y", "true", "1"}


def _get_robot(robot_id: str) -> Optional[Dict[str, str]]:
    key = _normalize_str(robot_id)
    for robot in ROBOT_REGISTRY:
        if robot["id"] == key or robot["full_id"] == key:
            return robot
    return None


def _robot_aliases(robot: Dict[str, str]) -> set[str]:
    short_id = robot.get("id", "")
    full_id = robot.get("full_id", "") or (f"12{short_id}" if short_id else "")
    return {
        short_id,
        full_id,
        f"Robot {short_id}",
        f"Robot {full_id}",
    }


def _robot_matches(report_robot_value: Any, robot: Dict[str, str]) -> bool:
    value = _normalize_str(report_robot_value).lower()
    if not value:
        return False

    if "tutti" in value:
        return True

    aliases = {alias.lower() for alias in _robot_aliases(robot) if alias}
    return any(alias in value for alias in aliases)


def _is_global_robot_report(row: Dict[str, Any]) -> bool:
    return "tutti" in _normalize_str(row.get("robot")).lower()


def _build_parti_coinvolte(row: Dict[str, Any]) -> str:
    blocks: List[str] = []

    errore = _normalize_str(row.get("errore"))
    risoluzione = _normalize_str(row.get("risoluzione"))
    zona = _normalize_str(row.get("zona"))
    cella = _normalize_str(row.get("cella"))
    scaffale = _normalize_str(row.get("scaffale"))

    if errore:
        blocks.append(_truncate(errore, 90))
    elif risoluzione:
        blocks.append(_truncate(risoluzione, 90))

    area_parts = [x for x in [zona, cella, scaffale] if x and x.lower() != "senza scaffale"]
    if area_parts:
        blocks.append("Area: " + " / ".join(area_parts))

    if not blocks:
        return "-"
    return " | ".join(blocks[:2])


def _extract_replaced_part(row: Dict[str, Any]) -> str:
    candidates = [
        row.get("risoluzione"),
        row.get("note"),
        row.get("update1"),
        row.get("update2"),
        row.get("errore"),
    ]
    keywords = ("sostit", "cambiat", "ricambio", "rimpiazz", "sostitu")

    for value in candidates:
        text = _normalize_str(value)
        if text and any(keyword in text.lower() for keyword in keywords):
            return _truncate(text, 120)

    categoria = _normalize_str(row.get("categoria")).lower()
    risoluzione = _normalize_str(row.get("risoluzione"))
    if categoria == "intervento manutenzione" and risoluzione:
        return _truncate(risoluzione, 120)

    return ""


def _get_related_rows_for_robot(robot: Dict[str, str], rows: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
    base_rows = rows if rows is not None else _read_excel_rows(limit=5000)
    return [row for row in base_rows if _robot_matches(row.get("robot"), robot)]


def _decorate_robot(robot: Dict[str, str], related_rows: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    related_rows = related_rows if related_rows is not None else _get_related_rows_for_robot(robot)
    status_row = next((row for row in related_rows if not _is_global_robot_report(row)), {})
    is_removed = _is_yes_value(status_row.get("rimosso")) if status_row else False

    decorated: Dict[str, Any] = dict(robot)
    decorated["clean_ip"] = _clean_ip(robot.get("ip", ""))
    decorated["display_id"] = f"{robot.get('id', '')} ({robot.get('full_id', '')})" if robot.get("full_id") else robot.get("id", "")
    decorated["is_removed"] = is_removed
    decorated["status_label"] = "Robot rimosso = SI" if is_removed else "Robot rimosso = NO"
    decorated["latest_report_label"] = _pick_first_non_empty(status_row.get("dt_label"), status_row.get("data"), "")
    return decorated


def _build_robot_list() -> List[Dict[str, Any]]:
    rows = _read_excel_rows(limit=5000)
    return [_decorate_robot(robot, _get_related_rows_for_robot(robot, rows)) for robot in ROBOT_REGISTRY]


def _build_robot_tables(robot: Dict[str, str], related_rows: Optional[List[Dict[str, Any]]] = None) -> Dict[str, List[Dict[str, Any]]]:
    related_rows = related_rows if related_rows is not None else _get_related_rows_for_robot(robot)

    replaced_parts: List[Dict[str, Any]] = []
    manutenzioni: List[Dict[str, Any]] = []
    incidenti: List[Dict[str, Any]] = []

    for row in related_rows:
        report_id = int(row.get("id", 0))
        link = f"/MedicairGeek/storicoReport/{report_id}" if report_id else "#"
        data_label = _pick_first_non_empty(row.get("dt_label"), row.get("data"), "-")
        descrizione = _pick_first_non_empty(row.get("note"), row.get("risoluzione"), row.get("errore"), "-")

        incidenti.append(
            {
                "data": data_label,
                "titolo": _pick_first_non_empty(row.get("titolo"), "-"),
                "tipo": _pick_first_non_empty(row.get("categoria"), "-"),
                "descrizione": _truncate(descrizione, 140) if descrizione != "-" else "-",
                "rimosso": _pick_first_non_empty(row.get("rimosso"), "-"),
                "link": link,
            }
        )

        if _normalize_str(row.get("categoria")).lower() == "intervento manutenzione":
            manutenzioni.append(
                {
                    "data": data_label,
                    "titolo": _pick_first_non_empty(row.get("titolo"), "-"),
                    "parti_coinvolte": _build_parti_coinvolte(row),
                    "descrizione": _truncate(descrizione, 140) if descrizione != "-" else "-",
                    "link": link,
                }
            )

        part_text = _extract_replaced_part(row)
        if part_text:
            replaced_parts.append(
                {
                    "data": data_label,
                    "pezzo": part_text,
                }
            )

    return {
        "replaced_parts": replaced_parts[:50],
        "manutenzioni": manutenzioni[:50],
        "incidenti": incidenti[:100],
        "related_count": len(related_rows),
    }


@info_impianto_bp.get("/MedicairGeek/infoImpianto")
def info_impianto_home():
    active_tab = _normalize_str(request.args.get("tab", "robot")).lower() or "robot"
    if active_tab not in {"robot", "hardware", "software"}:
        active_tab = "robot"

    return render_template(
        "infoImpianto.html",
        title="Info Impianto",
        active_tab=active_tab,
        robots=_build_robot_list(),
        hardware_items=HARDWARE_ITEMS,
        software_items=SOFTWARE_ITEMS,
    )


@info_impianto_bp.get("/MedicairGeek/infoImpianto/robot/<robot_id>")
def info_impianto_robot_detail(robot_id: str):
    robot = _get_robot(robot_id)
    if not robot:
        abort(404)

    related_rows = _get_related_rows_for_robot(robot)
    robot_view = _decorate_robot(robot, related_rows)
    tables = _build_robot_tables(robot, related_rows)

    return render_template(
        "infoImpiantoRobotDetail.html",
        title=f"Dettaglio {robot_view['label']}",
        robot=robot_view,
        replaced_parts=tables["replaced_parts"],
        manutenzioni=tables["manutenzioni"],
        incidenti=tables["incidenti"],
        related_count=tables["related_count"],
    )


@info_impianto_bp.get("/MedicairGeek/infoImpianto/placeholder/<section>/<slug>")
def info_impianto_placeholder(section: str, slug: str):
    label = _normalize_str(request.args.get("label")) or slug.replace("-", " ").title()
    section_label = {
        "hardware": "Hardware",
        "software": "Software",
        "robot": "Robot",
    }.get(section.lower(), "Info Impianto")

    return render_template(
        "infoImpiantoPlaceholder.html",
        title=f"{label} - Pagina in costruzione",
        section_label=section_label,
        item_label=label,
    )
