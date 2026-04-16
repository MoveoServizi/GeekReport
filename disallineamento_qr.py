from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List, Optional

from flask import Blueprint, jsonify, render_template, request

from consulta_report import _read_excel_rows
from log_utils import log_activity


disallineamento_qr_bp = Blueprint("disallineamento_qr", __name__)

# Registro dei robot (short id usato anche nelle route info impianto)
ROBOT_IDS = [
    "16278", "16279", "16292", "16294", "16302", "16306",
    "16313", "16314", "16325", "16337", "16339", "16340", "16348", "16349", "16350",
]

# Tipi evento
EVENT_TYPE_QR = "qr"
EVENT_TYPE_FW = "fw"
EVENT_TYPE_MCU = "mcu"
EVENT_TYPE_BAT = "bat"
EVENT_TYPE_MAN = "man"
EVENT_TYPE_MANSTRA = "manstra"

EVENT_TYPE_ORDER = [
    EVENT_TYPE_QR,
    EVENT_TYPE_FW,
    EVENT_TYPE_MCU,
    EVENT_TYPE_BAT,
    EVENT_TYPE_MAN,
    EVENT_TYPE_MANSTRA,
]

EVENT_LABEL_PREFIX = {
    EVENT_TYPE_QR: "QR",
    EVENT_TYPE_FW: "FW",
    EVENT_TYPE_MCU: "MCU",
    EVENT_TYPE_BAT: "BAT",
    EVENT_TYPE_MAN: "MAN",
    EVENT_TYPE_MANSTRA: "STR",
}


def _normalize_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_date(date_str: str) -> Optional[datetime]:
    raw = _normalize_str(date_str)
    if not raw:
        return None

    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt)
        except Exception:
            continue
    return None


def _format_date_short(date_str: str) -> str:
    dt = _parse_date(date_str)
    if not dt:
        return _normalize_str(date_str)
    return dt.strftime("%d/%m/%y")


def _build_text_blob(*values: Any) -> str:
    return " ".join(_normalize_str(v).lower() for v in values if _normalize_str(v))


def _contains_any(text: str, keywords: List[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def _extract_robot_ids(robot_value: Any) -> List[str]:
    text = _normalize_str(robot_value)
    if not text:
        return []

    found: List[str] = []
    seen: set[str] = set()
    for robot_id in ROBOT_IDS:
        if robot_id in text and robot_id not in seen:
            found.append(robot_id)
            seen.add(robot_id)
    return found


def _categorize_event(row: Dict[str, Any]) -> Optional[str]:
    categoria = _normalize_str(row.get("categoria")).lower()
    titolo = row.get("titolo")
    parti_coinvolte = row.get("parti_coinvolte")
    note = row.get("note")
    risoluzione = row.get("risoluzione")
    errore = row.get("errore")

    searchable = _build_text_blob(titolo, parti_coinvolte, note, risoluzione, errore)

    if "dissalineato" in categoria and "qr" in categoria:
        return EVENT_TYPE_QR

    is_manutenzione_stra = "intervento manutenzione straordinaria" in categoria
    is_manutenzione = is_manutenzione_stra or "intervento manutenzione" in categoria
    if not is_manutenzione:
        return None

    if _contains_any(searchable, ["batteria", "battery"]):
        return EVENT_TYPE_BAT

    if _contains_any(searchable, ["firmware", "firm ", "firm.", "fw ", " fw", "update fw"]):
        return EVENT_TYPE_FW

    if _contains_any(searchable, ["scheda madre", "mainboard", "motherboard", "mcu", "controller board", "control board"]):
        return EVENT_TYPE_MCU

    return EVENT_TYPE_MANSTRA if is_manutenzione_stra else EVENT_TYPE_MAN


def _read_events_from_excel() -> List[Dict[str, Any]]:
    events: List[Dict[str, Any]] = []

    try:
        for row in _read_excel_rows(limit=5000):
            date_str = _normalize_str(row.get("data"))
            date_obj = _parse_date(date_str)
            if not date_obj:
                continue

            robots = _extract_robot_ids(row.get("robot"))
            if not robots:
                continue

            event_type = _categorize_event(row)
            if not event_type:
                continue

            report_id_raw = row.get("id")
            try:
                report_id = int(report_id_raw)
            except Exception:
                report_id = None

            titolo = _normalize_str(row.get("titolo")) or _normalize_str(row.get("categoria")) or "Evento"

            events.append(
                {
                    "report_id": report_id,
                    "date": date_obj,
                    "date_str": date_str,
                    "date_short": _format_date_short(date_str),
                    "robots": robots,
                    "event_type": event_type,
                    "titolo": titolo,
                    "categoria": _normalize_str(row.get("categoria")),
                    "report_url": f"/MedicairGeek/storicoReport/{report_id}" if report_id is not None else "",
                }
            )
    except Exception as exc:
        print(f"[DisallineamentoQR] Errore lettura Excel: {exc}")

    return events


def _build_robot_events_table() -> tuple[
    List[str],
    Dict[str, List[Dict[str, Any]]],
    List[Dict[str, str]],
    Dict[str, Dict[str, int]],
]:
    events = _read_events_from_excel()

    robot_data: Dict[str, List[Dict[str, Any]]] = {}
    all_dates: Dict[str, str] = {}

    for event in events:
        all_dates[event["date_str"]] = event["date_short"]

        for robot_id in event["robots"]:
            robot_events = robot_data.setdefault(robot_id, [])
            robot_events.append(
                {
                    "report_id": event["report_id"],
                    "report_url": event["report_url"],
                    "date_str": event["date_str"],
                    "date_short": event["date_short"],
                    "date": event["date"],
                    "event_type": event["event_type"],
                    "titolo": event["titolo"],
                    "categoria": event["categoria"],
                }
            )

    robots_list = [robot_id for robot_id in ROBOT_IDS if robot_id in robot_data]

    unique_dates = [
        {"key": date_key, "label": all_dates[date_key]}
        for date_key in sorted(all_dates.keys(), key=lambda value: _parse_date(value) or datetime.max)
    ]

    robot_summary: Dict[str, Dict[str, int]] = {}

    for robot_id, robot_events in robot_data.items():
        robot_events.sort(
            key=lambda event: (
                event["date"],
                event.get("report_id") or 0,
                EVENT_TYPE_ORDER.index(event["event_type"]) if event["event_type"] in EVENT_TYPE_ORDER else 999,
            )
        )

        counters = {event_type: 0 for event_type in EVENT_TYPE_ORDER}
        for event in robot_events:
            counters[event["event_type"]] += 1
            prefix = EVENT_LABEL_PREFIX.get(event["event_type"], event["event_type"].upper())
            event["label"] = f"{prefix}{counters[event['event_type']]}"
            event["robot_url"] = f"/MedicairGeek/infoImpianto/robot/{robot_id}"

        robot_summary[robot_id] = dict(counters)

    return robots_list, robot_data, unique_dates, robot_summary


@disallineamento_qr_bp.get("/MedicairGeek/disallineamentoQr")
@disallineamento_qr_bp.get("/disallineamento-qr")
def disallineamento_qr_page():
    robots_list, robot_data, unique_dates, robot_summary = _build_robot_events_table()
    log_activity(f"view | page=disallineamentoQr | ip={request.remote_addr}")

    return render_template(
        "disallineamentoQr.html",
        title="Disallineamento QR",
        now=datetime.now(),
        robots_list=robots_list,
        robot_data=robot_data,
        unique_dates=unique_dates,
        robot_summary=robot_summary,
    )


@disallineamento_qr_bp.get("/MedicairGeek/DisallineamentoQR/data")
def api_disallineamento_qr_data():
    robots_list, robot_data, unique_dates, robot_summary = _build_robot_events_table()

    out: Dict[str, List[Dict[str, Any]]] = {}
    for robot_id in robots_list:
        out[robot_id] = []
        for event in robot_data.get(robot_id, []):
            out[robot_id].append(
                {
                    "report_id": event.get("report_id"),
                    "report_url": event.get("report_url", ""),
                    "date": event["date_str"],
                    "date_short": event["date_short"],
                    "event_type": event["event_type"],
                    "label": event.get("label", ""),
                    "titolo": event["titolo"],
                    "categoria": event.get("categoria", ""),
                }
            )

    return jsonify(
        {
            "ok": True,
            "robots": robots_list,
            "unique_dates": unique_dates,
            "summary": robot_summary,
            "data": out,
        }
    )
