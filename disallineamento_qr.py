"""Blueprint per la pagina Disallineamento QR.

Mostra una tabella con:
- Colonna 1: ID Robot
- Per ogni robot: date importante categorizzate per tipo evento
  1) Incidenti categoria disallineamento Qr
  2) Update firmware
  3) Sostituzione scheda madre
  4) Sostituzione batteria / motore sollevamento
"""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from flask import Blueprint, render_template, request
from openpyxl import load_workbook

from config import REPORT_BASE_DIR
from log_utils import log_activity


disallineamento_qr_bp = Blueprint("disallineamento_qr", __name__)

REPORT_DIR = REPORT_BASE_DIR
EXCEL_PATH = REPORT_DIR / "Incidenti_robot.xlsx"

# Registro dei robot
ROBOT_IDS = ["16278", "16279", "16292", "16294", "16302", "16306", 
             "16313", "16314", "16325", "16337", "16339", "16340", "16348", "16349", "16350"]


def _normalize_str(v: Any) -> str:
    """Normalizza un valore a stringa."""
    if v is None:
        return ""
    return str(v).strip()


def _safe_exists_excel() -> bool:
    """Controlla se il file Excel esiste."""
    return EXCEL_PATH.exists() and EXCEL_PATH.is_file()


def _get_header_map(ws) -> Dict[str, int]:
    """Mappa nome_colonna -> indice colonna (0-based)."""
    headers: Dict[str, int] = {}
    for idx, cell in enumerate(ws[1]):
        key = _normalize_str(cell.value)
        if key:
            headers[key] = idx
    return headers


def _cell_from_row(row: tuple, header_map: Dict[str, int], key: str) -> Any:
    """Estrae un valore da una riga usando la mappa header."""
    idx = header_map.get(key)
    if idx is None:
        return None
    if idx >= len(row):
        return None
    return row[idx]


def _parse_date(date_str: str) -> datetime | None:
    """Tenta di parsare una data nel formato GG/MM/YYYY."""
    try:
        d = _normalize_str(date_str)
        if not d:
            return None
        return datetime.strptime(d, "%d/%m/%Y")
    except Exception:
        return None


def _categorize_event(record: Dict[str, Any]) -> List[str]:
    """Categorizza un record in una lista di categorie.
    
    Ritorna una lista come:
    - "disallineamento_qr" se categoria è "disallineamento QR"
    - "firmware_update" se update1 o update2 contengono "firmware"
    - "hw_mother" se parti_coinvolte contiene "scheda madre"
    - "hw_battery" se parti_coinvolte contiene "batteria" o "motore sollevamento"
    """
    categories = []
    categoria = _normalize_str(record.get("categoria", "")).lower()
    
    if "disallineamento" in categoria and "qr" in categoria:
        categories.append("disallineamento_qr")
    
    update1 = _normalize_str(record.get("update1", "")).lower()
    update2 = _normalize_str(record.get("update2", "")).lower()
    if "firmware" in update1 or "firmware" in update2:
        categories.append("firmware_update")
    
    parti = _normalize_str(record.get("parti_coinvolte", "")).lower()
    if "scheda madre" in parti:
        categories.append("hw_mother")
    
    if "batteria" in parti or "motore" in parti or "sollevamento" in parti:
        categories.append("hw_battery")
    
    return categories if categories else ["other"]


def _read_disallineamento_qr_data() -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """Legge i dati e li organizza per robot e categoria.
    
    Ritorna:
    {
      "16278": {
        "disallineamento_qr": [
          {"data": "02/04/2026", "date_obj": datetime(...), "id": 1, "titolo": "...", ...},
          ...
        ],
        "firmware_update": [...],
        "hw_mother": [...],
        "hw_battery": [...],
      },
      ...
    }
    """
    result: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    
    # Inizializza la struttura per ogni robot
    for robot_id in ROBOT_IDS:
        result[robot_id] = {
            "disallineamento_qr": [],
            "firmware_update": [],
            "hw_mother": [],
            "hw_battery": [],
        }
    
    if not _safe_exists_excel():
        return result
    
    try:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb.active
        header_map = _get_header_map(ws)
        
        # Calcola la data di cutoff (settembre dell'anno scorso)
        # Assumendo che "settembre" significhi settembre dello scorso anno
        now = datetime.now()
        cutoff_date = datetime(now.year - 1, 9, 1)
        
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r:
                continue
            
            robot = _normalize_str(_cell_from_row(r, header_map, "robot"))
            if robot not in result:
                continue
            
            data_str = _normalize_str(_cell_from_row(r, header_map, "data"))
            date_obj = _parse_date(data_str)
            
            # Filtra solo records da settembre in poi
            if date_obj is None or date_obj < cutoff_date:
                continue
            
            # Costruisci il record
            record = {
                "id": _cell_from_row(r, header_map, "id"),
                "data": data_str,
                "date_obj": date_obj,
                "ora": _normalize_str(_cell_from_row(r, header_map, "ora")),
                "titolo": _normalize_str(_cell_from_row(r, header_map, "Titolo")),
                "categoria": _normalize_str(_cell_from_row(r, header_map, "Categoria")),
                "update1": _normalize_str(_cell_from_row(r, header_map, "update1")),
                "update2": _normalize_str(_cell_from_row(r, header_map, "update2")),
                "parti_coinvolte": _normalize_str(_cell_from_row(r, header_map, "parti_coinvolte")),
            }
            
            # Categorizza e inserisci nei relativi bucket
            categories = _categorize_event(record)
            for cat in categories:
                if cat in result[robot]:
                    result[robot][cat].append(record)
        
        wb.close()
    except Exception as exc:
        print(f"[DisallineamentoQR] Errore lettura Excel: {exc}")
    
    # Ordina i record per data (decrescente) per ogni categoria
    for robot_data in result.values():
        for category_list in robot_data.values():
            category_list.sort(key=lambda x: x["date_obj"], reverse=True)
    
    return result


@disallineamento_qr_bp.get("/MedicairGeek/disallineamentoQr")
def disallineamento_qr_page():
    """Pagina principale disallineamento QR."""
    data = _read_disallineamento_qr_data()
    log_activity(f"view | page=disallineamentoQr | ip={request.remote_addr}")
    
    return render_template(
        "disallineamentoQr.html",
        title="Disallineamento QR",
        now=datetime.now(),
        robots_data=data,
        robot_ids=ROBOT_IDS,
    )


@disallineamento_qr_bp.get("/MedicairGeek/DisallineamentoQR/data")
def api_disallineamento_qr_data():
    """API per ottenere i dati in JSON."""
    from flask import jsonify
    
    data = _read_disallineamento_qr_data()
    
    # Formatta per JSON
    out = {}
    for robot_id, categories in data.items():
        out[robot_id] = {}
        for category, records in categories.items():
            out[robot_id][category] = [
                {
                    "id": r["id"],
                    "data": r["data"],
                    "ora": r["ora"],
                    "titolo": r["titolo"],
                    "categoria": r["categoria"],
                }
                for r in records
            ]
    
    return jsonify({"ok": True, "data": out})
