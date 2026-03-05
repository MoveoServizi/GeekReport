# email_sender.py
import json
import mimetypes
import smtplib
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List, Optional, Union, Any


# =========================
# PATH BASE (robusto con PM2)
# =========================
BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs" / "email"
LOG_FILE = LOG_DIR / "email_log.jsonl"

TEMPLATES_JSON_PATH = BASE_DIR / "templates" / "email_templates.json"


# =========================
# CONFIG (modifica qui)
# =========================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "Server.moveo@gmail.com"
SENDER_PASSWORD = "bpxl dyil qdid ylyj"  # App Password (se a te funziona così, ok)
USE_TLS = True


# =========================
# TEMPLATE REGISTRY (DEFAULT)
# =========================
DEFAULT_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "REPORT INCIDENTE": {
        "required_fields": ["data", "robots", "note"],
        "subject": "REPORT INCIDENTE [robots] - [data]",
        "body": (
            "Buongiorno,\n"
            "Questa è un'email automatica in seguito ad un evento di incidente all'interno del magazzino Geek.\n\n"
            "EVENTO\n"
            "Data: [data]\n"
            "Robot coinvolti: [robots]\n\n"
            "Descrizione / Note:\n"
            "[note]\n\n"
            "In allegato i file.\n\n"
            "Grazie mille per l'attenzione.\n"
            "Moveo Servizi\n"
        ),
    }
}


@dataclass
class EmailSendResult:
    ok: bool
    recipient: str
    subject: str
    message_id: Optional[str] = None
    error: Optional[str] = None
    log_file: Optional[str] = None


class EmailSender:
    """
    - send_email(): invio diretto
    - send_template(): invio basato su template registry (JSON/dict) con validazione campi richiesti

    Logging:
    - 1 riga JSON per invio (ok/errore) in logs/email/email_log.jsonl (append)
    """

    def __init__(
        self,
        sender_email: str = SENDER_EMAIL,
        sender_password: str = SENDER_PASSWORD,
        smtp_server: str = SMTP_SERVER,
        smtp_port: int = SMTP_PORT,
        use_tls: bool = USE_TLS,
        templates: Optional[Dict[str, Dict[str, Any]]] = None,
        templates_json_path: Path = TEMPLATES_JSON_PATH,
        log_file: Path = LOG_FILE,
        autosync_templates_with_defaults: bool = True,
    ):
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.use_tls = use_tls

        self.templates_json_path = Path(templates_json_path)
        self.log_file = Path(log_file)

        LOG_DIR.mkdir(parents=True, exist_ok=True)
        self.log_file.parent.mkdir(parents=True, exist_ok=True)
        if not self.log_file.exists():
            self.log_file.touch()

        # Templates
        if templates is not None:
            self.templates = templates
        else:
            file_templates = self._load_templates_from_json(self.templates_json_path)

            if not file_templates:
                # Se non esiste, crea da default
                self.templates = dict(DEFAULT_TEMPLATES)
                self._save_templates_to_json(self.templates_json_path, self.templates)
            else:
                # Merge/upgrade per evitare mismatch (es. manca note)
                if autosync_templates_with_defaults:
                    self.templates = self._merge_templates_with_defaults(file_templates, DEFAULT_TEMPLATES)
                    self._save_templates_to_json(self.templates_json_path, self.templates)
                else:
                    self.templates = file_templates

    # -------------------------
    # Public API
    # -------------------------
    def send_email(
        self,
        destinatario: str,
        oggetto: str,
        testo: str,
        allegati: Optional[List[Union[str, Path]]] = None,
    ) -> EmailSendResult:
        allegati = allegati or []
        try:
            msg = self._build_message(to=destinatario, subject=oggetto, body=testo, attachments=allegati)
            message_id = self._smtp_send(msg)

            self._write_log(
                ok=True,
                kind="send_email",
                template_name=None,
                used_fields=None,
                recipient=destinatario,
                subject=oggetto,
                attachments=allegati,
                message_id=message_id,
                error=None,
            )

            return EmailSendResult(ok=True, recipient=destinatario, subject=oggetto, message_id=message_id, log_file=str(self.log_file))

        except Exception as e:
            self._write_log(
                ok=False,
                kind="send_email",
                template_name=None,
                used_fields=None,
                recipient=destinatario,
                subject=oggetto,
                attachments=allegati,
                message_id=None,
                error=str(e),
            )
            return EmailSendResult(ok=False, recipient=destinatario, subject=oggetto, error=str(e), log_file=str(self.log_file))

    def send_template(
        self,
        destinatario: str,
        template: str,
        campi: Dict[str, str],
        allegati: Optional[List[Union[str, Path]]] = None,
    ) -> EmailSendResult:
        allegati = allegati or []

        # normalizzazione campi
        campi = {str(k): ("" if v is None else str(v)) for k, v in (campi or {}).items()}

        try:
            tpl = self.templates.get(template)
            if not tpl:
                raise ValueError(f"Template non trovato: '{template}'. Disponibili: {', '.join(self.templates.keys())}")

            required = tpl.get("required_fields", [])
            missing = [k for k in required if k not in campi or campi.get(k, "").strip() == ""]
            if missing:
                raise ValueError(f"Campi mancanti per template '{template}': {missing}. Richiesti: {required}")

            subject_raw = tpl.get("subject", "")
            body_raw = tpl.get("body", "")

            subject = self._apply_placeholders(subject_raw, campi)
            body = self._apply_placeholders(body_raw, campi)

            msg = self._build_message(to=destinatario, subject=subject, body=body, attachments=allegati)
            message_id = self._smtp_send(msg)

            self._write_log(
                ok=True,
                kind="send_template",
                template_name=template,
                used_fields=campi,
                recipient=destinatario,
                subject=subject,
                attachments=allegati,
                message_id=message_id,
                error=None,
            )

            return EmailSendResult(ok=True, recipient=destinatario, subject=subject, message_id=message_id, log_file=str(self.log_file))

        except Exception as e:
            self._write_log(
                ok=False,
                kind="send_template",
                template_name=template,
                used_fields=campi,
                recipient=destinatario,
                subject="(template error)",
                attachments=allegati,
                message_id=None,
                error=str(e),
            )
            return EmailSendResult(ok=False, recipient=destinatario, subject="(template error)", error=str(e), log_file=str(self.log_file))

    # -------------------------
    # Templates I/O + merge/upgrade
    # -------------------------
    def _load_templates_from_json(self, path: Path) -> Dict[str, Dict[str, Any]]:
        try:
            if not path.exists():
                return {}
            raw = path.read_text(encoding="utf-8").strip()
            if not raw:
                return {}
            data = json.loads(raw)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_templates_to_json(self, path: Path, templates: Dict[str, Dict[str, Any]]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(templates, indent=2, ensure_ascii=False), encoding="utf-8")

    def _merge_templates_with_defaults(
        self,
        current: Dict[str, Dict[str, Any]],
        defaults: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Dict[str, Any]]:
        merged: Dict[str, Dict[str, Any]] = dict(current)

        for name, def_tpl in defaults.items():
            if name not in merged or not isinstance(merged.get(name), dict):
                merged[name] = dict(def_tpl)
                continue

            cur_tpl = merged[name]

            # assicura chiavi base
            for key in ["required_fields", "subject", "body"]:
                if key not in cur_tpl or cur_tpl.get(key) in (None, ""):
                    cur_tpl[key] = def_tpl.get(key)

            # upgrade specifico: REPORT INCIDENTE deve avere note e placeholder
            if name == "REPORT INCIDENTE":
                req = cur_tpl.get("required_fields", [])
                if not isinstance(req, list):
                    req = []
                for needed in ["data", "robots", "note"]:
                    if needed not in req:
                        req.append(needed)
                cur_tpl["required_fields"] = req

                body = str(cur_tpl.get("body", "") or "")
                if "[note]" not in body:
                    body = body.rstrip() + "\n\nDescrizione / Note:\n[note]\n"
                cur_tpl["body"] = body

            merged[name] = cur_tpl

        return merged

    # -------------------------
    # Internals (SMTP/MIME)
    # -------------------------
    def _smtp_send(self, msg: EmailMessage) -> str:
        with smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30) as server:
            server.ehlo()
            if self.use_tls:
                server.starttls()
                server.ehlo()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
        return msg.get("Message-ID", "")

    def _build_message(self, to: str, subject: str, body: str, attachments: List[Union[str, Path]]) -> EmailMessage:
        msg = EmailMessage()
        msg["From"] = self.sender_email
        msg["To"] = to
        msg["Subject"] = subject
        msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S")
        msg["Message-ID"] = f"<{datetime.now().strftime('%Y%m%d%H%M%S%f')}@moveo.local>"
        msg.set_content(body)

        for item in attachments:
            path = Path(item)
            if not path.exists() or not path.is_file():
                raise FileNotFoundError(f"Allegato non trovato o non è un file: {path}")

            ctype, encoding = mimetypes.guess_type(str(path))
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)

            msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)

        return msg

    def _apply_placeholders(self, text: str, fields: Dict[str, str]) -> str:
        out = text
        for k, v in (fields or {}).items():
            out = out.replace(f"[{k}]", str(v))
        return out

    # -------------------------
    # Logging (JSONL append su file unico)
    # -------------------------
    def _write_log(
        self,
        ok: bool,
        kind: str,
        template_name: Optional[str],
        used_fields: Optional[Dict[str, str]],
        recipient: str,
        subject: str,
        attachments: List[Union[str, Path]],
        message_id: Optional[str],
        error: Optional[str],
    ) -> None:
        record = {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "ok": ok,
            "kind": kind,
            "from": self.sender_email,
            "to": recipient,
            "subject": subject,
            "template": template_name,
            "fields": used_fields,
            "message_id": message_id or "",
            "attachments": [str(Path(a)) for a in (attachments or [])],
            "error": error,
        }
        line = json.dumps(record, ensure_ascii=False) + "\n"
        with self.log_file.open("a", encoding="utf-8") as f:
            f.write(line)