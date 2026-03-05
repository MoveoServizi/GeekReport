import re
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Iterable, Union, List
import shutil
from pathlib import Path

@dataclass
class ReportResult:
    model: str
    tex_path: Path
    pdf_path: Path
    build_dir: Path


class LatexReportManager:
    """
    - legge template .tex
    - sostituisce placeholder {{{CHIAVE}}}
    - compila con pdflatex in una cartella tmp dedicata
    - copia assets (LOGO.png)
    - gestisce AllegatiList in modo robusto (nomi file con _, %, &, ecc.)
    - cleanup: elimina la cartella tmp del report dopo uso
    """

    PLACEHOLDER_PATTERN = re.compile(r"\{\{\{([^}]+)\}\}\}")

    def __init__(
        self,
        base_dir: Optional[Path] = None,
        templates_dir: Optional[Path] = None,
        tmp_dir: Optional[Path] = None,
        pdflatex_cmd: str = "pdflatex",
        compile_passes: int = 2,
        keep_temp_on_error: bool = True,
    ):
        self.base_dir = base_dir or Path(__file__).resolve().parent
        self.templates_dir = templates_dir or (self.base_dir / "latex" / "templates")
        self.tmp_dir = tmp_dir or (self.base_dir / "latex" / "tmp")

        self.pdflatex_cmd = pdflatex_cmd
        self.compile_passes = max(1, int(compile_passes))
        self.keep_temp_on_error = bool(keep_temp_on_error)

        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.tmp_dir.mkdir(parents=True, exist_ok=True)

        self.model_map = {
            "modello": "report_incidente.tex",
        }

        # Campi che contengono LaTeX "voluto"
        self.raw_keys = {"AllegatiList"}

    def crea_report(self, modello: str, campi: Dict[str, Any], nome_file: Optional[str] = None) -> ReportResult:
        template_path = self._get_template_path(modello)

        if not nome_file:
            stamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
            nome_file = f"REPORT_{stamp}"

        build_dir = self.tmp_dir / nome_file
        build_dir.mkdir(parents=True, exist_ok=True)

        try:
            template_str = template_path.read_text(encoding="utf-8")

            # Normalizza AllegatiList (robusto) prima del render
            campi = dict(campi)
            campi["AllegatiList"] = self._normalize_allegati_list(campi.get("AllegatiList", ""))

            rendered = self._render_template(template_str, campi)

            tex_work_path = build_dir / f"{nome_file}.tex"
            tex_work_path.write_text(rendered, encoding="utf-8")

            self._copy_assets(build_dir)

            self._compile_pdf(tex_work_path, build_dir)

            pdf_work_path = build_dir / f"{nome_file}.pdf"
            if not pdf_work_path.exists():
                log = build_dir / f"{nome_file}.log"
                raise RuntimeError(f"PDF non generato. Controlla log: {log}")

            return ReportResult(
                model=modello,
                tex_path=tex_work_path,
                pdf_path=pdf_work_path,
                build_dir=build_dir,
            )

        except Exception:
            # se keep_temp_on_error=True NON cancelliamo: serve per debug
            if not self.keep_temp_on_error:
                self._safe_rmtree(build_dir)
            raise

    def cleanup_report(self, res: ReportResult) -> None:
        """Cancella tutta la cartella temporanea del report (latex/tmp/<nome_report>)."""
        self._safe_rmtree(res.build_dir)

    # -------------------------
    # Internals
    # -------------------------
    def _get_template_path(self, modello: str) -> Path:
        if modello not in self.model_map:
            raise ValueError(f"Modello '{modello}' non registrato. Disponibili: {list(self.model_map.keys())}")
        p = self.templates_dir / self.model_map[modello]
        if not p.exists():
            raise FileNotFoundError(f"Template non trovato: {p}")
        return p

    def _copy_assets(self, build_dir: Path) -> None:
        src_logo = self.base_dir / "static" / "img" / "LOGO.png"
        dst_img_dir = build_dir / "img"
        dst_img_dir.mkdir(parents=True, exist_ok=True)

        if not src_logo.exists():
            raise FileNotFoundError(f"Logo non trovato: {src_logo}")

        shutil.copy2(src_logo, dst_img_dir / "LOGO.png")

    def _render_template(self, template: str, campi: Dict[str, Any]) -> str:
        def repl(match: re.Match) -> str:
            key = match.group(1).strip()
            val = campi.get(key, "")
            if val is None:
                val = ""
            if key in self.raw_keys:
                return str(val)
            return self._latex_escape(str(val))

        return self.PLACEHOLDER_PATTERN.sub(repl, template)

    def _latex_escape(self, s: str) -> str:
        replacements = {
            "\\": r"\textbackslash{}",
            "&": r"\&",
            "%": r"\%",
            "$": r"\$",
            "#": r"\#",
            "_": r"\_",
            "{": r"\{",
            "}": r"\}",
            "~": r"\textasciitilde{}",
            "^": r"\textasciicircum{}",
        }
        return "".join(replacements.get(ch, ch) for ch in s)

    def _normalize_allegati_list(self, allegati: Any) -> str:
        """
        Produce SEMPRE una lista compilabile usando \FileItem{...}
        (la macro \FileItem deve esistere nel template).
        Accetta:
        - lista di Path/string
        - stringa con righe "\item nome"
        - stringa con righe "nome"
        - già in formato \FileItem{...}
        """
        if allegati is None:
            return ""

        if isinstance(allegati, (list, tuple, set)):
            return self._build_file_items(allegati)

        if isinstance(allegati, Path):
            return self._build_file_items([allegati])

        if isinstance(allegati, str):
            lines = [ln.strip() for ln in allegati.splitlines() if ln.strip()]
            if not lines:
                return ""

            if any(ln.lstrip().startswith(r"\FileItem") for ln in lines):
                return "\n".join(lines)

            converted: List[str] = []
            for ln in lines:
                if ln.startswith(r"\item"):
                    rest = ln[len(r"\item"):].strip().replace("\\", "/")
                    converted.append(rf"\FileItem{{{rest}}}")
                else:
                    rest = ln.replace("\\", "/")
                    converted.append(rf"\FileItem{{{rest}}}")
            return "\n".join(converted)

        return self._build_file_items([str(allegati)])

    def _build_file_items(self, items: Iterable[Union[str, Path]]) -> str:
        out: List[str] = []
        for it in items:
            p = it if isinstance(it, Path) else Path(str(it))
            name = p.name if getattr(p, "name", "") else str(it)
            name = name.replace("\\", "/")
            out.append(rf"\FileItem{{{name}}}")
        return "\n".join(out)

    def _compile_pdf(self, tex_path: Path, workdir: Path) -> None:
        cmd = [self.pdflatex_cmd, "-interaction=nonstopmode", "-halt-on-error", tex_path.name]

        for i in range(self.compile_passes):
            proc = subprocess.run(
                cmd,
                cwd=str(workdir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
            if proc.returncode != 0:
                out_txt = workdir / "pdflatex_stdout.txt"
                out_txt.write_text(proc.stdout, encoding="utf-8")
                snippet = self._extract_latex_error_snippet(proc.stdout)
                raise RuntimeError(
                    f"Errore pdflatex (pass {i+1}/{self.compile_passes}, code {proc.returncode}). "
                    f"Vedi: {out_txt} e/o il .log in {workdir}\n{snippet}"
                )

    def _extract_latex_error_snippet(self, stdout: str, max_lines: int = 25) -> str:
        lines = stdout.splitlines()
        idx = None
        for i, ln in enumerate(lines):
            if ln.startswith("! "):
                idx = i
                break
        if idx is None:
            return "Dettaglio (tail):\n" + "\n".join(lines[-max_lines:])
        return "Dettaglio (errore):\n" + "\n".join(lines[idx:idx + max_lines])

    def _safe_rmtree(self, path: Path) -> None:
        try:
            if path.exists():
                shutil.rmtree(path, ignore_errors=True)
        except Exception:
            pass


# Singleton + helpers
_manager_singleton: Optional[LatexReportManager] = None


def get_manager() -> LatexReportManager:
    global _manager_singleton
    if _manager_singleton is None:
        # keep_temp_on_error=True: se fallisce LaTeX resta la cartella tmp per debug
        _manager_singleton = LatexReportManager(keep_temp_on_error=True)
    return _manager_singleton


def crea_report(modello: str, campi: Dict[str, Any], nome_file: Optional[str] = None) -> ReportResult:
    return get_manager().crea_report(modello, campi, nome_file=nome_file)

def cleanup_latex_tmp() -> int:
    """
    Svuota completamente la cartella latex/tmp (file e sottocartelle).
    Ritorna il numero di elementi eliminati.
    """
    tmp_dir = Path("latex/tmp")  # Adjust the path as needed
    if not tmp_dir.exists() or not tmp_dir.is_dir():
        return 0

    deleted = 0
    for item in tmp_dir.iterdir():
        try:
            if item.is_dir():
                shutil.rmtree(item, ignore_errors=False)
            else:
                item.unlink()
            deleted += 1
        except Exception as e:
            print(f"[LATEX CLEANUP] Impossibile eliminare {item}: {e}")
    return deleted