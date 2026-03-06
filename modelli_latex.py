import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Union

from config import LATEX_PATH


DEFAULT_PDFLATEX_CMD = str(Path(LATEX_PATH) / "pdflatex.exe")


@dataclass
class ReportResult:
    model: str
    tex_path: Path
    pdf_path: Path
    build_dir: Path


class LatexReportManager:
    """
    Gestore report LaTeX:
    - legge template .tex
    - sostituisce placeholder {{{CHIAVE}}}
    - compila con pdflatex in una cartella tmp dedicata
    - copia assets (LOGO.png)
    - gestisce AllegatiList in modo robusto
    - salva file di debug per diagnosi errori
    - cleanup opzionale della cartella tmp
    """

    PLACEHOLDER_PATTERN = re.compile(r"\{\{\{([^}]+)\}\}\}")

    def __init__(
        self,
        base_dir: Optional[Path] = None,
        templates_dir: Optional[Path] = None,
        tmp_dir: Optional[Path] = None,
        pdflatex_cmd: Optional[str] = None,
        compile_passes: int = 2,
        keep_temp_on_error: bool = True,
        compile_timeout: int = 120,
    ):
        self.base_dir = base_dir or Path(__file__).resolve().parent
        self.templates_dir = templates_dir or (self.base_dir / "latex" / "templates")
        self.tmp_dir = tmp_dir or (self.base_dir / "latex" / "tmp")

        self.pdflatex_cmd = pdflatex_cmd or DEFAULT_PDFLATEX_CMD
        self.compile_passes = max(1, int(compile_passes))
        self.keep_temp_on_error = bool(keep_temp_on_error)
        self.compile_timeout = int(compile_timeout)

        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.tmp_dir.mkdir(parents=True, exist_ok=True)

        self.model_map = {
            "modello": "report_incidente.tex",
        }

        # Chiavi che contengono LaTeX voluto e NON vanno escape-ate
        self.raw_keys = {"AllegatiList"}

    def crea_report(
        self,
        modello: str,
        campi: Dict[str, Any],
        nome_file: Optional[str] = None
    ) -> ReportResult:
        template_path = self._get_template_path(modello)

        if not nome_file:
            stamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
            nome_file = f"REPORT_{stamp}"

        build_dir = self.tmp_dir / nome_file
        build_dir.mkdir(parents=True, exist_ok=True)

        try:
            self._write_debug_preflight(build_dir, modello, nome_file, template_path, campi)

            template_str = template_path.read_text(encoding="utf-8")

            campi = dict(campi)
            campi["AllegatiList"] = self._normalize_allegati_list(campi.get("AllegatiList", ""))

            rendered = self._render_template(template_str, campi)

            tex_work_path = build_dir / f"{nome_file}.tex"
            tex_work_path.write_text(rendered, encoding="utf-8")

            self._write_debug_json(
                build_dir / "debug_10_compile_start.json",
                {
                    "timestamp": datetime.now().isoformat(),
                    "tex_path": str(tex_work_path),
                    "build_dir": str(build_dir),
                    "pdflatex_cmd": self.pdflatex_cmd,
                    "tex_exists": tex_work_path.exists(),
                    "tex_size_bytes": tex_work_path.stat().st_size if tex_work_path.exists() else None,
                },
            )

            self._copy_assets(build_dir)
            self._compile_pdf(tex_work_path, build_dir)

            pdf_work_path = build_dir / f"{nome_file}.pdf"
            if not pdf_work_path.exists():
                log = build_dir / f"{nome_file}.log"
                raise RuntimeError(f"PDF non generato. Controlla log: {log}")

            self._write_debug_json(
                build_dir / "debug_20_success.json",
                {
                    "timestamp": datetime.now().isoformat(),
                    "pdf_path": str(pdf_work_path),
                    "pdf_exists": pdf_work_path.exists(),
                    "pdf_size_bytes": pdf_work_path.stat().st_size if pdf_work_path.exists() else None,
                },
            )

            return ReportResult(
                model=modello,
                tex_path=tex_work_path,
                pdf_path=pdf_work_path,
                build_dir=build_dir,
            )

        except Exception as exc:
            self._write_debug_json(
                build_dir / "debug_99_exception.json",
                {
                    "timestamp": datetime.now().isoformat(),
                    "error_type": type(exc).__name__,
                    "error_message": str(exc),
                    "pdflatex_cmd": self.pdflatex_cmd,
                    "build_dir": str(build_dir),
                },
            )

            if not self.keep_temp_on_error:
                self._safe_rmtree(build_dir)
            raise

    def cleanup_report(self, res: ReportResult) -> None:
        self._safe_rmtree(res.build_dir)

    # -------------------------
    # Internals
    # -------------------------
    def _get_template_path(self, modello: str) -> Path:
        if modello not in self.model_map:
            raise ValueError(
                f"Modello '{modello}' non registrato. Disponibili: {list(self.model_map.keys())}"
            )

        path = self.templates_dir / self.model_map[modello]
        if not path.exists():
            raise FileNotFoundError(f"Template non trovato: {path}")
        return path

    def _copy_assets(self, build_dir: Path) -> None:
        src_logo = self.base_dir / "static" / "img" / "LOGO.png"
        dst_img_dir = build_dir / "img"
        dst_img_dir.mkdir(parents=True, exist_ok=True)

        if not src_logo.exists():
            raise FileNotFoundError(f"Logo non trovato: {src_logo}")

        shutil.copy2(src_logo, dst_img_dir / "LOGO.png")

        self._write_debug_json(
            build_dir / "debug_05_assets.json",
            {
                "timestamp": datetime.now().isoformat(),
                "src_logo": str(src_logo),
                "src_logo_exists": src_logo.exists(),
                "dst_logo": str(dst_img_dir / "LOGO.png"),
                "dst_logo_exists": (dst_img_dir / "LOGO.png").exists(),
            },
        )

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
        resolved_cmd = Path(self.pdflatex_cmd)

        self._write_debug_json(
            workdir / "debug_11_command.json",
            {
                "timestamp": datetime.now().isoformat(),
                "pdflatex_cmd": self.pdflatex_cmd,
                "pdflatex_exists": resolved_cmd.exists(),
                "pdflatex_is_file": resolved_cmd.is_file(),
                "cwd": str(workdir),
                "compile_passes": self.compile_passes,
                "compile_timeout_sec": self.compile_timeout,
            },
        )

        try:
            version_proc = subprocess.run(
                [self.pdflatex_cmd, "--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=20,
            )
            (workdir / "pdflatex_version.txt").write_text(version_proc.stdout, encoding="utf-8")
        except Exception as exc:
            (workdir / "pdflatex_version.txt").write_text(
                f"Impossibile leggere versione pdflatex: {exc}",
                encoding="utf-8",
            )

        cmd = [
            self.pdflatex_cmd,
            "-interaction=nonstopmode",
            "-halt-on-error",
            tex_path.name,
        ]

        for i in range(self.compile_passes):
            pass_num = i + 1
            pass_out = workdir / f"pdflatex_pass_{pass_num}.txt"
            pass_json = workdir / f"debug_12_pass_{pass_num}.json"

            try:
                proc = subprocess.run(
                    cmd,
                    cwd=str(workdir),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    timeout=self.compile_timeout,
                )
            except subprocess.TimeoutExpired as exc:
                pass_out.write_text(
                    (exc.stdout or "") + "\n\n[TIMEOUT]\n" + str(exc),
                    encoding="utf-8",
                )
                self._write_debug_json(
                    pass_json,
                    {
                        "timestamp": datetime.now().isoformat(),
                        "pass": pass_num,
                        "status": "timeout",
                        "timeout_sec": self.compile_timeout,
                    },
                )
                raise RuntimeError(
                    f"Timeout pdflatex al pass {pass_num}. Controlla: {pass_out}"
                ) from exc

            pass_out.write_text(proc.stdout or "", encoding="utf-8")

            self._write_debug_json(
                pass_json,
                {
                    "timestamp": datetime.now().isoformat(),
                    "pass": pass_num,
                    "returncode": proc.returncode,
                    "output_file": str(pass_out),
                },
            )

            if proc.returncode != 0:
                out_txt = workdir / "pdflatex_stdout.txt"
                out_txt.write_text(proc.stdout or "", encoding="utf-8")
                snippet = self._extract_latex_error_snippet(proc.stdout or "")
                raise RuntimeError(
                    f"Errore pdflatex (pass {pass_num}/{self.compile_passes}, code {proc.returncode}). "
                    f"Vedi: {out_txt} e/o il .log in {workdir}\n{snippet}"
                )

    def _extract_latex_error_snippet(self, stdout: str, max_lines: int = 25) -> str:
        lines = stdout.splitlines()
        idx = None
        for i, line in enumerate(lines):
            if line.startswith("! "):
                idx = i
                break

        if idx is None:
            return "Dettaglio (tail):\n" + "\n".join(lines[-max_lines:])

        return "Dettaglio (errore):\n" + "\n".join(lines[idx: idx + max_lines])

    def _write_debug_preflight(
        self,
        build_dir: Path,
        modello: str,
        nome_file: str,
        template_path: Path,
        campi: Dict[str, Any],
    ) -> None:
        env_path = os.environ.get("PATH", "")
        pdflatex_path = Path(self.pdflatex_cmd)

        self._write_debug_json(
            build_dir / "debug_00_preflight.json",
            {
                "timestamp": datetime.now().isoformat(),
                "model": modello,
                "nome_file": nome_file,
                "base_dir": str(self.base_dir),
                "templates_dir": str(self.templates_dir),
                "tmp_dir": str(self.tmp_dir),
                "template_path": str(template_path),
                "template_exists": template_path.exists(),
                "build_dir": str(build_dir),
                "pdflatex_cmd": self.pdflatex_cmd,
                "pdflatex_exists": pdflatex_path.exists(),
                "pdflatex_is_file": pdflatex_path.is_file(),
                "cwd": str(Path.cwd()),
                "python_executable": str(Path(os.sys.executable)),
                "path_env": env_path,
                "campi_keys": sorted(list(campi.keys())),
            },
        )

    def _write_debug_json(self, path: Path, data: Dict[str, Any]) -> None:
        try:
            path.write_text(
                json.dumps(data, indent=2, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
        except Exception:
            pass

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
        _manager_singleton = LatexReportManager(
            pdflatex_cmd=DEFAULT_PDFLATEX_CMD,
            keep_temp_on_error=True,
        )
    return _manager_singleton


def crea_report(
    modello: str,
    campi: Dict[str, Any],
    nome_file: Optional[str] = None
) -> ReportResult:
    return get_manager().crea_report(modello, campi, nome_file=nome_file)


def cleanup_latex_tmp() -> int:
    """
    Svuota completamente la cartella latex/tmp.
    Ritorna il numero di elementi eliminati.
    """
    tmp_dir = Path(__file__).resolve().parent / "latex" / "tmp"
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
        except Exception as exc:
            print(f"[LATEX CLEANUP] Impossibile eliminare {item}: {exc}")
    return deleted