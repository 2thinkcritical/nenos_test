"""Render selected synthetic before/after examples in an isolated Linux runner.

Requires LibreOffice and Poppler. Uses a private LibreOffice profile; does not
open or alter a desktop LibreOffice session.
"""

import argparse
import re
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory

EXAMPLES = ("company-and-date", "selected-checkbox", "table-before-total", "null-and-zero")


def render(source: Path, destination: Path, libreoffice: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for name in EXAMPLES:
        for state in ("before", "after"):
            docx = (source / name / f"{state}.docx").resolve()
            if not docx.is_file():
                raise FileNotFoundError(docx)
            with TemporaryDirectory(prefix="docx-render-") as temporary:
                work = Path(temporary)
                process = subprocess.run(
                    [
                        str(libreoffice),
                        f"-env:UserInstallation={(work / 'profile').as_uri()}",
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(work),
                        str(docx),
                    ],
                    capture_output=True,
                    text=True,
                    timeout=60,
                )
                pdf = work / f"{state}.pdf"
                if process.returncode or not pdf.is_file():
                    raise RuntimeError(f"Rendering failed for {name}/{state}: {process.stderr}")
                info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True, text=True)
                pages = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
                if not pages or int(pages.group(1)) != 1:
                    raise RuntimeError(f"Expected one page for {name}/{state}; inspect layout before publishing")
                prefix = destination / f"{name}-{state}"
                subprocess.run(
                    ["pdftoppm", "-png", "-scale-to", "1400", "-singlefile", str(pdf), str(prefix)],
                    check=True,
                    timeout=30,
                )
                print(f"Rendered {name}/{state}: 1 page")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("evaluation/output"))
    parser.add_argument("--destination", type=Path, default=Path("evaluation/previews"))
    parser.add_argument("--libreoffice", type=Path, default=Path("/usr/bin/libreoffice"))
    args = parser.parse_args()
    render(args.source, args.destination, args.libreoffice)


if __name__ == "__main__":
    main()
