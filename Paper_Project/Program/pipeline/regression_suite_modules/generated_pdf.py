"""Small synthetic PDF fixture helpers for regression tests."""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Iterable


def poppler_available() -> bool:
    return bool(shutil.which("pdfinfo") and shutil.which("pdftotext"))


def write_text_pdf(
    path: Path,
    lines: Iterable[tuple[str, float, float, float]],
    *,
    page_width: float = 595.0,
    page_height: float = 842.0,
) -> None:
    """Write a simple one-page text PDF.

    Each line tuple is ``(text, size_pt, x_pt, y_pt)``. Text is limited to
    Latin-1-safe characters for this tiny PDF writer; product parsing itself is
    tested through Poppler, not through a handcrafted PDF engine.
    """

    content_parts = []
    for text, size, x, y in lines:
        content_parts.append(f"BT /F1 {size:g} Tf 1 0 0 1 {x:g} {y:g} Tm ({_escape_pdf_text(text)}) Tj ET\n")
    _write_pdf(path, "".join(content_parts).encode("latin-1"), page_width, page_height)


def write_blank_pdf(path: Path, *, page_width: float = 595.0, page_height: float = 842.0) -> None:
    _write_pdf(path, b"", page_width, page_height)


def _write_pdf(path: Path, content: bytes, page_width: float, page_height: float) -> None:
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_width:g} {page_height:g}] "
            "/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"
        ).encode("ascii"),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(content)).encode("ascii") + b" >>\nstream\n" + content + b"endstream",
    ]
    output = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for idx, body in enumerate(objects, start=1):
        offsets.append(len(output))
        output.extend(f"{idx} 0 obj\n".encode("ascii"))
        output.extend(body)
        output.extend(b"\nendobj\n")
    xref_offset = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(
        f"trailer << /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii")
    )
    path.write_bytes(output)


def _escape_pdf_text(text: str) -> str:
    return text.encode("latin-1", errors="replace").decode("latin-1").replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
