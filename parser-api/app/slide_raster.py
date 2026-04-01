"""
PPTX -> PDF (LibreOffice headless) -> per-page JPEG (PyMuPDF).

슬라이드 EMU 좌표 JSON과 동일 종횡비를 유지하려면 PDF 페이지가 슬라이드와 1:1이어야 하며,
LibreOffice 변환이 일반적으로 그렇게 맞춘다.
"""

from __future__ import annotations

import base64
import io
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any


def _page_to_jpeg_bytes(page: Any, scale: float, jpg_quality: int) -> bytes:
    """
    PDF 페이지를 JPEG로 인코딩.

    PyMuPDF 일부 버전에서 Pixmap(csRGB, rgba_source) 결과가 여전히 4채널이라
    tobytes(\"jpeg\")가 \"cannot have alpha\"로 실패하고, 예외가 삼켜지면 0/48이 된다.
    alpha=False + DeviceRGB로 3채널 pixmap을 직접 받아 JPEG로 만든다.
    """
    import fitz

    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(
        matrix=mat,
        alpha=False,
        colorspace=fitz.csRGB,
    )
    try:
        return pix.tobytes("jpeg", jpg_quality=jpg_quality)
    finally:
        if hasattr(pix, "close"):
            pix.close()


def collect_pptx_fonts(pptx_bytes: bytes) -> list[str]:
    """
    PPTX ZIP에서 슬라이드들의 XML을 직접 스캔해 사용된 폰트 이름 목록을 반환.
    python-pptx 의존 없이 동작하며 래스터 meta.missingFontHint에 사용.
    """
    import zipfile
    import xml.etree.ElementTree as ET

    fonts: set[str] = set()
    latin_tags = {
        "{http://schemas.openxmlformats.org/drawingml/2006/main}latin",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}ea",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}cs",
    }
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            targets = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
            for name in targets[:20]:
                try:
                    root = ET.fromstring(z.read(name))
                    for elem in root.iter():
                        if elem.tag in latin_tags:
                            typeface = elem.get("typeface", "").strip()
                            if typeface and not typeface.startswith("+"):
                                fonts.add(typeface)
                except Exception:
                    continue
    except Exception:
        pass
    return sorted(fonts)


def _check_system_fonts(font_names: list[str]) -> list[str]:
    """
    시스템에 설치된 폰트 목록과 대조해 없는 폰트 반환.
    fc-list(Linux/Mac) 또는 Windows 폰트 폴더 이름 스캔을 사용.
    """
    import platform

    missing: list[str] = []
    if not font_names:
        return missing

    system = platform.system()
    installed: set[str] = set()

    if system in ("Linux", "Darwin"):
        try:
            result = subprocess.run(
                ["fc-list", "--format=%{family}\n"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            for line in result.stdout.splitlines():
                for part in line.split(","):
                    installed.add(part.strip().lower())
        except Exception:
            return []
    elif system == "Windows":
        font_dirs = [
            Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts",
        ]
        for d in font_dirs:
            if d.is_dir():
                for f in d.iterdir():
                    installed.add(f.stem.lower())
    else:
        return []

    for font in font_names:
        normalized = font.lower().replace(" ", "").replace("-", "")
        found = any(
            normalized in installed_f.replace(" ", "").replace("-", "")
            for installed_f in installed
        )
        if not found:
            missing.append(font)
    return missing


def _find_soffice() -> str | None:
    env = os.getenv("PPTX_LIBREOFFICE_PATH", "").strip().strip('"')
    if env and Path(env).is_file():
        return env
    for c in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(c).is_file():
            return c
    return None


def _env_bool(key: str, default: bool) -> bool:
    v = os.getenv(key, "").strip().lower()
    if v in ("0", "false", "no", "off"):
        return False
    if v in ("1", "true", "yes", "on"):
        return True
    return default


def _raster_long_edge() -> int:
    raw = os.getenv("PPTX_RASTER_MAX_LONG_EDGE", "1280").strip()
    try:
        return max(320, min(int(raw), 4096))
    except ValueError:
        return 1280


def _jpeg_quality() -> int:
    raw = os.getenv("PPTX_RASTER_JPEG_QUALITY", "82").strip()
    try:
        return max(50, min(int(raw), 95))
    except ValueError:
        return 82


def _max_bytes_slide() -> int:
    raw = os.getenv("PPTX_RASTER_MAX_BYTES_PER_SLIDE", "2000000").strip()
    try:
        return max(100_000, int(raw))
    except ValueError:
        return 2_000_000


def _timeout_sec() -> int:
    raw = os.getenv("PPTX_RASTER_TIMEOUT_SEC", "240").strip()
    try:
        return max(30, int(raw))
    except ValueError:
        return 240


def render_slide_rasters_jpeg(
    pptx_bytes: bytes,
    slide_count: int,
) -> tuple[list[str | None], dict[str, Any]]:
    """
    Returns (per-slide data:image/jpeg;base64,... or None), meta for JSON meta.slideRaster.
    """
    meta: dict[str, Any] = {
        "enabled": True,
        "status": "skipped",
        "reason": "",
        "longEdgePx": _raster_long_edge(),
        "jpegQuality": _jpeg_quality(),
    }

    if slide_count <= 0:
        meta["status"] = "skipped"
        meta["reason"] = "no slides"
        return [], meta

    if not _env_bool("PPTX_SLIDE_RASTER", True):
        meta["status"] = "disabled"
        meta["reason"] = "PPTX_SLIDE_RASTER=0"
        return [None] * slide_count, meta

    soffice = _find_soffice()
    if not soffice:
        meta["status"] = "skipped"
        meta["reason"] = (
            "LibreOffice (soffice) not found; install LO or set PPTX_LIBREOFFICE_PATH"
        )
        return [None] * slide_count, meta

    # 폰트 진단: PPTX 사용 폰트 목록 + 시스템 미설치 폰트 감지
    pptx_fonts = collect_pptx_fonts(pptx_bytes)
    if pptx_fonts:
        meta["pptxFonts"] = pptx_fonts
    missing_fonts = _check_system_fonts(pptx_fonts)
    if missing_fonts:
        meta["missingFonts"] = missing_fonts
        meta["missingFontHint"] = (
            "LibreOffice가 이 폰트를 찾지 못하면 해당 텍스트 블록이 래스터에서 통째로 사라질 수 있습니다. "
            "시스템에 폰트를 설치하거나 LibreOffice 설정에서 대체 폰트를 지정하세요."
        )

    try:
        import fitz  # PyMuPDF
    except ImportError:
        meta["status"] = "error"
        meta["reason"] = "PyMuPDF (pymupdf) not installed"
        return [None] * slide_count, meta

    out_urls: list[str | None] = [None] * slide_count

    try:
        with tempfile.TemporaryDirectory(prefix="pptx_raster_") as tmp:
            tmp_path = Path(tmp)
            pptx_path = tmp_path / "deck.pptx"
            pptx_path.write_bytes(pptx_bytes)

            pdf_path = pptx_path.with_suffix(".pdf")
            timeout = _timeout_sec()
            proc = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(tmp_path),
                    str(pptx_path),
                ],
                capture_output=True,
                timeout=timeout,
            )
            if proc.returncode != 0 or not pdf_path.is_file():
                meta["status"] = "error"
                err = (proc.stderr or b"")[:2000].decode("utf-8", errors="replace")
                meta["reason"] = f"soffice pdf exit={proc.returncode}: {err}"
                return out_urls, meta

            doc = fitz.open(pdf_path)
            render_errors: list[str] = []
            try:
                n_pdf = len(doc)
                n_slides = slide_count
                meta["pdfPageCount"] = n_pdf
                if n_pdf != n_slides:
                    meta["pageCountMismatch"] = {"pdfPages": n_pdf, "pptxSlides": n_slides}

                max_edge = _raster_long_edge()
                q = _jpeg_quality()
                max_b = _max_bytes_slide()

                for i in range(min(n_pdf, n_slides)):
                    page = doc.load_page(i)
                    rect = page.rect
                    w, h = rect.width, rect.height
                    if w <= 0 or h <= 0:
                        continue
                    scale = max_edge / max(w, h)
                    if scale > 1.0:
                        scale = 1.0
                    q_use = q
                    try:
                        jpeg_bytes = _page_to_jpeg_bytes(page, scale, q_use)
                    except Exception as ex:
                        if len(render_errors) < 3:
                            render_errors.append(f"s{i + 1}: {ex!s}"[:240])
                        continue
                    if len(jpeg_bytes) > max_b:
                        try:
                            jpeg_bytes = _page_to_jpeg_bytes(
                                page,
                                scale * 0.72,
                                max(55, q_use - 12),
                            )
                        except Exception as ex:
                            if len(render_errors) < 3:
                                render_errors.append(
                                    f"s{i + 1} (retry): {ex!s}"[:240],
                                )
                            continue
                    if len(jpeg_bytes) <= max_b:
                        b64 = base64.b64encode(jpeg_bytes).decode("ascii")
                        out_urls[i] = f"data:image/jpeg;base64,{b64}"

                rendered = sum(1 for u in out_urls if u)
                meta["slidesRendered"] = rendered
                if rendered == 0 and n_slides > 0:
                    if n_pdf == 0:
                        meta["status"] = "error"
                        meta["reason"] = (
                            "PDF에 페이지가 없습니다. LibreOffice PPTX→PDF 결과를 확인하세요."
                        )
                    else:
                        meta["status"] = "partial"
                        meta["reason"] = (
                            "JPEG이 한 장도 생성되지 않았습니다. "
                            + (
                                render_errors[0]
                                if render_errors
                                else "PyMuPDF JPEG 또는 용량 한도를 확인하세요."
                            )
                        )
                        if render_errors:
                            meta["renderErrorsSample"] = render_errors
                else:
                    meta["status"] = "ok"
                    meta["reason"] = ""
            finally:
                doc.close()
    except subprocess.TimeoutExpired:
        meta["status"] = "error"
        meta["reason"] = f"LibreOffice timeout after {_timeout_sec()}s"
    except Exception as e:
        meta["status"] = "error"
        meta["reason"] = str(e)[:500]

    return out_urls, meta
