"""
PPTX -> PNG via PowerPoint COM automation (Windows only).

PowerPoint 자체로 렌더링하므로 폰트·테마·효과가 완벽하게 표현됩니다.
LibreOffice 방식과 달리 추가 설치 없이 pywin32 하나면 동작합니다.

Linux/Mac 환경에서는 자동으로 LibreOffice fallback이 사용됩니다.
"""

from __future__ import annotations

import base64
import io
import os
import platform
import tempfile
import zipfile
from pathlib import Path
from typing import Any


def _strip_animations(pptx_bytes: bytes) -> bytes:
    """
    PPTX ZIP 내 각 슬라이드의 <p:timing>...</p:timing> 블록을 정규식으로 제거.
    ElementTree 파싱 없이 원본 XML 바이트를 최대한 보존하여 PPT가 열 수 있게 함.
    """
    import re

    pattern = re.compile(rb"<p:timing\b[^>]*>.*?</p:timing>", re.DOTALL)

    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as src, \
         zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as dst:

        for item in src.infolist():
            data = src.read(item.filename)

            is_slide = (
                item.filename.startswith("ppt/slides/slide")
                and item.filename.endswith(".xml")
                and "slideLayout" not in item.filename
                and "slideMaster" not in item.filename
            )

            if is_slide and b"p:timing" in data:
                data = pattern.sub(b"", data)

            dst.writestr(item, data)

    return buf.getvalue()


def _long_edge() -> int:
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


def _max_bytes() -> int:
    raw = os.getenv("PPTX_RASTER_MAX_BYTES_PER_SLIDE", "2000000").strip()
    try:
        return max(100_000, int(raw))
    except ValueError:
        return 2_000_000


def _png_to_jpeg_bytes(png_path: Path, long_edge: int, quality: int) -> bytes | None:
    """PNG 파일을 JPEG bytes로 변환 (PIL/Pillow 또는 PyMuPDF 사용)."""
    try:
        from PIL import Image
        import io

        with Image.open(png_path) as img:
            # RGBA → RGB 변환
            if img.mode in ("RGBA", "P", "LA"):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode in ("RGBA", "LA"):
                    bg.paste(img, mask=img.split()[-1])
                else:
                    bg.paste(img)
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            # 긴 변 기준 리사이즈
            w, h = img.size
            if max(w, h) > long_edge:
                scale = long_edge / max(w, h)
                img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality, optimize=True)
            return buf.getvalue()
    except ImportError:
        pass

    # Pillow 없으면 PyMuPDF로 변환
    try:
        import fitz
        import io

        doc = fitz.open(str(png_path))
        page = doc.load_page(0)
        w, h = page.rect.width, page.rect.height
        scale = long_edge / max(w, h) if max(w, h) > long_edge else 1.0
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False, colorspace=fitz.csRGB)
        result = pix.tobytes("jpeg", jpg_quality=quality)
        doc.close()
        return result
    except Exception:
        pass

    return None


def render_slide_rasters_ppt(
    pptx_bytes: bytes,
    slide_count: int,
) -> tuple[list[str | None], dict[str, Any]]:
    """
    PowerPoint COM으로 PPTX 슬라이드를 PNG → JPEG base64 data URL로 변환.
    Returns (per-slide data URL or None, meta dict).
    """
    meta: dict[str, Any] = {
        "enabled": True,
        "status": "skipped",
        "reason": "",
        "engine": "powerpoint-com",
        "longEdgePx": _long_edge(),
        "jpegQuality": _jpeg_quality(),
    }

    if slide_count <= 0:
        meta["status"] = "skipped"
        meta["reason"] = "no slides"
        return [], meta

    if platform.system() != "Windows":
        meta["status"] = "skipped"
        meta["reason"] = "PowerPoint COM은 Windows 전용입니다. LibreOffice fallback을 사용하세요."
        return [None] * slide_count, meta

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        meta["status"] = "error"
        meta["reason"] = "pywin32 미설치: pip install pywin32"
        return [None] * slide_count, meta

    out_urls: list[str | None] = [None] * slide_count
    render_errors: list[str] = []
    long_edge = _long_edge()
    quality = _jpeg_quality()
    max_b = _max_bytes()

    try:
        with tempfile.TemporaryDirectory(prefix="pptx_ppt_") as tmp:
            tmp_path = Path(tmp)
            pptx_path = tmp_path / "deck.pptx"
            png_dir = tmp_path / "slides"
            png_dir.mkdir()
            pptx_path.write_bytes(pptx_bytes)

            pythoncom.CoInitialize()
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            # Visible=False는 일부 환경에서 오류 유발 — 최소화 상태로 열기
            ppt_app.WindowState = 2  # ppWindowMinimized

            presentation = None
            try:
                presentation = ppt_app.Presentations.Open(
                    str(pptx_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                )

                n_slides = presentation.Slides.Count
                meta["pptxSlideCount"] = n_slides

                # PPT 슬라이드 크기 기준으로 export 해상도 결정
                slide_w_pt = presentation.PageSetup.SlideWidth   # points
                slide_h_pt = presentation.PageSetup.SlideHeight
                scale = long_edge / max(slide_w_pt, slide_h_pt)
                export_w = int(slide_w_pt * scale)
                export_h = int(slide_h_pt * scale)

                for i in range(1, min(n_slides, slide_count) + 1):
                    png_path = png_dir / f"slide_{i:04d}.png"
                    try:
                        slide = presentation.Slides(i)
                        slide.Export(str(png_path), "PNG", export_w, export_h)
                    except Exception as ex:
                        if len(render_errors) < 3:
                            render_errors.append(f"s{i}: {ex!s}"[:200])
                        continue

                    if not png_path.is_file():
                        if len(render_errors) < 3:
                            render_errors.append(f"s{i}: PNG 파일 생성 안 됨")
                        continue

                    jpeg_bytes = _png_to_jpeg_bytes(png_path, long_edge, quality)
                    if jpeg_bytes is None:
                        if len(render_errors) < 3:
                            render_errors.append(f"s{i}: PNG→JPEG 변환 실패")
                        continue

                    # 용량 초과 시 품질 하향 재시도
                    if len(jpeg_bytes) > max_b:
                        jpeg_bytes = _png_to_jpeg_bytes(
                            png_path,
                            int(long_edge * 0.72),
                            max(55, quality - 12),
                        ) or jpeg_bytes

                    if len(jpeg_bytes) <= max_b:
                        b64 = base64.b64encode(jpeg_bytes).decode("ascii")
                        out_urls[i - 1] = f"data:image/jpeg;base64,{b64}"

            finally:
                if presentation is not None:
                    try:
                        presentation.Close()
                    except Exception:
                        pass
                try:
                    ppt_app.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()

        rendered = sum(1 for u in out_urls if u)
        meta["slidesRendered"] = rendered

        if rendered == 0:
            meta["status"] = "partial" if render_errors else "error"
            meta["reason"] = (
                render_errors[0] if render_errors else "슬라이드 내보내기 실패"
            )
            if render_errors:
                meta["renderErrorsSample"] = render_errors
        else:
            meta["status"] = "ok"
            if render_errors:
                meta["renderErrorsSample"] = render_errors

    except Exception as e:
        meta["status"] = "error"
        meta["reason"] = str(e)[:500]

    return out_urls, meta
