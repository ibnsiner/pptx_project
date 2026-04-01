import base64
import hashlib
import io
import os
import posixpath
import re
import uuid
import xml.etree.ElementTree as ET
import zipfile
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.shapes.group import GroupShape

from app.slide_raster import render_slide_rasters_jpeg
from app.slide_raster_ppt import render_slide_rasters_ppt

app = FastAPI(title="PPTX Parser API")

PARSER_API_BUILD = "2026-04-01-ppt-com-raster"

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _pct(value: int, total: int) -> str:
    if total <= 0:
        return "0%"
    return f"{(value / total) * 100:.4f}%"


def _shape_type_safe(shape: Any) -> Any:
    try:
        return shape.shape_type
    except (NotImplementedError, ValueError, AttributeError):
        return None


def _iter_shapes_placed(
    shapes: Any,
    offset_left: int = 0,
    offset_top: int = 0,
) -> Any:
    """
    Yield (shape, slide_offset_left, slide_offset_top).
    Child shapes inside a Group have left/top relative to the group origin;
    offsets must be accumulated so HTML % positions match PowerPoint.
    """
    for shape in shapes:
        st = _shape_type_safe(shape)
        if st == MSO_SHAPE_TYPE.GROUP:
            inner = getattr(shape, "shapes", None)
            gl = int(getattr(shape, "left", 0) or 0)
            gt = int(getattr(shape, "top", 0) or 0)
            nl = offset_left + gl
            nt = offset_top + gt
            if inner is not None:
                yield from _iter_shapes_placed(inner, nl, nt)
            elif isinstance(shape, GroupShape):
                yield from _iter_shapes_placed(shape.shapes, nl, nt)
        else:
            yield shape, offset_left, offset_top


def _is_placeholder_shape(shape: Any) -> bool:
    try:
        return bool(getattr(shape, "is_placeholder", False))
    except Exception:
        return False


def _should_skip_master_layout_shape(shape: Any) -> bool:
    """
    마스터/레이아웃에서 대부분의 플레이스홀더는 슬라이드에 복제되므로 건너뛴다.
    그림·클립아트·슬라이드 이미지 플레이스홀더는 슬라이드 XML에 없을 수 있어 포함한다.
    """
    if not _is_placeholder_shape(shape):
        return False
    try:
        t = shape.placeholder_format.type
        if t in (
            PP_PLACEHOLDER.PICTURE,
            PP_PLACEHOLDER.BITMAP,
            PP_PLACEHOLDER.SLIDE_IMAGE,
            PP_PLACEHOLDER.TABLE,
        ):
            return False
    except Exception:
        pass
    return True


def _iter_shapes_paint_order(slide: Any) -> Any:
    """
    Bottom-to-top painting order (matches typical PPT stacking).
    Shapes that exist only on slide master / layout (not copied to slide XML)
    are invisible to slide.shapes; include non-placeholder shapes from those parts.
    Placeholders are skipped on master/layout because the slide carries filled copies.
    """
    try:
        layout = slide.slide_layout
        master = layout.slide_master
    except Exception:
        yield from _iter_shapes_placed(slide.shapes)
        return

    for tree in (master.shapes, layout.shapes):
        for shape, gl, gt in _iter_shapes_placed(tree):
            if not _should_skip_master_layout_shape(shape):
                yield shape, gl, gt

    yield from _iter_shapes_placed(slide.shapes)


def _upload_image_if_configured(
    content: bytes,
    content_type: str,
    slide_num: int,
    index: int,
) -> str | None:
    import os

    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    bucket = os.getenv("STORAGE_BUCKET", "lecture-assets")
    if not url or not key:
        return None
    try:
        from supabase import create_client

        client = create_client(url, key)
        ext = "png"
        if "jpeg" in content_type or "jpg" in content_type:
            ext = "jpg"
        path = f"draft/{uuid.uuid4().hex}/s{slide_num}_{index}.{ext}"
        client.storage.from_(bucket).upload(
            path,
            content,
            file_options={"content-type": content_type or "application/octet-stream"},
        )
        pub = client.storage.from_(bucket).get_public_url(path)
        return pub
    except Exception:
        return None


def _blip_rel_ids(shape_elm: Any) -> list[str]:
    """Collect r:embed and r:link from a:blip (linked or embedded pictures)."""
    seen: list[str] = []
    found: set[str] = set()
    q_embed = qn("r:embed")
    q_link = qn("r:link")
    for node in shape_elm.iter():
        if node.tag != qn("a:blip"):
            continue
        for attr in (q_embed, q_link):
            rid = node.get(attr)
            if rid and rid not in found:
                found.add(rid)
                seen.append(rid)
    return seen


def _load_related_image_any(part: Any, slide_part: Any, r_id: str) -> tuple[bytes | None, str | None]:
    for p in (part, slide_part):
        if p is None:
            continue
        try:
            related = p.related_part(r_id)
            blob = related.blob
            ct = (getattr(related, "content_type", None) or "").strip()
            return blob, ct or None
        except Exception:
            continue
    return None, None


def _looks_like_raster(blob: bytes | None, content_type: str | None) -> bool:
    if not blob or len(blob) < 4:
        return False
    if content_type and content_type.lower().startswith("image/"):
        return True
    if blob.startswith(b"\xff\xd8\xff"):
        return True
    if blob.startswith(b"\x89PNG\r\n\x1a\n"):
        return True
    if blob[:4] in (b"GIF8", b"RIFF"):
        return True
    if blob.startswith(b"BM"):
        return True
    if len(blob) >= 12 and blob[:4] == b"RIFF" and blob[8:12] == b"WEBP":
        return True
    # EMF: 첫 레코드 EMR_HEADER(type=1, size>=88), placeable WMF
    if len(blob) >= 8:
        rtype = int.from_bytes(blob[0:4], "little")
        rsize = int.from_bytes(blob[4:8], "little")
        if rtype == 1 and 88 <= rsize <= len(blob):
            return True
    if blob.startswith(b"\xd7\xcd\xc6\x9a"):
        return True
    return False


def _table_text_from_tbl_xml(shape_elm: Any) -> str:
    """
    a:tbl에서 텍스트 추출.
    a:tr / a:tc 구조가 있으면 행·열 순서 유지, 없으면 a:t 순서 폴백.
    """
    q_tbl = qn("a:tbl")
    q_tr = qn("a:tr")
    q_tc = qn("a:tc")
    q_t = qn("a:t")
    tbl_blocks: list[str] = []

    for tbl in shape_elm.iter():
        if tbl.tag != q_tbl:
            continue
        trs = tbl.findall(q_tr)
        if trs:
            row_strs: list[str] = []
            for tr in trs:
                cells: list[str] = []
                for tc in tr.findall(q_tc):
                    chunks: list[str] = []
                    for node in tc.iter():
                        if node.tag == q_t and node.text:
                            chunks.append(node.text)
                    cells.append("".join(chunks).strip())
                row_strs.append(" | ".join(cells))
            tbl_blocks.append("\n".join(row_strs))
        else:
            flat: list[str] = []
            for node in tbl.iter():
                if node.tag == q_t and node.text and node.text.strip():
                    flat.append(node.text.strip())
            if flat:
                tbl_blocks.append("\n".join(flat))
    return "\n\n".join(tbl_blocks)


def _text_outside_tbl(shape_elm: Any) -> str:
    skip: set[Any] = set()
    for tbl in shape_elm.iter():
        if tbl.tag == qn("a:tbl"):
            skip.update(tbl.iter())
    parts: list[str] = []
    for node in shape_elm.iter():
        if node in skip:
            continue
        if node.tag == qn("a:t") and node.text:
            parts.append(node.text)
    return "\n".join(parts)


def _inline_image_max_bytes() -> int:
    raw = os.getenv("PPTX_IMAGE_INLINE_MAX_BYTES", "").strip()
    if raw.isdigit():
        return max(1_000_000, int(raw))
    return 12_000_000


def _blob_to_data_uri(blob: bytes, content_type: str, max_bytes: int | None = None) -> str | None:
    lim = max_bytes if max_bytes is not None else _inline_image_max_bytes()
    if not blob or len(blob) > lim:
        return None
    mime = (content_type or "image/png").split(";")[0].strip() or "image/png"
    if not mime.startswith("image/"):
        mime = "image/png"
    b64 = base64.b64encode(blob).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _table_text(shape: Any) -> str:
    rows_out: list[str] = []
    try:
        if not getattr(shape, "has_table", False):
            return ""
        for row in shape.table.rows:  # type: ignore[union-attr]
            cells: list[str] = []
            for cell in row.cells:
                try:
                    raw = (cell.text_frame.text or "").strip()
                    if not raw:
                        raw = (getattr(cell, "text", "") or "").strip()
                    cells.append(raw)
                except Exception:
                    try:
                        cells.append((getattr(cell, "text", "") or "").strip())
                    except Exception:
                        cells.append("")
            rows_out.append(" | ".join(cells))
    except Exception:
        try:
            el = getattr(shape, "element", None)
            if el is not None:
                return _table_text_from_tbl_xml(el)
        except Exception:
            pass
        return ""
    return "\n".join(rows_out)


def _shape_text(shape: Any) -> str:
    if getattr(shape, "has_table", False):
        t = _table_text(shape)
        if t.strip():
            return t

    t_tbl = _table_text_from_tbl_xml(shape.element)
    if t_tbl.strip():
        return t_tbl

    if getattr(shape, "has_text_frame", False):
        try:
            t = shape.text or ""
            if t.strip():
                return t
        except Exception:
            pass

    t2 = _text_outside_tbl(shape.element)
    if t2.strip():
        return t2

    return ""


def _pic_geom_emu(pic_elm: Any) -> tuple[int, int, int, int] | None:
    """p:pic 아래 p:spPr/a:xfrm 에서 위치·크기(EMU)."""
    sp_pr = pic_elm.find(qn("p:spPr"))
    if sp_pr is None:
        return None
    xfrm = sp_pr.find(qn("a:xfrm"))
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    try:
        x = int(off.get("x", 0))
        y = int(off.get("y", 0))
        cx = int(ext.get("cx", 0))
        cy = int(ext.get("cy", 0))
        if cx <= 0 or cy <= 0:
            return None
        return x, y, cx, cy
    except (TypeError, ValueError):
        return None


def _pic_blip_rel_ids(pic_elm: Any) -> list[str]:
    """p:blipFill 안의 a:blip r:embed / r:link."""
    out: list[str] = []
    found: set[str] = set()
    bf = pic_elm.find(qn("p:blipFill"))
    if bf is None:
        return out
    q_embed = qn("r:embed")
    q_link = qn("r:link")
    for blip in bf.iter():
        if blip.tag != qn("a:blip"):
            continue
        for attr in (q_embed, q_link):
            rid = blip.get(attr)
            if rid and rid not in found:
                found.add(rid)
                out.append(rid)
    return out


def _image_blob_fingerprint(blob: bytes) -> str:
    return hashlib.sha256(blob).hexdigest()


def _append_raster_dims(
    elements: list[dict[str, Any]],
    blob: bytes,
    ct: str,
    abs_l: int,
    abs_t: int,
    w_emu: int,
    h_emu: int,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    """같은 슬라이드에서 동일 바이너리가 여러 경로로 들어올 때 한 번만 표시."""
    if image_fps is not None and blob:
        fp = _image_blob_fingerprint(blob)
        if fp in image_fps:
            return img_idx
        image_fps.add(fp)
    left_pct = _pct(abs_l, slide_w)
    top_pct = _pct(abs_t, slide_h)
    width_pct = _pct(w_emu, slide_w)
    height_pct = _pct(h_emu, slide_h)
    url = _upload_image_if_configured(blob, ct, slide_num, img_idx)
    if not url:
        url = _blob_to_data_uri(blob, ct or "image/png")
    if not url:
        elements.append(
            {
                "type": "image",
                "src": "",
                "skipReason": "too_large_for_inline",
                "byteLength": len(blob),
                "contentType": ct or "",
                "style": {
                    "left": left_pct,
                    "top": top_pct,
                    "width": width_pct,
                    "height": height_pct,
                },
            }
        )
        return img_idx + 1
    elements.append(
        {
            "type": "image",
            "src": url,
            "style": {
                "left": left_pct,
                "top": top_pct,
                "width": width_pct,
                "height": height_pct,
            },
        }
    )
    return img_idx + 1


def _append_raster_element(
    elements: list[dict[str, Any]],
    blob: bytes,
    ct: str,
    abs_l: int,
    abs_t: int,
    shape: Any,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    return _append_raster_dims(
        elements,
        blob,
        ct,
        abs_l,
        abs_t,
        int(shape.width),
        int(shape.height),
        slide_w,
        slide_h,
        slide_num,
        img_idx,
        image_fps,
    )


def _xfrm_off_ext_emu(xfrm: Any) -> tuple[int, int, int, int] | None:
    if xfrm is None:
        return None
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    try:
        x = int(off.get("x", 0))
        y = int(off.get("y", 0))
        cx = int(ext.get("cx", 0))
        cy = int(ext.get("cy", 0))
        if cx <= 0 or cy <= 0:
            return None
        return x, y, cx, cy
    except (TypeError, ValueError):
        return None


def _graphic_frame_geom_emu(gf_elm: Any) -> tuple[int, int, int, int] | None:
    """p:graphicFrame 직속 p:xfrm."""
    xfrm = gf_elm.find(qn("p:xfrm"))
    return _xfrm_off_ext_emu(xfrm)


def _normalized_text_block(s: str) -> str:
    return "\n".join(
        ln.strip() for ln in s.strip().splitlines() if ln.strip()
    )


def _collect_existing_text_keys(elements: list[dict[str, Any]]) -> set[str]:
    keys: set[str] = set()
    for el in elements:
        if el.get("type") != "text":
            continue
        c = (el.get("content") or "").strip()
        if c:
            keys.add(_normalized_text_block(c))
            keys.add(c)
    return keys


def _enrich_tables_from_slide_oxml(
    slide_elm: Any,
    slide_w: int,
    slide_h: int,
    elements: list[dict[str, Any]],
    slide_text_parts: list[str],
) -> None:
    """
    slide.shapes / python-pptx에 안 올라오는 표(p:graphicFrame, mc:AlternateContent 등) 보강.
    """
    q_gf = qn("p:graphicFrame")
    seen = _collect_existing_text_keys(elements)
    for gf in slide_elm.iter():
        if gf.tag != q_gf:
            continue
        t = _table_text_from_tbl_xml(gf)
        if not t.strip():
            continue
        key = _normalized_text_block(t)
        if key in seen:
            continue
        seen.add(key)
        slide_text_parts.append(t.strip())
        geom = _graphic_frame_geom_emu(gf)
        if geom:
            abs_l, abs_t, w_emu, h_emu = geom
        else:
            abs_l, abs_t = 0, 0
            w_emu = max(slide_w // 2, 1)
            h_emu = max(slide_h // 8, 1)
        elements.append(
            {
                "type": "text",
                "content": t,
                "style": {
                    "left": _pct(abs_l, slide_w),
                    "top": _pct(abs_t, slide_h),
                    "width": _pct(w_emu, slide_w),
                    "height": _pct(h_emu, slide_h),
                },
            }
        )


def _blip_under_slide_background(blip: Any) -> bool:
    cur: Any = blip
    for _ in range(40):
        if cur is None:
            return False
        if cur.tag in (qn("p:bg"), qn("p:bgPr")):
            return True
        cur = cur.getparent()
    return False


def _geom_from_blip_ancestor(blip: Any) -> tuple[int, int, int, int] | None:
    """a:blip에서 상위 도형의 xfrm(off/ext) 추적."""
    cur: Any = blip.getparent()
    q_sp = qn("p:sp")
    q_pic = qn("p:pic")
    q_gf = qn("p:graphicFrame")
    q_grp = qn("p:grpSp")
    q_cxn = qn("p:cxnSp")
    while cur is not None:
        tag = cur.tag
        if tag == q_pic:
            g = _pic_geom_emu(cur)
            if g:
                return g
        if tag in (q_sp, q_cxn):
            sp_pr = cur.find(qn("p:spPr"))
            if sp_pr is not None:
                xf = sp_pr.find(qn("a:xfrm"))
                r = _xfrm_off_ext_emu(xf)
                if r:
                    return r
        if tag == q_gf:
            xf = cur.find(qn("p:xfrm"))
            r = _xfrm_off_ext_emu(xf)
            if r:
                return r
        if tag == q_grp:
            gsp = cur.find(qn("p:grpSpPr"))
            if gsp is not None:
                xf = gsp.find(qn("a:xfrm"))
                r = _xfrm_off_ext_emu(xf)
                if r:
                    return r
        cur = cur.getparent()
    return None


_OOXML_REL_EMBED = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
)
_OOXML_REL_LINK = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link"
)
_OOXML_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _xml_local_tag(el: Any) -> str:
    t = el.tag
    return t.rsplit("}", 1)[-1] if isinstance(t, str) and "}" in t else str(t)


def _collect_all_image_rids_from_element_tree(root: Any) -> list[str]:
    """
    a:blip 말고도 VML imagedata·일부 변형의 r:id / r:embed 등 그림 관계 ID 수집.
    (하이퍼링크 등 비이미지 r:id는 로드 후 _looks_like_raster로 걸러짐.)
    """
    out: list[str] = []
    seen: set[str] = set()
    q_embed = qn("r:embed")
    q_link = qn("r:link")
    q_rid = qn("r:id")
    direct_keys = (
        q_embed,
        q_link,
        q_rid,
        _OOXML_REL_EMBED,
        _OOXML_REL_LINK,
        _OOXML_REL_ID,
    )
    for el in root.iter():
        for k in direct_keys:
            v = el.get(k)
            if v and isinstance(v, str) and v.startswith("rId") and v not in seen:
                seen.add(v)
                out.append(v)
        loc = _xml_local_tag(el).lower()
        for ak, av in el.attrib.items():
            if not av or not isinstance(av, str) or not av.startswith("rId"):
                continue
            if av in seen:
                continue
            if ak in direct_keys or ak.endswith("}embed") or ak.endswith("}link"):
                seen.add(av)
                out.append(av)
            elif loc == "imagedata" and (ak.endswith("}id") or ak == q_rid):
                seen.add(av)
                out.append(av)
    return out


def _enrich_images_from_xml_rel_ids(
    part: Any,
    slide_part: Any,
    root_elm: Any,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    used_rids: set[str],
    elements: list[dict[str, Any]],
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    """XML에 등장하는 모든 rId 후보로 이미지 부품 로드(도형 API 누락 보강)."""
    for r_id in _collect_all_image_rids_from_element_tree(root_elm):
        if r_id in used_rids:
            continue
        blob, ct = _load_related_image_any(part, slide_part, r_id)
        if not blob or not _looks_like_raster(blob, ct):
            continue
        used_rids.add(r_id)
        abs_l = abs_t = 0
        w_emu = max(slide_w // 5, 1)
        h_emu = max(slide_h // 5, 1)
        img_idx = _append_raster_dims(
            elements,
            blob,
            ct or "image/png",
            abs_l,
            abs_t,
            w_emu,
            h_emu,
            slide_w,
            slide_h,
            slide_num,
            img_idx,
            image_fps,
        )
    return img_idx


def _merge_slide_xml_text_into_parts(slide_elm: Any, slide_text_parts: list[str]) -> None:
    """slide XML 전체 a:t 및 (drawing이 아닌) wps/w 단락 t — 도형 순회 누락분을 plainText에 합침."""
    seen: set[str] = set()
    for part in slide_text_parts:
        for ln in part.splitlines():
            s = ln.strip()
            if s:
                seen.add(s)
    q_at = qn("a:t")
    extra: list[str] = []
    for el in slide_elm.iter():
        if el.tag != q_at or not el.text:
            continue
        t = el.text.strip()
        if t and t not in seen:
            seen.add(t)
            extra.append(t)
    for el in slide_elm.iter():
        if el.tag == q_at:
            continue
        if _xml_local_tag(el) != "t" or not el.text:
            continue
        t = el.text.strip()
        if not t or t in seen:
            continue
        seen.add(t)
        extra.append(t)
    slide_text_parts.extend(extra)


def _enrich_all_blips_in_tree(
    part: Any,
    slide_part: Any,
    root_elm: Any,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    used_rids: set[str],
    elements: list[dict[str, Any]],
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    """
    p:pic 이외 blipFill(자동 도형·배경 등)의 a:blip까지 수집.
    관계는 주로 part에 있으나, 슬라이드 쪽 관계도 시도.
    """
    q_blip = qn("a:blip")
    q_embed = qn("r:embed")
    q_link = qn("r:link")
    for blip in root_elm.iter():
        if blip.tag != q_blip:
            continue
        for attr in (q_embed, q_link):
            r_id = blip.get(attr)
            if not r_id or r_id in used_rids:
                continue
            blob, ct = _load_related_image_any(part, slide_part, r_id)
            if not blob or not _looks_like_raster(blob, ct):
                continue
            used_rids.add(r_id)
            if _blip_under_slide_background(blip):
                abs_l, abs_t, w_emu, h_emu = 0, 0, slide_w, slide_h
            else:
                geom = _geom_from_blip_ancestor(blip)
                if geom is None:
                    abs_l, abs_t = 0, 0
                    w_emu = max(slide_w // 6, 1)
                    h_emu = max(slide_h // 6, 1)
                else:
                    abs_l, abs_t, w_emu, h_emu = geom
            img_idx = _append_raster_dims(
                elements,
                blob,
                ct or "image/png",
                abs_l,
                abs_t,
                w_emu,
                h_emu,
                slide_w,
                slide_h,
                slide_num,
                img_idx,
                image_fps,
            )
    return img_idx


def _enrich_slide_orphan_image_rels(
    slide_part: Any,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    used_rids: set[str],
    elements: list[dict[str, Any]],
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    """슬라이드 .rels에만 있는 이미지(비표준 XML 등) 보강. 레이아웃/마스터는 호출하지 않음."""
    try:
        rels = getattr(slide_part, "rels", None)
        if rels is None:
            return img_idx
        for rel in rels.values():
            rt = getattr(rel, "reltype", "") or ""
            if "relationships/image" not in rt:
                continue
            rid = getattr(rel, "rId", None)
            if not rid or rid in used_rids:
                continue
            try:
                tp = rel.target_part
                blob = getattr(tp, "blob", None)
                ct = (getattr(tp, "content_type", None) or "").strip() or None
            except Exception:
                continue
            if not blob or not _looks_like_raster(blob, ct):
                continue
            used_rids.add(rid)
            img_idx = _append_raster_dims(
                elements,
                blob,
                ct or "image/png",
                0,
                0,
                max(slide_w // 4, 1),
                max(slide_h // 4, 1),
                slide_w,
                slide_h,
                slide_num,
                img_idx,
                image_fps,
            )
    except Exception:
        pass
    return img_idx


def _enrich_oxml_pics(
    part: Any,
    slide_part: Any,
    root_elm: Any,
    slide_w: int,
    slide_h: int,
    slide_num: int,
    used_rids: set[str],
    elements: list[dict[str, Any]],
    img_idx: int,
    image_fps: set[str] | None = None,
) -> int:
    """
    python-pptx shape 트리 밖(mc:AlternateContent 등)의 p:pic 보강.
    part.related_part으로 해당 rId 해석.
    """
    for pic in root_elm.iter():
        if pic.tag != qn("p:pic"):
            continue
        for r_id in _pic_blip_rel_ids(pic):
            if r_id in used_rids:
                continue
            blob, ct = _load_related_image_any(part, slide_part, r_id)
            if not blob or not _looks_like_raster(blob, ct):
                continue
            used_rids.add(r_id)
            geom = _pic_geom_emu(pic)
            if geom is None:
                abs_l = abs_t = 0
                w_emu = max(slide_w // 20, 1)
                h_emu = max(slide_h // 20, 1)
            else:
                abs_l, abs_t, w_emu, h_emu = geom
            img_idx = _append_raster_dims(
                elements,
                blob,
                ct or "image/png",
                abs_l,
                abs_t,
                w_emu,
                h_emu,
                slide_w,
                slide_h,
                slide_num,
                img_idx,
                image_fps,
            )
    return img_idx


def _content_type_from_zip_path(name: str) -> str:
    n = name.lower()
    if n.endswith(".png"):
        return "image/png"
    if n.endswith(".jpg") or n.endswith(".jpeg"):
        return "image/jpeg"
    if n.endswith(".gif"):
        return "image/gif"
    if n.endswith(".webp"):
        return "image/webp"
    if n.endswith(".bmp"):
        return "image/bmp"
    if n.endswith(".tif") or n.endswith(".tiff"):
        return "image/tiff"
    if n.endswith(".emf"):
        return "image/x-emf"
    if n.endswith(".wmf"):
        return "image/x-wmf"
    if n.endswith(".svg"):
        return "image/svg+xml"
    return "application/octet-stream"


def _zip_resolve_part_target(part_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/").replace("\\", "/")
    return posixpath.normpath(posixpath.join(part_dir, target)).replace("\\", "/")


def _zip_iter_image_targets_from_rels(
    zf: zipfile.ZipFile,
    rels_internal_path: str,
    part_dir_for_targets: str,
) -> list[str]:
    """rels 파일에서 내부 이미지 타깃 경로(zip 내) 목록."""
    if rels_internal_path not in zf.namelist():
        return []
    out: list[str] = []
    try:
        rels_root = ET.fromstring(zf.read(rels_internal_path))
        for el in rels_root:
            loc = el.tag.split("}")[-1] if "}" in el.tag else el.tag
            if loc != "Relationship":
                continue
            typ = (el.get("Type") or "").lower()
            if "relationships/image" not in typ and "/image" not in typ:
                continue
            if el.get("TargetMode") == "External":
                continue
            target = el.get("Target")
            if not target:
                continue
            zip_path = _zip_resolve_part_target(part_dir_for_targets, target)
            if zip_path in zf.namelist():
                out.append(zip_path)
    except Exception:
        pass
    return out


def _slide_zip_package_enrich(
    zf: zipfile.ZipFile,
    slide_num: int,
    slide_w: int,
    slide_h: int,
    elements: list[dict[str, Any]],
    slide_text_parts: list[str],
    img_idx: int,
    image_fps: set[str] | None = None,
) -> tuple[int, dict[str, Any]]:
    """
    OOXML ZIP에서 slideN 관계·레이아웃·마스터 rels의 image Target까지 따라가 media 로드.
    """
    stats: dict[str, Any] = {
        "slideXmlInZip": False,
        "slideRelsInZip": False,
        "zipImageRelCount": 0,
        "zipImagesEmitted": 0,
        "zipLayoutMasterPartsFollowed": 0,
        "zipTextFragmentsAdded": 0,
    }
    slide_xml = f"ppt/slides/slide{slide_num}.xml"
    rels_xml = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
    base_dir = "ppt/slides"
    emitted_paths: set[str] = set()

    def try_emit_zip_paths(paths: list[str]) -> None:
        nonlocal img_idx
        for zip_path in paths:
            if zip_path in emitted_paths or zip_path not in zf.namelist():
                continue
            try:
                blob = zf.read(zip_path)
            except Exception:
                continue
            ct = _content_type_from_zip_path(zip_path)
            if not _looks_like_raster(blob, ct):
                continue
            emitted_paths.add(zip_path)
            prev_idx = img_idx
            img_idx = _append_raster_dims(
                elements,
                blob,
                ct,
                0,
                0,
                max(slide_w // 4, 1),
                max(slide_h // 4, 1),
                slide_w,
                slide_h,
                slide_num,
                img_idx,
                image_fps,
            )
            if img_idx > prev_idx:
                stats["zipImagesEmitted"] += 1

    if slide_xml in zf.namelist():
        stats["slideXmlInZip"] = True
    if rels_xml in zf.namelist():
        stats["slideRelsInZip"] = True
        try:
            rels_root = ET.fromstring(zf.read(rels_xml))
            slide_image_targets: list[str] = []
            layout_master_xmls: list[str] = []
            for el in rels_root:
                loc = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if loc != "Relationship":
                    continue
                typ = (el.get("Type") or "").lower()
                target = el.get("Target")
                if not target or el.get("TargetMode") == "External":
                    continue
                if "relationships/image" in typ or "/image" in typ:
                    stats["zipImageRelCount"] += 1
                    zp = _zip_resolve_part_target(base_dir, target)
                    if zp in zf.namelist():
                        slide_image_targets.append(zp)
                elif "slidelayout" in typ or "slidemaster" in typ:
                    lpath = _zip_resolve_part_target(base_dir, target)
                    if lpath in zf.namelist() and lpath.endswith(".xml"):
                        layout_master_xmls.append(lpath)

            try_emit_zip_paths(slide_image_targets)

            seen_lm: set[str] = set()
            for lm_xml in layout_master_xmls:
                if lm_xml in seen_lm:
                    continue
                seen_lm.add(lm_xml)
                stats["zipLayoutMasterPartsFollowed"] += 1
                lm_dir = posixpath.dirname(lm_xml)
                lm_base = posixpath.basename(lm_xml)
                lm_rels = posixpath.join(lm_dir, "_rels", lm_base + ".rels")
                nested = _zip_iter_image_targets_from_rels(zf, lm_rels, lm_dir)
                stats["zipImageRelCount"] += len(nested)
                try_emit_zip_paths(nested)
        except Exception:
            pass

    if slide_xml in zf.namelist():
        try:
            sroot = ET.fromstring(zf.read(slide_xml))
            seen: set[str] = set()
            for part in slide_text_parts:
                for ln in part.splitlines():
                    s = ln.strip()
                    if s:
                        seen.add(s)
            drawing_t = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
            extra: list[str] = []
            for el in sroot.iter():
                if el.text is None:
                    continue
                t = el.text.strip()
                if not t or t in seen:
                    continue
                tag = el.tag
                if tag == drawing_t:
                    seen.add(t)
                    extra.append(t)
                    continue
                local = tag.split("}")[-1] if "}" in tag else tag
                if local == "t" and "wordprocessing" in tag.lower():
                    seen.add(t)
                    extra.append(t)
            stats["zipTextFragmentsAdded"] = len(extra)
            slide_text_parts.extend(extra)
        except Exception:
            pass

    return img_idx, stats


def parse_presentation(content: bytes) -> dict[str, Any]:
    prs = Presentation(io.BytesIO(content))
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    all_plain: list[str] = []
    slides_out: list[dict[str, Any]] = []

    zf: zipfile.ZipFile | None = None
    try:
        zf = zipfile.ZipFile(io.BytesIO(content))
    except Exception:
        zf = None

    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        elements: list[dict[str, Any]] = []
        slide_text_parts: list[str] = []
        img_idx = 0
        image_fps: set[str] = set()

        slide_part = slide.part
        used_blip_rids: set[str] = set()

        for shape, grp_l, grp_t in _iter_shapes_paint_order(slide):
            abs_l = int(shape.left) + grp_l
            abs_t = int(shape.top) + grp_t

            st = _shape_type_safe(shape)
            skip_blip_after_picture = False
            if st in (
                MSO_SHAPE_TYPE.PICTURE,
                MSO_SHAPE_TYPE.LINKED_PICTURE,
            ):
                blob: bytes | None = None
                ct: str | None = None
                try:
                    blob = shape.image.blob  # type: ignore[attr-defined]
                    ct = shape.image.content_type  # type: ignore[attr-defined]
                except Exception:
                    blob, ct = None, None
                if blob and _looks_like_raster(blob, ct):
                    for rid in _blip_rel_ids(shape.element):
                        used_blip_rids.add(rid)
                    img_idx = _append_raster_element(
                        elements,
                        blob,
                        ct or "image/png",
                        abs_l,
                        abs_t,
                        shape,
                        slide_w,
                        slide_h,
                        slide_num,
                        img_idx,
                        image_fps,
                    )
                    skip_blip_after_picture = True

            if not skip_blip_after_picture:
                seen_rid: set[str] = set()
                for r_id in _blip_rel_ids(shape.element):
                    if r_id in seen_rid:
                        continue
                    seen_rid.add(r_id)
                    blob, ct = _load_related_image_any(shape.part, slide_part, r_id)
                    if blob and _looks_like_raster(blob, ct):
                        used_blip_rids.add(r_id)
                        img_idx = _append_raster_element(
                            elements,
                            blob,
                            ct or "image/png",
                            abs_l,
                            abs_t,
                            shape,
                            slide_w,
                            slide_h,
                            slide_num,
                            img_idx,
                            image_fps,
                        )

            txt = _shape_text(shape)
            if txt.strip():
                slide_text_parts.append(txt.strip())
            if txt.strip():
                left_pct = _pct(abs_l, slide_w)
                top_pct = _pct(abs_t, slide_h)
                width_pct = _pct(shape.width, slide_w)
                height_pct = _pct(shape.height, slide_h)
                font_size = None
                try:
                    tf = getattr(shape, "text_frame", None)
                    if tf is not None and tf.paragraphs:
                        run = tf.paragraphs[0].runs
                        if run and run[0].font.size:
                            pt = run[0].font.size.pt
                            if pt and slide_w > 0:
                                # 슬라이드 너비(pt) 대비 글꼴 pt → CSS cqw (미리보기 박스 기준)
                                slide_w_pt = float(slide_w) / 12700.0
                                if slide_w_pt > 0:
                                    cqw = (float(pt) / slide_w_pt) * 100.0
                                    cqw = min(max(cqw, 0.4), 14.0)
                                    font_size = f"{cqw:.4f}cqw"
                except Exception:
                    pass
                style = {
                    "left": left_pct,
                    "top": top_pct,
                    "width": width_pct,
                    "height": height_pct,
                }
                if font_size:
                    style["fontSize"] = font_size
                elements.append(
                    {
                        "type": "text",
                        "content": txt,
                        "style": style,
                    }
                )

        try:
            lo = slide.slide_layout
            ma = lo.slide_master
            oxml_roots: tuple[tuple[Any, Any], ...] = (
                (slide_part, slide._element),
                (lo.part, lo._element),
                (ma.part, ma._element),
            )
        except Exception:
            oxml_roots = ((slide_part, slide._element),)
        for oxml_part, root_elm in oxml_roots:
            img_idx = _enrich_oxml_pics(
                oxml_part,
                slide_part,
                root_elm,
                slide_w,
                slide_h,
                slide_num,
                used_blip_rids,
                elements,
                img_idx,
                image_fps,
            )
        for oxml_part, root_elm in oxml_roots:
            img_idx = _enrich_all_blips_in_tree(
                oxml_part,
                slide_part,
                root_elm,
                slide_w,
                slide_h,
                slide_num,
                used_blip_rids,
                elements,
                img_idx,
                image_fps,
            )
        img_idx = _enrich_images_from_xml_rel_ids(
            slide_part,
            slide_part,
            slide._element,
            slide_w,
            slide_h,
            slide_num,
            used_blip_rids,
            elements,
            img_idx,
            image_fps,
        )
        img_idx = _enrich_slide_orphan_image_rels(
            slide_part,
            slide_w,
            slide_h,
            slide_num,
            used_blip_rids,
            elements,
            img_idx,
            image_fps,
        )

        _enrich_tables_from_slide_oxml(
            slide._element,
            slide_w,
            slide_h,
            elements,
            slide_text_parts,
        )

        _merge_slide_xml_text_into_parts(slide._element, slide_text_parts)

        extract_stats: dict[str, Any] = {
            "zipPackageOk": zf is not None,
            "parserApiBuild": PARSER_API_BUILD,
        }
        if zf is not None:
            img_idx, zst = _slide_zip_package_enrich(
                zf,
                slide_num,
                slide_w,
                slide_h,
                elements,
                slide_text_parts,
                img_idx,
                image_fps,
            )
            extract_stats.update(zst)

        plain_slide = "\n".join(slide_text_parts)
        if plain_slide:
            all_plain.append(plain_slide)
        slides_out.append(
            {
                "slideNumber": slide_num,
                "elements": elements,
                "plainText": plain_slide,
                "extractStats": extract_stats,
            }
        )

    if zf is not None:
        try:
            zf.close()
        except Exception:
            pass

    full_plain = "\n\n".join(all_plain)
    words = re.findall(r"[\w가-힣]+", full_plain)
    freq: dict[str, int] = {}
    for w in words:
        if len(w) < 2:
            continue
        freq[w.lower()] = freq.get(w.lower(), 0) + 1
    top_tags = sorted(freq, key=lambda k: freq[k], reverse=True)[:5]

    title_guess = ""
    if slides_out and slides_out[0].get("elements"):
        for el in slides_out[0]["elements"]:
            if el.get("type") == "text" and el.get("content"):
                title_guess = (el["content"] or "").split("\n")[0].strip()[:120]
                break

    return {
        "slides": slides_out,
        "plainText": full_plain,
        "meta": {
            "title": title_guess or "Untitled lecture",
            "description": "",
            "tags": top_tags,
            "parserApiBuild": PARSER_API_BUILD,
            "slideSizeEmu": {
                "width": int(prs.slide_width),
                "height": int(prs.slide_height),
            },
        },
    }


@app.post("/api/parse-pptx")
async def parse_pptx(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Expected .pptx file")
    content = await file.read()
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="Empty file")
    try:
        payload = parse_presentation(content)
        slides = payload.get("slides") or []

        import asyncio, logging
        logger = logging.getLogger("pptx_raster")

        # PowerPoint COM은 STA COM이라 async 이벤트 루프 스레드에서 직접 호출하면
        # 아파트먼트 충돌이 발생함 → 전용 스레드에서 실행
        raster_urls, raster_meta = await asyncio.to_thread(
            render_slide_rasters_ppt, content, len(slides)
        )
        ppt_status = raster_meta.get("status")
        ppt_reason = raster_meta.get("reason", "")
        logger.warning("[raster] ppt-com status=%s reason=%s", ppt_status, ppt_reason)

        if ppt_status not in ("ok", "partial"):
            # PPT COM 실패 이유를 meta에 보존하고 LO로 재시도
            lo_urls, lo_meta = await asyncio.to_thread(
                render_slide_rasters_jpeg, content, len(slides)
            )
            lo_meta["pptComFallbackReason"] = ppt_reason or "ppt-com status=" + str(ppt_status)
            lo_meta["engine"] = "libreoffice"
            raster_urls, raster_meta = lo_urls, lo_meta
        
        for slide_dict, url in zip(slides, raster_urls):
            if url:
                slide_dict["rasterPreview"] = url
        meta = payload.setdefault("meta", {})
        meta["slideRaster"] = raster_meta
        return JSONResponse(
            content=payload,
            headers={"X-PPTX-Parser-Build": PARSER_API_BUILD},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@app.get("/health")
def health():
    # 브라우저가 JSON을 오래 캐시하면 parserApiBuild 갱신이 안 보일 수 있음
    return JSONResponse(
        content={"status": "ok", "parserApiBuild": PARSER_API_BUILD},
        headers={"Cache-Control": "no-store, max-age=0"},
    )


@app.get("/pptx-parser/build")
def parser_build_info():
    """같은 호스트에 다른 앱이 있을 때 /health만으로 헷갈리면 이 경로로 판별."""
    return JSONResponse(
        content={
            "parserApiBuild": PARSER_API_BUILD,
            "mainFile": __file__,
        },
        headers={"Cache-Control": "no-store, max-age=0"},
    )
