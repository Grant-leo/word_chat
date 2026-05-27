"""DOCX image extraction helpers for content parsing."""
from __future__ import annotations

import hashlib
import os
from typing import Any, Dict, List, Optional, Set


class ImageRegistry:
    """Per-extraction image registry.

    The same DOCX image relationship can be encountered more than once during
    verification or when Word duplicates drawing markup. Saving by a content
    hash avoids the historical bug where one logical figure produced many
    copied files. The registry is local to one extraction, so there is no
    university-, title-, or path-specific hardcoding.
    """

    def __init__(self, fig_dir: str, prefix: str = "img"):
        self.fig_dir = fig_dir
        self.prefix = prefix
        self.counter = 0
        self.by_hash: Dict[str, str] = {}
        self.failures: List[Dict[str, str]] = []

    def save_relationship_image(self, rel: Any) -> Optional[str]:
        if "image" not in getattr(rel, "reltype", ""):
            return None
        try:
            blob = rel.target_part.blob
            digest = hashlib.sha256(blob).hexdigest()[:20]
            if digest in self.by_hash:
                return self.by_hash[digest]

            ext = rel.target_ref.rsplit(".", 1)[-1].lower()
            if ext not in ("png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp"):
                ext = "png"
            self.counter += 1
            fname = f"{self.prefix}_{self.counter:03d}.{ext}"
            fpath = os.path.join(self.fig_dir, fname)
            with open(fpath, "wb") as f:
                f.write(blob)
            self.by_hash[digest] = fname
            return fname
        except Exception as exc:
            self.failures.append({
                "target": getattr(rel, "target_ref", ""),
                "error": str(exc)[:200],
            })
            return None


def extract_images_from_para(para: Any, fig_dir: str, prefix: str = "img", registry: Optional[ImageRegistry] = None) -> List[str]:
    """Extract inline images by relationship id. Returns filenames in paragraph order."""
    saved: List[str] = []
    registry = registry or ImageRegistry(fig_dir, prefix)
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    seen_rids: Set[str] = set()
    for run in para.runs:
        xml = run._element.xml
        if "w:drawing" not in xml and "wp:inline" not in xml and "wp:anchor" not in xml:
            continue
        for blip in run._element.iter(f"{{{a_ns}}}blip"):
            embed = blip.get(f"{{{r_ns}}}embed")
            if not embed or embed in seen_rids:
                continue
            seen_rids.add(embed)
            if embed in para.part.rels:
                fname = registry.save_relationship_image(para.part.rels[embed])
                if fname:
                    saved.append(fname)
            else:
                registry.failures.append({"target": embed, "error": "relationship id not found"})
    return saved


def images_from_run_ooxml(run_elem: Any, rels: Dict[str, Any], registry: ImageRegistry, seen_rids: Set[str], location: str = "body") -> List[Dict[str, Any]]:
    """Extract images from one OOXML run in its exact paragraph position."""
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    v_ns = "urn:schemas-microsoft-com:vml"
    out: List[Dict[str, Any]] = []

    def add_rid(rid: Optional[str]) -> None:
        if not rid or rid in seen_rids:
            return
        seen_rids.add(rid)
        if rid in rels:
            fname = registry.save_relationship_image(rels[rid])
            if fname:
                item: Dict[str, Any] = {"role": "image", "image": fname}
                if location and location != "body":
                    item["location"] = location
                out.append(item)
        else:
            registry.failures.append({"target": rid, "error": "relationship id not found"})

    for blip in run_elem.iter(f"{{{a_ns}}}blip"):
        add_rid(blip.get(f"{{{r_ns}}}embed") or blip.get(f"{{{r_ns}}}link"))
    for imagedata in run_elem.iter(f"{{{v_ns}}}imagedata"):
        add_rid(imagedata.get(f"{{{r_ns}}}id") or imagedata.get(f"{{{r_ns}}}embed"))
    return out


def image_items_from_ooxml(container_elem: Any, rels: Dict[str, Any], registry: ImageRegistry, location: str = "body") -> List[Dict[str, Any]]:
    """Extract all image runs from an arbitrary OOXML container."""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    seen_rids: Set[str] = set()
    out: List[Dict[str, Any]] = []
    for run_elem in container_elem.iter(f"{{{w_ns}}}r"):
        out.extend(images_from_run_ooxml(run_elem, rels, registry, seen_rids, location=location))
    return out


def non_body_image_entries(doc: Any) -> List[Dict[str, str]]:
    """Return header/footer images that are outside the body content stream."""
    entries: List[Dict[str, str]] = []
    seen = set()
    for sec_idx, section in enumerate(doc.sections):
        for attr in (
            "header",
            "first_page_header",
            "even_page_header",
            "footer",
            "first_page_footer",
            "even_page_footer",
        ):
            try:
                part = getattr(section, attr).part
            except Exception:
                continue
            for rid, rel in getattr(part, "rels", {}).items():
                if "image" not in getattr(rel, "reltype", ""):
                    continue
                try:
                    digest = hashlib.sha256(rel.target_part.blob).hexdigest()[:20]
                except Exception:
                    digest = f"{id(part)}:{rid}"
                key = (attr, digest)
                if key in seen:
                    continue
                seen.add(key)
                entries.append({
                    "location": f"section_{sec_idx + 1}_{attr}",
                    "target": getattr(rel, "target_ref", ""),
                })
    return entries
