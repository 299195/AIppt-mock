from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Dict, List, Optional


ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets"
CUSTOM_TEMPLATE_DIR = ASSETS_DIR / "custom_templates"
CUSTOM_BG_DIR = CUSTOM_TEMPLATE_DIR / "bgs"
CUSTOM_PREVIEW_DIR = ASSETS_DIR / "custom_template_previews"


def _slugify(value: str) -> str:
    normalized = re.sub(r"[^a-zA-Z0-9]+", "_", value.strip().lower())
    slug = normalized.strip("_")
    if slug:
        return slug
    digest = hashlib.md5(value.encode("utf-8")).hexdigest()[:10]
    return f"tpl_{digest}"


def _hex_color(rgb: tuple[int, int, int]) -> str:
    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"


def _derived_preview_colors(template_id: str) -> Dict[str, str]:
    digest = hashlib.md5(template_id.encode("utf-8")).digest()
    fg = (25 + digest[0] % 55, 45 + digest[1] % 65, 70 + digest[2] % 85)
    accent = (60 + digest[3] % 125, 95 + digest[4] % 130, 130 + digest[5] % 120)
    bg = (
        min(245, 220 + digest[6] % 35),
        min(248, 224 + digest[7] % 30),
        min(252, 228 + digest[8] % 24),
    )
    return {
        "preview_bg": _hex_color(bg),
        "preview_fg": _hex_color(fg),
        "preview_accent": _hex_color(accent),
    }


def _custom_bg_for(stem: str) -> Optional[Path]:
    candidates = [
        CUSTOM_TEMPLATE_DIR / f"{stem}.png",
        CUSTOM_TEMPLATE_DIR / f"{stem}.jpg",
        CUSTOM_TEMPLATE_DIR / f"{stem}.jpeg",
        CUSTOM_BG_DIR / f"{stem}.png",
        CUSTOM_BG_DIR / f"{stem}.jpg",
        CUSTOM_BG_DIR / f"{stem}.jpeg",
    ]
    for item in candidates:
        if item.exists():
            return item
    return None


def _preview_image_url_for(stem: str) -> Optional[str]:
    candidates = [
        CUSTOM_PREVIEW_DIR / f"{stem}.png",
        CUSTOM_PREVIEW_DIR / f"{stem}.jpg",
        CUSTOM_PREVIEW_DIR / f"{stem}.jpeg",
        CUSTOM_PREVIEW_DIR / f"{stem}.webp",
    ]
    for item in candidates:
        if item.exists():
            return f"/assets/custom_template_previews/{item.name}"
    return None


def _custom_template_index() -> Dict[str, Dict[str, object]]:
    CUSTOM_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    CUSTOM_PREVIEW_DIR.mkdir(parents=True, exist_ok=True)

    used_ids: set[str] = set()
    index: Dict[str, Dict[str, object]] = {}

    for pptx in sorted(CUSTOM_TEMPLATE_DIR.glob("*.pptx")):
        stem = pptx.stem
        base_id = f"custom_{_slugify(stem)}"
        template_id = base_id

        if template_id in used_ids:
            suffix = hashlib.md5(stem.encode("utf-8")).hexdigest()[:6]
            template_id = f"{base_id}_{suffix}"

        used_ids.add(template_id)

        colors = _derived_preview_colors(template_id)
        preview_image_url = _preview_image_url_for(stem)
        bg_path = _custom_bg_for(stem)

        index[template_id] = {
            "id": template_id,
            "name": stem.replace("_", " ").strip() or template_id,
            "subtitle": "Custom",
            "summary": "User-managed template from custom_templates.",
            "preview_image_url": preview_image_url,
            "pptx_path": pptx,
            "bg_path": bg_path,
            **colors,
        }

    return index


def list_templates() -> List[Dict[str, object]]:
    custom_index = _custom_template_index()
    custom_items = [
        {k: v for k, v in item.items() if k not in {"pptx_path", "bg_path"}}
        for item in custom_index.values()
    ]
    return [dict(item) for item in custom_items]


def template_exists(template_id: str) -> bool:
    return template_id in _custom_template_index()


def resolve_template_assets(template_id: str) -> Dict[str, Optional[Path]]:
    item = _custom_template_index().get(template_id)
    if not item:
        return {"pptx_path": None, "bg_path": None}
    return {
        "pptx_path": item.get("pptx_path"),
        "bg_path": item.get("bg_path"),
    }
