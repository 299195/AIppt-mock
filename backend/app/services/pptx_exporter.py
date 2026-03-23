from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict, List

from app.services.template_catalog import resolve_template_assets

_PPTX_GENERATOR_DIR = Path(__file__).resolve().parents[2] / "pptx_generator"
_PPTX_GENERATOR_SCRIPT = _PPTX_GENERATOR_DIR / "generate_deck.js"


def _content_slides(slides: List[Dict]) -> List[Dict]:
    out: List[Dict] = []
    for item in slides:
        title = str(item.get("title") or "")
        slide_type = str(item.get("slide_type") or "").lower()
        if slide_type == "title":
            continue
        low = title.lower()
        if "cover" in low or "agenda" in low:
            continue
        out.append(item)
    return out


def _ensure_pptx_generator_ready() -> None:
    if not _PPTX_GENERATOR_SCRIPT.exists():
        raise RuntimeError(f"pptx-generator script missing: {_PPTX_GENERATOR_SCRIPT}")

    deps_marker = _PPTX_GENERATOR_DIR / "node_modules" / "pptxgenjs"
    if not deps_marker.exists():
        raise RuntimeError(
            "pptx-generator dependencies are not installed. "
            "Run: cd backend/pptx_generator && npm install"
        )


def _default_topic(topic: str, body_slides: List[Dict]) -> str:
    if topic and topic.strip():
        return topic.strip()
    if body_slides:
        first = str(body_slides[0].get("title") or "").strip()
        if first:
            return first.split(" - ", 1)[0].strip() if " - " in first else first
    return "Report"


def export_slides_to_pptx(
    slides: List[Dict],
    out_path: Path,
    template_id: str = "executive_clean",
    topic: str = "",
    outline: List[str] | None = None,
) -> str:
    _ensure_pptx_generator_ready()

    assets = resolve_template_assets(template_id)
    template_pptx_path = assets.get("pptx_path")

    body_slides = _content_slides(slides)
    payload = {
        "topic": _default_topic(topic, body_slides),
        "templateId": template_id,
        "outline": outline[:] if outline else [str(item.get("title") or "") for item in body_slides],
        "slides": slides,
        "templatePptxPath": str(template_pptx_path) if template_pptx_path and template_pptx_path.exists() else None,
    }

    target_path = out_path if out_path.is_absolute() else (Path.cwd() / out_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as tmp:
        json.dump(payload, tmp, ensure_ascii=False)
        payload_path = Path(tmp.name)

    node_bin = os.getenv("PPTX_NODE_BIN", "node")
    cmd = [
        node_bin,
        str(_PPTX_GENERATOR_SCRIPT),
        "--input",
        str(payload_path),
        "--output",
        str(target_path),
    ]

    try:
        completed = subprocess.run(
            cmd,
            cwd=str(_PPTX_GENERATOR_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
    finally:
        payload_path.unlink(missing_ok=True)

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise RuntimeError(f"pptx-generator failed: {detail}")

    if not target_path.exists():
        raise RuntimeError("pptx-generator did not produce output file")

    return target_path.name
