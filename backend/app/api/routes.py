from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import List
from uuid import uuid4

from fastapi import APIRouter, File, HTTPException, UploadFile

from app.config import settings
from app.graph.workflow import llm, run_generation
from app.schemas import (
    GenerateRequest,
    GenerateResponse,
    HistoryItem,
    JobDetailResponse,
    ModelConfigResponse,
    OutlinePreviewRequest,
    OutlinePreviewResponse,
    RewriteRequest,
    SlideDTO,
    TemplateItem,
    UploadParseResponse,
)
from app.services.parser import parse_text_input, read_uploaded_file
from app.services.template_catalog import list_templates, template_exists
from app.storage.db import get_job, list_jobs

router = APIRouter()


@router.get("/model/config", response_model=ModelConfigResponse)
def model_config() -> ModelConfigResponse:
    return ModelConfigResponse(
        provider=settings.model_provider,
        model=settings.model_name,
        use_mock=settings.use_mock_llm,
        configured=bool(settings.model_base_url and settings.model_api_key and settings.model_name),
        base_url=settings.model_base_url,
    )


@router.get("/templates", response_model=List[TemplateItem])
def templates() -> List[TemplateItem]:
    return [TemplateItem(**item) for item in list_templates()]


@router.post("/parse-upload", response_model=UploadParseResponse)
async def parse_upload(file: UploadFile = File(...)) -> UploadParseResponse:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".md", ".docx"}:
        raise HTTPException(status_code=400, detail="仅支持 .md / .docx")

    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)

    try:
        extracted = read_uploaded_file(tmp_path)
        return UploadParseResponse(extracted_text=extracted)
    finally:
        tmp_path.unlink(missing_ok=True)


@router.post("/outline/preview", response_model=OutlinePreviewResponse)
def preview_outline(req: OutlinePreviewRequest) -> OutlinePreviewResponse:
    material = parse_text_input(req.title, req.outline_text, req.material_text)
    outline = llm.generate_outline(req.title, req.style, material, req.target_pages)
    return OutlinePreviewResponse(outline=outline)


@router.post("/jobs", response_model=GenerateResponse)
def create_job(req: GenerateRequest) -> GenerateResponse:
    if not template_exists(req.template_id):
        raise HTTPException(status_code=400, detail="模板不存在")

    job_id = str(uuid4())
    material = parse_text_input(req.title, req.outline_text, req.material_text)

    outline = [x.strip() for x in (req.outline or []) if x and x.strip()]
    result = run_generation(
        {
            "job_id": job_id,
            "title": req.title,
            "style": req.style,
            "template_id": req.template_id,
            "target_pages": req.target_pages,
            "material": material,
            "outline": outline if outline else None,
            "rewrite_action": "",
            "created_at": datetime.utcnow().isoformat(),
        }
    )
    return GenerateResponse(job_id=result["job_id"])


@router.get("/jobs/{job_id}", response_model=JobDetailResponse)
def job_detail(job_id: str) -> JobDetailResponse:
    row = get_job(job_id)
    if not row:
        raise HTTPException(status_code=404, detail="任务不存在")

    slides_raw = json.loads(row["slides_json"])
    slides = [
        SlideDTO(
            page=item["page"],
            title=item["title"],
            bullets=item["bullets"],
            notes=item.get("notes", ""),
            slide_type=item.get("slide_type"),
            evidence=item.get("evidence"),
        )
        for item in slides_raw
    ]

    return JobDetailResponse(
        job_id=row["job_id"],
        status=row["status"],
        style=row["style"],
        template_id=row["template_id"] if "template_id" in row.keys() else "executive_clean",
        title=row["title"],
        outline=json.loads(row["outline_json"]),
        slides=slides,
        pptx_url=row["pptx_url"],
        created_at=datetime.fromisoformat(row["created_at"]),
    )


@router.get("/jobs", response_model=List[HistoryItem])
def history() -> List[HistoryItem]:
    rows = list_jobs(100)
    return [
        HistoryItem(
            job_id=r["job_id"],
            title=r["title"],
            style=r["style"],
            template_id=r["template_id"] if "template_id" in r.keys() else "executive_clean",
            status=r["status"],
            created_at=datetime.fromisoformat(r["created_at"]),
        )
        for r in rows
    ]


@router.post("/jobs/{job_id}/rewrite", response_model=GenerateResponse)
def rewrite(job_id: str, req: RewriteRequest) -> GenerateResponse:
    row = get_job(job_id)
    if not row:
        raise HTTPException(status_code=404, detail="任务不存在")

    material = row["material_text"] if "material_text" in row.keys() else ""

    result = run_generation(
        {
            "job_id": job_id,
            "title": row["title"],
            "style": row["style"],
            "template_id": row["template_id"] if "template_id" in row.keys() else "executive_clean",
            "target_pages": len(json.loads(row["outline_json"])),
            "material": material,
            "rewrite_action": req.action,
            "created_at": datetime.utcnow().isoformat(),
            "outline": json.loads(row["outline_json"]),
            "slides": json.loads(row["slides_json"]),
        }
    )
    return GenerateResponse(job_id=result["job_id"])
