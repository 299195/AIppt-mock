from datetime import datetime
from typing import List, Literal, Optional

from pydantic import BaseModel, Field


StyleType = Literal["management", "technical"]
RewriteAction = Literal["concise", "management", "technical"]
TemplateId = str


class GenerateResponse(BaseModel):
    job_id: str


class SlideDTO(BaseModel):
    page: int
    title: str
    bullets: List[str]
    notes: str
    slide_type: Optional[str] = None
    evidence: Optional[List[str]] = None


class JobDetailResponse(BaseModel):
    job_id: str
    status: str
    style: StyleType
    template_id: TemplateId = Field(default="executive_clean", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    title: str
    outline: List[str]
    slides: List[SlideDTO]
    pptx_url: Optional[str] = None
    created_at: datetime


class HistoryItem(BaseModel):
    job_id: str
    title: str
    style: StyleType
    template_id: TemplateId = Field(default="executive_clean", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    status: str
    created_at: datetime


class ModelConfigResponse(BaseModel):
    provider: str
    model: str
    use_mock: bool
    configured: bool
    base_url: str


class GenerateRequest(BaseModel):
    title: str = Field(min_length=2, max_length=200)
    material_text: str = Field(default="")
    outline_text: str = Field(default="")
    outline: Optional[List[str]] = None
    style: StyleType = "management"
    template_id: TemplateId = Field(default="executive_clean", pattern=r"^[a-z0-9_\\-]+$", min_length=2, max_length=80)
    target_pages: int = Field(default=8, ge=8, le=12)


class RewriteRequest(BaseModel):
    action: RewriteAction


class UploadParseResponse(BaseModel):
    extracted_text: str


class OutlinePreviewRequest(BaseModel):
    title: str = Field(min_length=2, max_length=200)
    material_text: str = Field(default="")
    outline_text: str = Field(default="")
    style: StyleType = "management"
    target_pages: int = Field(default=8, ge=8, le=12)


class OutlinePreviewResponse(BaseModel):
    outline: List[str]


class TemplateItem(BaseModel):
    id: TemplateId
    name: str
    subtitle: str
    summary: str
    preview_bg: str
    preview_fg: str
    preview_accent: str
    preview_image_url: Optional[str] = None
