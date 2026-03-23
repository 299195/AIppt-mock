from __future__ import annotations



import json

import re

from datetime import datetime

from pathlib import Path

from typing import Any, Iterable

from uuid import uuid4



from app.config import settings

from app.services.banana_ai_service import (

    BananaAIService,

    BananaProjectContext,

    build_idea_prompt,

    enforce_target_pages,

    make_project_context_from_row,

)

from app.services.image_generator import image_generator

from app.services.pptx_exporter import export_slides_to_pptx

from app.storage.db import (

    get_project,

    list_pages,

    make_progress,

    replace_pages,

    update_page,

    update_project,

    update_task,

)





banana_ai = BananaAIService(use_mock=settings.use_mock_llm)





class OutlinePreviewAdapter:

    """Compatibility adapter used by /outline/preview."""



    def generate_outline(self, title: str, style: str, material: str, target_pages: int) -> list[str]:

        context = BananaProjectContext(

            idea_prompt=build_idea_prompt(title, style, material),

            creation_type="idea",

        )

        pages = banana_ai.generate_outline(context, language="zh")

        pages = enforce_target_pages(pages, target_pages)

        return [str(page.get("title") or f"第{i + 1}页") for i, page in enumerate(pages)]





llm = OutlinePreviewAdapter()





def utc_now_iso() -> str:

    return datetime.utcnow().isoformat()






def clean_outline_items(items: Iterable[str]) -> list[str]:
    out: list[str] = []
    for raw in items:
        txt = _cleanup_text(str(raw or ""), max_len=120)
        if not txt:
            continue
        txt = re.sub(r"^\d+\s*[\.\u3001\)\]\uff09]\s*", "", txt)
        txt = txt.strip()
        if txt:
            out.append(txt)
    return out


def _outline_list_to_pages(items: list[str]) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    for idx, item in enumerate(items, start=1):
        title = _cleanup_text(str(item).strip(), max_len=120)
        parts = re.split(r"[\uFF1A:]", title, maxsplit=1)
        if len(parts) == 2:
            title = _cleanup_text(parts[1].strip() or parts[0].strip(), max_len=120)
        if not title:
            title = f"Page {idx}"
        pages.append({"title": title, "points": []})
    return pages


def _context_for_project(project_row: dict[str, Any] | Any) -> BananaProjectContext:
    return make_project_context_from_row(project_row)


def get_outline_for_project(
    project_row: dict[str, Any] | Any,
    requested_outline: list[str] | None = None,
) -> list[dict[str, Any]]:
    target_pages = int(project_row["target_pages"])

    if requested_outline:
        pages = _outline_list_to_pages(clean_outline_items(requested_outline))
        return enforce_target_pages(pages, target_pages)

    outline_text = str(project_row["outline_text"] or "")
    creation_type = str(project_row["creation_type"] or "idea")

    context = _context_for_project(project_row)

    if creation_type == "outline" and outline_text.strip():
        context.creation_type = "outline"
        context.outline_text = outline_text
        pages = banana_ai.parse_outline_text(context, language="zh")
        return enforce_target_pages(pages, target_pages)

    pages = banana_ai.generate_outline(context, language="zh")
    return enforce_target_pages(pages, target_pages)


def rebuild_project_pages(project_id: str, outline_pages: list[dict[str, Any]]) -> None:
    now = utc_now_iso()
    rows: list[dict[str, Any]] = []
    outline_lines: list[str] = []

    for idx, item in enumerate(outline_pages):
        title = str(item.get("title") or f"Page {idx + 1}")
        points = [str(x) for x in list(item.get("points") or []) if str(x).strip()]

        payload: dict[str, Any] = {
            "title": title,
            "points": points,
        }
        if item.get("part"):
            payload["part"] = str(item.get("part"))

        rows.append(
            {
                "page_id": str(uuid4()),
                "project_id": project_id,
                "order_index": idx,
                "outline_content": json.dumps(payload, ensure_ascii=False),
                "description_content": None,
                "status": "DRAFT",
                "created_at": now,
                "updated_at": now,
            }
        )
        outline_lines.append(f"{idx + 1}. {title}")

    replace_pages(project_id, rows)
    update_project(
        project_id,
        {
            "outline_text": "\n".join(outline_lines),
            "status": "OUTLINE_GENERATED",
            "updated_at": now,
        },
    )


def _safe_load_json(raw: str | None, fallback: Any) -> Any:
    if not raw:
        return fallback
    try:
        return json.loads(raw)
    except Exception:
        return fallback


_TITLE_LABEL_KEYS = (
    "page title",
    "title",
    "\u9875\u9762\u6807\u9898",
    "\u6807\u9898",
)
_TEXT_LABEL_KEYS = (
    "page text",
    "content",
    "body",
    "\u9875\u9762\u6587\u5b57",
    "\u9875\u9762\u5185\u5bb9",
    "\u6b63\u6587",
    "\u5185\u5bb9",
)
_NOTES_LABEL_KEYS = (
    "notes",
    "note",
    "materials",
    "material",
    "reference",
    "\u56fe\u7247\u7d20\u6750",
    "\u5176\u4ed6\u9875\u9762\u7d20\u6750",
    "\u89c6\u89c9\u5143\u7d20",
    "\u89c6\u89c9\u7126\u70b9",
    "\u6392\u7248\u5e03\u5c40",
    "\u6f14\u8bb2\u8005\u5907\u6ce8",
    "\u7d20\u6750",
)
_ALL_LABEL_KEYS = _TITLE_LABEL_KEYS + _TEXT_LABEL_KEYS + _NOTES_LABEL_KEYS


def _normalize_newlines(text: str) -> str:
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "").strip()


def _strip_markdown_prefix(text: str) -> str:
    out = str(text or "")
    out = re.sub(r"^\s*#{1,6}\s*", "", out)
    out = re.sub(r"^\s*[-*]+\s*", "", out)
    out = re.sub(r"^\s*\d+\s*[\.\u3001\)\]\uff09]\s*", "", out)
    return out.strip()


def _strip_xml_like(text: str) -> str:
    out = str(text or "")
    out = re.sub(r"</?[A-Za-z_][A-Za-z0-9._:-]*(?:\s[^>\n]*)?>", " ", out)
    out = re.sub(r"&lt;/?[A-Za-z_][^&]{0,120}&gt;", " ", out, flags=re.I)
    out = re.sub(r"\bxmlns(?::\w+)?=\"[^\"]*\"", " ", out, flags=re.I)
    return out


def _cleanup_text(text: str, max_len: int | None = None) -> str:
    out = _normalize_newlines(text)
    out = _strip_markdown_prefix(out)
    out = _strip_xml_like(out)
    out = re.sub(r"\s+", " ", out).strip(" -:\t\n")
    if max_len and len(out) > max_len:
        out = out[:max_len].rstrip()
    return out


def _parse_labeled_line(line: str) -> tuple[str | None, str]:
    m = re.match(r"^\s*([^:\uFF1A]{1,40})[\uFF1A:]\s*(.*)$", str(line or ""))
    if not m:
        return None, ""
    return m.group(1).strip().lower(), m.group(2).strip()


def _is_label(label: str | None, keys: tuple[str, ...]) -> bool:
    if not label:
        return False
    val = str(label).strip().lower()
    return any(k in val for k in keys)


def _extract_labeled_section(raw_text: str, start_keys: tuple[str, ...], stop_keys: tuple[str, ...] | None = None) -> str:
    text_norm = _normalize_newlines(raw_text)
    lines = text_norm.split("\n")
    stop_set = stop_keys or _ALL_LABEL_KEYS

    capture = False
    out_lines: list[str] = []
    for raw_line in lines:
        line = raw_line.strip()
        label, value = _parse_labeled_line(line)

        if label and _is_label(label, start_keys):
            capture = True
            if value:
                out_lines.append(value)
            continue

        if capture and label and _is_label(label, stop_set):
            break

        if capture:
            out_lines.append(line)

    return "\n".join(out_lines).strip()


def _infer_slide_type(title: str, bullets: list[str], page_index: int, total: int) -> str:
    full = f"{title} {' '.join(bullets)}".lower()
    title_lower = str(title or "").lower()

    def contains_any(value: str, words: tuple[str, ...]) -> bool:
        return any(w in value for w in words)

    if page_index == 1 or contains_any(title_lower, ("\u5c01\u9762", "\u6807\u9898", "cover", "title")):
        return "title"
    if contains_any(full, ("\u98ce\u9669", "\u95ee\u9898", "\u6311\u6218", "\u963b\u585e", "\u9690\u60a3", "risk", "issue", "challenge", "blocker")):
        return "risk"
    if contains_any(full, ("\u8ba1\u5212", "\u8def\u7ebf", "\u91cc\u7a0b\u7891", "\u9636\u6bb5", "\u8fdb\u5ea6", "\u6392\u671f", "timeline", "plan", "roadmap", "milestone")):
        return "timeline"
    if contains_any(full, ("\u6570\u636e", "\u6307\u6807", "\u540c\u6bd4", "\u73af\u6bd4", "\u589e\u957f", "%", "roi", "gmv", "metric", "kpi", "trend")):
        return "data"
    if total >= 8 and page_index in (4, 5):
        return "data"
    return "summary"


def _extract_title(raw_text: str, fallback: str) -> str:
    raw = _normalize_newlines(raw_text)

    section_title = _extract_labeled_section(raw, _TITLE_LABEL_KEYS, _TEXT_LABEL_KEYS + _NOTES_LABEL_KEYS)
    title = _cleanup_text(section_title, max_len=90)
    if title:
        return title

    for line in raw.split("\n"):
        label, value = _parse_labeled_line(line)
        if label and _is_label(label, _TITLE_LABEL_KEYS):
            t = _cleanup_text(value, max_len=90)
            if t:
                return t
            continue
        if label and _is_label(label, _ALL_LABEL_KEYS):
            continue
        candidate = _cleanup_text(line, max_len=90)
        if candidate:
            return candidate

    cleaned_fallback = _cleanup_text(fallback, max_len=90)
    return cleaned_fallback or str(fallback or "Untitled")


def _extract_page_text_section(raw_text: str) -> str:
    raw = _normalize_newlines(raw_text)

    section = _extract_labeled_section(raw, _TEXT_LABEL_KEYS, _NOTES_LABEL_KEYS)
    if section.strip():
        return section.strip()

    lines: list[str] = []
    for line in raw.split("\n"):
        label, value = _parse_labeled_line(line)
        if label and _is_label(label, _TITLE_LABEL_KEYS):
            continue
        if label and _is_label(label, _NOTES_LABEL_KEYS):
            continue
        if label and _is_label(label, _TEXT_LABEL_KEYS):
            if value.strip():
                lines.append(value.strip())
            continue
        lines.append(line.strip())

    return "\n".join(lines).strip()


def _extract_bullets(text_section: str, fallback_points: list[str]) -> list[str]:
    bullet_lines: list[str] = []
    text_body = re.sub(r"```[\s\S]*?```", " ", _normalize_newlines(text_section))

    for raw in text_body.splitlines():
        line = _strip_markdown_prefix(raw)
        label, value = _parse_labeled_line(line)
        if label and _is_label(label, _ALL_LABEL_KEYS):
            line = value

        cleaned = _cleanup_text(line, max_len=150)
        if not cleaned:
            continue

        if re.search(r"</?[A-Za-z_][A-Za-z0-9._:-]*", cleaned):
            continue

        if cleaned not in bullet_lines:
            bullet_lines.append(cleaned)

    if not bullet_lines:
        for point in fallback_points:
            cleaned = _cleanup_text(point, max_len=150)
            if cleaned and cleaned not in bullet_lines:
                bullet_lines.append(cleaned)

    if not bullet_lines:
        bullet_lines = ["Core conclusion", "Key evidence", "Next action"]

    while len(bullet_lines) < 3:
        bullet_lines.append("Additional point")

    return bullet_lines[:5]


def _extract_notes(raw_text: str, extra_fields: dict[str, str] | None) -> str:
    notes_parts: list[str] = []

    notes_section = _extract_labeled_section(raw_text, _NOTES_LABEL_KEYS, ())
    cleaned_notes = _cleanup_text(notes_section, max_len=800)
    if cleaned_notes:
        notes_parts.append(cleaned_notes)

    if extra_fields:
        for name, value in extra_fields.items():
            value_s = _cleanup_text(str(value or ""), max_len=200)
            if value_s:
                notes_parts.append(f"{name}: {value_s}")

    merged = "\n".join(notes_parts).strip()
    return merged or "Generated from banana workflow"


def _description_to_slide_payload(
    description_text: str,
    page_outline: dict[str, Any],
    page_index: int,
    total_pages: int,
    extra_fields: dict[str, str] | None = None,
) -> dict[str, Any]:
    fallback_title = str(page_outline.get("title") or f"Page {page_index}")
    fallback_points = [str(x) for x in list(page_outline.get("points") or [])]

    normalized_text = _normalize_newlines(description_text)
    title = _extract_title(normalized_text, fallback_title)
    page_text = _extract_page_text_section(normalized_text)
    bullets = _extract_bullets(page_text, fallback_points)
    notes = _extract_notes(normalized_text, extra_fields)
    slide_type = _infer_slide_type(title, bullets, page_index, total_pages)

    evidence = bullets[:3]
    payload: dict[str, Any] = {
        "title": title,
        "bullets": bullets,
        "notes": notes,
        "slide_type": slide_type,
        "evidence": evidence,
        "chart_data": None,
        "text": normalized_text,
        "generated_image_path": None,
    }
    if extra_fields:
        payload["extra_fields"] = extra_fields
    return payload


def _outline_pages_from_db(pages: list[Any]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for idx, page in enumerate(pages, start=1):
        outline = _safe_load_json(page["outline_content"], {})
        out.append(
            {
                "title": str(outline.get("title") or f"Page {idx}"),
                "points": [str(x) for x in list(outline.get("points") or [])],
                "part": outline.get("part"),
            }
        )
    return out


def generate_descriptions_task(task_id: str, project_id: str) -> None:

    project = get_project(project_id)

    if not project:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": "project not found",

                "completed_at": utc_now_iso(),

            },

        )

        return



    pages = list_pages(project_id)

    if not pages:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": "no pages to generate",

                "completed_at": utc_now_iso(),

            },

        )

        return



    total = len(pages)

    update_task(

        task_id,

        {

            "status": "PROCESSING",

            "progress_json": make_progress(total, 0, 0, "generating_descriptions"),

        },

    )



    try:

        context = _context_for_project(project)

        outline = _outline_pages_from_db(pages)



        completed = 0

        failed = 0



        for idx, page in enumerate(pages):

            page_id = str(page["page_id"])

            page_outline = outline[idx]

            try:

                result = banana_ai.generate_page_description(

                    project_context=context,

                    outline=outline,

                    page_outline=page_outline,

                    page_index=idx + 1,

                    language="zh",

                    detail_level="default",

                )

                desc_text = str(result.get("text") or "")

                extra_fields = result.get("extra_fields") if isinstance(result.get("extra_fields"), dict) else None



                payload = _description_to_slide_payload(

                    description_text=desc_text,

                    page_outline=page_outline,

                    page_index=idx + 1,

                    total_pages=total,

                    extra_fields=extra_fields,

                )



                update_page(

                    page_id,

                    {

                        "description_content": json.dumps(payload, ensure_ascii=False),

                        "status": "DESCRIPTION_GENERATED",

                        "updated_at": utc_now_iso(),

                    },

                )

                completed += 1

            except Exception:

                failed += 1

                update_page(

                    page_id,

                    {

                        "status": "FAILED",

                        "updated_at": utc_now_iso(),

                    },

                )



            update_task(

                task_id,

                {

                    "progress_json": make_progress(total, completed, failed, "generating_descriptions"),

                },

            )



        final_status = "COMPLETED" if failed == 0 else "FAILED"

        if failed == 0:

            update_project(

                project_id,

                {

                    "status": "DESCRIPTIONS_GENERATED",

                    "updated_at": utc_now_iso(),

                },

            )



        update_task(

            task_id,

            {

                "status": final_status,

                "progress_json": make_progress(total, completed, failed, "descriptions_done"),

                "error_message": None if failed == 0 else f"{failed} pages failed",

                "completed_at": utc_now_iso(),

            },

        )

    except Exception as exc:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": str(exc),

                "completed_at": utc_now_iso(),

            },

        )





def _collect_project_slides(project: Any, pages: list[Any]) -> list[dict[str, Any]]:

    total = len(pages)

    outline = _outline_pages_from_db(pages)



    slides: list[dict[str, Any]] = []

    for idx, page in enumerate(pages):

        desc = _safe_load_json(page["description_content"], None)

        page_outline = outline[idx]



        if not isinstance(desc, dict):

            result = banana_ai.generate_page_description(

                project_context=_context_for_project(project),

                outline=outline,

                page_outline=page_outline,

                page_index=idx + 1,

                language="zh",

                detail_level="default",

            )

            desc = _description_to_slide_payload(

                description_text=str(result.get("text") or ""),

                page_outline=page_outline,

                page_index=idx + 1,

                total_pages=total,

                extra_fields=result.get("extra_fields") if isinstance(result.get("extra_fields"), dict) else None,

            )

            update_page(

                str(page["page_id"]),

                {

                    "description_content": json.dumps(desc, ensure_ascii=False),

                    "status": "DESCRIPTION_GENERATED",

                    "updated_at": utc_now_iso(),

                },

            )



        if "text" in desc and "bullets" in desc and "title" in desc:

            normalized = desc

        else:

            normalized = _description_to_slide_payload(

                description_text=str(desc.get("text") or ""),

                page_outline=page_outline,

                page_index=idx + 1,

                total_pages=total,

                extra_fields=desc.get("extra_fields") if isinstance(desc.get("extra_fields"), dict) else None,

            )

            update_page(

                str(page["page_id"]),

                {

                    "description_content": json.dumps(normalized, ensure_ascii=False),

                    "status": "DESCRIPTION_GENERATED",

                    "updated_at": utc_now_iso(),

                },

            )



        slide = {

            "page": idx + 1,

            "title": str(normalized.get("title") or page_outline["title"]),

            "bullets": [str(x) for x in list(normalized.get("bullets") or [])],

            "notes": str(normalized.get("notes") or ""),

            "slide_type": str(normalized.get("slide_type") or "summary"),

            "evidence": normalized.get("evidence"),

            "chart_data": normalized.get("chart_data"),

            "generated_image_path": normalized.get("generated_image_path"),

        }

        slides.append(slide)



    return slides



def _existing_image_path(raw: Any) -> str | None:

    if not raw:

        return None

    path = Path(str(raw))

    if path.exists() and path.is_file():

        return str(path)

    return None





def _ensure_slide_images(project: Any, project_id: str, pages: list[Any], slides: list[dict[str, Any]], task_id: str) -> None:

    if not image_generator.enabled():

        return



    total = max(1, len(slides))

    project_title = str(project["title"] or "")

    style = str(project["style"] or "management")



    for idx, slide in enumerate(slides):

        page_index = idx + 1

        if str(slide.get("slide_type") or "").lower() == "title":

            continue



        existing = _existing_image_path(slide.get("generated_image_path"))

        if existing:

            slide["generated_image_path"] = existing

            continue



        generated_path = image_generator.generate_for_slide(

            project_id=project_id,

            page_index=page_index,

            topic=project_title,

            title=str(slide.get("title") or f"第{page_index}页"),

            bullets=[str(x) for x in list(slide.get("bullets") or [])],

            notes=str(slide.get("notes") or ""),

            style=style,

        )

        if not generated_path:

            continue



        slide["generated_image_path"] = generated_path



        raw_desc = _safe_load_json(pages[idx]["description_content"], {})

        if isinstance(raw_desc, dict):

            raw_desc["generated_image_path"] = generated_path

            update_page(

                str(pages[idx]["page_id"]),

                {

                    "description_content": json.dumps(raw_desc, ensure_ascii=False),

                    "updated_at": utc_now_iso(),

                },

            )



        update_task(

            task_id,

            {

                "progress_json": make_progress(total, min(total, page_index), 0, "generating_images"),

            },

        )





def generate_ppt_task(task_id: str, project_id: str) -> None:

    project = get_project(project_id)

    if not project:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": "project not found",

                "completed_at": utc_now_iso(),

            },

        )

        return



    pages = list_pages(project_id)

    if not pages:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": "no pages to export",

                "completed_at": utc_now_iso(),

            },

        )

        return



    total = len(pages)

    update_task(

        task_id,

        {

            "status": "PROCESSING",

            "progress_json": make_progress(total, 0, 0, "building_slides"),

        },

    )



    try:

        slides = _collect_project_slides(project, pages)

        _ensure_slide_images(project, project_id, pages, slides, task_id)

        update_task(

            task_id,

            {

                "progress_json": make_progress(total, total, 0, "exporting_pptx"),

            },

        )



        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        filename = f"{project_id}_{ts}.pptx"

        out_path = settings.export_dir / filename



        outline = [str(_safe_load_json(page["outline_content"], {}).get("title") or "") for page in pages]



        exported = export_slides_to_pptx(

            slides,

            out_path,

            str(project["template_id"] or "executive_clean"),

            str(project["title"]),

            outline,

        )



        pptx_url = f"/exports/{exported}"

        update_project(

            project_id,

            {

                "status": "COMPLETED",

                "pptx_url": pptx_url,

                "updated_at": utc_now_iso(),

            },

        )

        update_task(

            task_id,

            {

                "status": "COMPLETED",

                "progress_json": make_progress(total, total, 0, "done"),

                "result_json": json.dumps({"pptx_url": pptx_url}, ensure_ascii=False),

                "completed_at": utc_now_iso(),

            },

        )

    except Exception as exc:

        update_task(

            task_id,

            {

                "status": "FAILED",

                "error_message": str(exc),

                "completed_at": utc_now_iso(),

            },

        )





def _slide_payload_to_description_text(payload: dict[str, Any]) -> str:
    title = str(payload.get("title") or "Untitled")
    bullets = [str(x) for x in list(payload.get("bullets") or [])]
    notes = str(payload.get("notes") or "")

    lines = [f"Page Title: {title}", "", "Page Text:"]
    for item in bullets[:5]:
        cleaned = _cleanup_text(item, max_len=180)
        if cleaned:
            lines.append(f"- {cleaned}")

    if notes:
        lines.extend(["", "Notes:", notes])

    return "\n".join(lines)

def _rewrite_requirement(action: str) -> str:

    mapping = {

        "concise": "请将所有页面内容精简为更短的表达，保留关键信息和结论。",

        "management": "请把所有页面改写成管理层汇报口径，强调结果、风险和决策建议。",

        "technical": "请把所有页面改写成技术汇报口径，强调现状、细节和实施计划。",

    }

    return mapping.get(action, "请优化页面描述表达。")





def rewrite_project(project_id: str, action: str) -> str:

    project = get_project(project_id)

    if not project:

        raise ValueError("project not found")



    pages = list_pages(project_id)

    if not pages:

        raise ValueError("no pages")



    outline = _outline_pages_from_db(pages)



    current_descriptions: list[dict[str, Any]] = []

    for idx, page in enumerate(pages):

        desc = _safe_load_json(page["description_content"], {})

        if isinstance(desc, dict) and desc.get("text"):

            raw_text = str(desc.get("text"))

        elif isinstance(desc, dict):

            raw_text = _slide_payload_to_description_text(desc)

        else:

            raw_text = ""



        current_descriptions.append(

            {

                "index": idx,

                "title": outline[idx]["title"],

                "description_content": {"text": raw_text},

            }

        )



    context = _context_for_project(project)

    user_requirement = _rewrite_requirement(action)



    refined = banana_ai.refine_descriptions(

        current_descriptions=current_descriptions,

        user_requirement=user_requirement,

        project_context=context,

        outline=outline,

        previous_requirements=None,

        language="zh",

    )



    if len(refined) < len(pages):

        for idx in range(len(refined), len(pages)):

            refined.append(current_descriptions[idx]["description_content"]["text"])

    elif len(refined) > len(pages):

        refined = refined[: len(pages)]

    rewritten_slides: list[dict[str, Any]] = []

    total = len(pages)

    project_title = str(project["title"] or "")

    image_style = action if action in {"management", "technical"} else str(project["style"] or "management")



    for idx, refined_text in enumerate(refined):

        page_id = str(pages[idx]["page_id"])

        payload = _description_to_slide_payload(

            description_text=str(refined_text),

            page_outline=outline[idx],

            page_index=idx + 1,

            total_pages=total,

            extra_fields=None,

        )

        payload["page"] = idx + 1



        if image_generator.enabled() and str(payload.get("slide_type") or "").lower() != "title":

            generated_path = image_generator.generate_for_slide(

                project_id=project_id,

                page_index=idx + 1,

                topic=project_title,

                title=str(payload.get("title") or f"第{idx + 1}页"),

                bullets=[str(x) for x in list(payload.get("bullets") or [])],

                notes=str(payload.get("notes") or ""),

                style=image_style,

            )

            if generated_path:

                payload["generated_image_path"] = generated_path



        rewritten_slides.append(payload)



        update_page(

            page_id,

            {

                "description_content": json.dumps(payload, ensure_ascii=False),

                "status": "DESCRIPTION_GENERATED",

                "updated_at": utc_now_iso(),

            },

        )

    style = str(project["style"])

    if action in {"management", "technical"}:

        style = action



    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    filename = f"{project_id}_{ts}.pptx"

    out_path = settings.export_dir / filename

    outline_titles = [item["title"] for item in outline]



    exported = export_slides_to_pptx(

        rewritten_slides,

        out_path,

        str(project["template_id"] or "executive_clean"),

        str(project["title"]),

        outline_titles,

    )



    pptx_url = f"/exports/{exported}"

    update_project(

        project_id,

        {

            "style": style,

            "status": "COMPLETED",

            "pptx_url": pptx_url,

            "updated_at": utc_now_iso(),

        },

    )

    return pptx_url















