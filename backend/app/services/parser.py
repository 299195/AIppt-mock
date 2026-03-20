from __future__ import annotations

from pathlib import Path

from docx import Document


def parse_text_input(title: str, outline_text: str, uploaded_text: str = "") -> str:
    parts = [f"主题: {title}"]
    if outline_text.strip():
        parts.append(f"用户提纲:\n{outline_text.strip()}")
    if uploaded_text.strip():
        parts.append(f"上传资料:\n{uploaded_text.strip()}")
    return "\n\n".join(parts)


def read_uploaded_file(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    if suffix == ".md":
        return file_path.read_text(encoding="utf-8", errors="ignore")
    if suffix == ".docx":
        doc = Document(str(file_path))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    raise ValueError("仅支持 .md / .docx")
