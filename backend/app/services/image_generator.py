from __future__ import annotations

import base64
import hashlib
import json
import logging
import ssl
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app.config import settings

logger = logging.getLogger(__name__)


def _parse_size(raw: str) -> tuple[int, int]:
    text = (raw or "").lower().strip()
    if "x" not in text:
        return (1536, 1024)
    left, right = text.split("x", 1)
    try:
        w = max(256, min(4096, int(left.strip())))
        h = max(256, min(4096, int(right.strip())))
    except ValueError:
        return (1536, 1024)
    return (w, h)


def _slug(text: str) -> str:
    token = "".join(ch if ch.isalnum() else "_" for ch in text.strip().lower()).strip("_")
    if token:
        return token[:48]
    return hashlib.md5(text.encode("utf-8")).hexdigest()[:12]


class ImageGenerator:
    def __init__(self) -> None:
        self.base_url = settings.image_base_url
        self.api_key = settings.image_api_key
        self.model = settings.image_model

    def enabled(self) -> bool:
        return bool(settings.enable_image_generation)

    def _remote_enabled(self) -> bool:
        return bool(self.base_url and self.api_key and self.model)

    @staticmethod
    def _build_image_prompt(
        topic: str,
        title: str,
        bullets: list[str],
        notes: str,
        style: str,
        page_index: int,
    ) -> str:
        points = "\n".join([f"- {item}" for item in bullets[:5]])
        notes_short = (notes or "").strip()
        if len(notes_short) > 260:
            notes_short = notes_short[:260] + "..."

        return (
            "你是一位资深PPT视觉设计师。请为当前页面生成一张可直接用于PPT的高质量插图或背景图。\n"
            "要求：\n"
            "1. 构图简洁、专业、信息聚焦，不要输出任何文字。\n"
            "2. 视觉风格与页面主题一致，适合商业汇报场景。\n"
            f"3. 当前页序号：第{page_index}页。\n"
            f"4. 汇报主题：{topic}\n"
            f"5. 页面标题：{title}\n"
            f"6. 风格偏好：{style}\n"
            "7. 请根据以下页面要点提取视觉隐喻并设计画面：\n"
            f"{points}\n"
            f"8. 备注信息：{notes_short if notes_short else '无'}\n"
            "输出仅用于图像模型输入，不要包含任何markdown或代码块。"
        )

    def _target_path(self, project_id: str, page_index: int, title: str) -> Path:
        folder = settings.generated_image_dir / project_id
        folder.mkdir(parents=True, exist_ok=True)
        return folder / f"{page_index:02d}_{_slug(title)}.png"

    def _save_bytes(self, raw: bytes, out_path: Path) -> str:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(raw)
        return str(out_path)

    def _download(self, url: str) -> bytes:
        req = Request(url=url, method="GET")
        context = ssl.create_default_context()
        with urlopen(req, timeout=settings.image_timeout_sec, context=context) as resp:
            return resp.read()

    def _generate_remote(self, prompt: str) -> bytes:
        payload: dict[str, Any] = {
            "model": self.model,
            "prompt": prompt,
            "size": settings.image_size,
            "response_format": "b64_json",
        }

        url = self.base_url.rstrip("/") + settings.image_gen_path
        req = Request(url=url, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {self.api_key}")

        context = ssl.create_default_context()
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

        try:
            with urlopen(req, body, timeout=settings.image_timeout_sec, context=context) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"image HTTPError: {exc.code} {detail}") from exc
        except URLError as exc:
            raise RuntimeError(f"image URLError: {exc}") from exc

        items = data.get("data")
        if not isinstance(items, list) or not items:
            raise RuntimeError("image response missing data")

        first = items[0]
        if isinstance(first, dict) and first.get("b64_json"):
            return base64.b64decode(first["b64_json"])

        if isinstance(first, dict) and first.get("url"):
            return self._download(str(first["url"]))

        raise RuntimeError("image response has no b64_json/url")

    def _generate_mock(self, out_path: Path, title: str, bullets: list[str]) -> str:
        try:
            from PIL import Image, ImageDraw
        except Exception as exc:
            logger.warning("Pillow unavailable for mock image: %s", exc)
            return ""

        width, height = _parse_size(settings.image_size)
        seed = int(hashlib.md5((title + "|" + "|".join(bullets)).encode("utf-8")).hexdigest()[:8], 16)

        base_r = 36 + (seed % 96)
        base_g = 62 + ((seed >> 8) % 96)
        base_b = 90 + ((seed >> 16) % 96)

        image = Image.new("RGB", (width, height), (base_r, base_g, base_b))
        draw = ImageDraw.Draw(image)

        for i in range(0, height, max(24, height // 36)):
            mix = i / max(1, height - 1)
            r = min(255, int(base_r + 90 * mix))
            g = min(255, int(base_g + 70 * mix))
            b = min(255, int(base_b + 60 * mix))
            draw.rectangle([(0, i), (width, min(height, i + max(18, height // 42)))], fill=(r, g, b))

        pad_x = int(width * 0.08)
        pad_y = int(height * 0.1)
        card_w = int(width * 0.84)
        card_h = int(height * 0.74)
        draw.rounded_rectangle(
            [(pad_x, pad_y), (pad_x + card_w, pad_y + card_h)],
            radius=max(16, int(min(width, height) * 0.02)),
            fill=(245, 248, 252),
            outline=(220, 228, 240),
            width=2,
        )

        line_y = pad_y + 26
        draw.rounded_rectangle(
            [(pad_x + 24, line_y), (pad_x + card_w - 24, line_y + 18)],
            radius=8,
            fill=(80, 112, 168),
        )

        y = line_y + 44
        for idx in range(min(4, len(bullets) if bullets else 4)):
            w = int(card_w * (0.72 - idx * 0.08))
            draw.rounded_rectangle(
                [(pad_x + 30, y), (pad_x + 30 + w, y + 14)],
                radius=7,
                fill=(128, 152, 186),
            )
            y += 30

        bubble_x = int(width * 0.58)
        bubble_y = int(height * 0.5)
        bubble_w = int(width * 0.26)
        bubble_h = int(height * 0.24)
        draw.ellipse(
            [(bubble_x, bubble_y), (bubble_x + bubble_w, bubble_y + bubble_h)],
            fill=(228, 236, 248),
            outline=(169, 190, 220),
            width=2,
        )

        buffer = BytesIO()
        image.save(buffer, format="PNG")
        return self._save_bytes(buffer.getvalue(), out_path)

    def generate_for_slide(
        self,
        project_id: str,
        page_index: int,
        topic: str,
        title: str,
        bullets: list[str],
        notes: str,
        style: str,
    ) -> str | None:
        if not self.enabled():
            return None

        out_path = self._target_path(project_id, page_index, title)
        prompt = self._build_image_prompt(topic, title, bullets, notes, style, page_index)

        if settings.use_mock_image or not self._remote_enabled():
            mock_path = self._generate_mock(out_path, title, bullets)
            return mock_path or None

        try:
            image_bytes = self._generate_remote(prompt)
            return self._save_bytes(image_bytes, out_path)
        except Exception as exc:
            logger.warning("image generation failed for project=%s page=%s: %s", project_id, page_index, exc)
            if settings.image_fallback_mock:
                mock_path = self._generate_mock(out_path, title, bullets)
                return mock_path or None
            return None


image_generator = ImageGenerator()
