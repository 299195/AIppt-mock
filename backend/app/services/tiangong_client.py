from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Dict
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app.config import settings


@dataclass
class TianGongClient:
    base_url: str = settings.tiangong_base_url
    api_key: str = settings.tiangong_api_key
    model: str = settings.tiangong_model

    def enabled(self) -> bool:
        return bool(self.base_url and self.api_key and self.model)

    def chat_json(self, system_prompt: str, user_prompt: str, temperature: float = 0.3) -> Dict[str, Any]:
        if not self.enabled():
            raise RuntimeError("TianGong配置不完整")

        payload = {
            "model": self.model,
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "response_format": {"type": "json_object"},
        }

        url = self.base_url.rstrip("/") + settings.tiangong_chat_path
        req = Request(url=url, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {self.api_key}")

        body = json.dumps(payload).encode("utf-8")
        try:
            with urlopen(req, body, timeout=settings.request_timeout_sec) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except HTTPError as e:
            detail = e.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"TianGong HTTPError: {e.code} {detail}")
        except URLError as e:
            raise RuntimeError(f"TianGong URLError: {e}")

        content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        if not content:
            raise RuntimeError("TianGong返回内容为空")

        return json.loads(content)
