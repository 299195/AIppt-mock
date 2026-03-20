import json
import os
from pathlib import Path

from pydantic import BaseModel


def _load_model_config(config_path: Path) -> dict:
    if not config_path.exists():
        return {}
    try:
        return json.loads(config_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def _env_or_default(name: str, default):
    value = os.getenv(name)
    if value is None:
        return default
    if isinstance(value, str) and value.strip() == "":
        return default
    return value


def _env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    try:
        return int(value)
    except ValueError:
        return default


MODEL_CONFIG_PATH = Path(__file__).resolve().parents[1] / "model_provider.json"
MODEL_CONFIG = _load_model_config(MODEL_CONFIG_PATH)


class Settings(BaseModel):
    project_name: str = "AI PPT Assistant"
    data_dir: Path = Path("data")
    export_dir: Path = Path("exports")
    database_path: Path = Path("data/history.db")

    model_provider: str = _env_or_default("MODEL_PROVIDER", MODEL_CONFIG.get("provider", "doubao"))
    use_mock_llm: bool = _env_bool("USE_MOCK_LLM", bool(MODEL_CONFIG.get("use_mock", True)))

    model_base_url: str = _env_or_default("MODEL_BASE_URL", MODEL_CONFIG.get("base_url", ""))
    model_api_key: str = _env_or_default("MODEL_API_KEY", MODEL_CONFIG.get("api_key", ""))
    model_name: str = _env_or_default("MODEL_NAME", MODEL_CONFIG.get("model", ""))
    model_chat_path: str = _env_or_default("MODEL_CHAT_PATH", MODEL_CONFIG.get("chat_path", "/v1/chat/completions"))
    request_timeout_sec: int = _env_int("MODEL_TIMEOUT", int(MODEL_CONFIG.get("timeout", 60)))


settings = Settings()
settings.data_dir.mkdir(parents=True, exist_ok=True)
settings.export_dir.mkdir(parents=True, exist_ok=True)

