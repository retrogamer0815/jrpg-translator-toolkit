#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Fetch and cache model catalogs for the JRPG Translator control panel."""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable


SCHEMA_VERSION = 1
DEFAULT_CACHE_HOURS = 24
PURPOSES = ("all", "screenshot", "explanation", "audio")
PROVIDERS = ("openai", "gemini")


def _configure_utf8_console() -> None:
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


def _load_environment() -> Path:
    root = Path(__file__).resolve().parents[1]
    try:
        from dotenv import find_dotenv, load_dotenv

        configured = (os.environ.get("SETTINGS_DIR", "") or "").strip()
        candidates = []
        if configured:
            candidates.append(Path(configured) / ".env")
        candidates.extend((root / "Settings" / ".env", root / ".env"))

        for candidate in candidates:
            if candidate.is_file():
                load_dotenv(candidate, override=False, encoding="utf-8-sig")
                break
        else:
            found = find_dotenv(usecwd=True)
            if found:
                load_dotenv(found, override=False, encoding="utf-8-sig")
    except Exception:
        pass
    return root


def _get_key(*names: str, file_var: str | None = None) -> str:
    bom = "\ufeff"
    for name in names:
        value = os.getenv(name) or os.getenv(bom + name, "")
        value = value.strip().strip('"').strip("'")
        if value:
            return value

    if file_var:
        key_path = os.getenv(file_var) or os.getenv(bom + file_var, "")
        key_path = key_path.strip().strip('"').strip("'")
        if key_path and Path(key_path).is_file():
            try:
                return Path(key_path).read_text(encoding="utf-8-sig").strip()
            except OSError:
                pass
    return ""


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _iso_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _parse_utc(value: str) -> datetime | None:
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(timezone.utc)
    except (TypeError, ValueError):
        return None


def _natural_key(value: str) -> list[tuple[int, Any]]:
    parts = re.split(r"(\d+)", value.lower())
    return [(0, int(part)) if part.isdigit() else (1, part) for part in parts if part]


def _clean_text(value: Any) -> str:
    return " ".join(str(value or "").replace("\t", " ").replace("\r", " ").splitlines()).strip()


def _entry(
    model_id: str,
    *,
    display_name: str = "",
    description: str = "",
    owner: str = "",
    supported_actions: Iterable[str] = (),
) -> dict[str, Any]:
    return {
        "id": model_id.removeprefix("models/"),
        "display_name": display_name or model_id.removeprefix("models/"),
        "description": description,
        "owner": owner,
        "supported_actions": sorted({_clean_text(action) for action in supported_actions if action}),
    }


def _fetch_openai() -> list[dict[str, Any]]:
    try:
        from openai import OpenAI
    except Exception as exc:
        raise RuntimeError(f"The OpenAI package could not be loaded: {exc}") from exc

    api_key = _get_key(
        "OPENAI_API_KEY",
        "OPENAI_LOCAL_KEY",
        "OPENAI_API_KEY_LOCAL",
        "OPENAI_KEY",
        file_var="OPENAI_API_KEY_FILE",
    )
    if not api_key:
        raise RuntimeError("No OpenAI API key was found in Settings\\.env or the Windows environment.")

    client = OpenAI(api_key=api_key)
    models = []
    for model in client.models.list():
        model_id = _clean_text(getattr(model, "id", ""))
        if model_id:
            models.append(
                _entry(
                    model_id,
                    owner=_clean_text(getattr(model, "owned_by", "")),
                )
            )
    return models


def _fetch_gemini() -> list[dict[str, Any]]:
    try:
        from google import genai
    except Exception as exc:
        raise RuntimeError(f"The google-genai package could not be loaded: {exc}") from exc

    api_key = _get_key(
        "GEMINI_API_KEY",
        "GOOGLE_API_KEY",
        "GEMINI_LOCAL_KEY",
        "GOOGLE_LOCAL_KEY",
        file_var="GEMINI_API_KEY_FILE",
    )
    if not api_key:
        raise RuntimeError("No Gemini API key was found in Settings\\.env or the Windows environment.")

    client = genai.Client(api_key=api_key)
    models = []
    for model in client.models.list():
        model_id = _clean_text(getattr(model, "name", ""))
        if not model_id:
            continue
        actions = getattr(model, "supported_actions", None) or []
        models.append(
            _entry(
                model_id,
                display_name=_clean_text(getattr(model, "display_name", "")),
                description=_clean_text(getattr(model, "description", "")),
                owner="google",
                supported_actions=actions,
            )
        )
    return models


def _is_general_openai_model(model_id: str, purpose: str) -> bool:
    name = model_id.lower()
    excluded = (
        "audio",
        "babbage",
        "codex",
        "dall-e",
        "davinci",
        "embedding",
        "image",
        "moderation",
        "realtime",
        "search",
        "speech",
        "transcribe",
        "transcription",
        "tts",
        "whisper",
    )
    if any(token in name for token in excluded):
        return False
    if not name.startswith(("gpt-", "chatgpt-", "o1", "o3", "o4", "ft:")):
        return False
    if purpose == "screenshot":
        if name.startswith("gpt-3.5") or name in ("gpt-4", "gpt-4-0613"):
            return False
    return True


def _is_general_gemini_model(item: dict[str, Any]) -> bool:
    name = item["id"].lower()
    actions = {action.lower() for action in item.get("supported_actions", [])}
    if actions and "generatecontent" not in actions:
        return False
    excluded = (
        "aqa",
        "audio",
        "embedding",
        "image",
        "imagen",
        "live",
        "nano-banana",
        "realtime",
        "robotics",
        "speech",
        "tts",
        "veo",
    )
    return name.startswith("gemini-") and not any(token in name for token in excluded)


def _filter_models(provider: str, purpose: str, entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if purpose == "all":
        selected = entries
    elif purpose == "audio":
        # live_audio_translator.py uses the providers' dedicated translation
        # protocols. General audio/realtime models may be listed by the APIs,
        # but they do not necessarily accept those translation endpoints and
        # session fields.
        if provider == "openai":
            selected = [
                item
                for item in entries
                if "realtime" in item["id"].lower() and "translate" in item["id"].lower()
            ]
        else:
            selected = [
                item
                for item in entries
                if item["id"].lower().startswith("gemini-")
                and "live" in item["id"].lower()
                and "translate" in item["id"].lower()
            ]
    elif provider == "openai":
        selected = [item for item in entries if _is_general_openai_model(item["id"], purpose)]
    else:
        selected = [item for item in entries if _is_general_gemini_model(item)]

    deduplicated = {item["id"].lower(): item for item in selected if item.get("id")}
    return sorted(deduplicated.values(), key=lambda item: _natural_key(item["id"]))


def _cache_path(cache_dir: Path, provider: str) -> Path:
    return cache_dir / f"{provider}_models.json"


def _read_cache(path: Path) -> dict[str, Any] | None:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return None
    if data.get("schema_version") != SCHEMA_VERSION or not isinstance(data.get("models"), list):
        return None
    return data


def _write_cache(path: Path, provider: str, entries: list[dict[str, Any]], fetched_at: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "schema_version": SCHEMA_VERSION,
        "provider": provider,
        "fetched_at": fetched_at,
        "models": entries,
    }
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)


def _cache_is_fresh(cache: dict[str, Any], max_age_hours: int) -> bool:
    fetched_at = _parse_utc(cache.get("fetched_at", ""))
    return bool(fetched_at and _utc_now() - fetched_at <= timedelta(hours=max_age_hours))


def _catalog(provider: str, purpose: str, cache_dir: Path, refresh: bool, cache_hours: int) -> dict[str, Any]:
    path = _cache_path(cache_dir, provider)
    cached = _read_cache(path)

    if cached and not refresh and _cache_is_fresh(cached, cache_hours):
        return {
            "ok": True,
            "provider": provider,
            "purpose": purpose,
            "source": "cache",
            "fetched_at": cached["fetched_at"],
            "models": _filter_models(provider, purpose, cached["models"]),
            "warnings": [],
        }

    try:
        entries = _fetch_openai() if provider == "openai" else _fetch_gemini()
        fetched_at = _iso_utc(_utc_now())
        entries = sorted(entries, key=lambda item: _natural_key(item["id"]))
        _write_cache(path, provider, entries, fetched_at)
        return {
            "ok": True,
            "provider": provider,
            "purpose": purpose,
            "source": "online",
            "fetched_at": fetched_at,
            "models": _filter_models(provider, purpose, entries),
            "warnings": [],
        }
    except Exception as exc:
        if cached:
            return {
                "ok": True,
                "provider": provider,
                "purpose": purpose,
                "source": "stale_cache",
                "fetched_at": cached.get("fetched_at", ""),
                "models": _filter_models(provider, purpose, cached["models"]),
                "warnings": [f"The online refresh failed; cached results are shown instead. {exc}"],
            }
        return {
            "ok": False,
            "provider": provider,
            "purpose": purpose,
            "source": "none",
            "fetched_at": "",
            "models": [],
            "warnings": [],
            "error": str(exc),
        }


def _as_ahk_lines(payload: dict[str, Any]) -> str:
    lines = [
        "JRPG_MODEL_CATALOG_V1",
        f"STATUS\t{'OK' if payload.get('ok') else 'ERROR'}",
        f"PROVIDER\t{_clean_text(payload.get('provider'))}",
        f"PURPOSE\t{_clean_text(payload.get('purpose'))}",
        f"SOURCE\t{_clean_text(payload.get('source'))}",
        f"FETCHED_AT\t{_clean_text(payload.get('fetched_at'))}",
    ]
    for model in payload.get("models", []):
        lines.append(f"MODEL\t{_clean_text(model.get('id'))}\t{_clean_text(model.get('display_name'))}")
    for warning in payload.get("warnings", []):
        lines.append(f"WARNING\t{_clean_text(warning)}")
    if payload.get("error"):
        lines.append(f"ERROR\t{_clean_text(payload['error'])}")
    return "\n".join(lines) + "\n"


def _arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="List provider models available to JRPG Translator.")
    parser.add_argument("--provider", required=True, choices=PROVIDERS)
    parser.add_argument("--purpose", default="all", choices=PURPOSES)
    parser.add_argument("--refresh", action="store_true", help="Ignore a fresh cache and query the provider.")
    parser.add_argument("--cache-hours", type=int, default=DEFAULT_CACHE_HOURS)
    parser.add_argument("--cache-dir", type=Path)
    parser.add_argument("--format", choices=("json", "ahk", "ids"), default="json")
    parser.add_argument("--output", type=Path, help="Write output to this file instead of stdout.")
    return parser.parse_args()


def main() -> int:
    _configure_utf8_console()
    root = _load_environment()
    args = _arguments()
    cache_dir = args.cache_dir or root / "Settings" / "model_catalog_cache"
    payload = _catalog(
        args.provider,
        args.purpose,
        cache_dir,
        args.refresh,
        max(0, args.cache_hours),
    )

    if args.format == "ahk":
        output = _as_ahk_lines(payload)
    elif args.format == "ids":
        output = "\n".join(model["id"] for model in payload.get("models", [])) + "\n"
    else:
        output = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output, encoding="utf-8")
    else:
        sys.stdout.write(output)
    return 0 if payload.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
