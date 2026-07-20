#!/usr/bin/env python
# -*- coding: utf-8 -*-

import configparser
import io
import os
import re
import sys
import tempfile
from typing import List, Tuple

from dotenv import load_dotenv

# Console UTF-8 (Windows safe)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# Shared temp folder (same as audio + overlay)
TEMP_DIR      = os.environ.get("TEMP") or tempfile.gettempdir()
OVERLAY_DIR   = os.path.join(TEMP_DIR, "JRPG_Overlay")
LAST_JP       = os.path.join(OVERLAY_DIR, "last_jp.txt")
EXPLAINER_TXT = os.path.join(OVERLAY_DIR, "explainer.txt")
os.makedirs(OVERLAY_DIR, exist_ok=True)

import time  # add near the other imports

def atomic_write_text(path: str, text: str):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)
    # Retry a few times in case another process has the file open without FILE_SHARE_DELETE
    for _ in range(10):              # total ≈ 500 ms
        try:
            os.replace(tmp, path)    # atomic on Win/NTFS if target isn’t locked
            break
        except PermissionError:
            time.sleep(0.05)
    else:
        # Last resort: try writing directly (non-atomic) to avoid leaving a .tmp around
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(text)
        try:
            os.remove(tmp)
        except Exception:
            pass

def read_text(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


def load_glossary(path: str) -> List[Tuple[str, str]]:
    """Load source-to-target mappings using the translator's accepted formats."""
    entries: List[Tuple[str, str]] = []
    if not path or not os.path.isfile(path):
        return entries

    encodings = [
        "utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be",
        "cp1252", "cp932",
    ]
    text = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding) as glossary_file:
                text = glossary_file.read()
            break
        except Exception:
            continue
    if text is None:
        return entries

    for raw_line in text.splitlines():
        line = raw_line.replace("\ufeff", "").strip()
        if not line or line.startswith("#"):
            continue
        source = target = None
        for separator in ("->", "→", "\t", ":", "="):
            if separator in line:
                source_part, target_part = line.split(separator, 1)
                source, target = source_part.strip(), target_part.strip()
                break
        if source and target:
            entries.append((source, target))
    return entries


def _sentence_case(text: str) -> str:
    lowered = text.lower()
    for index, char in enumerate(lowered):
        if char.isalpha():
            return lowered[:index] + char.upper() + lowered[index + 1:]
    return lowered


def _adapt_glossary_case(source: str, target: str, matched: str) -> str:
    """Adapt a replacement's case unless its source declares canonical casing."""
    source_first_letter = next((char for char in source if char.isalpha()), "")
    if source_first_letter.isupper():
        return target

    matched_letters = [char for char in matched if char.isalpha()]
    if not matched_letters:
        return target
    if all(char.isupper() for char in matched_letters):
        return target.upper()
    if all(char.islower() for char in matched_letters):
        return target.lower()

    if re.search(r"\s", source):
        words = re.findall(r"[^\W\d_]+", matched, flags=re.UNICODE)
        if words and all(
            word[0].isupper() and word[1:].islower() for word in words
        ):
            return target.title()

    if matched_letters[0].isupper() and all(
        char.islower() for char in matched_letters[1:]
    ):
        return _sentence_case(target)

    return target


def apply_target_glossary(
    text: str,
    glossary: List[Tuple[str, str]],
    protected_text: str = "",
) -> str:
    """Apply target-language replacements while preserving original source text."""
    protected_segments = []
    if protected_text:
        if protected_text in text:
            protected_segments = [protected_text]
        else:
            protected_segments = [
                line for line in protected_text.splitlines() if line.strip()
            ]

    protected_values = []
    out = text
    for index, segment in enumerate(protected_segments):
        token = f"\x00JRPG_PROTECTED_SOURCE_{index}\x00"
        if segment in out:
            out = out.replace(segment, token, 1)
            protected_values.append((token, segment))

    for source, target in glossary:
        if re.search(r"\s", source):
            pattern = re.compile(re.escape(source), flags=re.IGNORECASE)
            out = pattern.sub(
                lambda match: _adapt_glossary_case(
                    source, target, match.group(0)
                ),
                out,
            )
        else:
            pattern = re.compile(
                rf"\b(?P<core>{re.escape(source)})(?P<suf>s|'s|’s)?\b",
                flags=re.IGNORECASE,
            )
            out = pattern.sub(
                lambda match: (
                    _adapt_glossary_case(source, target, match.group("core"))
                    + (match.group("suf") or "")
                ),
                out,
            )

    for token, segment in protected_values:
        out = out.replace(token, segment)
    return out


def build_source_glossary_prompt(glossary: List[Tuple[str, str]]) -> str:
    if not glossary:
        return ""
    lines = [
        "Terminology overrides for this explanation:",
        "Keep the Original Japanese line unchanged.",
        "In meanings, glosses, and target-language paraphrases, use these exact "
        "Japanese-to-target mappings:",
    ]
    lines.extend(f"- {source} → {target}" for source, target in glossary)
    return "\n".join(lines)


def resolve_glossary_paths(project_root: str) -> Tuple[str, str]:
    """Resolve the profiles selected in the control panel's Terminology tab."""
    source_path = (
        os.environ.get("JP2TL_GLOSSARY_PATH", "").strip()
        or os.environ.get("JP2EN_GLOSSARY_PATH", "").strip()
    )
    target_path = (
        os.environ.get("TL2TL_GLOSSARY_PATH", "").strip()
        or os.environ.get("EN2EN_GLOSSARY_PATH", "").strip()
    )
    if source_path and target_path:
        return source_path, target_path

    settings_dir = (
        os.environ.get("SETTINGS_DIR", "").strip()
        or os.path.join(project_root, "Settings")
    )
    source_profile = target_profile = "default"
    for ini_name in ("control.ini", "config.ini"):
        ini_path = os.path.join(settings_dir, ini_name)
        if not os.path.isfile(ini_path):
            continue
        for encoding in (
            "utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be",
            "cp1252", "cp932",
        ):
            try:
                config = configparser.ConfigParser(interpolation=None)
                with open(ini_path, "r", encoding=encoding) as ini_file:
                    config.read_file(ini_file)
                source_profile = (
                    config.get(
                        "cfg", "jp2enGlossaryProfile", fallback="default"
                    ).strip()
                    or "default"
                )
                target_profile = (
                    config.get(
                        "cfg", "en2enGlossaryProfile", fallback="default"
                    ).strip()
                    or "default"
                )
                break
            except Exception:
                continue
        break

    glossary_dir = os.path.join(settings_dir, "glossaries")
    if not source_path:
        source_path = os.path.join(glossary_dir, source_profile, "jp2en.txt")
    if not target_path:
        target_path = os.path.join(glossary_dir, target_profile, "en2en.txt")
    return source_path, target_path

# Env + provider
try:
    from pathlib import Path
    from dotenv import find_dotenv

    # Project root (parent of /scripts or /bin, works in both cases)
    _ROOT = Path(__file__).resolve().parents[1]

    # 1) If Control Panel exposes SETTINGS_DIR, prefer that
    _SETTINGS_DIR = os.environ.get("SETTINGS_DIR", "").strip()
    _ENV_FROM_SETTINGS_DIR = Path(_SETTINGS_DIR) / ".env" if _SETTINGS_DIR else None

    # 2) Fallbacks
    _ENV_SETTINGS = _ROOT / "Settings" / ".env"
    _ENV_ROOT     = _ROOT / ".env"

    # Try in priority order
    for _p in (
        _ENV_FROM_SETTINGS_DIR if _ENV_FROM_SETTINGS_DIR and _ENV_FROM_SETTINGS_DIR.exists() else None,
        _ENV_SETTINGS if _ENV_SETTINGS.exists() else None,
        _ENV_ROOT if _ENV_ROOT.exists() else None,
    ):
        if _p:
            load_dotenv(_p, override=False, encoding="utf-8-sig")
            break
    else:
        # Last resort: search from current working directory
        p = find_dotenv(usecwd=True)
        if p:
            load_dotenv(p, override=False, encoding="utf-8-sig")
except Exception:
    pass
PROVIDER   = (os.environ.get("EXPLAIN_PROVIDER") or os.environ.get("PROVIDER", "openai")).strip().lower()
MODEL_NAME = os.environ.get("EXPLAIN_MODEL", "gpt-4o-mini")
GEM_MODEL  = os.environ.get("GEMINI_EXPLAIN_MODEL", "gemini-2.5-flash")
if not GEM_MODEL.startswith("models/"):
    GEM_MODEL = "models/" + GEM_MODEL

# Optional custom prompt from file
EXPLAIN_PROMPT_FILE = os.environ.get("EXPLAIN_PROMPT_FILE", "").strip()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
JP2TL_GLOSSARY_PATH, TL2TL_GLOSSARY_PATH = resolve_glossary_paths(PROJECT_ROOT)
JP2TL_GLOSSARY = load_glossary(JP2TL_GLOSSARY_PATH)
TL2TL_GLOSSARY = load_glossary(TL2TL_GLOSSARY_PATH)

BASE_PROMPT = """You are a friendly tutor for learners of Japanese (upper beginner to intermediate).
Input is a short Japanese line (from a JRPG). Produce a concise, readable explanation in PLAIN TEXT.
No markdown, no code fences.

Vocabulary:
- word (kana/kanji) – reading – succinct meaning; brief grammar/nuance if relevant

Grammar points:
- Particles, conjugations, set phrases; show tiny breakdowns when helpful.

Nuance & culture:
- Politeness level, speech style, cultural/cliché references if present.

Literal gloss (optional):
- A simple word-by-word gloss.

Natural English paraphrase:
- 1–2 smooth translations that fit likely context.

Key takeaways:
- 2–4 bullets to remember.

Japanese:
{jp}
"""


def gemini_safety_settings(types):
    """Disable Gemini's adjustable content filters for faithful explanation."""
    return [
        types.SafetySetting(
            category=category,
            threshold=types.HarmBlockThreshold.OFF,
        )
        for category in (
            types.HarmCategory.HARM_CATEGORY_HARASSMENT,
            types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
        )
    ]

prompt_tpl = BASE_PROMPT
if EXPLAIN_PROMPT_FILE and os.path.isfile(EXPLAIN_PROMPT_FILE):
    try:
        with open(EXPLAIN_PROMPT_FILE, "r", encoding="utf-8") as f:
            prompt_tpl = f.read()
    except Exception:
        pass

jp = read_text(LAST_JP).strip()
if not jp:
    print("(No last_jp.txt found or it is empty)", file=sys.stderr)
    sys.exit(2)

prompt = prompt_tpl.format(jp=jp)
source_glossary_prompt = build_source_glossary_prompt(JP2TL_GLOSSARY)
if source_glossary_prompt:
    prompt += "\n\n" + source_glossary_prompt

try:
    text = ""

    if PROVIDER == "gemini":
        try:
            from google import genai
            from google.genai import types
        except Exception as e:
            raise RuntimeError(
                "Missing google-genai package. Install with: python -m pip install -U google-genai"
            ) from e

        # Accept normal + local names
        api_key = (
            os.getenv("GEMINI_API_KEY")
            or os.getenv("GOOGLE_API_KEY")
            or os.getenv("GEMINI_LOCAL_KEY")
            or os.getenv("GOOGLE_LOCAL_KEY")
            or ""
        )
        api_key = api_key.strip().strip('"').strip("'")
        if not api_key:
            key_file = (os.getenv("GEMINI_API_KEY_FILE") or "").strip().strip('"').strip("'")
            if key_file and os.path.isfile(key_file):
                try:
                    with open(key_file, "r", encoding="utf-8") as kf:
                        api_key = kf.read().strip()
                except Exception:
                    pass
        if not api_key:
            raise RuntimeError("Missing GEMINI_API_KEY/GOOGLE_API_KEY (or *_LOCAL / _FILE)")

        model_name = GEM_MODEL
        if model_name.startswith("models/"):
            model_name = model_name[len("models/"):]

        client = genai.Client(api_key=api_key)
        resp = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                safety_settings=gemini_safety_settings(types),
            ),
        )

        # Handle the "blocked → no candidates" case cleanly.
        try:
            text = (getattr(resp, "text", "") or "").strip()
        except Exception:
            block_reason = None
            try:
                fb = getattr(resp, "prompt_feedback", None)
                block_reason = getattr(fb, "block_reason", None) if fb else None
            except Exception:
                pass
            if block_reason:
                text = f"(Gemini blocked the explanation; block_reason={block_reason})"
            else:
                text = "(Gemini returned no text candidates – likely blocked by safety settings.)"

        if not text:
            try:
                candidates = getattr(resp, "candidates", None) or []
                out_parts = []
                for cand in candidates:
                    content = getattr(cand, "content", None)
                    if not content:
                        continue
                    parts = getattr(content, "parts", None) or []
                    for part in parts:
                        t = getattr(part, "text", None)
                        if t:
                            out_parts.append(t)
                text = "".join(out_parts).strip()
            except Exception:
                pass
        if not text:
            try:
                prompt_feedback = getattr(resp, "prompt_feedback", None)
                if prompt_feedback:
                    text = f"(Gemini returned no text; prompt_feedback={prompt_feedback})"
            except Exception:
                pass
        if not text:
            text = "(Gemini returned no text candidates.)"

    elif PROVIDER == "openai":
        from openai import OpenAI

        # Handle BOM-prefixed names, common alternates, and *_FILE
        bom = "\ufeff"
        api_key = (
            os.getenv("OPENAI_API_KEY")
            or os.getenv(bom + "OPENAI_API_KEY")
            or os.getenv("OPENAI_LOCAL_KEY")
            or os.getenv(bom + "OPENAI_LOCAL_KEY")
            or os.getenv("OPENAI_API_KEY_LOCAL")
            or os.getenv(bom + "OPENAI_API_KEY_LOCAL")
            or os.getenv("OPENAI_KEY")
            or os.getenv(bom + "OPENAI_KEY")
            or ""
        )
        api_key = api_key.strip().strip('"').strip("'")

        if not api_key:
            key_file = (
                os.getenv("OPENAI_API_KEY_FILE")
                or os.getenv(bom + "OPENAI_API_KEY_FILE")
                or ""
            ).strip().strip('"').strip("'")
            if key_file and os.path.isfile(key_file):
                try:
                    with open(key_file, "r", encoding="utf-8") as kf:
                        api_key = kf.read().strip()
                except Exception:
                    pass

        if not api_key:
            raise RuntimeError("Missing OPENAI_API_KEY (or *_LOCAL / _FILE)")

        client = OpenAI(api_key=api_key)
        r = client.responses.create(model=MODEL_NAME, input=prompt)
        text = (getattr(r, "output_text", "") or "").strip()

    else:
        raise RuntimeError(f"Unknown provider: {PROVIDER}")

except Exception as e:
    print(f"(Python error) Explain call failed: {e}", file=sys.stderr)
    sys.exit(1)

if not text:
    text = "(No explanation returned)"

if TL2TL_GLOSSARY:
    text = apply_target_glossary(text, TL2TL_GLOSSARY, protected_text=jp)

# Always update the live explainer.txt (overlay reads this)
atomic_write_text(EXPLAINER_TXT, text)
print(f"Wrote explanation to: {EXPLAINER_TXT}")

# Optional: archive each explanation to a time-based text file
save_flag = (os.environ.get("SAVE_EXPLAINS", "0").strip() in ("1", "true", "yes"))
if save_flag:
    from datetime import datetime
    # Prefer EXPLAIN_SAVE_DIR, then SETTINGS_DIR\Explanations, then .\Settings\Explanations
    explains_dir = os.environ.get("EXPLAIN_SAVE_DIR", "").strip()
    if not explains_dir:
        settings_dir = os.environ.get("SETTINGS_DIR", "").strip()
        if settings_dir:
            explains_dir = os.path.join(settings_dir, "Explanations")
        else:
            explains_dir = os.path.join(os.getcwd(), "Settings", "Explanations")
    os.makedirs(explains_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_path = os.path.join(explains_dir, f"{ts}_explain.txt")

    # Include the JP source and the produced explanation for context
    try:
        with open(out_path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write("=== Japanese (source) ===\r\n")
            f.write(jp + "\r\n\r\n")
            f.write("=== Explanation ===\r\n")
            f.write(text + "\r\n")
        print(f"(Archived) {out_path}")
    except Exception as e:
        print(f"(Archive skipped) Could not write to {out_path}: {e}", file=sys.stderr)
