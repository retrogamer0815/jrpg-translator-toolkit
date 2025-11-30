#!/usr/bin/env python
# -*- coding: utf-8 -*-

import io, os, sys, tempfile
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

try:
    text = ""

    if PROVIDER == "gemini":
        import google.generativeai as genai

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

        # Try to relax safety filters so eroge / adult dialogue aren't blocked.
        safety_settings = None
        try:
            from google.generativeai.types.safety_types import HarmBlockThreshold, HarmCategory

            # NOTE:
            # If your account complains about BLOCK_NONE being "restricted",
            # change BLOCK_NONE below to BLOCK_ONLY_HIGH.
            safety_settings = {
                HarmCategory.HARM_CATEGORY_HARASSMENT:         HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_HATE_SPEECH:        HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT:  HarmBlockThreshold.BLOCK_NONE,
                HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT:  HarmBlockThreshold.BLOCK_NONE,
            }
        except Exception:
            safety_settings = None

        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(GEM_MODEL)

        if safety_settings is not None:
            resp = model.generate_content(
                prompt,
                generation_config={"temperature": 0.2},
                safety_settings=safety_settings,
            )
        else:
            resp = model.generate_content(
                prompt,
                generation_config={"temperature": 0.2},
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
