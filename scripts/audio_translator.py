import os, time, tempfile, sys, warnings
import traceback as _tb
# Quiet gRPC/absl when using Gemini (prevents “ALTS creds ignored …” spam)
os.environ.setdefault("GRPC_VERBOSITY", "ERROR")
os.environ.setdefault("GLOG_minloglevel", "2")   # 0=INFO, 1=WARNING, 2=ERROR
os.environ.setdefault("ABSL_LOG_LEVEL", "3")     # 0..3 == INFO..ERROR

from collections import deque

import numpy as np
import soundcard as sc
import soundfile as sf
from dotenv import load_dotenv

from typing import Optional

# Lazy resolver for WhisperModel (invoked only on first Local ASR use)
def _lazy_import_whisper():
    try:
        from faster_whisper import WhisperModel  # noqa: F401
        return WhisperModel
    except Exception as e:
        msg = str(e).lower()
        if isinstance(e, (FileNotFoundError, OSError)) and ("ctranslate2.dll" in msg or "dll" in msg):
            # Try silent VC++ redist install once, then retry import
            try: 
                import subprocess
                root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
                redist = os.path.join(root, "redist", "VC_redist.x64.exe")
                if os.path.exists(redist):
                    # Ask the user before triggering the installer (UAC may appear)
                    if not _confirm_vcredist_install():
                        return None  # user declined; stay in online mode / abort Local ASR

                    res = subprocess.run([redist, "/install", "/quiet", "/norestart"])
                    if res.returncode == 0:
                        from faster_whisper import WhisperModel  # noqa: F401
                        warnings.warn("Installed Microsoft VC++ runtime (x64) to enable faster-whisper.", RuntimeWarning)
                        return WhisperModel
                    # installer failed; fall through to final warning/None
            except Exception as e2:
                warnings.warn(f"VC++ runtime install failed: {type(e2).__name__}: {e2}", RuntimeWarning)
        # Final fallback: return None and let caller handle it (with a clear message)
        warnings.warn(f"faster-whisper unavailable: {type(e).__name__}: {e} (python={sys.executable})", RuntimeWarning)
        return None

# Do NOT resolve at import time—defer until Local ASR is actually used
WhisperModel = None
_FW_MODEL: Optional[object] = None

def _get_WhisperModel():
    """Resolve faster_whisper. Returns the class or None."""
    global WhisperModel
    if WhisperModel is None:
        WhisperModel = _lazy_import_whisper()
    return WhisperModel

def _confirm_vcredist_install() -> bool:
    """Ask the user before running the Microsoft VC++ installer. Returns True if user agrees."""
    try:
        import ctypes
        MB_YESNO   = 0x00000004
        MB_ICONINF = 0x00000040
        MB_TOPMOST = 0x00040000
        text  = (
            "Local ASR (faster-whisper) requires the Microsoft Visual C++ 2015–2022 "
            "Redistributable (x64).\n\n"
            "Install it now? (A Windows UAC prompt may appear.)"
        )
        title = "JRPG Translator – Enable Local ASR"
        # IDYES == 6
        return ctypes.windll.user32.MessageBoxW(
            None, text, title, MB_YESNO | MB_ICONINF | MB_TOPMOST
        ) == 6
    except Exception:
        # If the prompt fails for any reason, default to proceed (safer) and rely on installer outcome.
        return True

# --- .env bootstrap: prefer SETTINGS_DIR\.env, then Settings\.env, then root (BOM-safe) ---
try:
    from pathlib import Path
    from dotenv import find_dotenv

    # Project root = parent of /scripts
    _ROOT = Path(__file__).resolve().parents[1]

    # 1) If Control Panel exports SETTINGS_DIR, prefer that
    _SETTINGS_DIR = os.environ.get("SETTINGS_DIR", "").strip()
    _ENV_FROM_SETTINGS_DIR = Path(_SETTINGS_DIR) / ".env" if _SETTINGS_DIR else None

    # 2) Fallbacks
    _ENV_SETTINGS = _ROOT / "Settings" / ".env"
    _ENV_ROOT     = _ROOT / ".env"

    # Try in priority order
    for _p in (
        _ENV_FROM_SETTINGS_DIR if _ENV_FROM_SETTINGS_DIR and _ENV_FROM_SETTINGS_DIR.exists() else None,
        _ENV_SETTINGS if _ENV_SETTINGS.exists() else None,
        _ENV_ROOT     if _ENV_ROOT.exists()     else None,
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
    
def _get_key(*names, file_var=None):
    bom = "\ufeff"
    for n in names:
        v = os.getenv(n) or os.getenv(bom + n, "")
        if v:
            v = v.strip().strip('"').strip("'")
            if v:
                return v
    if file_var:
        p = (os.getenv(file_var) or os.getenv(bom + file_var, "")).strip().strip('"').strip("'")
        if p and os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    v = f.read().strip()
                    if v:
                        return v
            except Exception:
                pass
    return ""
   
OPENAI_API_KEY = _get_key(
    "OPENAI_API_KEY", "OPENAI_LOCAL_KEY", "OPENAI_API_KEY_LOCAL", "OPENAI_KEY",
    file_var="OPENAI_API_KEY_FILE",
)
GOOGLE_API_KEY = _get_key(
    "GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_LOCAL_KEY", "GOOGLE_LOCAL_KEY",
    file_var="GEMINI_API_KEY_FILE",
)

# Make sure the OpenAI SDK can see the key via environment
if OPENAI_API_KEY:
    os.environ["OPENAI_API_KEY"] = OPENAI_API_KEY

# (optional, extra belt-and-suspenders)
import logging
logging.getLogger("grpc").setLevel(logging.ERROR)
logging.getLogger("google").setLevel(logging.ERROR)

# Optional imports (lazy) so OpenAI/Gemini can be swapped freely
# We'll import whichever we need at runtime.
OpenAI = None
genai  = None

# ----------------- Paths for your AHK overlay (unchanged) -----------------
TEMP_DIR     = os.environ.get("TEMP") or tempfile.gettempdir()
OVERLAY_DIR  = os.path.join(TEMP_DIR, "JRPG_Overlay")
AUDIO_TXT    = os.path.join(OVERLAY_DIR, "audio.txt")        # AHK reads this
PAUSE_FLAG   = os.path.join(OVERLAY_DIR, "audio.pause")      # AHK toggle
# LOG_TXT kept for backward compat; actual writes go through audio_log() now
LOG_TXT      = os.path.join(OVERLAY_DIR, "audio_log.txt")
ERR_TXT      = os.path.join(OVERLAY_DIR, "audio_error.txt")  # FW/GPU errors get written here
os.makedirs(OVERLAY_DIR, exist_ok=True)
# --------------------------------------------------------------------------

# ----------------- Audio / VAD tunables (env overrides from AHK) ----------
def _f(envkey, default):
    v = os.environ.get(envkey, "")
    try:
        return float(v) if v else default
    except:
        return default

BLOCK_DUR    = 0.25                                   # seconds per audio block
RMS_THRESH   = _f("RMS_THRESH",     0.008)            # e.g., 0.006..0.010
MIN_SPEECH   = _f("MIN_VOICED_PCT", 0.50)             # treated as seconds
HANG_SIL     = _f("HANG_SIL",       0.25)
MAX_SEG_DUR  = 6.0
CAPTURE_RATE = 16000
PREROLL_SEC  = 0.20
# --------------------------------------------------------------------------

# ----------------- Provider + models (env from Control Panel) -------------
AUDIO_PROVIDER   = os.environ.get("AUDIO_PROVIDER", "openai").lower()  # "openai" or "gemini"

# OpenAI model names (from AHK)
ASR_MODEL        = os.environ.get("ASR_MODEL",       "gpt-4o-mini-transcribe")
TRANSLATE_MODEL  = os.environ.get("TRANSLATE_MODEL", "gpt-4o-mini")

# Local ASR (faster-whisper) options
FW_MODEL_NAME = (os.environ.get("FW_MODEL_NAME") or "small").strip()   # tiny|base|small|medium|large-v3 or CT2 dir
FW_COMPUTE    = (os.environ.get("FW_COMPUTE")    or "auto").strip()    # auto|int8|int8_float16|float16|float32
FW_CACHE_DIR  = (os.environ.get("FW_CACHE_DIR")  or "").strip() or None

# Gemini model (used for BOTH transcribe & translate when provider="gemini")
GEMINI_AUDIO_MODEL = os.environ.get("GEMINI_AUDIO_MODEL", "gemini-2.5-flash")
# Optional custom prompt injected by Control Panel (file or inline)
AUDIO_PROMPT_FILE = os.environ.get("AUDIO_PROMPT_FILE", "").strip()
AUDIO_PROMPT_ENV  = os.environ.get("AUDIO_PROMPT", "").strip()
# Optional: force a specific playback device by (substring) name.
# Leave empty to use the Windows Default Output device.
SPEAKER_NAME = os.environ.get("SPEAKER_NAME", "").strip()
# Debug logging (AHK sets JRPG_DEBUG to "1" or "0")
DEBUG = os.environ.get("JRPG_DEBUG", "0").strip() == "1"

# ---- Glossaries (profile-based only; no legacy defaults) ----
SCRIPT_DIR = os.path.dirname(__file__)
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
JP2EN_GLOSSARY_PATH = (os.environ.get("JP2EN_GLOSSARY_PATH", "").strip() or None)
EN2EN_GLOSSARY_PATH = (os.environ.get("EN2EN_GLOSSARY_PATH", "").strip() or None)

try:
    import configparser

    def _find_settings_ini(base_dir: str):
        cands = [
            os.path.join(base_dir, "Settings", "control.ini"),
            os.path.join(os.path.dirname(base_dir), "Settings", "control.ini"),
            os.path.join(base_dir, "Settings", "config.ini"),
            os.path.join(os.path.dirname(base_dir), "Settings", "config.ini"),
        ]
        for p in cands:
            if os.path.isfile(p):
                return p
        return None

    if JP2EN_GLOSSARY_PATH is None or EN2EN_GLOSSARY_PATH is None:
        cfg_path = _find_settings_ini(SCRIPT_DIR)
        prof_jp = prof_en = "default"

        def _read_profiles_from_ini(_path: str):
            encs = ["utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "cp932"]
            for enc in encs:
                try:
                    cfg = configparser.ConfigParser()
                    with open(_path, "r", encoding=enc) as fh:
                        cfg.read_file(fh)
                    jp = (cfg.get("cfg", "jp2enGlossaryProfile", fallback="default").strip() or "default")
                    en = (cfg.get("cfg", "en2enGlossaryProfile", fallback="default").strip() or "default")
                    return jp, en
                except Exception:
                    continue
            return "default", "default"

        if cfg_path and os.path.isfile(cfg_path):
            prof_jp, prof_en = _read_profiles_from_ini(cfg_path)

        settings_dir = os.path.dirname(cfg_path) if cfg_path else os.path.join(SCRIPT_DIR, "Settings")
        gloss_base   = os.path.join(settings_dir, "glossaries")

        cand_jp = os.path.join(gloss_base, prof_jp, "jp2en.txt")
        cand_en = os.path.join(gloss_base, prof_en, "en2en.txt")

        if JP2EN_GLOSSARY_PATH is None:
            JP2EN_GLOSSARY_PATH = cand_jp
        if EN2EN_GLOSSARY_PATH is None:
            EN2EN_GLOSSARY_PATH = cand_en

        # --- DEBUG: show profiles and the exact files we will use ---
        try:
            if DEBUG:
                jp_exists = "(exists)" if (JP2EN_GLOSSARY_PATH and os.path.isfile(JP2EN_GLOSSARY_PATH)) else ("(None)" if not JP2EN_GLOSSARY_PATH else "(MISSING)")
                en_exists = "(exists)" if (EN2EN_GLOSSARY_PATH and os.path.isfile(EN2EN_GLOSSARY_PATH)) else ("(None)" if not EN2EN_GLOSSARY_PATH else "(MISSING)")
                append_log(
                    "Audio glossary resolution:\n"
                    f"  control.ini: {cfg_path or '(not found)'}\n"
                    f"  jp2enProfile: {prof_jp}  -> {JP2EN_GLOSSARY_PATH or '(None)'}  {jp_exists}\n"
                    f"  en2enProfile: {prof_en}  -> {EN2EN_GLOSSARY_PATH or '(None)'}  {en_exists}"
                )
        except Exception:
            pass
except Exception:
    # Never let glossary resolution break the audio worker
    pass

from typing import List, Tuple
import re

def load_glossary(path) -> List[Tuple[str, str]]:
    """Load 'source -> target' pairs.
       - Accept separators: '->', '→', tab, ':', '='
       - Accept encodings: utf-8, utf-8-sig, utf-16, utf-16-le, utf-16-be, cp932, cp1252
       - Ignore blanks and lines starting with '#'
    """
    entries: List[Tuple[str, str]] = []
    if not path or not os.path.isfile(path):
        return entries

    encodings = ["utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "cp932"]
    text = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as f:
                text = f.read()
            break
        except Exception:
            continue
    if text is None:
        return entries  # unreadable; caller will log counts as 0

    seps = ["->", "→", "\t", ":", "="]
    for raw in text.splitlines():
        line = raw.replace("\ufeff", "").strip()
        if not line or line.startswith("#"):
            continue
        src = dst = None
        for sep in seps:
            if sep in line:
                s, d = line.split(sep, 1)
                src, dst = s.strip(), d.strip()
                break
        if src and dst:
            entries.append((src, dst))
    return entries

def apply_en_glossary(text: str, en2en: List[Tuple[str, str]]) -> str:
    """Apply EN→EN replacements.
       - Phrases: literal, case-insensitive
       - Single tokens: whole word + optional plural/possessive suffix kept (s, 's, ’s)
    """
    out = text
    for src, dst in en2en:
        if re.search(r"\s", src):
            # phrase: simple case-insensitive replace
            pattern = re.compile(re.escape(src), flags=re.IGNORECASE)
            out = pattern.sub(dst, out)
        else:
            # single token: allow plural/possessive and preserve suffix
            pattern = re.compile(
                rf"\b{re.escape(src)}(?P<suf>s|\'s|’s)?\b",
                flags=re.IGNORECASE
            )
            out = pattern.sub(lambda m: dst + (m.group("suf") or ""), out)
    return out

def build_jp2en_prompt(jp2en: List[Tuple[str, str]]) -> str:
    if not jp2en:
        return ""
    lines = ["When translating, apply these exact JP→EN mappings wherever they appear:"]
    for jp, en in jp2en:
        lines.append(f"- {jp} → {en}")
    return "\n".join(lines)

# Loaded once in main(); read by translate_* functions
JP2EN = []   # type: List[Tuple[str, str]]
EN2EN = []   # type: List[Tuple[str, str]]

# --- Glossary hot-reload (safe/optional) ---
JP2EN_MTIME = None  # type: Optional[float]
EN2EN_MTIME = None  # type: Optional[float]

def _file_mtime_or_none(p: Optional[str]) -> Optional[float]:
    try:
        return os.path.getmtime(p) if (p and os.path.isfile(p)) else None
    except Exception:
        return None

def _reload_glossaries_if_changed():
    """Reload JP2EN/EN2EN if their files changed."""
    global JP2EN, EN2EN, JP2EN_MTIME, EN2EN_MTIME
    try:
        cur_jp = _file_mtime_or_none(JP2EN_GLOSSARY_PATH)
        cur_en = _file_mtime_or_none(EN2EN_GLOSSARY_PATH)
        changed = False

        if cur_jp is not None and cur_jp != JP2EN_MTIME:
            JP2EN = load_glossary(JP2EN_GLOSSARY_PATH)
            JP2EN_MTIME = cur_jp
            changed = True

        if cur_en is not None and cur_en != EN2EN_MTIME:
            EN2EN = load_glossary(EN2EN_GLOSSARY_PATH)
            EN2EN_MTIME = cur_en
            changed = True

        if changed and DEBUG:
            append_log(f"Reloaded glossaries — JP2EN:{len(JP2EN)} EN2EN:{len(EN2EN)}")
    except Exception as e:
        append_log(f"Glossary hot-reload error: {e}")
# --- end hot-reload ---

def _rotate_if_big(path, max_bytes=2_000_000, backups=3):
    try:
        import os
        if os.path.getsize(path) <= max_bytes:
            return
    except FileNotFoundError:
        return
    except Exception:
        return
    for i in range(backups - 1, 0, -1):
        src = f"{path}.{i}"
        dst = f"{path}.{i+1}"
        try:
            if os.path.exists(src):
                os.replace(src, dst)
        except Exception:
            pass
    try:
        os.replace(path, f"{path}.1")
    except Exception:
        pass

def audio_log(line: str):
    """Append a line to audio_log.txt only if DEBUG is enabled."""
    if not DEBUG:
        return
    try:
        import os, tempfile
        folder = os.path.join(tempfile.gettempdir(), "JRPG_Overlay")
        os.makedirs(folder, exist_ok=True)
        log_path = os.path.join(folder, "audio_log.txt")
        _rotate_if_big(log_path)
        with open(log_path, "a", encoding="utf-8", errors="ignore") as f:
            f.write(line.rstrip("\r\n") + "\n")
    except Exception:
        # never let logging break the worker
        pass

def _pick_speaker():
    import soundcard as sc
    if SPEAKER_NAME:
        name_l = SPEAKER_NAME.lower()
        for s in sc.all_speakers():
            if name_l in s.name.lower():
                return s
    return sc.default_speaker()

DEFAULT_AUDIO_PROMPT = (
    "Translate the following Japanese into clear, natural English. "
    "Keep it concise. Do not add or omit meaning.\n\n"
    "Japanese:\n{JP_TEXT}"
)

def _load_audio_prompt_template() -> str:
    # 1) from file
    p = AUDIO_PROMPT_FILE
    if p and os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                t = f.read().strip()
                if t:
                    return t
        except Exception:
            pass
    # 2) from env
    if AUDIO_PROMPT_ENV:
        return AUDIO_PROMPT_ENV
    # 3) fallback
    return DEFAULT_AUDIO_PROMPT

def _render_audio_prompt(jp_text: str, jp2en: list) -> str:
    tpl = _load_audio_prompt_template()
    glossary_block = build_jp2en_prompt(jp2en)
    if glossary_block:
        tpl = f"{tpl}\n\n{glossary_block}"

    # support either explicit placeholder or auto-append
    if "{JP_TEXT}" in tpl:
        try:
            return tpl.replace("{JP_TEXT}", jp_text)
        except Exception:
            pass
    # fallback append
    return f"{tpl}\n\nJapanese:\n{jp_text}"

LANG_HINT        = "ja"
# --------------------------------------------------------------------------

# Silence benign discontinuity warnings from soundcard on HDMI/AVR
warnings.filterwarnings("ignore", message="data discontinuity in recording")

# ----------------- File helpers -----------------
def atomic_write_text(path: str, text: str):
    """UTF-8 atomic write so AHK never reads partial content (tolerant of brief locks)."""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)
    for _ in range(10):  # ~500 ms total
        try:
            os.replace(tmp, path)
            break
        except PermissionError:
            time.sleep(0.05)
    else:
        # Fallback (non-atomic) to avoid leaving '.tmp' behind
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(text)
        try:
            os.remove(tmp)
        except Exception:
            pass

def append_log(line: str):
    # Route all logging through audio_log(), which honors DEBUG and rotates.
    audio_log(line)

# -----------------------------------------------

def write_wav(path, audio_mono_np, sr):
    sf.write(path, audio_mono_np, sr, subtype="PCM_16")

def _gemini_safety_settings():
    """
    Return relaxed safety_settings for google.generativeai, or None if the
    safety_types module isn't available.

    NOTE: If your account ever complains about BLOCK_NONE being 'restricted',
    change BLOCK_NONE below to BLOCK_ONLY_HIGH.
    """
    try:
        from google.generativeai.types.safety_types import HarmBlockThreshold, HarmCategory

        return {
            HarmCategory.HARM_CATEGORY_HARASSMENT:        HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH:       HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
        }
    except Exception:
        # If anything goes wrong, just fall back to default safety settings.
        return None

# ----------------- Transcription backends -----------------
def transcribe_openai(audio_mono_np, sr):
    global OpenAI, OPENAI_API_KEY
    if OpenAI is None:
        from openai import OpenAI as _OpenAI
        OpenAI = _OpenAI
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY missing. Add it to .env or environment.")

    client = OpenAI(api_key=OPENAI_API_KEY)

    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        write_wav(tmp.name, audio_mono_np, sr)
        tmp_path = tmp.name
    try:
        with open(tmp_path, "rb") as f:
            tr = client.audio.transcriptions.create(
                model=ASR_MODEL,  # e.g., gpt-4o-mini-transcribe, whisper-1
                file=f,
                language=LANG_HINT
            )
        return getattr(tr, "text", "").strip()
    finally:
        try: os.remove(tmp_path)
        except: pass

def transcribe_gemini(audio_mono_np, sr):
    """
    Gemini is multimodal: we pass audio bytes + instruction and get Japanese text back.
    """
    global genai, GOOGLE_API_KEY, GEMINI_AUDIO_MODEL
    if genai is None:
        import google.generativeai as _genai
        genai = _genai
    if not GOOGLE_API_KEY:
        raise RuntimeError("GOOGLE_API_KEY missing. Add it to .env or environment.")

    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel(GEMINI_AUDIO_MODEL)

    safety_settings = _gemini_safety_settings()

    # Write a short WAV and read bytes (MediaFoundation likes PCM16 WAV)
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        write_wav(tmp.name, audio_mono_np, sr)
        tmp_path = tmp.name
    try:
        with open(tmp_path, "rb") as f:
            audio_bytes = f.read()
        prompt = (
            "Transcribe the Japanese speech accurately AS JAPANESE TEXT only. "
            "Do not translate. No explanations."
        )

        if safety_settings is not None:
            resp = model.generate_content(
                [
                    {"mime_type": "audio/wav", "data": audio_bytes},
                    prompt,
                ],
                safety_settings=safety_settings,
            )
        else:
            resp = model.generate_content(
                [
                    {"mime_type": "audio/wav", "data": audio_bytes},
                    prompt,
                ]
            )

        try:
            text = (getattr(resp, "text", "") or "").strip()
        except Exception as e:
            # If Gemini blocks the audio, just log (in DEBUG) and return empty string
            block_reason = None
            try:
                fb = getattr(resp, "prompt_feedback", None)
                block_reason = getattr(fb, "block_reason", None) if fb else None
            except Exception:
                pass
            if DEBUG:
                if block_reason:
                    append_log(f"[GEMINI] Transcribe blocked; reason={block_reason}")
                else:
                    append_log(f"[GEMINI] Transcribe: no text candidates; {e}")
            return ""

        return text
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
# ----------------------------------------------------------

# ----------------- Translation backends -------------------
def translate_openai(jp_text):
    global OpenAI, OPENAI_API_KEY
    if not jp_text:
        return ""
    if OpenAI is None:
        from openai import OpenAI as _OpenAI
        OpenAI = _OpenAI
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY missing. Add it to .env or environment.")

    client = OpenAI(api_key=OPENAI_API_KEY)
    prompt = _render_audio_prompt(jp_text, JP2EN)
    if DEBUG:
        # show if JP→EN block made it into the prompt
        append_log(f"[OPENAI] JP2EN in prompt: {bool(JP2EN)}; prompt preview:\n{prompt[:500]}")
    resp = client.responses.create(model=TRANSLATE_MODEL, input=prompt)
    en = (getattr(resp, "output_text", "") or "").strip()
    if EN2EN:
        en = apply_en_glossary(en, EN2EN)
        if DEBUG:
            append_log("[OPENAI] EN after EN2EN:\n" + en[:400])
    return en

def translate_gemini(jp_text):
    global genai, GOOGLE_API_KEY, GEMINI_AUDIO_MODEL
    if not jp_text:
        return ""
    if genai is None:
        import google.generativeai as _genai
        genai = _genai
    if not GOOGLE_API_KEY:
        raise RuntimeError("GOOGLE_API_KEY missing. Add it to .env or environment.")

    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel(GEMINI_AUDIO_MODEL)
    prompt = _render_audio_prompt(jp_text, JP2EN)
    if DEBUG:
        append_log(f"[GEMINI] JP2EN in prompt: {bool(JP2EN)}; prompt preview:\n{prompt[:500]}")

    safety_settings = _gemini_safety_settings()

    if safety_settings is not None:
        resp = model.generate_content(prompt, safety_settings=safety_settings)
    else:
        resp = model.generate_content(prompt)

    try:
        en = (getattr(resp, "text", "") or "").strip()
    except Exception as e:
        block_reason = None
        try:
            fb = getattr(resp, "prompt_feedback", None)
            block_reason = getattr(fb, "block_reason", None) if fb else None
        except Exception:
            pass

        if block_reason:
            en = f"(Gemini blocked the translation; block_reason={block_reason})"
        else:
            en = "(Gemini returned no text candidates – likely blocked by safety settings.)"

        if DEBUG:
            append_log(f"[GEMINI] translate_gemini text accessor failed: {e}; reason={block_reason}")

    if EN2EN:
        en = apply_en_glossary(en, EN2EN)
        if DEBUG:
            append_log("[GEMINI] EN after EN2EN:\n" + en[:400])
    return en
# ----------------------------------------------------------
# ===== Local ASR backend: faster-whisper (CTranslate2) =====
def fw_load_model(model_name: str, device_pref: str = "auto", compute: str = "auto", cache_dir=None):
    """Load a faster-whisper model once (lazy-resolves WhisperModel on first use)."""
    global _FW_MODEL
    WM = _get_WhisperModel()
    if WM is None:
        # write a clear reason to audio_error.txt, then raise
        try:
            with open(ERR_TXT, "a", encoding="utf-8", errors="ignore") as _f:
                _f.write(
                    "WhisperModel unresolved (faster-whisper failed to import).\n"
                    f"Python: {sys.executable}\n"
                )
        except Exception:
            pass
        raise RuntimeError(
            "Local ASR requires faster-whisper, but it failed to load. "
            f"See audio_error.txt for details (python={sys.executable})."
        )

    # Decide device
    use_cuda = False
    if device_pref in ("auto", "cuda"):
        try:
            import torch
            use_cuda = torch.cuda.is_available()
        except Exception:
            use_cuda = False
    device = "cuda" if use_cuda else "cpu"

    if compute == "auto":
        compute_type = "float16" if use_cuda else "int8"
    else:
        compute_type = compute

    _FW_MODEL = WM(
        model_size_or_path=model_name,
        device=device,
        compute_type=compute_type,
        download_root=cache_dir
    )

def _resample_to_16k(x: np.ndarray, sr: int) -> np.ndarray:
    """Lightweight linear resample to 16 kHz."""
    target = 16000
    if sr == target or x.size == 0:
        return x.astype(np.float32, copy=False)
    dur = x.shape[0] / float(sr)
    t_old = np.linspace(0.0, dur, num=x.shape[0], endpoint=False, dtype=np.float64)
    t_new = np.linspace(0.0, dur, num=int(round(dur * target)), endpoint=False, dtype=np.float64)
    y = np.interp(t_new, t_old, x.astype(np.float32, copy=False))
    return y.astype(np.float32, copy=False)

def fw_transcribe_float32(seg: np.ndarray, sr: int) -> str:
    """
    seg: mono float32 PCM (-1..1), length = one finalized segment from your VAD.
    Returns JP text.
    """
    if _FW_MODEL is None:
        raise RuntimeError("faster-whisper model not loaded")

    # Ensure 16 kHz mono for faster-whisper
    if seg.ndim > 1:
        seg = seg.mean(axis=1).astype(np.float32, copy=False)
    if sr != 16000:
        seg = _resample_to_16k(seg, sr)

    segments, info = _FW_MODEL.transcribe(
        seg,
        language="ja",
        vad_filter=True,
        beam_size=5,
        no_speech_threshold=0.4,
    )
    return "".join(s.text for s in segments).strip()

def capture_blocks(sc_speaker, samplerate, block_dur):
    """Yield (frames, channels) blocks from default speaker loopback."""
    block_size = int(samplerate * block_dur)
    mic = sc.get_microphone(id=sc_speaker.name, include_loopback=True)
    # Explicit stereo capture is robust across HDMI/AVR
    with mic.recorder(samplerate=samplerate, channels=2) as rec:
        while True:
            data = rec.record(numframes=block_size)  # float32, shape: (frames, channels)
            yield data

def main():
    # Clear old line at start so overlay isn’t stale
    try: atomic_write_text(AUDIO_TXT, "")
    except Exception: pass

    # Pick backend functions
    text_provider = (os.environ.get("TEXT_PROVIDER", "").strip().lower() or None)

            # --- Select transcriber and translator ---------------------------------
            # --- Select transcriber and translator ---------------------------------
    if AUDIO_PROVIDER == "gemini":
        # Online transcription via Gemini
        transcribe_fn = transcribe_gemini
        translate_fn  = translate_gemini if (text_provider == "gemini") else translate_openai

    elif AUDIO_PROVIDER in ("local", "faster-whisper", "fw"):
        # Local transcription via faster-whisper (lazy import & self-heal happens inside fw_load_model)
        # One-time probe: record context to audio_error.txt so failures are self-diagnosing
        try:
            with open(ERR_TXT, "w", encoding="utf-8", errors="ignore") as _f:
                _f.write(
                    "Starting Local ASR\n"
                    f"Python: {sys.executable}\n"
                    f"Model : {FW_MODEL_NAME}  compute={FW_COMPUTE}  cache={FW_CACHE_DIR or 'default'}\n"
                )
        except Exception:
            pass

        try:
            fw_load_model(FW_MODEL_NAME,
                          device_pref="auto",
                          compute=FW_COMPUTE,
                          cache_dir=FW_CACHE_DIR)
        except Exception as e:
            det = (f"FW load failed: {e}\n"
                   f"Model={FW_MODEL_NAME}  Compute={FW_COMPUTE}  Cache={FW_CACHE_DIR or 'default'}\n")
            print(det)
            print(_tb.format_exc())
            try:
                atomic_write_text(ERR_TXT, det + "\n" + _tb.format_exc())
                atomic_write_text(AUDIO_TXT, "FW error — see audio_error.txt")
            except Exception:
                pass
            time.sleep(0.5)
            sys.exit(1)

        transcribe_fn = fw_transcribe_float32
        # Translator independent of transcriber
        translate_fn  = translate_gemini if (text_provider == "gemini") else translate_openai

    else:
        # Online transcription via OpenAI (whisper-1 / 4o-mini-transcribe)
        transcribe_fn = transcribe_openai
        # Translator independent of transcriber
        translate_fn  = translate_gemini if (text_provider == "gemini") else translate_openai

    print("\n=== Live JP → EN Subtitles → AHK overlay (audio.txt) ===")
    print(f"Overlay dir : {OVERLAY_DIR}")
    print(f"Overlay file: {AUDIO_TXT}")
    print(f"Pause flag  : {PAUSE_FLAG}")

    import sys as _sys
    print(f"Python exec : {_sys.executable}")

    print(f"Provider    : {AUDIO_PROVIDER}")


    if AUDIO_PROVIDER == "gemini":
        print(f"Gemini model: {GEMINI_AUDIO_MODEL}")
    elif AUDIO_PROVIDER in ("local", "faster-whisper", "fw"):
        print(f"Local ASR   : faster-whisper ({FW_MODEL_NAME}, compute={FW_COMPUTE}, cache={FW_CACHE_DIR or 'default'})")
        print(f"Text TR     : {TRANSLATE_MODEL} (OpenAI)")
    else:
        print(f"OpenAI ASR  : {ASR_MODEL}")
        print(f"OpenAI TR   : {TRANSLATE_MODEL}")
        print("Tip: Set your DENON-AVR as Windows **default** output device.\n")
            # --- Resolve glossary paths (prefer Control Panel's SETTINGS_DIR) ---
    global JP2EN_GLOSSARY_PATH, EN2EN_GLOSSARY_PATH
    try:
        import configparser
        env_settings_dir = (os.environ.get("SETTINGS_DIR", "").strip() or None)

        # Find Settings folder
        if env_settings_dir:
            settings_dir = env_settings_dir
            cfg_path = os.path.join(settings_dir, "control.ini")
            if not os.path.isfile(cfg_path):
                cfg_path = os.path.join(settings_dir, "config.ini")
        else:
            # Try to locate a Settings near the project root
            cfg_path = None
            for cand in (
                os.path.join(PROJECT_ROOT, "Settings", "control.ini"),
                os.path.join(PROJECT_ROOT, "Settings", "config.ini"),
                os.path.join(os.path.dirname(SCRIPT_DIR), "Settings", "control.ini"),
                os.path.join(os.path.dirname(SCRIPT_DIR), "Settings", "config.ini"),
            ):
                if os.path.isfile(cand):
                    cfg_path = cand
                    break
            settings_dir = os.path.dirname(cfg_path) if cfg_path else os.path.join(PROJECT_ROOT, "Settings")

        prof_jp = prof_en = "default"
        if cfg_path and os.path.isfile(cfg_path):
            def _read_profiles_from_ini(_path: str):
                encs = ["utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "cp932"]
                for enc in encs:
                    try:
                        cfg = configparser.ConfigParser()
                        with open(_path, "r", encoding=enc) as fh:
                            cfg.read_file(fh)
                        jp = (cfg.get("cfg", "jp2enGlossaryProfile", fallback="default") or "default").strip()
                        en = (cfg.get("cfg", "en2enGlossaryProfile", fallback="default") or "default").strip()
                        return jp, en
                    except Exception:
                        continue
                return "default", "default"
            prof_jp, prof_en = _read_profiles_from_ini(cfg_path)

        # Only fill paths if not already provided via env
        if not JP2EN_GLOSSARY_PATH:
            JP2EN_GLOSSARY_PATH = os.path.join(settings_dir, "glossaries", prof_jp, "jp2en.txt")
        if not EN2EN_GLOSSARY_PATH:
            EN2EN_GLOSSARY_PATH = os.path.join(settings_dir, "glossaries", prof_en, "en2en.txt")

        if DEBUG:
            jp_ok = "(exists)" if (JP2EN_GLOSSARY_PATH and os.path.isfile(JP2EN_GLOSSARY_PATH)) else ("(None)" if not JP2EN_GLOSSARY_PATH else "(MISSING)")
            en_ok = "(exists)" if (EN2EN_GLOSSARY_PATH and os.path.isfile(EN2EN_GLOSSARY_PATH)) else ("(None)" if not EN2EN_GLOSSARY_PATH else "(MISSING)")
            append_log(
                "Audio glossary resolution:\n"
                f"  settings_dir : {settings_dir}\n"
                f"  control.ini  : {cfg_path or '(not found)'}\n"
                f"  jp2enProfile : {prof_jp} -> {JP2EN_GLOSSARY_PATH or '(None)'} {jp_ok}\n"
                f"  en2enProfile : {prof_en} -> {EN2EN_GLOSSARY_PATH or '(None)'} {en_ok}"
            )
    except Exception as e:
        if DEBUG:
            append_log(f"Glossary resolution error: {e}")
    # --- end resolver ---

    # Load glossaries once
    global JP2EN, EN2EN, JP2EN_MTIME, EN2EN_MTIME
    try:
        JP2EN = load_glossary(JP2EN_GLOSSARY_PATH)
        EN2EN = load_glossary(EN2EN_GLOSSARY_PATH)
        JP2EN_MTIME = _file_mtime_or_none(JP2EN_GLOSSARY_PATH)
        EN2EN_MTIME = _file_mtime_or_none(EN2EN_GLOSSARY_PATH)
        if DEBUG:
            preview_jp = ", ".join([f"{a}→{b}" for a,b in JP2EN[:5]])
            preview_en = ", ".join([f"{a}→{b}" for a,b in EN2EN[:5]])
            append_log(
                f"Loaded JP2EN: {len(JP2EN)} entries [{preview_jp}]\n"
                f"Loaded EN2EN: {len(EN2EN)} entries [{preview_en}]"
            )
    except Exception as e:
        append_log(f"Glossary load error: {e}")
        JP2EN, EN2EN = [], []

    speaker = _pick_speaker()
    if speaker is None:
        print("No playback device found. Check Windows sound settings.")
        sys.exit(1)

    print(f"Using default speaker: {speaker.name}  at {CAPTURE_RATE} Hz")
    print("Listening…  (Press Ctrl+C to stop)\n")

    # Pre-roll: short history before speech begins
    preroll_blocks = max(1, int(PREROLL_SEC / BLOCK_DUR))
    history = deque(maxlen=preroll_blocks)

    speech, speech_len, silence_run, in_voiced = [], 0.0, 0.0, False
    last_written = ""  # avoid duplicate writes

    try:
        for data in capture_blocks(speaker, CAPTURE_RATE, BLOCK_DUR):
            # Mix to mono float32
            if data.ndim == 2 and data.shape[1] > 1:
                mono = data.mean(axis=1).astype(np.float32, copy=False)
            else:
                mono = data.reshape(-1).astype(np.float32, copy=False)

            history.append(mono)

            # naive RMS VAD
            rms = float(np.sqrt(np.mean(mono**2)) + 1e-12)
            is_voiced = rms >= RMS_THRESH

            if is_voiced:
                if not in_voiced and len(history):
                    speech.extend(list(history))  # prepend pre-roll
                speech.append(mono)
                speech_len += BLOCK_DUR
                silence_run = 0.0
                in_voiced = True
            else:
                if in_voiced:
                    silence_run += BLOCK_DUR

            # end segment?
            should_cut = (
                in_voiced and
                (silence_run >= HANG_SIL or speech_len >= MAX_SEG_DUR) and
                speech_len >= MIN_SPEECH
            )

            if should_cut:
                seg = np.concatenate(speech, axis=0) if speech else np.empty(0, np.float32)
                seg = np.clip(seg, -1.0, 1.0).astype(np.float32)

                try:
                    jp_text = transcribe_fn(seg, CAPTURE_RATE)
                except Exception as e:
                    ts = time.strftime("%H:%M:%S")
                    msg = f"[{ts}] Transcribe error: {e}"
                    print(msg)
                    append_log(msg)
                    jp_text = ""

                if jp_text:
                    # Pick up any jp2en/en2en changes without restart
                    _reload_glossaries_if_changed()
                    try:
                        en_text = translate_fn(jp_text)
                    except Exception as e:
                        ts = time.strftime("%H:%M:%S")
                        msg = f"[{ts}] Translate error: {e}"
                        print(msg)
                        append_log(msg)
                        en_text = ""
                    ts = time.strftime("%H:%M:%S")
                    print(f"[{ts}] JA: {jp_text}")
                    print(f"[{ts}] EN: {en_text}\n")

                    if not os.path.exists(PAUSE_FLAG):
                        overlay_line = en_text  # (change to f"{jp_text}\r\n{en_text}" if you want both)
                        if overlay_line and overlay_line != last_written:
                            atomic_write_text(AUDIO_TXT, overlay_line)
                            last_written = overlay_line
                        append_log(f"[{ts}] EN: {en_text}")
                        # append_log(f"[{ts}] JA: {jp_text}")  # add if you want JA in log too

                # reset segment state
                speech, speech_len, silence_run, in_voiced = [], 0.0, 0.0, False
    except KeyboardInterrupt:
        print("\nStopped. Bye!")
    except Exception as e:
        msg = (
            "\nAudio failed to run.\n"
            f"Reason: {e}\n\n"
            "Tips:\n"
            " • Ensure your default output device is on (e.g., AVR/HDMI).\n"
            " • Or pick another default device in Windows Sound settings.\n"
            " • Or set a fixed 'Speaker name' in Settings > Audio.\n"
            " • Use --list-speakers to see device names.\n"
        )
        print(msg)
        append_log(msg)
        sys.exit(1)

if __name__ == "__main__":
    import sys
    if "--list-speakers" in sys.argv:
        # Print one name per line; force UTF-8 to avoid code page issues
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
        except Exception:
            pass
        import soundcard as sc
        for s in sc.all_speakers():
            print(s.name)
        sys.exit(0)
    main()
