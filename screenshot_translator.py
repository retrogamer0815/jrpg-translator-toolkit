#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
JRPG image translator with dual glossaries, Transcript/Translation output,
and speaker-name header normalization (kana→romaji fallback).

Usage:
  python screenshot_translator.py <image1> [<image2> ...]

Environment:
  PROVIDER                (optional) "openai" (default) or "gemini"

  --- OpenAI ---
  OPENAI_API_KEY          (required if PROVIDER=openai)
  MODEL_NAME              (optional, default: "gpt-4o")

  --- Google Gemini ---
  GEMINI_API_KEY          (required if PROVIDER=gemini)
  GEMINI_MODEL_NAME       (optional, default: "gemini-1.5-flash")

  --- Glossaries ---
  JP2EN_GLOSSARY_PATH     (optional, default: ./jp2en_glossary.txt)
  EN2EN_GLOSSARY_PATH     (optional, default: ./en2en_glossary.txt)

  --- Prompt override ---
  PROMPT_FILE             (optional) path to a UTF-8 prompt file
  PROMPT_TEXT             (optional) full prompt string (takes precedence over file)
  POSTPROC_MODE           (optional) "tt" (default) | "translation" | "none"
  PROMPT_PROFILE          (optional) name of a prompt in Settings/prompts/<name>.txt
                           or prompts/<name>.txt (overridden by PROMPT_FILE/TEXT)
"""

# --- Imports must come before using os.environ ---
import io
import os
import re
import sys
import base64
import mimetypes
import time
import tempfile
from typing import List, Tuple

# --- .env bootstrap: prefer SETTINGS_DIR\.env, then Settings\.env, then root (BOM-safe) ---
try:
    from dotenv import load_dotenv, find_dotenv
    from pathlib import Path
    # Project root = parent of /scripts
    _ROOT = Path(__file__).resolve().parents[1]

    # 1) If Control Panel exports SETTINGS_DIR, prefer that
    _SETTINGS_DIR = (os.environ.get("SETTINGS_DIR", "") or "").strip()
    _ENV_FROM_SETTINGS_DIR = (Path(_SETTINGS_DIR) / ".env") if _SETTINGS_DIR else None

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
        # Last resort: search from current working dir
        p = find_dotenv(usecwd=True)
        if p:
            load_dotenv(p, override=False, encoding="utf-8-sig")
except Exception:
    pass
    
def _get_key(*names, file_var=None):
    """Return first non-empty env var among names; if none, read from file path var."""
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

# (dotenv already loaded above; no second pass needed)
# --- Paths shared with the AHK overlay (same folder audio uses) ---
TEMP_DIR    = os.environ.get("TEMP") or tempfile.gettempdir()
OVERLAY_DIR = os.path.join(TEMP_DIR, "JRPG_Overlay")
LAST_JP     = os.path.join(OVERLAY_DIR, "last_jp.txt")
# (optional) keep track of the source images too
LAST_SRC    = os.path.join(OVERLAY_DIR, "last_src.txt")
OCR_TXT     = os.path.join(OVERLAY_DIR, "ocr.txt")
os.makedirs(OVERLAY_DIR, exist_ok=True)

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

# ---- Post-processing mode -----------------------------------------------------
# "tt" (default): enforce Transcript/Translation and normalize headers
# "translation": return only the Translation block (no headings)
# "none": return raw model text (no headings, no normalization)
POSTPROC_MODE = (
    os.environ.get("POSTPROC_MODE")
    or os.environ.get("SHOT_POSTPROC")
    or "tt"
).strip().lower()

# ---- Console UTF-8 on Windows ------------------------------------------------
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ---- Provider selection -------------------------------------------------------
PROVIDER = os.environ.get("PROVIDER", "openai").strip().lower()

# ---- OpenAI (optional) --------------------------------------------------------
_openai_client = None
if PROVIDER == "openai":
    try:
        from openai import OpenAI
    except Exception as e:
        print(f"OpenAI import failed: {e}", file=sys.stderr)
        sys.exit(1)

    OPENAI_API_KEY = _get_key(
    "OPENAI_API_KEY", "OPENAI_LOCAL_KEY", "OPENAI_API_KEY_LOCAL", "OPENAI_KEY",
    file_var="OPENAI_API_KEY_FILE",
    )
    if not OPENAI_API_KEY:
        print("Missing OPENAI_API_KEY (or *_LOCAL / _FILE).", file=sys.stderr)
        sys.exit(1)

    MODEL_NAME = os.environ.get("MODEL_NAME", "gpt-4o")
    _openai_client = OpenAI(api_key=OPENAI_API_KEY)


# ---- Gemini (optional) --------------------------------------------------------
if PROVIDER == "gemini":
    try:
        import google.generativeai as genai
    except Exception as e:
        print("Missing google-generativeai package. Install with: pip install google-generativeai", file=sys.stderr)
        print(f"Import error: {e}", file=sys.stderr)
        sys.exit(1)

    GEMINI_API_KEY = _get_key(
        "GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_LOCAL_KEY", "GOOGLE_LOCAL_KEY",
        file_var="GEMINI_API_KEY_FILE",
    )
    if not GEMINI_API_KEY:
        print("Missing GEMINI_API_KEY/GOOGLE_API_KEY (or *_LOCAL / _FILE).", file=sys.stderr)
        sys.exit(1)

    GEMINI_MODEL_NAME = os.environ.get("GEMINI_MODEL_NAME", "gemini-1.5-flash")
    genai.configure(api_key=GEMINI_API_KEY)

# ---- Common paths -------------------------------------------------------------
SCRIPT_DIR   = os.path.dirname(__file__)
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)  # one level up from /scripts


# 1) If env vars are set, they win. If not, leave as None for profile resolution.
JP2EN_GLOSSARY_PATH = (os.environ.get("JP2EN_GLOSSARY_PATH", "").strip() or None)
EN2EN_GLOSSARY_PATH = (os.environ.get("EN2EN_GLOSSARY_PATH", "").strip() or None)

# 2) Otherwise resolve via Control Panel profiles in Settings\control.ini
#    (fall back to Settings\config.ini only if control.ini is absent)
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

        if cfg_path:
            prof_jp, prof_en = _read_profiles_from_ini(cfg_path)

        # Prefer Control Panel's exported SETTINGS_DIR if present; otherwise fall back
        env_settings_dir = (os.environ.get("SETTINGS_DIR", "").strip() or None)
        if env_settings_dir:
            settings_dir = env_settings_dir
        else:
            settings_dir = os.path.dirname(cfg_path) if cfg_path else os.path.join(PROJECT_ROOT, "Settings")

        gloss_base = os.path.join(settings_dir, "glossaries")

        cand_jp = os.path.join(gloss_base, prof_jp, "jp2en.txt")
        cand_en = os.path.join(gloss_base, prof_en, "en2en.txt")

        if JP2EN_GLOSSARY_PATH is None:
            JP2EN_GLOSSARY_PATH = cand_jp
        if EN2EN_GLOSSARY_PATH is None:
            EN2EN_GLOSSARY_PATH = cand_en
except Exception:
    # Never break translation if INI/path handling fails
    pass

# ==============================================================================
# Glossaries
# ==============================================================================

def load_glossary(path) -> List[Tuple[str, str]]:
    """Load 'source -> target' pairs.
       - Accept separators: '->', '→', tab, ':', '='
       - Accept encodings: utf-8, utf-8-sig, utf-16, utf-16-le, utf-16-be, cp1252, cp932
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
        return entries

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
            pattern = re.compile(re.escape(src), flags=re.IGNORECASE)
            out = pattern.sub(dst, out)
        else:
            pattern = re.compile(
                rf"\b{re.escape(src)}(?P<suf>s|\'s|’s)?\b",
                flags=re.IGNORECASE
            )
            out = pattern.sub(lambda m: dst + (m.group("suf") or ""), out)
    return out

# --- NEW: mark guessed spans (anything backticked) as italics when enabled ---
def _mark_guessed_pronouns(text: str) -> str:
    import os, re

    # Allow changing the delimiter via env var; default to backtick `
    delim = os.getenv("SHOT_GUESS_DELIM", "`")
    escaped = re.escape(delim)

    # Support single-char delimiters (fast path) and multi-char (fallback).
    if len(delim) == 1:
        # Any content except the delimiter itself; allow spaces, punctuation, apostrophes, etc.
        pattern = re.compile(rf"{escaped}([^{escaped}\r\n]+){escaped}")
    else:
        # Non-greedy match between multi-char delimiters
        pattern = re.compile(rf"{escaped}(.+?){escaped}", flags=re.DOTALL)

    italics_on = (os.getenv("SHOT_ITALICIZE_GUESSED", "1") == "1")

    def repl(m: "re.Match[str]") -> str:
        span = m.group(1)
        return f"⟦i⟧{span}⟦/i⟧" if italics_on else span

    return pattern.sub(repl, text)

# --- NEW: optionally mark first Translation header line 「…」 as a name/title span ---
def _mark_translation_name_line(en_block: str) -> str:
    import os
    if (os.getenv("SHOT_COLOR_SPEAKER", "1") != "1"):
        return en_block
    lines = en_block.splitlines()
    # Find first non-empty line
    idx = next((i for i, ln in enumerate(lines) if ln.strip() != ""), None)
    if idx is None:
        return en_block
    first = lines[idx].strip()
    # Strict: must be enclosed in Japanese corner brackets
    if len(first) >= 2 and first[0] == "「" and first[-1] == "」":
        inner = first[1:-1].strip()  # strip the 「…」 so only the name remains
        # keep any left padding/indentation on that line
        left = lines[idx][:len(lines[idx]) - len(lines[idx].lstrip())]
        lines[idx] = f"{left}⟦name⟧{inner}⟦/name⟧"
        return "\n".join(lines)
    return en_block
    
# --- NEW: optionally mark first Transcript header line 「…」 as a name/title span ---
def _mark_transcript_name_line(jp_block: str) -> str:
    import os
    if (os.getenv("SHOT_COLOR_SPEAKER", "1") != "1"):
        return jp_block
    lines = jp_block.splitlines()
    # Find first non-empty line
    idx = next((i for i, ln in enumerate(lines) if ln.strip() != ""), None)
    if idx is None:
        return jp_block
    first = lines[idx].strip()
    # Strict: must be enclosed in Japanese corner brackets
    if len(first) >= 2 and first[0] == "「" and first[-1] == "」":
        inner = first[1:-1].strip()  # strip the 「…」
        left = lines[idx][:len(lines[idx]) - len(lines[idx].lstrip())]
        lines[idx] = f"{left}⟦name⟧{inner}⟦/name⟧"
        return "\n".join(lines)
    return jp_block

# ===============================================================

# ==============================================================================
# Kana → romaji (fallback name romanization)
# ==============================================================================

# Basic katakana mapping + digraphs + long vowels/small tsu handling
KATAKANA_ROMAJI = {
    "ア":"a","イ":"i","ウ":"u","エ":"e","オ":"o",
    "カ":"ka","キ":"ki","ク":"ku","ケ":"ke","コ":"ko",
    "サ":"sa","シ":"shi","ス":"su","セ":"se","ソ":"so",
    "タ":"ta","チ":"chi","ツ":"tsu","テ":"te","ト":"to",
    "ナ":"na","ニ":"ni","ヌ":"nu","ネ":"ne","ノ":"no",
    "ハ":"ha","ヒ":"hi","フ":"fu","ヘ":"he","ホ":"ho",
    "マ":"ma","ミ":"mi","ム":"mu","メ":"me","モ":"mo",
    "ヤ":"ya","ユ":"yu","ヨ":"yo",
    "ラ":"ra","リ":"ri","ル":"ru","レ":"re","ロ":"ro",
    "ワ":"wa","ヲ":"o","ン":"n",
    "ガ":"ga","ギ":"gi","グ":"gu","ゲ":"ge","ゴ":"go",
    "ザ":"za","ジ":"ji","ズ":"zu","ゼ":"ze","ゾ":"zo",
    "ダ":"da","ヂ":"ji","ヅ":"zu","デ":"de","ド":"do",
    "バ":"ba","ビ":"bi","ブ":"bu","ベ":"be","ボ":"bo",
    "パ":"pa","ピ":"pi","プ":"pu","ペ":"pe","ポ":"po",
    "ヴ":"vu",
    "ァ":"a","ィ":"i","ゥ":"u","ェ":"e","ォ":"o",
    "ャ":"ya","ュ":"yu","ョ":"yo",
    "ー":"-",
}
DIGRAPHS = {
    "キャ":"kya","キュ":"kyu","キョ":"kyo",
    "シャ":"sha","シュ":"shu","ショ":"sho",
    "ジャ":"ja","ジュ":"ju","ジョ":"jo",
    "チャ":"cha","チュ":"chu","チョ":"cho",
    "ニャ":"nya","ニュ":"nyu","ニョ":"nyo",
    "ヒャ":"hya","ヒュ":"hyu","ヒョ":"hyo",
    "ミャ":"mya","ミュ":"myu","ミョ":"myo",
    "リャ":"rya","リュ":"ryu","リョ":"ryo",
    "ギャ":"gya","ギュ":"gyu","ギョ":"gyo",
    "ビャ":"bya","ビュ":"byu","ビョ":"byo",
    "ピャ":"pya","ピュ":"pyu","ピョ":"pyo",
}

def hira_to_kata(s: str) -> str:
    out = []
    for ch in s:
        code = ord(ch)
        if 0x3041 <= code <= 0x3096:
            out.append(chr(code + 0x60))
        elif ch == "ゔ":
            out.append("ヴ")
        else:
            out.append(ch)
    return "".join(out)

def is_all_kana(s: str) -> bool:
    return re.fullmatch(r"[ぁ-ゖァ-ヶー・]+", s) is not None

def katakana_to_romaji(name: str) -> str:
    out = ""
    parts = name.split("・")
    rom_parts = []
    for segment in parts:
        seg = segment
        i = 0
        rom = ""
        while i < len(seg):
            if i+2 <= len(seg) and seg[i:i+2] in DIGRAPHS:
                rom += DIGRAPHS[seg[i:i+2]]; i += 2; continue
            ch = seg[i]
            if ch == "ッ" and i+1 < len(seg):
                nxt = seg[i+1:i+3] if i+3 <= len(seg) and seg[i+1:i+3] in DIGRAPHS else seg[i+1]
                base = DIGRAPHS.get(nxt) or KATAKANA_ROMAJI.get(nxt, "")
                if base: rom += base[0]
                i += 1; continue
            if ch == "ー":
                if rom:
                    for v in "aeiou"[::-1]:
                        if rom.endswith(v): rom += v; break
                i += 1; continue
            rom += KATAKANA_ROMAJI.get(ch, "")
            i += 1
        rom_parts.append(rom)
    romaji = " ".join(p.capitalize() for p in rom_parts if p)
    return romaji if romaji else name

def kana_to_romaji(s: str) -> str:
    return katakana_to_romaji(hira_to_kata(s))

# ==============================================================================
# Prompt & messages
# ==============================================================================

SYSTEM_PROMPT = """You are a translator for retro JRPG screenshots. Return PLAIN TEXT ONLY.
Do NOT use Markdown or code fences. Never output ``` or language tags.

Output EXACTLY two sections with these literal headings:

Transcript:
<Japanese transcription here, preserving all punctuation, line breaks, and full-width brackets 「」 as-is. Only text inside the dialogue/message box(es). Do not add or remove characters. If the speaker name is shown as a distinct header line without brackets, wrap it with Japanese corner brackets: 「name」.>

(blank line)

Translation:
<Fluent English translation of the Transcript block. If a speaker/tag line exists, output it on its own line inside corner brackets, e.g., 「Gus from Casta」. Apply the rule XのY → “Y from X”/“Y of X” only when it is a speaker identifier. Ignore everything outside the box. If no box text exists, output exactly: No Japanese text found.>
"""

def load_system_prompt() -> str:
    """
    Priority:
      1) PROMPT_FILE (path to a UTF-8 text file)
      2) PROMPT_TEXT (full prompt as env var)
      3) PROMPT_PROFILE (name looked up under ./Settings/prompts or ./prompts)
      4) ./prompt.txt next to this script
      5) built-in SYSTEM_PROMPT
    """
    # 1) PROMPT_FILE
    p = os.environ.get("PROMPT_FILE", "").strip()
    if p:
        try:
            if os.path.isfile(p):
                with open(p, "r", encoding="utf-8") as fh:
                    return fh.read()
        except Exception:
            pass

    # 2) PROMPT_TEXT
    prompt_text = os.environ.get("PROMPT_TEXT", "")
    if prompt_text:
        return prompt_text

    # 3) PROMPT_PROFILE (default/furigana/etc.)
    prof = os.environ.get("PROMPT_PROFILE", "").strip()
    if prof:
      for candidate in (
    os.path.join(PROJECT_ROOT, "Settings", "prompts", prof),
    os.path.join(PROJECT_ROOT, "Settings", "prompts", f"{prof}.txt"),
    os.path.join(PROJECT_ROOT, "prompts", prof),
    os.path.join(PROJECT_ROOT, "prompts", f"{prof}.txt"),
):
            try:
                if os.path.isfile(candidate):
                    with open(candidate, "r", encoding="utf-8") as fh:
                        return fh.read()
            except Exception:
                pass

    # 4) local prompt.txt (optional)
    try_path = os.path.join(SCRIPT_DIR, "prompt.txt")
    try:
        if os.path.isfile(try_path):
            with open(try_path, "r", encoding="utf-8") as fh:
                return fh.read()
    except Exception:
        pass

    # 5) fallback
    return SYSTEM_PROMPT

def build_jp2en_prompt(jp2en: List[Tuple[str, str]]) -> str:
    if not jp2en:
        return ""
    lines = ["When translating, apply these exact JP→EN mappings wherever they appear:"]
    for jp, en in jp2en:
        lines.append(f"- {jp} → {en}")
    return "\n".join(lines)

def file_to_data_url(path: str) -> str:
    mime, _ = mimetypes.guess_type(path)
    if not mime:
        mime = "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{b64}"

def make_messages_for_openai(image_paths: List[str], jp2en: List[Tuple[str, str]]) -> List[dict]:
    glossary_block = build_jp2en_prompt(jp2en)
    sys_prompt = load_system_prompt() + ("\n\n" + glossary_block if glossary_block else "")

    content = []
    for p in image_paths:
        abs_p = os.path.abspath(p)
        if not os.path.isfile(abs_p):
            continue
        try:
            content.append({
                "type": "image_url",
                "image_url": {"url": file_to_data_url(abs_p)}
            })
        except Exception as e:
            content.append({"type": "text", "text": f"(Failed to read image: {abs_p} – {e})"})
    if not content:
        content.append({"type": "text", "text": "No image provided."})

    return [
        {"role": "system", "content": sys_prompt},
        {"role": "user", "content": content},
    ]

def gemini_image_parts(paths: List[str]) -> List[dict]:
    """Gemini accepts parts like {'mime_type': 'image/png', 'data': b'...'}."""
    parts = []
    for p in paths:
        abs_p = os.path.abspath(p)
        if not os.path.isfile(abs_p):
            continue
        mime, _ = mimetypes.guess_type(abs_p)
        if not mime:
            mime = "image/png"
        with open(abs_p, "rb") as f:
            parts.append({"mime_type": mime, "data": f.read()})
    return parts

# ==============================================================================
# Output post-processing (sanitize / enforce / name header normalization)
# ==============================================================================

OPEN_BRACKETS  = "「『〈《［（[<"
CLOSE_BRACKETS = "」』〉》］）]>"
OPEN_CLASS     = "[" + re.escape(OPEN_BRACKETS)  + "]"
CLOSE_CLASS    = "[" + re.escape(CLOSE_BRACKETS) + "]"

NAME_LINE_RE     = re.compile(rf"^\s*{OPEN_CLASS}\s*(.+?)\s*{CLOSE_CLASS}\s*$")
INLINE_HEADER_RE = re.compile(rf"^\s*{OPEN_CLASS}\s*(.+?)\s*{CLOSE_CLASS}\s*[:：]?\s*(.+)$")

def strip_code_fences(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"^\s*(?:```|''')[^\r\n]*[\r\n]+", "", s)
    s = re.sub(r"[\r\n]*(?:```|''')\s*$", "", s)
    return s

def enforce_transcript_translation(s: str) -> str:
    if not s:
        return "Transcript:\n\nTranslation:\n"
    s = s.strip()
    # normalize legacy tags
    s = re.sub(r'^\s*<\s*block1(?:-jp)?\s*>\s*', 'Transcript:\n', s, flags=re.I|re.M)
    s = re.sub(r'^\s*<\s*block2(?:-en)?\s*>\s*', '\n\nTranslation:\n', s, flags=re.I|re.M)

    if re.search(r'(?mi)^Transcript:\s*', s) and re.search(r'(?mi)^Translation:\s*', s):
        s = re.sub(r'(?is)(^Transcript:\s*.*?)(?:\n{1,3})?^Translation:\s*',
                   lambda m: m.group(1).rstrip() + "\n\nTranslation:\n", s, count=1, flags=re.M)
        return s.strip()

    parts = re.split(r'\n\s*\n', s, maxsplit=1)
    if len(parts) == 2:
        jp, en = parts
    else:
        jp, en = s, ""
    return f"Transcript:\n{jp.strip()}\n\nTranslation:\n{en.strip()}"

def split_tt(s: str) -> Tuple[str, str]:
    s = s.replace("\r\n", "\n")
    m = re.search(r'(?is)^Transcript:\s*(.*?)\n\nTranslation:\s*(.*)\s*$', s, flags=re.M)
    if not m:
        s = enforce_transcript_translation(s)
        m = re.search(r'(?is)^Transcript:\s*(.*?)\n\nTranslation:\s*(.*)\s*$', s, flags=re.M)
    jp = (m.group(1) if m else "").strip()
    en = (m.group(2) if m else "").strip()
    return jp, en

def looks_like_jp_name(line: str) -> bool:
    ln = line.strip()
    if len(ln) == 0 or len(ln) > 12:
        return False
    return re.fullmatch(r"[ぁ-ゖァ-ヶ一-龯々〆ヶー・\s]+", ln) is not None

def normalize_jp_speaker_line(jp_block: str) -> Tuple[str, str]:
    if not jp_block:
        return jp_block, ""
    lines = [ln.rstrip() for ln in jp_block.replace("\r\n", "\n").split("\n")]
    idx = next((i for i, ln in enumerate(lines) if ln.strip() != ""), None)
    if idx is None:
        return jp_block, ""

    first = lines[idx].strip()
    m = NAME_LINE_RE.match(first)
    if m:
        name = m.group(1).strip()
        lines[idx] = f"「{name}」"
        return "\n".join(lines).strip(), name

    if looks_like_jp_name(first):
        lines[idx] = f"「{first}」"
        return "\n".join(lines).strip(), first

    return jp_block.strip(), ""

def translate_jp_name_to_en(name_jp: str, jp2en: List[Tuple[str, str]]) -> str:
    nm = name_jp.strip()
    for jp, en in jp2en:
        if nm == jp:
            return en
    if is_all_kana(nm):
        return kana_to_romaji(nm)
    return nm

def normalize_translation_name_line(en_block: str, jp2en: List[Tuple[str, str]], jp_name_hint: str = "") -> str:
    if not en_block:
        return en_block
    lines = [ln.rstrip() for ln in en_block.replace("\r\n", "\n").split("\n")]
    idx = next((i for i, ln in enumerate(lines) if ln.strip() != ""), None)
    if idx is None:
        return en_block

    first = lines[idx].strip()
    m = NAME_LINE_RE.match(first)
    if m:
        name = m.group(1).strip()
        name_en = translate_jp_name_to_en(name, jp2en)
        lines[idx] = f"「{name_en}」"
        return "\n".join(lines).strip()

    if jp_name_hint:
        name_en = translate_jp_name_to_en(jp_name_hint, jp2en)
        lines.insert(idx, f"「{name_en}」")
        return "\n".join(lines).strip()

    return en_block.strip()

# ==============================================================================
# Calls to providers
# ==============================================================================

def call_openai(image_paths: List[str], jp2en: List[Tuple[str, str]]) -> str:
    messages = make_messages_for_openai(image_paths, jp2en)
    resp = _openai_client.chat.completions.create(
        model=os.environ.get("MODEL_NAME", "gpt-4o"),
        messages=messages,
        temperature=0.2,
    )
    return resp.choices[0].message.content or ""

def call_gemini(image_paths: List[str], jp2en: List[Tuple[str, str]]) -> str:
    import google.generativeai as genai

    # Try to relax safety filters so adult game CGs don't get blocked as easily.
    # If the types module isn't available, we just fall back to the default settings.
    safety_settings = None
    try:
        from google.generativeai.types.safety_types import HarmBlockThreshold, HarmCategory

        # NOTE:
        # - BLOCK_NONE is the least restrictive, but on some accounts it is a
        #   "restricted" level and may require allowlisting / special billing.
        # - If you get an error mentioning a "restricted HarmBlockThreshold",
        #   change BLOCK_NONE to BLOCK_ONLY_HIGH here.
        safety_settings = {
            HarmCategory.HARM_CATEGORY_HARASSMENT:         HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH:        HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT:  HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT:  HarmBlockThreshold.BLOCK_NONE,
        }
    except Exception:
        safety_settings = None

    glossary_block = build_jp2en_prompt(jp2en)
    sys_prompt = load_system_prompt() + ("\n\n" + glossary_block if glossary_block else "")

    # Create model with system instruction so we can pass only a tiny user prompt
    model_name = os.environ.get("GEMINI_MODEL_NAME", "gemini-1.5-flash")

    # --- FIX: Add the required 'models/' prefix if it's missing ---
    if not model_name.startswith("models/"):
        model_name = f"models/{model_name}"
    # --- END FIX ---

    model = genai.GenerativeModel(model_name, system_instruction=sys_prompt)

    parts = gemini_image_parts(image_paths)
    if not parts:
        parts = [{"text": "No image provided."}]

    content = [{"text": "Process the attached image(s) and answer exactly as instructed."}] + parts

    if safety_settings is not None:
        resp = model.generate_content(
            content,
            generation_config={"temperature": 0.2},
            safety_settings=safety_settings,
        )
    else:
        resp = model.generate_content(
            content,
            generation_config={"temperature": 0.2},
        )

    # Gemini returns no candidates if the prompt was blocked. In that case
    # resp.text raises the "response.parts quick accessor" error you saw.
    try:
        return (resp.text or "")
    except Exception:
        block_reason = None
        try:
            fb = getattr(resp, "prompt_feedback", None)
            block_reason = getattr(fb, "block_reason", None) if fb else None
        except Exception:
            pass
        if block_reason:
            return f"(Gemini blocked the image; block_reason={block_reason})"
        return "(Gemini returned no text candidates – likely blocked by safety settings.)"

# ==============================================================================
# Main translation
# ==============================================================================

def translate_images(paths: List[str],
                     jp2en: List[Tuple[str, str]],
                     en2en: List[Tuple[str, str]]) -> str:
    try:
        if PROVIDER == "gemini":
            raw = call_gemini(paths, jp2en)
        else:
            raw = call_openai(paths, jp2en)
    except Exception as e:
        provider_name = "Gemini" if PROVIDER == "gemini" else "OpenAI"
        return f"(Python error) {provider_name} call failed: {e}"

    # Sanitize model text
    out = strip_code_fences(raw)

        # --- Sidecar: write latest JP transcript for Learning flow (normalized like overlay) ---
    try:
        enforced = enforce_transcript_translation(out)
        _jp_block, _en_block = split_tt(enforced)

        # Normalize the JP block so speaker headers are wrapped in full-width brackets,
        # exactly like the Translator overlay shows.
        _jp_block_norm, _ = normalize_jp_speaker_line(_jp_block)

        if _jp_block_norm.strip():
            atomic_write_text(LAST_JP, _jp_block_norm.strip())
            # Optional: remember the image paths that produced this JP
            if paths:
                atomic_write_text(LAST_SRC, "\r\n".join(os.path.abspath(p) for p in paths))
    except Exception:
        # Don’t break translation if sidecar write fails
        pass
    # --- end sidecar write ---

    # Raw output (no headings / normalization)
    if POSTPROC_MODE == "none":
        return out.replace("\r\n", "\n").strip()

    # Translation-only (extract Translation: if present, otherwise pass-through)
    if POSTPROC_MODE in ("translation", "en-only", "en"):
        enforced = enforce_transcript_translation(out)   # makes a best-effort TT split if needed
        _jp_block, en_block = split_tt(enforced)
        if en2en:
            en_block = apply_en_glossary(en_block, en2en)
        # NEW: add inline markers for overlay styling
        en_block = _mark_translation_name_line(en_block)
        en_block = _mark_guessed_pronouns(en_block)
        return en_block.replace("\r\n", "\n").strip()

    # Default: full TT pipeline with name normalization
    out = enforce_transcript_translation(out)
    jp_block, en_block = split_tt(out)

    jp_block, jp_name = normalize_jp_speaker_line(jp_block)  # bracket-header in JP
    jp_block = _mark_transcript_name_line(jp_block)          # ← NEW: wrap & strip brackets if color ON
    en_block = normalize_translation_name_line(en_block, jp2en, jp_name_hint=jp_name)

    if en2en:
        en_block = apply_en_glossary(en_block, en2en)

    # NEW: add inline markers for overlay styling
    en_block = _mark_translation_name_line(en_block)
    en_block = _mark_guessed_pronouns(en_block)

    final = f"Transcript:\n{jp_block.strip()}\n\nTranslation:\n{en_block.strip()}"
    return final.replace("\r\n", "\n").strip()

# ==============================================================================
# CLI
# ==============================================================================

def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python screenshot_translator.py <image1> [<image2> ...]", file=sys.stderr)
        sys.exit(2)

    images = sys.argv[1:]  # in chronological order
    jp2en = load_glossary(JP2EN_GLOSSARY_PATH)
    en2en = load_glossary(EN2EN_GLOSSARY_PATH)

    result = translate_images(images, jp2en, en2en)
    atomic_write_text(OCR_TXT, result)

if __name__ == "__main__":
    main()