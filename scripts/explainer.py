#!/usr/bin/env python
# -*- coding: utf-8 -*-

import configparser
import base64
import hashlib
import io
import mimetypes
import os
import re
import shutil
import sqlite3
import sys
import tempfile
from datetime import datetime
from typing import List, Tuple

# The packaged Python runtime is isolated by python312._pth and does not include
# the launched script's directory automatically. Add it explicitly before
# importing the sibling Study Library helper.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from dotenv import load_dotenv

from study_library_sections import (
    SECTION_SCHEMA,
    ensure_section_schema,
    replace_explanation_sections,
)

# Console UTF-8 (Windows safe)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# Shared temp folder (same as audio + overlay)
TEMP_DIR      = os.environ.get("TEMP") or tempfile.gettempdir()
OVERLAY_DIR   = os.path.join(TEMP_DIR, "JRPG_Overlay")
LAST_JP       = os.path.join(OVERLAY_DIR, "last_jp.txt")
LAST_SRC      = os.path.join(OVERLAY_DIR, "last_src.txt")
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


def read_last_source_paths() -> List[str]:
    paths = []
    for line in read_text(LAST_SRC).splitlines():
        path = line.strip()
        if path and os.path.isfile(path):
            paths.append(path)
    return paths


EXPLANATION_ARCHIVE_RE = re.compile(
    r"^(?P<timestamp>\d{8}-\d{6})_explain(?:_v(?P<version>\d{2,}))?\.txt$",
    re.IGNORECASE,
)
EXPLANATION_SOURCE_HEADER = "=== Japanese (source) ==="
EXPLANATION_BODY_HEADER = "=== Explanation ==="
SCREENSHOT_SOURCE_NOTE = "[Source supplied as cached screenshot(s)]"


def explanation_source_block(jp: str, source_paths: List[str]) -> str:
    """Return the source section written into an archived explanation."""
    if jp:
        return jp.strip()
    lines = [SCREENSHOT_SOURCE_NOTE]
    lines.extend(source_paths)
    return "\n".join(lines)


def normalize_explanation_source(source_block: str) -> str:
    """Ignore line wrapping and spacing when comparing repeated Japanese text."""
    source_block = source_block.replace("\ufeff", "").strip()
    if source_block.startswith(SCREENSHOT_SOURCE_NOTE):
        paths = [
            os.path.normcase(os.path.abspath(line.strip()))
            for line in source_block.splitlines()[1:]
            if line.strip()
        ]
        return "screenshots:" + "\n".join(paths)
    return "jp:" + re.sub(r"\s+", "", source_block)


def read_archived_explanation_source(path: str) -> str:
    """Read only the saved source section, returning an empty string on failure."""
    archived = read_text(path)
    start = archived.find(EXPLANATION_SOURCE_HEADER)
    end = archived.find(EXPLANATION_BODY_HEADER)
    if start < 0 or end < 0 or end <= start:
        return ""
    start += len(EXPLANATION_SOURCE_HEADER)
    return archived[start:end].strip()


def next_explanation_archive_path(
    explains_dir: str, jp: str, source_paths: List[str]
) -> str:
    """Keep consecutive repeats together as _v02, _v03, etc."""
    from datetime import datetime, timedelta

    newest = None
    newest_key = None
    for name in os.listdir(explains_dir):
        match = EXPLANATION_ARCHIVE_RE.match(name)
        if not match:
            continue
        path = os.path.join(explains_dir, name)
        try:
            key = (os.path.getmtime(path), name.lower())
        except OSError:
            continue
        if newest_key is None or key > newest_key:
            newest = (path, match)
            newest_key = key

    current_source = normalize_explanation_source(
        explanation_source_block(jp, source_paths)
    )
    if newest is not None:
        newest_path, newest_match = newest
        previous_source = normalize_explanation_source(
            read_archived_explanation_source(newest_path)
        )
        if previous_source and previous_source == current_source:
            base = newest_match.group("timestamp")
            version = int(newest_match.group("version") or "1") + 1
            while True:
                candidate = os.path.join(
                    explains_dir, f"{base}_explain_v{version:02d}.txt"
                )
                if not os.path.exists(candidate):
                    return candidate
                version += 1

    # Preserve the unversioned first-attempt format. If two different sources are
    # archived within the same second, advance the filename timestamp rather than
    # overwriting the earlier explanation or incorrectly marking it as a repeat.
    timestamp = datetime.now()
    while True:
        base = timestamp.strftime("%Y%m%d-%H%M%S")
        candidate = os.path.join(explains_dir, f"{base}_explain.txt")
        if not os.path.exists(candidate):
            return candidate
        timestamp += timedelta(seconds=1)


def env_flag(name: str, default: bool = False) -> bool:
    """Read a conventional boolean environment variable."""
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def archive_plain_text(
    explains_dir: str, jp: str, source_paths: List[str], text: str
) -> str:
    """Write the legacy text archive and return its path, or an empty string."""
    try:
        os.makedirs(explains_dir, exist_ok=True)
        out_path = next_explanation_archive_path(explains_dir, jp, source_paths)
        with open(out_path, "w", encoding="utf-8", newline="\r\n") as archive_file:
            archive_file.write(EXPLANATION_SOURCE_HEADER + "\r\n")
            archive_file.write(explanation_source_block(jp, source_paths) + "\r\n\r\n")
            archive_file.write(EXPLANATION_BODY_HEADER + "\r\n")
            archive_file.write(text + "\r\n")
        print(f"(Archived) {out_path}")
        return out_path
    except Exception as exc:
        print(f"(Archive skipped) Could not write explanation: {exc}", file=sys.stderr)
        return ""


STUDY_LIBRARY_SCHEMA = """
CREATE TABLE IF NOT EXISTS metadata (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS explanation_groups (
    id INTEGER PRIMARY KEY,
    game_profile TEXT NOT NULL DEFAULT '',
    source_hash TEXT NOT NULL,
    source_kind TEXT NOT NULL DEFAULT 'text',
    source_japanese TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(game_profile, source_hash)
);

CREATE TABLE IF NOT EXISTS explanations (
    id INTEGER PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES explanation_groups(id) ON DELETE CASCADE,
    version INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    provider TEXT NOT NULL DEFAULT '',
    model TEXT NOT NULL DEFAULT '',
    prompt_profile TEXT NOT NULL DEFAULT '',
    raw_text TEXT NOT NULL,
    manual_original_text TEXT NOT NULL DEFAULT '',
    manually_edited_at TEXT NOT NULL DEFAULT '',
    text_archive_path TEXT NOT NULL DEFAULT '',
    preferred INTEGER NOT NULL DEFAULT 0,
    UNIQUE(group_id, version)
);

CREATE TABLE IF NOT EXISTS media (
    id INTEGER PRIMARY KEY,
    explanation_id INTEGER NOT NULL REFERENCES explanations(id) ON DELETE CASCADE,
    sort_order INTEGER NOT NULL,
    relative_path TEXT NOT NULL,
    original_name TEXT NOT NULL DEFAULT '',
    mime_type TEXT NOT NULL DEFAULT '',
    sha256 TEXT NOT NULL DEFAULT '',
    UNIQUE(explanation_id, sort_order)
);

CREATE INDEX IF NOT EXISTS idx_groups_profile_updated
    ON explanation_groups(game_profile, updated_at DESC);
CREATE INDEX IF NOT EXISTS idx_explanations_group_version
    ON explanations(group_id, version DESC);
CREATE INDEX IF NOT EXISTS idx_media_explanation_order
    ON media(explanation_id, sort_order);
"""

STUDY_LIBRARY_SCHEMA += SECTION_SCHEMA


def file_sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source_file:
        for chunk in iter(lambda: source_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def study_source_identity(jp: str, source_paths: List[str]) -> Tuple[str, str]:
    """Return a stable source kind and hash without relying on temporary paths."""
    normalized_jp = re.sub(r"\s+", "", jp.replace("\ufeff", "").strip())
    if normalized_jp:
        identity = "jp:" + normalized_jp
        source_kind = "text"
    else:
        image_hashes = []
        for source_path in source_paths:
            try:
                image_hashes.append(file_sha256(source_path))
            except OSError:
                continue
        identity = "images:" + "|".join(image_hashes)
        source_kind = "images"
    return source_kind, hashlib.sha256(identity.encode("utf-8")).hexdigest()


def initialize_study_library(connection: sqlite3.Connection) -> None:
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 10000")
    connection.executescript(STUDY_LIBRARY_SCHEMA)
    ensure_section_schema(connection)
    connection.commit()


def archive_study_library_entry(
    study_dir: str,
    jp: str,
    source_paths: List[str],
    text: str,
    game_profile: str,
    provider: str,
    model: str,
    prompt_profile: str,
    include_screenshots: bool,
    text_archive_path: str = "",
) -> None:
    """Store one explanation and optional source images in the Stage 1 library."""
    os.makedirs(study_dir, exist_ok=True)
    media_dir = os.path.join(study_dir, "Media")
    if include_screenshots:
        os.makedirs(media_dir, exist_ok=True)

    database_path = os.path.join(study_dir, "study_library.db")
    source_kind, source_hash = study_source_identity(jp, source_paths)
    created_at = datetime.now().astimezone().isoformat(timespec="seconds")
    stored_text_archive_path = ""
    if text_archive_path:
        try:
            stored_text_archive_path = os.path.relpath(
                text_archive_path, study_dir
            ).replace(os.sep, "/")
        except (OSError, ValueError):
            stored_text_archive_path = text_archive_path
    connection = sqlite3.connect(database_path, timeout=10)
    copied_media = []
    try:
        initialize_study_library(connection)
        connection.execute("BEGIN IMMEDIATE")
        row = connection.execute(
            "SELECT id FROM explanation_groups WHERE game_profile = ? AND source_hash = ?",
            (game_profile, source_hash),
        ).fetchone()
        if row:
            group_id = int(row[0])
            connection.execute(
                "UPDATE explanation_groups SET updated_at = ?, source_japanese = ? "
                "WHERE id = ?",
                (created_at, jp.strip(), group_id),
            )
        else:
            cursor = connection.execute(
                "INSERT INTO explanation_groups("
                "game_profile, source_hash, source_kind, source_japanese, created_at, updated_at"
                ") VALUES(?, ?, ?, ?, ?, ?)",
                (
                    game_profile,
                    source_hash,
                    source_kind,
                    jp.strip(),
                    created_at,
                    created_at,
                ),
            )
            group_id = int(cursor.lastrowid)

        version = int(
            connection.execute(
                "SELECT COALESCE(MAX(version), 0) + 1 FROM explanations WHERE group_id = ?",
                (group_id,),
            ).fetchone()[0]
        )
        connection.execute(
            "UPDATE explanations SET preferred = 0 WHERE group_id = ?", (group_id,)
        )
        cursor = connection.execute(
            "INSERT INTO explanations("
            "group_id, version, created_at, provider, model, prompt_profile, raw_text, "
            "text_archive_path, preferred"
            ") VALUES(?, ?, ?, ?, ?, ?, ?, ?, 1)",
            (
                group_id,
                version,
                created_at,
                provider,
                model,
                prompt_profile,
                text,
                stored_text_archive_path,
            ),
        )
        explanation_id = int(cursor.lastrowid)
        replace_explanation_sections(connection, explanation_id, text)

        if include_screenshots:
            for sort_order, source_path in enumerate(source_paths, start=1):
                if not os.path.isfile(source_path):
                    continue
                extension = os.path.splitext(source_path)[1].lower()
                if not re.fullmatch(r"\.[a-z0-9]{1,8}", extension):
                    extension = ".png"
                filename = f"{explanation_id:08d}_{sort_order:02d}{extension}"
                destination = os.path.join(media_dir, filename)
                temporary = destination + ".tmp"
                try:
                    shutil.copy2(source_path, temporary)
                    os.replace(temporary, destination)
                finally:
                    try:
                        os.remove(temporary)
                    except OSError:
                        pass
                copied_media.append(destination)
                connection.execute(
                    "INSERT INTO media("
                    "explanation_id, sort_order, relative_path, original_name, mime_type, sha256"
                    ") VALUES(?, ?, ?, ?, ?, ?)",
                    (
                        explanation_id,
                        sort_order,
                        os.path.relpath(destination, study_dir).replace(os.sep, "/"),
                        os.path.basename(source_path),
                        mimetypes.guess_type(source_path)[0] or "application/octet-stream",
                        file_sha256(destination),
                    ),
                )

        connection.commit()
        profile_label = game_profile or "Unsorted"
        print(
            f"(Study Library) entry {explanation_id}, {profile_label}, version {version}"
        )
    except Exception:
        connection.rollback()
        for copied_path in copied_media:
            try:
                os.remove(copied_path)
            except OSError:
                pass
        raise
    finally:
        connection.close()


def file_to_data_url(path: str) -> str:
    mime_type = mimetypes.guess_type(path)[0] or "image/png"
    with open(path, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


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


def resolve_glossary_settings(project_root: str) -> Tuple[str, str, bool]:
    """Resolve the profiles selected in the control panel's Terminology tab."""
    enabled = (
        os.environ.get("USE_TERMINOLOGY_OVERRIDES", "1").strip().lower()
        not in {"0", "false", "no", "off"}
    )
    source_path = (
        os.environ.get("JP2TL_GLOSSARY_PATH", "").strip()
        or os.environ.get("JP2EN_GLOSSARY_PATH", "").strip()
    )
    target_path = (
        os.environ.get("TL2TL_GLOSSARY_PATH", "").strip()
        or os.environ.get("EN2EN_GLOSSARY_PATH", "").strip()
    )
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
                enabled = config.getboolean(
                    "cfg", "useTerminologyOverrides", fallback=True
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
    return source_path, target_path, enabled

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

PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
(
    JP2TL_GLOSSARY_PATH,
    TL2TL_GLOSSARY_PATH,
    USE_TERMINOLOGY_OVERRIDES,
) = resolve_glossary_settings(PROJECT_ROOT)
JP2TL_GLOSSARY = (
    load_glossary(JP2TL_GLOSSARY_PATH) if USE_TERMINOLOGY_OVERRIDES else []
)
TL2TL_GLOSSARY = (
    load_glossary(TL2TL_GLOSSARY_PATH) if USE_TERMINOLOGY_OVERRIDES else []
)

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
source_paths = read_last_source_paths()
if not jp and not source_paths:
    print("(No Japanese transcript or cached source screenshots found)", file=sys.stderr)
    sys.exit(2)

if jp:
    prompt = prompt_tpl.format(jp=jp)
else:
    prompt = prompt_tpl.format(
        jp="[Read the original Japanese from the attached screenshot(s), in order.]"
    )
    prompt += (
        "\n\nThe original Japanese source is supplied in the attached screenshot(s). "
        "Read only the Japanese text inside the dialogue, message, or menu box. "
        "When multiple images are attached, treat them as consecutive parts of one passage "
        "in the order supplied and join sentence fragments across image boundaries."
    )
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
        if source_paths:
            content_parts = [types.Part.from_text(text=prompt)]
            for source_path in source_paths:
                mime_type = mimetypes.guess_type(source_path)[0] or "image/png"
                with open(source_path, "rb") as image_file:
                    content_parts.append(
                        types.Part.from_bytes(data=image_file.read(), mime_type=mime_type)
                    )
            contents = [types.Content(role="user", parts=content_parts)]
        else:
            contents = prompt

        resp = client.models.generate_content(
            model=model_name,
            contents=contents,
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
        if source_paths:
            input_content = [{"type": "input_text", "text": prompt}]
            input_content.extend(
                {"type": "input_image", "image_url": file_to_data_url(source_path)}
                for source_path in source_paths
            )
            request_input = [{"role": "user", "content": input_content}]
        else:
            request_input = prompt
        r = client.responses.create(model=MODEL_NAME, input=request_input)
        text = (getattr(r, "output_text", "") or "").strip()

    else:
        raise RuntimeError(f"Unknown provider: {PROVIDER}")

except Exception as e:
    error_text = str(e)
    if "Missing OPENAI_API_KEY" in error_text:
        atomic_write_text(
            EXPLAINER_TXT,
            "OpenAI API key missing.\n\n"
            "Add it in the API Keys tab, or set OPENAI_API_KEY in Windows "
            "Environment Variables and restart JRPG Translator.",
        )
    elif "Missing GEMINI_API_KEY/GOOGLE_API_KEY" in error_text:
        atomic_write_text(
            EXPLAINER_TXT,
            "Gemini API key missing.\n\n"
            "Add it in the API Keys tab, or set GEMINI_API_KEY (or GOOGLE_API_KEY) "
            "in Windows Environment Variables and restart JRPG Translator.",
        )
    print(f"(Python error) Explain call failed: {e}", file=sys.stderr)
    sys.exit(1)

if not text:
    text = "(No explanation returned)"

if TL2TL_GLOSSARY:
    text = apply_target_glossary(text, TL2TL_GLOSSARY, protected_text=jp)

# Always update the live explainer.txt (overlay reads this)
atomic_write_text(EXPLAINER_TXT, text)
print(f"Wrote explanation to: {EXPLAINER_TXT}")

# Optional legacy plain-text archive. Keep this independent from the Study Library
# so either output can be enabled without making the other one a dependency.
settings_dir = os.environ.get("SETTINGS_DIR", "").strip()
if not settings_dir:
    settings_dir = os.path.join(os.getcwd(), "Settings")

text_archive_path = ""
if env_flag("SAVE_EXPLAINS"):
    explains_dir = os.environ.get("EXPLAIN_SAVE_DIR", "").strip()
    if not explains_dir:
        explains_dir = os.path.join(settings_dir, "Explanations")
    text_archive_path = archive_plain_text(explains_dir, jp, source_paths, text)

if env_flag("SAVE_STUDY_LIBRARY"):
    study_dir = os.environ.get("STUDY_LIBRARY_DIR", "").strip()
    if not study_dir:
        study_dir = os.path.join(settings_dir, "Study Library")
    try:
        archive_study_library_entry(
            study_dir=study_dir,
            jp=jp,
            source_paths=source_paths,
            text=text,
            game_profile=os.environ.get("STUDY_LIBRARY_PROFILE", "").strip(),
            provider=PROVIDER,
            model=(GEM_MODEL if PROVIDER == "gemini" else MODEL_NAME),
            prompt_profile=os.environ.get("EXPLAIN_PROMPT_PROFILE", "").strip(),
            include_screenshots=env_flag("STUDY_LIBRARY_SCREENSHOTS", True),
            text_archive_path=text_archive_path,
        )
    except Exception as exc:
        # Archiving must never turn a successful explanation request into a failed one.
        print(f"(Study Library skipped) {exc}", file=sys.stderr)
