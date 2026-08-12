#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build the local Study Library candidate lists used by Review for Anki."""

from __future__ import annotations

import argparse
import configparser
import hashlib
import json
import os
import re
import sqlite3
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from anki_bridge import (
    AnkiUnavailable,
    anki_endpoint_available,
    anki_process_running,
    normalize_japanese,
    notes_for_scope,
    plain_text,
    remove_readings,
)
from example_sentence import call_model, extract_json_object, load_project_environment


REVIEW_METADATA_KEY = "anki_candidates_reviewed_through"
HIDDEN_VOCABULARY_TABLE = "anki_candidate_hidden_vocabulary"
RECOMMENDATION_TABLE = "anki_candidate_recommendations"
RECOMMENDATION_PROMPT_VERSION = "3"
HIDDEN_MIGRATION_METADATA_KEY = "anki_candidate_hidden_migrated_global"
ENTRY_DASH_RE = re.compile(r"\s*(?:—|–|--|\s-\s)\s*")
BULLET_RE = re.compile(r"^\s*(?:[*•・]|[-–—]|\d+[.)])\s+")

PROMPT_LANGUAGE_NAMES = {
    "de": "German",
    "en": "English",
    "es": "Spanish",
    "fr": "French",
    "it": "Italian",
    "ja": "Japanese",
    "ko": "Korean",
    "nl": "Dutch",
    "pl": "Polish",
    "pt": "Portuguese",
    "ru": "Russian",
    "uk": "Ukrainian",
    "zh-cn": "Simplified Chinese",
    "zh-tw": "Traditional Chinese",
}

EXPLANATION_HEADING_LANGUAGES = {
    "natural english translation": "English",
    "natural english paraphrase": "English",
    "original japanese": "English",
    "detailed analysis": "English",
    "natürliche deutsche übersetzung": "German",
    "japanischer originaltext": "German",
    "detaillierte analyse": "German",
    "traducción natural al español": "Spanish",
    "español original": "Spanish",
    "análisis detallado": "Spanish",
    "traduction française naturelle": "French",
    "français original": "French",
    "analyse détaillée": "French",
    "traduzione naturale in italiano": "Italian",
    "italiano originale": "Italian",
    "analisi dettagliata": "Italian",
    "自然な日本語での言い換え": "Japanese",
    "日本語原文": "Japanese",
    "詳しい分析": "Japanese",
    "자연스러운 한국어 번역": "Korean",
    "일본어 원문": "Korean",
    "상세 분석": "Korean",
    "natuurlijke nederlandse vertaling": "Dutch",
    "origineel japans": "Dutch",
    "gedetailleerde analyse": "Dutch",
    "naturalne tłumaczenie na polski": "Polish",
    "oryginalny tekst japoński": "Polish",
    "szczegółowa analiza": "Polish",
    "tradução natural para português": "Portuguese",
    "japonês original": "Portuguese",
    "análise detalhada": "Portuguese",
    "естественный перевод на русский": "Russian",
    "оригинальный японский текст": "Russian",
    "подробный разбор": "Russian",
    "природний переклад українською": "Ukrainian",
    "оригінальний японський текст": "Ukrainian",
    "докладний розбір": "Ukrainian",
    "自然简体中文翻译": "Simplified Chinese",
    "详细分析": "Simplified Chinese",
    "自然繁體中文翻譯": "Traditional Chinese",
    "詳細分析": "Traditional Chinese",
}


def encode_field(value: object) -> str:
    if value is None:
        return ""
    return str(value).encode("utf-8").hex()


def decode_field(value: str) -> str:
    if not value:
        return ""
    return bytes.fromhex(value).decode("utf-8")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(text, encoding="utf-8", newline="\n")
    os.replace(temporary, path)


def write_rows(path: Path, rows: Iterable[Iterable[object]]) -> None:
    lines = ["\t".join(str(value) for value in row) for row in rows]
    write_text(path, "\n".join(lines) + ("\n" if lines else ""))


def connect_read_only(database: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(
        f"file:{database.resolve().as_posix()}?mode=ro", uri=True, timeout=10
    )
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA query_only = ON")
    connection.execute("PRAGMA busy_timeout = 10000")
    return connection


def connect_write(database: Path) -> sqlite3.Connection:
    connection = sqlite3.connect(database, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 10000")
    return connection


def ensure_triage_schema(preferences_database: Path) -> None:
    """Create the application-wide, reversible candidate-triage storage."""
    preferences_database.parent.mkdir(parents=True, exist_ok=True)
    connection = connect_write(preferences_database)
    try:
        connection.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {HIDDEN_VOCABULARY_TABLE} (
                normalized_front TEXT PRIMARY KEY,
                display_front TEXT NOT NULL DEFAULT '',
                hidden_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {RECOMMENDATION_TABLE} (
                candidate_hash TEXT PRIMARY KEY,
                candidate_kind TEXT NOT NULL,
                recommended INTEGER NOT NULL,
                score INTEGER NOT NULL DEFAULT 0,
                reason TEXT NOT NULL DEFAULT '',
                provider TEXT NOT NULL DEFAULT '',
                model TEXT NOT NULL DEFAULT '',
                generated_at TEXT NOT NULL
            )
            """
        )
        connection.commit()
    finally:
        connection.close()


def hidden_vocabulary(
    preferences_database: Path,
) -> dict[str, dict[str, str]]:
    ensure_triage_schema(preferences_database)
    connection = connect_read_only(preferences_database)
    try:
        rows = connection.execute(
            f"SELECT normalized_front, display_front, hidden_at "
            f"FROM {HIDDEN_VOCABULARY_TABLE}"
        ).fetchall()
        return {
            str(row[0]): {
                "display": str(row[1] or row[0]),
                "hidden_at": str(row[2] or ""),
            }
            for row in rows
        }
    finally:
        connection.close()


def candidate_hash(
    kind: str,
    source: str,
    context: str = "",
    recommendation_profile: str = "",
) -> str:
    normalized_source = normalize_japanese(
        source, remove_attached_readings=True
    )
    payload = "\0".join(
        (
            kind.strip().lower(),
            normalized_source,
            unicodedata.normalize("NFKC", context or "").strip(),
            recommendation_profile,
        )
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def recommendation_profile_signature(
    provider: str,
    model: str,
    learner_level: str,
    selection_style: str,
    focus_areas: str,
    additional_criteria: str,
) -> str:
    """Identify the exact settings under which a rating was generated."""
    payload = json.dumps(
        {
            "prompt_version": RECOMMENDATION_PROMPT_VERSION,
            "provider": provider.strip().lower(),
            "model": model.strip(),
            "learner_level": learner_level.strip().lower(),
            "selection_style": selection_style.strip().lower(),
            "focus_areas": sorted(
                value.strip().lower()
                for value in focus_areas.split(",")
                if value.strip()
            ),
            "additional_criteria": " ".join(additional_criteria.split())[:1000],
        },
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def decode_optional_hex(value: str) -> str:
    try:
        return bytes.fromhex(value.strip()).decode("utf-8") if value.strip() else ""
    except (ValueError, UnicodeDecodeError):
        return ""


def recommendation_reason_language(
    prompt_profile: str, section_headings: str, explanation_context: str
) -> str:
    """Resolve the language used by the selected saved explanation.

    Bundled prompt profiles have a stable language suffix.  Recognized localized
    section headings cover legacy/default profiles and renamed custom prompts.
    For an otherwise unknown custom prompt, the model receives a short excerpt
    and is told to match that excerpt rather than falling back to English.
    """
    profile = unicodedata.normalize("NFKC", prompt_profile or "").casefold()
    profile = profile.replace("_", "-")
    for code in sorted(PROMPT_LANGUAGE_NAMES, key=len, reverse=True):
        if profile == code or profile.endswith("-" + code):
            return PROMPT_LANGUAGE_NAMES[code]

    for heading in str(section_headings or "").splitlines():
        normalized = unicodedata.normalize("NFKC", heading).casefold()
        normalized = normalized.strip().rstrip(":：").strip()
        language = EXPLANATION_HEADING_LANGUAGES.get(normalized)
        if language:
            return language

    if str(explanation_context or "").strip():
        return "Match the language used in explanation_context"
    return "English"


def recommendation_language_context(value: str, limit: int = 600) -> str:
    """Keep enough saved explanation text for unknown custom-prompt languages."""
    text = " ".join(str(value or "").split())
    return text[:limit].rstrip()


def cached_recommendations(
    preferences_database: Path,
) -> dict[str, dict[str, object]]:
    ensure_triage_schema(preferences_database)
    connection = connect_read_only(preferences_database)
    try:
        rows = connection.execute(
            f"SELECT candidate_hash, recommended, score, reason, provider, model "
            f"FROM {RECOMMENDATION_TABLE}"
        ).fetchall()
        return {
            str(row[0]): {
                "recommended": bool(row[1]),
                "score": int(row[2] or 0),
                "reason": str(row[3] or ""),
                "provider": str(row[4] or ""),
                "model": str(row[5] or ""),
            }
            for row in rows
        }
    finally:
        connection.close()


def migrate_library_hidden_vocabulary(
    database: Path, preferences_database: Path
) -> None:
    """Move Stage 2 preview dismissals into the shared preference database once."""
    library = connect_write(database)
    try:
        migrated = library.execute(
            "SELECT value FROM metadata WHERE key = ?",
            (HIDDEN_MIGRATION_METADATA_KEY,),
        ).fetchone()
        if migrated:
            return
        try:
            rows = library.execute(
                f"SELECT normalized_front, display_front, hidden_at "
                f"FROM {HIDDEN_VOCABULARY_TABLE}"
            ).fetchall()
        except sqlite3.OperationalError:
            rows = []

        ensure_triage_schema(preferences_database)
        preferences = connect_write(preferences_database)
        try:
            for row in rows:
                preferences.execute(
                    f"INSERT INTO {HIDDEN_VOCABULARY_TABLE} "
                    "(normalized_front, display_front, hidden_at) VALUES(?, ?, ?) "
                    "ON CONFLICT(normalized_front) DO UPDATE SET "
                    "display_front = excluded.display_front",
                    (str(row[0]), str(row[1] or row[0]), str(row[2] or "")),
                )
            preferences.commit()
        finally:
            preferences.close()

        library.execute(
            "INSERT INTO metadata(key, value) VALUES(?, '1') "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (HIDDEN_MIGRATION_METADATA_KEY,),
        )
        library.commit()
    finally:
        library.close()


def compact(value: object, limit: int = 150) -> str:
    text = " ".join(str(value or "").split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def display_stamp(value: object) -> str:
    text = str(value or "").replace("T", " ")
    return text[:16]


def parse_stamp(value: object) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.astimezone() if parsed.tzinfo else parsed.astimezone()
    except ValueError:
        return None


def is_after(value: object, checkpoint: str) -> bool:
    if not checkpoint:
        return True
    candidate = parse_stamp(value)
    boundary = parse_stamp(checkpoint)
    if candidate is None or boundary is None:
        return str(value or "") > checkpoint
    return candidate > boundary


def load_anki_mappings(path: Path) -> dict[str, dict[str, str]]:
    if not path.is_file():
        return {}
    data = path.read_bytes()
    encoding = "utf-16" if data.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
    parser = configparser.ConfigParser(interpolation=None)
    parser.optionxform = str
    parser.read_string(data.decode(encoding))
    mappings: dict[str, dict[str, str]] = {}
    for section in parser.sections():
        profile = parser.get(section, "profile", fallback="").strip()
        if not profile:
            continue
        mappings[profile.casefold()] = {
            "profile": profile,
            "deck": parser.get(section, "deck", fallback="").strip(),
            "model": parser.get(section, "model", fallback="").strip(),
            "japanese_field": parser.get(
                section, "japaneseField", fallback=""
            ).strip(),
        }
    return mappings


def anki_matches(
    profiles: set[str], mappings: dict[str, dict[str, str]]
) -> tuple[dict[str, set[str]], str, str]:
    """Return normalized Japanese already present in each mapped parent deck."""
    matches: dict[str, set[str]] = {}
    mapped_profiles = [
        profile
        for profile in sorted(profiles, key=str.casefold)
        if (mapping := mappings.get(profile.casefold()))
        and all(mapping.get(key) for key in ("deck", "model", "japanese_field"))
    ]
    if mapped_profiles and not anki_endpoint_available():
        for profile in mapped_profiles:
            matches[profile.casefold()] = set()
        if anki_process_running():
            message = (
                "Anki is running, but AnkiConnect did not answer on "
                "127.0.0.1:8765. Install or enable AnkiConnect, then restart Anki."
            )
        else:
            message = (
                "Anki is not running. Open Anki and keep it running while "
                "refreshing links."
            )
        return matches, "unavailable", message
    attempted = 0
    connected = 0
    messages: list[str] = []
    cache: dict[tuple[str, str, str], set[str]] = {}
    for profile in sorted(profiles, key=str.casefold):
        mapping = mappings.get(profile.casefold())
        if not mapping or not all(
            mapping.get(key) for key in ("deck", "model", "japanese_field")
        ):
            continue
        attempted += 1
        key = (mapping["deck"], mapping["model"], mapping["japanese_field"])
        if key not in cache:
            try:
                notes = notes_for_scope(mapping["deck"], mapping["model"])
                normalized: set[str] = set()
                for note in notes:
                    field = note.get("fields", {}).get(mapping["japanese_field"], {})
                    value = field.get("value", "") if isinstance(field, dict) else field
                    value = normalize_japanese(value, remove_attached_readings=True)
                    if value:
                        normalized.add(value)
                cache[key] = normalized
                connected += 1
            except (AnkiUnavailable, OSError, ValueError) as exc:
                messages.append(str(exc))
                cache[key] = set()
        matches[profile.casefold()] = cache[key]
    if connected:
        return matches, "connected", "Exact matches already found in Anki are hidden."
    if attempted:
        message = messages[0] if messages else "AnkiConnect is unavailable."
        return matches, "unavailable", message
    return matches, "unmapped", "No saved Anki mapping applies to the current candidates."


def vocabulary_front(head: str) -> str:
    term = head.split("→")[-1].strip()
    # Keep candidate-driven Anki cards consistent with cards created from a
    # Reader selection.  The shared helper also handles okurigana, e.g.
    # 抜ける（ぬける） -> 抜ける, instead of only all-kanji terms.
    term = remove_readings(term)
    term = unicodedata.normalize("NFKC", term)
    term = re.sub(r"[\s\u3000]+", " ", term).strip(" ,.;:：。・")
    return term


@dataclass
class VocabularyEntry:
    display: str
    front: str
    back: str
    meaning: str


def parse_vocabulary(content: str) -> list[VocabularyEntry]:
    entries: list[str] = []
    current = ""
    for raw_line in str(content or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = BULLET_RE.sub("", raw_line.strip())
        if not line:
            if current:
                entries.append(current.strip())
                current = ""
            continue
        if ENTRY_DASH_RE.search(line):
            if current:
                entries.append(current.strip())
            current = line
        elif current:
            current += " " + line
    if current:
        entries.append(current.strip())

    parsed: list[VocabularyEntry] = []
    for entry in entries:
        split = ENTRY_DASH_RE.split(entry, maxsplit=1)
        if len(split) != 2:
            continue
        head, meaning = split[0].strip(), split[1].strip()
        front = vocabulary_front(head)
        if not front or not meaning:
            continue
        parsed.append(
            VocabularyEntry(
                display=head,
                front=front,
                back=entry,
                meaning=meaning,
            )
        )
    return parsed


def selected_groups(connection: sqlite3.Connection) -> list[sqlite3.Row]:
    return connection.execute(
        """
        WITH selected AS (
            SELECT e.*,
                   ROW_NUMBER() OVER (
                       PARTITION BY e.group_id
                       ORDER BY e.preferred DESC, e.version DESC
                   ) AS selection_order
            FROM explanations e
        ), counts AS (
            SELECT group_id, COUNT(*) AS version_count
            FROM explanations GROUP BY group_id
        )
        SELECT g.id AS group_id, g.game_profile, g.source_japanese,
               g.updated_at, s.id AS explanation_id, s.version,
               s.prompt_profile,
               c.version_count,
               COALESCE(d.added_to_anki_at, '') AS added_to_anki_at,
               COALESCE(a.status, 'not_checked') AS anki_status
        FROM explanation_groups g
        JOIN selected s ON s.group_id = g.id AND s.selection_order = 1
        JOIN counts c ON c.group_id = g.id
        LEFT JOIN explanation_group_details d ON d.group_id = g.id
        LEFT JOIN explanation_group_anki_links a ON a.group_id = g.id
        ORDER BY g.updated_at DESC, g.id DESC
        """
    ).fetchall()


def snapshot(
    database: Path,
    output_dir: Path,
    anki_config: Path,
    preferences_database: Path,
    scope: str,
    provider: str = "openai",
    model: str = "",
    learner_level: str = "intermediate",
    selection_style: str = "balanced",
    focus_areas: str = "vocabulary,grammar,natural_phrasing,reading",
    additional_criteria_hex: str = "",
) -> int:
    output_dir.mkdir(parents=True, exist_ok=True)
    sentence_path = output_dir / "candidate_sentences.tsv"
    vocabulary_path = output_dir / "candidate_vocabulary.tsv"
    status_path = output_dir / "candidate_status.tsv"
    for path in (sentence_path, vocabulary_path, status_path):
        try:
            path.unlink()
        except FileNotFoundError:
            pass

    snapshot_through = datetime.now().astimezone().isoformat(timespec="microseconds")
    if not database.is_file():
        write_text(sentence_path, "")
        write_text(vocabulary_path, "")
        write_rows(
            status_path,
            ((encode_field(snapshot_through), "", encode_field("unavailable"),
              encode_field("The active Study Library does not exist yet."), 0, 0),),
        )
        return 0

    recommendation_profile = recommendation_profile_signature(
        provider,
        model,
        learner_level,
        selection_style,
        focus_areas,
        decode_optional_hex(additional_criteria_hex),
    )
    migrate_library_hidden_vocabulary(database, preferences_database)
    connection = connect_read_only(database)
    try:
        checkpoint_row = connection.execute(
            "SELECT value FROM metadata WHERE key = ?", (REVIEW_METADATA_KEY,)
        ).fetchone()
        checkpoint = str(checkpoint_row[0]) if checkpoint_row else ""
        groups = selected_groups(connection)
        profiles = {str(row["game_profile"] or "").strip() for row in groups}
        profiles.discard("")
        hidden = hidden_vocabulary(preferences_database)
        recommendations = cached_recommendations(preferences_database)
        if scope == "hidden":
            matches: dict[str, set[str]] = {}
            anki_code = "hidden"
            anki_message = (
                "Ignored terms apply to every Study Library and can be restored here."
            )
        else:
            matches, anki_code, anki_message = anki_matches(
                profiles, load_anki_mappings(anki_config)
            )

        sentence_rows: list[tuple[object, ...]] = []
        vocabulary: dict[object, dict[str, object]] = {}
        for row in groups:
            profile = str(row["game_profile"] or "").strip()
            updated = str(row["updated_at"] or "")
            if scope == "new" and not is_after(updated, checkpoint):
                continue

            source = str(row["source_japanese"] or "")
            normalized_source = normalize_japanese(source, remove_attached_readings=True)
            profile_matches = matches.get(profile.casefold(), set())
            linked = bool(str(row["added_to_anki_at"] or "").strip()) or (
                str(row["anki_status"] or "") == "found"
            )
            section_rows = connection.execute(
                "SELECT section_key, heading, content FROM explanation_sections "
                "WHERE explanation_id = ? ORDER BY sort_order",
                (int(row["explanation_id"]),),
            ).fetchall()
            section_headings = "\n".join(
                str(section_row["heading"] or "") for section_row in section_rows
            )
            context_candidates = [
                str(section_row["content"] or "")
                for section_row in section_rows
                if str(section_row["section_key"] or "")
                in {"translation", "analysis", "vocabulary"}
            ]
            explanation_context = recommendation_language_context(
                next((value for value in context_candidates if value.strip()), "")
            )
            reason_language = recommendation_reason_language(
                str(row["prompt_profile"] or ""),
                section_headings,
                explanation_context,
            )
            reason_language_context = (
                explanation_context
                if reason_language.startswith("Match the language")
                else ""
            )
            language_profile = (
                recommendation_profile + "\0reason-language=" + reason_language
            )
            if reason_language_context:
                language_profile += "\0" + reason_language_context
            if (
                scope != "hidden"
                and not linked
                and normalized_source not in profile_matches
            ):
                recommendation_key = candidate_hash(
                    "sentence", source, recommendation_profile=language_profile
                )
                recommendation = recommendations.get(recommendation_key)
                sentence_rows.append(
                    (
                        int(row["group_id"]),
                        encode_field(profile or "Unsorted"),
                        encode_field(display_stamp(updated)),
                        encode_field(compact(source, 190)),
                        int(row["version_count"]),
                        recommendation_key,
                        -1 if recommendation is None else int(
                            bool(recommendation["recommended"])
                        ),
                        0 if recommendation is None else int(
                            recommendation["score"]
                        ),
                        encode_field(
                            "" if recommendation is None else recommendation["reason"]
                        ),
                        encode_field(reason_language),
                        encode_field(reason_language_context),
                    )
                )

            vocabulary_section = next(
                (
                    section_row
                    for section_row in section_rows
                    if str(section_row["section_key"] or "") == "vocabulary"
                ),
                None,
            )
            if vocabulary_section is None:
                continue
            seen_in_group: set[str] = set()
            for entry in parse_vocabulary(str(vocabulary_section["content"] or "")):
                normalized_front = normalize_japanese(
                    entry.front, remove_attached_readings=True
                )
                if not normalized_front:
                    continue
                is_hidden = normalized_front in hidden
                if scope == "hidden":
                    if not is_hidden:
                        continue
                elif is_hidden or normalized_front in profile_matches:
                    continue
                dedupe_key: object = (
                    normalized_front
                    if scope == "hidden"
                    else (profile.casefold(), normalized_front)
                )
                if dedupe_key in seen_in_group:
                    continue
                seen_in_group.add(dedupe_key)
                current = vocabulary.get(dedupe_key)
                if current is None:
                    vocabulary[dedupe_key] = {
                        "group_id": int(row["group_id"]),
                        "profile": profile or "Unsorted",
                        "updated": updated,
                        "display": entry.display,
                        "front": entry.front,
                        "back": entry.back,
                        "meaning": entry.meaning,
                        "occurrences": 1,
                        "reason_language": reason_language,
                        "explanation_context": reason_language_context,
                        "language_profile": language_profile,
                    }
                else:
                    current["occurrences"] = int(current["occurrences"]) + 1
                    if updated > str(current["updated"]):
                        current.update(
                            group_id=int(row["group_id"]),
                            updated=updated,
                            display=entry.display,
                            front=entry.front,
                            back=entry.back,
                            meaning=entry.meaning,
                            reason_language=reason_language,
                            explanation_context=reason_language_context,
                            language_profile=language_profile,
                        )

        if scope == "hidden":
            for normalized_front, hidden_item in hidden.items():
                if normalized_front in vocabulary:
                    continue
                display = hidden_item["display"] or normalized_front
                vocabulary[normalized_front] = {
                    "group_id": 0,
                    "profile": "",
                    "updated": hidden_item["hidden_at"],
                    "display": display,
                    "front": display,
                    "back": display,
                    "meaning": "Not present in the active Study Library",
                    "occurrences": 0,
                    "reason_language": "English",
                    "explanation_context": "",
                    "language_profile": recommendation_profile
                    + "\0reason-language=English",
                }

        vocabulary_rows = []
        ordered_vocabulary = sorted(
            vocabulary.values(),
            key=lambda item: (str(item["updated"]), str(item["front"])),
            reverse=True,
        )
        for item in ordered_vocabulary:
            recommendation_key = candidate_hash(
                "vocabulary",
                str(item["front"]),
                str(item["back"]),
                str(item["language_profile"]),
            )
            recommendation = recommendations.get(recommendation_key)
            vocabulary_rows.append(
                (
                    int(item["group_id"]),
                    encode_field(item["profile"]),
                    encode_field(display_stamp(item["updated"])),
                    encode_field(item["display"]),
                    encode_field(item["front"]),
                    encode_field(item["back"]),
                    encode_field(compact(item["meaning"], 180)),
                    int(item["occurrences"]),
                    recommendation_key,
                    -1 if recommendation is None else int(
                        bool(recommendation["recommended"])
                    ),
                    0 if recommendation is None else int(recommendation["score"]),
                    encode_field(
                        "" if recommendation is None else recommendation["reason"]
                    ),
                    encode_field(item["reason_language"]),
                    encode_field(item["explanation_context"]),
                )
            )

        write_rows(sentence_path, sentence_rows)
        write_rows(vocabulary_path, vocabulary_rows)
        assessed_count = sum(1 for item in sentence_rows if int(item[6]) >= 0)
        assessed_count += sum(1 for item in vocabulary_rows if int(item[9]) >= 0)
        recommended_count = sum(1 for item in sentence_rows if int(item[6]) == 1)
        recommended_count += sum(
            1 for item in vocabulary_rows if int(item[9]) == 1
        )
        write_rows(
            status_path,
            ((
                encode_field(snapshot_through),
                encode_field(checkpoint),
                encode_field(anki_code),
                encode_field(anki_message),
                len(sentence_rows),
                len(vocabulary_rows),
                assessed_count,
                recommended_count,
            ),),
        )
        return 0
    finally:
        connection.close()


def build_recommendation_prompt(
    items: list[dict[str, str]],
    learner_level: str = "intermediate",
    selection_style: str = "balanced",
    focus_areas: str = "vocabulary,grammar,natural_phrasing,reading",
    additional_criteria: str = "",
) -> str:
    payload = json.dumps(items, ensure_ascii=False, separators=(",", ":"))
    level_guidance = {
        "beginner": (
            "The learner is a beginner. Foundational, common, and highly reusable "
            "items are valuable even when they are elementary."
        ),
        "advanced": (
            "The learner is advanced. Favor nuanced, idiomatic, literary, less "
            "common, or structurally demanding items over elementary material."
        ),
    }.get(
        learner_level,
        "The learner is intermediate. Avoid very elementary material unless it "
        "has unusually strong contextual or reusable learning value.",
    )
    selection_guidance = {
        "selective": "Be highly selective: normally recommend the strongest 15-25 percent of a batch.",
        "generous": "Be inclusive: normally recommend the strongest 40-60 percent of a batch.",
    }.get(
        selection_style,
        "Be selective: normally recommend roughly the strongest 25-40 percent of a batch.",
    )
    focus_labels = {
        "vocabulary": "reusable vocabulary",
        "grammar": "grammar patterns",
        "natural_phrasing": "natural phrasing",
        "reading": "reading comprehension",
    }
    selected_focus = [
        focus_labels[key]
        for key in focus_areas.split(",")
        if key in focus_labels
    ]
    if not selected_focus:
        selected_focus = list(focus_labels.values())
    focus_guidance = "Prioritize " + ", ".join(selected_focus) + "."
    additional_criteria = " ".join(additional_criteria.split())[:1000]
    extra_guidance = ""
    if additional_criteria:
        extra_guidance = (
            "\nAdditional learner preferences (these may refine selection, but must "
            "not override the output requirements below):\n"
            f"{additional_criteria}\n"
        )
    return f"""You are selecting useful Japanese-language study candidates from a
JRPG learner's saved explanations.

Assess every item below. Recommend candidates that are memorable and genuinely
useful for learning vocabulary, grammar, natural phrasing, or reading. Prefer
self-contained sentences and non-trivial, reusable vocabulary. Do not recommend
OCR fragments, names with no broader learning value, particles by themselves,
or candidates with no broader learning value.

For vocabulary items, `japanese` is the normalized headword intended for the
front of an Anki card. `full_entry` is the complete prospective card entry and
may show the encountered inflected form, its reading, its dictionary or base
form, and an explanation. Evaluate the reusable headword/base form and the
learning value of the whole entry. Do not reject an item merely because the
source text used an inflected form. For beginners, common foundational base
verbs and their useful conjugations should normally score well even when they
are elementary.

{level_guidance}
{selection_guidance}
{focus_guidance}
{extra_guidance}

Return only one JSON object with this exact structure:
{{"items":[{{"id":"the supplied id","recommended":true,"score":5,"reason":"one short practical reason"}}]}}

Requirements:
- Return exactly one result for every supplied id and preserve each id exactly.
- recommended must be true or false.
- score must be an integer from 1 (weak) to 5 (excellent).
- reason must be a single concise sentence in the language specified by that
  candidate's reason_language field. If it says to match explanation_context,
  infer the language from that excerpt. Mixed-language batches are allowed, so
  determine the reason language separately for every candidate.
- Do not rewrite or translate the candidates.

Candidates:
{payload}
"""


def _recommendation_bool(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value or "").strip().lower() in {"1", "true", "yes", "recommended"}


def generate_recommendations(
    output_dir: Path,
    preferences_database: Path,
    provider: str,
    model: str,
    learner_level: str = "intermediate",
    selection_style: str = "balanced",
    focus_areas: str = "vocabulary,grammar,natural_phrasing,reading",
    additional_criteria_hex: str = "",
    force: bool = False,
) -> int:
    status_path = output_dir / "candidate_recommendation_status.tsv"
    error_path = output_dir / "candidate_recommendation_error.txt"
    for path in (status_path, error_path):
        try:
            path.unlink()
        except FileNotFoundError:
            pass

    try:
        provider = provider.strip().lower()
        model = model.strip()
        if provider not in {"gemini", "openai"} or not model:
            raise ValueError("The recommendation provider or model is missing.")
        learner_level = learner_level.strip().lower()
        if learner_level not in {"beginner", "intermediate", "advanced"}:
            learner_level = "intermediate"
        selection_style = selection_style.strip().lower()
        if selection_style not in {"selective", "balanced", "generous"}:
            selection_style = "balanced"
        focus_areas = focus_areas.strip().lower()
        additional_criteria = decode_optional_hex(additional_criteria_hex)

        pending: list[dict[str, str]] = []
        sentence_path = output_dir / "candidate_sentences.tsv"
        if sentence_path.is_file():
            for line in sentence_path.read_text(encoding="utf-8").splitlines():
                row = line.split("\t")
                if len(row) < 9 or (not force and int(row[6]) >= 0):
                    continue
                pending.append(
                    {
                        "id": row[5],
                        "kind": "sentence",
                        "japanese": decode_field(row[3]),
                        "reason_language": decode_field(row[9])
                        if len(row) > 9
                        else "English",
                        "explanation_context": decode_field(row[10])
                        if len(row) > 10
                        else "",
                    }
                )

        vocabulary_path = output_dir / "candidate_vocabulary.tsv"
        if vocabulary_path.is_file():
            for line in vocabulary_path.read_text(encoding="utf-8").splitlines():
                row = line.split("\t")
                if len(row) < 12 or (not force and int(row[9]) >= 0):
                    continue
                pending.append(
                    {
                        "id": row[8],
                        "kind": "vocabulary",
                        "japanese": decode_field(row[4]),
                        "full_entry": decode_field(row[5]),
                        "meaning_and_context": decode_field(row[6]),
                        "reason_language": decode_field(row[12])
                        if len(row) > 12
                        else "English",
                        "explanation_context": decode_field(row[13])
                        if len(row) > 13
                        else "",
                    }
                )

        # The same Japanese can appear under more than one profile.  One stable
        # content fingerprint should be assessed and billed only once.
        unique_pending: dict[str, dict[str, str]] = {}
        for item in pending:
            unique_pending.setdefault(item["id"], item)
        pending = list(unique_pending.values())

        if not pending:
            write_rows(
                status_path,
                ((0, 0, 0, 0, encode_field(provider), encode_field(model),
                  encode_field("Every visible candidate has already been assessed.")),),
            )
            return 0

        ensure_triage_schema(preferences_database)
        valid_ids = {item["id"] for item in pending}
        results: dict[str, tuple[bool, int, str]] = {}
        batch_size = 36
        for start in range(0, len(pending), batch_size):
            batch = pending[start : start + batch_size]
            raw = call_model(
                provider,
                model,
                build_recommendation_prompt(
                    batch,
                    learner_level,
                    selection_style,
                    focus_areas,
                    additional_criteria,
                ),
            )
            parsed = extract_json_object(raw)
            model_items = parsed.get("items", [])
            if not isinstance(model_items, list):
                raise ValueError("The model returned an invalid recommendation list.")
            for item in model_items:
                if not isinstance(item, dict):
                    continue
                item_id = str(item.get("id", "")).strip()
                if item_id not in valid_ids:
                    continue
                try:
                    score = max(1, min(5, int(item.get("score", 1))))
                except (TypeError, ValueError):
                    score = 1
                reason = " ".join(str(item.get("reason", "")).split())[:260]
                results[item_id] = (
                    _recommendation_bool(item.get("recommended")), score, reason
                )

        now = datetime.now().astimezone().isoformat(timespec="microseconds")
        connection = connect_write(preferences_database)
        try:
            for item in pending:
                item_id = item["id"]
                if item_id not in results:
                    continue
                recommended, score, reason = results[item_id]
                connection.execute(
                    f"INSERT INTO {RECOMMENDATION_TABLE} "
                    "(candidate_hash, candidate_kind, recommended, score, reason, "
                    "provider, model, generated_at) VALUES(?, ?, ?, ?, ?, ?, ?, ?) "
                    "ON CONFLICT(candidate_hash) DO UPDATE SET "
                    "candidate_kind = excluded.candidate_kind, "
                    "recommended = excluded.recommended, score = excluded.score, "
                    "reason = excluded.reason, provider = excluded.provider, "
                    "model = excluded.model, generated_at = excluded.generated_at",
                    (
                        item_id, item["kind"], int(recommended), score, reason,
                        provider, model, now,
                    ),
                )
            connection.commit()
        finally:
            connection.close()

        assessed = len(results)
        recommended_count = sum(1 for value in results.values() if value[0])
        remaining = len(pending) - assessed
        action = "Reassessed" if force else "Assessed"
        qualifier = "" if force else " new"
        message = (
            f"{action} {assessed}{qualifier} candidate{'s' if assessed != 1 else ''}; "
            f"{recommended_count} marked as recommended."
        )
        if remaining:
            message += f" {remaining} could not be assessed and remain eligible."
        write_rows(
            status_path,
            ((assessed, recommended_count, assessed - recommended_count, remaining,
              encode_field(provider), encode_field(model), encode_field(message)),),
        )
        return 0
    except Exception as exc:
        write_text(error_path, str(exc))
        return 2


def finish_review(database: Path, reviewed_through: str) -> int:
    if not reviewed_through.strip():
        raise ValueError("The review snapshot timestamp is missing.")
    connection = connect_write(database)
    try:
        connection.execute(
            "INSERT INTO metadata(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (REVIEW_METADATA_KEY, reviewed_through.strip()),
        )
        connection.commit()
        return 0
    finally:
        connection.close()


def set_vocabulary_hidden(
    preferences_database: Path, front_hex: str, hidden: bool
) -> int:
    try:
        front = bytes.fromhex(front_hex).decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError("The selected vocabulary term is invalid.") from exc
    display = vocabulary_front(front)
    normalized = normalize_japanese(display, remove_attached_readings=True)
    if not normalized:
        raise ValueError("The selected vocabulary term is empty.")

    ensure_triage_schema(preferences_database)
    connection = connect_write(preferences_database)
    try:
        if hidden:
            connection.execute(
                f"INSERT INTO {HIDDEN_VOCABULARY_TABLE} "
                "(normalized_front, display_front, hidden_at) VALUES(?, ?, ?) "
                "ON CONFLICT(normalized_front) DO UPDATE SET "
                "display_front = excluded.display_front, "
                "hidden_at = excluded.hidden_at",
                (
                    normalized,
                    display,
                    datetime.now().astimezone().isoformat(timespec="microseconds"),
                ),
            )
        else:
            connection.execute(
                f"DELETE FROM {HIDDEN_VOCABULARY_TABLE} "
                "WHERE normalized_front = ?",
                (normalized,),
            )
        connection.commit()
        return 0
    finally:
        connection.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "command",
        choices=(
            "snapshot",
            "generate-recommendations",
            "finish-review",
            "hide-vocabulary",
            "restore-vocabulary",
        ),
    )
    parser.add_argument("--db", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--anki-config", type=Path, default=Path())
    parser.add_argument("--preferences-db", required=True, type=Path)
    parser.add_argument("--scope", choices=("new", "all", "hidden"), default="new")
    parser.add_argument("--reviewed-through", default="")
    parser.add_argument("--front-hex", default="")
    parser.add_argument("--provider", choices=("gemini", "openai"), default="openai")
    parser.add_argument("--model", default="")
    parser.add_argument(
        "--learner-level",
        choices=("beginner", "intermediate", "advanced"),
        default="intermediate",
    )
    parser.add_argument(
        "--selection-style",
        choices=("selective", "balanced", "generous"),
        default="balanced",
    )
    parser.add_argument(
        "--focus-areas",
        default="vocabulary,grammar,natural_phrasing,reading",
    )
    parser.add_argument("--additional-criteria-hex", default="")
    parser.add_argument("--force", action="store_true")
    return parser


def main() -> int:
    load_project_environment()
    arguments = build_parser().parse_args()
    if arguments.command == "snapshot":
        return snapshot(
            arguments.db,
            arguments.output_dir,
            arguments.anki_config,
            arguments.preferences_db,
            arguments.scope,
            arguments.provider,
            arguments.model,
            arguments.learner_level,
            arguments.selection_style,
            arguments.focus_areas,
            arguments.additional_criteria_hex,
        )
    if arguments.command == "generate-recommendations":
        return generate_recommendations(
            arguments.output_dir,
            arguments.preferences_db,
            arguments.provider,
            arguments.model,
            arguments.learner_level,
            arguments.selection_style,
            arguments.focus_areas,
            arguments.additional_criteria_hex,
            arguments.force,
        )
    if arguments.command == "finish-review":
        return finish_review(arguments.db, arguments.reviewed_through)
    return set_vocabulary_hidden(
        arguments.preferences_db,
        arguments.front_hex,
        arguments.command == "hide-vocabulary",
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Anki candidate bridge failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
