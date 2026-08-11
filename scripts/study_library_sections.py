#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Local, lossless section extraction for Study Library explanations."""

from __future__ import annotations

import re
import sqlite3
import unicodedata
from dataclasses import dataclass
from typing import Iterable


@dataclass(frozen=True)
class ExplanationSection:
    key: str
    heading: str
    content: str


_SPEAKER_HEADER_RE = re.compile(r"^\s*「([^「」\r\n]{1,80})」\s*$")
_SPEAKER_SENTENCE_PUNCTUATION_RE = re.compile(r"[。.!！、,，；;：:…｡､]")


def extract_strict_speaker_header(source_text: str) -> str:
    """Return a speaker only from a strict standalone opening header.

    A valid source begins with ``「speaker」`` on its own line and contains
    dialogue on a later non-empty line. Mixed scripts and symbols are allowed,
    including ``?``/``？`` for unknown speakers. Sentence-like punctuation is
    rejected to avoid treating quoted dialogue as a speaker name.
    """
    lines = str(source_text or "").splitlines()
    first_index = next((index for index, line in enumerate(lines) if line.strip()), -1)
    if first_index < 0:
        return ""

    match = _SPEAKER_HEADER_RE.fullmatch(lines[first_index])
    if not match:
        return ""

    # A properly formatted header identifies following dialogue; a lone quoted
    # line is not enough evidence that its contents are a speaker label.
    if not any(line.strip() for line in lines[first_index + 1 :]):
        return ""

    speaker = match.group(1).strip()
    if not speaker or _SPEAKER_SENTENCE_PUNCTUATION_RE.search(speaker):
        return ""
    return speaker


# These are the headings requested by the bundled explanation prompts. Aliases
# from the compact fallback prompt are included as well. Matching a controlled
# list is intentional: a speaker name such as "Estelle:" inside a literal gloss
# must not become a new section.
SECTION_ALIASES: dict[str, tuple[str, ...]] = {
    "original": (
        "Original Japanese", "Japanese", "Japanischer Originaltext",
        "Japonés original", "Japonais original", "Giapponese originale",
        "日本語原文", "일본어 원문", "Origineel Japans",
        "Oryginalny tekst japoński", "Japonês original",
        "Оригинальный японский текст", "Оригінальний японський текст",
        "日文原文",
    ),
    "translation": (
        "Natural English translation", "Natural English paraphrase",
        "Natürliche deutsche Übersetzung", "Traducción natural al español",
        "Traduction française naturelle", "Traduzione naturale in italiano",
        "自然な日本語での言い換え", "자연스러운 한국어 번역",
        "Natuurlijke Nederlandse vertaling", "Naturalne tłumaczenie na polski",
        "Tradução natural para português", "Естественный перевод на русский",
        "Природний переклад українською", "自然简体中文翻译", "自然繁體中文翻譯",
        "Translation", "Natural translation", "Paraphrase",
    ),
    "analysis": (
        "Detailed analysis", "Analysis", "Detaillierte Analyse",
        "Análisis detallado", "Analyse détaillée", "Analisi dettagliata",
        "詳しい分析", "상세 분석", "Gedetailleerde analyse",
        "Szczegółowa analiza", "Análise detalhada", "Подробный разбор",
        "Докладний розбір", "详细分析", "詳細分析",
    ),
    "vocabulary": (
        "Key vocabulary", "Vocabulary", "Wichtiger Wortschatz",
        "Vocabulario clave", "Vocabulaire essentiel", "Vocabolario chiave",
        "重要語彙", "핵심 어휘", "Belangrijke woordenschat",
        "Kluczowe słownictwo", "Vocabulário essencial", "Ключевая лексика",
        "Ключова лексика", "重点词汇", "重點詞彙",
    ),
    "grammar": ("Grammar points", "Grammar"),
    "nuance": (
        "Nuance and tone", "Nuance & culture", "Nuance and culture",
        "Nuancen und Tonfall", "Matices y tono", "Nuances et ton",
        "Sfumature e tono", "ニュアンスと口調", "뉘앙스와 말투",
        "Nuance en toon", "Niuanse i ton", "Nuances e tom",
        "Нюансы и тон", "Нюанси й тон", "语气与细微差别", "語氣與細微差別",
    ),
    "literal": (
        "Literal structure", "Literal gloss", "Literal gloss (optional)",
        "Wörtliche Struktur", "Estructura literal", "Structure littérale",
        "Struttura letterale", "文の直訳構造", "직역 구조",
        "Letterlijke structuur", "Struktura dosłowna", "Estrutura literal",
        "Буквальная структура", "Буквальна структура", "直译结构", "直譯結構",
    ),
    "takeaways": (
        "Key takeaways", "Takeaways", "Wichtige Punkte", "Puntos clave",
        "Points essentiels", "Punti chiave", "重要ポイント", "핵심 포인트",
        "Belangrijkste punten", "Najważniejsze informacje", "Pontos principais",
        "Ключевые моменты", "Ключові моменти", "要点", "重點",
    ),
}


SECTION_SCHEMA = """
CREATE TABLE IF NOT EXISTS explanation_sections (
    id INTEGER PRIMARY KEY,
    explanation_id INTEGER NOT NULL REFERENCES explanations(id) ON DELETE CASCADE,
    sort_order INTEGER NOT NULL,
    section_key TEXT NOT NULL DEFAULT '',
    heading TEXT NOT NULL,
    content TEXT NOT NULL,
    parser TEXT NOT NULL DEFAULT 'local-v1',
    UNIQUE(explanation_id, sort_order)
);

CREATE INDEX IF NOT EXISTS idx_sections_explanation_order
    ON explanation_sections(explanation_id, sort_order);

CREATE TABLE IF NOT EXISTS explanation_group_details (
    group_id INTEGER PRIMARY KEY REFERENCES explanation_groups(id) ON DELETE CASCADE,
    chapter TEXT NOT NULL DEFAULT '',
    speaker TEXT NOT NULL DEFAULT '',
    tags TEXT NOT NULL DEFAULT '',
    added_to_anki_at TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_group_details_chapter
    ON explanation_group_details(chapter COLLATE NOCASE);
CREATE INDEX IF NOT EXISTS idx_group_details_speaker
    ON explanation_group_details(speaker COLLATE NOCASE);

CREATE TABLE IF NOT EXISTS study_tags (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL COLLATE NOCASE UNIQUE
);

CREATE TABLE IF NOT EXISTS explanation_group_tags (
    group_id INTEGER NOT NULL REFERENCES explanation_groups(id) ON DELETE CASCADE,
    tag_id INTEGER NOT NULL REFERENCES study_tags(id) ON DELETE CASCADE,
    PRIMARY KEY(group_id, tag_id)
);

CREATE INDEX IF NOT EXISTS idx_group_tags_tag_group
    ON explanation_group_tags(tag_id, group_id);

CREATE TABLE IF NOT EXISTS explanation_group_anki_links (
    group_id INTEGER PRIMARY KEY REFERENCES explanation_groups(id) ON DELETE CASCADE,
    status TEXT NOT NULL DEFAULT 'not_checked',
    note_id INTEGER,
    checked_at TEXT NOT NULL DEFAULT '',
    profile_name TEXT NOT NULL DEFAULT '',
    deck_name TEXT NOT NULL DEFAULT '',
    model_name TEXT NOT NULL DEFAULT '',
    japanese_field TEXT NOT NULL DEFAULT '',
    explanation_field TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_group_anki_links_status
    ON explanation_group_anki_links(status);
"""


def normalize_tags(value: str) -> list[str]:
    """Return stable, case-insensitively unique comma-separated tags."""
    result: list[str] = []
    seen: set[str] = set()
    for raw_tag in str(value or "").split(","):
        tag = re.sub(r"\s+", " ", raw_tag).strip()[:120]
        folded = tag.casefold()
        if not tag or folded in seen:
            continue
        seen.add(folded)
        result.append(tag)
    return result


def replace_group_tags(
    connection: sqlite3.Connection, group_id: int, value: str
) -> str:
    """Synchronize normalized tag links and return canonical legacy text."""
    tags = normalize_tags(value)
    connection.execute(
        "DELETE FROM explanation_group_tags WHERE group_id = ?", (group_id,)
    )
    for tag in tags:
        connection.execute("INSERT OR IGNORE INTO study_tags(name) VALUES(?)", (tag,))
        tag_row = connection.execute(
            "SELECT id FROM study_tags WHERE name = ? COLLATE NOCASE", (tag,)
        ).fetchone()
        if tag_row is not None:
            connection.execute(
                "INSERT OR IGNORE INTO explanation_group_tags(group_id, tag_id) "
                "VALUES(?, ?)",
                (group_id, int(tag_row[0])),
            )
    connection.execute(
        "DELETE FROM study_tags WHERE NOT EXISTS ("
        "SELECT 1 FROM explanation_group_tags gt WHERE gt.tag_id = study_tags.id)"
    )
    return ", ".join(tags)


def _normalize_heading(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).strip()
    value = re.sub(r"^#{1,6}\s*", "", value)
    value = re.sub(r"^\d{1,2}[.)]\s*", "", value)
    value = re.sub(r"^[*_]{1,3}|[*_]{1,3}$", "", value).strip()
    value = value.rstrip(":：").strip()
    value = re.sub(r"\s+", " ", value)
    return value.casefold()


ALIAS_TO_KEY = {
    _normalize_heading(alias): key
    for key, aliases in SECTION_ALIASES.items()
    for alias in aliases
}


def _recognized_heading(line: str) -> tuple[str, str] | None:
    stripped = line.strip()
    if not stripped or len(stripped) > 100:
        return None
    display = re.sub(r"^#{1,6}\s*", "", stripped)
    display = re.sub(r"^\d{1,2}[.)]\s*", "", display).strip()
    emphasized = re.fullmatch(
        r"[*_]{1,3}\s*(.+?)\s*[*_]{1,3}\s*([:：])?", display
    )
    if emphasized:
        display = emphasized.group(1).rstrip(":：").strip()
        had_marker = True
    else:
        had_marker = display.endswith((":", "：")) or stripped.startswith("#")
        display = display.rstrip(":：").strip()
    if not had_marker:
        return None
    key = ALIAS_TO_KEY.get(_normalize_heading(display))
    return (key, display) if key else None


def parse_explanation_sections(raw_text: str) -> list[ExplanationSection]:
    """Extract recognized sections while preserving all source text elsewhere."""
    raw_text = str(raw_text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = raw_text.split("\n")
    found: list[ExplanationSection] = []
    preamble: list[str] = []
    current: tuple[str, str] | None = None
    content: list[str] = []

    def finish() -> None:
        nonlocal content
        if current is None:
            return
        body = "\n".join(content).strip()
        if body:
            found.append(ExplanationSection(current[0], current[1], body))
        content = []

    for line in lines:
        heading = _recognized_heading(line)
        if heading:
            if current is None:
                preamble = content
            else:
                finish()
            current = heading
            content = []
        else:
            content.append(line)
    finish()

    # A sectioned view is useful only when the response actually has structure.
    # Otherwise expose one lossless fallback instead of implying that arbitrary
    # model formatting was parsed successfully.
    if len(found) < 2:
        return [ExplanationSection("full", "Full explanation", raw_text.strip())]

    preamble_text = "\n".join(preamble).strip()
    if preamble_text:
        found.insert(0, ExplanationSection("overview", "Overview", preamble_text))
    return found


def ensure_section_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(SECTION_SCHEMA)
    explanation_columns = {
        str(row[1])
        for row in connection.execute("PRAGMA table_info(explanations)")
    }
    if "manual_original_text" not in explanation_columns:
        connection.execute(
            "ALTER TABLE explanations "
            "ADD COLUMN manual_original_text TEXT NOT NULL DEFAULT ''"
        )
    if "manually_edited_at" not in explanation_columns:
        connection.execute(
            "ALTER TABLE explanations "
            "ADD COLUMN manually_edited_at TEXT NOT NULL DEFAULT ''"
        )
    detail_columns = {
        str(row[1])
        for row in connection.execute("PRAGMA table_info(explanation_group_details)")
    }
    if "added_to_anki_at" not in detail_columns:
        connection.execute(
            "ALTER TABLE explanation_group_details "
            "ADD COLUMN added_to_anki_at TEXT NOT NULL DEFAULT ''"
        )
    # Migrate Stage 4 comma-separated tags once, without changing their visible
    # representation. Keeping the text column synchronized provides a safe
    # fallback for older builds and existing search behavior.
    legacy_rows = connection.execute(
        "SELECT d.group_id, d.tags FROM explanation_group_details d "
        "WHERE d.tags <> '' AND NOT EXISTS ("
        "SELECT 1 FROM explanation_group_tags gt WHERE gt.group_id = d.group_id)"
    ).fetchall()
    for group_id, tags in legacy_rows:
        canonical = replace_group_tags(connection, int(group_id), str(tags or ""))
        connection.execute(
            "UPDATE explanation_group_details SET tags = ? WHERE group_id = ?",
            (canonical, int(group_id)),
        )
    connection.execute(
        "INSERT OR REPLACE INTO metadata(key, value) VALUES('schema_version', '6')"
    )
    connection.execute("PRAGMA user_version = 6")


def replace_explanation_sections(
    connection: sqlite3.Connection, explanation_id: int, raw_text: str
) -> int:
    sections = parse_explanation_sections(raw_text)
    connection.execute(
        "DELETE FROM explanation_sections WHERE explanation_id = ?",
        (explanation_id,),
    )
    connection.executemany(
        "INSERT INTO explanation_sections("
        "explanation_id, sort_order, section_key, heading, content, parser"
        ") VALUES(?, ?, ?, ?, ?, 'local-v1')",
        (
            (explanation_id, order, section.key, section.heading, section.content)
            for order, section in enumerate(sections, start=1)
        ),
    )
    return len(sections)


def backfill_missing_sections(connection: sqlite3.Connection) -> tuple[int, int]:
    rows: Iterable[sqlite3.Row | tuple] = connection.execute(
        "SELECT e.id, e.raw_text FROM explanations e "
        "WHERE NOT EXISTS (SELECT 1 FROM explanation_sections s "
        "WHERE s.explanation_id = e.id) ORDER BY e.id"
    )
    explanation_count = 0
    section_count = 0
    for explanation_id, raw_text in rows:
        section_count += replace_explanation_sections(
            connection, int(explanation_id), str(raw_text or "")
        )
        explanation_count += 1
    return explanation_count, section_count
