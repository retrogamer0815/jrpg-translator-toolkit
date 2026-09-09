"""Local vocabulary parsing shared by the Reader and Review for Anki.

No AI or Anki requests are made by this module.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from anki_bridge import remove_readings

ENTRY_DASH_RE = re.compile(r"\s*(?:—|–|--|\s-\s)\s*")
BULLET_RE = re.compile(r"^\s*(?:[*•・]|[-–—]|\d+[.)])\s+")


def vocabulary_front(head: str) -> str:
    term = head.split("→")[-1].strip()
    # Strip attached pronunciation, not kana belonging to the word itself.
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
    # Preserve the existing Review parser's entry boundaries and card contents.
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
        parsed.append(VocabularyEntry(display=head, front=front, back=entry, meaning=meaning))
    return parsed
