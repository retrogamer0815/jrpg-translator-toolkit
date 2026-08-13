#!/usr/bin/env python
"""Focused regression tests for hidden Study Library grammar metadata."""

from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from study_library_sections import ensure_section_schema, extract_key_grammar_metadata


def main() -> None:
    visible, values = extract_key_grammar_metadata(
        "Explanation text\n\n"
        "[[JRPG_TRANSLATOR_METADATA]]\n"
        '{"key_grammar":["～てほしい (want someone to...)",'
        '"～ながら (while...)","ignored third"]}\n'
        "[[/JRPG_TRANSLATOR_METADATA]]"
    )
    assert visible == "Explanation text"
    assert values == [
        "～てほしい (want someone to...)",
        "～ながら (while...)",
    ]

    visible, values = extract_key_grammar_metadata(
        "Still useful\n[[JRPG_TRANSLATOR_METADATA]]\n{not json}\n"
        "[[/JRPG_TRANSLATOR_METADATA]]"
    )
    assert visible == "Still useful"
    assert values == []

    visible, values = extract_key_grammar_metadata(
        "Explanation\n[[JRPG_TRANSLATOR_METADATA]]\n"
        "```json\n{\"key_grammar\":[\"～ながら (while...)\"]}\n```\n"
        "[[/JRPG_TRANSLATOR_METADATA]]"
    )
    assert visible == "Explanation"
    assert values == ["～ながら (while...)"]

    visible, values = extract_key_grammar_metadata("Old explanation without metadata")
    assert visible == "Old explanation without metadata"
    assert values == []

    connection = sqlite3.connect(":memory:")
    connection.executescript(
        "CREATE TABLE metadata(key TEXT PRIMARY KEY, value TEXT NOT NULL);"
        "CREATE TABLE explanations("
        "id INTEGER PRIMARY KEY, raw_text TEXT NOT NULL, "
        "manual_original_text TEXT NOT NULL DEFAULT '', "
        "manually_edited_at TEXT NOT NULL DEFAULT ''"
        ");"
        "CREATE TABLE explanation_groups(id INTEGER PRIMARY KEY);"
    )
    ensure_section_schema(connection)
    columns = {
        str(row[1]) for row in connection.execute("PRAGMA table_info(explanations)")
    }
    assert "key_grammar" in columns
    assert connection.execute("PRAGMA user_version").fetchone()[0] == 7
    connection.close()


if __name__ == "__main__":
    main()
