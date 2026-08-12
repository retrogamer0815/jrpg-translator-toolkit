#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""AnkiConnect bridge for Study Library link checks and explicit note adds."""

from __future__ import annotations

import argparse
import base64
import hashlib
import html
import io
import json
import os
import re
import socket
import sqlite3
import subprocess
import sys
import unicodedata
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Iterable

# The bundled embeddable Python uses an isolated ._pth file, so it does not
# automatically add the launched script's folder to sys.path.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from study_library_sections import _recognized_heading

ANKI_URL = "http://127.0.0.1:8765"
ANKI_CONNECT_VERSION = 6
READING_RE = re.compile(
    r"([一-龯々〆ヵヶ][一-龯々〆ヵヶぁ-ゖー]*)[ \t]*[\(（][ぁ-ゖー・ \t]+[\)）]"
)
BR_RE = re.compile(r"(?is)<\s*br\s*/?\s*>")
BLOCK_END_RE = re.compile(
    r"(?is)</\s*(?:div|p|li|tr|h[1-6]|blockquote|pre|table|ul|ol)\s*>"
)
TAG_RE = re.compile(r"(?is)<[^>]+>")


class AnkiUnavailable(RuntimeError):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


def anki_endpoint_available(timeout: float = 0.15) -> bool:
    """Return quickly when the local AnkiConnect listener is not available.

    On Windows, urllib can spend roughly two seconds establishing that a closed
    localhost port is unavailable.  Candidate-list refreshes only need a cheap
    reachability check before attempting the normal AnkiConnect requests.
    """
    try:
        with socket.create_connection(("127.0.0.1", 8765), timeout=timeout):
            return True
    except OSError:
        return False


def encode_field(value: object) -> str:
    if value is None:
        return ""
    return str(value).encode("utf-8").hex()


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(text, encoding="utf-8", newline="\n")
    os.replace(temporary, path)


def write_rows(path: Path, rows: Iterable[Iterable[object]]) -> None:
    lines = ["\t".join(str(field) for field in row) for row in rows]
    write_text(path, "\n".join(lines) + ("\n" if lines else ""))


def clear_outputs(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for name in (
        "anki_status.tsv",
        "anki_decks.tsv",
        "anki_models.tsv",
        "anki_refresh.tsv",
        "anki_add.tsv",
    ):
        try:
            (output_dir / name).unlink()
        except FileNotFoundError:
            pass


def anki_process_running() -> bool:
    """Best-effort distinction between Anki being closed and AnkiConnect missing."""
    if os.name != "nt":
        return False
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            check=False,
            capture_output=True,
            text=True,
            timeout=4,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return False
    process_text = result.stdout.casefold()
    return '"anki.exe"' in process_text or '"anki-console.exe"' in process_text


def invoke(action: str, params: dict | None = None, timeout: float = 8.0):
    payload = json.dumps(
        {"action": action, "version": ANKI_CONNECT_VERSION, "params": params or {}},
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        ANKI_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as exc:
        if anki_process_running():
            raise AnkiUnavailable(
                "connect_missing",
                "Anki is running, but AnkiConnect did not answer on 127.0.0.1:8765. "
                "Install or enable AnkiConnect, then restart Anki.",
            ) from exc
        raise AnkiUnavailable(
            "anki_not_running",
            "Anki is not running. Open Anki and keep it running while refreshing links.",
        ) from exc
    try:
        decoded = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AnkiUnavailable(
            "invalid_response", "AnkiConnect returned an invalid response."
        ) from exc
    error = decoded.get("error")
    if error:
        raise AnkiUnavailable("anki_error", str(error))
    return decoded.get("result")


def write_status(
    output_dir: Path, code: str, version: object = "", message: str = ""
) -> None:
    write_rows(
        output_dir / "anki_status.tsv",
        ((code, str(version or ""), encode_field(message)),),
    )


def probe(output_dir: Path) -> tuple[bool, object]:
    try:
        version = invoke("version")
    except AnkiUnavailable as exc:
        write_status(output_dir, exc.code, "", exc.message)
        return False, ""
    write_status(
        output_dir,
        "connected",
        version,
        f"Connected to AnkiConnect (API {version}).",
    )
    return True, version


def discover(output_dir: Path) -> int:
    clear_outputs(output_dir)
    connected, _version = probe(output_dir)
    if not connected:
        write_text(output_dir / "anki_decks.tsv", "")
        write_text(output_dir / "anki_models.tsv", "")
        return 0

    decks = sorted((str(name) for name in (invoke("deckNames") or [])), key=str.casefold)
    models = sorted((str(name) for name in (invoke("modelNames") or [])), key=str.casefold)
    write_rows(output_dir / "anki_decks.tsv", ((encode_field(name),) for name in decks))
    model_rows: list[tuple[str, str]] = []
    for model in models:
        fields = invoke("modelFieldNames", {"modelName": model}) or []
        if not fields:
            model_rows.append((encode_field(model), ""))
            continue
        for field in fields:
            model_rows.append((encode_field(model), encode_field(str(field))))
    write_rows(output_dir / "anki_models.tsv", model_rows)
    return 0


def remove_readings(value: str) -> str:
    return READING_RE.sub(r"\1", value)


def plain_text(value: object) -> str:
    text = str(value or "")
    text = BR_RE.sub("\n", text)
    text = BLOCK_END_RE.sub("\n", text)
    text = TAG_RE.sub("", text)
    return html.unescape(text).replace("\xa0", " ")


def anki_html(value: str) -> str:
    """Preserve reviewed plain-text line breaks in Anki's HTML fields."""
    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    return html.escape(normalized).replace("\n", "<br>")


def _anki_source_image_html(filename: str) -> str:
    """Return the shared, phone-friendly source-image markup."""
    safe_filename = html.escape(filename, quote=True)
    return (
        '<div class="jrpg-translator-source-image" style="margin:0.75em 0">'
        f'<img src="{safe_filename}" style="max-width:100%;height:auto">'
        "</div>"
    )


def anki_explanation_html(value: str, filename: str = "") -> str:
    """Format a structured explanation with portable in-card navigation.

    Only headings recognized by the same conservative parser as the Study
    Reader become navigation targets. Custom/unstructured responses keep the
    former escaped plain-text layout. Anki's card WebViews do not reliably
    navigate ordinary fragment links inside field HTML, so the generated
    controls use fixed local ``scrollIntoView()`` calls. Model/user text is
    never inserted into those calls.
    """
    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    heading_rows: list[tuple[int, str, str]] = []
    for index, line in enumerate(lines):
        recognized = _recognized_heading(line)
        if recognized is not None:
            heading_rows.append((index, recognized[0], recognized[1]))

    # Do not imply reliable structure when the model/custom prompt returned
    # fewer than two recognized sections. Preserve the established fallback,
    # including the conservative screenshot placement used by older cards.
    if len(heading_rows) < 2:
        if not filename:
            return anki_html(value)
        insert_at: int | None = None
        original_heading = next(
            (index for index, key, _heading in heading_rows if key == "original"),
            None,
        )
        if original_heading is not None:
            insert_at = next(
                (
                    index
                    for index, _key, _heading in heading_rows
                    if index > original_heading
                ),
                len(lines),
            )
        else:
            insert_at = next(
                (index for index, key, _heading in heading_rows if key == "translation"),
                None,
            )
        if insert_at is None:
            saw_text = False
            for index, line in enumerate(lines):
                if line.strip():
                    saw_text = True
                elif saw_text:
                    insert_at = index
                    break
        if insert_at is None:
            insert_at = len(lines)
        before = anki_html("\n".join(lines[:insert_at]).rstrip("\n"))
        after = anki_html("\n".join(lines[insert_at:]).lstrip("\n"))
        image = _anki_source_image_html(filename)
        if before and after:
            return before + "<br><br>" + image + "<br><br>" + after
        if before:
            return before + "<br><br>" + image
        if after:
            return image + "<br><br>" + after
        return image

    id_counts: dict[str, int] = {}
    sections: list[dict[str, object]] = []
    for position, (line_index, key, heading) in enumerate(heading_rows):
        end_index = (
            heading_rows[position + 1][0]
            if position + 1 < len(heading_rows)
            else len(lines)
        )
        id_counts[key] = id_counts.get(key, 0) + 1
        suffix = "" if id_counts[key] == 1 else f"-{id_counts[key]}"
        sections.append(
            {
                "key": key,
                "heading": heading,
                "anchor": f"jrpg-section-{key}{suffix}",
                "content": "\n".join(lines[line_index + 1 : end_index]).strip(),
            }
        )

    link_style = "color:#55baf2;text-decoration:underline;cursor:pointer"
    toc_links = " &middot; ".join(
        f'<span class="tappable" '
        f'onclick="document.getElementById(\'{section["anchor"]}\')'
        f'.scrollIntoView();" style="{link_style}">'
        f'{html.escape(str(section["heading"]))}</span>'
        for section in sections
    )
    parts = [
        '<div id="jrpg-explanation-top"></div>',
        '<div class="jrpg-translator-toc" '
        'style="border:1px solid #888;padding:10px;margin-bottom:18px;'
        'border-radius:6px">'
        "<b>📑 Jump to section:</b><br>"
        f"{toc_links}</div>",
    ]

    preamble = "\n".join(lines[: heading_rows[0][0]]).strip()
    if preamble:
        parts.append(f'<div class="jrpg-explanation-preamble">{anki_html(preamble)}</div>')

    image_inserted = False
    has_original = any(section["key"] == "original" for section in sections)
    for section_index, section in enumerate(sections):
        # If a custom prompt omitted Original Japanese, retain the useful old
        # behavior of placing context immediately before the translation. When
        # there is no translation either, place it after the first section.
        if (
            filename
            and not image_inserted
            and not has_original
            and section["key"] == "translation"
        ):
            parts.append(_anki_source_image_html(filename))
            image_inserted = True

        parts.append("<hr>")
        parts.append(
            f'<h3 id="{section["anchor"]}">'
            f'{html.escape(str(section["heading"]))}:</h3>'
        )
        content = str(section["content"])
        if content:
            parts.append(f'<div class="jrpg-explanation-section">{anki_html(content)}</div>')

        if filename and not image_inserted and section["key"] == "original":
            parts.append(_anki_source_image_html(filename))
            image_inserted = True
        elif (
            filename
            and not image_inserted
            and not has_original
            and section_index == 0
            and not any(item["key"] == "translation" for item in sections)
        ):
            parts.append(_anki_source_image_html(filename))
            image_inserted = True

        parts.append(
            '<p><span class="tappable" '
            'onclick="document.getElementById(\'jrpg-explanation-top\')'
            f'.scrollIntoView();" style="{link_style}">↑ Back to top</span></p>'
        )

    if filename and not image_inserted:
        parts.append(_anki_source_image_html(filename))
    return "".join(parts)


def anki_html_with_screenshot(value: str, filename: str) -> str:
    """Backward-compatible wrapper for structured explanation export."""
    return anki_explanation_html(value, filename)


def anki_html_with_screenshot_at_end(value: str, filename: str) -> str:
    """Append context media to a short reviewed vocabulary definition."""
    image = _anki_source_image_html(filename)
    content = anki_html(value)
    return content + ("<br><br>" if content else "") + image


def prepare_anki_screenshot(source: Path) -> tuple[str, str, int, tuple[int, int]]:
    """Return a compact, phone-friendly JPEG and deterministic Anki filename."""
    try:
        from PIL import Image, ImageOps

        with Image.open(source) as opened:
            image = ImageOps.exif_transpose(opened)
            if image.mode in ("RGBA", "LA") or (
                image.mode == "P" and "transparency" in image.info
            ):
                rgba = image.convert("RGBA")
                background = Image.new("RGB", rgba.size, "black")
                background.paste(rgba, mask=rgba.getchannel("A"))
                image = background
            else:
                image = image.convert("RGB")

            # 1280 px remains readable for game dialogue on a phone while avoiding
            # multi-megabyte PNG copies in Anki's synchronized media collection.
            image.thumbnail((1280, 1280), Image.Resampling.LANCZOS)
            size = image.size
            encoded = io.BytesIO()
            image.save(
                encoded,
                format="JPEG",
                quality=82,
                optimize=True,
                progressive=True,
                subsampling=2,
            )
    except ImportError as exc:
        raise ValueError(
            "Screenshot export needs the Pillow image package in the bundled Python."
        ) from exc
    except (OSError, ValueError) as exc:
        raise ValueError(f"The selected source screenshot could not be converted: {exc}")

    data = encoded.getvalue()
    digest = hashlib.sha256(data).hexdigest()[:20]
    filename = f"jrpg_translator_{digest}.jpg"
    return filename, base64.b64encode(data).decode("ascii"), len(data), size


def normalize_japanese(value: object, remove_attached_readings: bool = False) -> str:
    text = plain_text(value).replace("\r\n", "\n").replace("\r", "\n")
    if remove_attached_readings:
        text = remove_readings(text)
    text = unicodedata.normalize("NFKC", text)
    lines = [re.sub(r"[ \t\u3000]+", " ", line).strip() for line in text.split("\n")]
    return "\n".join(line for line in lines if line).strip()


def deck_scope_names(selected: str, available: Iterable[str]) -> list[str]:
    selected_folded = selected.casefold()
    prefix = selected_folded + "::"
    scoped = [
        deck
        for deck in available
        if deck.casefold() == selected_folded or deck.casefold().startswith(prefix)
    ]
    return sorted(scoped, key=str.casefold)


def notes_for_scope(deck_name: str, model_name: str) -> list[dict]:
    available = [str(name) for name in (invoke("deckNames") or [])]
    scoped_decks = deck_scope_names(deck_name, available)
    if not scoped_decks:
        raise ValueError(f"The mapped Anki deck was not found: {deck_name}")

    note_ids: set[int] = set()
    for scoped_deck in scoped_decks:
        found = invoke("findNotes", {"query": f'deck:"{scoped_deck}"'}) or []
        note_ids.update(int(note_id) for note_id in found)

    notes: list[dict] = []
    ordered_ids = sorted(note_ids)
    for offset in range(0, len(ordered_ids), 250):
        batch = ordered_ids[offset : offset + 250]
        notes.extend(invoke("notesInfo", {"notes": batch}, timeout=15.0) or [])
    if model_name:
        notes = [note for note in notes if str(note.get("modelName", "")) == model_name]
    return notes


def ensure_link_schema(connection: sqlite3.Connection) -> None:
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 10000")
    connection.execute(
        """
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
        )
        """
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_group_anki_links_status "
        "ON explanation_group_anki_links(status)"
    )


def record_group_link(
    database: Path,
    group_id: int,
    note_id: int,
    checked_at: str,
    profile_name: str,
    deck_name: str,
    model_name: str,
    japanese_field: str,
    explanation_field: str,
) -> None:
    connection = sqlite3.connect(database, timeout=10)
    try:
        ensure_link_schema(connection)
        connection.execute(
            """
            INSERT INTO explanation_group_anki_links(
                group_id, status, note_id, checked_at, profile_name,
                deck_name, model_name, japanese_field, explanation_field
            ) VALUES(?, 'found', ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(group_id) DO UPDATE SET
                status='found', note_id=excluded.note_id,
                checked_at=excluded.checked_at,
                profile_name=excluded.profile_name,
                deck_name=excluded.deck_name,
                model_name=excluded.model_name,
                japanese_field=excluded.japanese_field,
                explanation_field=excluded.explanation_field
            """,
            (
                group_id,
                note_id,
                checked_at,
                profile_name,
                deck_name,
                model_name,
                japanese_field,
                explanation_field,
            ),
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def refresh(database: Path, output_dir: Path) -> int:
    clear_outputs(output_dir)
    connected, _version = probe(output_dir)
    if not connected:
        # Connection failures deliberately leave the last successful link audit intact.
        return 0

    profile_name = os.environ.get("STUDY_ANKI_PROFILE", "").strip()
    deck_name = os.environ.get("STUDY_ANKI_DECK", "").strip()
    model_name = os.environ.get("STUDY_ANKI_MODEL", "").strip()
    japanese_field = os.environ.get("STUDY_ANKI_JAPANESE_FIELD", "").strip()
    explanation_field = os.environ.get("STUDY_ANKI_EXPLANATION_FIELD", "").strip()
    if not profile_name:
        raise ValueError("No Study Library profile was selected.")
    if not deck_name:
        raise ValueError("No Anki deck or parent deck was selected.")
    if not model_name:
        raise ValueError("No Anki note type was selected.")
    if not japanese_field:
        raise ValueError("No Japanese field was selected.")
    if not database.is_file():
        raise ValueError("The selected Study Library database does not exist.")

    notes = notes_for_scope(deck_name, model_name)
    note_lookup: dict[str, list[int]] = {}
    for note in notes:
        fields = note.get("fields") or {}
        field = fields.get(japanese_field) or {}
        normalized = normalize_japanese(field.get("value", ""), False)
        if normalized:
            note_lookup.setdefault(normalized, []).append(int(note.get("noteId", 0)))

    connection = sqlite3.connect(database, timeout=10)
    connection.row_factory = sqlite3.Row
    ensure_link_schema(connection)
    try:
        groups = connection.execute(
            "SELECT id, source_japanese FROM explanation_groups "
            "WHERE game_profile = ? ORDER BY id",
            (profile_name,),
        ).fetchall()
        checked_at = datetime.now().astimezone().isoformat(timespec="seconds")
        found_count = 0
        duplicate_count = 0
        connection.execute("BEGIN IMMEDIATE")
        for group in groups:
            normalized = normalize_japanese(group["source_japanese"], True)
            matches = note_lookup.get(normalized, []) if normalized else []
            status = "found" if matches else "not_found"
            note_id = min(matches) if matches else None
            if matches:
                found_count += 1
            if len(matches) > 1:
                duplicate_count += 1
            connection.execute(
                """
                INSERT INTO explanation_group_anki_links(
                    group_id, status, note_id, checked_at, profile_name,
                    deck_name, model_name, japanese_field, explanation_field
                ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(group_id) DO UPDATE SET
                    status=excluded.status,
                    note_id=excluded.note_id,
                    checked_at=excluded.checked_at,
                    profile_name=excluded.profile_name,
                    deck_name=excluded.deck_name,
                    model_name=excluded.model_name,
                    japanese_field=excluded.japanese_field,
                    explanation_field=excluded.explanation_field
                """,
                (
                    int(group["id"]),
                    status,
                    note_id,
                    checked_at,
                    profile_name,
                    deck_name,
                    model_name,
                    japanese_field,
                    explanation_field,
                ),
            )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    write_rows(
        output_dir / "anki_refresh.tsv",
        ((len(groups), found_count, len(groups) - found_count, duplicate_count, encode_field(checked_at)),),
    )
    return 0


def write_add_result(
    output_dir: Path, code: str, note_id: object = "", message: str = ""
) -> None:
    write_rows(
        output_dir / "anki_add.tsv",
        ((code, str(note_id or ""), encode_field(message)),),
    )


def add_note(database: Path, output_dir: Path) -> int:
    """Create one explicitly reviewed explanation or vocabulary note."""
    clear_outputs(output_dir)
    connected, _version = probe(output_dir)
    if not connected:
        write_add_result(
            output_dir,
            "unavailable",
            "",
            "AnkiConnect is unavailable. Nothing was added to Anki.",
        )
        return 0

    profile_name = os.environ.get("STUDY_ANKI_PROFILE", "").strip()
    parent_deck = os.environ.get("STUDY_ANKI_PARENT_DECK", "").strip()
    deck_name = os.environ.get("STUDY_ANKI_DECK", "").strip()
    model_name = os.environ.get("STUDY_ANKI_MODEL", "").strip()
    japanese_field = os.environ.get("STUDY_ANKI_JAPANESE_FIELD", "").strip()
    explanation_field = os.environ.get(
        "STUDY_ANKI_EXPLANATION_FIELD", ""
    ).strip()
    card_kind = os.environ.get("STUDY_ANKI_CARD_KIND", "explanation").strip().lower()
    group_id_raw = os.environ.get("STUDY_ANKI_GROUP_ID", "").strip()
    version_raw = os.environ.get("STUDY_ANKI_VERSION", "").strip()
    front_path = Path(os.environ.get("STUDY_ANKI_FRONT_PATH", ""))
    back_path = Path(os.environ.get("STUDY_ANKI_BACK_PATH", ""))
    include_screenshot = os.environ.get("STUDY_ANKI_INCLUDE_SCREENSHOT", "") == "1"
    screenshot_path_raw = os.environ.get("STUDY_ANKI_SCREENSHOT_PATH", "").strip()

    if card_kind not in {"explanation", "vocabulary"}:
        raise ValueError("The requested Anki card type is not supported.")
    link_explanation = card_kind == "explanation"

    if not profile_name:
        raise ValueError("No Study Library profile is selected.")
    if not parent_deck:
        raise ValueError("No mapped Anki parent deck is available.")
    if not deck_name:
        raise ValueError("No destination Anki deck is selected.")
    if not model_name or not japanese_field or not explanation_field:
        raise ValueError("The mapped note type or field mapping is incomplete.")
    if japanese_field == explanation_field:
        raise ValueError("The Japanese and explanation fields must be different.")
    if not group_id_raw.isdigit() or int(group_id_raw) <= 0:
        raise ValueError("No valid Study Library explanation is selected.")
    if not database.is_file():
        raise ValueError("The selected Study Library database does not exist.")
    if not front_path.is_file() or not back_path.is_file():
        raise ValueError("The reviewed Anki entry could not be read.")

    front = front_path.read_text(encoding="utf-8").strip()
    back = back_path.read_text(encoding="utf-8").strip()
    if not front:
        raise ValueError("The Japanese/front field is empty.")
    if not back:
        raise ValueError("The explanation/back field is empty.")

    group_id = int(group_id_raw)
    screenshot_source: Path | None = None
    with sqlite3.connect(database, timeout=10) as verification_connection:
        exists = verification_connection.execute(
            "SELECT 1 FROM explanation_groups WHERE id = ? AND game_profile = ?",
            (group_id, profile_name),
        ).fetchone()
        if not exists:
            raise ValueError(
                "The selected Study Library explanation no longer belongs to this profile."
            )
        if include_screenshot:
            if not version_raw.isdigit() or int(version_raw) <= 0:
                raise ValueError("No valid Study Library explanation version is selected.")
            if not screenshot_path_raw:
                raise ValueError("No source screenshot is selected for this card.")
            screenshot_source = Path(screenshot_path_raw).resolve()
            media_rows = verification_connection.execute(
                "SELECT m.relative_path FROM media AS m "
                "JOIN explanations AS e ON e.id = m.explanation_id "
                "WHERE e.group_id = ? AND e.version = ?",
                (group_id, int(version_raw)),
            ).fetchall()
            allowed_media = {
                (database.parent / Path(str(row[0] or ""))).resolve()
                for row in media_rows
            }
            if screenshot_source not in allowed_media or not screenshot_source.is_file():
                raise ValueError(
                    "The selected screenshot is missing or does not belong to this explanation."
                )
    available_decks = [str(name) for name in (invoke("deckNames") or [])]
    scoped_decks = deck_scope_names(parent_deck, available_decks)
    scoped_folded = {name.casefold() for name in scoped_decks}
    if deck_name.casefold() not in scoped_folded:
        raise ValueError(
            "The destination deck is not the mapped parent deck or one of its subdecks."
        )

    normalized_front = normalize_japanese(front, False)
    for existing in notes_for_scope(parent_deck, model_name):
        fields = existing.get("fields") or {}
        field = fields.get(japanese_field) or {}
        if normalize_japanese(field.get("value", ""), False) == normalized_front:
            note_id = int(existing.get("noteId", 0))
            if link_explanation:
                record_group_link(
                    database,
                    group_id,
                    note_id,
                    datetime.now().astimezone().isoformat(timespec="seconds"),
                    profile_name,
                    parent_deck,
                    model_name,
                    japanese_field,
                    explanation_field,
                )
            write_add_result(
                output_dir,
                "duplicate",
                note_id,
                "An exact normalized Japanese match already exists in the mapped deck.",
            )
            return 0

    screenshot_filename = ""
    screenshot_data = ""
    if include_screenshot and screenshot_source is not None:
        (
            screenshot_filename,
            screenshot_data,
            _byte_count,
            _image_size,
        ) = prepare_anki_screenshot(screenshot_source)

    back_html = (
        anki_html(back)
        if card_kind == "vocabulary"
        else anki_explanation_html(back, screenshot_filename)
    )
    if screenshot_filename and card_kind == "vocabulary":
        back_html = anki_html_with_screenshot_at_end(back, screenshot_filename)

    note = {
        "deckName": deck_name,
        "modelName": model_name,
        "fields": {
            japanese_field: anki_html(front),
            explanation_field: back_html,
        },
        "options": {"allowDuplicate": False},
        "tags": [],
    }
    validation = invoke("canAddNotesWithErrorDetail", {"notes": [note]}) or []
    if not validation or not bool(validation[0].get("canAdd")):
        reason = "Anki rejected this entry."
        if validation:
            reason = str(validation[0].get("error") or reason)
        write_add_result(output_dir, "blocked", "", reason)
        return 0

    media_was_present = False
    stored_media_filename = screenshot_filename
    if screenshot_filename:
        media_was_present = bool(
            invoke("retrieveMediaFile", {"filename": screenshot_filename})
        )
        stored_filename = invoke(
            "storeMediaFile",
            {"filename": screenshot_filename, "data": screenshot_data},
        )
        if not stored_filename:
            write_add_result(
                output_dir, "error", "", "Anki could not store the source screenshot."
            )
            return 0
        stored_media_filename = str(stored_filename)
        if stored_media_filename != screenshot_filename:
            note["fields"][explanation_field] = (
                anki_html_with_screenshot_at_end(back, stored_media_filename)
                if card_kind == "vocabulary"
                else anki_explanation_html(back, stored_media_filename)
            )

    try:
        note_id = invoke("addNote", {"note": note})
    except Exception:
        if screenshot_filename and not media_was_present:
            try:
                invoke("deleteMediaFile", {"filename": stored_media_filename})
            except Exception:
                pass
        raise
    if not note_id:
        if screenshot_filename and not media_was_present:
            try:
                invoke("deleteMediaFile", {"filename": stored_media_filename})
            except Exception:
                pass
        write_add_result(output_dir, "error", "", "Anki did not confirm the new entry.")
        return 0

    if not link_explanation:
        write_add_result(
            output_dir,
            "added",
            int(note_id),
            f"Added the vocabulary entry to {deck_name}.",
        )
        return 0

    checked_at = datetime.now().astimezone().isoformat(timespec="seconds")
    try:
        record_group_link(
            database,
            group_id,
            int(note_id),
            checked_at,
            profile_name,
            parent_deck,
            model_name,
            japanese_field,
            explanation_field,
        )
    except Exception as exc:
        write_add_result(
            output_dir,
            "added_unlinked",
            int(note_id),
            "The entry was added to Anki, but the local link could not be saved. "
            f"Refresh Anki status to recover it. ({exc})",
        )
        return 0

    write_add_result(
        output_dir,
        "added",
        int(note_id),
        f"Added the explanation to {deck_name}.",
    )
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="JRPG Translator Anki bridge")
    parser.add_argument("command", choices=("discover", "refresh", "add-note"))
    parser.add_argument("--db", required=True)
    parser.add_argument("--output-dir", required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output_dir)
    try:
        if args.command == "discover":
            return discover(output_dir)
        if args.command == "add-note":
            return add_note(Path(args.db), output_dir)
        return refresh(Path(args.db), output_dir)
    except Exception as exc:
        # Expected configuration and Anki query errors are part of the bridge
        # protocol. Returning a status row lets the themed UI explain the exact
        # problem instead of showing a generic hidden-process exit code.
        write_status(output_dir, "error", "", str(exc))
        if args.command == "add-note":
            write_add_result(output_dir, "error", "", str(exc))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
