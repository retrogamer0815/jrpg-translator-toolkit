#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Bridge between the AutoHotkey Study Library UI and SQLite."""

from __future__ import annotations

import argparse
import os
import shutil
import sqlite3
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

# The bundled portable Python uses python312._pth isolation, which deliberately
# omits the launched script's directory from sys.path. Add this trusted local
# directory explicitly so sibling Study Library modules remain importable.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from study_library_sections import (
    backfill_missing_sections,
    ensure_section_schema,
    normalize_tags,
    replace_explanation_sections,
    replace_group_tags,
)


def encode_field(value: object) -> str:
    """Hex-encode UTF-8 so tabs/newlines can never corrupt bridge rows."""
    if value is None:
        return ""
    return str(value).encode("utf-8").hex()


def compact(value: object, limit: int = 140) -> str:
    text = " ".join(str(value or "").split())
    if len(text) > limit:
        return text[: limit - 1].rstrip() + "…"
    return text


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(text, encoding="utf-8", newline="\n")
    os.replace(temporary, path)


def write_rows(path: Path, rows: Iterable[Iterable[object]]) -> None:
    lines = ["\t".join(str(field) for field in row) for row in rows]
    write_text(path, "\n".join(lines) + ("\n" if lines else ""))


def clear_outputs(output_dir: Path, names: Iterable[str]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for name in names:
        try:
            (output_dir / name).unlink()
        except FileNotFoundError:
            pass


def ensure_database(database: Path, output_dir: Path) -> int:
    """Upgrade derived Study Library data and backfill it without touching raw text."""
    clear_outputs(output_dir, ("ensure.tsv",))
    if not database.is_file():
        write_rows(output_dir / "ensure.tsv", ((0, 0),))
        return 0
    connection = sqlite3.connect(database, timeout=10)
    try:
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 10000")
        ensure_section_schema(connection)
        connection.commit()
        connection.execute("BEGIN IMMEDIATE")
        explanations, sections = backfill_missing_sections(connection)
        connection.commit()
        write_rows(output_dir / "ensure.tsv", ((explanations, sections),))
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


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
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 10000")
    return connection


def directory_usage(path: Path) -> tuple[int, int]:
    """Return file count and byte size without following directory links."""
    if not path.exists():
        return 0, 0
    if path.is_file():
        try:
            return 1, path.stat().st_size
        except OSError:
            return 0, 0
    file_count = 0
    byte_count = 0
    for root, directories, files in os.walk(path, followlinks=False):
        # Junctions and symbolic directory links should not make a library scan
        # escape its own storage folder or count the same files repeatedly.
        directories[:] = [
            name for name in directories
            if not (Path(root) / name).is_symlink()
        ]
        for filename in files:
            candidate = Path(root) / filename
            try:
                if candidate.is_symlink():
                    continue
                byte_count += candidate.stat().st_size
                file_count += 1
            except OSError:
                continue
    return file_count, byte_count


def storage_snapshot(database: Path, output_dir: Path) -> int:
    """Report active-library disk usage without modifying library contents."""
    clear_outputs(output_dir, ("storage.tsv",))
    study_dir = database.parent
    study_dir.mkdir(parents=True, exist_ok=True)

    database_files = (
        database,
        Path(str(database) + "-wal"),
        Path(str(database) + "-shm"),
    )
    database_bytes = 0
    for database_file in database_files:
        try:
            database_bytes += database_file.stat().st_size
        except OSError:
            pass

    media_count, media_bytes = directory_usage(study_dir / "Media")
    backup_count, backup_bytes = directory_usage(study_dir / "Backups")
    trash_count, trash_bytes = directory_usage(study_dir / "Trash")
    _total_count, total_bytes = directory_usage(study_dir)
    other_bytes = max(
        0,
        total_bytes - database_bytes - media_bytes - backup_bytes - trash_bytes,
    )

    source_count = 0
    explanation_count = 0
    if database.is_file():
        connection = connect_read_only(database)
        try:
            source_count = int(
                connection.execute(
                    "SELECT COUNT(*) FROM explanation_groups"
                ).fetchone()[0]
            )
            explanation_count = int(
                connection.execute(
                    "SELECT COUNT(*) FROM explanations"
                ).fetchone()[0]
            )
        finally:
            connection.close()

    volume = shutil.disk_usage(study_dir)
    write_rows(
        output_dir / "storage.tsv",
        ((
            database_bytes,
            media_bytes,
            media_count,
            backup_bytes,
            backup_count,
            trash_bytes,
            trash_count,
            other_bytes,
            total_bytes,
            volume.free,
            volume.total,
            source_count,
            explanation_count,
        ),),
    )
    return 0


def parse_local_datetime(value: str) -> datetime:
    """Parse a UI date/time as local time and attach the applicable UTC offset."""
    value = value.strip()
    for date_format in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, date_format).astimezone()
        except ValueError:
            continue
    raise ValueError(f"Invalid Study Library date/time: {value!r}")


def date_filter_bounds(mode: str, custom_from: str, custom_to: str) -> tuple[
    datetime | None, datetime | None, bool
]:
    """Return local date bounds and whether the upper boundary is inclusive."""
    now = datetime.now().astimezone()
    today_naive = datetime(now.year, now.month, now.day)
    if mode == "today":
        return today_naive.astimezone(), (today_naive + timedelta(days=1)).astimezone(), False
    if mode == "yesterday":
        return (today_naive - timedelta(days=1)).astimezone(), today_naive.astimezone(), False
    if mode == "last24":
        return now - timedelta(hours=24), now, True
    if mode == "last7":
        return now - timedelta(days=7), now, True
    if mode == "custom":
        start = parse_local_datetime(custom_from)
        end = parse_local_datetime(custom_to)
        if end < start:
            raise ValueError("The end of the generated-date range is before its start.")
        return start, end, True
    return None, None, False


def snapshot(database: Path, output_dir: Path) -> int:
    output_names = (
        "profiles.tsv",
        "chapters.tsv",
        "speakers.tsv",
        "tags.tsv",
        "groups.tsv",
    )
    clear_outputs(output_dir, output_names)
    if not database.is_file():
        for output_name in output_names:
            write_text(output_dir / output_name, "")
        return 0

    query = os.environ.get("STUDY_LIBRARY_QUERY", "").strip()
    profile_mode = os.environ.get("STUDY_LIBRARY_PROFILE_MODE", "all").strip()
    profile = os.environ.get("STUDY_LIBRARY_PROFILE_FILTER", "")
    chapter_mode = os.environ.get("STUDY_LIBRARY_CHAPTER_MODE", "all").strip()
    chapter = os.environ.get("STUDY_LIBRARY_CHAPTER_FILTER", "")
    speaker_mode = os.environ.get("STUDY_LIBRARY_SPEAKER_MODE", "all").strip()
    speaker = os.environ.get("STUDY_LIBRARY_SPEAKER_FILTER", "")
    tag_mode = os.environ.get("STUDY_LIBRARY_TAG_MODE", "all").strip()
    tag = os.environ.get("STUDY_LIBRARY_TAG_FILTER", "")
    anki_mode = os.environ.get("STUDY_LIBRARY_ANKI_MODE", "all").strip()
    date_mode = os.environ.get("STUDY_LIBRARY_DATE_MODE", "all").strip()
    date_from = os.environ.get("STUDY_LIBRARY_DATE_FROM", "").strip()
    date_to = os.environ.get("STUDY_LIBRARY_DATE_TO", "").strip()
    generated_from, generated_to, inclusive_to = date_filter_bounds(
        date_mode, date_from, date_to
    )

    connection = connect_read_only(database)
    try:
        profiles = [
            row[0]
            for row in connection.execute(
                "SELECT DISTINCT game_profile FROM explanation_groups "
                "WHERE game_profile <> '' ORDER BY game_profile COLLATE NOCASE"
            )
        ]
        write_rows(output_dir / "profiles.tsv", ((encode_field(p),) for p in profiles))
        chapters = [
            row[0]
            for row in connection.execute(
                "SELECT DISTINCT chapter FROM explanation_group_details "
                "WHERE chapter <> '' ORDER BY chapter COLLATE NOCASE"
            )
        ]
        speakers = [
            row[0]
            for row in connection.execute(
                "SELECT DISTINCT speaker FROM explanation_group_details "
                "WHERE speaker <> '' ORDER BY speaker COLLATE NOCASE"
            )
        ]
        tags = [
            row[0]
            for row in connection.execute(
                "SELECT name FROM study_tags ORDER BY name COLLATE NOCASE"
            )
        ]
        write_rows(output_dir / "chapters.tsv", ((encode_field(v),) for v in chapters))
        write_rows(output_dir / "speakers.tsv", ((encode_field(v),) for v in speakers))
        write_rows(output_dir / "tags.tsv", ((encode_field(v),) for v in tags))

        where = []
        parameters: list[object] = []
        if profile_mode == "unsorted":
            where.append("g.game_profile = ''")
        elif profile_mode == "profile":
            where.append("g.game_profile = ?")
            parameters.append(profile)
        if chapter_mode == "empty":
            where.append("COALESCE(d.chapter, '') = ''")
        elif chapter_mode == "value":
            where.append("COALESCE(d.chapter, '') = ? COLLATE NOCASE")
            parameters.append(chapter)
        if speaker_mode == "empty":
            where.append("COALESCE(d.speaker, '') = ''")
        elif speaker_mode == "value":
            where.append("COALESCE(d.speaker, '') = ? COLLATE NOCASE")
            parameters.append(speaker)
        if tag_mode == "empty":
            where.append(
                "NOT EXISTS (SELECT 1 FROM explanation_group_tags empty_gt "
                "WHERE empty_gt.group_id = g.id) AND COALESCE(d.tags, '') = ''"
            )
        elif tag_mode == "value":
            where.append(
                "EXISTS (SELECT 1 FROM explanation_group_tags filter_gt "
                "JOIN study_tags filter_t ON filter_t.id = filter_gt.tag_id "
                "WHERE filter_gt.group_id = g.id "
                "AND filter_t.name = ? COLLATE NOCASE)"
            )
            parameters.append(tag)
        if anki_mode == "found":
            where.append("COALESCE(a.status, 'not_checked') = 'found'")
        elif anki_mode == "not-found":
            where.append("COALESCE(a.status, 'not_checked') = 'not_found'")
        elif anki_mode == "not-checked":
            where.append("COALESCE(a.status, 'not_checked') = 'not_checked'")
        if generated_from is not None:
            where.append("julianday(g.updated_at) >= julianday(?)")
            parameters.append(generated_from.isoformat(timespec="seconds"))
        if generated_to is not None:
            where.append(
                "julianday(g.updated_at) <= julianday(?)" if inclusive_to
                else "julianday(g.updated_at) < julianday(?)"
            )
            parameters.append(generated_to.isoformat(timespec="seconds"))
        if query:
            where.append(
                "(g.source_japanese LIKE ? OR COALESCE(d.chapter, '') LIKE ? "
                "OR COALESCE(d.speaker, '') LIKE ? OR COALESCE(d.tags, '') LIKE ? "
                "OR EXISTS (SELECT 1 FROM explanation_group_tags search_gt "
                "JOIN study_tags search_t ON search_t.id = search_gt.tag_id "
                "WHERE search_gt.group_id = g.id AND search_t.name LIKE ?) "
                "OR EXISTS ("
                "SELECT 1 FROM explanations search_e "
                "WHERE search_e.group_id = g.id "
                "AND (search_e.raw_text LIKE ? OR search_e.key_grammar LIKE ?)))"
            )
            like = f"%{query}%"
            parameters.extend((like, like, like, like, like, like, like))

        sql = """
            SELECT g.id, g.game_profile, g.updated_at, g.source_japanese,
                   g.source_kind, COUNT(e.id) AS version_count,
                   COALESCE(d.chapter, '') AS chapter,
                   COALESCE(d.speaker, '') AS speaker,
                   COALESCE((
                       SELECT GROUP_CONCAT(tag_name, ', ') FROM (
                           SELECT display_t.name AS tag_name
                           FROM explanation_group_tags display_gt
                           JOIN study_tags display_t ON display_t.id = display_gt.tag_id
                           WHERE display_gt.group_id = g.id
                           ORDER BY display_t.name COLLATE NOCASE
                       )
                   ), COALESCE(d.tags, '')) AS tags,
                   COALESCE(d.added_to_anki_at, '') AS added_to_anki_at,
                   COALESCE(a.status, 'not_checked') AS anki_status,
                   COALESCE(a.note_id, 0) AS anki_note_id,
                   COALESCE(a.checked_at, '') AS anki_checked_at,
                   COALESCE((
                       SELECT latest_e.key_grammar
                       FROM explanations latest_e
                       WHERE latest_e.group_id = g.id
                       ORDER BY latest_e.version DESC
                       LIMIT 1
                   ), '') AS key_grammar
            FROM explanation_groups g
            JOIN explanations e ON e.group_id = g.id
            LEFT JOIN explanation_group_details d ON d.group_id = g.id
            LEFT JOIN explanation_group_anki_links a ON a.group_id = g.id
        """
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " GROUP BY g.id ORDER BY g.updated_at DESC, g.id DESC"

        rows = []
        for row in connection.execute(sql, parameters):
            preview = compact(row["source_japanese"])
            if not preview:
                preview = "[Source supplied as screenshot]"
            updated = str(row["updated_at"] or "").replace("T", " ")[:16]
            rows.append(
                (
                    int(row["id"]),
                    encode_field(row["game_profile"]),
                    encode_field(updated),
                    int(row["version_count"]),
                    encode_field(preview),
                    encode_field(row["source_kind"]),
                    encode_field(row["chapter"]),
                    encode_field(row["speaker"]),
                    encode_field(row["tags"]),
                    encode_field(row["added_to_anki_at"]),
                    encode_field(row["anki_status"]),
                    int(row["anki_note_id"]),
                    encode_field(row["anki_checked_at"]),
                    encode_field(row["key_grammar"]),
                )
            )
        write_rows(output_dir / "groups.tsv", rows)
        return 0
    finally:
        connection.close()


def detail(database: Path, output_dir: Path, group_id: int, version: int) -> int:
    names = (
        "detail.tsv",
        "versions.tsv",
        "media.tsv",
        "sections.tsv",
        "source.txt",
        "explanation.txt",
    )
    clear_outputs(output_dir, names)
    for old_section in output_dir.glob("section_*.txt"):
        try:
            old_section.unlink()
        except FileNotFoundError:
            pass
    if not database.is_file():
        return 0

    connection = connect_read_only(database)
    try:
        group = connection.execute(
            "SELECT g.*, COALESCE(d.chapter, '') AS chapter, "
            "COALESCE(d.speaker, '') AS speaker, "
            "COALESCE((SELECT GROUP_CONCAT(tag_name, ', ') FROM ("
            "SELECT detail_t.name AS tag_name "
            "FROM explanation_group_tags detail_gt "
            "JOIN study_tags detail_t ON detail_t.id = detail_gt.tag_id "
            "WHERE detail_gt.group_id = g.id "
            "ORDER BY detail_t.name COLLATE NOCASE)), COALESCE(d.tags, '')) AS tags, "
            "COALESCE(d.added_to_anki_at, '') AS added_to_anki_at, "
            "COALESCE(a.status, 'not_checked') AS anki_status, "
            "COALESCE(a.note_id, 0) AS anki_note_id, "
            "COALESCE(a.checked_at, '') AS anki_checked_at "
            "FROM explanation_groups g "
            "LEFT JOIN explanation_group_details d ON d.group_id = g.id "
            "LEFT JOIN explanation_group_anki_links a ON a.group_id = g.id "
            "WHERE g.id = ?",
            (group_id,),
        ).fetchone()
        if group is None:
            return 0

        versions = connection.execute(
            "SELECT id, version, preferred, created_at, provider, model, prompt_profile, "
            "manually_edited_at "
            "FROM explanations WHERE group_id = ? ORDER BY version DESC",
            (group_id,),
        ).fetchall()
        write_rows(
            output_dir / "versions.tsv",
            (
                (
                    int(row["id"]),
                    int(row["version"]),
                    int(row["preferred"]),
                    encode_field(str(row["created_at"] or "").replace("T", " ")[:19]),
                    encode_field(str(row["manually_edited_at"] or "").replace("T", " ")[:19]),
                )
                for row in versions
            ),
        )

        selected = None
        if version > 0:
            selected = connection.execute(
                "SELECT * FROM explanations WHERE group_id = ? AND version = ?",
                (group_id, version),
            ).fetchone()
        if selected is None:
            selected = connection.execute(
                "SELECT * FROM explanations WHERE group_id = ? "
                "ORDER BY preferred DESC, version DESC LIMIT 1",
                (group_id,),
            ).fetchone()
        if selected is None:
            return 0

        write_text(output_dir / "source.txt", str(group["source_japanese"] or ""))
        write_text(output_dir / "explanation.txt", str(selected["raw_text"] or ""))
        section_rows = connection.execute(
            "SELECT sort_order, section_key, heading, content "
            "FROM explanation_sections WHERE explanation_id = ? ORDER BY sort_order",
            (int(selected["id"]),),
        ).fetchall()
        section_output = []
        for section in section_rows:
            filename = f"section_{int(section['sort_order']):03d}.txt"
            write_text(output_dir / filename, str(section["content"] or ""))
            section_output.append(
                (
                    int(section["sort_order"]),
                    encode_field(section["section_key"]),
                    encode_field(section["heading"]),
                    encode_field(filename),
                )
            )
        write_rows(output_dir / "sections.tsv", section_output)
        write_rows(
            output_dir / "detail.tsv",
            (
                (
                    int(group["id"]),
                    int(selected["id"]),
                    int(selected["version"]),
                    int(selected["preferred"]),
                    encode_field(group["game_profile"]),
                    encode_field(str(selected["created_at"] or "").replace("T", " ")[:19]),
                    encode_field(selected["provider"]),
                    encode_field(selected["model"]),
                    encode_field(selected["prompt_profile"]),
                    encode_field(group["source_kind"]),
                    encode_field(group["chapter"]),
                    encode_field(group["speaker"]),
                    encode_field(group["tags"]),
                    encode_field(group["added_to_anki_at"]),
                    encode_field(
                        str(selected["manually_edited_at"] or "").replace("T", " ")[:19]
                    ),
                    encode_field(group["anki_status"]),
                    int(group["anki_note_id"]),
                    encode_field(group["anki_checked_at"]),
                ),
            ),
        )

        media_rows = connection.execute(
            "SELECT sort_order, relative_path, original_name, mime_type, sha256 "
            "FROM media WHERE explanation_id = ? ORDER BY sort_order",
            (int(selected["id"]),),
        ).fetchall()
        study_dir = database.parent
        output_media = []
        for row in media_rows:
            relative_path = str(row["relative_path"] or "")
            media_path = (study_dir / Path(relative_path)).resolve()
            output_media.append(
                (
                    int(row["sort_order"]),
                    encode_field(str(media_path)),
                    encode_field(relative_path),
                    encode_field(row["original_name"]),
                    encode_field(row["mime_type"]),
                    encode_field(row["sha256"]),
                    1 if media_path.is_file() else 0,
                )
            )
        write_rows(output_dir / "media.tsv", output_media)
        return 0
    finally:
        connection.close()


def set_metadata(database: Path, output_dir: Path, group_id: int) -> int:
    clear_outputs(output_dir, ("mutation.tsv",))
    chapter = os.environ.get("STUDY_LIBRARY_CHAPTER", "").strip()[:240]
    speaker = os.environ.get("STUDY_LIBRARY_SPEAKER", "").strip()[:240]
    tags = ", ".join(normalize_tags(os.environ.get("STUDY_LIBRARY_TAGS", "")))[:1000]
    anki_setting = os.environ.get("STUDY_LIBRARY_ADDED_TO_ANKI", "keep").strip()
    connection = connect_write(database)
    try:
        connection.execute("BEGIN IMMEDIATE")
        exists = connection.execute(
            "SELECT 1 FROM explanation_groups WHERE id = ?", (group_id,)
        ).fetchone()
        if exists is None:
            raise ValueError("The selected explanation no longer exists.")
        now = datetime.now().astimezone().isoformat(timespec="seconds")
        current_detail = connection.execute(
            "SELECT added_to_anki_at FROM explanation_group_details WHERE group_id = ?",
            (group_id,),
        ).fetchone()
        current_anki_at = str(current_detail[0] or "") if current_detail else ""
        if anki_setting == "1":
            added_to_anki_at = current_anki_at or now
        elif anki_setting == "0":
            added_to_anki_at = ""
        else:
            added_to_anki_at = current_anki_at
        connection.execute(
            "INSERT INTO explanation_group_details("
            "group_id, chapter, speaker, tags, added_to_anki_at, updated_at"
            ") VALUES(?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(group_id) DO UPDATE SET "
            "chapter=excluded.chapter, speaker=excluded.speaker, "
            "tags=excluded.tags, added_to_anki_at=excluded.added_to_anki_at, "
            "updated_at=excluded.updated_at",
            (group_id, chapter, speaker, tags, added_to_anki_at, now),
        )
        tags = replace_group_tags(connection, group_id, tags)
        connection.execute(
            "UPDATE explanation_group_details SET tags = ? WHERE group_id = ?",
            (tags, group_id),
        )
        connection.commit()
        write_rows(
            output_dir / "mutation.tsv",
            ((
                group_id,
                encode_field(chapter),
                encode_field(speaker),
                encode_field(tags),
                encode_field(added_to_anki_at),
            ),),
        )
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def parse_bulk_group_ids() -> list[int]:
    raw_value = os.environ.get("STUDY_LIBRARY_GROUP_IDS", "")
    group_ids: list[int] = []
    seen: set[int] = set()
    for raw_id in raw_value.split(","):
        raw_id = raw_id.strip()
        if not raw_id:
            continue
        try:
            group_id = int(raw_id)
        except ValueError as exc:
            raise ValueError(f"Invalid bulk-edit group id: {raw_id!r}") from exc
        if group_id <= 0:
            raise ValueError("Bulk-edit group ids must be positive.")
        if group_id not in seen:
            seen.add(group_id)
            group_ids.append(group_id)
    if not group_ids:
        raise ValueError("Bulk metadata editing requires at least one group id.")
    if len(group_ids) > 10000:
        raise ValueError("Bulk metadata editing is limited to 10,000 entries at once.")
    return group_ids


def bulk_metadata(database: Path, output_dir: Path) -> int:
    """Atomically apply explicit metadata operations to selected groups."""
    clear_outputs(output_dir, ("mutation.tsv",))
    group_ids = parse_bulk_group_ids()
    chapter_mode = os.environ.get(
        "STUDY_LIBRARY_BULK_CHAPTER_MODE", "keep"
    ).strip()
    speaker_mode = os.environ.get(
        "STUDY_LIBRARY_BULK_SPEAKER_MODE", "keep"
    ).strip()
    tags_mode = os.environ.get(
        "STUDY_LIBRARY_BULK_TAGS_MODE", "keep"
    ).strip()
    anki_mode = os.environ.get(
        "STUDY_LIBRARY_BULK_ANKI_MODE", "keep"
    ).strip()
    chapter_value = os.environ.get("STUDY_LIBRARY_BULK_CHAPTER", "").strip()[:240]
    speaker_value = os.environ.get("STUDY_LIBRARY_BULK_SPEAKER", "").strip()[:240]
    requested_tags = normalize_tags(
        os.environ.get("STUDY_LIBRARY_BULK_TAGS", "")
    )
    if chapter_mode not in {"keep", "set", "clear"}:
        raise ValueError(f"Invalid bulk chapter operation: {chapter_mode!r}")
    if speaker_mode not in {"keep", "set", "clear"}:
        raise ValueError(f"Invalid bulk speaker operation: {speaker_mode!r}")
    if tags_mode not in {"keep", "add", "remove", "replace", "clear"}:
        raise ValueError(f"Invalid bulk tag operation: {tags_mode!r}")
    if anki_mode not in {"keep", "1", "0"}:
        raise ValueError(f"Invalid bulk Anki operation: {anki_mode!r}")
    if chapter_mode == "set" and not chapter_value:
        raise ValueError("Set chapter requires a non-empty value.")
    if speaker_mode == "set" and not speaker_value:
        raise ValueError("Set speaker requires a non-empty value.")
    if tags_mode in {"add", "remove", "replace"} and not requested_tags:
        raise ValueError(f"{tags_mode.title()} tags requires at least one tag.")
    if all(
        mode == "keep"
        for mode in (chapter_mode, speaker_mode, tags_mode, anki_mode)
    ):
        raise ValueError("No bulk metadata changes were selected.")

    connection = connect_write(database)
    mutation_rows: list[tuple[object, ...]] = []
    try:
        connection.execute("BEGIN IMMEDIATE")
        placeholders = ",".join("?" for _ in group_ids)
        existing_ids = {
            int(row[0])
            for row in connection.execute(
                f"SELECT id FROM explanation_groups WHERE id IN ({placeholders})",
                group_ids,
            )
        }
        missing_ids = [group_id for group_id in group_ids if group_id not in existing_ids]
        if missing_ids:
            raise ValueError(
                "Selected explanations no longer exist: "
                + ", ".join(str(group_id) for group_id in missing_ids)
            )
        now = datetime.now().astimezone().isoformat(timespec="seconds")
        requested_tag_text = ", ".join(requested_tags)
        remove_tag_keys = {tag.casefold() for tag in requested_tags}
        for group_id in group_ids:
            current = connection.execute(
                "SELECT COALESCE(d.chapter, '') AS chapter, "
                "COALESCE(d.speaker, '') AS speaker, "
                "COALESCE((SELECT GROUP_CONCAT(tag_name, ', ') FROM ("
                "SELECT t.name AS tag_name FROM explanation_group_tags gt "
                "JOIN study_tags t ON t.id = gt.tag_id "
                "WHERE gt.group_id = g.id ORDER BY t.name COLLATE NOCASE"
                ")), COALESCE(d.tags, '')) AS tags, "
                "COALESCE(d.added_to_anki_at, '') AS added_to_anki_at "
                "FROM explanation_groups g "
                "LEFT JOIN explanation_group_details d ON d.group_id = g.id "
                "WHERE g.id = ?",
                (group_id,),
            ).fetchone()
            current_chapter = str(current["chapter"] or "")
            current_speaker = str(current["speaker"] or "")
            current_tags = normalize_tags(str(current["tags"] or ""))
            current_anki_at = str(current["added_to_anki_at"] or "")

            chapter = (
                current_chapter if chapter_mode == "keep"
                else chapter_value if chapter_mode == "set" else ""
            )
            speaker = (
                current_speaker if speaker_mode == "keep"
                else speaker_value if speaker_mode == "set" else ""
            )
            if tags_mode == "keep":
                tags = current_tags
            elif tags_mode == "add":
                tags = normalize_tags(", ".join(current_tags + requested_tags))
            elif tags_mode == "remove":
                tags = [tag for tag in current_tags if tag.casefold() not in remove_tag_keys]
            elif tags_mode == "replace":
                tags = normalize_tags(requested_tag_text)
            else:
                tags = []
            tag_text = ", ".join(tags)
            if anki_mode == "1":
                added_to_anki_at = current_anki_at or now
            elif anki_mode == "0":
                added_to_anki_at = ""
            else:
                added_to_anki_at = current_anki_at

            connection.execute(
                "INSERT INTO explanation_group_details("
                "group_id, chapter, speaker, tags, added_to_anki_at, updated_at"
                ") VALUES(?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(group_id) DO UPDATE SET "
                "chapter=excluded.chapter, speaker=excluded.speaker, "
                "tags=excluded.tags, added_to_anki_at=excluded.added_to_anki_at, "
                "updated_at=excluded.updated_at",
                (group_id, chapter, speaker, tag_text, added_to_anki_at, now),
            )
            tag_text = replace_group_tags(connection, group_id, tag_text)
            connection.execute(
                "UPDATE explanation_group_details SET tags = ? WHERE group_id = ?",
                (tag_text, group_id),
            )
            mutation_rows.append(
                (
                    group_id,
                    encode_field(chapter),
                    encode_field(speaker),
                    encode_field(tag_text),
                    encode_field(added_to_anki_at),
                )
            )
        connection.commit()
        write_rows(output_dir / "mutation.tsv", mutation_rows)
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def set_anki(database: Path, output_dir: Path, group_id: int) -> int:
    """Update only the review marker without touching chapter/speaker/tags."""
    clear_outputs(output_dir, ("mutation.tsv",))
    anki_setting = os.environ.get("STUDY_LIBRARY_ADDED_TO_ANKI", "").strip()
    if anki_setting not in {"0", "1"}:
        raise ValueError("set-anki requires STUDY_LIBRARY_ADDED_TO_ANKI=0 or 1")
    connection = connect_write(database)
    try:
        connection.execute("BEGIN IMMEDIATE")
        group = connection.execute(
            "SELECT g.id, COALESCE(d.chapter, '') AS chapter, "
            "COALESCE(d.speaker, '') AS speaker, COALESCE(d.tags, '') AS tags, "
            "COALESCE(d.added_to_anki_at, '') AS added_to_anki_at "
            "FROM explanation_groups g "
            "LEFT JOIN explanation_group_details d ON d.group_id = g.id "
            "WHERE g.id = ?",
            (group_id,),
        ).fetchone()
        if group is None:
            raise ValueError("The selected explanation no longer exists.")
        now = datetime.now().astimezone().isoformat(timespec="seconds")
        current_anki_at = str(group["added_to_anki_at"] or "")
        added_to_anki_at = (current_anki_at or now) if anki_setting == "1" else ""
        connection.execute(
            "INSERT INTO explanation_group_details("
            "group_id, chapter, speaker, tags, added_to_anki_at, updated_at"
            ") VALUES(?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(group_id) DO UPDATE SET "
            "added_to_anki_at=excluded.added_to_anki_at, "
            "updated_at=excluded.updated_at",
            (
                group_id,
                str(group["chapter"] or ""),
                str(group["speaker"] or ""),
                str(group["tags"] or ""),
                added_to_anki_at,
                now,
            ),
        )
        connection.commit()
        write_rows(
            output_dir / "mutation.tsv",
            ((
                group_id,
                encode_field(group["chapter"]),
                encode_field(group["speaker"]),
                encode_field(group["tags"]),
                encode_field(added_to_anki_at),
            ),),
        )
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def save_manual_edit(
    database: Path, output_dir: Path, group_id: int, version: int
) -> int:
    """Save a reversible user edit and rebuild its derived section index."""
    clear_outputs(output_dir, ("mutation.tsv",))
    edit_path = output_dir / "manual_edit.txt"
    if not edit_path.is_file():
        raise ValueError("The edited explanation text was not supplied.")
    edited_text = edit_path.read_text(encoding="utf-8")
    try:
        edit_path.unlink()
    except OSError:
        pass
    edited_text = edited_text.replace("\r\n", "\n").replace("\r", "\n")
    if not edited_text.strip():
        raise ValueError("An explanation cannot be saved empty.")

    connection = connect_write(database)
    try:
        ensure_section_schema(connection)
        connection.commit()
        connection.execute("BEGIN IMMEDIATE")
        selected = connection.execute(
            "SELECT id, raw_text, manual_original_text FROM explanations "
            "WHERE group_id = ? AND version = ?",
            (group_id, version),
        ).fetchone()
        if selected is None:
            raise ValueError("The selected explanation version no longer exists.")
        explanation_id = int(selected["id"])
        current_text = str(selected["raw_text"] or "")
        original_text = str(selected["manual_original_text"] or "") or current_text
        edited_at = datetime.now().astimezone().isoformat(timespec="seconds")
        connection.execute(
            "UPDATE explanations SET raw_text = ?, manual_original_text = ?, "
            "manually_edited_at = ? WHERE id = ?",
            (edited_text, original_text, edited_at, explanation_id),
        )
        section_count = replace_explanation_sections(
            connection, explanation_id, edited_text
        )
        connection.commit()
        write_rows(
            output_dir / "mutation.tsv",
            ((explanation_id, encode_field(edited_at), section_count),),
        )
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def revert_manual_edit(
    database: Path, output_dir: Path, group_id: int, version: int
) -> int:
    """Restore the untouched model output preserved by the first manual edit."""
    clear_outputs(output_dir, ("mutation.tsv",))
    connection = connect_write(database)
    try:
        ensure_section_schema(connection)
        connection.commit()
        connection.execute("BEGIN IMMEDIATE")
        selected = connection.execute(
            "SELECT id, manual_original_text FROM explanations "
            "WHERE group_id = ? AND version = ?",
            (group_id, version),
        ).fetchone()
        if selected is None:
            raise ValueError("The selected explanation version no longer exists.")
        original_text = str(selected["manual_original_text"] or "")
        if not original_text:
            raise ValueError("This explanation has no manual edits to revert.")
        explanation_id = int(selected["id"])
        connection.execute(
            "UPDATE explanations SET raw_text = ?, manual_original_text = '', "
            "manually_edited_at = '' WHERE id = ?",
            (original_text, explanation_id),
        )
        section_count = replace_explanation_sections(
            connection, explanation_id, original_text
        )
        connection.commit()
        write_rows(
            output_dir / "mutation.tsv",
            ((explanation_id, "", section_count),),
        )
        return 0
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def backup_database(connection: sqlite3.Connection, database: Path) -> Path:
    backup_dir = database.parent / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    backup_path = backup_dir / f"{stamp}_before_remove.db"
    backup_connection = sqlite3.connect(backup_path)
    try:
        connection.backup(backup_connection)
    finally:
        backup_connection.close()
    return backup_path


def remove_version(
    database: Path, output_dir: Path, group_id: int, version: int
) -> int:
    """Remove one version after making a recoverable database/media backup."""
    clear_outputs(output_dir, ("mutation.tsv",))
    connection = connect_write(database)
    moved_media: list[tuple[Path, Path]] = []
    backup_path: Path | None = None
    try:
        selected = connection.execute(
            "SELECT id FROM explanations WHERE group_id = ? AND version = ?",
            (group_id, version),
        ).fetchone()
        if selected is None:
            raise ValueError("The selected explanation version no longer exists.")
        explanation_id = int(selected["id"])
        media_rows = connection.execute(
            "SELECT relative_path FROM media WHERE explanation_id = ?",
            (explanation_id,),
        ).fetchall()
        backup_path = backup_database(connection, database)
        trash_dir = database.parent / "Trash" / backup_path.stem

        connection.execute("BEGIN IMMEDIATE")
        for media_row in media_rows:
            relative = Path(str(media_row["relative_path"] or ""))
            source = (database.parent / relative).resolve()
            if not source.is_relative_to(database.parent.resolve()) or not source.is_file():
                continue
            trash_dir.mkdir(parents=True, exist_ok=True)
            destination = trash_dir / source.name
            suffix = 2
            while destination.exists():
                destination = trash_dir / f"{source.stem}_{suffix}{source.suffix}"
                suffix += 1
            shutil.move(str(source), str(destination))
            moved_media.append((source, destination))

        connection.execute("DELETE FROM explanations WHERE id = ?", (explanation_id,))
        remaining = connection.execute(
            "SELECT id, version, created_at FROM explanations "
            "WHERE group_id = ? ORDER BY version DESC LIMIT 1",
            (group_id,),
        ).fetchone()
        if remaining is None:
            connection.execute("DELETE FROM explanation_groups WHERE id = ?", (group_id,))
            group_exists = 0
            next_version = 0
        else:
            connection.execute(
                "UPDATE explanations SET preferred = CASE WHEN id = ? THEN 1 ELSE 0 END "
                "WHERE group_id = ?",
                (int(remaining["id"]), group_id),
            )
            connection.execute(
                "UPDATE explanation_groups SET updated_at = ? WHERE id = ?",
                (str(remaining["created_at"]), group_id),
            )
            group_exists = 1
            next_version = int(remaining["version"])
        connection.execute(
            "DELETE FROM study_tags WHERE NOT EXISTS ("
            "SELECT 1 FROM explanation_group_tags gt WHERE gt.tag_id = study_tags.id)"
        )
        connection.commit()
        write_rows(
            output_dir / "mutation.tsv",
            (
                (
                    group_exists,
                    next_version,
                    encode_field(str(backup_path)),
                    encode_field(str(trash_dir) if moved_media else ""),
                ),
            ),
        )
        return 0
    except Exception:
        connection.rollback()
        for source, destination in reversed(moved_media):
            try:
                source.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(destination), str(source))
            except OSError:
                pass
        raise
    finally:
        connection.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="JRPG Translator Study Library bridge")
    parser.add_argument(
        "command",
        choices=(
            "ensure", "snapshot", "detail", "set-metadata", "bulk-metadata",
            "set-anki", "save-edit", "revert-edit", "remove-version", "storage",
        ),
    )
    parser.add_argument("--db", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--group-id", type=int, default=0)
    parser.add_argument("--version", type=int, default=0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    database = Path(args.db)
    output_dir = Path(args.output_dir)
    if args.command == "ensure":
        return ensure_database(database, output_dir)
    if args.command == "snapshot":
        return snapshot(database, output_dir)
    if args.command == "storage":
        return storage_snapshot(database, output_dir)
    if args.command == "bulk-metadata":
        return bulk_metadata(database, output_dir)
    if args.group_id <= 0:
        raise ValueError(f"{args.command} requires a positive --group-id")
    if args.command == "set-metadata":
        return set_metadata(database, output_dir, args.group_id)
    if args.command == "set-anki":
        return set_anki(database, output_dir, args.group_id)
    if args.command in {"save-edit", "revert-edit"}:
        if args.version <= 0:
            raise ValueError(f"{args.command} requires a positive --version")
        if args.command == "save-edit":
            return save_manual_edit(
                database, output_dir, args.group_id, args.version
            )
        return revert_manual_edit(database, output_dir, args.group_id, args.version)
    if args.command == "remove-version":
        if args.version <= 0:
            raise ValueError("remove-version requires a positive --version")
        return remove_version(database, output_dir, args.group_id, args.version)
    return detail(database, output_dir, args.group_id, args.version)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Study Library bridge error: {exc}", file=sys.stderr)
        raise SystemExit(1)
