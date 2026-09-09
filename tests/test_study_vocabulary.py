from __future__ import annotations

import hashlib
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import study_library as library
import study_vocabulary as vocabulary
import anki_candidates as candidates


class VocabularyTests(unittest.TestCase):
    def test_front_keeps_okurigana_and_back_keeps_complete_entry(self):
        text = "• 抜ける（ぬける） — to slip out\n  Used when leaving a place.\n\nお城(しろ) — castle"
        entries = vocabulary.parse_vocabulary(text)
        self.assertEqual([e.front for e in entries], ["抜ける", "お城"])
        self.assertEqual(entries[0].display, "抜ける（ぬける）")
        self.assertEqual(entries[0].back, "抜ける（ぬける） — to slip out Used when leaving a place.")
        self.assertEqual(entries[0].meaning, "to slip out Used when leaving a place.")

    def test_review_uses_identical_parser(self):
        self.assertIs(candidates.parse_vocabulary, vocabulary.parse_vocabulary)
        self.assertIs(candidates.vocabulary_front, vocabulary.vocabulary_front)
        self.assertEqual(vocabulary.vocabulary_front("行った（いった） → 行く（いく）"), "行く")
        self.assertEqual(vocabulary.vocabulary_front("ありがとう"), "ありがとう")
        self.assertEqual(vocabulary.parse_vocabulary("Unstructured explanation without entries."), [])

    def test_detail_exports_only_selected_version_and_clears_stale_data(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            database = root / "library.db"
            output = root / "detail"
            with sqlite3.connect(database) as db:
                db.executescript("""
                    CREATE TABLE explanation_groups(id INTEGER, game_profile TEXT, source_japanese TEXT, source_kind TEXT);
                    CREATE TABLE explanation_group_details(group_id INTEGER, chapter TEXT, speaker TEXT, tags TEXT, added_to_anki_at TEXT);
                    CREATE TABLE explanation_group_tags(group_id INTEGER, tag_id INTEGER);
                    CREATE TABLE study_tags(id INTEGER, name TEXT);
                    CREATE TABLE explanation_group_anki_links(group_id INTEGER, status TEXT, note_id INTEGER, checked_at TEXT);
                    CREATE TABLE explanations(id INTEGER, group_id INTEGER, version INTEGER, preferred INTEGER,
                        created_at TEXT, provider TEXT, model TEXT, prompt_profile TEXT, manually_edited_at TEXT, raw_text TEXT);
                    CREATE TABLE explanation_sections(explanation_id INTEGER, sort_order INTEGER, section_key TEXT, heading TEXT, content TEXT);
                    CREATE TABLE media(explanation_id INTEGER, sort_order INTEGER, relative_path TEXT, original_name TEXT, mime_type TEXT, sha256 TEXT);
                    INSERT INTO explanation_groups VALUES(1, 'Demo', '日本語', 'screenshot');
                    INSERT INTO explanations VALUES(11, 1, 1, 0, '2026-09-01', 'fixture', 'fixture', 'default', '', 'Old');
                    INSERT INTO explanations VALUES(12, 1, 2, 1, '2026-09-02', 'fixture', 'fixture', 'default', '', 'New');
                    INSERT INTO explanations VALUES(13, 1, 3, 0, '2026-09-03', 'fixture', 'fixture', 'default', '', 'No vocabulary');
                    INSERT INTO explanation_sections VALUES(11, 1, 'vocabulary', 'Key vocabulary', '抜ける（ぬける） — to slip out');
                    INSERT INTO explanation_sections VALUES(12, 1, 'vocabulary', 'Key vocabulary', '城（しろ） — castle');
                    INSERT INTO explanation_sections VALUES(12, 2, 'analysis', 'Analysis', '別（べつ） — not a vocabulary entry here');
                """)
            db.close()
            before = hashlib.sha256(database.read_bytes()).digest()
            with patch('socket.create_connection', side_effect=AssertionError('Unexpected network request')):
                self.assertEqual(library.detail(database, output, 1, 1), 0)
                row = (output / 'reader_vocabulary.tsv').read_text().strip().split('\t')
                self.assertEqual(row[:2], ['1', '1'])
                self.assertEqual(bytes.fromhex(row[3]).decode(), '抜ける')
                self.assertEqual(bytes.fromhex(row[4]).decode(), '抜ける（ぬける） — to slip out')
                library.detail(database, output, 1, 0)
                rows = (output / 'reader_vocabulary.tsv').read_text().splitlines()
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0].split('\t')[:2], ['1', '2'])
                self.assertEqual(bytes.fromhex(rows[0].split('\t')[3]).decode(), '城')
                library.detail(database, output, 1, 3)
                self.assertEqual((output / 'reader_vocabulary.tsv').read_text(), '')
                library.detail(database, output, 1, 1)
                library.detail(database, output, 999, 1)
                self.assertFalse((output / 'reader_vocabulary.tsv').exists())
            self.assertEqual(hashlib.sha256(database.read_bytes()).digest(), before)


if __name__ == '__main__':
    unittest.main()
