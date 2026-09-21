"""Synthetic Anki service and database only; never contacts real Anki."""
import importlib.util
import os
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

SOURCE = Path(__file__).resolve().parents[1] / "scripts" / "anki_bridge.py"
SPEC = importlib.util.spec_from_file_location("anki_outcomes_bridge", SOURCE)
bridge = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(bridge)


class AddOutcomeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.db = self.root / "fixture.sqlite3"
        self.output = self.root / "output"
        self.output.mkdir()
        with closing(sqlite3.connect(self.db)) as connection:
            connection.executescript("""
                CREATE TABLE explanation_groups(id INTEGER PRIMARY KEY, game_profile TEXT);
                CREATE TABLE explanations(id INTEGER PRIMARY KEY, group_id INTEGER, version INTEGER);
                CREATE TABLE media(explanation_id INTEGER, relative_path TEXT);
                INSERT INTO explanation_groups VALUES(1, 'fixture');
                INSERT INTO explanations VALUES(1, 1, 1);
                INSERT INTO media VALUES(1, 'image.png');
            """)
        (self.root / "image.png").write_bytes(b"synthetic image; encoding is stubbed")
        (self.root / "front.txt").write_text("日本語", encoding="utf-8")
        (self.root / "back.txt").write_text("A synthetic explanation.", encoding="utf-8")
        env = {
            "STUDY_ANKI_PROFILE": "fixture", "STUDY_ANKI_PARENT_DECK": "Test",
            "STUDY_ANKI_DECK": "Test", "STUDY_ANKI_MODEL": "Basic",
            "STUDY_ANKI_JAPANESE_FIELD": "Front", "STUDY_ANKI_EXPLANATION_FIELD": "Back",
            "STUDY_ANKI_CARD_KIND": "explanation", "STUDY_ANKI_GROUP_ID": "1",
            "STUDY_ANKI_VERSION": "1", "STUDY_ANKI_INCLUDE_SCREENSHOT": "1",
            "STUDY_ANKI_SCREENSHOT_PATH": str(self.root / "image.png"),
            "STUDY_ANKI_FRONT_PATH": str(self.root / "front.txt"),
            "STUDY_ANKI_BACK_PATH": str(self.root / "back.txt"),
        }
        self.enterContext(patch.dict(os.environ, env))
        self.enterContext(patch.object(bridge, "probe", return_value=(True, 6)))
        self.enterContext(patch.object(bridge, "prepare_anki_screenshot", return_value=("fixture.jpg", "aW1hZ2U=", 5, (1, 1))))
        self.enterContext(patch.object(bridge, "invoke", side_effect=self.invoke))
        self.notes, self.media, self.calls = [], {}, []
        self.commit = True
        self.reply = 123
        self.read_failure = False
        self.corrupt_back = False
        self.duplicate_match = False

    def invoke(self, action, params=None, **kwargs):
        params = params or {}
        self.calls.append(action)
        if action == "deckNames":
            return ["Test"]
        if action == "findNotes":
            if self.read_failure and "addNote" in self.calls:
                raise TimeoutError("reconciliation unavailable")
            return [note["noteId"] for note in self.notes]
        if action == "notesInfo":
            return self.notes
        if action == "canAddNotesWithErrorDetail":
            return [{"canAdd": True}]
        if action == "storeMediaFile":
            self.media[params["filename"]] = params["data"]
            return params["filename"]
        if action == "addNote":
            if self.commit:
                note = params["note"]
                fields = {key: {"value": value} for key, value in note["fields"].items()}
                if self.corrupt_back:
                    fields["Back"]["value"] = "a different card"
                self.notes.append({"noteId": 123, "modelName": note["modelName"], "fields": fields})
                if self.duplicate_match:
                    self.notes.append({**self.notes[0], "noteId": 124})
            if isinstance(self.reply, Exception):
                raise self.reply
            return self.reply
        raise AssertionError(f"Unexpected or destructive Anki action: {action}")

    def run_add(self):
        self.assertEqual(bridge.add_note(self.db, self.output), 0)
        code, note_id, message_hex = (self.output / "anki_add.tsv").read_text().strip().split("\t")
        self.assertNotIn("deleteMediaFile", self.calls)
        return code, note_id, bytes.fromhex(message_hex).decode("utf-8")

    def assert_link(self):
        with closing(sqlite3.connect(self.db)) as connection:
            self.assertEqual(connection.execute("SELECT note_id FROM explanation_group_anki_links WHERE group_id=1").fetchone(), (123,))

    def test_normal_success(self):
        self.assertEqual(self.run_add()[:2], ("added", "123"))
        self.assert_link()

    def test_committed_but_reply_lost_preserves_media_and_recovers_link(self):
        self.reply = TimeoutError("reply lost")
        code, note_id, message = self.run_add()
        self.assertEqual((code, note_id), ("added", "123"))
        self.assertIn("verified", message)
        self.assertIn("fixture.jpg", self.media)
        self.assertEqual(self.calls.count("addNote"), 1)
        self.assert_link()
        # An explicit later retry performs duplicate detection before mutation.
        self.assertEqual(self.run_add()[:2], ("duplicate", "123"))
        self.assertEqual(self.calls.count("addNote"), 1)

    def test_uncommitted_timeout_stays_uncertain_and_retains_media(self):
        self.commit = False
        self.reply = TimeoutError("request outcome unknown")
        code, _, message = self.run_add()
        self.assertEqual(code, "uncertain")
        self.assertIn("may already", message)
        self.assertIn("fixture.jpg", self.media)
        self.assertEqual(self.calls.count("addNote"), 1)

    def test_failed_reconciliation_does_not_claim_failure_or_delete_media(self):
        self.reply = bridge.AnkiUnavailable("connect_missing", "response lost")
        self.read_failure = True
        self.assertEqual(self.run_add()[0], "uncertain")
        self.assertEqual(len(self.notes), 1)
        self.assertIn("fixture.jpg", self.media)

    def test_explicit_rejection_preserves_shared_media(self):
        self.commit = False
        self.reply = bridge.AnkiUnavailable("anki_error", "rejected")
        self.assertEqual(self.run_add()[0], "blocked")
        self.assertIn("fixture.jpg", self.media)

    def test_invalid_success_replies_are_reconciled_not_blindly_accepted(self):
        for reply in (None, 0, True, "123", {}, []):
            with self.subTest(reply=reply):
                self.notes.clear()
                self.calls.clear()
                self.reply = reply
                self.assertEqual(self.run_add()[:2], ("added", "123"))
                self.assertEqual(self.calls.count("addNote"), 1)

    def test_matching_front_but_different_back_is_not_confirmation(self):
        self.reply = None
        self.corrupt_back = True
        self.assertEqual(self.run_add()[0], "uncertain")

    def test_multiple_exact_matches_are_not_confirmation(self):
        self.reply = None
        self.duplicate_match = True
        self.assertEqual(self.run_add()[0], "uncertain")

    def test_recovered_vocabulary_does_not_link_explanation_group(self):
        with patch.dict(os.environ, {"STUDY_ANKI_CARD_KIND": "vocabulary"}), patch.object(bridge, "record_group_link") as link:
            self.reply = TimeoutError("lost")
            self.assertEqual(self.run_add()[:2], ("added", "123"))
            link.assert_not_called()

    def test_local_link_failure_after_recovery_is_reported_accurately(self):
        self.reply = TimeoutError("lost")
        with patch.object(bridge, "record_group_link", side_effect=OSError("db locked")):
            self.assertEqual(self.run_add()[:2], ("added_unlinked", "123"))
        self.assertIn("fixture.jpg", self.media)

    def test_uncertain_without_screenshot(self):
        with patch.dict(os.environ, {"STUDY_ANKI_INCLUDE_SCREENSHOT": "0"}):
            self.commit = False
            self.reply = TimeoutError("lost")
            code, _, message = self.run_add()
            self.assertEqual(code, "uncertain")
            self.assertNotIn("screenshot", message)
            self.assertFalse(self.media)


if __name__ == "__main__":
    unittest.main()
