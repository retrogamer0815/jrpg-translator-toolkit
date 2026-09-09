from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "scripts"))
import anki_candidates as candidates


class RecommendationPromptTests(unittest.TestCase):
    def test_custom_instructions_preserve_automatic_parts(self):
        items = [{"id": "fixture", "japanese": "日本語", "reason_language": "English"}]
        prompt = candidates.build_recommendation_prompt(
            items, "beginner", "generous", "grammar", "Avoid proper names.",
            "Prefer reusable dialogue.\nKeep useful inflections.")
        self.assertTrue(prompt.startswith("Prefer reusable dialogue.\nKeep useful inflections."))
        self.assertNotIn("Assess every item below", prompt)
        for automatic in ("The learner is a beginner", "40-60 percent", "Prioritize grammar patterns",
                          "Avoid proper names.", "Return only one JSON object", "reason_language"):
            self.assertIn(automatic, prompt)
        self.assertEqual(json.loads(prompt.split("\nCandidates:\n", 1)[1]), items)

    def test_default_prompt_and_cache_remain_compatible(self):
        # Golden values verified against the pre-editor implementation (prompt v3).
        settings = ("beginner", "generous", "grammar", "Extra guidance")
        self.assertEqual(hashlib.sha256(candidates.build_recommendation_prompt([], *settings).encode()).hexdigest(),
                         "8fe40be44163296880fd8c4006943ec258dff2be34b95369d375512176c80214")
        signature_args = ("openai", "fixture", *settings)
        previous = "057ac74a2afbd3acfa38f834eb84ae890591171dbe169e77389bbceda7c3f660"
        self.assertEqual(candidates.recommendation_profile_signature(*signature_args), previous)
        custom = candidates.recommendation_profile_signature(*signature_args, "Custom instructions")
        self.assertNotEqual(custom, previous)
        self.assertNotEqual(custom, candidates.recommendation_profile_signature(*signature_args, "Edited instructions"))
        self.assertEqual(candidates.build_recommendation_prompt([], *settings, "  "),
                         candidates.build_recommendation_prompt([], *settings))

    def test_utf8_file_transport_and_validation(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "instructions.txt"
            path.write_text('日本語 "quotes"\nA second line.', encoding="utf-8-sig")
            self.assertEqual(candidates.read_prompt_instructions(path), '日本語 "quotes"\nA second line.')
            self.assertEqual(candidates.read_prompt_instructions(None), "")
            with self.assertRaises(FileNotFoundError):
                candidates.read_prompt_instructions(Path(temp) / "missing.txt")
            path.write_text("x" * 16001, encoding="utf-8")
            with self.assertRaises(ValueError):
                candidates.read_prompt_instructions(path)

    def test_generation_uses_instruction_file_without_real_api_calls(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            prompt_file = root / "instructions.txt"
            prompt_file.write_text("Custom generation instructions 日本語", encoding="utf-8-sig")
            candidate_id = "a" * 64
            row = ["", "", "", candidates.encode_field("日本語"), "", candidate_id, "-1", "0", ""]
            (root / "candidate_sentences.tsv").write_text("\t".join(row), encoding="utf-8")
            response = json.dumps({"items": [{"id": candidate_id, "recommended": True,
                                              "score": 4, "reason": "Useful dialogue."}]})
            with patch.object(candidates, "call_model", return_value=response) as call:
                result = candidates.generate_recommendations(
                    root, root / "preferences.db", "openai", "fixture", prompt_file=prompt_file)
            self.assertEqual(result, 0)
            self.assertTrue(call.call_args.args[2].startswith("Custom generation instructions 日本語"))
            self.assertIn("Return only one JSON object", call.call_args.args[2])
            self.assertIn(candidate_id, call.call_args.args[2])

    def test_snapshot_uses_custom_instructions_for_cache_identity(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            database = root / "library.db"
            database.touch()
            prompt_file = root / "instructions.txt"
            prompt_file.write_text("Custom profile instructions", encoding="utf-8")
            # Stop at the profile calculation, before database migration or Anki discovery.
            with patch.object(candidates, "recommendation_profile_signature", side_effect=RuntimeError("stop")) as signature:
                with self.assertRaisesRegex(RuntimeError, "stop"):
                    candidates.snapshot(database, root, root / "anki.ini", root / "preferences.db",
                                        "all", prompt_file=prompt_file)
            self.assertEqual(signature.call_args.args[-1], "Custom profile instructions")


if __name__ == "__main__":
    unittest.main()
